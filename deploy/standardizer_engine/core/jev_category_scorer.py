# -*- coding: utf-8 -*-
"""
jev_category_scorer.py
基于 TypeSafe Jev (System One 判别式模型) + 标准数据 (category_lexicon.db + 工料机类别.xlsx) 的双档位分类识别与打分引擎
"""
import json
import os
import sqlite3
from functools import lru_cache
from typing import Dict, List, Any, Optional, Tuple, Set
import requests

from config import TYPESAFE_API_KEY, TYPESAFE_BASE_URL, TYPESAFE_MODEL, JEV_MIN_CONFIDENCE

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
import config
LEXICON_DB = config.CATEGORY_LEXICON_DB
FLYWHEEL_LOG = os.path.join(BASE_DIR, 'data', 'lexicon_flywheel.jsonl')


def _load_catalog():
    from category_attribute_service import _load_catalog as _lc, _CATALOG_ROWS
    _lc()
    return _CATALOG_ROWS


def catalog_by_t2() -> Dict[str, List[Dict[str, Any]]]:
    """{二级类别名: [该二级类下所有三级类别字典]}，用于第二档全量展开兜底"""
    out: Dict[str, List[Dict[str, Any]]] = {}
    for t1, t2, t3, code in _load_catalog():
        out.setdefault(t2, []).append({'level1': t1, 'level2': t2, 'level3': t3, 'code': code})
    return out


def catalog_by_t1() -> Dict[str, List[Dict[str, Any]]]:
    """{一级大类名: [该大类下的全部三级节点]}"""
    out: Dict[str, List[Dict[str, Any]]] = {}
    for t1, t2, t3, code in _load_catalog():
        out.setdefault(t1, []).append({'level1': t1, 'level2': t2, 'level3': t3, 'code': code})
    return out


@lru_cache(maxsize=8192)
def _lex_rows(nk: str, limit: int) -> tuple:
    if not os.path.exists(LEXICON_DB):
        return tuple(), 0
    con = sqlite3.connect(LEXICON_DB)
    try:
        rows = con.execute(
            "select t1,t2,t3,companies,rows from lex where nk=? order by companies desc, rows desc limit ?",
            (nk, limit)).fetchall()
        total = con.execute("select sum(companies) from lex where nk=?", (nk,)).fetchone()[0] or 1
    finally:
        con.close()
    return tuple(rows), total


_CODE_OF: Dict[Tuple[str, str, str], str] = {}


def _code_of() -> Dict[Tuple[str, str, str], str]:
    if not _CODE_OF:
        for t1, t2, t3, code in _load_catalog():
            _CODE_OF.setdefault((t1, t2, t3), code)
    return _CODE_OF


def lexicon_candidates(name: str, limit: int = 6) -> List[Dict[str, Any]]:
    """查广材词表 (category_lexicon.db)：返回 (一级,二级,三级,企业支持度)；同一个品名可能对应多个类别（分布）"""
    if not os.path.exists(LEXICON_DB):
        return []
    from name_norm import normalize_material_name
    nk = normalize_material_name(name)
    if not nk:
        return []
    rows, total = _lex_rows(nk, limit)
    code_of = _code_of()
    out = []
    for t1, t2, t3, companies, nrows in rows:
        code = code_of.get((t1, t2, t3))
        if not code:
            continue
        out.append({
            'level1': t1,
            'level2': t2,
            'level3': t3,
            'code': code,
            'share': round(companies / max(total, 1), 3),
            'companies': companies,
            'reason': f"广材词表 {companies}家企业({round(100*companies/max(total,1))}%)"
        })
    return out


@lru_cache(maxsize=4096)
def spec_lexicon_lookup(spec_str: str) -> Optional[Dict[str, Any]]:
    """
    查广材规格分类词库 (category_lexicon.db -> spec_lex)：
    根据物料的规格属性 JSON 串或核心属性签名，直达精确对应的标准工料机分类。
    """
    if not spec_str or not os.path.exists(LEXICON_DB):
        return None
    spec_clean = spec_str.strip()
    con = sqlite3.connect(LEXICON_DB)
    try:
        cur = con.cursor()
        # 1. 尝试全字匹配
        row = cur.execute(
            "SELECT t1, t2, t3, code, sample_unit, sample_brand, total_occurrences, unique_companies "
            "FROM spec_lex WHERE spec_json = ?", (spec_clean,)
        ).fetchone()
        if not row:
            # 2. 若传入是字典序列化有空格差异，尝试归一化 json 匹配
            try:
                parsed = json.loads(spec_clean)
                canon_json = json.dumps(parsed, ensure_ascii=False, separators=(',', ':'))
                row = cur.execute(
                    "SELECT t1, t2, t3, code, sample_unit, sample_brand, total_occurrences, unique_companies "
                    "FROM spec_lex WHERE spec_json = ?", (canon_json,)
                ).fetchone()
            except Exception:
                row = None
        if row:
            t1, t2, t3, code, unit, brand, occ, comp = row
            return {
                'level1': t1,
                'level2': t2,
                'level3': t3,
                'code': code,
                'unit': unit,
                'brand': brand,
                'occurrences': occ,
                'companies': comp,
                'confidence': 1.0,
                'reason': f"广材规格词库精准命中 ({comp}家企业/{occ}条记录验证)"
            }
    finally:
        con.close()
    return None



class JevCategoryScorer:
    """
    利用 Jev 模型 (choice + score 原语) 对工料机候选类别进行批量概率判别与匹配度打分。
    内置两档机制：
      - 档位 1 (精排消歧)：对词表 (category_lexicon.db) + BM25 + Sheet2 枚举召回的 Top-K 候选做带上下文的精确判别；
      - 档位 2 (全量二级树兜底)：当词表数据未覆盖该新词 (档位 1 选了 NONE 或低分) 时，
        利用 Jev 单次支持 255 个选项的能力，将该二级大类下的全部三级类别全量摊开由 Jev 闭卷直选。
    """

    def __init__(self):
        self.api_key = TYPESAFE_API_KEY
        self.base_url = TYPESAFE_BASE_URL.rstrip("/")
        self.model = TYPESAFE_MODEL or "jev-latest"
        self._session = requests.Session()

    @property
    def enabled(self) -> bool:
        return bool(self.api_key)

    def _call_jev(self, state: Dict[str, Any], questions: Dict[str, Any], timeout: int = 15) -> Optional[Dict[str, Any]]:
        if not self.enabled or not questions:
            return None
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": self.model,
            "state": state,
            "questions": questions,
        }
        try:
            resp = self._session.post(self.base_url, headers=headers, json=payload, timeout=timeout)
            if resp.status_code == 200:
                data = resp.json()
                return data.get("answers") or data.get("results") or data
        except Exception as e:
            print(f"[JevScorer Warning] API 调用异常: {e}")
        return None

    @staticmethod
    def _build_option_criterion(cand: Dict[str, Any]) -> str:
        """
        将单个候选类别组装为富语义判别标准 (解决同名三级类如“输入模块”无法区分的问题)：
        包含：【一级 > 二级 > 三级】+ 标准属性特征 + 广材词表真实企业支持度/召回依据
        """
        l1 = cand.get("level1", "")
        l2 = cand.get("level2", "")
        l3 = cand.get("level3", "")
        attrs = cand.get("allowed_attributes") or []
        attr_text = f" | 核心属性:[{', '.join(attrs[:6])}]" if attrs else ""
        reason = cand.get("reason", "")
        reason_text = f" | 证据:{reason}" if reason else ""
        return f"【{l1} > {l2} > {l3}】{attr_text}{reason_text}"

    def score_batch_items(
        self,
        items: List[Dict[str, Any]],
        sheet_context: Dict[str, Any],
        catalog_t2_map: Optional[Dict[str, List[Dict[str, Any]]]] = None,
    ) -> Dict[str, Dict[str, Any]]:
        """
        批量对整张表的物料进行 Jev 分类识别与打分。
        参数:
          - items: 每项含 {"key": 唯一键, "std_name": 品名, "raw_cat": 原表分类, "raw_spec": 规格, "candidates": 初步Top-K候选}
          - sheet_context: {"supplier": 供应商, "brand": 品牌, "dominant_t2": list(sheet_prior_t2), "sample_names": [...]}
          - catalog_t2_map: {二级类别名: [该二级类下所有三级类别字典]}，用于第二档“词表未覆盖”时的全量展开兜底
        返回:
          {item_key: {"code": 选中编码, "level1": l1, "level2": l2, "level3": l3, "prob": 概率, "quality_score": 打分, "confidence": 综合置信度, "tier": "tier1/tier2/fallback"}}
        """
        if not self.enabled or not items:
            return {}

        # 1. 按 (std_name, raw_cat, raw_spec[:40]) 去重，避免同表重复品名重复计费
        unique_map: Dict[str, Dict[str, Any]] = {}
        key_to_ukey: Dict[str, str] = {}
        for idx, it in enumerate(items):
            dedup_sig = f"{it.get('std_name', '')}|{it.get('raw_cat', '')}|{str(it.get('raw_spec', ''))[:40]}"
            if dedup_sig not in unique_map:
                ukey = f"u_{len(unique_map)}"
                unique_map[dedup_sig] = {"ukey": ukey, "item": it}
            key_to_ukey[it.get("key", str(idx))] = unique_map[dedup_sig]["ukey"]

        # 2. 构造第一档 Jev 批量请求 (Choice + Score 双原语协同)
        state_items = {}
        questions = {}
        ukey_to_cand_map: Dict[str, Dict[str, Dict[str, Any]]] = {}

        for u_info in unique_map.values():
            ukey = u_info["ukey"]
            it = u_info["item"]
            cands = it.get("candidates") or []
            if not cands:
                continue

            state_items[ukey] = {
                "material_name": it.get("std_name", ""),
                "raw_category_hint": it.get("raw_cat", ""),
                "raw_specification": it.get("raw_spec", ""),
            }

            criteria_dict = {}
            c_map = {}
            for c in cands[:15]:  # 第一档最多取前 15 个高潜候选
                code = c.get("code")
                if not code:
                    continue
                criteria_dict[code] = self._build_option_criterion(c)
                c_map[code] = c
            # 关键设计：必须加入 NONE 逃生槽，防止词表没给全时在错误候选里硬选
            criteria_dict["NONE"] = "以上候选类别均不符合该物料的真实物理品类（需触发全库二级树重新检索）"
            ukey_to_cand_map[ukey] = c_map

            questions[f"{ukey}_choice"] = {
                "type": "choice",
                "instructions": (
                    f"结合 sheet_context 整表主业背景与 items.{ukey} 的名称规格，"
                    "选出唯一最匹配的工料机标准分类编码。若所有候选均不匹配，必须选 NONE。"
                ),
                "criteria": criteria_dict,
            }
            questions[f"{ukey}_score"] = {
                "type": "score",
                "instructions": f"评估 items.{ukey} 与所选最佳工料机分类的语义及属性匹配质量：",
                "criteria": [
                    "0_完全不符: 品类错误或规格属性与该类严重冲突",
                    "1_仅大类相近: 所属专业大类正确，但细分三级品类不准确",
                    "2_品类正确: 三级品类准确，规格属性基本兼容",
                    "3_精准匹配: 品名、整表上下文、规格属性及广材词表支持度完全吻合",
                ],
            }

        state = {
            "sheet_context": sheet_context,
            "items": state_items,
        }
        raw_answers = self._call_jev(state, questions)
        if not raw_answers:
            return {}

        # 3. 解析第一档结果，并收集未命中/低置信度项进入【第二档：二级大类全量展开兜底】
        ukey_results: Dict[str, Dict[str, Any]] = {}
        tier2_ukeys: List[str] = []

        for u_info in unique_map.values():
            ukey = u_info["ukey"]
            ans_c = raw_answers.get(f"{ukey}_choice") or {}
            ans_s = raw_answers.get(f"{ukey}_score") or {}

            winner = str(ans_c.get("value") or ans_c.get("choice") or ans_c.get("answer") or "NONE")
            prob = float(ans_c.get("probability") or ans_c.get("confidence") or (ans_c.get("probabilities") or {}).get(winner, 0.0) or 0.0)
            q_score = float(ans_s.get("score") or ans_s.get("value") or prob)
            # 归一化 score 到 0~1
            if q_score > 1.0:
                q_score = min(q_score / 3.0, 1.0)
            combined_conf = round(0.6 * prob + 0.4 * q_score, 4)

            c_map = ukey_to_cand_map.get(ukey, {})
            if winner != "NONE" and winner in c_map and combined_conf >= JEV_MIN_CONFIDENCE:
                win_cand = c_map[winner]
                ukey_results[ukey] = {
                    "code": winner,
                    "level1": win_cand.get("level1", ""),
                    "level2": win_cand.get("level2", ""),
                    "level3": win_cand.get("level3", ""),
                    "prob": round(prob, 4),
                    "quality_score": round(q_score, 4),
                    "confidence": combined_conf,
                    "tier": "tier1_lex_rerank",
                }
            else:
                tier2_ukeys.append(ukey)

        # 4. 第二档：专门应对“词表数据还没给全”——将该表主流二级大类下的全部三级类别（<=200个）全量喂给 Jev 直选！
        if tier2_ukeys:
            # 注意：参数名不能叫 catalog_by_t2，否则会把模块级同名函数遮蔽掉（这里曾踩过）
            t2_catalog = catalog_t2_map or catalog_by_t2()
            dom_t2_list = sheet_context.get("dominant_t2") or []
            full_branch_cands = []
            for t2_name in dom_t2_list:
                full_branch_cands.extend(t2_catalog.get(t2_name, []))

            # 若没有主业先验或候选太多，补充常见二级大类
            if not full_branch_cands:
                for t2_cands in list(t2_catalog.values())[:3]:
                    full_branch_cands.extend(t2_cands)

            if full_branch_cands and len(full_branch_cands) <= 250:
                t2_questions = {}
                t2_cand_map = {c["code"]: c for c in full_branch_cands if c.get("code")}
                branch_criteria = {
                    c["code"]: f"【{c.get('level1','')} > {c.get('level2','')} > {c.get('level3','')}】"
                    for c in full_branch_cands if c.get("code")
                }
                branch_criteria["NONE"] = "不属于本二级专业大类下的任何品类"

                for ukey in tier2_ukeys:
                    t2_questions[f"{ukey}_branch_choice"] = {
                        "type": "choice",
                        "instructions": (
                            f"词表未直接覆盖 items.{ukey} 的俗称。请在本表所属专业大类的完整标准三级类别列表中，"
                            "根据物理本质选出最匹配的一项（若完全不属于本专业大类则选 NONE）："
                        ),
                        "criteria": branch_criteria,
                    }

                t2_answers = self._call_jev(state, t2_questions)
                if t2_answers:
                    for ukey in tier2_ukeys:
                        ans_b = t2_answers.get(f"{ukey}_branch_choice") or {}
                        win_b = str(ans_b.get("value") or ans_b.get("choice") or "NONE")
                        prob_b = float(ans_b.get("probability") or ans_b.get("confidence") or (ans_b.get("probabilities") or {}).get(win_b, 0.0) or 0.0)
                        if win_b != "NONE" and win_b in t2_cand_map and prob_b >= JEV_MIN_CONFIDENCE:
                            wc = t2_cand_map[win_b]
                            ukey_results[ukey] = {
                                "code": win_b,
                                "level1": wc.get("level1", ""),
                                "level2": wc.get("level2", ""),
                                "level3": wc.get("level3", ""),
                                "prob": round(prob_b, 4),
                                "quality_score": round(prob_b, 4),
                                "confidence": round(prob_b, 4),
                                "tier": "tier2_full_branch",
                            }

        # 5. 映射回每一行原始 item key，并记录数据自增长飞轮日志
        final_out = {}
        for orig_key, ukey in key_to_ukey.items():
            if ukey in ukey_results:
                final_out[orig_key] = ukey_results[ukey]

        # 数据飞轮日志
        try:
            os.makedirs(os.path.dirname(FLYWHEEL_LOG), exist_ok=True)
            with open(FLYWHEEL_LOG, 'a', encoding='utf-8') as f:
                for it in items:
                    k = it.get("key", "")
                    if k in final_out and final_out[k]["tier"] == "tier2_full_branch" and final_out[k]["confidence"] >= 0.8:
                        f.write(json.dumps({
                            "name": it.get("std_name", ""),
                            "code": final_out[k]["code"],
                            "level3": final_out[k]["level3"],
                            "confidence": final_out[k]["confidence"],
                        }, ensure_ascii=False) + '\n')
        except Exception:
            pass

        return final_out


def score_items(items: List[Dict[str, Any]],
                supplier: str = '', brand: str = '',
                sheet_prior_t2: Optional[Set[str]] = None,
                min_prob: float = None) -> Dict[str, Dict[str, Any]]:
    """便捷包装函数，对接 JevCategoryScorer"""
    scorer = JevCategoryScorer()
    if not scorer.enabled:
        return {}
    sheet_context = {
        "supplier": supplier,
        "brand": brand,
        "dominant_t2": sorted(sheet_prior_t2 or []),
    }
    return scorer.score_batch_items(items, sheet_context)
