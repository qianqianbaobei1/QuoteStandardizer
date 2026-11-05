# -*- coding: utf-8 -*-
"""
jev_category_scorer.py - 双档位 Jev 分类判定（词表先验精排 + 全分支展开兜底）

档位 1：把「广材词表分布 + 树检索 + Sheet2 标准属性」拼成富语义选项，Jev 一次 choice 选一个，
        内置 NONE 逃生槽；再配一个 score 评估匹配质量。
档位 2：档位 1 选 NONE 或低分（= 词表还没覆盖的新叫法）时，把该表主业**一级大类**下
        的全部三级节点摊开给 Jev 闭卷直选（choice 上限 255 项，够用）。

API 契约以官方 OpenAPI 为准：
  choice 答案 = {type, choice, confidence, probabilities}   ← 没有 value / probability 字段
  score  答案 = {type, score, confidence, legend, probabilities}  ← score 是 0..(档数-1) 的期望值
  choice.criteria 上限 255 项；score.criteria 上限 10 档

注意：选项池里没有正确类别时，模型无法选中它（"候选覆盖"是前提）——
所以档位 2 用一级大连同族全量展开，而不是再靠字面检索一次。
"""
import json
import os
import sqlite3
from typing import Any, Dict, List, Optional, Set, Tuple

from config import JEV_MIN_CONFIDENCE

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LEXICON_DB = os.path.join(BASE_DIR, 'category_lexicon.db')
FLYWHEEL_LOG = os.path.join(BASE_DIR, 'data', 'lexicon_flywheel.jsonl')

CHUNK_ITEMS = 12          # 每请求最多 12 个物料（12 choice + 12 score = 24 问，留出余量）
MAX_T1_OPTIONS = 255      # choice 硬上限


def _load_catalog():
    from category_attribute_service import _load_catalog as _lc, _CATALOG_ROWS
    _lc()
    return _CATALOG_ROWS


def catalog_by_t1() -> Dict[str, List[Dict[str, str]]]:
    """{一级大类名: [该大一下的全部三级节点]}，档位 2 的选项池"""
    out: Dict[str, List[Dict[str, str]]] = {}
    for t1, t2, t3, code in _load_catalog():
        out.setdefault(t1, []).append({'level1': t1, 'level2': t2, 'level3': t3, 'code': code})
    return out


def t1_of(t2_names: Set[str]) -> List[str]:
    """二级主业先验 → 一级大类（二级到三级常常只有 1 个孩子，必须升到一级才有候选）"""
    hit = []
    for t1, t2, _t3, _code in _load_catalog():
        if t2 in t2_names and t1 not in hit:
            hit.append(t1)
    return hit


def lexicon_candidates(name: str, limit: int = 6) -> List[Dict[str, Any]]:
    """查广材词表：返回 (一级,二级,三级,企业支持度)；同一个品名可能对多个类别（分布）"""
    if not os.path.exists(LEXICON_DB):
        return []
    from name_norm import normalize_material_name
    nk = normalize_material_name(name)
    if not nk:
        return []
    con = sqlite3.connect(LEXICON_DB)
    try:
        rows = con.execute(
            "select t1,t2,t3,companies,rows from lex where nk=? order by companies desc, rows desc limit ?",
            (nk, limit)).fetchall()
        total = con.execute("select sum(companies) from lex where nk=?", (nk,)).fetchone()[0] or 1
    finally:
        con.close()
    code_of = {(t1, t2, t3): code for t1, t2, t3, code in _load_catalog()}
    out = []
    for t1, t2, t3, companies, nrows in rows:
        code = code_of.get((t1, t2, t3))
        if not code:
            continue
        out.append({'level1': t1, 'level2': t2, 'level3': t3, 'code': code,
                    'share': round(companies / max(total, 1), 3),
                    'reason': f"广材词表 {companies}家企业({round(100*companies/max(total,1))}%)"})
    return out


def _attrs_brief(code_or_name: str, limit: int = 5) -> str:
    from category_attribute_service import _ATTR_SCHEMA_MAP
    from category_attribute_service import _CODE_MAP
    t3 = code_or_name
    if code_or_name in _CODE_MAP:
        t3 = _CODE_MAP[code_or_name][2]
    attrs = _ATTR_SCHEMA_MAP.get(t3, [])
    names = []
    for a in attrs[:limit]:
        n = str(a.get('name') or '')
        if not n:
            continue
        enum = str(a.get('enum') or '').split(',')[0].strip()
        names.append(f"{n}(如{enum})" if enum else n)
    return '、'.join(names)


def _criterion(c: Dict[str, Any]) -> str:
    path = f"{c['level1']} > {c['level2']} > {c['level3']}"
    attrs = _attrs_brief(c['code'])
    parts = [f"【{path}】"]
    if attrs:
        parts.append(f"标准属性: {attrs}")
    if c.get('reason'):
        parts.append(f"证据: {c['reason']}")
    return ' | '.join(parts)


def _judge():
    from llm_agent import JevJudge
    j = JevJudge()
    return j if j.available else None


def score_items(items: List[Dict[str, Any]],
                supplier: str = '', brand: str = '',
                sheet_prior_t2: Optional[Set[str]] = None,
                min_prob: float = None) -> Dict[str, Dict[str, Any]]:
    """
    items: [{'key','name','raw_cat','raw_spec','candidates'(可选)}]
    返回: {key: {'code','level1','level2','level3','prob','quality','tier','reason'}}
    """
    judge = _judge()
    if not judge or not items:
        return {}
    min_prob = JEV_MIN_CONFIDENCE if min_prob is None else min_prob
    by_t1 = catalog_by_t1()

    # 去重 + 组装档位 1 候选
    uniq: Dict[str, Dict[str, Any]] = {}
    key2u: Dict[str, str] = {}
    for it in items:
        sig = f"{it.get('name','')}|{it.get('raw_cat','')}"
        if sig not in uniq:
            cands = it.get('candidates') or []
            merged = {c['code']: c for c in lexicon_candidates(it.get('name', ''))}
            for c in cands:                       # 词表优先级更高：同名 code 时保留词表那条
                merged.setdefault(c.get('code'), c)
            uniq[sig] = {'ukey': f"u{len(uniq)}", 'item': it, 'cands': list(merged.values())[:15]}
        key2u[it['key']] = uniq[sig]['ukey']

    results: Dict[str, Dict[str, Any]] = {}
    todo = list(uniq.values())

    # ---------- 档位 1 ----------
    for start in range(0, len(todo), CHUNK_ITEMS):
        chunk = todo[start:start + CHUNK_ITEMS]
        state_items, questions, cand_map = {}, {}, {}
        for u in chunk:
            cands = u['cands']
            if not cands:
                continue
            it = u['item']
            state_items[u['ukey']] = {
                'material_name': it.get('name', ''),
                'source_category_hint': it.get('raw_cat', ''),
                'specification': str(it.get('raw_spec', ''))[:300],
            }
            crit = {c['code']: _criterion(c) for c in cands if c.get('code')}
            crit['NONE'] = '以上候选都与该物料物理品类不符（词表未覆盖该叫法）'
            cand_map[u['ukey']] = {c['code']: c for c in cands if c.get('code')}
            questions[f"{u['ukey']}_cat"] = {
                'type': 'choice',
                'instructions': f"结合整表专业背景与 items.{u['ukey']} 的品名、规格，选出唯一匹配的工料机三级类别；都不符就选 NONE。",
                'criteria': crit,
            }
            questions[f"{u['ukey']}_q"] = {
                'type': 'score',
                'instructions': f"评估 items.{u['ukey']} 与所选类别在品名、规格、专业背景上的一致程度：",
                'criteria': ['完全不符', '仅大类相近', '三级正确但规格有出入', '精准匹配'],
            }
        if not questions:
            continue
        ans = judge.ask({'sheet_context': {'supplier': supplier, 'brand': brand,
                                           'dominant_t2': sorted(sheet_prior_t2 or [])},
                         'items': state_items}, questions, timeout=15) or {}
        for u in chunk:
            uk = u['ukey']
            a = ans.get(f"{uk}_cat") or {}
            winner = a.get('choice')
            prob = float((a.get('probabilities') or {}).get(winner, 0) or 0)
            sc = float((ans.get(f"{uk}_q") or {}).get('score') or 0) / 3.0
            conf = round(0.6 * prob + 0.4 * sc, 3)
            if winner and winner != 'NONE' and winner in cand_map.get(uk, {}) and prob >= min_prob:
                c = cand_map[uk][winner]
                results[uk] = {'code': c['code'], 'level1': c['level1'], 'level2': c['level2'],
                               'level3': c['level3'], 'prob': round(prob, 3), 'quality': round(sc, 3),
                               'confidence': conf, 'tier': 'tier1_lexicon_rerank',
                               'reason': c.get('reason', '')}

    # ---------- 档位 2：一级大类全量展开 ----------
    rest = [u for u in todo if u['ukey'] not in results]
    if rest:
        t1_list = t1_of(sheet_prior_t2 or set())
        if not t1_list and sheet_prior_t2:
            t1_list = []
        # 没有明确主业先验时，先让 Jev 从全部一级大类里挑（<=255 项）
        if not t1_list:
            options = {t1: f"一级大类：{t1}" for t1 in list(by_t1)[:MAX_T1_OPTIONS]}
            a1 = judge.ask({'items': {u['ukey']: {'material_name': u['item'].get('name', ''),
                                                  'specification': str(u['item'].get('raw_spec', ''))[:200]}
                                      for u in rest[:CHUNK_ITEMS]}},
                           {f"{u['ukey']}_t1": {'type': 'choice',
                                                'instructions': f"items.{u['ukey']} 属于哪个一级工料机大类？",
                                                'criteria': options} for u in rest[:CHUNK_ITEMS]},
                           timeout=15) or {}
            t1_list = sorted({(a1.get(f"{u['ukey']}_t1") or {}).get('choice') for u in rest[:CHUNK_ITEMS]} - {None})
        branch: List[Dict[str, Any]] = []
        for t1 in t1_list:
            branch.extend(by_t1.get(t1, []))
        branch = branch[:MAX_T1_OPTIONS]
        if branch:
            crit = {c['code']: _criterion(c) for c in branch}
            crit['NONE'] = '不属于该一级大类下的任何品类'
            code2c = {c['code']: c for c in branch}
            for start in range(0, len(rest), CHUNK_ITEMS):
                chunk = rest[start:start + CHUNK_ITEMS]
                state = {'sheet_context': {'supplier': supplier, 'brand': brand,
                                           'dominant_t1': t1_list},
                         'items': {u['ukey']: {'material_name': u['item'].get('name', ''),
                                               'source_category_hint': u['item'].get('raw_cat', ''),
                                               'specification': str(u['item'].get('raw_spec', ''))[:300]}
                                   for u in chunk}}
                qs = {f"{u['ukey']}_branch": {
                        'type': 'choice',
                        'instructions': f"词表未覆盖 items.{u['ukey']} 的叫法。请在本表所属一级大类的完整三级类别里按物理本质选出最匹配的一项；不属于则选 NONE。",
                        'criteria': crit} for u in chunk}
                ans2 = judge.ask(state, qs, timeout=20) or {}
                for u in chunk:
                    a = ans2.get(f"{u['ukey']}_branch") or {}
                    winner = a.get('choice')
                    prob = float((a.get('probabilities') or {}).get(winner, 0) or 0)
                    if winner and winner != 'NONE' and winner in code2c and prob >= min_prob:
                        c = code2c[winner]
                        results[u['ukey']] = {'code': c['code'], 'level1': c['level1'],
                                              'level2': c['level2'], 'level3': c['level3'],
                                              'prob': round(prob, 3), 'quality': '', 'confidence': round(prob, 3),
                                              'tier': 'tier2_full_branch', 'reason': f"一级大类 {c['level1']} 全量展开"}

    # 映射回行 key，并给高分 tier2 结果留飞轮日志
    out: Dict[str, Dict[str, Any]] = {}
    for key, uk in key2u.items():
        if uk in results:
            out[key] = results[uk]
    try:
        os.makedirs(os.path.dirname(FLYWHEEL_LOG), exist_ok=True)
        with open(FLYWHEEL_LOG, 'a', encoding='utf-8') as f:
            for key, r in out.items():
                if r['tier'] == 'tier2_full_branch' and r['confidence'] >= 0.8:
                    name = next((it.get('name') for it in items if it['key'] == key), '')
                    f.write(json.dumps({'name': name, 'code': r['code'], 'level3': r['level3'],
                                        'confidence': r['confidence']}, ensure_ascii=False) + '\n')
    except Exception:
        pass
    return out
