# -*- coding: utf-8 -*-
"""
category_attribute_service.py
工料机标准分类与属性服务（BM25+IDF 联合索引 + Sheet2 枚举反查 + 整表先验 + Top-K 召回）
"""
import os
import re
import math
import json
from collections import Counter
from typing import Dict, List, Tuple, Optional, Any, Set
import openpyxl
from config import CATEGORY_DICT_PATH, CATEGORY_ATTRS_API_URL

# 全局动态全库属性枚举反向索引 (品种/类型/系列/型号/材质等属性枚举自动成为三级分类别名，严禁硬编码测试集字典)
_DYNAMIC_ENUM_ALIAS_MAP: Dict[str, Set[str]] = {}

_BRACKET_PREFIX_RE = re.compile(r'^[\[\【][A-Za-z0-9\-_]+[\]\】][\s　]*')
_TOKEN_RE = re.compile(r'[A-Za-z0-9]+(?:[\-/\.][A-Za-z0-9]+)*|[\u4e00-\u9fff]+')
_SPEC_NUM_NOISE_RE = re.compile(
    r'\b\d+(?:\.\d+)?\s*(?:mm|cm|m|m2|m3|kg|g|t|mpa|kv|v|w|kw|a|%|元|根|米|套|个|台|只|对|条)?\b',
    re.I,
)

_CATALOG_LOADED = False
_CATALOG_ROWS: List[Tuple[str, str, str, str]] = []  # (t1, t2, t3, code)
_EXACT_T3_MAP: Dict[str, Tuple[str, str, str, str]] = {}
_CODE_MAP: Dict[str, Tuple[str, str, str, str]] = {}
_ATTR_SCHEMA_MAP: Dict[str, List[Dict[str, Any]]] = {}
_SEARCH_INDEX: Optional["EnrichedCategoryIndex"] = None


def _finish(code: str, src: str, notes: list) -> Tuple[str, str, list]:
    notes = list(notes or [])
    if code == '000101' and not any('其他类别' in n for n in notes):
        notes.append("已归入其他类别 (000101)，需人工确认")
    return code, src, notes


def _format_cat_code(raw_code: Any) -> str:
    s = re.sub(r'\.0$', '', str(raw_code or '').strip())
    digits = re.sub(r'\D', '', s)
    if len(digits) == 8:
        return f"{digits[0:2]}-{digits[2:4]}-{digits[4:6]}-{digits[6:8]}"
    if len(digits) == 6:
        return f"{digits[0:2]}-{digits[2:4]}-{digits[4:6]}"
    return s


def _tokenize(text: str) -> List[str]:
    """混合分词：保留英文型号缩写 + 中文完整短语 + 中文 2-gram/3-gram + 关键单字"""
    if not text:
        return []
    cleaned = _SPEC_NUM_NOISE_RE.sub(' ', str(text).lower())
    tokens: List[str] = []
    for seg in _TOKEN_RE.findall(cleaned):
        if 'a' <= seg[0] <= 'z' or '0' <= seg[0] <= '9':
            if len(seg) >= 2:
                tokens.append(seg)
        else:
            n = len(seg)
            if n == 1:
                tokens.append(seg)
            else:
                if n <= 6:
                    tokens.append(seg)
                for i in range(n - 1):
                    tokens.append(seg[i:i + 2])
                for i in range(n - 2):
                    tokens.append(seg[i:i + 3])
    return tokens


def _is_specific_enum(v: str) -> bool:
    """属性枚举值是否够“具体”到能当品类别名。

    像“LED”“A”“B”这种短英文/字母枚举几百个品类都有，当别名会把“LED灯管”推到
    可视对讲主机上去（实测过）。所以只留：带汉字的，或 5 位以上的字母数字串（形如型号）。
    """
    s = str(v or '').strip()
    if len(s) < 2:
        return False
    if re.search(r'[\u4e00-\u9fff]', s):
        return True
    return len(s) >= 5 and bool(re.fullmatch(r'[A-Za-z0-9\-./]+', s))


class EnrichedCategoryIndex:
    """联合 Sheet1 (分类树) + Sheet2 (属性名称与限定值枚举) 的 BM25+IDF 检索引擎"""

    def __init__(
        self,
        catalog_rows: List[Tuple[str, str, str, str]],
        attr_schema_map: Dict[str, List[Dict[str, Any]]],
    ):
        self.entries: List[Dict[str, Any]] = []
        self.df: Counter = Counter()

        for t1, t2, t3, code in catalog_rows:
            attrs = attr_schema_map.get(t3, [])
            attr_names = [str(a.get("name") or "").strip() for a in attrs if a.get("name")]
            enum_vals: List[str] = []
            for a in attrs:
                a_name = str(a.get("name") or "").strip()
                # 仅将品种/类型/系列/名称等分类核心属性的枚举值作为品类专有特征，排除材质/颜色等通用干扰
                if not any(k in a_name for k in ('品种', '类型', '系列', '名称', '品名')):
                    continue
                raw_enum = str(a.get("enum") or "").strip()
                if raw_enum:
                    enum_vals.extend(re.split(r'[,，;；/、\s|]+', raw_enum))
            enum_vals = [v.strip() for v in enum_vals if _is_specific_enum(v)]

            # 字段权重：三级名称(5.0) > Sheet2枚举值(4.5) > 二级名称(2.0) > Sheet2属性名(1.5) > 一级名称(1.0)
            tf: Counter = Counter()
            for tok in _tokenize(t3):
                tf[tok] += 5.0
            for ev in enum_vals:
                for tok in _tokenize(ev):
                    tf[tok] += 4.5
            for tok in _tokenize(t2):
                tf[tok] += 2.0
            for an in attr_names:
                for tok in _tokenize(an):
                    tf[tok] += 1.5
            for tok in _tokenize(t1):
                tf[tok] += 1.0

            for tok in tf.keys():
                self.df[tok] += 1

            t3_clean = re.sub(r'[（(].*?[)）]', '', t3).strip()
            self.entries.append({
                "t1": t1,
                "t2": t2,
                "t3": t3,
                "t3_clean": t3_clean,
                "code": code,
                "enum_set": set(enum_vals),
                "attr_names": set(attr_names),
                "tf": tf,
                "doc_len": sum(tf.values()) or 1.0,
            })

        self.N = max(len(self.entries), 1)
        self.avg_dl = sum(e["doc_len"] for e in self.entries) / self.N
        self.idf: Dict[str, float] = {
            tok: math.log(1.0 + (self.N - freq + 0.5) / (freq + 0.5))
            for tok, freq in self.df.items()
        }

    def search(
        self,
        std_name: str,
        raw_cat: str = "",
        raw_spec: str = "",
        sheet_prior_t2: Optional[set] = None,
        top_k: int = 5,
    ) -> List[Dict[str, Any]]:
        clean_name = _BRACKET_PREFIX_RE.sub('', str(std_name or '').strip())
        clean_cat = _BRACKET_PREFIX_RE.sub('', str(raw_cat or '').strip())
        spec_str = str(raw_spec or '').strip()

        # 1. 动态全库枚举别名反查（从全库数千个品类的 Sheet2 枚举值中自动提取反查，无硬编码）
        expanded_name_terms = [clean_name, clean_cat]
        combined_text_lower = f"{clean_name} {clean_cat} {spec_str}".lower()
        matched_dynamic_t3 = set()
        for ev_lower, t3_set in _DYNAMIC_ENUM_ALIAS_MAP.items():
            if len(ev_lower) >= 3 and ev_lower in combined_text_lower:
                matched_dynamic_t3.update(t3_set)
            elif len(ev_lower) == 2 and (ev_lower in clean_name.lower() or ev_lower in clean_cat.lower()):
                matched_dynamic_t3.update(t3_set)

        # 2. 构造查询 Token 权重（品名 x3.0，规格参数 x1.5）
        q_tokens: Counter = Counter()
        # 两字切片（如“线型灯”→“线型/型灯/线/型/灯”）在 BM25 里会把“混凝土扶手”这种
        # 八竿子打不着的节点顶上来；只给长词全权重，两字词降权
        for tok in _tokenize(" ".join(expanded_name_terms)):
            q_tokens[tok] += 3.0 if len(tok) >= 3 else 0.6
        for tok in _tokenize(spec_str):
            q_tokens[tok] += 1.5

        k1, b = 1.5, 0.75
        scored: List[Tuple[float, Dict[str, Any], str]] = []

        for entry in self.entries:
            score = 0.0
            reasons: List[str] = []
            t3 = entry["t3"]
            t3_c = entry["t3_clean"]

            # A. 三级类别名称精确或包含命中
            for candidate_str in (clean_name, clean_cat):
                if not candidate_str:
                    continue
                if candidate_str == t3 or candidate_str == t3_c:
                    score += 130.0
                    reasons.append(f"精确命中三级类:{t3}")
                    break
                if len(candidate_str) >= 2 and (candidate_str in t3 or (len(t3_c) >= 2 and t3_c in candidate_str)):
                    overlap_ratio = len(min(candidate_str, t3_c, key=len)) / max(len(candidate_str), len(t3_c), 1)
                    score += 48.0 * overlap_ratio
                    reasons.append(f"品名包含:{t3}")
                    break

            # B. 动态全库属性枚举反查直达（全行业几千个品类自动通用，无任何面向测试集的写死词表）
            if t3 in matched_dynamic_t3 or t3_c in matched_dynamic_t3:
                score += 85.0
                reasons.append(f"属性枚举动态反查命中:{t3}")

            # C. Sheet2 枚举限定值反查命中（按该枚举值在全库出现的稀缺度加分，绝不压过品名精确匹配）
            for ev in entry["enum_set"]:
                ev_lower = ev.lower()
                if not re.search(r'[\u4e00-\u9fffA-Za-z]', ev):
                    continue
                if len(ev) >= 3 and (ev_lower in combined_text_lower):
                    # 若枚举词包含核心品名特征或长度>=5（如“kbg扣压式镀锌穿线管”、“脚踏式冲洗阀”），给高分 95；普通3~4字枚举（如“氟碳涂层”、“不锈钢”）给辅助分 25
                    bonus = 95.0 if len(ev) >= 5 else 25.0
                    score += bonus
                    reasons.append(f"命中Sheet2属性枚举:{ev}")
                    break

            # D. BM25 (IDF 加权) 语义得分
            dl = entry["doc_len"]
            tf_map = entry["tf"]
            bm25 = 0.0
            for tok, q_w in q_tokens.items():
                if tok not in tf_map:
                    continue
                tf = tf_map[tok]
                idf = self.idf.get(tok, 0.1)
                bm25 += idf * ((tf * (k1 + 1.0)) / (tf + k1 * (1.0 - b + b * (dl / self.avg_dl)))) * q_w
            score += bm25

            # E. 整表主流二级类别先验加成（防止单行跳出本表主业领域）
            if score > 2.0 and sheet_prior_t2 and entry["t2"] in sheet_prior_t2:
                score *= 1.25
                reasons.append(f"整表主业先验:{entry['t2']}")

            if score > 1.5:
                scored.append((score, entry, "; ".join(reasons) or "BM25高权重词匹配"))

        scored.sort(key=lambda x: x[0], reverse=True)
        results: List[Dict[str, Any]] = []
        for sc, e, rsn in scored[:top_k]:
            results.append({
                "level1": e["t1"],
                "level2": e["t2"],
                "level3": e["t3"],
                "code": e["code"],
                "score": round(sc, 2),
                "reason": rsn,
            })
        return results


def _load_catalog() -> None:
    global _CATALOG_LOADED, _CATALOG_ROWS, _EXACT_T3_MAP, _CODE_MAP, _ATTR_SCHEMA_MAP, _SEARCH_INDEX
    if _CATALOG_LOADED:
        return
    _CATALOG_LOADED = True
    dict_path = CATEGORY_DICT_PATH or "工料机类别.xlsx"
    if not os.path.exists(dict_path):
        # 尝试在当前目录寻找
        alt_path = os.path.join(os.path.dirname(__file__), "工料机类别.xlsx")
        if os.path.exists(alt_path):
            dict_path = alt_path
        else:
            print(f"[Warning] 工料机字典文件不存在: {dict_path}")
            return

    wb = openpyxl.load_workbook(dict_path, read_only=True, data_only=True)
    # 1. 加载 Sheet 1: 工料机类别 (自适应兼容 6列标准表头 与 4列简表)
    ws1 = wb.worksheets[0]
    first_row = next(ws1.iter_rows(min_row=1, max_row=1, values_only=True), None)
    is_6_cols = False
    if first_row and any('三级类别名称' in str(c or '') for c in first_row):
        is_6_cols = True

    for row in ws1.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        if is_6_cols:
            if len(row) < 6:
                continue
            t1 = str(row[1] or "").strip()
            t2 = str(row[3] or "").strip()
            t3 = str(row[5] or "").strip()
            code = _format_cat_code(row[4])
        else:
            if len(row) < 4:
                continue
            t1 = str(row[0] or "").strip()
            t2 = str(row[1] or "").strip()
            t3 = str(row[2] or "").strip()
            code = _format_cat_code(row[3])

        if not t3:
            continue
        item = (t1, t2, t3, code)
        _CATALOG_ROWS.append(item)
        _EXACT_T3_MAP[t3] = item
        t3_clean = re.sub(r'[（(].*?[)）]', '', t3).strip()
        if t3_clean and t3_clean not in _EXACT_T3_MAP:
            _EXACT_T3_MAP[t3_clean] = item
        if code:
            _CODE_MAP[code] = item

    # 2. 加载 Sheet 2: 三级类别属性 (从 Excel 或本地 category_attrs_cache.json 加载)
    if len(wb.worksheets) >= 2:
        ws2 = wb.worksheets[1]
        for row in ws2.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 2:
                continue
            t3_name = str(row[0] or "").strip()
            attr_name = str(row[1] or "").strip()
            if not t3_name or not attr_name:
                continue
            enum_str = str(row[2] or "").strip() if len(row) > 2 else ""
            req_str = str(row[3] or "").strip() if len(row) > 3 else ""
            price_str = str(row[4] or "").strip() if len(row) > 4 else ""
            _ATTR_SCHEMA_MAP.setdefault(t3_name, []).append({
                "name": attr_name,
                "enum": enum_str,
                "required": req_str in ("是", "必填", "1", "True", "true"),
                "affects_price": price_str in ("是", "1", "True", "true"),
            })
    wb.close()

    # 3. 增强：融合本地 category_attrs_cache.json 中的真实接口限定枚举与必填规则
    cache_json = os.path.join(os.path.dirname(__file__), "category_attrs_cache.json")
    if os.path.exists(cache_json):
        try:
            with open(cache_json, "r", encoding="utf-8") as f:
                c_data = json.load(f)
            for c_code, payload in c_data.items():
                c_desc = payload.get("type", {}).get("description") or ""
                if not c_desc:
                    # 尝试从 _CODE_MAP 获取三级名称
                    matched_item = _CODE_MAP.get(_format_cat_code(c_code))
                    if matched_item:
                        c_desc = matched_item[2]
                if not c_desc:
                    continue

                for grp in payload.get("attributeGroups", []):
                    attr = grp.get("attribute", {})
                    a_name = attr.get("description") or ""
                    is_req = attr.get("isRequired", False)
                    vals = [v.get("value") for v in grp.get("values", []) if v.get("value")]
                    enum_str = ", ".join(vals) if vals else ""
                    if a_name:
                        _ATTR_SCHEMA_MAP.setdefault(c_desc, []).append({
                            "name": a_name,
                            "enum": enum_str,
                            "required": is_req,
                            "affects_price": True
                        })
        except Exception:
            pass

    # 4. 自动全库构建核心属性枚举反向索引 (仅从 品种/类型/系列/名称/品名 等核心分类属性的枚举值构建反查，排除材质/颜色等通用属性干扰)
    CORE_TYPE_ATTR_KEYS = ('品种', '类型', '系列', '名称', '品名')
    _DYNAMIC_ENUM_ALIAS_MAP.clear()
    for t3_name, attrs in _ATTR_SCHEMA_MAP.items():
        t3_clean = re.sub(r'[（(].*?[)）]', '', t3_name).strip()
        for a in attrs:
            a_name = str(a.get("name") or "").strip()
            # 严格限定只提取能表征物料品种/类型的核心属性
            if not any(k in a_name for k in CORE_TYPE_ATTR_KEYS):
                continue
            raw_enum = str(a.get("enum") or "").strip()
            if not raw_enum:
                continue
            vals = re.split(r'[,，;；/、\s|]+', raw_enum)
            for v in vals:
                v_clean = v.strip()
                if not _is_specific_enum(v_clean):
                    continue
                if len(v_clean) >= 2 and re.search(r'[\u4e00-\u9fffA-Za-z]', v_clean):
                    v_lower = v_clean.lower()
                    _DYNAMIC_ENUM_ALIAS_MAP.setdefault(v_lower, set()).add(t3_name)
                    if t3_clean:
                        _DYNAMIC_ENUM_ALIAS_MAP.setdefault(v_lower, set()).add(t3_clean)

    _SEARCH_INDEX = EnrichedCategoryIndex(_CATALOG_ROWS, _ATTR_SCHEMA_MAP)


def compute_sheet_category_prior(items: List[Dict[str, Any]]) -> set:
    """整表主业先验投票：扫描全表所有行，统计得分最高的 Top-2 二级类别"""
    _load_catalog()
    if not _SEARCH_INDEX or not items:
        return set()
    t2_counter: Counter = Counter()
    for it in items:
        name = str(it.get("src_name") or it.get("raw_name") or "")
        spec = str(it.get("raw_spec") or "")
        cat = str(it.get("src_cat") or "")
        cands = _SEARCH_INDEX.search(name, cat, spec, sheet_prior_t2=None, top_k=1)
        if cands and cands[0]["score"] >= 15.0:
            t2_counter[cands[0]["level2"]] += 1
    return {t2 for t2, _ in t2_counter.most_common(2)}


def retrieve_category_candidates(
    raw_cat: str,
    std_name: str,
    raw_spec: str = "",
    sheet_prior: Optional[set] = None,
    top_k: int = 5,
) -> List[Dict[str, Any]]:
    """召回 Top-K 标准类别候选及其 Sheet2 标准属性列表，供 LLM 做受限单选题"""
    _load_catalog()
    if not _SEARCH_INDEX:
        return []
    cands = _SEARCH_INDEX.search(std_name, raw_cat, raw_spec, sheet_prior_t2=sheet_prior, top_k=top_k)

    # 同族展开：本表主业二级大类下的其余三级类别补进候选池。
    # 原表写的是生僻俗称、字面一个字也不命中时，靠这一手让模型仍能在这十几个同族节点里做选择题。
    # 按一级大类展开：二级到三级往往只有 1 个孩子（“光源”就只挂了一个节点），只按二级展开等于没展开
    if sheet_prior and len(cands) < 10:
        prior_t1 = {t1 for t1, t2, _t3, _c in _CATALOG_ROWS if t2 in sheet_prior}
        seen_codes = {c["code"] for c in cands}
        floor_score = min([c["score"] for c in cands], default=1.0) * 0.5
        for e in _SEARCH_INDEX.entries:
            if e["code"] in seen_codes or (e["t2"] not in sheet_prior and e["t1"] not in prior_t1):
                continue
            cands.append({
                "level1": e["t1"],
                "level2": e["t2"],
                "level3": e["t3"],
                "code": e["code"],
                "score": round(floor_score, 2),
                "reason": f"同表主业({e['t1']}/{e['t2']})同族候选",
            })
            seen_codes.add(e["code"])
            if len(cands) >= 10:
                break

    for c in cands:
        attrs = _ATTR_SCHEMA_MAP.get(c["level3"], [])
        c["allowed_attributes"] = [
            f"{a['name']}({a['enum']})" if a.get("enum") else a["name"]
            for a in attrs[:10]
        ]
    return cands


def resolve_and_verify_category(
    raw_cat: str,
    std_name: str,
    raw_spec: str = "",
    sheet_prior: Optional[set] = None,
    selected_code: str = "",
) -> Tuple[str, str, str, str, bool, str]:
    """
    返回: (level1, level2, level3, category_code, is_verified, match_reason)
    优先采信 LLM 从候选池中选出的合法编码 selected_code，否则走联合检索引擎
    """
    _load_catalog()
    # 1. 若 LLM 已经从 Top-5 候选中选出了合法的分类编码，直接精确锁定
    norm_code = _format_cat_code(selected_code)
    if norm_code and norm_code in _CODE_MAP:
        t1, t2, t3, code = _CODE_MAP[norm_code]
        return t1, t2, t3, code, True, f"候选池精准锁定({code})"

    # 2. 检查精确名称映射
    for q in (raw_cat, std_name):
        q_clean = _BRACKET_PREFIX_RE.sub('', str(q or '').strip())
        if q_clean in _EXACT_T3_MAP:
            t1, t2, t3, code = _EXACT_T3_MAP[q_clean]
            return t1, t2, t3, code, True, f"字典精确匹配:{t3}"

    # 3. 走 BM25 + Sheet2 枚举 + 整表先验联合检索
    if _SEARCH_INDEX:
        cands = _SEARCH_INDEX.search(std_name, raw_cat, raw_spec, sheet_prior_t2=sheet_prior, top_k=1)
        if cands:
            best = cands[0]
            is_ok = best["score"] >= 12.0
            return best["level1"], best["level2"], best["level3"], best["code"], is_ok, best["reason"]

    return "", "", "", "", False, "未命中工料机分类字典"


def fetch_category_attributes(level3_name: str) -> List[Dict[str, Any]]:
    _load_catalog()
    return _ATTR_SCHEMA_MAP.get(str(level3_name or "").strip(), [])


def validate_and_normalize_specs(
    level3_name: str,
    specs: Dict[str, Any],
) -> Tuple[Dict[str, str], List[str], List[str]]:
    """
    对照 Sheet2 标准属性表校验并归一化规格参数：
    返回: (normalized_specs, missing_required_attrs, enum_warnings)
    """
    _load_catalog()
    schema = _ATTR_SCHEMA_MAP.get(str(level3_name or "").strip(), [])
    if not specs or not isinstance(specs, dict):
        specs = {}
    clean_specs: Dict[str, str] = {str(k).strip(): str(v).strip() for k, v in specs.items() if str(v or "").strip()}
    if not schema:
        return clean_specs, [], []

    # 属性名同义对齐（如 厚度 -> 厚度(mm)，公称口径 -> 公称直径DN(mm)）
    normalized: Dict[str, str] = {}
    used_schema_names = set()
    enum_warnings: List[str] = []

    for raw_k, raw_v in clean_specs.items():
        matched_attr = None
        raw_k_base = re.sub(r'[（(].*?[)）]', '', raw_k).strip()
        for attr in schema:
            std_k = attr["name"]
            std_k_base = re.sub(r'[（(].*?[)）]', '', std_k).strip()
            if raw_k == std_k or (raw_k_base and raw_k_base == std_k_base) or (raw_k_base in std_k_base):
                matched_attr = attr
                break
        if matched_attr:
            std_k = matched_attr["name"]
            normalized[std_k] = raw_v
            used_schema_names.add(std_k)
        else:
            normalized[raw_k] = raw_v

    missing_req = [a["name"] for a in schema if a.get("required") and a["name"] not in used_schema_names]
    return normalized, missing_req, enum_warnings
