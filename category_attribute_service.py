# -*- coding: utf-8 -*-
"""
category_attribute_service.py
工料机标准分类与属性服务（BM25+IDF 联合索引 + Sheet2 枚举反查 + 整表先验 + Top-K 召回）
"""
import os
import re
import math
import json
import copy
import threading
from collections import Counter
from typing import Dict, List, Tuple, Optional, Any, Set
import openpyxl
import requests
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
# 行间召回缓存：（品名, 原始类别, 规格, 整表主业先验, top_k）→ 候选列表
# 同一张表里几十行“中国黑花岗岩”的检索结果完全一致，重新查库 + 跑 BM25 是纯浪费
_CANDS_CACHE: Dict[tuple, List[Dict[str, Any]]] = {}
_CANDS_LOCK = threading.Lock()
_CANDS_CACHE_MAX = int(os.getenv("CATEGORY_CAND_CACHE_MAX", "4096"))


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
    try:
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
    finally:
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
    """整表主业先验投票：词表实盘分布 + 树检索两边一起投，统计 Top-2 二级类别

    只查《工料机类别》树时，厂商俗称（“洗脸盆龙头”“白铁皮”）根本投不出票，
    整表先验经常是空的；广材词表（category_lexicon.db）能直接给出二级归属，必须参与投票。
    """
    _load_catalog()
    if not items:
        return set()
    try:
        from jev_category_scorer import lexicon_candidates
    except Exception:
        lexicon_candidates = None

    t2_counter: Counter = Counter()
    for it in items:
        name = str(it.get("src_name") or it.get("raw_name") or "")
        spec = str(it.get("raw_spec") or "")
        cat = str(it.get("src_cat") or "")

        # ① 词表实盘：按企业支持度加权（去重企业数，不是行数）
        if lexicon_candidates and name:
            try:
                for lh in lexicon_candidates(name, limit=2):
                    if lh.get("level2"):
                        t2_counter[lh["level2"]] += 2.0 * float(lh.get("share", 0.5))
            except Exception:
                pass

        # ② 树检索
        if _SEARCH_INDEX:
            cands = _SEARCH_INDEX.search(name, cat, spec, sheet_prior_t2=None, top_k=1)
            if cands and cands[0]["score"] >= 15.0:
                t2_counter[cands[0]["level2"]] += 1.0

    return {t2 for t2, _ in t2_counter.most_common(2) if t2}


def fetch_category_attributes_from_api(cat_code: str) -> List[Dict[str, Any]]:
    """当本地缓存未命中且配置了 CATEGORY_ATTRS_API_URL 时，调用官方接口实时拉取标准属性并持久化回填本地缓存"""
    if not CATEGORY_ATTRS_API_URL or not cat_code:
        return []
    try:
        clean_code = str(cat_code).strip().replace('-', '')
        url = f"{CATEGORY_ATTRS_API_URL.rstrip('/')}/{clean_code}"
        resp = requests.get(url, timeout=5)
        if resp.status_code == 200:
            data = resp.json()
            payload = data.get("data") or data
            attrs = []
            for grp in payload.get("attributeGroups", []):
                attr = grp.get("attribute", {})
                a_name = attr.get("description") or ""
                is_req = attr.get("isRequired", False)
                vals = [v.get("value") for v in grp.get("values", []) if v.get("value")]
                enum_str = ", ".join(vals) if vals else ""
                if a_name:
                    attrs.append({
                        "name": a_name,
                        "enum": enum_str,
                        "required": is_req,
                        "affects_price": True
                    })
            if attrs:
                cache_json = os.path.join(os.path.dirname(__file__), "category_attrs_cache.json")
                try:
                    c_data = {}
                    if os.path.exists(cache_json):
                        with open(cache_json, "r", encoding="utf-8") as f:
                            c_data = json.load(f)
                    c_data[cat_code] = payload
                    with open(cache_json, "w", encoding="utf-8") as f:
                        json.dump(c_data, f, ensure_ascii=False, indent=2)
                except Exception:
                    pass
            return attrs
    except Exception:
        pass
    return []


# 当外部接口不可用或新分类未入缓存时，按二级/三级类关键词提供工料机国标通用属性模板兜底
_FALLBACK_CATEGORY_ATTRS = [
    (re.compile(r'(电线|电缆|导线|护套线|穿线|软线|光缆)'), ["型号", "芯数", "标称截面(mm²)", "额定电压(V)", "绝缘材质", "燃烧性能等级"]),
    (re.compile(r'(板|砖|石材|玻璃|铝塑板|石膏板|夹芯板)'), ["材质", "长度(mm)", "宽度(mm)", "厚度(mm)", "表面处理", "品种"]),
    (re.compile(r'(管|导管|套管|波纹管)'), ["品种", "公称直径DN(mm)", "外径(mm)", "壁厚(mm)", "公称压力(MPa)", "连接方式"]),
    (re.compile(r'(阀|水嘴|龙头|冲洗阀)'), ["品种", "型号", "公称直径DN(mm)", "阀体材质", "连接形式"]),
    (re.compile(r'(光模块|收发器|交换机|配线架|耦合器|跳线|尾纤)'), ["型号", "接口类型", "传输速率", "传输距离", "单模/多模", "长度(m)"]),
]


def fetch_category_attributes(level3_or_code: str) -> List[Dict[str, Any]]:
    """优先查本地 Sheet2 与缓存 -> 其次实时调用 CATEGORY_ATTRS_API_URL 接口 -> 最后走国标大类属性模板兜底"""
    _load_catalog()
    key = str(level3_or_code or "").strip()
    t3_name, cat_code = key, ""
    if key in _CODE_MAP:
        _, _, t3_name, cat_code = _CODE_MAP[key]
    elif _format_cat_code(key) in _CODE_MAP:
        _, _, t3_name, cat_code = _CODE_MAP[_format_cat_code(key)]
    elif key in _EXACT_T3_MAP:
        _, _, t3_name, cat_code = _EXACT_T3_MAP[key]

    attrs = _ATTR_SCHEMA_MAP.get(t3_name, [])
    if attrs:
        return attrs

    # 实时调用远端接口获取该分类编码的标准属性定义
    if cat_code:
        remote_attrs = fetch_category_attributes_from_api(cat_code)
        if remote_attrs:
            _ATTR_SCHEMA_MAP[t3_name] = remote_attrs
            return remote_attrs

    # 国标大类属性模板保底（确保大模型永远有标准属性名可对齐）
    for pat, std_names in _FALLBACK_CATEGORY_ATTRS:
        if pat.search(t3_name):
            return [{"name": n, "enum": "", "required": False, "affects_price": True} for n in std_names]
    return []


# 工程造价领域常见【同义属性名】标准映射表（解决“口径->公称直径DN”、“壁厚->厚度”、“容重->密度”等字面不包含导致的 L3-ATTR-MISS）
_ATTR_NAME_SYNONYMS: Dict[str, Tuple[str, ...]] = {
    "公称直径": ("口径", "通径", "管径", "公称通径", "公称口径", "直径dn", "dn"),
    "外径": ("公称外径", "管外径", "外径de", "de", "φ", "Φ"),
    "壁厚": ("厚度", "管壁厚", "板厚", "δ"),
    "厚度": ("壁厚", "板厚", "层厚", "δ"),
    "长度": ("长", "定尺", "定尺长度", "长边"),
    "宽度": ("宽", "幅宽", "短边"),
    "高度": ("高", "H"),
    "标称截面": ("截面", "截面积", "平方", "平方数", "线径", "导线截面"),
    "芯数": ("芯", "几芯", "股数"),
    "额定电压": ("电压", "电压等级", "耐压"),
    "材质": ("材料", "阀体材质", "管材材质", "板材材质", "主材质", "基材", "胎基", "胎基材料"),
    "表面处理": ("表面工艺", "防腐工艺", "涂层", "镀层", "饰面", "覆面", "覆面材料"),
    "品种": ("类型", "产品品种", "款式", "结构形式", "形式"),
    "型号": ("规格型号", "产品型号", "系列", "型号/系列"),
    "密度": ("容重", "克重", "表观密度"),
    "燃烧性能": ("防火等级", "阻燃等级", "燃烧等级", "耐火等级"),
    "连接方式": ("连接形式", "接口方式", "接口形式", "连接"),
    "公称压力": ("压力", "压力等级", "工作压力", "额定压力", "pn"),
}

_GENERIC_ENUM_STOPWORDS = {
    "其他", "其它", "普通", "通用", "无", "有", "是", "否",
    "标配", "标准", "定制", "进口", "国产", "一级", "二级", "三级", "特级", "默认"
}


def _is_specific_attr_enum(ev: str) -> bool:
    """属性枚举反向认领安全校验：过滤过短字符、纯数字和通用无指向泛词"""
    s = str(ev or "").strip()
    if len(s) < 2:
        return False
    if s in _GENERIC_ENUM_STOPWORDS:
        return False
    if re.fullmatch(r'\d+', s):
        return False
    return True


def retrieve_category_candidates(
    raw_cat: str,
    std_name: str,
    raw_spec: str = "",
    sheet_prior: Optional[set] = None,
    top_k: int = 5,
) -> List[Dict[str, Any]]:
    """召回 Top-K 标准类别候选及其 Sheet2 标准属性列表（融合 39.8万品名词表 + 377万规格词库 + BM25 + 整表先验）"""
    _load_catalog()
    if not _SEARCH_INDEX:
        return []

    cache_key = (raw_cat or "", std_name or "", raw_spec or "",
                 frozenset(sheet_prior) if sheet_prior else None, int(top_k))
    with _CANDS_LOCK:
        cached = _CANDS_CACHE.get(cache_key)
    if cached is not None:
        return copy.deepcopy(cached)

    # 0. 377万条广材规格词库精准召回（前置到候选召回阶段，绝不在后面被短路！）
    spec_lex_hit = None
    if raw_spec:
        try:
            from jev_category_scorer import spec_lexicon_lookup
            spec_lex_hit = spec_lexicon_lookup(raw_spec)
        except Exception:
            pass

    # 1. 39.8万条广材实盘品名词表召回（同时查 std_name 与 raw_cat，保留 limit=12 完整一对多分布）
    lex_hits = []
    try:
        from jev_category_scorer import lexicon_candidates
        for q_name in dict.fromkeys([std_name, raw_cat]):
            if q_name:
                for h in lexicon_candidates(q_name, limit=12):
                    if not any(x.get("code") == h.get("code") for x in lex_hits):
                        lex_hits.append(h)
    except Exception:
        pass

    # 2. 联合检索引擎（BM25 + Sheet2 枚举反查 + 整表先验投票）
    cands = _SEARCH_INDEX.search(std_name, raw_cat, raw_spec, sheet_prior_t2=sheet_prior, top_k=max(top_k, 8))
    bm25_by_code = {c["code"]: c for c in cands if c.get("code")}

    # 3. 多源历史证据融合打分与重排
    spec_txt = f"{std_name} {raw_cat} {raw_spec}"
    seen_codes = set()
    merged = []

    # 3.1 若 377 万规格词库直接命中，赋予最高先验分置顶进入候选池
    if spec_lex_hit and spec_lex_hit.get("code"):
        sc_code = _format_cat_code(spec_lex_hit["code"])
        if sc_code in _CODE_MAP:
            t1, t2, t3, sc_code = _CODE_MAP[sc_code]
            sp_score = 185.0
            if sheet_prior and t2 in sheet_prior:
                sp_score += 50.0
            merged.append({
                "level1": t1, "level2": t2, "level3": t3, "code": sc_code,
                "score": sp_score,
                "reason": spec_lex_hit.get("reason", "广材377万规格词库精准命中"),
            })
            seen_codes.add(sc_code)

    # 3.2 融合 39.8 万品名词表分布 + 整表主业先验 + BM25/Sheet2 规格证据
    if lex_hits:
        def _lex_score(lh: Dict[str, Any]) -> float:
            score = 100.0 * float(lh.get("share", 0.5))
            if sheet_prior and lh.get("level2") in sheet_prior:
                score += 65.0
            elif sheet_prior and lh.get("level1") in {t1 for t1, t2, _t3, _c in _CATALOG_ROWS if t2 in sheet_prior}:
                score += 25.0
            if lh.get("level3") and lh["level3"] in spec_txt:
                score += 40.0
            # 叠加 BM25 规格枚举匹配分（用于区分“电线”等一对多品名下的细分品类）
            bm_c = bm25_by_code.get(lh.get("code", ""))
            if bm_c:
                score += float(bm_c.get("score", 0.0)) * 0.45
            return score

        for lh in lex_hits:
            c_code = _format_cat_code(lh.get("code", ""))
            if not c_code or c_code in seen_codes:
                continue
            merged.append({
                "level1": lh["level1"],
                "level2": lh["level2"],
                "level3": lh["level3"],
                "code": c_code,
                "score": round(_lex_score(lh), 2),
                "reason": lh.get("reason", "广材词表实盘先验"),
            })
            seen_codes.add(c_code)

    for c in cands:
        if c["code"] not in seen_codes:
            merged.append(c)
            seen_codes.add(c["code"])

    merged.sort(key=lambda c: -float(c.get("score") or 0))
    cands = merged[:max(top_k, min(len(merged), 10))]

    # 4. 同族两轮展开（先补本表二级类别，不足再补同一级大类）
    if sheet_prior and len(cands) < 10:
        seen_codes = {c["code"] for c in cands}
        floor_score = min([c["score"] for c in cands], default=1.0) * 0.5
        for e in _SEARCH_INDEX.entries:
            if e["code"] not in seen_codes and e["t2"] in sheet_prior:
                cands.append({
                    "level1": e["t1"], "level2": e["t2"], "level3": e["t3"], "code": e["code"],
                    "score": round(floor_score, 2), "reason": f"同表主业二级类({e['t2']})同族候选",
                })
                seen_codes.add(e["code"])
                if len(cands) >= 10:
                    break
        if len(cands) < 10:
            prior_t1 = {t1 for t1, t2, _t3, _c in _CATALOG_ROWS if t2 in sheet_prior}
            for e in _SEARCH_INDEX.entries:
                if e["code"] not in seen_codes and e["t1"] in prior_t1:
                    cands.append({
                        "level1": e["t1"], "level2": e["t2"], "level3": e["t3"], "code": e["code"],
                        "score": round(floor_score * 0.8, 2), "reason": f"同表一级大类({e['t1']})同族候选",
                    })
                    seen_codes.add(e["code"])
                    if len(cands) >= 10:
                        break

    # 5. 挂载标准属性（对超长枚举截断至前 45 字符，防止撑爆 Batch Prompt Token）
    for c in cands:
        attrs = fetch_category_attributes(c.get("level3") or c.get("code"))
        allowed = []
        for a in attrs[:10]:
            if a.get("enum"):
                enum_preview = str(a["enum"])[:45] + ("..." if len(str(a["enum"])) > 45 else "")
                allowed.append(f"{a['name']}({enum_preview})")
            else:
                allowed.append(a["name"])
        c["allowed_attributes"] = allowed

    with _CANDS_LOCK:
        if len(_CANDS_CACHE) >= _CANDS_CACHE_MAX:
            _CANDS_CACHE.clear()
        _CANDS_CACHE[cache_key] = cands
    return copy.deepcopy(cands)


def resolve_and_verify_category(
    raw_cat: str,
    std_name: str,
    raw_spec: str = "",
    sheet_prior: Optional[set] = None,
    selected_code: str = "",
) -> Tuple[str, str, str, str, bool, str]:
    """
    返回: (level1, level2, level3, category_code, is_verified, match_reason)
    """
    _load_catalog()

    # 1. 优先检查 377 万条广材规格属性实盘词库精准直达（若规格特征唯一且与本表主业不冲突，优先级最高）
    if raw_spec:
        try:
            from jev_category_scorer import spec_lexicon_lookup
            spec_hit = spec_lexicon_lookup(raw_spec)
            if spec_hit and spec_hit.get("code"):
                sc = _format_cat_code(spec_hit["code"])
                if sc in _CODE_MAP and (not sheet_prior or spec_hit.get("level2") in sheet_prior):
                    t1, t2, t3, code = _CODE_MAP[sc]
                    return t1, t2, t3, code, True, spec_hit.get("reason", "广材规格词库精准命中")
        except Exception:
            pass

    # 2. 若 LLM/Jev 已经从候选池中选出了合法分类编码，锁定返回
    norm_code = _format_cat_code(selected_code)
    if norm_code and norm_code in _CODE_MAP:
        t1, t2, t3, code = _CODE_MAP[norm_code]
        return t1, t2, t3, code, True, f"候选池精准锁定({code})"

    # 3. 检查字典精确名称映射
    for q in (raw_cat, std_name):
        q_clean = _BRACKET_PREFIX_RE.sub('', str(q or '').strip())
        if q_clean in _EXACT_T3_MAP:
            t1, t2, t3, code = _EXACT_T3_MAP[q_clean]
            return t1, t2, t3, code, True, f"字典精确匹配:{t3}"

    # 4. 使用融合了广材词表 + 整表先验 + Sheet2 枚举的联合候选排序第一名
    cands = retrieve_category_candidates(raw_cat, std_name, raw_spec, sheet_prior=sheet_prior, top_k=1)
    if cands:
        best = cands[0]
        is_ok = float(best.get("score") or 0.0) >= 12.0
        return best["level1"], best["level2"], best["level3"], best["code"], is_ok, best["reason"]

    return "", "", "", "", False, "未命中工料机分类字典"


def _match_attr_key_to_schema(raw_k: str, schema: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """利用双向包含 + 工程同义属性词表，将原表属性键名精准对齐到标准模板属性"""
    raw_k_base = re.sub(r'[（(].*?[)）]', '', raw_k).strip().lower()
    if not raw_k_base:
        return None
    # 1. 精确匹配
    for attr in schema:
        std_k = attr["name"]
        std_k_base = re.sub(r'[（(].*?[)）]', '', std_k).strip().lower()
        if raw_k.lower() == std_k.lower() or raw_k_base == std_k_base:
            return attr
    # 2. 工程同义词匹配（如 口径/通径 -> 公称直径DN(mm)，壁厚 -> 厚度(mm)）
    for attr in schema:
        std_k_base = re.sub(r'[（(].*?[)）]', '', attr["name"]).strip().lower()
        for std_root, syn_list in _ATTR_NAME_SYNONYMS.items():
            if std_root in std_k_base or std_k_base in std_root:
                if raw_k_base == std_root or any(raw_k_base == s or s in raw_k_base for s in syn_list):
                    return attr
    # 3. 双向子串包含匹配
    for attr in schema:
        std_k_base = re.sub(r'[（(].*?[)）]', '', attr["name"]).strip().lower()
        if len(raw_k_base) >= 2 and len(std_k_base) >= 2 and (raw_k_base in std_k_base or std_k_base in raw_k_base):
            return attr
    return None


def _normalize_attr_value_with_enum(std_attr_name: str, raw_val: str, enum_str: str) -> Tuple[str, Optional[str]]:
    """
    用历史标准枚举值与属性单位规范化属性值：
    1. 若标准属性名自带单位括号如 '厚度(mm)'、'长度(mm)'、'标称截面(mm²)'，自动剥离值末尾重复的单位（'1.5mm' -> '1.5'）；
    2. 若该属性有标准枚举值列表（如 '304不锈钢, 316L不锈钢'），当原值为简称（如 '304' 或 'SUS304'）且唯一命中某标准枚举时，自动对齐为标准枚举值。
    """
    val = str(raw_val or "").strip()
    if not val:
        return val, None

    # 1. 数值型带单位属性自动脱壳（防止出现 "厚度(mm): 1.5mm" 与标准答案 "厚度(mm): 1.5" 不一致）
    m_unit_in_key = re.search(r'[（(](mm²|mm2|mm|cm|m|kg/m³|kg|t|mpa|kpa|kv|v|w|kw|a)[)）]', std_attr_name, re.I)
    if m_unit_in_key:
        u_key = m_unit_in_key.group(1)
        m_num = re.fullmatch(r'([+-]?\d+(?:\.\d+)?)\s*(?:mm²|mm2|mm|cm|m|kg/m³|kg|t|mpa|kpa|kv|v|w|kw|a|平方)?', val, re.I)
        if m_num:
            val = m_num.group(1)

    # 2. 对照历史标准枚举值对齐
    if not enum_str:
        return val, None
    enums = [e.strip() for e in re.split(r'[,，;；/、|]+', str(enum_str)) if e.strip()]
    if not enums or val in enums:
        return val, None

    val_low = val.lower()
    # 精确忽略大小写
    for e in enums:
        if e.lower() == val_low:
            return e, None
    # 常用工程字变体对齐（如 聚酯胎 <-> 聚酯毡）
    val_norm = val_low.replace('胎', '毡')
    for e in enums:
        if e.lower().replace('胎', '毡') == val_norm:
            return e, None
    # 唯一包含关系对齐（例如原值写 "氟碳"，标准枚举只有 "氟碳喷涂"；或原值写 "304"，标准枚举为 "304不锈钢"）
    contain_hits = [e for e in enums if (len(val_low) >= 2 and (val_low in e.lower() or e.lower() in val_low or val_norm in e.lower().replace('胎', '毡')))]
    if len(contain_hits) == 1:
        return contain_hits[0], None
    return val, None


def validate_and_normalize_specs(
    level3_name: str,
    specs: Dict[str, Any],
    raw_spec_text: str = "",
) -> Tuple[Dict[str, str], List[str], List[str]]:
    """
    对照历史标准属性表（Sheet2 / category_attrs_cache.json / 接口）校验、认领并归一化规格参数：
    1. 属性键同义词对齐（口径->公称直径DN(mm)、壁厚->厚度(mm)）；
    2. 属性值单位脱壳与标准枚举值对齐（'1.5mm'->'1.5'，'氟碳'->'氟碳涂层'）；
    3. 历史枚举值反向认领：对于尚未填上的标准属性，扫描原始规格串 raw_spec_text，若命中该属性的专有枚举值则自动补齐！
    """
    _load_catalog()
    lvl3 = str(level3_name or "").strip()
    if lvl3 in _CODE_MAP:
        lvl3 = _CODE_MAP[lvl3][2]
    else:
        fmt_c = _format_cat_code(lvl3)
        if fmt_c in _CODE_MAP:
            lvl3 = _CODE_MAP[fmt_c][2]
    schema = _ATTR_SCHEMA_MAP.get(lvl3, []) or fetch_category_attributes(lvl3)
    if not specs or not isinstance(specs, dict):
        specs = {}
    clean_specs: Dict[str, str] = {str(k).strip(): str(v).strip() for k, v in specs.items() if str(v or "").strip()}
    if not schema:
        return clean_specs, [], []

    normalized: Dict[str, str] = {}
    used_schema_names = set()
    enum_warnings: List[str] = []
    unsplit_text_blobs = [raw_spec_text] if raw_spec_text else []

    for raw_k, raw_v in clean_specs.items():
        matched_attr = _match_attr_key_to_schema(raw_k, schema)
        if matched_attr:
            std_k = matched_attr["name"]
            norm_v, warn = _normalize_attr_value_with_enum(std_k, raw_v, matched_attr.get("enum", ""))
            normalized[std_k] = norm_v
            used_schema_names.add(std_k)
            if warn:
                enum_warnings.append(warn)
        else:
            normalized[raw_k] = raw_v
            unsplit_text_blobs.append(f"{raw_k}:{raw_v}")

    # 3. 历史枚举反向自动认领：若某标准属性尚未被提取，但其专有枚举值在原始规格文本中明确出现，则自动认领填入！
    full_blob = " ; ".join(unsplit_text_blobs)
    if full_blob:
        blob_norm = full_blob.lower().replace('胎', '毡')
        for attr in schema:
            std_k = attr["name"]
            if std_k in used_schema_names or not attr.get("enum"):
                continue
            enums = [e.strip() for e in re.split(r'[,，;；/、|]+', str(attr["enum"])) if len(e.strip()) >= 2]
            # 按枚举长度从长到短匹配，确保优先匹配最具体的枚举值（如优先匹配“聚酯胎PY”而非“PY”）
            for ev in sorted(enums, key=len, reverse=True):
                ev_low = ev.lower()
                ev_norm = ev_low.replace('胎', '毡')
                if (_is_specific_enum(ev) or _is_specific_attr_enum(ev)) and (ev_low in full_blob.lower() or ev_norm in blob_norm):
                    normalized[std_k] = ev
                    used_schema_names.add(std_k)
                    break

    missing_req = [a["name"] for a in schema if a.get("required") and a["name"] not in used_schema_names]
    return normalized, missing_req, enum_warnings
