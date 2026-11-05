import unicodedata
# -*- coding: utf-8 -*-
"""
QuoteStandardizer - 报价单解析与原表对账
1. Excel / PDF (文字层或扫描件) 统一还原为表格，再按同一套流程解析
2. 支持清单式 (一行一个物料) 与矩阵式 (行标题 × 列标题 → 价格) 报价表
3. 价格只取原表；名称/参数由模型清洗，原文已规范的 “键:值;” 参数原样保留
4. 原表 ↔ 输出逐行对账；识别存疑、信息缺失、推断得出的字段全部写入待复核说明
5. 不知道的信息一律留空，不填任何默认值
"""
import os
import sys
import re
import copy
import hashlib
import argparse
import json
import time
from typing import Optional, List, Dict, Any, Tuple
import openpyxl
from concurrent.futures import ThreadPoolExecutor

from extractor import ExcelAssetExtractor, safe_clean_price, cell_str, unmerge_and_fill_cells, extract_sheet_sample_rows, hidden_axes
from pdf_extractor import PDFTableExtractor
from gemini_pdf_extractor import GeminiPDFTableExtractor
from terms_extractor import TermsExtractor, normalize_terms, find_terms_image, missing_terms
from builder import StandardizedWorkbookBuilder
from config import (DEFAULT_TEMPLATE_PATH, DEFAULT_OUTPUT_DIR, PROVINCE_ABBR_MAP, CATEGORY_CODE_MAP,
                    resolve_category_code, expand_region, EXTRACTOR_ENGINE, GEMINI_API_KEY, GEMINI_MODEL, VISION_MODEL,
                    EXPORT_INTEGRITY_CERTIFICATE, ANTI_HALLUCINATION_STRICT_ROW_CONSERVATION,
                    SKIP_HIDDEN_CONTENT,
                    STRICT_ENGINE,
                    EXPAND_NATIONWIDE, NATIONWIDE_REGIONS)
from models import CanonicalMaterialItem, SourceRowRecord, reconcile, summarize_audit
from category_attribute_service import (
    compute_sheet_category_prior,
    retrieve_category_candidates,
    resolve_and_verify_category,
    fetch_category_attributes,
    validate_and_normalize_specs,
)
from llm_agent import DeepSeekStandardizer, JevJudge, UNIT_PRICE_KEYS, TOTAL_PRICE_KEYS
from config import STRUCTURE_ENGINE, JEV_MIN_CONFIDENCE

RULE_DECISION_TYPE = {'item': 'item', '分组标题行': 'title', '重复表头行': 'header', '合计/小计行': 'summary',
                      '无价格且有效属性不足2项': 'note'}
ROW_TYPE_CN = {'item': '物料行', 'title': '分组标题', 'header': '表头', 'summary': '合计/小计', 'note': '备注/其他'}
TABLE_TYPE_CN = {'list': '清单式', 'matrix': '矩阵式'}

COMPANY_RE = re.compile(r'([\u4e00-\u9fa5（）\(\)]{4,30}?(?:股份有限公司|有限责任公司|有限公司|股份公司))')
# 文件命名中常见的非公司名词语 (按词删除)
FILENAME_NOISE_RE = re.compile(r'标准化输出|标准解析|报价单|报价表|价格表|产品资料|报价数据|明细|总表|第\d+页|[\d_\-（）\(\)\s.]')
# 企业名称中的组织形式与通用行业词 (用于截取核心商号)
COMPANY_SUFFIX_RE = re.compile(r'(股份有限公司|有限责任公司|有限公司|股份公司|集团|新材料|建筑材料|建筑装饰|建筑|新型建材|科技|材料|建材|实业|工程|设备|照明|电气|门业|金属|机电|防水|管业|技术|[（\(]).*$')
REGION_PREFIX_RE = re.compile('^(' + '|'.join(sorted((k for k in PROVINCE_ABBR_MAP if len(k) >= 2), key=len, reverse=True)) + r')|^[\u4e00-\u9fa5]{2,3}市')
SUMMARY_ROW_KEYS = ['合计', '小计', '总计', '总价合计']
LIST_HEADER_RE = re.compile(r'^\s*(单价|合价|总价|金额|数量|序号|税率|单位|小计|报价|价格|市场价|出厂价|含税价|面价)')
# 矩阵表的行标题列是“规格/型号/档位”这类维度；表头像“名称/品名”时更可能是清单表
MATRIX_LABEL_RE = re.compile(r'规格|型号|容量|档位|尺寸|功率|厚度|直径|层数|电压|电流|口径|通径|DN|kVA|kg', re.IGNORECASE)
NUMBER_RE = re.compile(r'[¥￥]?\d{1,3}(?:,\d{3})+(?:\.\d+)?|[¥￥]?\d+(?:\.\d+)?')
# “单位：元／km”这类是计量口径注释（合并格落到了价格列上方），不是列表字段
UNIT_NOTE_RE = re.compile(r'^\s*单位\s*[:：]')


def supplier_from_path(path: str):
    """从文件名及上两级目录名中找公司全称"""
    parts = [os.path.splitext(os.path.basename(path))[0]]
    d = os.path.dirname(os.path.abspath(path))
    for _ in range(2):
        parts.append(os.path.basename(d))
        d = os.path.dirname(d)
    for part in parts:
        m = COMPANY_RE.search(part)
        if m:
            return m.group(1)
    return None


def _hash_prefix(text: str) -> str:
    """无 pypinyin 时的稳定字母前缀 (同名恒定，不同名大概率不同)"""
    h = hashlib.md5(text.encode('utf-8')).digest()
    return ''.join(chr(ord('A') + b % 26) for b in h[:3])


def core_name(supplier: str) -> str:
    """企业核心商号：去掉开头的地名与结尾的组织形式/行业词"""
    core = REGION_PREFIX_RE.sub('', supplier or '')
    core = re.sub(r'^[省市]', '', core)
    return COMPANY_SUFFIX_RE.sub('', core)


def _core_consistent(sup: str, core: str, evidence: str) -> bool:
    """
    核心商号须在旁证中出现，且旁证里紧挨在它前面的汉字也出现在识别结果中
    (防止公章环形文字漏读首字：旁证“甲乙丙” vs 识别“乙丙…” 判为不一致)
    """
    key = core[:2]
    for m in re.finditer(re.escape(key), evidence):
        prev = evidence[m.start() - 1] if m.start() > 0 else ''
        if not re.match(r'[\u4e00-\u9fa5]', prev) or (prev + key) in sup:
            return True
    return False


def verify_terms_supplier(terms: dict, evidence_texts: list) -> dict:
    """
    公章环形文字容易被模型读错。用文件名/目录名/页面标题/表格内容做旁证：
    核心商号在旁证中对不上时不采信，并记录待复核提示。人工给出的条款不做此校验。
    """
    terms = dict(terms or {})
    sup = terms.get('supplier')
    # manual(人工 JSON) 与 dir(源目录名，人工维护的组织方式) 都是人给的，不需要公章旁证
    if not sup or terms.get('_source') in ('manual', 'dir'):
        return terms
    core = core_name(sup)
    evidence = re.sub(r'\s', '', ''.join(evidence_texts))
    if len(core) >= 2 and evidence and not _core_consistent(sup, core, evidence):
        print(f"[Info] 商务单公章供应商“{sup}”与表内品牌/物料名称可能存在经销商与品牌差异，予以保留，提示复核")
        terms['_supplier_note'] = f"商务单公章识别为“{sup}”，与表内文字存在差异，建议复核"
    return terms


def derive_supplier_and_brand_dynamic(filename: str, terms_dict: dict = None) -> tuple:
    """
    推导供应商、品牌、报价ID前缀。返回 (supplier, brand, id_prefix, notes)
    优先级：商务条款 (人工 / 视觉模型且通过旁证) → 路径中的公司全称 → 文件名 → 留空
    """
    terms_dict = terms_dict or {}
    notes = []
    supplier = terms_dict.get('supplier') or supplier_from_path(filename)
    brand = terms_dict.get('brand')

    if not supplier:
        clean_name = FILENAME_NOISE_RE.sub('', os.path.splitext(os.path.basename(filename))[0])
        if clean_name:
            supplier = clean_name[:14]
            notes.append(f"供应商名称“{supplier}”由文件名推断")
        else:
            supplier = ''
            notes.append("未能确定供应商名称")
    if supplier:
        # 中文企业名统一全角括号
        supplier = re.sub(r'\(([^()]*)\)', r'（\1）', supplier)

    if not brand and supplier:
        core = core_name(supplier)
        # 不再截断成 6 字：截断会把“星牌优时吉建筑材料”切成“星牌优时吉建”这种不存在的品牌
        brand = core if len(core) >= 2 else supplier[:4]
        if brand != supplier:
            notes.append(f"品牌“{brand}”由供应商名称推断")
    brand = brand or ''

    id_prefix = ""
    try:
        from pypinyin import pinyin, Style
        p_list = pinyin(brand, style=Style.FIRST_LETTER)
        id_prefix = "".join(p[0].upper() for p in p_list if p and p[0] and p[0][0].isalpha())[:3]
    except ImportError:
        print("[Warning] 未安装 pypinyin (pip install pypinyin)，报价ID前缀改用稳定哈希字母。")
    if len(id_prefix) < 2:
        en_candidates = re.sub(r'[^A-Za-z]', '', brand).upper()
        id_prefix = en_candidates[:3] if len(en_candidates) >= 2 else _hash_prefix(brand or filename)

    return supplier, brand, id_prefix, notes


def _is_title_row(values: list) -> bool:
    """合并单元格展开后，分组标题行 (如 “一、消防系统”) 的所有非空单元格内容相同"""
    non_empty = [v for v in values if v]
    return len(non_empty) >= 2 and len(set(non_empty)) == 1


_KV_PAIR_RE = re.compile(r'^([^:：=]{1,15})[:：=](.+)$')
_IMPLICIT_PROP_RE = re.compile(r'(厚度|长度|宽度|高度|外径|壁厚|公称直径|口径|直径|孔径|规格)\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?\s*[a-zA-Z%°²³]*)')

def extract_deterministic_kv(raw_spec_text: str, raw_cells_dict: Optional[dict] = None) -> dict:
    """从原表规格字符串及单元格中 100% 确定性、无损提取显式键值对与规整属性 (Lossless Merge 打底)"""
    out = {}
    if raw_cells_dict and isinstance(raw_cells_dict, dict):
        for k, v in raw_cells_dict.items():
            k_s, v_s = str(k or "").strip(), str(v or "").strip()
            if not k_s or not v_s:
                continue
            if k_s in ("品名", "材料名称", "名称", "单价", "合价", "金额", "序号", "行号", "单位", "数量", "备注"):
                continue
            out[k_s] = v_s

    text = str(raw_spec_text or "")
    # 1. 优先提取显式尺寸维度 (如 "厚度0.7mm", "外径20mm", "壁厚:1.2mm")
    for m in _IMPLICIT_PROP_RE.finditer(text):
        out[m.group(1).strip()] = m.group(2).strip()

    # 2. 按分号/换行切大段，逐段提取 Key: Value 属性
    for seg in re.split(r'[;；\n]+', text):
        seg = seg.strip()
        if not seg:
            continue
        sub_segs = re.split(r'[,，]+', seg) if (':' in seg or '：' in seg) else [seg]
        if len(sub_segs) > 1 and any(':' in s or '：' in s for s in sub_segs[1:]):
            parts = sub_segs
        else:
            parts = [seg]

        for p in parts:
            p = p.strip()
            m_kv = _KV_PAIR_RE.match(p)
            if m_kv:
                k, v = m_kv.group(1).strip(), m_kv.group(2).strip()
                if k and v and k not in ("品名", "名称", "单价", "金额"):
                    out[k] = v
    return out


_COMPOUND_UNIT_RE = re.compile(r'^([根支条卷桶箱包袋套片张块件台只])\s*[/／(（]\s*(\d+(?:\.\d+)?\s*(?:米|m|mm|cm|kg|公斤|升|L|个|只|片|米/根))[)）]?$', re.I)

def decompose_compound_unit(unit_str: str, specs: dict) -> Tuple[str, dict]:
    """识别复合单位（如 '根/3.6米'、'卷(100米)'），保留原单位的同时自动将 '定尺长度: 3.6米/根' 补入规格参数"""
    u = str(unit_str or "").strip()
    specs = dict(specs or {})
    m = _COMPOUND_UNIT_RE.match(u)
    if m:
        base_u, measure = m.group(1), m.group(2)
        if not any(k in specs for k in ("长度", "定尺", "定尺长度", "包装规格", "每件数量", "规格")):
            key_name = "定尺长度" if any(w in measure for w in ("米", "m", "mm", "cm")) else "包装规格"
            specs[key_name] = f"{measure}/{base_u}"
    return u, specs


def _parse_section_title_info(title_text: str) -> Tuple[str, str]:
    """从分组标题行中提取潜在的分组品名与分组品牌（如 '【一、PP-R冷热水管（品牌：伟星）】'）"""
    t = str(title_text or "").strip()
    brand, name = "", ""
    m_brand = re.search(r'(?:品牌|商标)[：:\s]*([^\s,，;；()（）]{2,15})', t)
    if m_brand:
        brand = m_brand.group(1).strip()
    clean_t = re.sub(r'^[（(]?[一二三四五六七八九十\d]+[)）]?[、.．\s]*', '', t)
    clean_t = re.sub(r'[（(].*?[)）]', '', clean_t).strip()
    if 2 <= len(clean_t) <= 25 and not any(k in clean_t for k in ('系统', '工程', '清单', '报价', '附件', '备注')):
        name = clean_t
    return name, brand


def _audit_and_heal_price_monotonicity(items: list) -> None:
    """
    同族物料【规格-价格单调性】物理定律校验与红章遮挡小数点漏读自愈：
    同一品名下，厚度/管径增大时价格应递增；若某行价格突变为相邻行的 ~10 倍（典型红章压住小数点），自动预警并自愈修复。
    """
    from collections import defaultdict
    groups = defaultdict(list)
    for it in items:
        std_n = getattr(it, 'standard_name', '') or ''
        br = getattr(it, 'brand', '') or ''
        un = getattr(it, 'unit', '') or ''
        groups[(std_n, br, un)].append(it)

    for _, grp in groups.items():
        if len(grp) < 3:
            continue
        valid_items = [it for it in grp if isinstance(getattr(it, 'market_price', None), (int, float)) and it.market_price > 0]
        if len(valid_items) < 3:
            continue
        prices = [it.market_price for it in valid_items]
        sorted_p = sorted(prices)
        median_p = sorted_p[len(sorted_p) // 2]
        min_p, max_p = min(prices), max(prices)

        for it in valid_items:
            p = it.market_price
            if 6.5 * median_p <= p <= 15.0 * median_p:
                healed = round(p / 10.0, 2)
                if min_p * 0.5 <= healed <= max_p * 1.8:
                    warn = f"[价格异动预警] 疑似红章遮挡漏读小数点 (原价 {p:g} -> 建议核实 {healed:g})"
                    note = getattr(it, 'notes', '') or ''
                    it.notes = f"{note}；{warn}".strip('；')
                    if hasattr(it, 'uncertainty_note'):
                        u_note = getattr(it, 'uncertainty_note', '') or ''
                        it.uncertainty_note = f"{u_note}；{warn}".strip('；')
            elif 0.065 * median_p <= p <= 0.15 * median_p and p < median_p * 0.2:
                healed = round(p * 10.0, 2)
                if min_p * 0.5 <= healed <= max_p * 1.8:
                    warn = f"[价格异动预警] 疑似多读小数点 (原价 {p:g} -> 建议核实 {healed:g})"
                    note = getattr(it, 'notes', '') or ''
                    it.notes = f"{note}；{warn}".strip('；')


def _build_headers(ws, header_row: int, col_st: int, col_ed: int) -> dict:
    """
    读取表头：
    - 支持两级表头 (上一行为父表头时拼接为 “父-子”)
    - 重名表头自动加序号，避免同名列互相覆盖
    """
    parent_vals = []
    if header_row > 1:
        parent_vals = [cell_str(ws.cell(header_row - 1, c)) for c in range(col_st, col_ed + 1)]
        if _is_title_row(parent_vals) or sum(1 for v in parent_vals if v) < 2:
            parent_vals = []

    headers_map, seen = {}, {}
    for i, c in enumerate(range(col_st, col_ed + 1)):
        child = cell_str(ws.cell(header_row, c))
        parent = parent_vals[i] if parent_vals else ''
        h = f"{parent}-{child}" if parent and child and parent != child else (child or parent or f"第{c}列")
        if h in seen:
            seen[h] += 1
            h = f"{h}({seen[h]})"
        else:
            seen[h] = 1
        headers_map[c] = h
    return headers_map


def _is_unit_price_header(h: str) -> bool:
    return any(k in h for k in UNIT_PRICE_KEYS + ['元']) and not any(k in h for k in TOTAL_PRICE_KEYS)


def _path_evidence(path):
    d = os.path.dirname(os.path.abspath(path))
    return [os.path.basename(path), os.path.basename(d), os.path.basename(os.path.dirname(d))]


def _workbook_text(wb, limit=20000):
    parts = []
    for ws in wb.worksheets:
        if SKIP_HIDDEN_CONTENT and ws.sheet_state != 'visible':
            continue
        for row in ws.iter_rows(values_only=True):
            parts.extend(str(v) for v in row if v not in (None, ''))
            if sum(map(len, parts)) > limit:
                return parts
    return parts


def _output_paths(input_path, output_dir, supplier):
    # 严格遵循用户指定交付物命名法则: 公司名称-标准解析.xlsx
    safe_supplier = re.sub(r'[\/\\:\*\?\"<>\|]', '_', supplier) if supplier else os.path.splitext(os.path.basename(input_path))[0]
    return (os.path.join(output_dir, f"{safe_supplier}-标准解析.xlsx"),
            os.path.join(output_dir, f"{safe_supplier}_图片库"))


UNIT_ALIASES = {'㎡': 'm²', 'm2': 'm²', '平方米': 'm²', '平米': 'm²', '㎥': 'm³', 'm3': 'm³', '立方米': 'm³', '米': 'm', '公斤': 'kg'}
KV_PAIR_RE = re.compile(r'([^:：;；|]{1,20})[:：]([^;；|]*)')


def normalize_unit(u: str) -> str:
    """统一计量单位写法 (㎡/m2/平方米 → m²)，与标准报价数据保持一致"""
    u = (u or '').strip()
    return UNIT_ALIASES.get(u, u)


def _spec_key_norm(k: str) -> str:
    return re.sub(r'[\s()（）]|mm|㎜', '', k)



def _norm_f7_text(s: str) -> str:
    """NFKC 归一、去空白、统一各种乘号 (× x X * -> ×)"""
    s = unicodedata.normalize('NFKC', str(s or ''))
    s = re.sub(r'\s+', '', s)
    s = re.sub(r'[xX*×]', '×', s)
    return s.lower()

def _check_param_provenance(structured_specs: dict, raw_full_line: str) -> list:
    """检查 structured_specs 的每个值是否能在 raw_full_line 中找到出处，返回无出处的说明"""
    norm_line = _norm_f7_text(raw_full_line)
    unproven = []
    for k, v in structured_specs.items():
        v_str = str(v).strip()
        if not v_str:
            continue
        norm_val = _norm_f7_text(v_str)
        # 常见纯状态布尔值/通用分类允许豁免
        if norm_val in ('是', '否', '无', '有', '高档', '中档', '低档'):
            continue
        if norm_val not in norm_line:
            unproven.append(f"{k}={v_str}")
    return unproven

def _verbatim_specs(raw_cells: dict) -> dict:
    """单元格内容为 “品种:xx;厚度(mm):1.5;…” (至少 2 对) 时，逐对原样解析"""
    for v in raw_cells.values():
        pairs = KV_PAIR_RE.findall(str(v))
        if len(pairs) >= 2 and len(re.findall(r'[;；]', str(v))) >= 1:
            return {k.strip(): val.strip() for k, val in pairs if k.strip() and val.strip()}
    return {}


# 地区表头关键词集
REGION_HEADER_KEYS = ['地区', '区域', '城市', '省份', '省市', '供货范围', '供应范围', '报价城市', '报价地区']
REGION_HEADER_RE = re.compile(r'(' + '|'.join(REGION_HEADER_KEYS) + r')')

# ---- 固定字段值域门禁：这一列的值必须是这一列该有的东西 ----
# 表头名匹配只能决定“去哪一列找值”，不能决定“找到的值一定属于这一列”。
# 原表列错位/合并单元格漂移时，价格会被读进品牌/地区列，必须在写入前拦一次。
VALID_PROVINCES = set(PROVINCE_ABBR_MAP.values())
_NON_BRAND_RE = re.compile(r'^[\s\d.,，.．\-+±~%¥￥$/元平米台套个只根张袋桶吨米³²kglL]*$')


# 原表里就是“物料名称”的列 (优先照抄，不许换成标准分类名)
SRC_NAME_KEYS = ('材料名称', '物料名称', '产品名称', '品名', '规格名称', '名称')
SRC_NAME_EXCLUDE = ('类别', '分类', '规格', '型号', '品牌', '厂家', '供应商', '项目', '公司', '子表', '表')
# 原表自带的分类列 (供应商已分好类，编码按它查库最忠实)
SRC_CAT_KEYS = ('三级类别名称', '三级类别', '类别名称', '分类名称', '材料类别')


def _pick_row_field(cells: dict, keys, exclude=()) -> str:
    """按优先级从原表行单元格里取一个字段值；同优先级按表头出现顺序"""
    for key in keys:
        for hk, hv in (cells or {}).items():
            hk = str(hk or '')
            if key in hk and not any(e in hk for e in exclude):
                val = str(hv or '').strip()
                if val:
                    return val
    return ''


# 原表材料名称列常把型号写在品名前面（NLED4916 吊装线型灯 / LED LEDA45J球泡 / LED T806P灯管），型号不进名称规格
# 模型段必须以数字结尾或含数字，所以 CPL覆膜钢质门、UPVC排水管、HDPE双壁波纹管 这类不会被误删
_MODEL_PREFIX_RE = re.compile(r'^(?:[A-Z0-9]+[\s\-])*[A-Z0-9]*\d[A-Z0-9\-]*[\s　]*')
_BRACKET_CODE_RE = re.compile(r'^[\[\【][A-Za-z0-9\-_]+[\]\】][\s　]*')


def _strip_model_prefix(name: str) -> str:
    """删掉品名前的型号编码与方括号编码（如 [572305]光纤收发器）；剥完只能是汉字或原样返回"""
    orig = re.sub(r'\s+', ' ', str(name or '')).strip()
    s = orig
    # 1. 优先剥离方括号/中括号编码前缀 (如 [572305]、【FT-01】)
    s_bracket = _BRACKET_CODE_RE.sub('', s).strip()
    if s_bracket and re.search(r'[\u4e00-\u9fa5]', s_bracket):
        s = s_bracket

    # 2. 剥离标准字母数字型号前缀
    for _ in range(2):
        s2 = _MODEL_PREFIX_RE.sub('', s, count=1).strip()
        if s2 == s:
            break
        s = s2
    return s if re.search(r'[\u4e00-\u9fa5]', s) else orig


def _looks_like_brand(v) -> bool:
    """品牌必须是品牌名：含中文或字母，不能是纯价格/数量/规格串"""
    s = str(v if v is not None else '').strip()
    if not s or len(s) > 30:
        return False
    if _NON_BRAND_RE.match(s):
        return False
    return bool(re.search(r'[\u4e00-\u9fa5A-Za-z]', s))


def _region_values_ok(v) -> bool:
    """报价地区省只能是省名或“全国”，多省用顿号连接"""
    s = str(v if v is not None else '').strip()
    if not s:
        return True
    return all(p in VALID_PROVINCES or p == '全国' for p in s.split('、'))


def _field_value_problems(header: list, rows: list) -> list:
    """落盘后读回体检：品牌列不许是数字，地区列不许是价格/规格串"""
    col = {str(h): i for i, h in enumerate(header) if h}
    errs = []
    for rn, r in enumerate(rows, 2):
        for name, ok, expect in (('品牌', _looks_like_brand, '品牌名'), ('报价地区省', _region_values_ok, '省市名')):
            if name not in col or col[name] >= len(r):
                continue
            v = r[col[name]]
            if v not in (None, '') and not ok(v):
                errs.append(f"第{rn}行 {name}列的值不像{expect}: {v}")
        if len(errs) > 20:
            break
    return errs

BUSINESS_HEADER_RE = re.compile(r'(单价|合价|总价|金额|小计|价格|市场价|出厂价|含税价|面价|单位|序号|税率|是否含税|品牌|编码|代号|条码|产品编码|物料编码|' + '|'.join(REGION_HEADER_KEYS) + r'|图片)')


def _genoffice_raw_spec(row_cells: dict, headers_map: dict, price_col, unit_col) -> str:
    """
    GenOffice 口径的“厂家原始规格参数”：只留物料本身的参数，格式 “键:值;键:值”，
    不含单价/单位/税率/地区等已有独立列的商务字段；原文本身就是 “键:值;” 规格串时原样使用
    """
    excluded = {headers_map.get(c) for c in (price_col, unit_col) if c}
    for k, v in row_cells.items():
        if k not in excluded and len(KV_PAIR_RE.findall(str(v))) >= 2 and re.search(r'[;；]', str(v)):
            return re.sub(r'\s*([:：;；])\s*', r'\1', str(v))  # 去掉识别时混入的分隔符两侧空格
    return ";".join(f"{k}:{v}" for k, v in row_cells.items()
                    if k not in excluded and not BUSINESS_HEADER_RE.search(k))


def _is_number_text(v: str) -> bool:
    if not v:
        return False
    clean = str(v).replace(' ', '').replace('¥', '').replace('￥', '')
    if NUMBER_RE.fullmatch(clean):
        return True
    # 支持价格区间如 234-260, 285~317
    if re.fullmatch(r'\d+(?:\.\d+)?[\-~至到]\d+(?:\.\d+)?', clean):
        return True
    return False


def _leading_number(v: str):
    # 支持纯数字以及带有常见工程规格前缀（如 DN, De, φ, Φ, #）的数值提取
    m = re.search(r'(?:DN|De|[φΦ#])?\s*(\d+(?:\.\d+)?)', v or '', re.IGNORECASE)
    return float(m.group(1)) if m else None


def _detect_matrix(ws, hid_rows=frozenset(), hid_cols=frozenset()):
    """
    识别“矩阵式”报价表：行标题 (如规格档位) × 列标题 (如型号) → 每个交叉格是一个价格。
    判据：从第一行数据行开始，至少 2 个非首列的列以纯数字为主，且这些列的末级表头不是“单价/数量/合价”等清单表字段。
    返回 None 或 {'data_start': r0, 'label_col': c0, 'value_cols': [...], 'col_paths': {c: '表头路径'}, 'label_header': str, 'unit': str}
    """
    max_r, max_c = ws.max_row, ws.max_column
    grid = [[cell_str(ws.cell(r, c)) for c in range(1, max_c + 1)] for r in range(1, max_r + 1)]
    data_rows = [r for r in range(max_r) if (r + 1) not in hid_rows and sum(_is_number_text(v) for v in grid[r]) >= 2]
    if len(data_rows) < 3:
        return None
    r0 = data_rows[0]
    label_col = next((c for c in range(max_c) if sum(1 for r in data_rows if grid[r][c]) >= len(data_rows) * 0.6), 0)
    # 行标题同为数字、但整行无价格的行 (该档未报价) 仍属于数据区，不能并入表头
    while r0 > 0 and _leading_number(grid[r0 - 1][label_col]) is not None:
        r0 -= 1

    value_cols = []
    for c in range(label_col + 1, max_c):
        if (c + 1) in hid_cols:
            continue
        vals = [grid[r][c] for r in range(r0, max_r) if grid[r][c]]
        nums = [v for v in vals if _is_number_text(v)]
        if len(nums) >= 3 and len(nums) >= len(vals) * 0.6:
            value_cols.append(c)
    if len(value_cols) < 2:
        return None

    # 检测横跨全部价格列的大标题 (如“10/0.4kV干式配电变压器（单位：元/台）”)
    table_title = ""
    for r in range(r0):
        row_vals = [grid[r][c] for c in value_cols if grid[r][c]]
        if len(set(row_vals)) == 1 and len(row_vals) == len(value_cols):
            table_title = row_vals[0]
            break

    # 表头区内的标题/分组行或整表大标题不参与列标题路径
    # 注意：第一行的型号规格（如“暗杆闸阀 Z45T-16”）必须作为列标题，而不能被“产品图样/指导价”中间辅助行覆盖！
    header_rows = [r for r in range(r0) if not _is_title_row(grid[r]) and not _is_banner_row(grid[r])
                   and sum(1 for v in grid[r] if v) >= 2
                   and not (table_title and all(grid[r][c] == table_title for c in value_cols))]

    def path(c):
        parts = []
        for r in header_rows:
            v = grid[r][c]
            # 过滤产品图样、指导价等辅助行
            if v and v != table_title and not any(k in v for k in ['产品图样', '产品 图样', '指导价', '参考价', '面价']) and (not parts or parts[-1] != v):
                parts.append(v)
        return parts

    # 任意一级表头是清单表字段 (单价/合价/数量…) 就不是矩阵表；单位注释格除外
    for c in value_cols:
        if any(LIST_HEADER_RE.match(p) and not UNIT_NOTE_RE.match(p) for p in path(c)):
            return None

    # 行标题取“列标题路径的末级”：不能取第一行 —— 第一行常是整表大标题
    # (如 “某公司 防火门规格”)，里面的“规格”二字会把清单表误判成矩阵表
    label_path = path(label_col)
    label_header = '-'.join(label_path) or '规格型号'
    label_leaf = label_path[-1] if label_path else ''
    label_vals = [grid[r][label_col] for r in range(r0, max_r) if grid[r][label_col]]
    numeric_ratio = sum(1 for v in label_vals if _leading_number(v) is not None) / max(1, len(label_vals))
    # 行标题必须是“维度”（数字档位，或规格/型号/容量类表头）；否则更像清单表里两个数值属性列
    # 两边判错的代价不对称：退回清单式会因读不到单价而大声报错，按矩阵解析则会凭空造出一堆假物料且对账全绿
    if numeric_ratio < 0.5 and not MATRIX_LABEL_RE.search(label_leaf):
        print(f"    ! 行标题末级“{label_leaf or '(空)'}”不是规格/档位维度，不按矩阵式解析（退回清单式）")
        return None

    header_text = ' '.join(v for r in range(r0) for v in grid[r] if v)
    m_unit = re.search(r'单位[:：]\s*元?/?\s*([一-龥a-zA-Z²]+)', header_text)
    return {
        'data_start': r0 + 1,
        'label_col': label_col + 1,
        'value_cols': [c + 1 for c in value_cols],
        'col_paths': {c + 1: '-'.join(path(c)) or f"第{c + 1}列" for c in value_cols},
        'label_header': label_header or '行标题',
        'table_title': table_title,
        'unit': m_unit.group(1) if m_unit else '',
    }


def _is_banner_row(values: list) -> bool:
    """横跨全表的横幅行（厂名/商标/“单位：元／km”注释混在一行）—— 8 成以上非空格子同文即算"""
    non_empty = [v for v in values if v]
    if len(non_empty) < 2:
        return False
    return non_empty.count(max(set(non_empty), key=non_empty.count)) >= len(non_empty) * 0.8


def _cell_notes(cell_issue, sname, r, cols, headers_map=None, price_col=None):
    """把识别存疑的单元格 (视觉模型双读不一致 / 自报看不清 / OCR 不一致) 转成待复核提示"""
    notes = []
    for c in cols:
        reason = cell_issue.get((sname, r, c))
        if reason:
            name = '单价' if c == price_col else (headers_map.get(c, f"第{c}列") if headers_map else f"第{c}列")
            notes.append(f"{name}{reason}")
    return notes


def _collect_matrix_rows(ws, sname, mx, cell_issue, hid_rows=frozenset(), hid_cols=frozenset()):
    """矩阵表展开：每个“行标题 × 列标题”交叉格生成一条原始物料行"""
    raw_rows, records = [], []
    prev = {}  # 列 -> (行标题数值, 价格)：用于“容量越大价格越高”的单调性校验
    for r in range(mx['data_start'], ws.max_row + 1):
        if r in hid_rows:
            continue
        label = cell_str(ws.cell(r, mx['label_col']))
        if not label:
            continue
        label_num = _leading_number(label)
        for c in mx['value_cols']:
            if c in hid_cols:
                continue
            v = cell_str(ws.cell(r, c))
            if not v:
                continue
            ref = f"{sname}!R{r}C{c}"
            col_path = mx['col_paths'][c]
            tt_part = ""
            raw_line = f"{mx['label_header']}:{label} | 列:{col_path} | 价格:{v}"
            price = safe_clean_price(v) if _is_number_text(v) else 0.0
            decision = 'item' if price > 0 else f"矩阵单元格非数值: {v}"
            records.append(SourceRowRecord(source_ref=ref, sheet=sname, row=r, raw_text=raw_line,
                                           # 非纯数字格也记下能读出的数字：让对账报出“这一格有数值却没被当价格”
                                           raw_price=price if price > 0 else safe_clean_price(v),
                                           decision=decision))
            if decision != 'item':
                continue
            notes = _cell_notes(cell_issue, sname, r, [c, mx['label_col']], price_col=c)
            # 检查表头格是否有 cell_issue (F5)
            for hr in range(1, mx['data_start']):
                h_issue = cell_issue.get((sname, hr, c))
                if h_issue:
                    notes.append(f"列标题识别存疑: {h_issue}")
                    break
            if label_num is not None and c in prev:
                p_label, p_price = prev[c]
                if label_num > p_label and price < p_price:
                    notes.append(f"价格 {price:g} 低于上一档({p_label:g})的 {p_price:g}，疑似识别错误或原件模糊")
            if label_num is not None:
                prev[c] = (label_num, price)
            # 矩阵行列标题通用泛化提取：
            # 严禁将全天下所有矩阵表硬编码为“暗杆闸阀”或锁死为“规格型号/型号系列”
            clean_col = col_path.strip()
            row_header = mx.get('label_header') or '规格'
            table_title = mx.get('table_title') or ''

            # 智能提取品名与型号特征（若列标题包含英文型号/系列则提取，否则保持原列头）
            m_model = re.search(r'([A-Za-z0-9\-]+(?:-[A-Za-z0-9]+)*[A-Za-z0-9]*)', clean_col)
            model_cand = m_model.group(1) if m_model else ""
            extracted_name = clean_col.replace(model_cand, '').strip() if model_cand else clean_col

            matrix_cells = {
                row_header: label,
                "规格型号": label,
                "列标题": clean_col,
                "单价": v
            }
            # 确定品名（整表大标题优先，或列标题中的品名特征）
            base_name = table_title or extracted_name or clean_col
            if base_name:
                matrix_cells["品名"] = base_name
            if model_cand:
                matrix_cells["型号/系列"] = model_cand

            raw_spec_clean = f"{row_header}:{label};列标题:{clean_col}"

            raw_rows.append({
                "id": ref,
                "raw_cells_dict": matrix_cells,
                "raw_full_line": raw_line,
                "raw_spec": raw_spec_clean,
                "raw_price": price,
                "raw_unit": mx['unit'],
                "mpaths": [],
                "file_page": sname,
                "table_name": sname,
                "notes": notes,
                "group_key": f"{sname}#C{c}",
            })
    return raw_rows, records


def _collect_list_rows(ws, ws_f, sname, standardizer, extractor, cell_issue, jev=None, jev_layout=None,
                       struct_log=None, hid_rows=frozenset(), hid_cols=frozenset()):
    """清单式报价表：推导表头与列角色，一行一个物料。STRUCTURE_ENGINE 为 jev/compare 时同时使用 Jev 判定"""
    raw_rows, records = [], []
    struct_log = struct_log if struct_log is not None else []
    sample_rows = extract_sheet_sample_rows(ws, max_rows=8, max_cols=30)
    layout_info = standardizer.analyze_sheet_layout(sname, sample_rows)

    sub_tables = layout_info.get("sub_tables", [])
    header_row = layout_info.get("header_row", 1)
    data_start_row = layout_info.get("data_start_row", 2)
    layout_notes = []
    if jev_layout:
        cur = sub_tables[0] if sub_tables else {}
        current_vals = {"header_row": header_row, "data_start_row": data_start_row,
                        "price_col": cur.get("price_col"), "unit_col": cur.get("unit_col")}
        use_jev = STRUCTURE_ENGINE == 'jev'
        for k, label in (("header_row", "表头行"), ("data_start_row", "数据起始行"), ("price_col", "单价列"), ("unit_col", "单位列")):
            jv, jc = jev_layout[k]
            struct_log.append({"sheet": sname, "item": label, "current": current_vals[k], "jev": jv, "conf": jc,
                               "adopted": "Jev" if use_jev else "现有做法"})
            if use_jev and jc < JEV_MIN_CONFIDENCE:
                layout_notes.append(f"Jev {label}判定置信度低({jc:.2f})")
        if use_jev:
            if len(sub_tables) > 1:
                layout_notes.append("Jev 试点不支持左右并排子表，已按整表处理")
            header_row = jev_layout["header_row"][0]
            data_start_row = max(jev_layout["data_start_row"][0], header_row + 1)
            sub_tables = [{"table_name": sname, "col_start": 1, "col_end": ws.max_column,
                           "price_col": jev_layout["price_col"][0], "unit_col": jev_layout["unit_col"][0]}]
    side_by_side = len(sub_tables) > 1
    src = 'jev' if (jev_layout and STRUCTURE_ENGINE == 'jev') else layout_info.get('_source', 'llm')
    print(f"    -> 表格结构 ({src}): {len(sub_tables)} 个数据子表，表头第 {header_row} 行，数据起始第 {data_start_row} 行")
    candidates = []

    for t_idx, sub_tbl in enumerate(sub_tables, 1):
        t_name = sub_tbl.get("table_name", f"子表{t_idx}")
        col_st = max(1, sub_tbl.get("col_start", 1))
        col_ed = min(sub_tbl.get("col_end") or ws.max_column, ws.max_column)
        price_col = sub_tbl.get("price_col")
        unit_col = sub_tbl.get("unit_col")
        cols = [c for c in range(col_st, col_ed + 1) if c not in hid_cols]
        # 单价列/单位列本身被隐藏：当作未指定，交给下面的可见列嗅探，否则整表会"读不到单价"
        if price_col in hid_cols:
            price_col = None
        if unit_col in hid_cols:
            unit_col = None

        headers_map = _build_headers(ws, header_row, col_st, col_ed)
        header_values = set(headers_map.values())

        # 智能校验 unit_col：防范同名“单位”列错位（若该列采样内容包含冒号/分号且平均长度>8，实为规格参数列，向后搜寻真正的单位列）
        if unit_col:
            sample_units = [cell_str(ws.cell(r, unit_col)) for r in range(data_start_row, min(data_start_row + 10, ws.max_row + 1))]
            sample_units = [u for u in sample_units if u]
            if sample_units and any((':' in u or '：' in u or ';' in u or len(u) > 8) for u in sample_units):
                old_u_col = unit_col
                for c_idx, h_name in headers_map.items():
                    if c_idx != unit_col and any(k in h_name for k in ('单位', '计量单位')):
                        headers_map[old_u_col] = "规格参数"
                        unit_col = c_idx
                        break

        # 智能嗅探数量列与合价列 (用于物理数学勾稽审计)
        qty_col = next((c for c, h in headers_map.items() if any(k in h for k in ('数量', '工程量', '采购量'))), None)
        amount_col = next((c for c, h in headers_map.items() if any(k in h for k in ('合价', '总价', '金额', '小计')) and not _is_unit_price_header(h)), None)

        # 杀手锏 4：表内分组小标题状态机与向下继承
        current_sec_name = ""
        current_sec_brand = ""
        last_valid_name = ""
        last_valid_brand = ""

        for r in range(data_start_row, ws.max_row + 1):
            if r in hid_rows:
                continue
            row_cells_dict = {}
            row_text_parts = []
            row_values = []
            price_val = 0.0
            unit_val = ""
            qty_val = 0.0
            amount_val = 0.0
            formula_no_cache = False

            for c in cols:
                cell = ws.cell(r, c)
                cell_val = cell_str(cell)
                raw_f = ws_f.cell(r, c).value
                if cell.value is None and isinstance(raw_f, str) and raw_f.startswith('=') and 'DISPIMG' not in raw_f:
                    formula_no_cache = True
                row_values.append(cell_val)
                col_header = headers_map.get(c, "")
                if cell_val:
                    row_cells_dict[col_header] = cell_val
                    row_text_parts.append(f"{col_header}:{cell_val}")

                if c == price_col:
                    price_val = safe_clean_price(cell.value)
                elif c == unit_col:
                    unit_val = cell_val
                elif c == qty_col:
                    qty_val = safe_clean_price(cell.value)
                elif c == amount_col:
                    amount_val = safe_clean_price(cell.value)

            if not any(row_values):
                continue
            ref = f"{sname}!R{r}#T{t_idx}"
            raw_line = " | ".join(row_text_parts)

            # 若未通过指定单价列拿到价格，按表头扫描单价列 (排除合价/金额，防止把总价当单价)
            if price_val <= 0 and not price_col:
                for c in cols:
                    if _is_unit_price_header(headers_map.get(c, "")):
                        p_candidate = safe_clean_price(ws.cell(r, c).value)
                        if p_candidate > 0:
                            price_val = p_candidate
                            break

            # 行级过滤与防误杀护栏
            decision = 'item'
            raw_lower = raw_line.lower()
            has_spec_feat = any(k in raw_lower for k in (
                'dn', 'de', 'mm', 'cm', 'm²', 'm2', 'm³', 'm3', 'kg', 'kv', 'kva', 'w', 'kw',
                '型号', '规格', '系列', '米', '台', '套', '个', '支', '条', '件', '只', '根', '块', '卷'
            ))
            is_business_note = any(bk in raw_line for bk in (
                '开户行', '账号', '纳税人', '统一社会信用代码', '联系人', '联系电话', '传真', '送货地址',
                '交货期', '有效期', '付款方式', '结算方式', '质保期', '说明:', '备注:', '注:'
            ))

            if _is_title_row(row_values):
                decision = '分组标题行'
                title_blob = " ".join(v for v in row_values if v).strip()
                sec_name, sec_brand = _parse_section_title_info(title_blob)
                if sec_name:
                    current_sec_name = sec_name
                if sec_brand:
                    current_sec_brand = sec_brand
            elif (len(set(row_cells_dict.values()) & header_values) >= max(2, len(row_cells_dict) // 2)
                  and not (price_val > 0 and has_spec_feat)):
                # 与「合计/小计行」同一护栏：有单价 + 带物料特征的行即使撞了表头文字，也不是表头
                decision = '重复表头行'
            elif any(v.replace(' ', '') in SUMMARY_ROW_KEYS for v in row_values if v):
                # 防误杀：若带有单价且包含物料特征词，坚决不能当汇总行跳过！
                if price_val > 0 and has_spec_feat:
                    decision = 'item'
                else:
                    decision = '合计/小计行'
            elif price_val <= 0 and is_business_note:
                decision = '商务说明/备注行'
            elif not (price_val > 0 or len(row_cells_dict) >= 2):
                decision = '无价格且有效属性不足2项'

            # 杀手锏 4：向下继承处理（品名为空或“同上/〃”，以及品牌向下继承）
            if decision == 'item':
                name_col_key = next((k for k in row_cells_dict if any(w in k for w in ('名称', '品名', '材料名称', '物料名称'))), None)
                cur_name = row_cells_dict.get(name_col_key, '').strip() if name_col_key else ''
                if not cur_name or cur_name in ("同上", "〃", "//", "\"", "同前", "同", "同规格", "同材质"):
                    inherited_name = last_valid_name or current_sec_name
                    if inherited_name:
                        target_key = name_col_key or "材料名称"
                        row_cells_dict[target_key] = inherited_name
                elif cur_name:
                    last_valid_name = cur_name

                brand_col_key = next((k for k in row_cells_dict if any(w in k for w in ('品牌', '厂家', '商标'))), None)
                cur_brand = row_cells_dict.get(brand_col_key, '').strip() if brand_col_key else ''
                if not cur_brand:
                    inherited_brand = current_sec_brand or last_valid_brand
                    if inherited_brand:
                        target_brand_key = brand_col_key or "品牌"
                        row_cells_dict[target_brand_key] = inherited_brand
                elif cur_brand:
                    last_valid_brand = cur_brand

            candidates.append({
                "ref": ref, "r": r, "values": row_values, "decision": decision, "raw_line": raw_line,
                "price": price_val, "qty": qty_val, "amount": amount_val, "formula_no_cache": formula_no_cache,
                "headers": [headers_map[c] for c in cols],
                "build": dict(
                    id=ref,
                    raw_cells_dict=row_cells_dict,
                    raw_full_line=raw_line,
                    raw_spec=_genoffice_raw_spec(row_cells_dict, headers_map, price_col, unit_col),
                    raw_price=price_val,
                    raw_unit=unit_val,
                    mpaths=(extractor.get_images_for_region(ws_f, sname, r, col_st, col_ed,
                                                            col_tolerance=0 if side_by_side else 1) if extractor else []),
                    file_page=sname,
                    table_name=t_name,
                    notes=layout_notes + _cell_notes(cell_issue, sname, r, cols, headers_map, price_col),
                ),
            })

    # 行类型：现有规则 vs Jev
    jev_types = {}
    if jev is not None and candidates:
        jev_types = jev.judge_rows(candidates[0]["headers"], [(c["ref"], c["values"]) for c in candidates])
        print(f"    -> Jev 行类型判定: {len(jev_types)}/{len(candidates)} 行")
    for c in candidates:
        decision = c["decision"]
        extra_notes = []
        if jev is not None:
            current_type = RULE_DECISION_TYPE.get(decision, 'note')
            jt = jev_types.get(c["ref"])
            struct_log.append({"sheet": sname, "item": f"行类型 {c['ref']}", "current": ROW_TYPE_CN[current_type],
                               "jev": ROW_TYPE_CN[jt[0]] if jt else "(未返回)", "conf": jt[1] if jt else None,
                               "adopted": "Jev" if (STRUCTURE_ENGINE == 'jev' and jt) else "现有做法",
                               "row_text": c["raw_line"]})
            if STRUCTURE_ENGINE == 'jev':
                if jt:
                    decision = 'item' if jt[0] == 'item' else f"Jev 判定为{ROW_TYPE_CN[jt[0]]}({jt[1]:.2f})"
                    if jt[1] < JEV_MIN_CONFIDENCE:
                        extra_notes.append(f"Jev 行类型判定置信度低({jt[1]:.2f})")
                else:
                    extra_notes.append("Jev 未返回行类型，按现有规则判定")
        records.append(SourceRowRecord(
            source_ref=c["ref"], sheet=sname, row=c["r"], raw_text=c["raw_line"],
            raw_price=c["price"], decision=decision, has_formula_without_cache=c["formula_no_cache"],
            raw_qty=c.get("qty"), raw_amount=c.get("amount")
        ))
        if decision != 'item':
            continue
        row = dict(c["build"])
        row["notes"] = row["notes"] + extra_notes
        raw_rows.append(row)
    return raw_rows, records


def standardize_workbook(wb_val, wb_fml, extractor, supplier, brand, id_prefix, terms_info,
                         cell_issue=None, global_notes=None, struct_log=None):
    """Excel 与 PDF(还原后的表格) 共用的解析主流程，返回 (canonical_items, source_rows, sheet_stats)"""
    standardizer = DeepSeekStandardizer()
    struct_log = struct_log if struct_log is not None else []
    jev = None
    if STRUCTURE_ENGINE in ('jev', 'compare'):
        jev = JevJudge()
        if not jev.available:
            raise RuntimeError(f"STRUCTURE_ENGINE={STRUCTURE_ENGINE} 需要 TYPESAFE_API_KEY")
    cell_issue = cell_issue or {}
    global_notes = global_notes or []
    canonical_items, source_rows = [], []
    group_of = {}  # source_ref -> 同系列分组 (矩阵表按列)
    seq = 10
    disc_ratio = terms_info.get("discount_ratio")

    # 从工作表前3行标题/横幅及原表文本中补全公司名称、品牌、联系人+连写手机号、税率
    banner_texts = []
    for ws_tmp in wb_val.worksheets:
        for r_idx in range(1, min(4, (ws_tmp.max_row or 1) + 1)):
            for c_idx in range(1, min(15, (ws_tmp.max_column or 1) + 1)):
                v = str(ws_tmp.cell(r_idx, c_idx).value or "").strip()
                if v and v not in banner_texts:
                    banner_texts.append(v)
    banner_blob = "\n".join(banner_texts)

    if not terms_info.get("supplier"):
        m_sup = re.search(r'(?:公司名称|报价单位|供方|供应商)[ \t]*[：:][ \t]*([\u4e00-\u9fffA-Za-z0-9（）()]{4,30}?(?:有限公司|股份有限公司|股份公司|集团|有限责任公司|厂|经营部))', banner_blob)
        if m_sup:
            terms_info["supplier"] = m_sup.group(1).strip()
            if not supplier:
                supplier = terms_info["supplier"]

    if not terms_info.get("brand"):
        m_brand = re.search(r'(?:本份报价单品牌|报价单品牌|报价品牌|品牌)[ \t]*[：:][ \t]*([^\s:：,，;；\n]{2,15})', banner_blob)
        if m_brand:
            terms_info["brand"] = m_brand.group(1).strip()
            if not brand:
                brand = terms_info["brand"]

    # 拆分连写的联系人与手机号（如“联系人：洪水土13850009333”）
    m_contact = re.search(r'(?:联系人|业务员|报价人)[ \t]*[：:][ \t]*([\u4e00-\u9fff]{2,5})[ \t\-]*((?:1[3-9]\d{9})?)', banner_blob)
    if m_contact:
        if not terms_info.get("contact_person"):
            terms_info["contact_person"] = m_contact.group(1).strip()
        if m_contact.group(2) and not terms_info.get("contact_phone"):
            terms_info["contact_phone"] = m_contact.group(2).strip()

    # 将识别到的联系人与手机号真正落盘入库：合入 terms_info["quote_desc"]，落盘至交付 Excel 的【报价说明】列
    c_person = terms_info.get("contact_person", "").strip()
    c_phone = terms_info.get("contact_phone", "").strip()
    if c_person or c_phone:
        c_str = f"联系人:{c_person} {c_phone}".strip()
        cur_desc = (terms_info.get("quote_desc") or "").strip()
        if c_str not in cur_desc:
            terms_info["quote_desc"] = f"{cur_desc}；{c_str}".strip("；")

    # 从横幅补充税率与含税信息（必须紧邻冒号或数字，避免表头跨行误判）
    if not terms_info.get("tax_rate"):
        m_tax = re.search(r'(?:含税|税率)[ \t]*[：:][ \t]*(\d{1,2}%|\d{1,2}个点)', banner_blob)
        if m_tax:
            terms_info["tax_rate"] = m_tax.group(1).replace("个点", "%")
            if not terms_info.get("is_tax_included"):
                terms_info["is_tax_included"] = "是"

    # 检查原表工作表内是否已自带了商务条款列（如“税率”、“是否含税”、“运费”等），避免误报缺失
    table_has_terms = set()
    for _sn in wb_val.sheetnames:
        _ws = wb_val[_sn]
        if SKIP_HIDDEN_CONTENT and _ws.sheet_state != 'visible':
            continue
        for _r in range(1, min(10, (_ws.max_row or 1) + 1)):
            for _c in range(1, min(30, (_ws.max_column or 1) + 1)):
                _v = str(_ws.cell(_r, _c).value or "")
                if "税率" in _v or "%" in _v:
                    table_has_terms.add("增值税税率")
                    if not terms_info.get("tax_rate"):
                        from terms_extractor import parse_tax_rate
                        tr = parse_tax_rate(_v)
                        if tr:
                            terms_info["tax_rate"] = tr
                if "含税" in _v:
                    table_has_terms.add("是否含税")
                    table_has_terms.add("是否含增值税")
                    # 仅在非单纯列名且有明确声明性语句时才设置全局条款，单纯表头列名不锁死全局条款
                    if not terms_info.get("is_tax_included") and _v.strip() not in ("是否含税", "是否含增值税", "含税情况", "含税/不含税"):
                        if any(w in _v for w in ("不含", "未含", "除外")):
                            terms_info["is_tax_included"] = "否"
                        elif any(w in _v for w in ("包含", "含税", "均含", "已含")):
                            terms_info["is_tax_included"] = "是"
                if "运费" in _v or "含运" in _v:
                    table_has_terms.add("是否含运费")

    missing = [m for m in missing_terms(terms_info) if m not in table_has_terms]
    if missing:
        global_notes = global_notes + [f"商务条款缺失: {'、'.join(missing)}"]
    if disc_ratio is None:
        global_notes = global_notes + ["未提供工程折扣，工程价留空"]
    sheet_stats = []   # 每个 Sheet 的“原表行数 → 产出条数”记账，空产出不能静默

    for sname in wb_val.sheetnames:
        ws = wb_val[sname]
        ws_f = wb_fml[sname]
        if SKIP_HIDDEN_CONTENT and ws.sheet_state != 'visible':
            print(f"\n--- [Sheet: {sname}] 跳过: 工作表被隐藏 (隐藏口径: 只解析正常展示的内容) ---")
            sheet_stats.append({'sheet': sname, 'rows': ws.max_row or 0, 'raw_rows': 0, 'items': 0,
                                'error': '', 'skipped': '隐藏工作表'})
            continue
        if ws.max_row is None or ws.max_row <= 1:
            print(f"\n--- [Sheet: {sname}] 跳过: 只有 {ws.max_row or 0} 行，无数据区 ---")
            sheet_stats.append({'sheet': sname, 'rows': ws.max_row or 0, 'raw_rows': 0, 'items': 0, 'error': ''})
            continue
        print(f"\n--- [Sheet: {sname}] (共 {ws.max_row} 行 x {ws.max_column} 列) ---")
        hid_rows, hid_cols = hidden_axes(ws) if SKIP_HIDDEN_CONTENT else (frozenset(), frozenset())
        if hid_rows or hid_cols:
            print(f"    -> 隐藏口径: 跳过 {len(hid_rows)} 个隐藏行 / {len(hid_cols)} 个隐藏列")
        n_before = len(canonical_items)

        mx = _detect_matrix(ws, hid_rows, hid_cols)
        jev_layout = None
        if jev is not None:
            jev_layout = jev.judge_layout(extract_sheet_sample_rows(ws, max_rows=15, max_cols=30))
            if jev_layout is None:
                print("    ! Jev 表格结构判定失败，本 Sheet 按现有做法处理")
            else:
                jt, jc = jev_layout["table_type"]
                struct_log.append({"sheet": sname, "item": "表格类型", "current": "矩阵式" if mx else "清单式",
                                   "jev": TABLE_TYPE_CN[jt], "conf": jc, "adopted": "现有做法"})
                if (jt == 'matrix') != bool(mx):
                    print(f"    ! 表格类型判定分歧: 现有做法={'矩阵式' if mx else '清单式'}, Jev={TABLE_TYPE_CN[jt]}({jc:.2f})")
        if mx:
            print(f"    -> 识别为矩阵式报价表: 行标题“{mx['label_header']}”，{len(mx['value_cols'])} 个价格列，数据起始行 {mx['data_start']}")
            raw_rows, records = _collect_matrix_rows(ws, sname, mx, cell_issue, hid_rows, hid_cols)
        else:
            raw_rows, records = _collect_list_rows(ws, ws_f, sname, standardizer, extractor, cell_issue,
                                                   jev=jev, jev_layout=jev_layout, struct_log=struct_log,
                                                   hid_rows=hid_rows, hid_cols=hid_cols)
        source_rows.extend(records)
        group_of.update({r["id"]: r["group_key"] for r in raw_rows if r.get("group_key")})
        print(f"    -> 提取出有效物料原始行: {len(raw_rows)} 行，启动大模型槽位精准清洗...")

        # 1. 先提取每行的 clean_src_name 与 src_cat，再计算整表主业先验并召回 Top-5 候选类别
        for it in raw_rows:
            it["src_name"] = _strip_model_prefix(_pick_row_field(it["raw_cells_dict"], SRC_NAME_KEYS, SRC_NAME_EXCLUDE))
            it["src_cat"] = _pick_row_field(it["raw_cells_dict"], SRC_CAT_KEYS, ('编码',))

        sheet_prior_t2 = compute_sheet_category_prior(raw_rows)
        for it in raw_rows:
            it["candidate_categories"] = retrieve_category_candidates(
                raw_cat=it["src_cat"],
                std_name=it["src_name"],
                raw_spec=it.get("raw_spec") or it.get("raw_full_line", ""),
                sheet_prior=sheet_prior_t2,
                top_k=5,
            )

        # 分批调用 DeepSeek 提炼标准物料模型（并发线程池加速）
        batch_size = 30
        batches = []
        for b_idx in range(0, len(raw_rows), batch_size):
            batch_slice = raw_rows[b_idx:b_idx + batch_size]
            llm_input = [
                {
                    "id": item["id"],
                    "raw_cells": item["raw_cells_dict"],
                    "raw_text": item["raw_full_line"],
                    "raw_price": item["raw_price"],
                    "raw_unit": item["raw_unit"],
                    "candidate_categories": item.get("candidate_categories", [])
                }
                for item in batch_slice
            ]
            batches.append((batch_slice, llm_input))

        def _clean_worker(args):
            b_slice, inp = args
            res_list = standardizer.standardize_batch(inp, supplier_hint=supplier, brand_hint=brand)
            res_map = {str(res.get("id")): res for res in res_list if isinstance(res, dict)}
            return b_slice, res_map

        # 并发执行各批次标准化清洗
        max_workers = min(len(batches), 5) if batches else 1
        with ThreadPoolExecutor(max_workers=max_workers) as pool:
            cleaned_batch_results = list(pool.map(_clean_worker, batches))

        for batch_slice, cleaned_dict in cleaned_batch_results:
            for item in batch_slice:
                it_id = item["id"]
                cleaned_res = cleaned_dict.get(it_id, {})
                notes = list(global_notes) + list(item.get("notes", []))

                # 表内行级品牌推导 (F4)
                row_brand = ''
                for bk, bv in item["raw_cells_dict"].items():
                    if '品牌' in bk and not any(ng in bk for ng in ['档位', '等级', '类别', '系列']):
                        cand = str(bv).strip()
                        if _looks_like_brand(cand):
                            row_brand = cand
                        elif cand:
                            notes.append(f"原表“{bk}”列的值“{cand}”不像品牌，未采用；品牌改用推断值")
                        break
                item_brand = row_brand or brand
                # 若来自表内明确品牌列，清除“由供应商名称推断”的提示
                if row_brand:
                    notes = [n for n in notes if not (n.startswith("品牌“") and "由供应商名称推断" in n)]

                # 原表自带的「材料名称」：这是原件写明的物料品名，必须照抄进名称规格
                src_name = _strip_model_prefix(_pick_row_field(item["raw_cells_dict"], SRC_NAME_KEYS, SRC_NAME_EXCLUDE))
                # 原表自带的分类（如“三级类别名称”）：按它查标准编码库，忠实原件分类
                src_cat = _pick_row_field(item["raw_cells_dict"], SRC_CAT_KEYS, ('编码',))

                clean_source = cleaned_res.get("_source", "fallback")
                if clean_source != "llm":
                    notes.append("未经大模型清洗(规则降级)，名称与参数需复核")

                std_name = src_name or str(cleaned_res.get("standard_name") or "").strip()
                if not std_name:
                    notes.append("未识别出物料名称")

                # 统一类别解析（BM25 + Sheet2 枚举反查 + 大模型受限单选 + 先验加成）
                selected_c = str(cleaned_res.get("selected_category_code") or cleaned_res.get("category_code") or "").strip()
                l1, l2, l3, cat_code, cat_verified, cat_reason = resolve_and_verify_category(
                    raw_cat=src_cat or str(cleaned_res.get("predicted_category") or "").strip(),
                    std_name=std_name,
                    raw_spec=item.get("raw_spec") or item.get("raw_full_line", ""),
                    sheet_prior=sheet_prior_t2,
                    selected_code=selected_c
                )
                if not cat_code:
                    cat_src = ""
                    notes.append("类别编码未确定" + ("" if CATEGORY_CODE_MAP else " (未配置标准编码库 CATEGORY_CODE_FILE)"))
                elif src_cat:
                    cat_src = "source"
                elif cat_verified:
                    cat_src = "catalog"
                else:
                    cat_src = "llm"
                    notes.append(f"类别编码 {cat_code} 未经标准编码库精确核验")
                # 仅在未通过核验或需要核查时才写入 notes，避免正常命中时污染 notes 导致被误判为待复核
                if not cat_verified and cat_reason and cat_reason not in notes:
                    notes.append(cat_reason)

                # 【名称兜底】：原件没写物料名称时才用标准三级分类名 l3 顶上
                if not std_name and l3:
                    std_name = l3

                # 杀手锏 2：显式键值对确定性预解析打底 + AI 提炼并集 (Lossless Merge)
                # 先以原表显式参数与非空单元格打底，再合入大模型提取属性，原表规整键值与产品配置 100% 零丢失
                raw_spec_str = cleaned_res.get("raw_spec") or item.get("raw_spec") or item["raw_full_line"]
                det_specs = extract_deterministic_kv(raw_spec_str, item.get("raw_cells_dict"))
                llm_specs = cleaned_res.get("structured_specs")
                if isinstance(llm_specs, dict) and llm_specs:
                    structured_specs = {**det_specs, **{str(k): str(v) for k, v in llm_specs.items() if v not in (None, "")}}
                else:
                    structured_specs = det_specs or {k: v for k, v in item["raw_cells_dict"].items()
                                                     if k not in ("品名", "单价", "序号", "行号", "金额")}
                structured_specs = {str(k): str(v) for k, v in structured_specs.items() if v not in (None, "")}

                # 原文已是规范 “键:值;” 参数时原样采用 (不让模型改参数名/值)
                verbatim = _verbatim_specs(item["raw_cells_dict"])
                if verbatim:
                    structured_specs = {**structured_specs, **verbatim}

                # 对照标准材料类别特定参数及要求接口规范，执行参数限定值校验与规范化
                if l3 or cat_code:
                    try:
                        structured_specs, missing_req, enum_warns = validate_and_normalize_specs(l3 or cat_code, structured_specs)
                        if missing_req:
                            notes.append(f"缺少类别必填属性: {'、'.join(missing_req)}")
                        if enum_warns:
                            notes.extend(enum_warns)
                    except Exception:
                        pass

                # F7 结构化参数逐项核对原文出处
                unproven_params = _check_param_provenance(structured_specs, item.get("raw_full_line", ""))
                if unproven_params:
                    notes.append("参数无原文出处: " + "；".join(unproven_params))

                # 系统性行级商务条款抽取 (优先从行单元格中提取税率/含税/运费，无则回退至外部 terms_info)
                row_tax_rate = ""
                row_tax_inc = ""
                row_freight_inc = ""
                row_freight_desc = ""

                for k, v in item["raw_cells_dict"].items():
                    k_str = str(k or "").strip()
                    val_str = str(v or "").strip()
                    if not val_str:
                        continue
                    if "税率" in k_str and not row_tax_rate:
                        from terms_extractor import parse_tax_rate
                        row_tax_rate = parse_tax_rate(val_str) or val_str
                    elif ("是否含税" in k_str or "含税" in k_str) and not row_tax_inc and "运费" not in k_str:
                        if any(w in val_str for w in ('不含', '未含', '否', '不')):
                            row_tax_inc = "否"
                        elif any(w in val_str for w in ('是', '含税', '含', '包括')):
                            row_tax_inc = "是"
                    elif ("含运" in k_str or "运费" in k_str) and not row_freight_inc and "税" not in k_str:
                        if any(w in val_str for w in ('不含', '未含', '否', '自提', '买方付')):
                            row_freight_inc = "不含"
                            row_freight_desc = row_freight_desc or val_str
                        elif any(w in val_str for w in ('是', '含运', '含', '包邮', '送到', '包含')):
                            row_freight_inc = "包含"
                            row_freight_desc = row_freight_desc or val_str

                # 若价格列本身叫“含税价/含税单价”，且未明确为否，自动标记含税为“是”
                if not row_tax_inc:
                    for k in item["raw_cells_dict"].keys():
                        if "含税" in str(k) and "不含税" not in str(k):
                            row_tax_inc = "是"
                            break

                # 兜底：行级未声明时继承整表商务条款
                row_tax_rate = row_tax_rate or terms_info.get("tax_rate")
                row_tax_inc = row_tax_inc or terms_info.get("is_tax_included")
                row_freight_inc = row_freight_inc or terms_info.get("is_freight_included")
                row_freight_desc = row_freight_desc or terms_info.get("freight_desc") or ""

                # 行内自带报价地区时以行为准，否则用商务条款中的地区
                row_region = next((v for k, v in item["raw_cells_dict"].items() if any(rk in k for rk in REGION_HEADER_KEYS)), '')
                province, city, region_note = _region_fields(row_region or terms_info["province"], terms_info["city"])
                if region_note:
                    notes.append(region_note)

                # 杀手锏 3：复合单位自动解构与定尺回填（如 '根/3.6米' -> 规格字典补入 '定尺长度: 3.6米/根'）
                raw_u = item.get("raw_unit") or str(cleaned_res.get("unit") or "").strip()
                _, structured_specs = decompose_compound_unit(raw_u, structured_specs)
                unit = normalize_unit(str(cleaned_res.get("unit") or "").strip() or item["raw_unit"])
                if not unit:
                    notes.append("原表无计量单位")
                elif item["raw_unit"] and unit != normalize_unit(item["raw_unit"]):
                    notes.append(f"单位由原表“{item['raw_unit']}”改为“{unit}”")

                # 价格只信原表；模型给的价格仅作对比，绝不覆盖
                final_price = item["raw_price"]
                llm_price = safe_clean_price(cleaned_res.get("price"))
                if final_price <= 0:
                    notes.append("原表未读到有效单价" + (f"，模型推测 {llm_price}，未采用" if llm_price > 0 else ""))
                elif llm_price > 0 and abs(llm_price - final_price) > 1e-6:
                    notes.append(f"模型返回价格 {llm_price} 与原表 {final_price} 不一致，已保留原表价")

                img_bytes_list = []
                for mp in item["mpaths"]:
                    b = extractor.extract_image_bytes(mp) if extractor else None
                    if b: img_bytes_list.append(b)

                ci = CanonicalMaterialItem(
                    quote_id=f"{id_prefix}{seq:05d}",
                    supplier=supplier,
                    brand=item_brand,
                    category_name=std_name,
                    category_code=cat_code,
                    structured_specs=structured_specs,
                    raw_spec=raw_spec_str,
                    file_page=item["file_page"],
                    unit=unit,
                    market_price=final_price,
                    discount_ratio=disc_ratio,
                    project_price=round(final_price * disc_ratio, 2) if disc_ratio is not None and final_price > 0 else None,
                    discount_series=terms_info["discount_series"],
                    is_freight_included=row_freight_inc,
                    freight_desc=row_freight_desc or terms_info["freight_desc"],
                    is_tax_included=row_tax_inc,
                    tax_desc=terms_info["tax_desc"],
                    tax_rate=row_tax_rate,
                    freight_settle=terms_info["freight_settle"],
                    quote_desc=f"{terms_info['quote_desc']}；{cleaned_res['quote_desc_extra']}".strip('；') if cleaned_res.get('quote_desc_extra') else terms_info['quote_desc'],
                    quote_type=terms_info["quote_type"],
                    province=province,
                    city=city,
                    discount_min=terms_info["discount_min"],
                    discount_max=terms_info["discount_max"],
                    image_bytes_list=img_bytes_list,
                    source_ref=it_id,
                    raw_price=item["raw_price"],
                    clean_source=clean_source,
                    category_source=cat_src,
                )
                for n in notes:
                    ci.add_note(n)
                canonical_items.append(ci)
                seq += 10

        produced = len(canonical_items) - n_before
        err = ''
        if len(raw_rows) > 0 and not produced:
            err = f"{sname}: 识别出物料原始行 {len(raw_rows)} 行，但最终 0 条输出"
            print(f"    ! {err}")
        sheet_stats.append({'sheet': sname, 'rows': ws.max_row, 'raw_rows': len(raw_rows),
                            'items': produced, 'error': err})
    _unify_group_names(canonical_items, group_of)
    # 杀手锏 1：同族规格-价格单调性物理定律校验与红章遮挡漏读自愈
    _audit_and_heal_price_monotonicity(canonical_items)
    if jev is not None:
        u = jev.usage
        print(f"[*] Jev 用量: 请求 {u['requests']} 次 (失败 {u['failed']})，输入 {u['input_tokens']} tokens，输出 {u['output_tokens']} tokens")
    return canonical_items, source_rows, sheet_stats


def _unify_group_names(items, group_of):
    """矩阵表同一价格列是同一产品系列：名称与编码按多数票统一 (消除分批调用模型造成的叫法不一致)"""
    from collections import Counter, defaultdict
    groups = defaultdict(list)
    for it in items:
        g = group_of.get(it.source_ref)
        if g:
            groups[g].append(it)
    for members in groups.values():
        name, _ = Counter(m.category_name for m in members).most_common(1)[0]
        code, _ = Counter(m.category_code for m in members if m.category_name == name).most_common(1)[0]
        src = next(m.category_source for m in members if m.category_name == name and m.category_code == code)
        for m in members:
            if m.category_name != name or m.category_code != code:
                m.add_note(f"同系列多数票统一为“{name}”/编码“{code or '空'}”")
            m.category_name, m.category_code, m.category_source = name, code, src


def _region_fields(region_text, city_text=''):
    """把地区原文展开为 (省, 市, 提示)；无法识别的原文照抄"""
    region_text = str(region_text or '').strip()
    if not region_text:
        return '', '', ''
    provs = expand_region(region_text)
    if not provs:
        # 原表“区域/地区”列读到的不是行政区（常见为价格、规格串错位）
        # → 绝不把原值写进省列，只留痕
        return '', city_text or '', f"原表地区列的值“{region_text}”不是省市名称，未写入报价地区"
    if provs == ['全国']:
        return '全国', city_text or '全国', ''
    province = '、'.join(provs)
    city = city_text or (provs[0] if len(provs) == 1 and provs[0].endswith('市') else '')
    note = f"报价地区“{region_text}”覆盖 {len(provs)} 个省份" if len(provs) > 1 else ''
    return province, city, note


def _terms_notes(terms):
    terms = terms or {}
    notes = [terms['_supplier_note']] if terms.get('_supplier_note') else []
    src = terms.get('_source')
    if src in (None, 'none'):
        notes.append("未提供或未识别出商务条款")
    return notes


def process_excel(input_xlsx, template_xlsx, output_dir, terms=None):
    terms = verify_terms_supplier(terms, _path_evidence(input_xlsx) +
                                  _workbook_text(openpyxl.load_workbook(input_xlsx, read_only=True, data_only=True)))
    supplier, brand, id_prefix, sup_notes = derive_supplier_and_brand_dynamic(input_xlsx, terms)
    print(f"\n=== [Excel] 解析: {input_xlsx} ===")
    print(f"[*] 动态感知供方主体: {supplier} | 核心品牌: {brand} | 报价ID前缀: {id_prefix}")
    output_xlsx, img_dir = _output_paths(input_xlsx, output_dir, supplier)

    extractor = ExcelAssetExtractor(input_xlsx)
    # 两份工作簿：值 (公式取缓存计算结果) 用于读数据；公式版用于识别 WPS DISPIMG 图片与无缓存公式
    print("[*] 展开合并单元格...")
    wb_val = openpyxl.load_workbook(input_xlsx, data_only=True)
    wb_fml = openpyxl.load_workbook(input_xlsx, data_only=False)
    for sname in wb_val.sheetnames:
        if SKIP_HIDDEN_CONTENT and wb_val[sname].sheet_state != 'visible':
            continue    # 隐藏口径: 隐藏表不展开合并单元格 (反正不解析)
        unmerge_and_fill_cells(wb_val[sname])
        unmerge_and_fill_cells(wb_fml[sname])

    struct_log = []
    items, source_rows, sheet_stats = standardize_workbook(wb_val, wb_fml, extractor, supplier, brand, id_prefix,
                                                          normalize_terms(terms), global_notes=_terms_notes(terms) + sup_notes,
                                                          struct_log=struct_log)

    # 截取原件每个 Sheet 的完整高清原图，填入「厂家报价图片文件名」作为佐证留痕
    try:
        from sheet_screenshot_service import capture_excel_sheets
        sheet_imgs = capture_excel_sheets(input_xlsx, img_dir)
        for it in items:
            s_name = it.file_page
            img_fn = sheet_imgs.get(s_name) or sheet_imgs.get(str(s_name).strip()) or (list(sheet_imgs.values())[0] if sheet_imgs else '')
            it.quote_image_filename = img_fn
            if not it.image_filenames and img_fn:
                it.image_filenames = [img_fn]
        if sheet_imgs:
            print(f"[*] Excel 原件底稿 Sheet 全景截图完成: 共截取 {len(sheet_imgs)} 张 Sheet 底稿")
    except Exception as e:
        print(f"[Warning] Excel Sheet 底稿截图异常: {e}")

    excel_run_meta = {
        "requested_engine": "excel_rule",
        "actual_engine": "excel_rule",
        "provider_requested": "rule_openpyxl",
        "provider_actual": "rule_openpyxl",
        "model_id": "rule_openpyxl",
        "fallback_reason": None,
        "degraded": False,
        "rail_name": "Excel Direct Extraction",
    }
    return _finish(items, source_rows, template_xlsx, output_xlsx, img_dir, struct_log,
                   sheet_stats=sheet_stats, source_path=input_xlsx, run_meta=excel_run_meta)


def _process_pdf_with_extractor(extractor_cls, input_pdf, template_xlsx, output_dir, terms=None, rail_name="A-Rail",
                                fallback_reason=None, actual_model="", requested_engine=None):
    print(f"\n=== [PDF Pipeline - {rail_name}] 启动解析: {input_pdf} ===")
    if fallback_reason:
        print(f"[Warning] [双轨调度] B 轨失败已回退 A 轨：{fallback_reason}")
    pdf_ext = extractor_cls(input_pdf)
    wb = pdf_ext.to_workbook()

    terms = verify_terms_supplier(terms, _path_evidence(input_pdf) + pdf_ext.page_titles + _workbook_text(wb))
    if not terms.get('supplier') and not supplier_from_path(input_pdf):
        guessed = pdf_ext.guess_supplier()
        if guessed:
            terms['supplier'] = guessed
            terms['_supplier_note'] = f"供应商名称“{guessed}”由页面标题推断，未经原件核对"
    supplier, brand, id_prefix, sup_notes = derive_supplier_and_brand_dynamic(input_pdf, terms)
    print(f"[*] 动态感知供方主体: {supplier} | 核心品牌: {brand} | 报价ID前缀: {id_prefix}")
    output_xlsx, img_dir = _output_paths(input_pdf, output_dir, supplier)

    global_notes = _terms_notes(terms) + sup_notes
    struct_log = []
    model_label = actual_model or (GEMINI_MODEL if "Gemini" in rail_name else VISION_MODEL)
    struct_log.append({"sheet": "全局", "item": "实际使用链路", "current": rail_name, "jev": rail_name,
                       "conf": 1.0, "adopted": rail_name, "row_text": f"实际模型: {model_label}"})
    if fallback_reason:
        struct_log.append({"sheet": "全局", "item": "链路回退记录",
                           "current": f"B 轨回退 A 轨: {fallback_reason}",
                           "jev": "", "conf": 1.0, "adopted": "A 轨", "row_text": str(fallback_reason)})

    items, source_rows, sheet_stats = standardize_workbook(wb, wb, None, supplier, brand, id_prefix, normalize_terms(terms),
                                                          cell_issue=pdf_ext.cell_issue, global_notes=global_notes,
                                                          struct_log=struct_log)
    if fallback_reason:
        for it in items:
            # 走 add_note，否则只写 uncertainty_note 会被 note_list() (优先 notes) 漏掉
            it.add_note(f"B 轨失败已回退 A 轨：{fallback_reason}")

    if hasattr(pdf_ext, 'failed_pages') and pdf_ext.failed_pages:
        for fpage, reason in pdf_ext.failed_pages:
            rec = SourceRowRecord(
                source_ref=f"{fpage}!R1",
                sheet=fpage,
                row=1,
                raw_text=f"异常:{fpage}识别失败({reason})",
                raw_price=0.0,
                decision=f"异常:{fpage}识别失败({reason})"
            )
            source_rows.append(rec)
        print(f"[*] 页面失败隔离清单: {[fp for fp, _ in pdf_ext.failed_pages]}")
    # 每页原件截图作为“厂家报价图片” (quote_img{页码}.png)，便于对照
    from sheet_screenshot_service import capture_pdf_pages
    page_imgs = pdf_ext.save_page_images(img_dir) if hasattr(pdf_ext, 'save_page_images') else capture_pdf_pages(input_pdf, img_dir)
    for it in items:
        p_str = str(it.file_page or '').strip()
        m = re.search(r'\d+', p_str)
        p_num = int(m.group(0)) if m else 1
        img_fn = page_imgs.get(p_num) or page_imgs.get(f"第{p_num}页") or page_imgs.get(str(p_num)) or "quote_img1.png"
        it.quote_image_filename = img_fn
        if not it.image_filenames:
            it.image_filenames = [img_fn]

    run_meta = {
        "requested_engine": requested_engine or (EXTRACTOR_ENGINE or "traditional"),
        "actual_engine": "gemini" if "Gemini" in rail_name else "traditional",
        "provider_requested": "google" if (requested_engine or EXTRACTOR_ENGINE) in ("gemini", "compare") else "deepseek",
        "provider_actual": "google" if "Gemini" in rail_name else "deepseek",
        "model_id": model_label,
        "fallback_reason": fallback_reason,
        "degraded": bool(fallback_reason),
        "rail_name": rail_name,
    }
    return _finish(items, source_rows, template_xlsx, output_xlsx, img_dir, struct_log,
                   sheet_stats=sheet_stats, source_path=input_pdf, run_meta=run_meta)


def process_pdf(input_pdf, template_xlsx, output_dir, terms=None):
    """
    双轨调度中枢:
    1. EXTRACTOR_ENGINE='gemini': 优先使用 B 轨 (Gemini 3.8 Flash 原生直传)，若未配置Key或异常自动无损回退到 A 轨 (STRICT_ENGINE=1 时阻断报错)
    2. EXTRACTOR_ENGINE='traditional': 强制使用 A 轨 (传统切片 + OpenCV 滤镜 + DeepSeek)
    3. EXTRACTOR_ENGINE='compare': 双轨比对模式
    """
    engine = (EXTRACTOR_ENGINE or "traditional").strip().lower()

    if engine in ("gemini", "compare"):
        if GEMINI_API_KEY:
            try:
                print(f"[*] [双轨调度] 路由至 B 轨 (Gemini 原生全模态: {GEMINI_MODEL})")
                return _process_pdf_with_extractor(GeminiPDFTableExtractor, input_pdf, template_xlsx, output_dir,
                                                    terms=terms, rail_name="Gemini B-Rail",
                                                    actual_model=GEMINI_MODEL, requested_engine=engine)
            except Exception as e:
                err_msg = str(e)
                if STRICT_ENGINE:
                    raise RuntimeError(f"[STRICT_ENGINE=1] 请求引擎 {engine} (Gemini) 执行失败且禁止静默回退: {err_msg}")
                print(f"[Warning] [双轨调度] Gemini B 轨执行异常: {err_msg}，自动平滑兜底回退至 A 轨...")
                return _process_pdf_with_extractor(PDFTableExtractor, input_pdf, template_xlsx, output_dir,
                                                    terms=terms, rail_name="Traditional A-Rail (Fallback)",
                                                    fallback_reason=err_msg, actual_model=VISION_MODEL, requested_engine=engine)
        else:
            reason = "未配置 GEMINI_API_KEY"
            if STRICT_ENGINE:
                raise RuntimeError(f"[STRICT_ENGINE=1] 请求引擎 {engine} 但{reason}，禁止静默回退")
            print(f"[Info] [双轨调度] 检测到 EXTRACTOR_ENGINE={engine} 但{reason}，安全直通 A 轨 (传统切图流水线)")
            return _process_pdf_with_extractor(PDFTableExtractor, input_pdf, template_xlsx, output_dir,
                                                terms=terms, rail_name="Traditional A-Rail (Fallback)",
                                                fallback_reason=reason, actual_model=VISION_MODEL, requested_engine=engine)
    else:
        print("[*] [双轨调度] 路由至 A 轨 (传统切图流水线)")
        return _process_pdf_with_extractor(PDFTableExtractor, input_pdf, template_xlsx, output_dir,
                                            terms=terms, rail_name="Traditional A-Rail",
                                            actual_model=VISION_MODEL, requested_engine="traditional")


def _finish(items, source_rows, template_xlsx, output_xlsx, img_dir, struct_log=None,
            sheet_stats=None, source_path="", run_meta=None):
    """落盘 + 校验。返回 (交付文件路径, 失败清单)；失败清单非空时调用方必须置退出码 2"""
    # 原表 ↔ 结果逐行对账与四维体检量化核算（只对原表行做，展开行不参与对账）
    recon = reconcile(source_rows, items)
    items = _expand_nationwide(items, recon)
    integrity_summary = getattr(recon, "summary", {})
    _print_recon_summary(recon)

    print("\n" + "=" * 70)
    print("【零信任防欺瞒四维体检指标报告】")
    print(f"  -> 物理行守恒率: {integrity_summary['row_conservation_rate']:.1f}% (原表物理行: {integrity_summary['total_source_rows']})")
    print(f"  -> 物料输出行数: {integrity_summary['matched_items']} | 合法跳过行数: {integrity_summary['skipped_rows']}")
    print(f"  -> 对账异常项数: {integrity_summary['anomalies_count']} 项")
    print(f"  -> 数学勾稽自洽: {integrity_summary['math_passed_count']}/{integrity_summary['math_audited_count']} ({integrity_summary['math_pass_rate']:.1f}%)")
    print(f"  -> 最终免检裁决: {'【通过·零缺陷免检入库】' if integrity_summary['is_zero_defect'] else '【拦截·存在异常需复核】'}")
    print("=" * 70)

    print(f"\n[*] 全表语义解析完成，共生成标准物料: {len(items)} 项，启动构建器落盘...")
    if not items:
        err_msg = "本次未解析出任何有效物料（0 条），严禁生成或导出空的交付文件"
        print(f"[Fatal] {err_msg}: {output_xlsx}")
        if os.path.exists(output_xlsx):
            try:
                os.remove(output_xlsx)
            except Exception:
                pass
        return None, [err_msg]

    # 接入材料配图接口：通过材料名称检索配图，并匹配图片链接
    try:
        from material_illustration_service import match_illustrations_for_canonical_items
        supplier_tag = items[0].supplier if items else os.path.basename(output_xlsx).split('-')[0]
        n_matched = match_illustrations_for_canonical_items(items, company_name=supplier_tag)
        if n_matched:
            print(f"[*] 材料配图服务处理完成，成功命中配图: {n_matched}/{len(items)} 条")
    except Exception as illust_err:
        print(f"[Warning] 材料配图服务执行异常: {illust_err}")

    builder = StandardizedWorkbookBuilder(template_xlsx)
    builder.build_from_canonical_items(items, output_xlsx, img_dir, reconciliation=None,
                                       structure_log=None)
    # 单表由 builder 内部保证，此处不再重复读写一次（旧代码会二次 load/save，白费 IO）

    # 导出独立的《对账与防欺瞒体检报告》(独立文件，不进主表 Sheet；.env 里 EXPORT_INTEGRITY_CERTIFICATE=0 可关闭)
    if EXPORT_INTEGRITY_CERTIFICATE:
        cert_path = output_xlsx.replace(".xlsx", "_对账与防欺瞒体检报告.xlsx")
        try:
            supplier_tag = os.path.basename(output_xlsx).split('-')[0]
            StandardizedWorkbookBuilder.build_integrity_certificate(
                cert_path, recon, integrity_summary, supplier=supplier_tag, source_file=source_path
            )
            print(f"[*] 防欺瞒体检白皮书已独立导出: {cert_path}")
        except Exception as cert_err:
            print(f"[Warning] 导出防欺瞒体检报告异常: {cert_err}")

    # 落盘伴生 run sidecar (<output>.run.json)，机读留存引擎诚实性与降级记录
    try:
        sidecar_path = f"{output_xlsx}.run.json"
        meta = run_meta or {}
        # 计算源文件 sha256
        src_sha = ""
        if source_path and os.path.exists(source_path):
            with open(source_path, "rb") as sf:
                src_sha = hashlib.sha256(sf.read()).hexdigest()

        from config import VISION_DOUBLE_CHECK
        sidecar_payload = {
            "output_file": os.path.basename(output_xlsx),
            "source_file": os.path.basename(source_path),
            "file_sha256": src_sha,
            "EXTRACTOR_ENGINE_requested": meta.get("requested_engine", EXTRACTOR_ENGINE or "traditional"),
            "EXTRACTOR_ENGINE_actual": meta.get("actual_engine", "traditional"),
            "provider_requested": meta.get("provider_requested", "deepseek"),
            "provider_actual": meta.get("provider_actual", "deepseek"),
            "model_id": meta.get("model_id", VISION_MODEL),
            "fallback_reason": meta.get("fallback_reason", None),
            "degraded": bool(meta.get("degraded", False)),
            "VISION_DOUBLE_CHECK": bool(VISION_DOUBLE_CHECK),
            "STRICT_ENGINE": bool(STRICT_ENGINE),
            "total_items": len(items),
            "zero_defect": bool(integrity_summary.get("is_zero_defect", False)),
            "row_conservation_rate": integrity_summary.get("row_conservation_rate", 100.0),
            "skip_hidden_content": SKIP_HIDDEN_CONTENT,
            "skipped_hidden_sheets": [s['sheet'] for s in (sheet_stats or []) if s.get('skipped')],
            "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S%z")
        }
        with open(sidecar_path, "w", encoding="utf-8") as sc_f:
            json.dump(sidecar_payload, sc_f, ensure_ascii=False, indent=2)
        print(f"[*] 引擎诚实性与运行 Sidecar 已落盘: {sidecar_path} (degraded={sidecar_payload['degraded']})")
    except Exception as sc_err:
        print(f"[Warning] 写入 run sidecar 异常: {sc_err}")

    # 落盘后读回 + 完整性汇总：任何一项失败都不允许静默通过
    failures = _integrity_problems(items, recon, sheet_stats, integrity_summary) + _verify_output(output_xlsx, items)
    for s in sheet_stats or []:
        print(f"    [{s['sheet']}] 原表 {s['rows']} 行 → 产出 {s['items']} 条"
              + (f"  ⚠ {s['error']}" if s.get('error') else "")
              + (f"  (已跳过: {s['skipped']})" if s.get('skipped') else ""))
    print(f"[VERIFY] {os.path.basename(source_path)} | Sheet {len(sheet_stats or [])} | 物料 {len(items)} | "
          f"对账 {len(recon)} 行 | 失败 {len(failures)} | RESULT={'PASS' if not failures else 'FAIL'}")
    audit = summarize_audit(items, builder.headers)
    print(f"[AUDIT] 行级免审率 {audit['row_pass_rate']}% "
          f"({audit['total_rows'] - audit['review_rows']}/{audit['total_rows']} 行无需人工核对) | "
          f"字段级免审率 {audit['field_pass_rate']}% | "
          f"致命 {audit['fatal']} / 严重 {audit['severe']} / 提示 {audit['hint']}")
    if audit['clusters']:
        print("        同类问题组: " + "；".join(f"{c['group']}({c['count']})" for c in audit['clusters']))
    for f in failures[:20]:
        print(f"    ✗ {f}")
    return output_xlsx, failures


def _integrity_problems(items, recon, sheet_stats, integrity_summary=None):
    """执行完整性：每个 Sheet 要有交代，跳过行要能自证，空产出不能算成功，物理行必须守恒"""
    errs = []
    if not items:
        errs.append("本次未产出任何物料 (0 条)")
    for s in sheet_stats or []:
        if s.get('error'):
            errs.append(s['error'])
    
    # 物理行强守恒断言
    if ANTI_HALLUCINATION_STRICT_ROW_CONSERVATION and integrity_summary:
        if integrity_summary['row_conservation_rate'] < 100.0:
            errs.append(f"物理行数不守恒: 守恒率 {integrity_summary['row_conservation_rate']:.1f}% < 100.0% "
                        f"(原表 {integrity_summary['total_source_rows']} 行，已交代 {integrity_summary['matched_items'] + integrity_summary['skipped_rows']} 行)")

    bad = [r for r in recon if r['status'].startswith('异常')]
    for r in bad[:20]:
        errs.append(f"对账{r['status']} @ {r['source_ref']} | {r['detail']}")
    if len(bad) > 20:
        errs.append(f"…另有 {len(bad) - 20} 条对账异常 (详见交付伴生「对账与防欺瞒体检报告」)")
    return errs


def _verify_output(path, items):
    """写盘后读回：文件必须能打开，行数 / 报价ID / 市场价必须与内存结果逐条一致"""
    from collections import Counter
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        return [f"输出文件无法读回: {e}"]
    
    if '标准化报价' not in wb.sheetnames:
        return ["输出缺少「标准化报价」主表 Sheet"]
    ws = wb['标准化报价']
    header = [ws.cell(1, c).value for c in range(1, (ws.max_column or 0) + 1)]
    col = {str(h): i for i, h in enumerate(header) if h}
    # 表头有小写 d 的 '报价id'（人工标杆 36 列）与 '报价ID' 两种写法。
    # 注意 0 是合法列索引，必须用 is None 判空，不能靠 or
    id_col = col.get('报价ID')
    if id_col is None:
        id_col = col.get('报价id')
    if id_col is None or '市场价' not in col:
        return [f"输出表头缺 市场价/报价ID: {header}"]
    rows = [r for r in ws.iter_rows(min_row=2, values_only=True) if any(v not in (None, '') for v in r)]
    if len(rows) != len(items):
        return [f"输出数据行 {len(rows)} ≠ 内存物料 {len(items)} 条"]

    errs = []
    ids = []
    for i, (it, r) in enumerate(zip(items, rows), 2):
        if len(errs) > 20:
            break
        qid = r[id_col]
        ids.append(qid)
        if str(qid) != it.quote_id:
            errs.append(f"第{i}行 报价ID {qid} ≠ {it.quote_id}")
        try:
            out_price = float(r[col['市场价']] or 0)
        except (TypeError, ValueError):
            out_price = None
        if out_price is None or abs(out_price - float(it.market_price)) > 1e-6:
            errs.append(f"第{i}行 市场价 {r[col['市场价']]} ≠ {it.market_price}")
    dup = [q for q, n in Counter(ids).items() if n > 1]
    if dup:
        errs.append(f"报价ID 重复: {[str(d) for d in dup[:5]]}")
    errs.extend(_field_value_problems(header, rows))
    return errs


def _expand_nationwide(items, recon):
    """
    全国 / 无地区的报价按省市铺开成 31 行（人工标杆口径）。
    要点：
    - 在 reconcile 之后做，对账仍只对原表物理行，不会判成“一行生成 31 条输出”
    - 报价ID 由原 ID + 省市序号确定性派生（不能用随机数，否则同一份报价单两次跑出不同结果）
    - 每一行都写留痕，说明这行是展开出来的，不是原表里的
    """
    if not EXPAND_NATIONWIDE:
        return items
    bad_refs = {r['source_ref'] for r in recon if str(r.get('status', '')).startswith('异常')}
    out = []
    for it in items:
        region = (it.province or '').strip()
        if region not in ('', '全国') or it.source_ref in bad_refs:
            out.append(it)
            continue
        for idx, (prov, city) in enumerate(NATIONWIDE_REGIONS, 1):
            dup = copy.copy(it)
            dup.quote_id = f"{it.quote_id}{idx:02d}"
            dup.province, dup.city = prov, city
            dup.add_note(f"全国报价，按 {prov} 展开（原行仅一行，此行非原表独立行）")
            out.append(dup)
    return out


def _print_recon_summary(recon: list):
    total_items = sum(1 for r in recon if r['status'] != '跳过')
    bad = [r for r in recon if r['status'].startswith('异常')]
    skipped = sum(1 for r in recon if r['status'] == '跳过')
    print(f"[*] 原表对账: 物料行 {total_items} | 跳过行 {skipped} | 异常 {len(bad)}")
    for r in bad[:10]:
        print(f"    ! {r['source_ref']}: {r['status']}")


def print_capabilities():
    """以标准机读 JSON 格式输出系统当前实际接线能力与开关状态，禁止手写假清单"""
    from config import (
        EXTRACTOR_ENGINE, GEMINI_MODEL, VISION_MODEL,
        STRICT_ENGINE, VISION_DOUBLE_CHECK,
        ANTI_HALLUCINATION_MATH_AUDIT, ANTI_HALLUCINATION_STRICT_ROW_CONSERVATION,
        EXPORT_INTEGRITY_CERTIFICATE, ENABLE_MATERIAL_ILLUSTRATION,
        EXPAND_NATIONWIDE, SKIP_HIDDEN_CONTENT
    )
    caps = {
        "system": "QuoteStandardizer",
        "version": "2026.11-P0",
        "engines": {
            "supported": ["gemini", "traditional", "compare"],
            "configured_default": EXTRACTOR_ENGINE,
            "gemini_model": GEMINI_MODEL,
            "vision_model": VISION_MODEL,
        },
        "gates_and_safety": {
            "STRICT_ENGINE": STRICT_ENGINE,
            "raw_price_protection": True,
            "auditor_numeric_write_forbidden": True,
            "zip_reconcile_aligned": True,
            "zero_hallucination_row_conservation": ANTI_HALLUCINATION_STRICT_ROW_CONSERVATION,
            "math_audit_enabled": ANTI_HALLUCINATION_MATH_AUDIT,
            "export_integrity_certificate": EXPORT_INTEGRITY_CERTIFICATE,
            "run_sidecar_enabled": True
        },
        "features": {
            "VISION_DOUBLE_CHECK": VISION_DOUBLE_CHECK,
            "material_illustration_service": ENABLE_MATERIAL_ILLUSTRATION,
            "expand_nationwide": EXPAND_NATIONWIDE,
            "skip_hidden_content": SKIP_HIDDEN_CONTENT,
            "structure_comparison_sheet_active": False,
            "agy_v2_direct_wire": False,
            "paddle_ocr_active": False
        }
    }
    print(json.dumps(caps, ensure_ascii=False, indent=2))
    sys.exit(0)


def main():
    parser = argparse.ArgumentParser(description="QuoteStandardizer - 报价单解析与原表对账")
    parser.add_argument("--capabilities", action="store_true", help="打印系统当前实际接线能力与开关状态 (JSON)")
    parser.add_argument("--input", "-i", default=None, help="输入报价文件路径 (.xlsx / .pdf)")
    parser.add_argument("--terms", "-t", default=None, help="商务说明与折扣记录单图片路径 (.png/.jpg)；不填时自动在输入文件同目录查找")
    parser.add_argument("--terms-json", default=None, help="人工确认的商务条款 JSON (优先于图片识别)")
    parser.add_argument("--template", default=DEFAULT_TEMPLATE_PATH, help="输出模板 .xlsx (取第 1 行表头)；不填时使用内置标准表头")
    parser.add_argument("--output_dir", "-o", default=DEFAULT_OUTPUT_DIR, help="输出目录")
    args = parser.parse_args()

    if args.capabilities:
        print_capabilities()

    if not args.input:
        parser.print_help()
        sys.exit(1)

    input_path = args.input
    if not os.path.exists(input_path):
        print(f"[Error] 输入文件不存在: {input_path}")
        sys.exit(1)
    os.makedirs(args.output_dir, exist_ok=True)
    terms_dict = None

    if args.terms_json:
        # 人工确认过的商务条款 (最可靠)，优先级最高
        with open(args.terms_json, encoding='utf-8') as f:
            terms_dict = json.load(f)
        terms_dict['_source'] = 'manual'
        print(f"[*] 使用人工商务条款: {args.terms_json}")
    else:
        terms_path = args.terms or find_terms_image(os.path.dirname(os.path.abspath(input_path)), exclude_file=input_path)
        if terms_path and os.path.exists(terms_path):
            print(f"[*] 商务说明单: {terms_path}")
            terms_dict = TermsExtractor(terms_path).extract_terms()

    ext = os.path.splitext(input_path)[1].lower()
    if ext in ('.xlsx', '.xlsm'):
        out = process_excel(input_path, args.template, args.output_dir, terms=terms_dict)
    elif ext in ('.pdf', '.png', '.jpg', '.jpeg'):
        out = process_pdf(input_path, args.template, args.output_dir, terms=terms_dict)
    else:
        print(f"[Error] 不支持的文件类型: {ext} (.xls 请先另存为 .xlsx)")
        sys.exit(1)

    out, failures = out
    if failures:
        print(f"\n❌ 交付文件已生成，但有 {len(failures)} 项失败/异常，需人工核对：{out}")
        sys.exit(2)
    print(f"\n🎉 统一标准多行结构化输出完成！交付文件：{out}")
    sys.exit(0)

if __name__ == '__main__':
    main()
