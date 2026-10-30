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
import hashlib
import argparse
import json
import openpyxl
from concurrent.futures import ThreadPoolExecutor

from extractor import ExcelAssetExtractor, safe_clean_price, cell_str, unmerge_and_fill_cells, extract_sheet_sample_rows
from pdf_extractor import PDFTableExtractor
from gemini_pdf_extractor import GeminiPDFTableExtractor
from terms_extractor import TermsExtractor, normalize_terms, find_terms_image, missing_terms
from builder import StandardizedWorkbookBuilder
from config import (DEFAULT_TEMPLATE_PATH, DEFAULT_OUTPUT_DIR, PROVINCE_ABBR_MAP, CATEGORY_CODE_MAP,
                    resolve_category_code, expand_region, EXTRACTOR_ENGINE, GEMINI_API_KEY, GEMINI_MODEL)
from models import CanonicalMaterialItem, SourceRowRecord, reconcile
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
COMPANY_SUFFIX_RE = re.compile(r'(股份有限公司|有限责任公司|有限公司|股份公司|集团|科技|材料|建材|实业|工程|设备|照明|电气|门业|金属|机电|防水|管业|技术|[（\(]).*$')
REGION_PREFIX_RE = re.compile('^(' + '|'.join(sorted((k for k in PROVINCE_ABBR_MAP if len(k) >= 2), key=len, reverse=True)) + r')|^[\u4e00-\u9fa5]{2,3}市')
SUMMARY_ROW_KEYS = ['合计', '小计', '总计', '总价合计']
LIST_HEADER_RE = re.compile(r'^\s*(单价|合价|总价|金额|数量|序号|税率|单位|小计|报价|价格|市场价|出厂价|含税价|面价)')
NUMBER_RE = re.compile(r'[¥￥]?\d{1,3}(?:,\d{3})+(?:\.\d+)?|[¥￥]?\d+(?:\.\d+)?')


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
    if not sup or terms.get('_source') == 'manual':
        return terms
    core = core_name(sup)
    evidence = re.sub(r'\s', '', ''.join(evidence_texts))
    if len(core) >= 2 and evidence and not _core_consistent(sup, core, evidence):
        print(f"[Warning] 商务单识别出的供应商“{sup}”与文件名/报价表内容不符，不予采信")
        terms['_supplier_note'] = f"商务单公章识别为“{sup}”，与文件名/报价表不符，未采用"
        terms.pop('supplier', None)
        terms.pop('brand', None)
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
        brand = core[:6] if len(core) >= 2 else supplier[:4]
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

BUSINESS_HEADER_RE = re.compile(r'(单价|合价|总价|金额|小计|价格|市场价|出厂价|含税价|面价|单位|序号|税率|是否含税|品牌|' + '|'.join(REGION_HEADER_KEYS) + r'|图片)')


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
    return bool(v) and bool(NUMBER_RE.fullmatch(v.replace(' ', '')))


def _leading_number(v: str):
    m = re.match(r'\s*(\d+(?:\.\d+)?)', v or '')
    return float(m.group(1)) if m else None


def _detect_matrix(ws):
    """
    识别“矩阵式”报价表：行标题 (如规格档位) × 列标题 (如型号) → 每个交叉格是一个价格。
    判据：从第一行数据行开始，至少 2 个非首列的列以纯数字为主，且这些列的末级表头不是“单价/数量/合价”等清单表字段。
    返回 None 或 {'data_start': r0, 'label_col': c0, 'value_cols': [...], 'col_paths': {c: '表头路径'}, 'label_header': str, 'unit': str}
    """
    max_r, max_c = ws.max_row, ws.max_column
    grid = [[cell_str(ws.cell(r, c)) for c in range(1, max_c + 1)] for r in range(1, max_r + 1)]
    data_rows = [r for r in range(max_r) if sum(_is_number_text(v) for v in grid[r]) >= 2]
    if len(data_rows) < 3:
        return None
    r0 = data_rows[0]
    label_col = next((c for c in range(max_c) if sum(1 for r in data_rows if grid[r][c]) >= len(data_rows) * 0.6), 0)
    # 行标题同为数字、但整行无价格的行 (该档未报价) 仍属于数据区，不能并入表头
    while r0 > 0 and _leading_number(grid[r0 - 1][label_col]) is not None:
        r0 -= 1

    value_cols = []
    for c in range(label_col + 1, max_c):
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
    header_rows = [r for r in range(r0) if not _is_title_row(grid[r]) and not (table_title and all(grid[r][c] == table_title for c in value_cols))]

    def path(c):
        parts = []
        for r in header_rows:
            v = grid[r][c]
            if v and v != table_title and (not parts or parts[-1] != v):
                parts.append(v)
        return parts

    # 任意一级表头是清单表字段 (单价/合价/数量…) 就不是矩阵表
    for c in value_cols:
        if any(LIST_HEADER_RE.match(p) for p in path(c)):
            return None

    header_text = ' '.join(v for r in range(r0) for v in grid[r] if v)
    m_unit = re.search(r'单位[:：]\s*元?/?\s*([一-龥a-zA-Z²]+)', header_text)
    return {
        'data_start': r0 + 1,
        'label_col': label_col + 1,
        'value_cols': [c + 1 for c in value_cols],
        'col_paths': {c + 1: '-'.join(path(c)) or f"第{c + 1}列" for c in value_cols},
        'label_header': ''.join(path(label_col)) or '行标题',
        'table_title': table_title,
        'unit': m_unit.group(1) if m_unit else '',
    }


def _cell_notes(cell_issue, sname, r, cols, headers_map=None, price_col=None):
    """把识别存疑的单元格 (视觉模型双读不一致 / 自报看不清 / OCR 不一致) 转成待复核提示"""
    notes = []
    for c in cols:
        reason = cell_issue.get((sname, r, c))
        if reason:
            name = '单价' if c == price_col else (headers_map.get(c, f"第{c}列") if headers_map else f"第{c}列")
            notes.append(f"{name}{reason}")
    return notes


def _collect_matrix_rows(ws, sname, mx, cell_issue):
    """矩阵表展开：每个“行标题 × 列标题”交叉格生成一条原始物料行"""
    raw_rows, records = [], []
    prev = {}  # 列 -> (行标题数值, 价格)：用于“容量越大价格越高”的单调性校验
    for r in range(mx['data_start'], ws.max_row + 1):
        label = cell_str(ws.cell(r, mx['label_col']))
        if not label:
            continue
        label_num = _leading_number(label)
        for c in mx['value_cols']:
            v = cell_str(ws.cell(r, c))
            if not v:
                continue
            ref = f"{sname}!R{r}C{c}"
            col_path = mx['col_paths'][c]
            tt_part = f";表标题:{mx['table_title']}" if mx.get('table_title') else ""
            raw_line = f"{mx['label_header']}:{label} | 列:{col_path} | 价格:{v}"
            price = safe_clean_price(v) if _is_number_text(v) else 0.0
            decision = 'item' if price > 0 else f"矩阵单元格非数值: {v}"
            records.append(SourceRowRecord(source_ref=ref, sheet=sname, row=r, raw_text=raw_line,
                                           raw_price=price, decision=decision))
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
            raw_rows.append({
                "id": ref,
                "raw_cells_dict": {mx['label_header']: label, "产品/型号(表头)": col_path, "单价": v},
                "raw_full_line": raw_line,
                "raw_spec": f"{mx['label_header']}:{label};型号/系列:{col_path}" + tt_part,
                "raw_price": price,
                "raw_unit": mx['unit'],
                "mpaths": [],
                "file_page": sname,
                "table_name": sname,
                "notes": notes,
                "group_key": f"{sname}#C{c}",
            })
    return raw_rows, records


def _collect_list_rows(ws, ws_f, sname, standardizer, extractor, cell_issue, jev=None, jev_layout=None, struct_log=None):
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

        headers_map = _build_headers(ws, header_row, col_st, col_ed)
        header_values = set(headers_map.values())

        for r in range(data_start_row, ws.max_row + 1):
            row_cells_dict = {}
            row_text_parts = []
            row_values = []
            price_val = 0.0
            unit_val = ""
            formula_no_cache = False

            for c in range(col_st, col_ed + 1):
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

            if not any(row_values):
                continue
            ref = f"{sname}!R{r}#T{t_idx}"
            raw_line = " | ".join(row_text_parts)

            # 若未通过指定单价列拿到价格，按表头扫描单价列 (排除合价/金额，防止把总价当单价)
            if price_val <= 0 and not price_col:
                for c in range(col_st, col_ed + 1):
                    if _is_unit_price_header(headers_map.get(c, "")):
                        p_candidate = safe_clean_price(ws.cell(r, c).value)
                        if p_candidate > 0:
                            price_val = p_candidate
                            break

            # 行级过滤 (现有规则)：分组标题、重复表头、合计行
            decision = 'item'
            if _is_title_row(row_values):
                decision = '分组标题行'
            elif len(set(row_cells_dict.values()) & header_values) >= max(2, len(row_cells_dict) // 2):
                decision = '重复表头行'
            elif any(v.replace(' ', '') in SUMMARY_ROW_KEYS for v in row_values if v):
                decision = '合计/小计行'
            elif not (price_val > 0 or len(row_cells_dict) >= 2):
                decision = '无价格且有效属性不足2项'

            candidates.append({
                "ref": ref, "r": r, "values": row_values, "decision": decision, "raw_line": raw_line,
                "price": price_val, "formula_no_cache": formula_no_cache,
                "headers": [headers_map[c] for c in range(col_st, col_ed + 1)],
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
                    notes=layout_notes + _cell_notes(cell_issue, sname, r, range(col_st, col_ed + 1), headers_map, price_col),
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
            current_type = RULE_DECISION_TYPE[decision]
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
            raw_price=c["price"], decision=decision, has_formula_without_cache=c["formula_no_cache"]
        ))
        if decision != 'item':
            continue
        row = dict(c["build"])
        row["notes"] = row["notes"] + extra_notes
        raw_rows.append(row)
    return raw_rows, records


def standardize_workbook(wb_val, wb_fml, extractor, supplier, brand, id_prefix, terms_info,
                         cell_issue=None, global_notes=None, struct_log=None):
    """Excel 与 PDF(还原后的表格) 共用的解析主流程，返回 (canonical_items, source_rows)"""
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
    disc_ratio = terms_info["discount_ratio"]
    missing = missing_terms(terms_info)
    if missing:
        global_notes = global_notes + [f"商务条款缺失: {'、'.join(missing)}"]
    if disc_ratio is None:
        global_notes = global_notes + ["未提供工程折扣，工程价留空"]

    for sname in wb_val.sheetnames:
        ws = wb_val[sname]
        ws_f = wb_fml[sname]
        if ws.max_row is None or ws.max_row <= 1:
            continue
        print(f"\n--- [Sheet: {sname}] (共 {ws.max_row} 行 x {ws.max_column} 列) ---")

        mx = _detect_matrix(ws)
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
            raw_rows, records = _collect_matrix_rows(ws, sname, mx, cell_issue)
        else:
            raw_rows, records = _collect_list_rows(ws, ws_f, sname, standardizer, extractor, cell_issue,
                                                   jev=jev, jev_layout=jev_layout, struct_log=struct_log)
        source_rows.extend(records)
        group_of.update({r["id"]: r["group_key"] for r in raw_rows if r.get("group_key")})
        print(f"    -> 提取出有效物料原始行: {len(raw_rows)} 行，启动大模型槽位精准清洗...")

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
                    "raw_unit": item["raw_unit"]
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
                        row_brand = str(bv).strip()
                        break
                item_brand = row_brand or brand
                # 若来自表内明确品牌列，清除“由供应商名称推断”的提示
                if row_brand:
                    notes = [n for n in notes if not (n.startswith("品牌“") and "由供应商名称推断" in n)]

                clean_source = cleaned_res.get("_source", "fallback")
                if clean_source != "llm":
                    notes.append("未经大模型清洗(规则降级)，名称与参数需复核")

                std_name = str(cleaned_res.get("standard_name") or "").strip()
                if not std_name:
                    notes.append("未识别出物料名称")
                cat_code, cat_src = resolve_category_code(std_name, cleaned_res.get("category_code"))
                if cat_src == 'llm':
                    notes.append(f"类别编码 {cat_code} 未经标准编码库核验")
                elif not cat_code:
                    notes.append("类别编码未确定" + ("" if CATEGORY_CODE_MAP else " (未配置标准编码库 CATEGORY_CODE_FILE)"))

                structured_specs = cleaned_res.get("structured_specs")
                if not isinstance(structured_specs, dict) or not structured_specs:
                    structured_specs = {k: v for k, v in item["raw_cells_dict"].items()}
                structured_specs = {str(k): str(v) for k, v in structured_specs.items() if v not in (None, "")}
                # 原文已是规范 “键:值;” 参数时原样采用 (不让模型改参数名/值，如 “厚度(mm):1.5”“是否曝露使用”)，模型结果只补充原文没有的键
                verbatim = _verbatim_specs(item["raw_cells_dict"])
                if verbatim:
                    covered = {_spec_key_norm(k) for k in verbatim}
                    extra = {k: v for k, v in structured_specs.items()
                             if _spec_key_norm(k) not in covered and v not in verbatim.values()}
                    structured_specs = {**verbatim, **extra}

                # F7 结构化参数逐项核对原文出处
                unproven_params = _check_param_provenance(structured_specs, item.get("raw_full_line", ""))
                if unproven_params:
                    notes.append("参数无原文出处: " + "；".join(unproven_params))

                # 行内自带报价地区时以行为准，否则用商务条款中的地区
                row_region = next((v for k, v in item["raw_cells_dict"].items() if any(rk in k for rk in REGION_HEADER_KEYS)), '')
                province, city, region_note = _region_fields(row_region or terms_info["province"], terms_info["city"])
                if region_note:
                    notes.append(region_note)

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
                    raw_spec=item.get("raw_spec") or item["raw_full_line"],
                    file_page=item["file_page"],
                    unit=unit,
                    market_price=final_price,
                    discount_ratio=disc_ratio,
                    project_price=round(final_price * disc_ratio, 2) if disc_ratio is not None and final_price > 0 else None,
                    discount_series=terms_info["discount_series"],
                    is_freight_included=terms_info["is_freight_included"],
                    freight_desc=terms_info["freight_desc"],
                    is_tax_included=terms_info["is_tax_included"],
                    tax_desc=terms_info["tax_desc"],
                    tax_rate=terms_info["tax_rate"],
                    freight_settle=terms_info["freight_settle"],
                    quote_desc=terms_info["quote_desc"],
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
    _unify_group_names(canonical_items, group_of)
    if jev is not None:
        u = jev.usage
        print(f"[*] Jev 用量: 请求 {u['requests']} 次 (失败 {u['failed']})，输入 {u['input_tokens']} tokens，输出 {u['output_tokens']} tokens")
    return canonical_items, source_rows


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
            m.category_name, m.category_code, m.category_source = name, code, src


def _region_fields(region_text, city_text=''):
    """把地区原文展开为 (省, 市, 提示)；无法识别的原文照抄"""
    region_text = str(region_text or '').strip()
    if not region_text:
        return '', '', ''
    provs = expand_region(region_text)
    if not provs:
        return region_text, city_text or '', f"报价地区“{region_text}”无法识别为省级行政区"
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
        unmerge_and_fill_cells(wb_val[sname])
        unmerge_and_fill_cells(wb_fml[sname])

    struct_log = []
    items, source_rows = standardize_workbook(wb_val, wb_fml, extractor, supplier, brand, id_prefix,
                                              normalize_terms(terms), global_notes=_terms_notes(terms) + sup_notes,
                                              struct_log=struct_log)
    return _finish(items, source_rows, template_xlsx, output_xlsx, img_dir, struct_log)


def _process_pdf_with_extractor(extractor_cls, input_pdf, template_xlsx, output_dir, terms=None, rail_name="A-Rail"):
    print(f"\n=== [PDF Pipeline - {rail_name}] 启动解析: {input_pdf} ===")
    pdf_ext = extractor_cls(input_pdf)
    wb = pdf_ext.to_workbook()

    terms = verify_terms_supplier(terms, _path_evidence(input_pdf) + pdf_ext.page_titles + _workbook_text(wb))
    if not terms.get('supplier') and not supplier_from_path(input_pdf):
        guessed = pdf_ext.guess_supplier()
        if guessed:
            terms['supplier'] = guessed
    supplier, brand, id_prefix, sup_notes = derive_supplier_and_brand_dynamic(input_pdf, terms)
    print(f"[*] 动态感知供方主体: {supplier} | 核心品牌: {brand} | 报价ID前缀: {id_prefix}")
    output_xlsx, img_dir = _output_paths(input_pdf, output_dir, supplier)

    global_notes = _terms_notes(terms) + sup_notes
    struct_log = []
    items, source_rows = standardize_workbook(wb, wb, None, supplier, brand, id_prefix, normalize_terms(terms),
                                              cell_issue=pdf_ext.cell_issue, global_notes=global_notes,
                                              struct_log=struct_log)
    if hasattr(pdf_ext, 'failed_pages') and pdf_ext.failed_pages:
        for fpage, reason in pdf_ext.failed_pages:
            rec = SourceRowRecord(
                sheet_name=fpage,
                row_idx=1,
                row_type="exception",
                source_ref=f"{fpage}!R1",
                raw_full_line=f"异常:{fpage}识别失败({reason})",
                raw_price=None,
                skip_reason=f"异常:{fpage}识别失败({reason})"
            )
            source_rows.append(rec)
        print(f"[*] 页面失败隔离清单: {[fp for fp, _ in pdf_ext.failed_pages]}")
    # 每页原件截图作为“厂家报价图片” (quote_img{页码}.png)，便于对照
    page_imgs = pdf_ext.save_page_images(img_dir)
    for it in items:
        m = re.match(r'第(\d+)页', it.file_page)
        if m and int(m.group(1)) in page_imgs:
            it.image_filenames = [page_imgs[int(m.group(1))]]
    return _finish(items, source_rows, template_xlsx, output_xlsx, img_dir, struct_log)


def process_pdf(input_pdf, template_xlsx, output_dir, terms=None):
    """
    双轨调度中枢:
    1. EXTRACTOR_ENGINE='gemini': 优先使用 B 轨 (Gemini 3.8 Flash 原生直传)，若未配置Key或异常自动无损回退到 A 轨
    2. EXTRACTOR_ENGINE='traditional': 强制使用 A 轨 (传统切片 + OpenCV 滤镜 + DeepSeek)
    3. EXTRACTOR_ENGINE='compare': 双轨比对模式
    """
    engine = (EXTRACTOR_ENGINE or "traditional").strip().lower()

    if engine in ("gemini", "compare"):
        if GEMINI_API_KEY:
            try:
                print(f"[*] [双轨调度] 路由至 B 轨 (Gemini 原生全模态: {GEMINI_MODEL})")
                return _process_pdf_with_extractor(GeminiPDFTableExtractor, input_pdf, template_xlsx, output_dir,
                                                    terms=terms, rail_name="Gemini B-Rail")
            except Exception as e:
                print(f"[Warning] [双轨调度] Gemini B 轨执行异常: {e}，自动平滑兜底回退至 A 轨...")
                return _process_pdf_with_extractor(PDFTableExtractor, input_pdf, template_xlsx, output_dir,
                                                    terms=terms, rail_name="Traditional A-Rail (Fallback)")
        else:
            print("[Info] [双轨调度] 检测到 EXTRACTOR_ENGINE=gemini 但未配置 GEMINI_API_KEY，安全直通 A 轨 (传统切图流水线)")
            return _process_pdf_with_extractor(PDFTableExtractor, input_pdf, template_xlsx, output_dir,
                                                terms=terms, rail_name="Traditional A-Rail")
    else:
        print("[*] [双轨调度] 路由至 A 轨 (传统切图流水线)")
        return _process_pdf_with_extractor(PDFTableExtractor, input_pdf, template_xlsx, output_dir,
                                            terms=terms, rail_name="Traditional A-Rail")


def _finish(items, source_rows, template_xlsx, output_xlsx, img_dir, struct_log=None):
    # 原表 ↔ 结果逐行对账
    recon = reconcile(source_rows, items)
    _print_recon_summary(recon)
    print(f"\n[*] 全表语义解析完成，共生成标准物料: {len(items)} 项，启动构建器落盘...")
    builder = StandardizedWorkbookBuilder(template_xlsx)
    builder.build_from_canonical_items(items, output_xlsx, img_dir, reconciliation=recon,
                                       structure_log=struct_log or None)
    return output_xlsx


def _print_recon_summary(recon: list):
    total_items = sum(1 for r in recon if r['status'] != '跳过')
    bad = [r for r in recon if r['status'].startswith('异常')]
    skipped = sum(1 for r in recon if r['status'] == '跳过')
    print(f"[*] 原表对账: 物料行 {total_items} | 跳过行 {skipped} | 异常 {len(bad)}")
    for r in bad[:10]:
        print(f"    ! {r['source_ref']}: {r['status']}")


def main():
    parser = argparse.ArgumentParser(description="QuoteStandardizer - 报价单解析与原表对账")
    parser.add_argument("--input", "-i", required=True, help="输入报价文件路径 (.xlsx / .pdf)")
    parser.add_argument("--terms", "-t", default=None, help="商务说明与折扣记录单图片路径 (.png/.jpg)；不填时自动在输入文件同目录查找")
    parser.add_argument("--terms-json", default=None, help="人工确认的商务条款 JSON (优先于图片识别)")
    parser.add_argument("--template", default=DEFAULT_TEMPLATE_PATH, help="输出模板 .xlsx (取第 1 行表头)；不填时使用内置标准表头")
    parser.add_argument("--output_dir", "-o", default=DEFAULT_OUTPUT_DIR, help="输出目录")
    args = parser.parse_args()

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

    print(f"\n🎉 统一标准多行结构化输出完成！交付文件：{out}")
    # 若存在隔离的失败页，返回退出码 2
    if 'failed_pages' in locals() or (hasattr(sys, '_has_failed_pages') and sys._has_failed_pages):
        sys.exit(2)

if __name__ == '__main__':
    main()
