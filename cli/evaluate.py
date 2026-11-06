import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
评差引擎：把本地解析结果与线上系统的结果逐格比对，按「单份文件零差异才算 PASS」的严口径出报告。

口径说明（客户的验收标准）：
  · 文件级：一份文件里任何一处对不上 → 该文件 FAIL。10 份文件 9 份各错 1 处 = 1/10 PASS，不是 99% 准确率。
  · 字段级：仅作参考，不作为验收依据，打印出来是为了让人看清"平均值会把问题藏起来"。
  · 差异 ≠ 我们错。差异只说明两套系统不一致，需对照原件裁决；裁决结果写进 verdicts.csv，
    之后回归会按裁决算"谁的错"。引擎绝不替人判定真相。

命令：
  python evaluate.py compare  --ours A.xlsx --theirs B.xlsx [-o out/] [--verdicts v.csv]
  python evaluate.py batch    --cases cases.csv   [-o out/] [--verdicts v.csv]     # 一批文件
  python evaluate.py archive  --cases cases.csv --dest fixtures/                   # 存成回归用例
  python evaluate.py regression [--update]                                        # 全量回归，任一 FAIL 退出码 1
  python evaluate.py selftest                                                     # 引擎自检

cases.csv 列：file_id,ours,theirs（首行表头；路径相对 CSV 所在目录）
verdicts.csv 列：file_id,field,ours,theirs,truth  (truth = ours / theirs / both_wrong / equiv)
"""

import argparse
import csv
import json
import os
import re
import sys
import time
import unicodedata
from collections import Counter, defaultdict

import openpyxl

# ---------------------------------------------------------------- 错误码字典
# 每个码都绑定"该去查哪里"，报表直接告诉人下一个动作，而不是从头排查。
CODE_META = {
    'E00_EQUIV': ('写法差异(已归一，不判错)', 'INFO', '无：数值/全半角/单位同义词等价'),
    'E01_ROW_OURS_MISSING': ('我们少一条', 'FATAL', 'main._collect_list_rows 行过滤 / pdf 切块去重 / 页丢失'),
    'E02_ROW_OURS_EXTRA': ('我们多一条', 'FATAL', 'main._detect_matrix 矩阵误判 / 表头行判错 / 合计行被当物料'),
    'E03_PRICE': ('价格不一致', 'FATAL', 'extractor.safe_clean_price / 单价列判定 price_col / 矩阵交叉格取值'),
    'E04_UNIT': ('单位不一致', 'FATAL', 'main.normalize_unit / unit_col 判定'),
    'E05_NAME_TYPO': ('名称错字(编辑距离小)', 'FATAL', 'llm_agent 清洗 / 视觉模型转录 / 合并单元格展开'),
    'E06_NAME_DIFF': ('名称语义不同', 'FATAL', 'main._unify_group_names 多数票统一 / 表头路径拼接'),
    'E07_SPEC_MISSING': ('规格参数缺项/多项', 'FATAL', 'main._verbatim_specs / _genoffice_raw_spec / 单元格混写拆分'),
    'E08_SPEC_VALUE': ('规格参数值不一致', 'FATAL', 'llm_agent 清洗(参数值必须照抄原文) / 视觉转录'),
    'E09_CODE': ('类别编码不一致', 'ERROR', 'config.resolve_category_code / CATEGORY_CODE_FILE 未配置'),
    'E10_BRAND': ('品牌不一致', 'ERROR', 'main.derive_supplier_and_brand_dynamic / 表内品牌列'),
    'E11_TERMS': ('商务条款不一致', 'ERROR', 'terms_extractor / 商务单图片识别 / 折扣解析'),
    'E12_REGION': ('报价地区不一致', 'ERROR', 'config.expand_region / 行内地区列'),
    'E15_PAGE': ('原表页码不一致', 'ERROR', 'pdf_extractor 页序 / gemini page_number / 切块合并'),
    'E14_STRUCTURE': ('疑似整表结构判错(已折叠逐行差异)', 'FATAL', 'main._detect_matrix / _collect_list_rows 表头行'),
    'E16_HEADER': ('输出表结构/列名对不上', 'FATAL', '线上系统列名与标准 33 列映射(ALIASES)'),
}

# 参与判定的字段：(字段名, 错误码, 归一类型, 容差)
FIELDS = [
    ('名称规格', 'name', 'name', 0),
    ('品牌', 'text', 'E10_BRAND', 0),
    ('类别编码', 'text', 'E09_CODE', 0),
    ('规格参数', 'spec', 'E08_SPEC_VALUE', 0),
    ('单位', 'unit', 'E04_UNIT', 0),
    ('市场价', 'num', 'E03_PRICE', 1e-6),
    ('工程价', 'num', 'E03_PRICE', 0.011),
    ('工程价折扣比例', 'ratio', 'E11_TERMS', 1e-6),
    ('报价地区省', 'text', 'E12_REGION', 0),
    ('报价地区市', 'text', 'E12_REGION', 0),
    ('折扣系列', 'text', 'E11_TERMS', 0),
    ('折扣下限', 'ratio', 'E11_TERMS', 1e-6),
    ('折扣上限', 'ratio', 'E11_TERMS', 1e-6),
    ('是否含运费', 'yn', 'E11_TERMS', 0),
    ('运费说明', 'text', 'E11_TERMS', 0),
    ('运费结算方式', 'text', 'E11_TERMS', 0),
    ('是否含增值税', 'yn', 'E11_TERMS', 0),
    ('税费说明', 'text', 'E11_TERMS', 0),
    ('增值税税率', 'tax', 'E11_TERMS', 1e-6),
    ('价格填报类型', 'text', 'E11_TERMS', 0),
    ('厂家报价文件页码', 'page', 'E15_PAGE', 0),
]
# 不参与判定：编号规则各系统不同 / 自由文本自带本系统提示 / 图片列为附件
IGNORE_FIELDS = ['报价ID', '顺序号', '厂家原始规格参数', '报价说明', '图片文件夹', '图片']

# 线上系统列名可能不同：按别名解析成标准列名
ALIASES = {
    '名称规格': ['名称规格', '物料名称', '材料名称', '品名', '名称'],
    '规格参数': ['规格参数', '技术参数', '规格型号', '规格'],
    '市场价': ['市场价', '单价', '含税单价', '原表单价', '报价'],
    '工程价': ['工程价', '折后价', '工程单价'],
    '工程价折扣比例': ['工程价折扣比例', '折扣比例', '折扣率'],
    '报价地区省': ['报价地区省', '地区省', '省份', '报价地区'],
    '报价地区市': ['报价地区市', '地区市', '城市'],
    '厂家报价文件页码': ['厂家报价文件页码', '页码', '原表页码', '文件页码'],
    '类别编码': ['类别编码', '分类编码', '物料编码', '编码'],
    '增值税税率': ['增值税税率', '增值税税率%', '税率', '增值税率'],
    '是否含运费': ['是否含运费', '含运费'],
    '是否含增值税': ['是否含增值税', '含税'],
    '厂家报价图片文件名': ['厂家报价图片文件名', '图片文件名', '切片文件名', '图片名', '凭证图片文件名'],
    '报价ID': ['报价ID', '报价id', '报价Id'],
    '供应商名称': ['供应商名称', '供应商', '供方'],
}
# 中文标点 → 半角（NFKC 不处理这些，业务上属排版差异不算错）
PUNCT = {'，': ',', '：': ':', '；': ';', '（': '(', '）': ')', '【': '[', '】': ']',
         '“': '"', '”': '"', '‘': "'", '’': "'", '　': ' ', '－': '-', '－': '-'}
UNIT_MAP = {'㎡': 'm2', '平方米': 'm2', '平米': 'm2', 'm²': 'm2', '㎥': 'm3', 'm³': 'm3', '立方米': 'm3',
            '米': 'm', '公斤': 'kg', '千克': 'kg', '个': '个'}
YES = {'是', '包含', '含', '有', '√', '1', 'true'}
NO = {'否', '不含', '不包含', '无', '×', '0', 'false'}


# ---------------------------------------------------------------- 归一化
def nfkc(s):
    return unicodedata.normalize('NFKC', str(s if s is not None else ''))


def canon_text(s):
    """文本归一：NFKC + 去全部空白 + 中文标点转半角。逐字不同才算不同（不模糊匹配）"""
    s = nfkc(s)
    for k, v in PUNCT.items():
        s = s.replace(k, v)
    return re.sub(r'\s+', '', s)


def canon_num(s):
    """数值归一：'¥10,500' / '10500.00' / '10 500' → 10500.0；空/非数字 → None"""
    if s is None or s == '':
        return None
    if isinstance(s, bool):
        return None
    if isinstance(s, (int, float)):
        return float(s)
    t = re.sub(r'[¥$￥,\s、]', '', nfkc(s))
    m = re.search(r'-?\d+(?:\.\d+)?', t)
    return float(m.group()) if m else None


def canon_ratio(s):
    """折扣归一：0.9 / '90%' / '9折' → 0.9"""
    if s is None or s == '':
        return None
    t = nfkc(s)
    x = canon_num(t)
    if x is None:
        return None
    if '折' in t and x > 1:
        x = x / 10
    if x > 1:
        x = x / 100 if x <= 100 else x
    return round(x, 6)


def canon_tax(s):
    """税率归一：13 / '13%' / 0.13 → 13.0"""
    x = canon_ratio(s)
    return None if x is None else round(x * 100, 6) if x <= 1 else x


def canon_yn(s):
    t = canon_text(s)
    if not t:
        return None
    if t in YES:
        return '是'
    if t in NO:
        return '否'
    if '不含' in t or '不包含' in t:
        return '否'
    if '含' in t or '包' in t:
        return '是'
    return t


def canon_unit(s):
    t = canon_text(s)
    return UNIT_MAP.get(t, t.lower())


def canon_page(s):
    t = canon_text(s)
    m = re.search(r'(\d+)', t)
    return int(m.group(1)) if m else (t or None)


def canon_value(kind, raw):
    return {'num': canon_num, 'ratio': canon_ratio, 'tax': canon_tax, 'yn': canon_yn,
            'unit': canon_unit, 'page': canon_page, 'text': canon_text}.get(kind, canon_text)(raw)


def parse_specs(s):
    """'容量kVA:30\\n产品/型号:S13型（NX3）' → {'容量kva':'30', '产品/型号':'s13型(nx3)'}"""
    out = {}
    for line in str(s or '').split('\n'):
        for kv in re.split(r'[;；]', line):
            if ':' in nfkc(kv) or '：' in kv:
                k, v = re.split(r'[:：]', kv, maxsplit=1)
                k, v = canon_text(k), canon_text(v)
                if k and v:
                    out[k] = v
    return out


def spec_sig(s):
    """规格签名，用于行对齐：排好序的 键=值 串"""
    return '|'.join(f'{k}={v}' for k, v in sorted(parse_specs(s).items()))


def lev(a, b):
    """编辑距离（两行滚动数组，短串够用）"""
    if a == b:
        return 0
    if not a or not b:
        return len(a) or len(b)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


def sim(a, b):
    if not a and not b:
        return 1.0
    if not a or not b:
        return 0.0
    return 1 - lev(a, b) / max(len(a), len(b))


# ---------------------------------------------------------------- 读表
def _pick_sheet(wb, want=None):
    if want and want in wb.sheetnames:
        return want
    for name in wb.sheetnames:
        if any(k in name for k in ('标准化', '标准解析', '报价')):
            return name
    return wb.sheetnames[0]


def load_table(path, sheet=None):
    """读一张标准报价表 → (rows, problems)。rows 为标准字段名的 dict 列表"""
    problems = []
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        return [], [f'无法打开 {path}: {e}']
    ws = wb[_pick_sheet(wb, sheet)]
    header = [str(ws.cell(1, c).value or '').strip() for c in range(1, (ws.max_column or 0) + 1)]
    col = {}
    for idx, h in enumerate(header):
        if not h:
            continue
        for std in [f[0] for f in FIELDS] + IGNORE_FIELDS + ['厂家报价图片文件名']:
            names = ALIASES.get(std, [std])
            if std not in col and h in names:
                col[std] = idx
                break

    known_col_indices = set(col.values())
    dyn_spec_cols = [(idx, h) for idx, h in enumerate(header) if idx not in known_col_indices and h and not any(k in h for k in ('图片', '文件夹', '序号', '项目类型', '工程用量', '报价时间'))]

    sheet_cat_code = None
    if '类别编码' not in col:
        s_name = str(ws.title).strip()
        if re.match(r'^[A-Za-z0-9\-_]{4,}$', s_name) and s_name not in ('Sheet', 'Sheet1'):
            sheet_cat_code = s_name

    missing = [f[0] for f in FIELDS if f[0] not in col and f[0] not in ('厂家报价文件页码',)]
    if '规格参数' in missing and dyn_spec_cols:
        missing.remove('规格参数')
    if '类别编码' in missing and sheet_cat_code:
        missing.remove('类别编码')

    if '名称规格' not in col and '规格参数' not in col and not dyn_spec_cols:
        problems.append(f'E16_HEADER: 表头没找到物料名列，实际表头: {header[:8]}…')
    if missing:
        problems.append(f'E16_HEADER: 缺列 {missing} (该列按"双方都空"处理)')
    rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, '') for v in r):
            continue
        row_dict = {std: (r[i] if i < len(r) else None) for std, i in col.items()}
        if '规格参数' not in row_dict or not row_dict['规格参数']:
            if dyn_spec_cols:
                parts = [f"{h}:{r[i]}" for i, h in dyn_spec_cols if i < len(r) and r[i] not in (None, '')]
                if parts:
                    row_dict['规格参数'] = ';'.join(parts)
        if ('类别编码' not in row_dict or not row_dict['类别编码']) and sheet_cat_code:
            row_dict['类别编码'] = sheet_cat_code
        rows.append(row_dict)
    return rows, problems


# ---------------------------------------------------------------- 行对齐
def _sig(row):
    """行身份签名：价格 + 规格。名称在多型号报表里常常整表相同，不能单独作为身份"""
    return (canon_num(row.get('市场价')), spec_sig(row.get('规格参数')))


def align(A, B):
    """A=我们的行, B=他们的行。返回 (pairs, only_A, only_B)。
    分轮匹配：同价同规格 → 同规格 → 同价同名 → 相似度贪心，避免少一条导致后面整体错位成 N 条差异。"""
    pairs, used_a, used_b = [], set(), set()

    def run(keyf, cond=None):
        idx = defaultdict(list)
        for j, r in enumerate(B):
            if j not in used_b:
                idx[keyf(r)].append(j)
        for i, r in enumerate(A):
            if i in used_a:
                continue
            k = keyf(r)
            while idx.get(k):
                j = idx[k].pop(0)
                if j in used_b:
                    continue
                if cond and not cond(r, B[j]):
                    continue
                pairs.append((i, j)); used_a.add(i); used_b.add(j)
                break

    def _exact_img_sig(r):
        """四级级联第一级：切片图片文件名 / (页码+价格+规格) 物理原件切片精确对齐"""
        img = str(r.get('厂家报价图片文件名') or '').strip()
        p = str(r.get('厂家报价文件页码') or '').strip()
        pr = canon_num(r.get('市场价'))
        sp = spec_sig(r.get('规格参数'))
        if img and pr and sp:
            return ('img_pr_sp', img, pr, sp)
        if p and pr and sp:
            return ('page_pr_sp', p, pr, sp)
        return None

    run(lambda r: _exact_img_sig(r) or ('__none__', id(r)))
    run(_sig)
    run(lambda r: spec_sig(r.get('规格参数')))
    run(lambda r: (canon_num(r.get('市场价')), canon_text(r.get('名称规格'))))
    # 最后一轮：相似度贪心
    cand = []
    for i, a in enumerate(A):
        if i in used_a:
            continue
        for j, b in enumerate(B):
            if j in used_b:
                continue
            score = (0.5 * sim(spec_sig(a.get('规格参数')), spec_sig(b.get('规格参数')))
                     + 0.3 * sim(canon_text(a.get('名称规格')), canon_text(b.get('名称规格')))
                     + 0.2 * (1.0 if canon_num(a.get('市场价')) == canon_num(b.get('市场价')) else 0.0))
            cand.append((score, i, j))
    for score, i, j in sorted(cand, reverse=True):
        if score < 0.45:
            break
        if i in used_a or j in used_b:
            continue
        pairs.append((i, j)); used_a.add(i); used_b.add(j)

    return (sorted(pairs), [i for i in range(len(A)) if i not in used_a],
            [j for j in range(len(B)) if j not in used_b])


# ---------------------------------------------------------------- 逐格比对
def _row_label(row, i):
    """人工定位用：用见过的最大信息量描述这一行"""
    parts = [str(row.get('名称规格') or ''), spec_sig(row.get('规格参数'))[:40],
             f"价格={row.get('市场价')}"]
    return f'#{i + 1} ' + ' '.join(p for p in parts if p)


def check_math_consistency(rows: list, tol: float = 0.01) -> dict:
    """
    数学勾稽 warning 模式：
    当行同时存在 qty、unit_price/price、amount/total 时检查 abs(qty*price - amount) <= max(tol, amount * 0.002)
    仅告警与统计，绝对不改值。
    """
    audited = 0
    suspects = []
    for idx, r in enumerate(rows, 1):
        p_val = canon_num(r.get('市场价') or r.get('单价') or r.get('price'))
        q_val = canon_num(r.get('数量') or r.get('qty'))
        a_val = canon_num(r.get('合价') or r.get('金额') or r.get('amount') or r.get('总价'))

        if p_val is not None and p_val > 0 and q_val is not None and q_val > 0 and a_val is not None and a_val > 0:
            audited += 1
            expected = p_val * q_val
            diff = abs(expected - a_val)
            allowed_tol = max(tol, a_val * 0.002)
            if diff > allowed_tol:
                suspects.append({
                    "row": idx,
                    "price": p_val,
                    "qty": q_val,
                    "amount": a_val,
                    "expected": round(expected, 2),
                    "diff": round(diff, 2)
                })

    return {
        "audited_rows": audited,
        "suspect_count": len(suspects),
        "suspects": suspects
    }


def diff_pair(a, b, label):
    diffs = []
    for field, kind, code, tol in FIELDS:
        va, vb = a.get(field), b.get(field)
        if va in (None, '') and vb in (None, ''):
            continue
        if kind == 'spec':
            sa, sb = parse_specs(va), parse_specs(vb)
            for k in sorted(set(sa) | set(sb)):
                if k not in sa:
                    diffs.append(dict(label=label, field=f'规格参数.{k}', ours='', theirs=sb[k],
                                      code='E07_SPEC_MISSING', level='FATAL', note='我们缺这个参数'))
                elif k not in sb:
                    diffs.append(dict(label=label, field=f'规格参数.{k}', ours=sa[k], theirs='',
                                      code='E07_SPEC_MISSING', level='FATAL', note='他们缺这个参数'))
                elif sa[k] != sb[k]:
                    diffs.append(dict(label=label, field=f'规格参数.{k}', ours=sa[k], theirs=sb[k],
                                      code='E08_SPEC_VALUE', level='FATAL', note='参数值不同'))
            continue
        ca, cb = canon_value(kind, va), canon_value(kind, vb)
        if ca is None and cb is None:
            continue

        # P0-4 步①过渡等价：针对折扣比例字段，一边为空且另一边为 1.0 时视为等价，不判错
        if field in ('工程价折扣比例', '折扣下限', '折扣上限'):
            if (ca is None and cb == 1.0) or (ca == 1.0 and cb is None):
                diffs.append(dict(label=label, field=field, ours=va, theirs=vb,
                                  code='E00_EQUIV', level='INFO', note='空折扣与1.0过渡等价'))
                continue

        if ca is None or cb is None:
            # 一边有值一边为空：本质上是缺失（E03/E11 等），而不是写法差异
            diffs.append(dict(label=label, field=field, ours=va, theirs=vb, code=code,
                              level=CODE_META[code][1], note='一边为空'))
            continue
        same = (abs(float(ca) - float(cb)) <= tol) if (isinstance(ca, float) or isinstance(cb, float)) else (ca == cb)
        if same:
            # 归一后一致但原文不同 → 记一条 INFO，让审计者看得到引擎做的每一次宽容
            if str(va) != str(vb):
                diffs.append(dict(label=label, field=field, ours=va, theirs=vb,
                                  code='E00_EQUIV', level='INFO', note='写法不同，归一后一致'))
            continue
        if field == '名称规格':
            code = 'E05_NAME_TYPO' if sim(ca or '', cb or '') >= 0.6 else 'E06_NAME_DIFF'
        diffs.append(dict(label=label, field=field, ours=va, theirs=vb, code=code,
                          level=CODE_META[code][1]))
    return diffs


def eval_file(file_id, ours_path, theirs_path, ours_sheet=None, theirs_sheet=None):
    res = dict(file_id=file_id, ours=ours_path, theirs=theirs_path, diffs=[], problems=[])
    A, pa = load_table(ours_path, ours_sheet)
    B, pb = load_table(theirs_path, theirs_sheet)
    res['problems'] = pa + pb
    res['rows_ours'], res['rows_theirs'] = len(A), len(B)
    if pa or pb:
        for p in pa + pb:
            code = p.split(':')[0]
            res['diffs'].append(dict(label='(文件级)', field='-', ours='', theirs='', code=code,
                                     level=CODE_META.get(code, ('', 'FATAL', ''))[1], note=p))
    pairs, only_a, only_b = align(A, B)
    for i, j in pairs:
        res['diffs'].extend(diff_pair(A[i], B[j], _row_label(A[i], i)))
    for i in only_a:
        res['diffs'].append(dict(label=_row_label(A[i], i), field='(整行)', ours=json.dumps(
            {k: str(A[i].get(k)) for k in ('名称规格', '规格参数', '市场价')}, ensure_ascii=False),
            theirs='', code='E01_ROW_OURS_MISSING', level='FATAL', note='我们有、他们没有'))
    for j in only_b:
        res['diffs'].append(dict(label=_row_label(B[j], j), field='(整行)', ours='', theirs=json.dumps(
            {k: str(B[j].get(k)) for k in ('名称规格', '规格参数', '市场价')}, ensure_ascii=False),
            code='E02_ROW_OURS_EXTRA', level='FATAL', note='他们有、我们没有'))
    # 整表结构判错的折叠：行/字段级错位会刷出几十条噪音，先聚合成一条，明细仍在报告里
    fatal_rows = {d['label'] for d in res['diffs'] if d['level'] == 'FATAL'}
    total = max(len(A), len(B), 1)
    if len(fatal_rows) >= max(3, total * 0.2):
        res['diffs'].insert(0, dict(label='(文件级)', field='-', ours=f'{len(A)} 行', theirs=f'{len(B)} 行',
                                    code='E14_STRUCTURE', level='FATAL',
                                    note=f'受影响行 {len(fatal_rows)}/{total}，疑似整表结构判错(表头行/矩阵判定/漏页)，逐行明细见下'))
    res['fatal'] = sum(1 for d in res['diffs'] if d['level'] == 'FATAL')
    res['errors'] = sum(1 for d in res['diffs'] if d['level'] != 'INFO')
    res['equiv'] = sum(1 for d in res['diffs'] if d['code'] == 'E00_EQUIV')
    res['status'] = 'PASS' if res['errors'] == 0 else 'FAIL'
    codes = Counter(d['code'] for d in res['diffs'] if d['level'] != 'INFO')
    # 结构错是根因：它会间接刷出上百条行级差异，报表上必须让它当主错误码
    res['top_code'] = 'E14_STRUCTURE' if 'E14_STRUCTURE' in codes else (codes.most_common(1)[0][0] if codes else '-')
    return res


# ---------------------------------------------------------------- 报告
def summarize(results):
    n = len(results)
    ok = sum(1 for r in results if r['status'] == 'PASS')
    fields_total = sum(max(r['rows_ours'], r['rows_theirs'], 1) * len(FIELDS) for r in results)
    field_bad = sum(r['errors'] for r in results)
    codes = Counter()
    for r in results:
        codes.update(d['code'] for d in r['diffs'] if d['level'] != 'INFO')
    return dict(files=n, passed=ok, failed=n - ok,
                file_rate=(ok / n * 100) if n else 0.0,
                field_rate=(1 - field_bad / fields_total) * 100 if fields_total else 100.0,
                codes=codes,
                equiv=sum(r['equiv'] for r in results))


def print_report(results, verdicts=None, vstats=None):
    s = summarize(results)
    print("\n=== 评差结果（严口径：单份文件零差异才 PASS）===")
    print(f"{'文件':32s}{'结论':6s}{'行(我们/他们)':16s}{'差异':6s}{'主错误码':22s}裁决")
    for r in results:
        v = (verdicts or {}).get(r['file_id'], '')
        print(f"{r['file_id'][:30]:32s}{r['status']:6s}"
              f"{str(r['rows_ours']) + '/' + str(r['rows_theirs']):16s}{r['errors']:<6d}{r['top_code']:22s}{v}")
        for d in r['diffs'][:2]:
            if d['code'] == 'E14_STRUCTURE':
                print(f"    └ {d['note']}")
    print(f"\n文件级：{s['passed']}/{s['files']} PASS（{s['file_rate']:.1f}%）   ← 验收口径：任一处对不上即 FAIL")
    print(f"字段级：{s['field_rate']:.2f}%（仅参考，不作为验收；会把个别文件的硬伤平均掉）")
    if s['equiv']:
        print(f"归一化抹平：{s['equiv']} 处写法差异（E00，已按数值/全半角/单位同义词等价，明细见报告）")
    if s['codes']:
        print('\n--- 错误归纳 ---')
        for code, cnt in s['codes'].most_common():
            name, level, hint = CODE_META.get(code, (code, '', ''))
            files = sum(1 for r in results if any(d['code'] == code for d in r['diffs']))
            print(f"  {code:24s}{cnt:5d} 次 / {files:3d} 份文件  {name:34s}→ 查: {hint}")
    print_verdict_stats(results, vstats, verdicts)


def write_report(results, outdir, verdicts=None, vstats=None):
    if outdir.endswith('.xlsx'):
        path = outdir
        parent_dir = os.path.dirname(outdir)
        if parent_dir:
            os.makedirs(parent_dir, exist_ok=True)
    else:
        os.makedirs(outdir, exist_ok=True)
        path = os.path.join(outdir, '评差报告.xlsx')
    s = summarize(results)
    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = '文件级结果'
    ws.append([f"文件级 {s['passed']}/{s['files']} PASS（{s['file_rate']:.1f}%）| "
               f"验收口径：单份文件零差异才 PASS | 验收结论 {'PASS' if s['failed'] == 0 else 'FAIL'}"])
    ws.append(['文件ID', '结论', '我们行数', '他们行数', '差异数', 'FATAL', '主错误码', '裁决', '我们的文件', '他们的文件'])
    for r in results:
        ws.append([r['file_id'], r['status'], r['rows_ours'], r['rows_theirs'], r['errors'], r['fatal'],
                   r['top_code'], (verdicts or {}).get(r['file_id'], ''), r['ours'], r['theirs']])
    for col, w in zip('ABCDEFGHIJ', [34, 8, 10, 10, 8, 8, 24, 14, 40, 40]):
        ws.column_dimensions[col].width = w

    ws2 = wb.create_sheet('差异明细')
    ws2.append(['文件ID', '行定位', '字段', '我们的值', '他们的值', '错误码', '级别', '含义', '该查哪里', '备注'])
    for r in results:
        for d in r['diffs']:
            name, level, hint = CODE_META.get(d['code'], (d['code'], '', ''))
            ws2.append([r['file_id'], d['label'], d['field'], str(d['ours'])[:200], str(d['theirs'])[:200],
                        d['code'], d['level'], name, hint, d.get('note', '')])
    for col, w in zip('ABCDEFGHIJ', [28, 34, 22, 34, 34, 24, 8, 30, 46, 20]):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = 'A2'

    ws3 = wb.create_sheet('错误归纳')
    ws3.append(['错误码', '含义', '级别', '次数', '涉及文件数', '建议排查位置'])
    for code, cnt in s['codes'].most_common():
        name, level, hint = CODE_META.get(code, (code, '', ''))
        files = sum(1 for r in results if any(d['code'] == code for d in r['diffs']))
        ws3.append([code, name, level, cnt, files, hint])
    for col, w in zip('ABCDEF', [26, 34, 8, 8, 12, 50]):
        ws3.column_dimensions[col].width = w

    ws4 = wb.create_sheet('错误码x文件')
    code_list = [c for c, _ in s['codes'].most_common()]
    ws4.append(['文件ID'] + code_list)
    for r in results:
        cnt = Counter(d['code'] for d in r['diffs'] if d['level'] != 'INFO')
        ws4.append([r['file_id']] + [cnt.get(c, 0) for c in code_list])

    if vstats:
        ws5 = wb.create_sheet('交锋统计')
        ow, tw = vstats.get('他们错', 0), vstats.get('我们错', 0)
        decided = ow + tw + vstats.get('双方都错', 0)
        ws5.append(['项目', '数值', '口径'])
        ws5.append(['他们错', ow, '裁决为 ours（我们的值对）'])
        ws5.append(['我们错', tw, '裁决为 theirs（他们的值对）'])
        ws5.append(['双方都错', vstats.get('双方都错', 0), '原件里有第三个正确值'])
        ws5.append(['未裁决', max(0, sum(r['errors'] for r in results) - decided), '差异未逐条裁决，不计入胜负'])
        ws5.append(['我们错误占比', f'{tw / decided * 100:.1f}%' if decided else 'N/A', '按分歧处数，不按文件数'])
        for col, w in zip('ABC', [18, 14, 40]):
            ws5.column_dimensions[col].width = w

    wb.save(path)
    return path


# ---------------------------------------------------------------- 命令
def read_cases(path):
    base = os.path.dirname(os.path.abspath(path))
    out = []
    with open(path, encoding='utf-8-sig', newline='') as f:
        for row in csv.DictReader(f):
            if not row.get('file_id'):
                continue
            fix = lambda p: p if (not p or os.path.isabs(p)) else os.path.join(base, p)
            out.append((row['file_id'].strip(), fix(row.get('ours', '')), fix(row.get('theirs', ''))))
    return out


def read_verdicts(path):
    """verdicts.csv: file_id,field,ours,theirs,truth → (每文件裁决标签, 全局统计)"""
    if not path or not os.path.exists(path):
        return {}, {}
    per_file = {}
    st = Counter()
    with open(path, encoding='utf-8-sig', newline='') as f:
        for row in csv.DictReader(f):
            truth = (row.get('truth') or '').strip()
            if truth not in ('ours', 'theirs', 'both_wrong', 'equiv'):
                continue
            label = {'ours': '他们错', 'theirs': '我们错', 'both_wrong': '双方都错', 'equiv': '等价'}[truth]
            per_file.setdefault(row['file_id'], Counter())[label] += 1
            st[label] += 1
    labels = {fid: ' '.join(f'{k}{v}' for k, v in c.items()) for fid, c in per_file.items()}
    return labels, dict(st)


def print_verdict_stats(results, vstats, verdicts=None):
    """裁决后才能真正回答“我们比线上系统强多少”——口径按错误处数，不按文件数平均"""
    if not vstats:
        print('\n（未提供 verdicts.csv：以下只报差异，不判谁错。差异 ≠ 我们错，需对照原件裁决）')
        return
    ow, tw = vstats.get('他们错', 0), vstats.get('我们错', 0)
    decided = ow + tw + vstats.get('双方都错', 0)
    undecided = sum(r['errors'] for r in results) - decided
    print(f"\n裁决统计：他们错 {ow} 处 | 我们错 {tw} 处 | 双方都错 {vstats.get('双方都错', 0)} 处 | "
          f"未裁决 {max(0, undecided)} 处")
    if decided:
        print(f"交锋：分歧 {decided} 处，我们胜 {ow} 处（{ow / decided * 100:.1f}%），"
              f"他们胜 {tw} 处（{tw / decided * 100:.1f}%）")
    bad = sorted(fid for fid, lab in (verdicts or {}).items() if '我们错' in str(lab))
    if bad:
        print(f"我们出错的文件：{bad}")


def cmd_compare(args):
    r = eval_file('single', args.ours, args.theirs, args.ours_sheet, args.theirs_sheet)
    verdicts, vstats = read_verdicts(args.verdicts)
    results = [r]
    print_report(results, verdicts, vstats)
    print(f"\n报告: {write_report(results, args.out, verdicts, vstats)}")
    return 0 if r['status'] == 'PASS' else 1


def cmd_batch(args):
    cases = read_cases(args.cases)
    if not cases:
        print(f'[Error] {args.cases} 里没有有效用例')
        return 2
    verdicts, vstats = read_verdicts(args.verdicts)
    results = [eval_file(fid, o, t, args.ours_sheet, args.theirs_sheet) for fid, o, t in cases]
    print_report(results, verdicts, vstats)
    path = write_report(results, args.out, verdicts, vstats)
    print(f"\n报告: {path}")
    bad = sum(1 for r in results if r['status'] != 'PASS')
    print(f"验收结论：{'PASS' if bad == 0 else f'FAIL（{bad}/{len(results)} 份文件不达标）'}")
    return 0 if bad == 0 else 1


def cmd_archive(args):
    """把一批已裁决过的用例存成回归集：fixtures/<id>/{ours.xlsx,theirs.xlsx,case.json}"""
    cases = read_cases(args.cases)
    os.makedirs(args.dest, exist_ok=True)
    n = 0
    for fid, ours, theirs in cases:
        d = os.path.join(args.dest, re.sub(r'[^\w\-.]', '_', fid))
        os.makedirs(d, exist_ok=True)
        for src, name in ((ours, 'ours.xlsx'), (theirs, 'theirs.xlsx')):
            if src and os.path.exists(src):
                with open(src, 'rb') as fi, open(os.path.join(d, name), 'wb') as fo:
                    fo.write(fi.read())
        with open(os.path.join(d, 'case.json'), 'w', encoding='utf-8') as f:
            json.dump({'file_id': fid, 'archived_at': time.strftime('%Y-%m-%d %H:%M:%S')},
                      f, ensure_ascii=False)
        n += 1
    print(f'已归档 {n} 个回归用例到 {args.dest}（新增文件的裁决结果请写进 verdicts.csv 后再回归）')
    return 0


def cmd_regression(args):
    """全量回归：任一用例比上次多出错 = 回归失败。--update 用当前结果刷新基线。"""
    root = args.dest
    dirs = sorted(d for d in os.listdir(root) if os.path.isdir(os.path.join(root, d))) if os.path.isdir(root) else []
    if not dirs:
        print(f'[Info] {root} 还没有回归用例，先跑 archive')
        return 0
    regress, worse = [], 0
    for d in dirs:
        p = os.path.join(root, d)
        ours, theirs = os.path.join(p, 'ours.xlsx'), os.path.join(p, 'theirs.xlsx')
        if not (os.path.exists(ours) and os.path.exists(theirs)):
            continue
        r = eval_file(d, ours, theirs)
        base_path = os.path.join(p, 'baseline.json')
        base = None
        if os.path.exists(base_path):
            with open(base_path, encoding='utf-8') as f:
                base = json.load(f)
        cur = {'status': r['status'], 'errors': r['errors'],
               'codes': sorted({x['code'] for x in r['diffs'] if x['level'] != 'INFO'})}
        flag = ''
        if base:
            if cur['errors'] > base['errors'] or set(cur['codes']) - set(base['codes']):
                flag = f"◀ REGRESSION 比基线差 ({base['errors']}→{cur['errors']} 处, 新增码 {sorted(set(cur['codes']) - set(base['codes']))})"
                worse += 1
            elif cur['errors'] < base['errors']:
                flag = f"◀ 改善 ({base['errors']}→{cur['errors']} 处)"
        if args.update or not base:
            with open(base_path, 'w', encoding='utf-8') as f:
                json.dump(cur, f, ensure_ascii=False, indent=2)
        print(f"  {d[:34]:36s}{r['status']:6s}{cur['errors']:4d} 处  {cur['codes']}  {flag}")
        if flag.startswith('◀ REGRESSION'):
            regress.append(d)
    print(f"\n回归用例 {len(dirs)} 个 | 回归失败 {len(regress)} 个")
    return 1 if regress else 0


# ---------------------------------------------------------------- 自检
def cmd_scan(args):
    """每天新文件进来时用它生成 cases.csv：按文件名主干配对 ours/ 与 theirs/。
    只有一边有文件的也算缺失（缺一份 = 没法评 = 当作未完成），不静默跳过。"""
    def index(d):
        out = {}
        for f in os.listdir(d):
            if f.startswith(('~$', '.')) or not f.lower().endswith(('.xlsx', '.xlsm')):
                continue
            out[os.path.splitext(f)[0]] = os.path.join(d, f)
        return out

    ours, theirs = index(args.ours_dir), index(args.theirs_dir)
    keys = sorted(set(ours) | set(theirs))
    rows, lonely = [], []
    for k in keys:
        if not (k in ours and k in theirs):
            lonely.append(f"{k}（缺 {'他们' if k not in theirs else '我们'}）")
            continue
        rows.append((k, ours[k], theirs[k]))
    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or '.', exist_ok=True)
    with open(args.out, 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f)
        w.writerow(['file_id', 'ours', 'theirs'])
        for r in rows:
            w.writerow([os.path.basename(r[1]), os.path.abspath(r[1]), os.path.abspath(r[2])])
    print(f'配对成功 {len(rows)} 份 → {args.out}')
    if lonely:
        print(f'❗ {len(lonely)} 份只有一边，无法评差（必须补上后才能算完成）：')
        for x in lonely:
            print(f'    - {x}')
    return 0 if rows else 2


def selftest():
    """不依赖外部件：证明 ①自己vs自己=0差异 ②每一类错误都被精确命中 ③严口径不被平均值稀释"""
    base = [
        {'名称规格': '变压器', '规格参数': '容量kVA:30\n产品/型号:S13型（NX3）', '市场价': 10500,
         '工程价': 9450, '工程价折扣比例': 0.9, '单位': '台', '品牌': '望变', '类别编码': '310502',
         '报价地区省': '全国', '是否含运费': '包含', '增值税税率': 13, '厂家报价文件页码': '第1页'},
        {'名称规格': '变压器', '规格参数': '容量kVA:50\n产品/型号:S13型（NX3）', '市场价': 13500,
         '工程价': 12150, '工程价折扣比例': 0.9, '单位': '台', '品牌': '望变', '类别编码': '310502',
         '报价地区省': '全国', '是否含运费': '包含', '增值税税率': 13, '厂家报价文件页码': '第1页'},
        {'名称规格': '变压器', '规格参数': '容量kVA:80\n产品/型号:S20型（NX2）', '市场价': 18000,
         '工程价': 16200, '工程价折扣比例': 0.9, '单位': '台', '品牌': '望变', '类别编码': '310502',
         '报价地区省': '全国', '是否含运费': '包含', '增值税税率': 13, '厂家报价文件页码': '第1页'},
    ]

    def diff(a, b):
        pairs, only_a, only_b = align(a, b)
        ds = []
        for i, j in pairs:
            ds += diff_pair(a[i], b[j], _row_label(a[i], i))
        ds += [dict(code='E01_ROW_OURS_MISSING')] * len(only_a)
        ds += [dict(code='E02_ROW_OURS_EXTRA')] * len(only_b)
        return [d['code'] for d in ds if d.get('level') != 'INFO'], ds

    def clone(mutate=None, drop=None, add=None):
        out = [dict(r) for r in base]
        if mutate:
            mutate(out)
        if drop is not None:
            out.pop(drop)
        if add:
            out.append(add)
        return out

    # ① 完全一致
    codes, ds = diff(base, clone())
    assert codes == [], f'自己vs自己应当 0 差异，实际 {codes}'

    # ② 各类错误逐项命中
    cases = {
        'E03_PRICE': clone(lambda r: r[0].update(市场价=105000)),                      # 价格多一位
        'E04_UNIT': clone(lambda r: r[0].update(单位='个')),                           # 单位不同
        'E05_NAME_TYPO': clone(lambda r: r[0].update(名称规格='变压哭')),               # 错一个字
        'E06_NAME_DIFF': clone(lambda r: r[0].update(名称规格='干式电力变压器成套设备')),  # 语义不同
        'E08_SPEC_VALUE': clone(lambda r: r[0].update(规格参数='容量kVA:30\n产品/型号:S15型（NX5）')),
        'E09_CODE': clone(lambda r: r[0].update(类别编码='999999')),
        'E10_BRAND': clone(lambda r: r[0].update(品牌='望变电气')),
        'E11_TERMS': clone(lambda r: r[0].update(是否含运费='不含')),
        'E12_REGION': clone(lambda r: r[0].update(报价地区省='重庆市')),
        'E15_PAGE': clone(lambda r: r[0].update(厂家报价文件页码='第3页')),
        'E01_ROW_OURS_MISSING': clone(drop=1),                                        # 我们少一条
        'E02_ROW_OURS_EXTRA': clone(add=dict(base[0], 规格参数='容量kVA:100\n产品/型号:S22型', 市场价=26000)),
        'E07_SPEC_MISSING': clone(lambda r: r[0].update(规格参数='容量kVA:30')),        # 少一个参数
    }
    for want, theirs in cases.items():
        codes, _ = diff(base, theirs)
        assert want in codes, f'{want} 未被命中，实际命中 {codes}'

    # ③ 归一化等价不能误判成错误（这条防线防止"永远 FAIL"）
    ok = clone(lambda r: r[0].update(市场价='¥10,500.00', 单位='台 ', 规格参数='容量kVA：30\n产品/型号:S13型(NX3)'))
    codes, _ = diff(base, ok)
    assert codes == [], f'等价写法不应判错，实际 {codes}'

    # ④ 整表结构错位必须聚合识别成 E14，而不是刷 30 条行差异
    shifted = [dict(r, 规格参数=f'容量kVA:{i}', 市场价=1000 + i) for i, r in enumerate(base * 3)]
    pairs, only_a, only_b = align(base * 3, shifted)
    assert len(only_a) >= 3, '结构错位应当留下大量未匹配行'
    print(f'  结构错位: 未匹配 我们{len(only_a)} 他们{len(only_b)} → 会触发 E14_STRUCTURE 折叠')

    print('selftest: 全部通过（0 差异基线 / 13 类错误全命中 / 等价写法不误报 / 结构错位可识别）')
    return 0


def main():
    ap = argparse.ArgumentParser(description='QuoteStandardizer 评差引擎（文件级严口径）')
    sub = ap.add_subparsers(dest='cmd', required=True)

    def common(p):
        p.add_argument('--ours-sheet', default=None)
        p.add_argument('--theirs-sheet', default=None)
        p.add_argument('--verdicts', default=None, help='人工裁决 verdicts.csv')
        p.add_argument('-o', '--out', default='eval_out')

    p = sub.add_parser('compare'); common(p)
    p.add_argument('--ours', required=True); p.add_argument('--theirs', required=True)
    p.set_defaults(func=cmd_compare)

    p = sub.add_parser('batch'); common(p)
    p.add_argument('--cases', required=True)
    p.set_defaults(func=cmd_batch)

    p = sub.add_parser('scan')
    p.add_argument('--ours-dir', required=True, help='我们输出的目录')
    p.add_argument('--theirs-dir', required=True, help='线上系统输出的目录')
    p.add_argument('-o', '--out', default='cases.csv')
    p.set_defaults(func=cmd_scan)

    p = sub.add_parser('archive')
    p.add_argument('--cases', required=True)
    p.add_argument('--dest', default='fixtures')
    p.set_defaults(func=cmd_archive)

    p = sub.add_parser('regression')
    p.add_argument('--dest', default='fixtures')
    p.add_argument('--update', action='store_true')
    p.set_defaults(func=cmd_regression)

    p = sub.add_parser('selftest')
    p.set_defaults(func=lambda a: selftest())

    args = ap.parse_args()
    sys.exit(args.func(args))


if __name__ == '__main__':
    main()
