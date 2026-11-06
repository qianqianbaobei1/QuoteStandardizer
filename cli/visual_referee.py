#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cli/visual_referee.py —— 三方原图裁决评测（线上 vs 标准 vs 原图）

解决的核心问题：线上值与标准值不一致时，谁对？
人工裁决太慢，且旧脚本里的"原图证据"往往是写死的字符串。这里在出现分歧时
自动把**该行所在页的原图** + 线上值 + 标准值一起交给视觉模型，由它看原图判案。

裁决枚举（唯一真值是原图）：
  ONLINE_WIN   原图支持线上值 → 线上不算错，标准写进「标准待修正」
  STD_WIN      原图支持标准值 → 线上判错
  BOTH_WRONG   两边都错       → 线上判错，原图真值写进「标准待修正」
  PENDING      原图确实没体现该字段（外部商务配置） → 不计胜负，待人工
  UNADJUDICATED 没挂到原图 / 视觉模型不可用 → **如实记为未裁决**，绝不默认判线上错

用法：
  python3 cli/visual_referee.py --cases cases.csv --source-dir 原件根目录 [-o out/]
  cases.csv 列：file_id,ours,theirs,source(可选，覆盖 --source-dir)
     ours   = 我们的解析（标准答案/基准）
     theirs = 线上系统导出结果
  ——与 evaluate.py 的 cases.csv 同语义：ours 是我们，theirs 是线上。
"""
import argparse
import base64
import csv
import json
import os
import re
import sys
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed

import fitz
import openpyxl
import requests
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import VISION_API_KEY, VISION_BASE_URL, VISION_MODEL  # noqa: E402

REFEREE_SYSTEM_PROMPT = """【报价单质量 · 三方原图裁决】
你是严格的报价单解析质量审查法官。唯一真值是你看到的原始报价单图像，标准答案也可能错。

裁决枚举（verdict）：
- "ONLINE_WIN"：原图明确支持「线上解析值」（标准答案存在 OCR 错字/漏提/算错/无依据编造）
- "STD_WIN"：原图明确支持「标准答案值」（线上解析漏提/错提/幻觉/分类错误）
- "BOTH_WRONG"：线上值与标准值都与原图真实内容不符
- "PENDING"：原图中确实看不到该字段（如外部商务配置、系统侧折扣），无法判定

只输出 JSON，不要任何解释文字：
{"verdict": "ONLINE_WIN|STD_WIN|BOTH_WRONG|PENDING",
 "ground_truth_from_image": "你从原图亲眼读到的该字段真实值（看不到就写空串）",
 "evidence_location": "证据在原图的位置，如：第2页表头/第5行单价列/页脚备注第2条",
 "std_fix_reason": "若标准答案有误，一句话说明错在哪；否则空串"}"""

VERDICTS = ('ONLINE_WIN', 'STD_WIN', 'BOTH_WRONG', 'PENDING', 'UNADJUDICATED')
_WIN_LABEL = {
    'ONLINE_WIN': '线上对/标准错',
    'STD_WIN': '标准对/线上错',
    'BOTH_WRONG': '双方都错',
    'PENDING': '待人工(原图未体现)',
    'UNADJUDICATED': '未裁决',
}

# 手册第二章 C 类（外部商务配置）：原图未体现时一律 PENDING，不得判任何一方错
C_CLASS_FIELDS = {'增值税税率', '工程价折扣比例', '折扣系列', '折扣下限', '折扣上限',
                  '价格填报类型', '报价地区省', '报价地区市'}

# 手册第三章：evaluate.py 的字段错误码 → 四层检验图谱编码
E_TO_L = {
    'E03_PRICE': 'L1-PRICE', 'E04_UNIT': 'L1-UNIT',
    'E05_NAME_TYPO': 'L1-NAME', 'E06_NAME_DIFF': 'L1-NAME',
    'E10_BRAND': 'L1-BRAND', 'E11_TERMS': 'L1-TERMS',
    'E09_CODE': 'L2-CODE',
    'E07_SPEC_MISSING': 'L3-ATTR-MISS', 'E08_SPEC_VALUE': 'L3-VALUE',
    'E15_PAGE': 'L4-SRC-PAGE', 'E12_REGION': 'L1-REGION',
}

def layer_of(field):
    """字段 → 四层中的哪一层（L1 图片解析 / L2 分类 / L3 规格 / L4 凭证）"""
    if field.startswith('规格参数'):
        return 'L3'
    if field == '类别编码':
        return 'L2'
    if field in ('厂家报价文件页码',):
        return 'L4'
    return 'L1'


def _load_evaluate():
    """复用 evaluate.py 的归一化/对齐/逐字段比对，绝不另起一套口径"""
    import importlib.util
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'evaluate.py')
    spec = importlib.util.spec_from_file_location('evaluate_engine', path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# ------------------------------------------------------------------ 原图
def find_source_file(source, hint=''):
    """source 可以是单份 PDF/图片，也可以是一个目录；hint 为文件名线索"""
    if not source:
        return None
    if os.path.isfile(source):
        return source
    if not os.path.isdir(source):
        return None
    files = []
    for root, _, fns in os.walk(source):
        for fn in fns:
            if fn.lower().endswith(('.pdf', '.png', '.jpg', '.jpeg')):
                files.append(os.path.join(root, fn))
    if not files:
        return None
    if hint:
        key = os.path.basename(str(hint)).lower()
        hits = [f for f in files
                if key and (key in os.path.basename(f).lower() or os.path.basename(f).lower() in key)]
        if hits:
            return sorted(hits, key=len)[0]
    pdfs = [f for f in files if f.lower().endswith('.pdf')]
    return sorted(pdfs or files)[0]


def page_number_of(row):
    m = re.search(r'\d+', str((row or {}).get('厂家报价文件页码') or ''))
    return max(1, int(m.group(0))) if m else 1


def page_image_b64(path, page):
    if not path or not os.path.exists(path):
        return ''
    try:
        if path.lower().endswith('.pdf'):
            with fitz.open(path) as doc:
                idx = max(0, min(int(page or 1) - 1, len(doc) - 1))
                return base64.b64encode(doc[idx].get_pixmap(dpi=170).tobytes('png')).decode()
        with open(path, 'rb') as f:
            return base64.b64encode(f.read()).decode()
    except Exception as e:
        print(f"  [原图] 渲染失败 {path} p{page}: {e}")
        return ''


# ------------------------------------------------------------------ 裁决
def adjudicate(task, cache=None):
    """单处分歧的原图裁决（结果按 图+字段+两值 缓存，同构行不重复花钱）"""
    if task.get('pre_verdict'):
        # 手册公理 2：C 类外部配置字段原图未体现，直接 PENDING，不浪费一次视觉调用，也不冤枉任一方
        return {'verdict': task['pre_verdict'], 'ground_truth_from_image': '',
                'evidence_location': 'C 类外部配置字段，原图未体现', 'std_fix_reason': ''}
    sig = json.dumps([task.get('source_file'), task['page'], task['field'],
                      str(task['online']), str(task['standard'])], ensure_ascii=False)
    if cache is not None and sig in cache:
        return cache[sig]

    undecided = {'verdict': 'UNADJUDICATED', 'ground_truth_from_image': '',
                 'evidence_location': '', 'std_fix_reason': ''}
    if not VISION_API_KEY or not VISION_BASE_URL or not VISION_MODEL:
        return dict(undecided, std_fix_reason='未配置视觉模型(VISION_*)')
    img = page_image_b64(task.get('source_file'), task['page'])
    if not img:
        return dict(undecided, std_fix_reason='没找到该行所在页的原图')

    prompt = (
        f"【待裁决分歧】\n"
        f"- 原图页码：第 {task['page']} 页\n"
        f"- 物料：品名={task.get('name', '')} | 规格={task.get('raw_spec', '')}\n"
        f"- 争议字段：【{task['field']}】（错误码 {task.get('code', '-')}）\n"
        f"- 标准答案值：\"{task['standard']}\"\n"
        f"- 线上解析值：\"{task['online']}\"\n\n"
        f"请只看原图，裁决这两个值谁与原图一致。"
    )
    payload = {
        'model': VISION_MODEL, 'temperature': 0,
        'max_tokens': 400,
        'messages': [
            {'role': 'system', 'content': REFEREE_SYSTEM_PROMPT},
            {'role': 'user', 'content': [
                {'type': 'text', 'text': prompt},
                {'type': 'image_url', 'image_url': {'url': f"data:image/png;base64,{img}"}},
            ]},
        ],
    }
    try:
        resp = requests.post(f"{VISION_BASE_URL.rstrip('/')}/chat/completions",
                             headers={'Authorization': f"Bearer {VISION_API_KEY}"},
                             json=payload, timeout=90)
        if resp.status_code != 200:
            return dict(undecided, std_fix_reason=f'裁判接口 HTTP {resp.status_code}')
        content = resp.json()['choices'][0]['message']['content']
        m = re.search(r'\{.*\}', content, re.S)
        parsed = json.loads(m.group(0)) if m else {}
        verdict = str(parsed.get('verdict') or '').upper()
        if verdict not in VERDICTS:
            return dict(undecided, std_fix_reason=f'裁判输出无法解析: {content[:80]}')
        out = {
            'verdict': verdict,
            'ground_truth_from_image': str(parsed.get('ground_truth_from_image') or '').strip(),
            'evidence_location': str(parsed.get('evidence_location') or '').strip(),
            'std_fix_reason': str(parsed.get('std_fix_reason') or '').strip(),
        }
    except Exception as e:
        out = dict(undecided, std_fix_reason=f'裁判调用异常: {e}')
    if cache is not None:
        cache[sig] = out
    return out


def adjudicate_batch(tasks, workers=6):
    if not tasks:
        return []
    cache = {}
    results = [None] * len(tasks)
    with ThreadPoolExecutor(max_workers=max(1, min(workers, len(tasks)))) as pool:
        futs = {pool.submit(adjudicate, t, cache): i for i, t in enumerate(tasks)}
        for fut in as_completed(futs):
            results[futs[fut]] = fut.result()
    return results


# ------------------------------------------------------------------ 组装
def collect_discrepancies(ev, ours_path, theirs_path, source, file_hint='', max_per_file=200):
    """对齐 + 逐字段比对，只留下需要看原图判定的分歧（归一后等价的不算）"""
    A, pa = ev.load_table(ours_path)
    B, pb = ev.load_table(theirs_path)
    if pa or pb:
        print(f"  ! 表结构告警: {pa + pb}")
    pairs, only_a, only_b = ev.align(A, B)
    src_file = find_source_file(source, file_hint)
    if not src_file:
        print(f"  ! 没找到原件（source={source or '(空)'} hint={file_hint or '(空)'}）——分歧将如实记为未裁决")
    tasks, skipped = [], []
    for i, j in pairs:
        row = A[i]
        for d in ev.diff_pair(A[i], B[j], ev._row_label(A[i], i)):
            if d['level'] == 'INFO' or d['code'] in ('E00_EQUIV',) or d['field'] == '(整行)':
                continue
            if len(tasks) >= max_per_file:
                skipped.append(d['label'])
                continue
            # 注意方向：evaluate.py 里 A/d['ours'] = 我们的解析（基准），B/d['theirs'] = 线上系统
            tasks.append({
                'file_id': file_hint,
                'source_file': src_file,
                'page': page_number_of(row),
                'row_label': d['label'],
                'name': str(row.get('名称规格') or ''),
                'raw_spec': str(row.get('规格参数') or ''),
                'field': d['field'],
                'code': d['code'],
                'l_code': E_TO_L.get(d['code'], 'L1-OTHER'),
                'layer': layer_of(d['field']),
                'pre_verdict': ('PENDING'
                                if d['field'] in C_CLASS_FIELDS
                                and (d['ours'] in (None, '') or d['theirs'] in (None, ''))
                                else ''),
                'online': d['theirs'],
                'standard': d['ours'],
            })
    return dict(rows_ours=len(A), rows_theirs=len(B), diffs=tasks,
                all_rows=[ev._row_label(A[i], i) for i in range(len(A))],
                only_a=len(only_a), only_b=len(only_b), skipped=skipped)


def read_cases(path):
    base = os.path.dirname(os.path.abspath(path))
    out = []
    with open(path, encoding='utf-8-sig', newline='') as f:
        for row in csv.DictReader(f):
            if not row.get('file_id'):
                continue
            fix = lambda p: p if (not p or os.path.isabs(p)) else os.path.join(base, p)
            out.append({'file_id': row['file_id'].strip(), 'ours': fix(row.get('ours', '')),
                        'theirs': fix(row.get('theirs', '')), 'source': fix(row.get('source', ''))})
    return out


# ------------------------------------------------------------------ 报告
def _cell(v, n=120):
    """Excel 单元格显示：None/空不写成字符串 'None'"""
    return '' if v in (None, 'None', 'nan') else str(v)[:n]


def _sheet(wb, title, header, rows, widths=None, freeze=True):
    ws = wb.create_sheet(title)
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', fgColor='EEF2F7')
    for r in rows:
        ws.append(r)
    for idx, w in enumerate(widths or [], 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    if freeze:
        ws.freeze_panes = 'A2'
    return ws


def _row_status_map(diffs, all_rows=None):
    """手册第五章：逐行计算 L1/L2/L3/L4 状态与行级综合状态（无分歧的行也是 PASS，必须计入）"""
    rows = defaultdict(lambda: {'L1': 'PASS', 'L2': 'PASS', 'L3': 'PASS', 'L4': 'PASS', 'fields': []})
    for label in (all_rows or []):
        rows[label]  # 先把所有行建出来，避免"只统计有分歧的行"把通过率算成 0
    for d in diffs:
        st = rows[d['row_label']]
        layer = d.get('layer') or layer_of(d['field'])
        if d['verdict'] in ('STD_WIN', 'BOTH_WRONG'):
            st[layer] = 'FAIL'
        elif d['verdict'] in ('PENDING', 'UNADJUDICATED') and st[layer] == 'PASS':
            st[layer] = 'PENDING'
        st['fields'].append(f"{d['field']}={_WIN_LABEL[d['verdict']]}")
    for st in rows.values():
        vals = [st[k] for k in ('L1', 'L2', 'L3', 'L4')]
        st['状态'] = 'FAIL' if 'FAIL' in vals else ('PENDING' if 'PENDING' in vals else 'PASS')
    return rows


def write_reports(per_file, out_dir, source_label):
    """5 个 Sheet（对齐手册第七章）+ 一块看板 HTML"""
    os.makedirs(out_dir, exist_ok=True)
    all_diffs = [d for f in per_file for d in f['diffs']]
    counts = Counter(d['verdict'] for d in all_diffs)
    decided = counts['ONLINE_WIN'] + counts['STD_WIN'] + counts['BOTH_WRONG']
    online_right = counts['ONLINE_WIN']

    # 分母口径必须说清楚：这里是基准库行数，不是原图物理行数（原图行数需要版面探测，本脚本不假装有）
    base_rows = sum(f['rows_ours'] for f in per_file)
    row_pass = row_pending = 0
    for f in per_file:
        for st in _row_status_map(f['diffs'], f.get('all_rows')).values():
            row_pass += st['状态'] == 'PASS'
            row_pending += st['状态'] == 'PENDING'

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '1_汇总'
    ws.append([f"线上正确率（已裁决分歧处）: "
               f"{(online_right / decided * 100) if decided else 0:.1f}%"])
    ws['A1'].font = Font(size=20, bold=True)
    ws.append([])
    ws.append(['口径', '数值', '说明'])
    ws.append(['文件数', len(per_file), '每份文件按标准答案 xlsx 为一个用例'])
    ws.append(['基准行数', base_rows, '基准库(标准答案)行数——不是原图物理行数，不含原图行数校核'])
    ws.append(['行级 PASS 数', row_pass, f'该行所有分歧裁决后无 FAIL；PENDING 行 {row_pending} 行不计 PASS'])
    ws.append(['行级 PASS 率', f"{row_pass / base_rows * 100:.1f}%" if base_rows else 'N/A', '行 PASS / 基准行数'])
    ws.append(['分歧处数', len(all_diffs), '归一化后仍不一致的字段处数'])
    ws.append(['已裁决', decided, '看过原图并有明确结论'])
    ws.append(['未裁决+待人工', counts['UNADJUDICATED'] + counts['PENDING'],
               '没挂到原图 / C 类配置字段 / 原图未体现，不计胜负'])
    for v in VERDICTS:
        ws.append([_WIN_LABEL[v], counts[v], f'{v}'])
    ws.append([])
    ws.append(['L1~L4 层裁决正确率（仅统计该层已裁决分歧）'])
    ws.append(['层', '线上对', '标准对', '双方都错', '层正确率'])
    for layer in ('L1', 'L2', 'L3', 'L4'):
        sub = [d for d in all_diffs if (d.get('layer') or layer_of(d['field'])) == layer]
        c = Counter(d['verdict'] for d in sub)
        dec = c['ONLINE_WIN'] + c['STD_WIN'] + c['BOTH_WRONG']
        ws.append([layer, c['ONLINE_WIN'], c['STD_WIN'], c['BOTH_WRONG'],
                   f"{c['ONLINE_WIN'] / dec * 100:.1f}%" if dec else '无已裁决分歧'])
    for col, w in zip('ABCDE', [34, 12, 46, 10, 16]):
        ws.column_dimensions[col].width = w

    row_sheet = []
    for f in per_file:
        for label, st in _row_status_map(f['diffs'], f.get('all_rows')).items():
            row_sheet.append([f['file_id'], label, st['L1'], st['L2'], st['L3'], st['L4'],
                              st['状态'], '; '.join(st['fields'])])
    _sheet(wb, '2_逐行判定',
           ['文件ID', '行定位', 'L1 图片解析', 'L2 分类', 'L3 规格', 'L4 凭证', '行状态', '问题明细'],
           row_sheet,
           [26, 34, 12, 10, 10, 10, 10, 60])

    _sheet(wb, '3_差异明细',
           ['文件ID', '行定位', '字段', '错误码', '线上值', '标准值', '裁决', '原图真值', '证据位置',
            '标准修正原因', '四层编码', '页码'],
           [[f['file_id'], d['row_label'], d['field'], d['code'], _cell(d['online']),
             _cell(d['standard']), f"{_WIN_LABEL[d['verdict']]}({d['verdict']})",
             d['ground_truth_from_image'], d['evidence_location'], d['std_fix_reason'],
             d.get('l_code') or E_TO_L.get(d['code'], ''), d.get('page', '')]
            for f in per_file for d in f['diffs']],
           [26, 32, 20, 22, 30, 30, 22, 24, 30, 36, 18, 8])

    # 双重归因：受影响行数（影响面） vs 根因去重数（同文件同页同层只算 1 个根因）
    root = defaultdict(lambda: {'rows': set(), 'pages': set(), 'count': 0})
    for f in per_file:
        for d in f['diffs']:
            key = (d.get('l_code') or E_TO_L.get(d['code'], ''), d['verdict'])
            root[key]['rows'].add((f['file_id'], d['row_label']))
            root[key]['pages'].add((f['file_id'], d.get('page') or 1))
            root[key]['count'] += 1
    _sheet(wb, '4_问题归因',
           ['四层编码', '裁决', '影响处数', '受影响行数', '根因去重数', '建议'],
           [[k[0], _WIN_LABEL[k[1]], v['count'], len(v['rows']), len(v['pages']),
             '线上侧修' if k[1] == 'STD_WIN' else ('标准侧修' if k[1] == 'ONLINE_WIN' else '待人工确认')]
            for k, v in sorted(root.items(), key=lambda kv: -kv[1]['count'])],
           [22, 22, 10, 12, 12, 22])

    fixes = [(f['file_id'], d) for f in per_file for d in f['diffs']
             if d['verdict'] in ('ONLINE_WIN', 'BOTH_WRONG')]
    _sheet(wb, '5_标准待修正',
           ['文件ID', '行定位', '字段', '原标准值', '原图真值', '裁决', '修正原因', '证据位置'],
           [[fid, d['row_label'], d['field'], _cell(d['standard']),
             d['ground_truth_from_image'] or '(原图无法确认)', d['verdict'],
             d['std_fix_reason'], d['evidence_location']] for fid, d in fixes],
           [26, 32, 20, 26, 26, 20, 40, 30])

    xlsx_path = os.path.join(out_dir, '三方裁决报告.xlsx')
    wb.save(xlsx_path)

    html_path = os.path.join(out_dir, '报价单解析质量看板.html')
    rate = (online_right / decided * 100) if decided else 0.0
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(f"""<!doctype html><html lang="zh"><head><meta charset="utf-8">
<title>报价单解析质量看板</title><style>
body{{margin:0;background:#f6f7f9;font:16px/1.6 -apple-system,"PingFang SC",sans-serif;color:#1b1f24}}
.wrap{{max-width:1180px;margin:0 auto;padding:48px 125px 48px 125px}}
h1{{font-size:1.1rem;letter-spacing:.14em;color:#6b7280;font-weight:600;margin:0 0 8px}}
.rate{{font-size:4.2rem;font-weight:800;line-height:1.05;color:{'#0f9d58' if rate >= 80 else '#d93025'}}}
.sub{{color:#6b7280;margin:6px 0 32px}}
.cards{{display:grid;grid-template-columns:repeat(4,1fr);gap:16px}}
.card{{background:#fff;border:1px solid #e5e7eb;border-radius:14px;padding:20px}}
.card .n{{font-size:1.9rem;font-weight:700}}.card .t{{color:#6b7280;font-size:.9rem}}
table{{width:100%;border-collapse:collapse;margin-top:28px;background:#fff;border-radius:14px;overflow:hidden}}
th,td{{padding:12px 14px;text-align:left;border-bottom:1px solid #eef0f3;font-size:.94rem}}
th{{background:#eef2f7;color:#475569}}
</style></head><body><div class="wrap">
<h1>线上 vs 标准 vs 原图 · 已裁决分歧处的线上正确率</h1>
<div class="rate">{rate:.1f}%</div>
<div class="sub">分歧 {len(all_diffs)} 处 · 已裁决 {decided} 处 · 未裁决 {counts['UNADJUDICATED'] + counts['PENDING']} 处 · 原件来源 {source_label}</div>
<div class="cards">
<div class="card"><div class="n">{len(per_file)}</div><div class="t">文件数</div></div>
<div class="card"><div class="n">{counts['ONLINE_WIN']}</div><div class="t">线上对/标准错</div></div>
<div class="card"><div class="n">{counts['STD_WIN']}</div><div class="t">标准对/线上错</div></div>
<div class="card"><div class="n">{counts['BOTH_WRONG']}</div><div class="t">双方都错</div></div>
</div>
<table><tr><th>文件</th><th>分歧</th><th>线上对</th><th>标准对</th><th>双方都错</th><th>未裁决</th></tr>
{''.join(f"<tr><td>{f['file_id']}</td><td>{len(f['diffs'])}</td>"
         f"<td>{sum(1 for d in f['diffs'] if d['verdict'] == 'ONLINE_WIN')}</td>"
         f"<td>{sum(1 for d in f['diffs'] if d['verdict'] == 'STD_WIN')}</td>"
         f"<td>{sum(1 for d in f['diffs'] if d['verdict'] == 'BOTH_WRONG')}</td>"
         f"<td>{sum(1 for d in f['diffs'] if d['verdict'] in ('UNADJUDICATED', 'PENDING'))}</td></tr>"
         for f in per_file)}
</table></div></body></html>""")
    return xlsx_path, html_path


# ------------------------------------------------------------------ CLI
def main():
    ap = argparse.ArgumentParser(description='线上 vs 标准 vs 原图 三方裁决评测')
    ap.add_argument('--cases', required=True, help='cases.csv: file_id,ours,theirs[,source]')
    ap.add_argument('--source-dir', default='', help='原件根目录（PDF/图片）；cases 里有 source 列则优先用列')
    ap.add_argument('-o', '--out', default='out_visual_referee')
    ap.add_argument('--workers', type=int, default=6, help='并发裁决线路数')
    ap.add_argument('--max-per-file', type=int, default=200, help='单文件最多裁决多少处分歧')
    args = ap.parse_args()

    cases = read_cases(args.cases)
    if not cases:
        print(f'[Error] {args.cases} 没有有效用例')
        return 2
    ev = _load_evaluate()

    per_file = []
    for c in cases:
        print(f"\n[裁决] {c['file_id']}")
        collected = collect_discrepancies(ev, c['ours'], c['theirs'],
                                          c['source'] or args.source_dir,
                                          c['file_id'], args.max_per_file)
        tasks = collected['diffs']
        print(f"  {collected['rows_ours']} vs {collected['rows_theirs']} 行，分歧 {len(tasks)} 处"
              + (f"，超出上限跳过 {len(collected['skipped'])} 处" if collected['skipped'] else ''))
        verdicts = adjudicate_batch(tasks, args.workers)
        diffs = [dict(t, **v) for t, v in zip(tasks, verdicts)]
        c2 = Counter(d['verdict'] for d in diffs)
        print("  " + ' | '.join(f"{_WIN_LABEL[v]} {c2[v]}" for v in VERDICTS if c2[v]))
        per_file.append(dict(file_id=c['file_id'], rows_ours=collected['rows_ours'],
                             all_rows=collected['all_rows'],
                             rows_theirs=collected['rows_theirs'], only_a=collected['only_a'],
                             only_b=collected['only_b'], diffs=diffs))

    xlsx_path, html_path = write_reports(per_file, args.out, args.source_dir or 'cases 内 source 列')
    undecided = sum(1 for f in per_file for d in f['diffs'] if d['verdict'] == 'UNADJUDICATED')
    print(f"\n报告: {xlsx_path}\n看板: {html_path}")
    if undecided:
        print(f"! 有 {undecided} 处分歧没裁决成（缺原图或视觉模型不可用），已如实计入未裁决，未默认判线上错")
    return 0


if __name__ == '__main__':
    sys.exit(main())
