# -*- coding: utf-8 -*-
"""
batch_50_runner.py - 50 批报价单流水线解析与线上解析全量对比验证器
1. 遍历 deliveries/batch_50_eval/pdfs 下的全部 50 个 PDF
2. 调用 main.process_pdf 解析为标准交付 Excel
3. 与 deliveries/batch_50_eval/theirs 下的线上 报价数据.xlsx 进行逐行、逐字段高维对账
4. 产出企业级《50家报价单解析对比与质检总览.xlsx》报告
"""
import os
import sys
import glob
import re
import time
import json
import traceback
import unicodedata
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

from main import process_pdf

PDF_DIR = os.path.join(BASE_DIR, 'deliveries', 'batch_50_eval', 'pdfs')
OURS_DIR = os.path.join(BASE_DIR, 'deliveries', 'batch_50_eval', 'ours')
THEIRS_DIR = os.path.join(BASE_DIR, 'deliveries', 'batch_50_eval', 'theirs')
REPORT_PATH = os.path.join(BASE_DIR, 'deliveries', 'batch_50_eval', '50家报价单解析对比验证总览报告.xlsx')


def norm_str(s):
    s = unicodedata.normalize('NFKC', str(s or '')).lower()
    return re.sub(r'[\s\-_（）()\[\]【】/:：,，、;；.。\'"*×]', '', s)


def norm_money(v):
    if v is None or str(v).strip() == '':
        return None
    m = re.search(r'-?\d+(?:\.\d+)?', str(v).replace(',', ''))
    return float(m.group(0)) if m else None


def norm_unit(s):
    s = unicodedata.normalize('NFKC', str(s or '')).strip().lower()
    return {'㎡': 'm²', 'm2': 'm²', '平方米': 'm²', '平米': 'm²', '㎥': 'm³', 'm3': 'm³',
            '立方米': 'm³', '米': 'm', '公斤': 'kg', 't': 't', '吨': 't', '支': '根', '件': '个'}.get(s, s)


def load_ours_data(xlsx_path):
    if not os.path.exists(xlsx_path):
        return []
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        if '标准化报价' not in wb.sheetnames:
            return []
        ws = wb['标准化报价']
        headers = [c.value for c in ws[1]]
        items = []
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not any(r):
                continue
            d = dict(zip(headers, r))
            name = str(d.get('名称规格') or '').strip()
            if not name:
                continue
            items.append({
                'seq': d.get('顺序号'),
                'name': name,
                'brand': str(d.get('品牌') or '').strip(),
                'unit': norm_unit(d.get('单位')),
                'raw_unit': str(d.get('单位') or '').strip(),
                'price': norm_money(d.get('市场价')),
                'raw_price': d.get('市场价'),
                'eng_price': norm_money(d.get('工程价')),
                'spec': str(d.get('规格参数') or d.get('厂家原始规格参数') or ''),
                'raw_spec': str(d.get('厂家原始规格参数') or ''),
                'cat_code': re.sub(r'\D', '', str(d.get('类别编码') or '')),
                'raw_cat_code': str(d.get('类别编码') or ''),
                'province': d.get('报价地区省'),
                'city': d.get('报价地区市')
            })
        return items
    except Exception as e:
        print(f"Error loading {xlsx_path}: {e}")
        return []


def load_theirs_data(vendor_theirs_dir):
    xlsx_p = os.path.join(vendor_theirs_dir, '报价数据.xlsx')
    if not os.path.exists(xlsx_p):
        return []
    try:
        wb = openpyxl.load_workbook(xlsx_p, data_only=True)
        items = []
        for ws in wb.worksheets:
            if ws.sheet_state != 'visible':
                continue
            headers = [c.value for c in ws[1]]
            sheet_cat = ws.title.strip()
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not any(r):
                    continue
                d = dict(zip(headers, r))
                name = str(d.get('名称规格') or '').strip()
                if not name:
                    continue
                items.append({
                    'seq': d.get('顺序号'),
                    'name': name,
                    'brand': str(d.get('品牌') or '').strip(),
                    'unit': norm_unit(d.get('单位')),
                    'raw_unit': str(d.get('单位') or '').strip(),
                    'price': norm_money(d.get('市场价')),
                    'raw_price': d.get('市场价'),
                    'eng_price': norm_money(d.get('工程价')),
                    'spec': str(d.get('厂家原始规格参数') or ''),
                    'raw_spec': str(d.get('厂家原始规格参数') or ''),
                    'cat_code': re.sub(r'\D', '', sheet_cat),
                    'raw_cat_code': sheet_cat,
                    'province': d.get('报价地区省'),
                    'city': d.get('报价地区市'),
                    'tax': d.get('是否含增值税'),
                    'rate': d.get('增值税税率%') or d.get('增值税税率')
                })
        return items
    except Exception as e:
        print(f"Error loading theirs in {vendor_theirs_dir}: {e}")
        return []


def fold_theirs(rows):
    """折叠线上因地区展开产生的重复行"""
    groups = defaultdict(list)
    for i, r in enumerate(rows):
        key = (norm_str(r['name']), norm_unit(r['unit']), r['price'], norm_str(r['raw_spec'][:60]))
        groups[key].append(r)
    out = []
    for key, grp in groups.items():
        provs = {str(g['province']) for g in grp if g['province']}
        if len(grp) > 1 and len(provs) > 1:
            base = dict(grp[0])
            base['n_dup'] = len(grp)
            out.append(base)
        else:
            for g in grp:
                d = dict(g)
                d['n_dup'] = 1
                out.append(d)
    return out


def find_matched_ours(vname):
    # exact
    p1 = os.path.join(OURS_DIR, f"{vname}-标准解析.xlsx")
    if os.path.exists(p1) and os.path.getsize(p1) > 5000:
        return p1
    # strip trailing digits or tags
    clean = re.sub(r'(?:解析\d+|\d+)$', '', vname)
    p2 = os.path.join(OURS_DIR, f"{clean}-标准解析.xlsx")
    if os.path.exists(p2) and os.path.getsize(p2) > 5000:
        return p2
    return None


def compare_vendor(vname):
    ours_xlsx = find_matched_ours(vname)
    theirs_dir = os.path.join(THEIRS_DIR, vname)
    
    ours = load_ours_data(ours_xlsx) if ours_xlsx else []
    theirs_raw = load_theirs_data(theirs_dir)
    theirs = fold_theirs(theirs_raw)
    
    n_ours = len(ours)
    n_theirs_raw = len(theirs_raw)
    n_theirs_folded = len(theirs)
    
    # 贪心匹配
    matched_pairs = []
    used_o = set()
    used_t = set()
    
    # 第一轮：名称 + 价格 + 单位 完全匹配
    for i, o in enumerate(ours):
        for j, t in enumerate(theirs):
            if j in used_t:
                continue
            if norm_str(o['name']) == norm_str(t['name']) and o['price'] == t['price'] and (o['unit'] == t['unit'] or not o['unit'] or not t['unit']):
                matched_pairs.append((o, t))
                used_o.add(i)
                used_t.add(j)
                break
                
    # 第二轮：价格一致 + 模糊名称匹配
    for i, o in enumerate(ours):
        if i in used_o:
            continue
        for j, t in enumerate(theirs):
            if j in used_t:
                continue
            no = norm_str(o['name'])
            nt = norm_str(t['name'])
            if o['price'] == t['price'] and (no in nt or nt in no or not no or not nt):
                matched_pairs.append((o, t))
                used_o.add(i)
                used_t.add(j)
                break

    # 第三轮：名称完全一致
    for i, o in enumerate(ours):
        if i in used_o:
            continue
        for j, t in enumerate(theirs):
            if j in used_t:
                continue
            if norm_str(o['name']) == norm_str(t['name']):
                matched_pairs.append((o, t))
                used_o.add(i)
                used_t.add(j)
                break

    diffs = []
    cat_match_cnt = 0
    price_match_cnt = 0
    unit_match_cnt = 0
    name_match_cnt = 0
    
    for o, t in matched_pairs:
        # 分类比对
        is_cat_match = (o['cat_code'] == t['cat_code']) if (o['cat_code'] and t['cat_code']) else False
        if is_cat_match:
            cat_match_cnt += 1
        else:
            diffs.append({
                'vendor': vname, 'type': '分类差异',
                'name': o['name'], 'ours': f"{o['raw_cat_code']}", 'theirs': f"{t['raw_cat_code']}",
                'remark': '三级品类归属编码不一致'
            })
            
        # 价格比对
        is_price_match = (o['price'] == t['price'])
        if is_price_match:
            price_match_cnt += 1
        else:
            diffs.append({
                'vendor': vname, 'type': '市场价差异',
                'name': o['name'], 'ours': o['raw_price'], 'theirs': t['raw_price'],
                'remark': '单价读数差异'
            })
            
        # 单位比对
        is_unit_match = (o['unit'] == t['unit'])
        if is_unit_match:
            unit_match_cnt += 1
        elif o['unit'] and t['unit']:
            diffs.append({
                'vendor': vname, 'type': '单位差异',
                'name': o['name'], 'ours': o['raw_unit'], 'theirs': t['raw_unit'],
                'remark': '单位规范化差异'
            })
            
        # 名称比对
        is_name_match = (norm_str(o['name']) == norm_str(t['name']))
        if is_name_match:
            name_match_cnt += 1
        else:
            diffs.append({
                'vendor': vname, 'type': '名称差异',
                'name': o['name'], 'ours': o['name'], 'theirs': t['name'],
                'remark': '品名清洗命名口径差异'
            })

    unmatched_ours = [ours[i] for i in range(len(ours)) if i not in used_o]
    unmatched_theirs = [theirs[j] for j in range(len(theirs)) if j not in used_t]
    
    for uo in unmatched_ours:
        diffs.append({
            'vendor': vname, 'type': '我方独有行',
            'name': uo['name'], 'ours': f"单价:{uo['raw_price']} 单位:{uo['raw_unit']}", 'theirs': '-',
            'remark': '线上未收录或识别漏行'
        })
    for ut in unmatched_theirs:
        diffs.append({
            'vendor': vname, 'type': '线上独有行',
            'name': ut['name'], 'ours': '-', 'theirs': f"单价:{ut['raw_price']} 单位:{ut['raw_unit']}",
            'remark': '我方漏行或线上虚增行'
        })

    return {
        'vendor': vname,
        'ours_rows': n_ours,
        'theirs_raw_rows': n_theirs_raw,
        'theirs_folded_rows': n_theirs_folded,
        'matched_count': len(matched_pairs),
        'cat_match_cnt': cat_match_cnt,
        'price_match_cnt': price_match_cnt,
        'unit_match_cnt': unit_match_cnt,
        'diffs': diffs,
        'ours_extra_cnt': len(unmatched_ours),
        'theirs_extra_cnt': len(unmatched_theirs)
    }


def main():
    pdfs = sorted(glob.glob(os.path.join(PDF_DIR, '*.pdf')))
    print(f"=== 找到待评测 PDF 文件: {len(pdfs)} 个 ===")
    
    start_all = time.time()
    
    # 1. 批量运行解析
    def run_one_pdf(pdf_p):
        vname = os.path.splitext(os.path.basename(pdf_p))[0]
        ours_p = find_matched_ours(vname)
        if ours_p:
            print(f"[*] 已存在有效成稿，跳过重复解析: {vname}")
            return vname, ours_p, 0
        try:
            print(f"[*] 启动端到端解析: {vname} ...")
            out_file, failures = process_pdf(pdf_p, '', OURS_DIR, terms=None)
            print(f"[+] 完成解析: {vname} -> {out_file}")
            return vname, out_file, len(failures)
        except Exception as e:
            print(f"[-] 解析失败 {vname}: {e}")
            return vname, None, -1

    # 控制并发为 3，避免 API 速率限制
    with ThreadPoolExecutor(max_workers=3) as pool:
        parse_results = list(pool.map(run_one_pdf, pdfs))

    print(f"\n全量解析运行完成，耗时: {time.time() - start_all:.1f}s")
    print("开始生成对账对比报告...")

    # 2. 对账对比
    vendors = [os.path.splitext(os.path.basename(p))[0] for p in pdfs]
    comparison_results = []
    all_diff_records = []
    
    for v in vendors:
        res = compare_vendor(v)
        comparison_results.append(res)
        all_diff_records.extend(res['diffs'])

    # 3. 写入 Excel 总览与明细
    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = '1-50家解析对比总览'
    
    ws_summary.append([
        '序号', '供应商名称', '我方解析行数', '线上原始行数', '线上折叠后行数',
        '双向匹配物料数', '分类一致数', '分类一致率', '价格一致数', '价格一致率',
        '我方独有条目', '线上独有条目', '整体健康度裁决'
    ])

    for idx, c in enumerate(comparison_results, 1):
        m = c['matched_count']
        cat_rate = f"{(c['cat_match_cnt'] / max(m, 1))*100:.1f}%" if m > 0 else '0.0%'
        price_rate = f"{(c['price_match_cnt'] / max(m, 1))*100:.1f}%" if m > 0 else '0.0%'
        
        status = '通过(PASS)'
        if c['ours_rows'] == 0:
            status = '失败(解析空表)'
        elif c['theirs_extra_cnt'] > 0 and c['theirs_extra_cnt'] > c['ours_rows'] * 0.3:
            status = '需关注(漏行较多)'
        elif m > 0 and (c['price_match_cnt'] / m) < 0.8:
            status = '需关注(单价差异)'
            
        ws_summary.append([
            idx, c['vendor'], c['ours_rows'], c['theirs_raw_rows'], c['theirs_folded_rows'],
            c['matched_count'], c['cat_match_cnt'], cat_rate, c['price_match_cnt'], price_rate,
            c['ours_extra_cnt'], c['theirs_extra_cnt'], status
        ])

    ws_diff = wb.create_sheet('2-字段级差异明细表')
    ws_diff.append(['序号', '供应商名称', '差异类型', '物料名称', '我方解析值', '线上系统值', '差异分析与说明'])
    
    for didx, d in enumerate(all_diff_records, 1):
        ws_diff.append([
            didx, d['vendor'], d['type'], d['name'], str(d['ours']), str(d['theirs']), d['remark']
        ])

    # 咨询公司商业级美化
    header_fill = PatternFill('solid', fgColor='0B192C')
    header_font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
    border_thin = Border(
        left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0')
    )

    for ws in [ws_summary, ws_diff]:
        ws.freeze_panes = 'A2'
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws.row_dimensions[1].height = 28
        
        for r in range(2, ws.max_row + 1):
            ws.row_dimensions[r].height = 20
            for col in range(1, ws.max_column + 1):
                c = ws.cell(r, col)
                c.font = Font(name='微软雅黑', size=10)
                c.border = border_thin
                if col in (1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13):
                    c.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    c.alignment = Alignment(horizontal='left', vertical='center')
                    
        # 自动调整列宽
        for col in ws.columns:
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col_letter].width = min(max(max_len * 1.6, 12), 45)

    wb.save(REPORT_PATH)
    print(f"\n[Done] 对比质检总览报告成功落盘: {REPORT_PATH}")


if __name__ == '__main__':
    main()
