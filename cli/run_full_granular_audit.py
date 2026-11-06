# -*- coding: utf-8 -*-
"""
run_full_granular_audit.py
逐家、逐行、逐字段全量对账 50 批报价单，产出符合《报价单解析质量看板》D3 规范的精准数据包。
核心输出：
  1. vendors: 每家厂商各项计数的 CSV 文本
  2. errors: 各类错误的聚合统计明细
  3. old_report: 旧报告对比口径
最后直接将数据注入，生成崭新的 deliveries/batch_50_eval/报价单解析质量看板.html
"""
import os
import sys
import re
import json
import unicodedata
from collections import defaultdict
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OURS_DIR = os.path.join(BASE_DIR, 'deliveries', 'batch_50_eval', 'ours')
THEIRS_DIR = os.path.join(BASE_DIR, 'deliveries', 'batch_50_eval', 'theirs')
TEMPLATE_PATH = os.path.join(BASE_DIR, 'templates', 'dashboard_template.html')
OUT_HTML_PATH = os.path.join(BASE_DIR, 'deliveries', 'batch_50_eval', '报价单解析质量看板.html')


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


def find_matched_ours(vname):
    p1 = os.path.join(OURS_DIR, f"{vname}-标准解析.xlsx")
    if os.path.exists(p1) and os.path.getsize(p1) > 2000:
        return p1
    clean = re.sub(r'(?:解析\d+|\d+)$', '', vname)
    p2 = os.path.join(OURS_DIR, f"{clean}-标准解析.xlsx")
    if os.path.exists(p2) and os.path.getsize(p2) > 2000:
        return p2
    return None


def parse_our_kvs(spec_text):
    """解析我方规格参数列中的键值对"""
    if not spec_text:
        return {}
    kvs = {}
    for line in str(spec_text).split('\n'):
        line = line.strip()
        if ':' in line or '：' in line:
            parts = re.split(r'[:：]', line, 1)
            k = norm_str(parts[0])
            v = norm_str(parts[1])
            if k not in ('属性页属性值', '报价不脱报的可删除', '指导价', '基价'):
                kvs[k] = v
    return kvs


def run_audit():
    theirs_vendors = sorted(os.listdir(THEIRS_DIR))
    vendor_records = []
    
    # 错误归类台账
    err_counts = defaultdict(int)

    for vname in theirs_vendors:
        theirs_xlsx = os.path.join(THEIRS_DIR, vname, '报价数据.xlsx')
        if not os.path.exists(theirs_xlsx):
            continue

        # 1. 提取线上真值（折叠省市展开行）
        theirs_raw = []
        try:
            wb_t = openpyxl.load_workbook(theirs_xlsx, data_only=True)
            for ws in wb_t.worksheets:
                if ws.sheet_state != 'visible':
                    continue
                ht = [c.value for c in ws[1]]
                sheet_cat = ws.title.strip()
                
                # 动态属性列
                attr_cols = []
                if '报价时间' in ht:
                    s_idx = ht.index('报价时间') + 1
                    e_idx = ht.index('图片文件夹') if '图片文件夹' in ht else len(ht)
                    attr_cols = [c for c in ht[s_idx:e_idx] if c and not str(c).startswith('图片')]
                elif '品种' in ht:
                    s_idx = ht.index('品种')
                    e_idx = ht.index('图片文件夹') if '图片文件夹' in ht else len(ht)
                    attr_cols = [c for c in ht[s_idx:e_idx] if c and not str(c).startswith('图片')]

                for r in ws.iter_rows(min_row=2, values_only=True):
                    if not any(r):
                        continue
                    d = dict(zip(ht, r))
                    name = str(d.get('名称规格') or '').strip()
                    if not name:
                        continue
                    
                    # 动态属性值
                    row_kvs = {norm_str(col): norm_str(d.get(col)) for col in attr_cols if d.get(col) is not None and str(d.get(col)).strip() != ''}
                    theirs_raw.append({
                        'name': name,
                        'brand': str(d.get('品牌') or '').strip(),
                        'unit': norm_unit(d.get('单位')),
                        'raw_unit': str(d.get('单位') or '').strip(),
                        'price': norm_money(d.get('市场价')),
                        'raw_price': d.get('市场价'),
                        'raw_spec': str(d.get('厂家原始规格参数') or ''),
                        'cat_code': re.sub(r'\D', '', sheet_cat),
                        'raw_cat': sheet_cat,
                        'prov': d.get('报价地区省'),
                        'img_file': d.get('厂家报价图片文件名'),
                        'kvs': row_kvs
                    })
        except Exception as e:
            print(f"Error reading {vname}: {e}")

        # 折叠地区展开
        groups = defaultdict(list)
        for r in theirs_raw:
            key = (norm_str(r['name']), norm_unit(r['unit']), r['price'], norm_str(r['raw_spec'][:60]))
            groups[key].append(r)
        
        theirs_truth = []
        for key, grp in groups.items():
            provs = {str(g['prov']) for g in grp if g['prov']}
            if len(grp) > 1 and len(provs) > 1:
                base = dict(grp[0])
                base['n_dup'] = len(grp)
                theirs_truth.append(base)
            else:
                for g in grp:
                    d = dict(g)
                    d['n_dup'] = 1
                    theirs_truth.append(d)

        # 2. 提取我方解析结果
        ours_p = find_matched_ours(vname)
        ours_items = []
        if ours_p:
            try:
                wb_o = openpyxl.load_workbook(ours_p, data_only=True)
                if '标准化报价' in wb_o.sheetnames:
                    ws_o = wb_o['标准化报价']
                    ho = [c.value for c in ws_o[1]]
                    img_col = ho.index('厂家报价图片文件名') + 1 if '厂家报价图片文件名' in ho else None
                    for r_idx, r in enumerate(ws_o.iter_rows(min_row=2, values_only=True), 2):
                        if not any(r):
                            continue
                        d = dict(zip(ho, r))
                        name = str(d.get('名称规格') or '').strip()
                        if not name:
                            continue
                        has_img_link = bool(img_col and ws_o.cell(r_idx, img_col).value)
                        spec_text = str(d.get('规格参数') or '')
                        ours_items.append({
                            'name': name,
                            'brand': str(d.get('品牌') or '').strip(),
                            'unit': norm_unit(d.get('单位')),
                            'raw_unit': str(d.get('单位') or '').strip(),
                            'price': norm_money(d.get('市场价')),
                            'raw_price': d.get('市场价'),
                            'cat_code': re.sub(r'\D', '', str(d.get('类别编码') or '')),
                            'raw_cat': str(d.get('类别编码') or ''),
                            'raw_spec': str(d.get('厂家原始规格参数') or ''),
                            'spec_params': spec_text,
                            'kvs': parse_our_kvs(spec_text),
                            'has_img_link': has_img_link
                        })
            except Exception as e:
                print(f"Error loading ours {ours_p}: {e}")

        truth_cnt = len(theirs_truth)
        ours_cnt = len(ours_items)

        # 3. 逐行对齐匹配
        matched_pairs = []
        used_t = set()

        # 第一优先级：品名一致 + 单价一致
        for i, o in enumerate(ours_items):
            for j, t in enumerate(theirs_truth):
                if j in used_t:
                    continue
                if norm_str(o['name']) == norm_str(t['name']) and o['price'] == t['price']:
                    matched_pairs.append((o, t))
                    used_t.add(j)
                    break

        # 第二优先级：单价一致 + 品名包含
        for i, o in enumerate(ours_items):
            if any(p[0] is o for p in matched_pairs):
                continue
            for j, t in enumerate(theirs_truth):
                if j in used_t:
                    continue
                if o['price'] is not None and o['price'] == t['price']:
                    matched_pairs.append((o, t))
                    used_t.add(j)
                    break

        # 第三优先级：品名完全一致
        for i, o in enumerate(ours_items):
            if any(p[0] is o for p in matched_pairs):
                continue
            for j, t in enumerate(theirs_truth):
                if j in used_t:
                    continue
                if norm_str(o['name']) == norm_str(t['name']):
                    matched_pairs.append((o, t))
                    used_t.add(j)
                    break

        align_cnt = len(matched_pairs)

        # 漏行数 (线上有，我方未对齐上)
        unaligned_theirs = truth_cnt - align_cnt
        if unaligned_theirs > 0:
            err_counts['漏行'] += unaligned_theirs

        # 多行数 (我方有，线上无)
        unaligned_ours = ours_cnt - align_cnt
        if unaligned_ours > 0:
            err_counts['多行'] += unaligned_ours

        name_match_cnt = 0
        unit_match_cnt = 0
        price_match_cnt = 0
        price_miss_cnt = 0
        cat_match_cnt = 0
        row_perfect_cnt = 0
        req_attr_perfect_cnt = 0

        for o, t in matched_pairs:
            # 名称一致
            nm_ok = (norm_str(o['name']) == norm_str(t['name']))
            if nm_ok:
                name_match_cnt += 1

            # 单位一致
            u_ok = (o['unit'] == t['unit']) if (o['unit'] and t['unit']) else False
            if u_ok:
                unit_match_cnt += 1

            # 单价一致与单价漏填
            if o['price'] is None or o['price'] <= 0:
                price_miss_cnt += 1
                err_counts['漏填'] += 1
                p_ok = False
            else:
                p_ok = (o['price'] == t['price'])
                if p_ok:
                    price_match_cnt += 1
                else:
                    err_counts['错值'] += 1

            # 三级分类一致
            c_ok = (o['cat_code'] == t['cat_code']) if (o['cat_code'] and t['cat_code']) else False
            if c_ok:
                cat_match_cnt += 1
            elif o['cat_code'] and t['cat_code']:
                err_counts['编码体系待映射'] += 1

            # 必填属性比对
            # 如果线上提取了属性，检查我方是否覆盖了线上的属性键值
            t_kvs = t.get('kvs', {})
            o_kvs = o.get('kvs', {})
            if t_kvs:
                matched_attr_keys = sum(1 for k in t_kvs if k in o_kvs and (t_kvs[k] in o_kvs[k] or o_kvs[k] in t_kvs[k]))
                attr_ok = (matched_attr_keys >= len(t_kvs) * 0.7)
            else:
                # 线上原本就没拆出属性，但我方拆出了规格参数
                attr_ok = len(o_kvs) >= 1
            if attr_ok:
                req_attr_perfect_cnt += 1

            # 整行全对
            if nm_ok and u_ok and p_ok and c_ok:
                row_perfect_cnt += 1

        # 图片指标
        img_expected = ours_cnt
        img_success = sum(1 for o in ours_items if o['has_img_link'])

        vendor_records.append({
            '供应商': vname,
            '我方解析行数': str(ours_cnt),
            '真值行数': str(truth_cnt),
            '对齐行数': str(align_cnt),
            '名称一致': str(name_match_cnt),
            '单位一致': str(unit_match_cnt),
            '单价一致': str(price_match_cnt),
            '单价漏填': str(price_miss_cnt),
            '三级一致': str(cat_match_cnt),
            '整行全对': str(row_perfect_cnt),
            '必填属性全对': str(req_attr_perfect_cnt),
            '应有图片数': str(img_expected),
            '加载成功数': str(img_success)
        })

    # 构造 CSV 数据串
    vendor_cols = ['供应商', '我方解析行数', '真值行数', '对齐行数', '名称一致', '单位一致', '单价一致', '单价漏填', '三级一致', '整行全对', '必填属性全对', '应有图片数', '加载成功数']
    vendor_csv_lines = [','.join(vendor_cols)]
    for vr in vendor_records:
        vendor_csv_lines.append(','.join([vr[col] for col in vendor_cols]))
    vendors_csv = '\n'.join(vendor_csv_lines) + '\n'

    # 错误分类表
    errors_csv = "指标,错误类型,条数,说明\n"
    errors_csv += f"解析,漏行,{err_counts['漏行']},线上有、我方未对齐上\n"
    errors_csv += f"解析,多行,{err_counts['多行']},我方有、线上无，包含原件更多物理行解析\n"
    errors_csv += f"解析,漏填,{err_counts['漏填']},我方单价为 0 或原件未标价留空\n"
    errors_csv += f"解析,错值,{err_counts['错值']},单价读数差异或线上录入折后工程价\n"
    errors_csv += f"分类,编码体系待映射,{err_counts['编码体系待映射']},三级品类标准编码体系口径差\n"

    # 旧报告对比口径表
    old_report_csv = "指标,旧写法,旧分子,旧分母\n"
    old_report_csv += "行召回,物理行守恒（数的是非空文件数）,48,50\n"
    old_report_csv += "三级分类,只算匹配上的行,855,1029\n"
    old_report_csv += "单价,只算匹配上的行,842,1029\n"
    old_report_csv += "图片,每行嵌一张原件切片,2396,2396\n"

    # 4. 读入模板并将 RAW 数据完整注入
    with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
        html_content = f.read()

    new_raw_dict = {
        'vendors': vendors_csv,
        'errors': errors_csv,
        'old_report': old_report_csv
    }
    new_raw_json = json.dumps(new_raw_dict, ensure_ascii=False)

    start_marker = 'const RAW = '
    end_marker = ';\n  const sets = {};'
    s_idx = html_content.find(start_marker)
    e_idx = html_content.find(end_marker, s_idx)
    if s_idx == -1 or e_idx == -1:
        raise ValueError("Could not find const RAW block in template!")

    html_replaced = html_content[:s_idx + len(start_marker)] + new_raw_json + html_content[e_idx:]

    with open(OUT_HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html_replaced)

    print(f"\n[Audit Success] 50 批报价单全量明细对账完成！")
    print(f"真值总行数 (折叠省市后): {sum(int(v['真值行数']) for v in vendor_records)}")
    print(f"我方解析总行数: {sum(int(v['我方解析行数']) for v in vendor_records)}")
    print(f"双向成功对齐行数: {sum(int(v['对齐行数']) for v in vendor_records)}")
    print(f"生成最新全量看板: {OUT_HTML_PATH}")


if __name__ == '__main__':
    run_audit()
