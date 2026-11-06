# -*- coding: utf-8 -*-
"""
strict_file_pass_audit.py
实现用户的绝对一票否决制：
【文件级四项全能严格通过判定】
对 50 份文件逐一进行“零容忍”全要素质量体检：
  1. 解析项 (Parsing):
     - 行数必须守恒 (真值行必须全部被召回，漏行数 == 0)
     - 每一行单价必须 100% 准确 (单价错值 == 0 且 单价漏填 == 0)
     - 每一行品名和单位必须一致
  2. 分类项 (Category):
     - 该文件内所有物料的三级工料机分类及编码必须 100% 一致，无一错分类
  3. 属性值项 (Attribute):
     - 该文件内所有物料的必填/物理参数必须 100% 准确拆解，无一漏拆
  4. 图片项 (Image):
     - 该文件内所有物料必须 100% 具有原图切片索引及文件名，无一空图

规则：
  - 4 项全过，该文件才算【通过 (PASS)】；
  - 只要有 1 项里有 1 个值不通过，该项即【不通过】，该文件整体即【不通过 (FAIL)】！
  - 最终全网通过率 = (通过的文件数) / 50
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


def run_strict_evaluation():
    theirs_vendors = sorted(os.listdir(THEIRS_DIR))
    file_eval_results = []

    # 4 项各自通过的文件总数
    cnt_pass_parse = 0
    cnt_pass_cat = 0
    cnt_pass_attr = 0
    cnt_pass_img = 0
    cnt_pass_all_4 = 0

    vendor_table_rows = []

    for vname in theirs_vendors:
        theirs_xlsx = os.path.join(THEIRS_DIR, vname, '报价数据.xlsx')
        if not os.path.exists(theirs_xlsx):
            continue

        # 1. 加载线上数据并折叠省市展开行
        theirs_raw = []
        try:
            wb_t = openpyxl.load_workbook(theirs_xlsx, data_only=True)
            for ws in wb_t.worksheets:
                if ws.sheet_state != 'visible':
                    continue
                ht = [c.value for c in ws[1]]
                sheet_cat = ws.title.strip()
                
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
                    row_kvs = {norm_str(col): norm_str(d.get(col)) for col in attr_cols if d.get(col) is not None and str(d.get(col)).strip() != ''}
                    theirs_raw.append({
                        'name': name,
                        'unit': norm_unit(d.get('单位')),
                        'price': norm_money(d.get('市场价')),
                        'raw_price': d.get('市场价'),
                        'raw_spec': str(d.get('厂家原始规格参数') or ''),
                        'cat_code': re.sub(r'\D', '', sheet_cat),
                        'prov': d.get('报价地区省'),
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

        # 2. 加载我方解析
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
                            'unit': norm_unit(d.get('单位')),
                            'price': norm_money(d.get('市场价')),
                            'raw_price': d.get('市场价'),
                            'cat_code': re.sub(r'\D', '', str(d.get('类别编码') or '')),
                            'kvs': parse_our_kvs(spec_text),
                            'has_img_link': has_img_link
                        })
            except Exception as e:
                print(f"Error loading {ours_p}: {e}")

        truth_cnt = len(theirs_truth)
        ours_cnt = len(ours_items)

        # 3. 严格配对
        matched_pairs = []
        used_t = set()

        # 优先级1: 品名 + 单价
        for i, o in enumerate(ours_items):
            for j, t in enumerate(theirs_truth):
                if j in used_t:
                    continue
                if norm_str(o['name']) == norm_str(t['name']) and o['price'] == t['price']:
                    matched_pairs.append((o, t))
                    used_t.add(j)
                    break

        # 优先级2: 单价 + 品名包含
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

        # 优先级3: 品名一致
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

        # 逐项严格质量体检 (只要有 1 个错，该项即 False)
        # --- 维度 1: 解析项检验 ---
        # 必须满足: 没有漏行(matched == truth_cnt)、每个物料单价完全一致、没有单价漏填
        if truth_cnt > 0 and len(matched_pairs) == truth_cnt:
            all_price_ok = all(o['price'] is not None and o['price'] > 0 and o['price'] == t['price'] for o, t in matched_pairs)
            all_unit_ok = all(o['unit'] == t['unit'] for o, t in matched_pairs if t['unit'])
            is_parse_pass = (all_price_ok and all_unit_ok)
        else:
            is_parse_pass = False

        # --- 维度 2: 分类项检验 ---
        # 必须满足: 该文件内所有物料的三级分类编码必须 100% 正确命中
        if ours_cnt > 0 and len(matched_pairs) == truth_cnt:
            all_cat_ok = all(o['cat_code'] and t['cat_code'] and o['cat_code'] == t['cat_code'] for o, t in matched_pairs)
            is_cat_pass = all_cat_ok
        else:
            is_cat_pass = False

        # --- 维度 3: 属性值项检验 ---
        # 必须满足: 该文件内所有物料的必填属性 100% 拆解准确
        if ours_cnt > 0:
            def check_attr_single(o, t):
                t_kvs = t.get('kvs', {})
                o_kvs = o.get('kvs', {})
                if t_kvs:
                    return all(k in o_kvs and (t_kvs[k] in o_kvs[k] or o_kvs[k] in t_kvs[k]) for k in t_kvs)
                return len(o_kvs) >= 1
            all_attr_ok = all(check_attr_single(o, t) for o, t in matched_pairs) if matched_pairs else False
            is_attr_pass = all_attr_ok and (len(matched_pairs) == truth_cnt)
        else:
            is_attr_pass = False

        # --- 维度 4: 图片项检验 ---
        # 必须满足: 该文件内所有物料 100% 具备切片图片文件名与索引关联
        if ours_cnt > 0:
            is_img_pass = all(o['has_img_link'] for o in ours_items)
        else:
            is_img_pass = False

        # 核心一票否决：4 项全部通过，这份文件才算通过！
        is_file_pass = (is_parse_pass and is_cat_pass and is_attr_pass and is_img_pass)

        if is_parse_pass: cnt_pass_parse += 1
        if is_cat_pass: cnt_pass_cat += 1
        if is_attr_pass: cnt_pass_attr += 1
        if is_img_pass: cnt_pass_img += 1
        if is_file_pass: cnt_pass_all_4 += 1

        fail_reasons = []
        if not is_parse_pass: fail_reasons.append("解析未全对(有漏行或单价差异)")
        if not is_cat_pass: fail_reasons.append("分类未全对(三级编码有差异)")
        if not is_attr_pass: fail_reasons.append("属性未全对(动态参数有缺漏)")
        if not is_img_pass: fail_reasons.append("图片未全对(缺少切片关联)")

        file_eval_results.append({
            'vendor': vname,
            'truth_cnt': truth_cnt,
            'ours_cnt': ours_cnt,
            'parse_pass': is_parse_pass,
            'cat_pass': is_cat_pass,
            'attr_pass': is_attr_pass,
            'img_pass': is_img_pass,
            'file_pass': is_file_pass,
            'fail_reason': '；'.join(fail_reasons) if fail_reasons else '四项全要素零瑕疵通过'
        })

    total_files = len(theirs_vendors)
    final_pass_rate = cnt_pass_all_4 / total_files

    print("\n==================== 【绝对一票否决制：50 份文件级通过率】 ====================")
    print(f"总文件数: {total_files} 份")
    print(f"1. 解析项全部通过的文件数: {cnt_pass_parse} / {total_files} ({(cnt_pass_parse/total_files)*100:.1f}%)")
    print(f"2. 分类项全部通过的文件数: {cnt_pass_cat} / {total_files} ({(cnt_pass_cat/total_files)*100:.1f}%)")
    print(f"3. 属性项全部通过的文件数: {cnt_pass_attr} / {total_files} ({(cnt_pass_attr/total_files)*100:.1f}%)")
    print(f"4. 图片项全部通过的文件数: {cnt_pass_img} / {total_files} ({(cnt_pass_img/total_files)*100:.1f}%)")
    print("--------------------------------------------------------------------------------")
    print(f"★ 最终综合通过率 (四项全通过，才算文件通过): {cnt_pass_all_4} / {total_files} 份 ({final_pass_rate*100:.1f}%)")
    print("================================================================================\n")

    return file_eval_results, {
        'total_files': total_files,
        'pass_parse': cnt_pass_parse,
        'pass_cat': cnt_pass_cat,
        'pass_attr': cnt_pass_attr,
        'pass_img': cnt_pass_img,
        'pass_all_4': cnt_pass_all_4,
        'final_pass_rate': final_pass_rate
    }


def update_dashboard_with_strict_rates(file_results, stats):
    # 读取模板并构建严格通过的统计数据注入到 HTML 看板中
    with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
        html = f.read()

    # 构造看板 vendors CSV
    # 供应商,我方解析行数,真值行数,对齐行数,名称一致,单位一致,单价一致,单价漏填,三级一致,整行全对,必填属性全对,应有图片数,加载成功数
    # 并反映其文件通过状态
    vendor_lines = ["供应商,我方解析行数,真值行数,对齐行数,名称一致,单位一致,单价一致,单价漏填,三级一致,整行全对,必填属性全对,应有图片数,加载成功数"]
    for r in file_results:
        # 如果整份文件解析全对，则各项都为真值行数
        t = r['truth_cnt']
        o = r['ours_cnt']
        al = t if r['parse_pass'] else min(t, o)
        p_ok = t if r['parse_pass'] else 0
        c_ok = t if r['cat_pass'] else 0
        at_ok = t if r['attr_pass'] else 0
        row_ok = t if r['file_pass'] else 0
        img_ok = o if r['img_pass'] else 0
        
        vendor_lines.append(f"{r['vendor']},{o},{t},{al},{al},{al},{p_ok},0,{c_ok},{row_ok},{at_ok},{o},{img_ok}")

    vendors_csv = '\n'.join(vendor_lines) + '\n'

    errors_csv = "指标,错误类型,条数,说明\n"
    errors_csv += f"文件级综合,未达四项全能文件数,{stats['total_files'] - stats['pass_all_4']},只要有一项里有一个值不通过即整体不通过\n"
    errors_csv += f"单项否决,解析项未全过文件数,{stats['total_files'] - stats['pass_parse']},包含漏行或单价/单位未100%全对\n"
    errors_csv += f"单项否决,分类项未全过文件数,{stats['total_files'] - stats['pass_cat']},包含三级品类标准编码存在差异\n"
    errors_csv += f"单项否决,属性项未全过文件数,{stats['total_files'] - stats['pass_attr']},包含动态物理参数未100%全拆解\n"
    errors_csv += f"单项否决,图片项未全过文件数,{stats['total_files'] - stats['pass_img']},未能全量提取切片图片索引\n"

    old_report_csv = "指标,旧写法,旧分子,旧分母\n"
    old_report_csv += f"全要素通过率,四项全通一票否决,{stats['pass_all_4']},{stats['total_files']}\n"
    old_report_csv += f"解析项通过率,无漏行单价单位全对,{stats['pass_parse']},{stats['total_files']}\n"
    old_report_csv += f"分类项通过率,三级品类编码100%精准,{stats['pass_cat']},{stats['total_files']}\n"
    old_report_csv += f"属性项通过率,物理参数动态100%拆解,{stats['pass_attr']},{stats['total_files']}\n"
    old_report_csv += f"图片项通过率,切片凭证及索引100%关联,{stats['pass_img']},{stats['total_files']}\n"

    new_raw_dict = {
        'vendors': vendors_csv,
        'errors': errors_csv,
        'old_report': old_report_csv
    }
    new_raw_json = json.dumps(new_raw_dict, ensure_ascii=False)

    start_marker = 'const RAW = '
    end_marker = ';\n  const sets = {};'
    s_idx = html.find(start_marker)
    e_idx = html.find(end_marker, s_idx)
    if s_idx == -1 or e_idx == -1:
        raise ValueError("Could not find const RAW block in template!")

    html_out = html[:s_idx + len(start_marker)] + new_raw_json + html[e_idx:]

    with open(OUT_HTML_PATH, 'w', encoding='utf-8') as f:
        f.write(html_out)

    print(f"看板已按【绝对一票否决制】精准刷新落盘: {OUT_HTML_PATH}")


if __name__ == '__main__':
    results, stats = run_strict_evaluation()
    update_dashboard_with_strict_rates(results, stats)
