# -*- coding: utf-8 -*-
"""对 deliveries 下全部交付件做独立体检（只读，不改文件）"""
import glob
import os
import re
import sys
import openpyxl

DELIV = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'deliveries')

# 规格参数里不该出现的东西（README/提示词约定：基价、单价、表标题、指导价、税率等）
FORBIDDEN_IN_SPEC = ('铝锭基价', '铜基价', '基价', '指导价', '表标题', '表头', '单价', '总价', '金额', '税率', '运费')
# 报价说明里不该出现的系统内部告警（原则：只记客户商务原文）
INTERNAL_PREFIXES = ('未经大模型清洗', '参数无原文出处', '原表未读到有效单价', '模型返回价格', '商务条款缺失')
TRACE_MARK = re.compile(r'\[模型自愈\]|\[规格回填\]')


def audit_one(path):
    errs, warns, info = [], [], {}
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        return [f'无法打开: {e}'], [], {}
    names = wb.sheetnames
    info['sheets'] = names
    if names != ['标准化报价']:
        (errs if len(names) > 1 else warns).append(f'Sheet 非单表: {names}')
    if '标准化报价' not in names:
        return errs + ['缺少标准化报价'], warns, info
    ws = wb['标准化报价']
    col = {str(c.value): c.column for c in ws[1] if c.value}
    need = ['报价ID', '市场价', '名称规格', '单位', '规格参数', '报价说明']
    missing = [h for h in need if h not in col]
    if missing:
        return errs + [f'缺列: {missing}'], warns, info

    rows = 0
    ids = []
    no_unit = 0
    traced = 0
    spec_dirty = 0
    desc_dirty = 0
    zero_price = 0
    for r in range(2, ws.max_row + 1):
        qid = ws.cell(r, col['报价ID']).value
        if not qid:
            continue
        rows += 1
        ids.append(str(qid))
        p = ws.cell(r, col['市场价']).value
        if not isinstance(p, (int, float)):
            errs.append(f'第{r}行 市场价非数值: {p!r}')
        elif p <= 0:
            zero_price += 1
        if not ws.cell(r, col['单位']).value:
            no_unit += 1
        spec = str(ws.cell(r, col['规格参数']).value or '')
        hit = [k for k in FORBIDDEN_IN_SPEC if k in spec]
        if hit:
            spec_dirty += 1
        desc = str(ws.cell(r, col['报价说明']).value or '')
        if any(k in desc for k in INTERNAL_PREFIXES):
            desc_dirty += 1
        if TRACE_MARK.search(desc) or TRACE_MARK.search(spec):
            traced += 1
        if not str(ws.cell(r, col['名称规格']).value or '').strip():
            errs.append(f'第{r}行 名称规格为空')

    dup = [q for q in set(ids) if ids.count(q) > 1]
    if dup:
        errs.append(f'报价ID 重复: {dup[:3]}')
    if rows == 0:
        errs.append('零数据行')

    info.update(rows=rows, no_unit=no_unit, traced=traced,
                spec_dirty=spec_dirty, desc_dirty=desc_dirty, zero_price=zero_price)
    if zero_price:
        warns.append(f'{zero_price} 行市场价<=0')
    if spec_dirty:
        warns.append(f'{spec_dirty} 行规格参数含非技术参数(基价/单价/表标题等)')
    if desc_dirty:
        warns.append(f'{desc_dirty} 行报价说明含系统内部告警')
    if no_unit:
        warns.append(f'{no_unit} 行无计量单位')
    if traced:
        warns.append(f'{traced} 行带留痕标记')
    return errs, warns, info


def main():
    files = sorted(glob.glob(os.path.join(DELIV, '*-标准解析.xlsx')))
    if len(sys.argv) > 1:
        files = files[:int(sys.argv[1])]
    bad = 0
    agg = {'files': 0, 'rows': 0, 'no_unit': 0, 'spec_dirty': 0, 'desc_dirty': 0,
           'traced': 0, 'zero_price': 0, 'no_cert': 0}
    for f in files:
        errs, warns, info = audit_one(f)
        agg['files'] += 1
        for k in ('rows', 'no_unit', 'spec_dirty', 'desc_dirty', 'traced', 'zero_price'):
            agg[k] += info.get(k, 0)
        base = os.path.basename(f)
        if not os.path.exists(f.replace('.xlsx', '_对账与防欺瞒体检报告.xlsx')):
            agg['no_cert'] += 1
            warns.append('无独立体检报告')
        if errs:
            bad += 1
            print(f'[ERR ] {base}')
            for e in errs[:4]:
                print(f'        ✗ {e}')
        elif warns:
            print(f'[WARN] {base} | {info.get("rows",0)} 行')
        for w in warns[:5]:
            print(f'        - {w}')

    print('\n' + '=' * 72)
    print(f"交付件 {agg['files']} 个 | 总行数 {agg['rows']} | 有硬错误 {bad} 个")
    print(f"  规格参数被非技术参数污染: {agg['spec_dirty']} 行")
    print(f"  报价说明含系统内部告警  : {agg['desc_dirty']} 行")
    print(f"  无计量单位              : {agg['no_unit']} 行")
    print(f"  市场价<=0               : {agg['zero_price']} 行")
    print(f"  带留痕标记(自愈/回填)   : {agg['traced']} 行")
    print(f"  缺独立体检报告          : {agg['no_cert']} 个")
    print('=' * 72)


if __name__ == '__main__':
    main()
