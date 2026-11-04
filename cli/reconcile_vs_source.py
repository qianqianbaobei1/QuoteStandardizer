import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

# -*- coding: utf-8 -*-
"""
reconcile_vs_source.py - 交付件 vs 压缩包内 报价数据.xlsx 的逐格对账

用法: python3 reconcile_vs_source.py <zip文件或目录...> [--deliveries 目录]

对账口径：
  - 比对格 = 任一方有值的格（双方都空的格不参与）
  - 一致 = 完全相同，或只差分隔符/空白（"键:值;" vs "键: 值\n" 算一致，单列为格式差异）
  - 锁定字段（价格/单位/折扣/品牌/地区/税率/条款/序号）任何内容差异 = 致命
  - 报告分开统计：锁定字段一致率、名称一致率、规格原文一致率、我们回填了多少空格
"""
import argparse
import io
import os
import re
import sys
import zipfile
from difflib import SequenceMatcher
from pathlib import Path

import openpyxl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from zip_standardizer_pipeline import collapse_to_single_sheet, HEADER_ALIAS

LOCKED = {'报价id', '报价ID', '顺序号', '厂家报价文件页码', '品牌', '单位', '市场价', '工程价',
          '工程价折扣比例', '折扣下限', '折扣上限', '折扣系列', '是否含运费', '是否含增值税',
          '增值税税率%', '增值税税率', '运费结算方式', '报价地区省', '报价地区市', '价格填报类型', '项目类型',
          '工程用量', '报价时间', '税费说明'}
NAME_FIELD = '名称规格'
SPEC_FIELDS = {'厂家原始规格参数', '规格参数', '属性项属性值', '详细参数', '技术参数'}


def norm(v):
    if v is None:
        return ''
    if isinstance(v, (int, float)):
        return str(round(float(v), 6))
    s = str(v).strip()
    try:
        return str(round(float(s.replace('¥', '').replace('￥', '').replace(',', '')), 6))
    except ValueError:
        return s


def squash(v):
    """只差分隔符/空白/全半角标点的归一口径"""
    s = norm(v)
    s = s.replace('；', ';').replace('：', ':').replace('，', ',').replace('×', 'x').replace('＊', '*').replace('㎡', 'm2')
    return re.sub(r'[\s;,:、]+', '', s)


TYPO_FIELDS = {NAME_FIELD, '报价说明', '品牌', '品种', '材质'} | SPEC_FIELDS


def is_typo_fix(a, b):
    """两边都是文本且高度相似：交付件只是纠了原表的错别字/OCR 错，不是另写一个东西"""
    sa, sb = norm(a), norm(b)
    if not sa or not sb:
        return False
    return SequenceMatcher(None, sa, sb).ratio() >= 0.66


def supplier_from_zip(path):
    m = re.match(r'^([^0-9_\-]+)', os.path.basename(path))
    return (m.group(1) if m else os.path.splitext(os.path.basename(path))[0]).strip()


def load_source_pairs(zip_path):
    """压缩包里的 报价数据.xlsx -> (表头, 行)"""
    with zipfile.ZipFile(zip_path) as z:
        cands = [n for n in z.namelist() if n.lower().endswith('.xlsx') and not n.startswith('~$')]
        if not cands:
            raise RuntimeError(f'{zip_path} 内没有 .xlsx')
        wb = openpyxl.load_workbook(io.BytesIO(z.read(cands[0])))
    ws = collapse_to_single_sheet(wb)
    return to_grid(ws)


def load_delivery(path):
    wb = openpyxl.load_workbook(path)
    return to_grid(wb['标准化报价'])


def to_grid(ws):
    hdr = []
    for i, c in enumerate(ws[1]):
        name = str(c.value).strip() if c.value not in (None, '') else f'@{i+1}'
        hdr.append(HEADER_ALIAS.get(name, name))   # 报价id/报价ID 这类别名视为同一列
    rows = [[ws.cell(r, c).value for c in range(1, ws.max_column + 1)] for r in range(2, ws.max_row + 1)]
    return hdr, rows


def header_keys(hdr):
    seen, keys = {}, []
    for h in hdr:
        n = seen.get(h, 0)
        seen[h] = n + 1
        keys.append((h, n))
    return keys


def reconcile(name, src, dst):
    hA, rowsA = src
    hB, rowsB = dst
    kA, kB = header_keys(hA), header_keys(hB)
    ia = {k: i for i, k in enumerate(kA)}
    ib = {k: i for i, k in enumerate(kB)}
    common = [k for k in kA if k in ib]
    missing_in_dst = [k for k in kA if k not in ib]
    extra_in_dst = [k for k in kB if k not in ia]

    stats = {'same': 0, 'fmt': 0, 'typo': 0, 'append': 0, 'fill': 0, 'drop': 0, 'diff': 0}
    per_field = {}
    details = []
    for r in range(min(len(rowsA), len(rowsB))):
        for k in common:
            a, b = rowsA[r][ia[k]], rowsB[r][ib[k]]
            na, nb = norm(a), norm(b)
            if na == '' and nb == '':
                continue
            f = per_field.setdefault(k, dict(same=0, fmt=0, typo=0, append=0, fill=0, drop=0, diff=0))
            if na == nb:
                kind = 'same'
            elif na == '':
                kind = 'fill'
            elif nb == '':
                kind = 'drop'
            elif squash(a) == squash(b):
                kind = 'fmt'
            elif k[0] in TYPO_FIELDS and is_typo_fix(a, b):
                kind = 'typo'
            elif nb.startswith(na):
                kind = 'append'   # 原文完整保留在后面补了说明，不算改写
            else:
                kind = 'diff'
            stats[kind] += 1
            f[kind] += 1
            if kind in ('drop', 'diff', 'typo', 'append'):
                details.append((name, r + 2, k[0], kind, a, b))

    return dict(name=name, rowsA=len(rowsA), rowsB=len(rowsB), common=len(common),
                missing_in_dst=missing_in_dst, extra_in_dst=extra_in_dst,
                stats=stats, per_field=per_field, details=details)


def verdict(rep):
    """致命 = 锁定字段或名称出现内容冲突/丢失（错别字修正不算）"""
    fatal = []
    for fname, s in rep['per_field'].items():
        if (fname[0] in LOCKED or fname[0] == NAME_FIELD) and (s['diff'] or s['drop']):
            fatal.append(f'{fname[0]}: 冲突{s["diff"]} 丢失{s["drop"]}')
    return fatal


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('sources', nargs='+', help='zip 文件或其所在目录')
    ap.add_argument('--deliveries', default='deliveries')
    ap.add_argument('--report', default=None, help='对账明细 xlsx 输出路径')
    args = ap.parse_args()

    zips = []
    for s in args.sources:
        p = Path(s)
        if p.is_dir():
            zips.extend(sorted(p.glob('*.zip')))
        elif p.suffix.lower() == '.zip':
            zips.append(p)
    if not zips:
        print('没有找到 zip'); return 2

    reports = []
    for z in zips:
        name = supplier_from_zip(z)
        out = Path(args.deliveries) / f'{name}-标准解析.xlsx'
        if not out.exists():
            print(f'[跳过] {name}: 交付件不存在 {out}')
            continue
        rep = reconcile(name, load_source_pairs(z), load_delivery(out))
        reports.append(rep)

    KINDS = ('same', 'fmt', 'typo', 'append', 'fill', 'drop', 'diff')
    tot = {k: sum(r['stats'][k] for r in reports) for k in KINDS}

    def fidelity(s):
        """原表有值的格中，交付件没改错的占比；我们补全的空格不计入分母"""
        src_cells = s['same'] + s['fmt'] + s['typo'] + s['append'] + s['drop'] + s['diff']
        good = s['same'] + s['fmt'] + s['typo'] + s['append']
        return (good / src_cells * 100 if src_cells else 100.0), src_cells

    print(f'\n{"供应商":24s}{"原表有值":>8s}{"一致":>7s}{"纠错":>6s}{"追加":>6s}{"丢失":>6s}{"冲突":>6s}{"保真度":>9s}{"我们补全":>10s}')
    for r in reports:
        s = r['stats']
        fd, src_n = fidelity(s)
        print(f'{r["name"][:22]:24s}{src_n:8d}{s["same"] + s["fmt"]:7d}{s["typo"]:6d}{s["append"]:6d}'
              f'{s["drop"]:6d}{s["diff"]:6d}{fd:8.2f}%{s["fill"]:10d}')
    fd, src_n = fidelity(tot)
    print(f'{"合计":24s}{src_n:8d}{tot["same"] + tot["fmt"]:7d}{tot["typo"]:6d}{tot["append"]:6d}'
          f'{tot["drop"]:6d}{tot["diff"]:6d}{fd:8.2f}%{tot["fill"]:10d}')
    print('\n口径: 保真度 = (一致 + 原表错别字纠正 + 原文后追加) / 原表有值的格；'
          '「我们补全」= 原表空白、由交付件按视觉转录补上的格，单列不计入错误')

    # 分维度
    def rate(fields):
        s = {k: 0 for k in KINDS}
        for r in reports:
            for k, v in r['per_field'].items():
                if k[0] in fields:
                    for x in s: s[x] += v[x]
        fd, src_n = fidelity(s)
        return src_n, s, fd

    for label, fields in [('锁定字段(价格/单位/折扣/品牌/地区/税率/条款)', LOCKED),
                          ('名称', {NAME_FIELD}),
                          ('规格原文', SPEC_FIELDS)]:
        src_n, s, rt = rate(fields)
        print(f'{label:42s} 原表有值{src_n:6d} 冲突{s["diff"]:3d} 丢失{s["drop"]:3d} '
              f'纠错{s["typo"]:3d} 追加{s["append"]:3d} 补全{s["fill"]:5d} 保真度 {rt:6.2f}%')

    fatals = [(r['name'], f) for r in reports for f in verdict(r)]
    print('\n致命项: ' + ('无' if not fatals else ''))
    for n, f in fatals:
        print(f'  ✗ {n}: {f}')

    if args.report is None:
        args.report = str(Path(args.deliveries) / '对账_vs压缩包原表.xlsx')
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '差异明细'
    ws.append(['供应商', '行号', '字段', '类型', '原表(他们的结果)', '交付(我们的结果)'])
    for r in reports:
        for d in r['details']:
            ws.append([str(x)[:500] if x is not None else '' for x in d])
    ws2 = wb.create_sheet('按字段统计')
    ws2.append(['供应商', '字段', '一致', '仅格式差异', '错别字纠正', '原文后追加', '我们补全', '我们丢失', '内容冲突'])
    for r in reports:
        for k, v in sorted(r['per_field'].items(), key=lambda x: -(x[1]['diff'] + x[1]['drop'])):
            if v['diff'] or v['drop'] or v['fmt'] or v['typo'] or v['append']:
                ws2.append([r['name'], k[0], v['same'], v['fmt'], v['typo'], v['append'], v['fill'], v['drop'], v['diff']])
    os.makedirs(os.path.dirname(os.path.abspath(args.report)), exist_ok=True)
    wb.save(args.report)
    print(f'\n明细: {args.report}')
    return 1 if fatals else 0


if __name__ == '__main__':
    raise SystemExit(main())
