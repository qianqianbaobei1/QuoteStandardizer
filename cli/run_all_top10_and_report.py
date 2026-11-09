# -*- coding: utf-8 -*-
"""
cli/run_all_top10_and_report.py
全量自跑 Top10 供应商报价单，逐个生成标准交付物、对账底稿与质量看板，并输出全量准确率汇总大盘
"""
import os
import sys
import re
import time
import zipfile
import json
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from cli.run_batch_standardize import run_single_vendor
from config import DEFAULT_TEMPLATE_PATH

DESKTOP_TOP10 = Path("/Users/qianqianawodebaobei/Desktop/top10_download_20261106")
DELIVERIES_ROOT = ROOT / "deliveries"

def extract_theirs_xlsx(vendor_dir: Path) -> Path:
    """从供应商导出结果 zip 中解压出 报价数据.xlsx"""
    zdir = vendor_dir / "导出结果"
    if not zdir.exists():
        return None
    for zf in zdir.glob("*.zip"):
        try:
            with zipfile.ZipFile(zf, 'r') as z:
                for n in z.namelist():
                    if n.endswith("报价数据.xlsx"):
                        out_p = vendor_dir / "线上结果_解压"
                        out_p.mkdir(parents=True, exist_ok=True)
                        target = out_p / "报价数据.xlsx"
                        with z.open(n) as s, open(target, 'wb') as d:
                            d.write(s.read())
                        return target
        except Exception as e:
            print(f"[Warning] 解压线上结果失败 {zf.name}: {e}")
    return None

def load_items_from_xlsx(path: Path):
    """读取我方或线上 Excel 物料数据"""
    if not path or not path.exists():
        return []
    wb = openpyxl.load_workbook(path, data_only=True)
    items = []
    # 如果有 标准化报价 Sheet
    if "标准化报价" in wb.sheetnames:
        ws = wb["标准化报价"]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) > 1:
            hm = {name: i for i, name in enumerate(rows[0])}
            for r in rows[1:]:
                items.append({
                    'name': str(r[hm.get('名称规格', 2)] or '').strip(),
                    'cat_code': str(r[hm.get('类别编码', 3)] or '').strip(),
                    'brand': str(r[hm.get('品牌', 4)] or '').strip(),
                    'raw_spec': str(r[hm.get('厂家原始规格参数', 6)] or '').strip(),
                    'page': str(r[hm.get('厂家报价文件页码', 7)] or '').strip(),
                    'unit': str(r[hm.get('单位', 12)] or '').strip(),
                    'price': r[hm.get('市场价', 13)],
                    'tax_inc': r[hm.get('是否含增值税', 24)],
                    'tax_rate': r[hm.get('增值税税率', 22)],
                })
    else:
        # 多 Sheet 分类形式 (线上系统)
        for sn in wb.sheetnames:
            ws = wb[sn]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) < 2: continue
            for r in rows[1:]:
                items.append({
                    'name': str(r[1] or '').strip(),
                    'cat_code': sn,
                    'brand': str(r[2] or '').strip(),
                    'raw_spec': str(r[3] or '').strip(),
                    'page': r[4] if len(r)>4 else '',
                    'unit': str(r[9] or '').strip() if len(r)>9 else '',
                    'price': r[10] if len(r)>10 else None,
                    'tax_inc': r[18] if len(r)>18 else None,
                    'tax_rate': r[20] if len(r)>20 else None,
                })
    return items

def generate_vendor_review(vendor_name: str, our_xlsx: Path, their_xlsx: Path, out_dir: Path):
    """为单家供应商生成双向对账底稿 Excel 与 D3 看板 HTML"""
    ours = load_items_from_xlsx(our_xlsx)
    theirs = load_items_from_xlsx(their_xlsx)
    
    # 统计指标
    our_cnt = len(ours)
    their_cnt = len(theirs)
    
    # 单价匹配统计
    price_match = 0
    cat_match = 0
    typo_cnt = 0
    tax_err_cnt = 0
    unknown_cnt = sum(1 for t in theirs if t['cat_code'] == '未知' or not t['name'])
    
    for t in theirs:
        spec_text = f"{t['name']} {t['raw_spec']}"
        if any(w in spec_text for w in ["然轧", "工字钠", "王字钢", "方草厚", "方干", "壁焊", "市厚", "品对"]):
            typo_cnt += 1
        if t.get('tax_inc') == '不含':
            tax_err_cnt += 1
            
    # 单价吻合度
    used_t = set()
    for o in ours:
        for t_idx, t in enumerate(theirs):
            if t_idx in used_t: continue
            if o['price'] is not None and t['price'] is not None and str(o['price']) == str(t['price']):
                price_match += 1
                used_t.add(t_idx)
                if o['cat_code'].replace('-', '') == t['cat_code'].replace('-', ''):
                    cat_match += 1
                break
                
    price_acc = round(price_match / max(1, our_cnt) * 100, 1)
    cat_acc = round(cat_match / max(1, our_cnt) * 100, 1)
    
    # 生成 Excel 对账底稿
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "对账总览仪表盘"
    ws2 = wb.create_sheet(title="逐行明细三方核对表")
    
    f_title = Font(name="Microsoft YaHei", size=14, bold=True, color="1F4E78")
    f_sec = Font(name="Microsoft YaHei", size=11, bold=True, color="002060")
    f_hdr = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
    f_cell = Font(name="Microsoft YaHei", size=9)
    f_bold = Font(name="Microsoft YaHei", size=9, bold=True)
    f_kpi = Font(name="Microsoft YaHei", size=16, bold=True, color="1F4E78")
    
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_match = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_diff = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_fix = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    
    ws1.merge_cells("A1:G1")
    ws1["A1"] = f"【标准解析 vs 线上对比】{vendor_name} 深度对账底稿"
    ws1["A1"].font = f_title
    
    kpis = [
        ("B3:C4", "我方解析物料", f"{our_cnt} 行"),
        ("D3:E4", "线上提取物料", f"{their_cnt} 行"),
        ("F3:G4", "单价吻合率", f"{price_acc}%"),
        ("H3:I4", "线上质量缺陷", f"{tax_err_cnt}税错/{typo_cnt}错字"),
    ]
    for rng, lbl, val in kpis:
        ws1.merge_cells(rng)
        c = ws1[rng.split(":")[0]]
        c.value = f"{val}\n{lbl}"
        c.font = f_kpi
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border_thin
        
    # 逐行明细
    headers_ws2 = ["序号", "【我方】品名", "【我方】分类", "【我方】单价", "【我方】规格参数",
                   "【线上】品名", "【线上】分类", "【线上】单价", "【线上】规格参数", "裁决结论"]
    for ci, h in enumerate(headers_ws2, 1):
        c = ws2.cell(1, ci, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal="center", vertical="center")
        
    for idx, o in enumerate(ours, 1):
        t = theirs[idx - 1] if idx - 1 < len(theirs) else {}
        verdict = "✅ 完全一致" if str(o.get('price')) == str(t.get('price')) else "⚠️ 存在差异"
        vals = [idx, o.get('name'), o.get('cat_code'), o.get('price'), o.get('raw_spec'),
                t.get('name', '-'), t.get('cat_code', '-'), t.get('price', '-'), t.get('raw_spec', '-'), verdict]
        for ci, v in enumerate(vals, 1):
            c = ws2.cell(idx + 1, ci, v)
            c.font = f_cell
            c.border = border_thin
            if ci in (1, 3, 4, 7, 8, 10):
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
            if ci == 10:
                c.fill = fill_match if "一致" in verdict else fill_diff
                c.font = f_bold
                
    review_xlsx_path = out_dir / f"{vendor_name}_双向对账核对底稿.xlsx"
    wb.save(review_xlsx_path)
    
    return {
        'vendor': vendor_name,
        'our_cnt': our_cnt,
        'their_cnt': their_cnt,
        'price_acc': price_acc,
        'cat_acc': cat_acc,
        'typo_cnt': typo_cnt,
        'tax_err_cnt': tax_err_cnt,
        'unknown_cnt': unknown_cnt,
        'review_xlsx': review_xlsx_path,
        'out_dir': out_dir
    }

def main():
    print("=== [Top10 全量自跑与评测汇总引擎启动] ===")
    dirs = sorted([d for d in DESKTOP_TOP10.iterdir() if d.is_dir() and not d.name.startswith(".")])
    print(f"检测到 {len(dirs)} 家供应商目录，启动逐一执行...")
    
    summary_stats = []
    
    for idx, vd in enumerate(dirs, 1):
        vname = vd.name
        print(f"\n==================================================")
        print(f"[{idx}/{len(dirs)}] 正在标准化解析: {vname}")
        print(f"==================================================")
        
        target_out = DELIVERIES_ROOT / vname
        target_out.mkdir(parents=True, exist_ok=True)
        
        # 1. 执行我方标准解析
        res = run_single_vendor(vd, target_out)
        our_xlsx = Path(res["output_file"]) if res.get("output_file") else None
        if not our_xlsx or not our_xlsx.exists():
            cand = [f for f in target_out.glob("*标准解析.xlsx") if not f.name.startswith("~$")]
            if cand:
                our_xlsx = cand[0]
        
        # 2. 解压线上系统报价数据
        their_xlsx = extract_theirs_xlsx(vd)
        
        # 3. 生成双向对账核对底稿
        rev_info = generate_vendor_review(vname, our_xlsx, their_xlsx, target_out)
        rev_info['status'] = res['status']
        rev_info['duration'] = res['duration']
        rev_info['failures'] = len(res['failures'] or [])
        summary_stats.append(rev_info)
        
        print(f"  -> 完成: 我方 {rev_info['our_cnt']} 行 vs 线上 {rev_info['their_cnt']} 行 | 单价吻合: {rev_info['price_acc']}% | 耗时: {res['duration']}s")
        
    # 生成全量大盘汇总
    build_top10_master_summary(summary_stats)

def build_top10_master_summary(stats):
    """生成包含整体解析准确率的 Top10 全量汇总大盘 (Excel + HTML)"""
    # 汇总计算
    total_our = sum(s['our_cnt'] for s in stats)
    total_their = sum(s['their_cnt'] for s in stats)
    avg_price_acc = round(sum(s['price_acc'] for s in stats) / max(1, len(stats)), 1)
    avg_cat_acc = round(sum(s['cat_acc'] for s in stats) / max(1, len(stats)), 1)
    total_typo = sum(s['typo_cnt'] for s in stats)
    total_tax_err = sum(s['tax_err_cnt'] for s in stats)
    total_unknown = sum(s['unknown_cnt'] for s in stats)
    
    master_xlsx = DELIVERIES_ROOT / "Top10_报价单整体解析质量汇总底稿.xlsx"
    master_html = DELIVERIES_ROOT / "Top10_报价单全量解析评测汇总看板.html"
    
    # 1. Master Excel
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Top10全量质量总览"
    
    f_title = Font(name="Microsoft YaHei", size=14, bold=True, color="1F4E78")
    f_hdr = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")
    f_cell = Font(name="Microsoft YaHei", size=9)
    f_bold = Font(name="Microsoft YaHei", size=9, bold=True)
    f_kpi = Font(name="Microsoft YaHei", size=16, bold=True, color="1F4E78")
    
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_match = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    border_thin = Border(left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
                         top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9'))
    
    ws.merge_cells("A1:I1")
    ws["A1"] = "【Top10 供应商报价单】整体解析准确率与质量评测汇总总览"
    ws["A1"].font = f_title
    
    # KPI Cards
    kpis = [
        ("B3:C4", "总体解析物料量", f"{total_our} 行"),
        ("D3:E4", "整体单价吻合率", f"{avg_price_acc}%"),
        ("F3:G4", "标准分类匹配度", f"{avg_cat_acc}%"),
        ("H3:I4", "纠正线上缺陷数", f"{total_tax_err}税错/{total_typo}错字"),
    ]
    for rng, lbl, val in kpis:
        ws.merge_cells(rng)
        c = ws[rng.split(":")[0]]
        c.value = f"{val}\n{lbl}"
        c.font = f_kpi
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border_thin
        
    headers = ["序号", "供应商名称", "我方解析行", "线上解析行", "单价吻合率", "分类吻合率", "线上错别字", "线上税判错", "线上未知行", "四维体检结果"]
    for ci, h in enumerate(headers, 1):
        c = ws.cell(6, ci, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = Alignment(horizontal="center", vertical="center")
        
    for ri, s in enumerate(stats, 7):
        vals = [
            ri - 6, s['vendor'], s['our_cnt'], s['their_cnt'],
            f"{s['price_acc']}%", f"{s['cat_acc']}%", s['typo_cnt'], s['tax_err_cnt'], s['unknown_cnt'],
            "PASS" if s.get('status') == 'PASS' else "需复核"
        ]
        for ci, v in enumerate(vals, 1):
            c = ws.cell(ri, ci, v)
            c.font = f_cell
            c.border = border_thin
            c.alignment = Alignment(horizontal="center" if ci != 2 else "left", vertical="center")
            if ci == 5:
                c.fill = fill_match
                c.font = f_bold
                
    wb.save(master_xlsx)
    print(f"\n[OK] Top10 全量汇总底稿已生成: {master_xlsx}")
    
    # 2. Master HTML 看板
    rows_html = "".join([f"""
        <tr>
            <td>{i+1}</td>
            <td><strong>{s['vendor']}</strong></td>
            <td>{s['our_cnt']}</td>
            <td>{s['their_cnt']}</td>
            <td><span class="tag-match">{s['price_acc']}%</span></td>
            <td>{s['cat_acc']}%</td>
            <td>{s['typo_cnt']} 处</td>
            <td>{s['tax_err_cnt']} 行</td>
            <td>{s['unknown_cnt']} 行</td>
            <td><span class="tag-match">{"100% PASS" if s.get('status') == 'PASS' else "需复核"}</span></td>
            <td>
                <a href="file://{s['out_dir']}/报价单解析质量看板.html" target="_blank" style="color:#002FA7;font-weight:600;text-decoration:none;">质量看板 ↗</a>
                <a href="file://{s['review_xlsx']}" target="_blank" style="color:#555;margin-left:8px;text-decoration:none;">对账底稿 ↗</a>
            </td>
        </tr>
    """ for i, s in enumerate(stats)])
    
    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Top10 报价单全量解析评测汇总看板</title>
    <style>
        :root {{ --bg: #FAFAF8; --ink: #0A0A0A; --accent: #002FA7; --card: #FFFFFF; --border: #E5E5E0; --green: #2E7D32; }}
        body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: var(--bg); color: var(--ink); padding: 24px; }}
        .container {{ max-width: 1400px; margin: 0 auto; }}
        .header {{ margin-bottom: 24px; border-bottom: 1px solid var(--border); padding-bottom: 16px; }}
        .header h1 {{ margin: 0 0 8px 0; font-size: 26px; }}
        .kpi-grid {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; margin-bottom: 24px; }}
        .kpi-card {{ background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 20px; }}
        .kpi-card .num {{ font-size: 32px; font-weight: 700; font-family: monospace; color: var(--accent); }}
        .kpi-card .label {{ font-size: 13px; color: #666; margin-top: 6px; }}
        .section-card {{ background: var(--card); border: 1px solid var(--border); border-radius: 8px; padding: 20px; }}
        table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
        th {{ background: #F4F4F2; text-align: left; padding: 10px 12px; border-bottom: 1px solid var(--border); }}
        td {{ padding: 10px 12px; border-bottom: 1px solid var(--border); }}
        .tag-match {{ background: #E8F5E9; color: var(--green); padding: 2px 6px; border-radius: 4px; font-weight: 600; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Top10 报价单全量标准化解析质量评测总览</h1>
            <div class="meta">评测样本：全量 10 家供应商原单 | 对比基准：线上解析系统导出数据 | 生成时间：2026-11-06</div>
        </div>
        <div class="kpi-grid">
            <div class="kpi-card"><div class="num">{total_our} 行</div><div class="label">我方总产出物料行</div></div>
            <div class="kpi-card"><div class="num">{avg_price_acc}%</div><div class="label">整体单价综合准确率</div></div>
            <div class="kpi-card"><div class="num">{avg_cat_acc}%</div><div class="label">标准分类匹配吻合率</div></div>
            <div class="kpi-card"><div class="num">{total_tax_err} 行 / {total_typo} 处</div><div class="label">纠偏线上税率与OCR错字</div></div>
        </div>
        <div class="section-card">
            <h2>各供应商详细评测指标横向对照</h2>
            <table>
                <thead>
                    <tr>
                        <th>#</th><th>供应商名称</th><th>我方行数</th><th>线上行数</th><th>单价准确率</th><th>分类吻合率</th><th>线上错字</th><th>线上税错</th><th>线上未知</th><th>体检状态</th><th>报告直达</th>
                    </tr>
                </thead>
                <tbody>
                    {rows_html}
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>"""
    with open(master_html, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"[OK] Top10 全量汇总看板已生成: {master_html}")

if __name__ == "__main__":
    main()
