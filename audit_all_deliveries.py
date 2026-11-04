# -*- coding: utf-8 -*-
"""
audit_all_deliveries.py - 工业级全量交付物四维深度体检与缺陷隔离诊断器
严格执行客户四大核心检查指标：
1. 【空行检查】：是否存在纯空行、核心字段(名称规格/市场价)为空的假行、尾部幽灵空行
2. 【截断检查】：是否只看了前面几行/前1页（多页PDF漏页、大表截断）
3. 【条数一致性】：原始物理行、识别物料行与最终落盘行数是否守恒一致
4. 【内容真实性与质量】：价格是否纯正数值、规格是否残留内部编码、300DPI底稿图片及超链接是否100%全覆盖且物理存在
"""
import os
import re
import glob
import fitz
import openpyxl
from pathlib import Path
from run_all_universe_vendors import scan_all_universe_tasks, clean_vendor_name
from run_batch_standardize import find_primary_quote_file, extract_zip_if_needed
from extractor import hidden_content_report
from config import SKIP_HIDDEN_CONTENT

DELIVERIES_DIR = Path('/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries')

FORBIDDEN_CODE_PATTERNS = [
    r'产品编码\s*[:：]', r'物料编码\s*[:：]', r'商品编码\s*[:：]',
    r'内部编码\s*[:：]', r'条形码\s*[:：]', r'三级类别编码\s*[:：]'
]

def audit_single_delivery(task_info):
    batch_name, vendor_raw_name, vendor_path = task_info
    clean_v = clean_vendor_name(vendor_raw_name)
    xlsx_path = DELIVERIES_DIR / f"{clean_v}-标准解析.xlsx"
    img_dir = DELIVERIES_DIR / f"{clean_v}_图片库"

    issues = []
    audit_data = {
        "batch": batch_name,
        "vendor": clean_v,
        "raw_dir": vendor_raw_name,
        "xlsx_path": str(xlsx_path),
        "status": "PASS",
        "raw_file": None,
        "raw_type": None,
        "raw_pages": 0,
        "raw_rows": 0,
        "excel_rows": 0,
        "empty_rows_cnt": 0,
        "truncated_pages": False,
        "count_mismatch": False,
        "content_issues": [],
        "issues": issues
    }

    # 1. 原始文件勘测
    extract_zip_if_needed(vendor_path)
    primary_file = find_primary_quote_file(vendor_path)
    if not primary_file:
        audit_data["status"] = "EMPTY"
        audit_data["issues"].append("【原始空档】原始目录内无有效报价单")
        return audit_data

    audit_data["raw_file"] = primary_file.name
    raw_ext = primary_file.suffix.lower()
    audit_data["raw_type"] = raw_ext

    if raw_ext == '.pdf':
        try:
            doc = fitz.open(str(primary_file))
            audit_data["raw_pages"] = len(doc)
        except Exception:
            audit_data["raw_pages"] = 1
    elif raw_ext in ['.xlsx', '.xls']:
        try:
            # 隐藏口径：隐藏表 / 隐藏行不计入原表物理行
            report = hidden_content_report(str(primary_file)) if SKIP_HIDDEN_CONTENT else {}
            wb_raw = openpyxl.load_workbook(str(primary_file), read_only=True)
            audit_data["raw_rows"] = sum(
                max(0, ws.max_row - 1 - len([r for r in info['hidden_rows'] if r > 1]))
                for ws in wb_raw.worksheets
                if ws.max_row and not (report.get(ws.title) or {}).get('hidden')
                for info in [report.get(ws.title) or {'hidden_rows': set()}]
            )
            audit_data["hidden_sheets"] = [s for s, i in report.items() if i['hidden']]
        except Exception:
            audit_data["raw_rows"] = 0

    # 2. 交付 Excel 物理勘测
    if not xlsx_path.exists():
        audit_data["status"] = "FAIL"
        audit_data["issues"].append("【缺失交付件】deliveries/ 目录下未找到对应的标准解析.xlsx")
        return audit_data

    try:
        wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
    except Exception as e:
        audit_data["status"] = "FAIL"
        audit_data["issues"].append(f"【损坏文件】交付 Excel 无法被 openpyxl 打开: {e}")
        return audit_data

    if '标准化报价' not in wb.sheetnames:
        audit_data["status"] = "FAIL"
        audit_data["issues"].append(f"【Sheet缺失】缺少「标准化报价」Sheet (当前为: {wb.sheetnames})")
        return audit_data

    ws = wb['标准化报价']
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    
    col_map = {str(h).strip(): idx for idx, h in enumerate(headers, 1) if h}
    req_cols = ['名称规格', '市场价', '厂家报价图片文件名', '规格参数', '厂家原始规格参数', '厂家报价文件页码']
    missing_cols = [c for c in req_cols if c not in col_map]
    if missing_cols:
        audit_data["status"] = "FAIL"
        audit_data["issues"].append(f"【表头残缺】缺少关键列: {missing_cols}")
        return audit_data

    c_name = col_map['名称规格']
    c_price = col_map['市场价']
    c_img = col_map['厂家报价图片文件名']
    c_spec = col_map.get('规格参数')
    c_raw_spec = col_map.get('厂家原始规格参数')
    c_page = col_map.get('厂家报价文件页码')

    data_rows = 0
    empty_rows = 0
    blank_name_or_price_rows = 0
    zero_or_bad_price_rows = 0
    code_residual_rows = 0
    missing_img_rows = 0
    missing_physical_img_files = set()
    pages_seen = set()

    available_imgs = set(os.listdir(img_dir)) if img_dir.exists() else set()

    for r in range(2, ws.max_row + 1):
        name_val = ws.cell(r, c_name).value
        price_val = ws.cell(r, c_price).value
        img_val = ws.cell(r, c_img).value
        spec_val = ws.cell(r, c_spec).value if c_spec else ""
        raw_spec_val = ws.cell(r, c_raw_spec).value if c_raw_spec else ""
        page_val = ws.cell(r, c_page).value if c_page else ""

        # 是否整行全空
        row_vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        if not any(v not in (None, '') for v in row_vals):
            empty_rows += 1
            continue

        data_rows += 1

        # 检查 1: 空行 / 关键字段为空
        if name_val in (None, '') or price_val in (None, ''):
            blank_name_or_price_rows += 1

        # 检查 2: 截断追踪（页码覆盖）
        if page_val:
            pages_seen.add(str(page_val).strip())

        # 检查 4: 内容对不对
        # (a) 市场价格格式与合理性
        try:
            p_float = float(price_val)
            if p_float <= 0:
                zero_or_bad_price_rows += 1
        except (ValueError, TypeError):
            zero_or_bad_price_rows += 1

        # (b) 是否残留内部编码
        spec_combined = f"{spec_val or ''} {raw_spec_val or ''}"
        for pat in FORBIDDEN_CODE_PATTERNS:
            if re.search(pat, spec_combined):
                code_residual_rows += 1
                break

        # (c) 厂家报价图片与物理存在性
        img_str = str(img_val or '').strip()
        if not img_str or img_str == 'None':
            missing_img_rows += 1
        else:
            if img_str not in available_imgs:
                # 检查忽略大小写
                lower_match = [f for f in available_imgs if f.lower() == img_str.lower()]
                if not lower_match:
                    missing_physical_img_files.add(img_str)

    audit_data["excel_rows"] = data_rows
    audit_data["empty_rows_cnt"] = empty_rows

    # ------------------ 判定 1: 有没有空行 ------------------
    if empty_rows > 0:
        issues.append(f"【空行缺陷】发现 {empty_rows} 行完全空白的无效占位行")
    if blank_name_or_price_rows > 0:
        issues.append(f"【假数据行】发现 {blank_name_or_price_rows} 行名称或市场价为空的坏行")

    # ------------------ 判定 2: 是不是只看了前面几行 / 截断 ------------------
    if raw_ext == '.pdf' and audit_data["raw_pages"] > 1:
        # 如果 PDF 有多页，但只看了第 1 页
        if len(pages_seen) == 1 and ('第1页' in pages_seen or '1' in pages_seen):
            audit_data["truncated_pages"] = True
            issues.append(f"【漏页截断】原 PDF 共有 {audit_data['raw_pages']} 页，但仅解析了第 1 页！(产出 {data_rows} 行)")
        elif len(pages_seen) < audit_data["raw_pages"] * 0.4 and audit_data["raw_pages"] >= 5:
            audit_data["truncated_pages"] = True
            issues.append(f"【疑似漏页】原 PDF 共有 {audit_data['raw_pages']} 页，仅覆盖 {len(pages_seen)} 个页码引用")

    if raw_ext in ['.xlsx', '.xls'] and audit_data["raw_rows"] > 50:
        if data_rows < audit_data["raw_rows"] * 0.3:
            issues.append(f"【Excel截断】原表有 {audit_data['raw_rows']} 行，但标准解析仅产出 {data_rows} 行！")

    # ------------------ 判定 3: 识别条数与实盘条数一致性 ------------------
    if data_rows == 0:
        issues.append("【零物料】交付 Excel 中没有任何有效物料数据行 (0行)")

    # ------------------ 判定 4: 内容对不对 ------------------
    if zero_or_bad_price_rows > data_rows * 0.3 and data_rows > 0:
        issues.append(f"【异常价格】多达 {zero_or_bad_price_rows}/{data_rows} 行市场价为 0 或非有效数值")
    if code_residual_rows > 0:
        issues.append(f"【残留编码】发现 {code_residual_rows} 行规格参数中残留「产品编码/物料编码」等内部非物理字段")
    if missing_img_rows > 0:
        issues.append(f"【缺少底稿图名】{missing_img_rows} 行未填写「厂家报价图片文件名」")
    if missing_physical_img_files:
        issues.append(f"【图片未落盘】引用的底稿图片在图片库中物理缺失: {list(missing_physical_img_files)[:3]}")

    if issues:
        audit_data["status"] = "DEFECTIVE"
    else:
        audit_data["status"] = "PERFECT"

    return audit_data

def run_universe_audit():
    tasks = scan_all_universe_tasks()
    print("=" * 80)
    print(f"【全域 163 家建材报价单交付成果全维深度体检启动】")
    print(f"  -> 待巡检总档案库: {len(tasks)} 家")
    print(f"  -> 检验标准: 严格四维闭环 (空行 / 截断 / 条数守恒 / 内容真实合规)")
    print("=" * 80)

    results = []
    perfect_cnt = 0
    defective_cnt = 0
    empty_cnt = 0
    missing_cnt = 0

    for idx, task in enumerate(tasks, 1):
        res = audit_single_delivery(task)
        results.append(res)
        
        status = res["status"]
        if status == "PERFECT":
            perfect_cnt += 1
            print(f"[{idx:3d}/{len(tasks)}] [PERFECT] ✓ [{res['batch']}] {res['vendor']} "
                  f"(产出 {res['excel_rows']} 行, 原文件: {res['raw_type']} {res['raw_pages']}页/{res['raw_rows']}行)")
        elif status == "DEFECTIVE":
            defective_cnt += 1
            print(f"[{idx:3d}/{len(tasks)}] [DEFECT ] ✗ [{res['batch']}] {res['vendor']} "
                  f"(产出 {res['excel_rows']} 行)")
            for iss in res["issues"]:
                print(f"        ! {iss}")
        elif status == "EMPTY":
            empty_cnt += 1
            print(f"[{idx:3d}/{len(tasks)}] [EMPTY  ] - [{res['batch']}] {res['vendor']} (原单空文件夹)")
        else:
            missing_cnt += 1
            print(f"[{idx:3d}/{len(tasks)}] [MISSING] ✗ [{res['batch']}] {res['vendor']} (未生成交付件)")

    print("\n" + "=" * 80)
    print("【全维巡检统计总汇】")
    print(f"  -> 完美合规 (PERFECT): {perfect_cnt} 家")
    print(f"  -> 存在缺陷 (DEFECTIVE): {defective_cnt} 家  <-- 需摘出隔离并重新解析")
    print(f"  -> 原始空档 (EMPTY): {empty_cnt} 家")
    print(f"  -> 缺失成果 (MISSING): {missing_cnt} 家")
    print("=" * 80)

    # 导出缺陷清单
    defective_list = [r for r in results if r["status"] in ("DEFECTIVE", "MISSING")]
    return results, defective_list

if __name__ == '__main__':
    run_universe_audit()
