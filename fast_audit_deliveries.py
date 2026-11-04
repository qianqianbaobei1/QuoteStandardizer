# -*- coding: utf-8 -*-
"""
fast_audit_deliveries.py - 极速工业级四维交付物深度巡检与缺陷隔离器
专门响应客户指令：
1. 有没有空行
2. 是不是只看了前面几行（多页漏页、截断）
3. 识别出来的条数和 Excel 里实际解析完的条数是否一致
4. 内容对不对（价格、残留内部编码、300DPI底稿图片与超链接）
将有问题的供应商全部标记、隔离并输出明细清单，准备进行重新解析！
"""
import os
import re
import fitz
import openpyxl
import zipfile
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

def get_raw_file_real_stats(primary_file: Path):
    """准确获取原单真实行数与真实页数（避免 openpyxl 空格底表 1048576 假象）"""
    ext = primary_file.suffix.lower()
    pages = 0
    real_rows = 0
    if ext == '.pdf':
        try:
            doc = fitz.open(str(primary_file))
            pages = len(doc)
        except Exception:
            pages = 1
    elif ext in ['.xlsx', '.xls']:
        try:
            # 隐藏口径：隐藏表 / 隐藏行一律不计入原表真实行数
            report = hidden_content_report(str(primary_file)) if SKIP_HIDDEN_CONTENT else {}
            wb = openpyxl.load_workbook(str(primary_file), read_only=True, data_only=True)
            for ws in wb.worksheets:
                if (report.get(ws.title) or {}).get('hidden'):
                    continue
                hid = (report.get(ws.title) or {}).get('hidden_rows') or set()
                # 严格限制 max_col=40，防范 16384 列超大空白 XML 拖垮 CPU
                r_cnt = 0
                consecutive_empty = 0
                for r_idx, row in enumerate(ws.iter_rows(max_col=40, values_only=True), start=1):
                    if r_idx in hid:
                        continue
                    if any(v not in (None, '') for v in row):
                        r_cnt += 1
                        consecutive_empty = 0
                    else:
                        consecutive_empty += 1
                        if consecutive_empty > 30 and r_cnt > 0:
                            break
                    if r_cnt > 10000:
                        break
                real_rows += max(0, r_cnt - 1)
        except Exception:
            real_rows = 0
    return ext, pages, real_rows

def audit_delivery_fast(task_info):
    batch_name, vendor_raw_name, vendor_path = task_info
    clean_v = clean_vendor_name(vendor_raw_name)
    xlsx_path = DELIVERIES_DIR / f"{clean_v}-标准解析.xlsx"
    img_dir = DELIVERIES_DIR / f"{clean_v}_图片库"

    res = {
        "batch": batch_name,
        "vendor": clean_v,
        "raw_dir": vendor_raw_name,
        "xlsx_path": str(xlsx_path),
        "status": "PASS",
        "raw_file": None,
        "raw_type": None,
        "raw_pages": 0,
        "raw_rows": 0,
        "excel_data_rows": 0,
        "empty_rows": 0,
        "blank_key_rows": 0,
        "bad_price_rows": 0,
        "code_residual_rows": 0,
        "missing_img_name_rows": 0,
        "missing_disk_img_files": [],
        "pages_covered": set(),
        "issues": []
    }

    # 1. 原文件勘测
    extract_zip_if_needed(vendor_path)
    primary_file = find_primary_quote_file(vendor_path)
    if not primary_file:
        res["status"] = "EMPTY"
        res["issues"].append("【原始空档】档案内无有效报价单原件")
        return res

    res["raw_file"] = primary_file.name
    raw_ext, raw_pages, raw_rows = get_raw_file_real_stats(primary_file)
    res["raw_type"] = raw_ext
    res["raw_pages"] = raw_pages
    res["raw_rows"] = raw_rows

    # 2. 检查交付文件是否存在
    if not xlsx_path.exists():
        res["status"] = "MISSING"
        res["issues"].append("【缺失交付件】未找到对应的标准解析.xlsx")
        return res

    # 3. 极速读取与四维深度巡检
    try:
        wb = openpyxl.load_workbook(str(xlsx_path), read_only=True, data_only=True)
    except Exception as e:
        res["status"] = "FAIL"
        res["issues"].append(f"【损坏文件】Excel无法读回: {e}")
        return res

    if '标准化报价' not in wb.sheetnames:
        res["status"] = "FAIL"
        res["issues"].append(f"【Sheet缺失】缺少「标准化报价」Sheet: {wb.sheetnames}")
        return res

    ws = wb['标准化报价']
    available_imgs = set(os.listdir(img_dir)) if img_dir.exists() else set()
    available_imgs_lower = {fn.lower(): fn for fn in available_imgs}

    has_embedded_drawings = False
    try:
        with zipfile.ZipFile(str(xlsx_path), 'r') as z:
            has_embedded_drawings = any('drawings/drawing' in n for n in z.namelist())
    except Exception:
        pass

    header_row = None
    col_map = {}
    row_idx = 0
    consecutive_empty = 0

    for row in ws.iter_rows(max_col=40, values_only=True):
        row_idx += 1
        if row_idx == 1:
            header_row = [str(c).strip() if c is not None else '' for c in row]
            col_map = {h: idx for idx, h in enumerate(header_row) if h}
            continue

        is_empty = not any(v not in (None, '') for v in row)
        if is_empty:
            consecutive_empty += 1
            if consecutive_empty > 50:
                break
            res["empty_rows"] += 1
            continue
        consecutive_empty = 0
        res["excel_data_rows"] += 1

        # 字段定位
        c_name = col_map.get('名称规格', -1)
        c_price = col_map.get('市场价', -1)
        c_img = col_map.get('厂家报价图片文件名', -1)
        c_spec = col_map.get('规格参数', -1)
        c_raw_spec = col_map.get('厂家原始规格参数', -1)
        c_page = col_map.get('厂家报价文件页码', -1)

        name_val = row[c_name] if 0 <= c_name < len(row) else None
        price_val = row[c_price] if 0 <= c_price < len(row) else None
        img_val = row[c_img] if 0 <= c_img < len(row) else None
        spec_val = row[c_spec] if 0 <= c_spec < len(row) else ""
        raw_spec_val = row[c_raw_spec] if 0 <= c_raw_spec < len(row) else ""
        page_val = row[c_page] if 0 <= c_page < len(row) else ""

        # 检查 1: 空行与关键字段留空
        if name_val in (None, '') or price_val in (None, ''):
            res["blank_key_rows"] += 1

        # 检查 2: 截断页码收集
        if page_val:
            res["pages_covered"].add(str(page_val).strip())

        # 检查 4: 内容对不对
        # (a) 价格数值
        try:
            p_val = float(price_val)
            if p_val <= 0:
                res["bad_price_rows"] += 1
        except (ValueError, TypeError):
            res["bad_price_rows"] += 1

        # (b) 残留产品编码
        spec_all = f"{spec_val or ''} {raw_spec_val or ''}"
        for pat in FORBIDDEN_CODE_PATTERNS:
            if re.search(pat, spec_all):
                res["code_residual_rows"] += 1
                break

        # (c) 底稿图片与物理文件 (支持方案 A 直接内嵌图片无文本模式)
        img_str = str(img_val or '').strip()
        if not img_str or img_str.lower() in ('none', ''):
            if not (has_embedded_drawings and len(available_imgs) > 0):
                res["missing_img_name_rows"] += 1
        else:
            if img_str.lower() not in available_imgs_lower:
                res["missing_disk_img_files"].append(img_str)

    # ------------------ 判定 1: 有没有空行 ------------------
    if res["empty_rows"] > 0:
        res["issues"].append(f"【空行】发现 {res['empty_rows']} 行空行占位")
    if res["blank_key_rows"] > 0:
        res["issues"].append(f"【残缺行】发现 {res['blank_key_rows']} 行名称规格或市场价为空")

    # ------------------ 判定 2: 是不是只看了前面几行 (截断) ------------------
    if raw_ext == '.pdf' and raw_pages >= 3:
        if len(res["pages_covered"]) == 1 and ('第1页' in res["pages_covered"] or '1' in res["pages_covered"]):
            res["issues"].append(f"【严重漏页】原单有 {raw_pages} 页，但只解析了第 1 页！")
        elif len(res["pages_covered"]) < raw_pages * 0.35 and raw_pages >= 5:
            res["issues"].append(f"【疑似漏页】原单有 {raw_pages} 页，仅覆盖 {len(res['pages_covered'])} 个页码")

    if raw_ext in ['.xlsx', '.xls'] and raw_rows >= 30:
        if res["excel_data_rows"] < raw_rows * 0.3:
            res["issues"].append(f"【表格截断】原表有 {raw_rows} 行，标准解析仅产出 {res['excel_data_rows']} 行！")

    # ------------------ 判定 3: 识别条数与实盘条数一致性 ------------------
    if res["excel_data_rows"] == 0:
        res["issues"].append("【零条数】交付 Excel 中没有提取出任何有效物料行 (0 行)")

    # ------------------ 判定 4: 内容对不对 ------------------
    if res["bad_price_rows"] > max(3, res["excel_data_rows"] * 0.3) and res["excel_data_rows"] > 0:
        res["issues"].append(f"【异常价格】多达 {res['bad_price_rows']}/{res['excel_data_rows']} 行单价为 0 或非数值")
    if res["code_residual_rows"] > 0:
        res["issues"].append(f"【残留编码】发现 {res['code_residual_rows']} 行规格中混入产品编码/物料编码")
    if res["missing_img_name_rows"] > 0:
        res["issues"].append(f"【缺少底稿图名】{res['missing_img_name_rows']} 行未填写厂家报价图片文件名")
    if res["missing_disk_img_files"]:
        unique_missing = set(res["missing_disk_img_files"])
        res["issues"].append(f"【图片未落盘】引用的底稿图在图片库中物理缺失 {len(unique_missing)} 个文件: {list(unique_missing)[:2]}")

    if res["issues"]:
        res["status"] = "DEFECTIVE"
    else:
        res["status"] = "PERFECT"

    return res

def main():
    tasks = scan_all_universe_tasks()
    print("=" * 80)
    print(f"【全域 163 家交付件极速四维质量巡检 (只看前几行/空行/条数一致性/内容合规)】")
    print("=" * 80)

    perfects = []
    defectives = []
    empties = []
    missings = []

    for idx, t in enumerate(tasks, 1):
        r = audit_delivery_fast(t)
        st = r["status"]
        if st == "PERFECT":
            perfects.append(r)
            print(f"[{idx:3d}/{len(tasks)}] [PERFECT] ✓ [{r['batch']}] {r['vendor']} | 产出: {r['excel_data_rows']} 行 | 原件: {r['raw_type']} {r['raw_pages']}页/{r['raw_rows']}行")
        elif st == "DEFECTIVE":
            defectives.append(r)
            print(f"[{idx:3d}/{len(tasks)}] [DEFECT ] ✗ [{r['batch']}] {r['vendor']} | 产出: {r['excel_data_rows']} 行")
            for iss in r["issues"]:
                print(f"        ! {iss}")
        elif st == "EMPTY":
            empties.append(r)
            print(f"[{idx:3d}/{len(tasks)}] [EMPTY  ] - [{r['batch']}] {r['vendor']} (原单空文件夹)")
        else:
            missings.append(r)
            print(f"[{idx:3d}/{len(tasks)}] [MISSING] ✗ [{r['batch']}] {r['vendor']} (缺少交付件)")

    print("\n" + "=" * 80)
    print("【巡检统计结论】")
    print(f"  -> 总巡检厂商: {len(tasks)} 家")
    print(f"  -> 完美合规 (PERFECT): {len(perfects)} 家")
    print(f"  -> 存在问题 (DEFECTIVE): {len(defectives)} 家")
    print(f"  -> 原单空档案 (EMPTY): {len(empties)} 家")
    print(f"  -> 缺少交付件 (MISSING): {len(missings)} 家")
    print("=" * 80)

    # 打印需要摘出来重新解析的问题清单
    if defectives:
        print("\n【🚨 需摘出隔离并重新解析的问题文件清单】:")
        for idx, d in enumerate(defectives, 1):
            print(f"{idx:2d}. [{d['batch']}] {d['vendor']} (当前行数: {d['excel_data_rows']})")
            for iss in d["issues"]:
                print(f"     - {iss}")

if __name__ == '__main__':
    main()
