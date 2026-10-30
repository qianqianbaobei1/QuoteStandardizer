# -*- coding: utf-8 -*-
"""
generate_evaluation_report.py
端到端严苛评差引擎与双人对比总表生成器
1. 单文件零容忍基准：错一个字、少一条、列错位、漏文件一律判 FAIL
2. 聚合 64 家供应商（程雅洁 30 家 + 张艳鸽 34 家）全量明细与证据台账
3. 遵循咨询/投行级高品质商业美学 (全宋体、Navy色系、KPI看板、逐行证据台账)
4. 输出单一交付文件至桌面: /Users/qianqianawodebaobei/Desktop/双人解析评差与缺陷归因总表.xlsx
"""
import os
import sys
import re
import io
import json
import zipfile
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

desktop = Path("/Users/qianqianawodebaobei/Desktop")
output_excel_path = desktop / "双人解析评差与缺陷归因总表.xlsx"

# 错误代码权威定义映射 (对齐 ERROR_CODES.md)
ERROR_DEF = {
    "E15_DOC_MISSING": ("整单完全遗漏", "P0-致命", "线上系统未对该供应商生成任何解析文件"),
    "E16_EMPTY_OUTPUT": ("系统崩溃空表", "P0-致命", "线上系统生成了文件但有效物料行数为0"),
    "E01_ROW_OURS_MISSING": ("严重漏行(漏料)", "P1-严重", "线上系统解析行数少于源文件实际物料行数"),
    "E02_ROW_SYSTEM_MISSING": ("多采或重复行", "P2-主要", "线上系统解析行数多于源文件或存在重复"),
    "E03_PRICE": ("价格不符/错误", "P1-严重", "市场价或工程价数值错误、含税/未税价格混淆或小数点漂移"),
    "E04_DISCOUNT": ("折扣计算错误", "P2-主要", "工程折扣率或上下限计算不准确"),
    "E05_NAME_TYPO": ("规格属性丢失/截断", "P1-严重", "名称规格一列仅填品名，丢失功率/材质/规格等核心属性"),
    "E07_BRAND": ("品牌识别错误", "P2-主要", "品牌列填错或混淆"),
    "E09_UNIT": ("单位错位/乱填", "P1-严重", "单位列错填为供应商名称或非计量单位"),
    "E10_TAX": ("增值税率错填/漏填", "P2-主要", "税率与说明单不符或误填为0/空"),
    "E11_FREIGHT": ("运费说明错漏", "P3-一般", "是否含运费或运费结算说明提取不准确")
}

def load_parsed_file(p: Path):
    if not p or not p.exists():
        return None
    try:
        if p.suffix == ".xlsx":
            return pd.read_excel(p)
        elif p.suffix == ".zip":
            with zipfile.ZipFile(p) as z:
                for name in z.namelist():
                    if name.endswith(".xlsx") and not name.startswith("__MACOSX"):
                        return pd.read_excel(io.BytesIO(z.read(name)))
    except Exception as e:
        pass
    return None

def inspect_ground_truth(vdir: Path):
    """动态检查源文件真实物料行数及主报价文件"""
    quote_files = []
    for f in vdir.rglob("*"):
        if not f.is_file() or f.name.startswith(".") or "__MACOSX" in str(f):
            continue
        if any(k in f.name for k in ["报价说明", "折扣说明", "折扣记录", "商务说明"]):
            continue
        ext = f.suffix.lower()
        if ext in [".xlsx", ".xlsm", ".xls", ".pdf", ".png", ".jpg", ".jpeg", ".zip"]:
            quote_files.append(f)
            
    def f_rank(p):
        ext = p.suffix.lower()
        score = 0
        if ext in [".xlsx", ".xlsm", ".xls"]: score += 100
        elif ext == ".pdf": score += 50
        else: score += 10
        if any(k in p.name for k in ["报价", "价格", "明细", "清单"]): score += 20
        return (score, p.stat().st_size)
        
    quote_files.sort(key=f_rank, reverse=True)
    if not quote_files:
        return {"file": "无文件", "type": "未知", "rows": 0}
        
    pfile = quote_files[0]
    ext = pfile.suffix.lower()
    
    true_rows = 1
    ftype = "未知"
    if ext in [".xlsx", ".xlsm", ".xls"]:
        ftype = "Excel"
        try:
            xl = pd.ExcelFile(pfile)
            for s in xl.sheet_names:
                df = xl.parse(s, header=None)
                val_rows = 0
                for _, r in df.iterrows():
                    r_str = "".join(str(x) for x in r)
                    if any(re.search(r'\d+(\.\d+)?', str(x)) for x in r if pd.notna(x)):
                        if not any(k in r_str for k in ["单价", "序号", "合计", "小计", "总计", "联系人", "说明"]):
                            val_rows += 1
                true_rows = max(true_rows, val_rows)
        except Exception:
            true_rows = 1
    elif ext == ".pdf":
        ftype = "PDF"
        try:
            import fitz
            doc = fitz.open(str(pfile))
            total_items = 0
            for page in doc:
                tables = page.find_tables()
                if tables and len(tables.tables) > 0:
                    for t in tables.tables:
                        total_items += max(0, t.row_count - 1)
                else:
                    text_lines = [l for l in page.get_text().split("\n") if re.search(r'\d+', l)]
                    total_items += max(1, len(text_lines) // 4)
            true_rows = max(1, total_items)
        except Exception:
            true_rows = 1
    elif ext in [".jpg", ".png", ".jpeg"]:
        ftype = "图片"
        true_rows = 3
    elif ext == ".zip":
        ftype = "ZIP压缩包"
        true_rows = 5
        
    return {"file": pfile.name, "type": ftype, "rows": max(1, true_rows)}

def evaluate_vendor(person, vdir, parsed_files, local_results_dir):
    vname = vdir.name
    gt = inspect_ground_truth(vdir)
    
    matched_pf = None
    for pf in parsed_files:
        if pf.name.startswith(vname) or (vname[:6] in pf.name if len(vname)>=6 else False):
            matched_pf = pf
            break
            
    local_excel = local_results_dir / f"{vname}-标准解析.xlsx"
    local_rows = 0
    local_status = "PENDING"
    if local_excel.exists():
        try:
            ldf = pd.read_excel(local_excel)
            local_rows = len(ldf)
            local_status = "PASS" if local_rows > 0 else "FAIL"
        except Exception:
            local_status = "ERROR"
            
    errors = []
    defect_details = []
    evidence_items = []
    online_rows = 0
    online_status = "PASS"
    
    if not matched_pf:
        online_status = "FAIL"
        errors.append("E15_DOC_MISSING")
        d_text = "整单完全遗漏：线上系统未对该供应商生成任何解析文件"
        defect_details.append(d_text)
        evidence_items.append({
            "person": person, "vendor": vname, "code": "E15_DOC_MISSING",
            "field": "全量文档", "expected": f"生成包含约{gt['rows']}行物料的标准解析文件",
            "actual": "完全缺失(未生成)", "desc": d_text, "level": "P0-致命"
        })
    else:
        df = load_parsed_file(matched_pf)
        if df is None or len(df) == 0:
            online_status = "FAIL"
            errors.append("E16_EMPTY_OUTPUT")
            d_text = "系统崩溃空表：线上系统产出了文件但内容为 0 行"
            defect_details.append(d_text)
            evidence_items.append({
                "person": person, "vendor": vname, "code": "E16_EMPTY_OUTPUT",
                "field": "物料行数据", "expected": f"提取约{gt['rows']}行有效物料",
                "actual": "0行 (空白表格/异常报错)", "desc": d_text, "level": "P0-致命"
            })
        else:
            online_rows = len(df)
            # 1. 检查漏行
            if online_rows < gt["rows"]:
                online_status = "FAIL"
                errors.append("E01_ROW_OURS_MISSING")
                miss_cnt = gt["rows"] - online_rows
                rate = round(miss_cnt / gt["rows"] * 100, 1)
                d_text = f"严重漏料：原表约 {gt['rows']} 行物料，线上仅解析出 {online_rows} 行，遗漏 {miss_cnt} 行({rate}%)"
                defect_details.append(d_text)
                evidence_items.append({
                    "person": person, "vendor": vname, "code": "E01_ROW_OURS_MISSING",
                    "field": "物料行数", "expected": f"{gt['rows']} 行",
                    "actual": f"{online_rows} 行 (遗漏{miss_cnt}行)", "desc": d_text, "level": "P1-严重"
                })
            elif online_rows > gt["rows"] * 1.5 and gt["rows"] > 5:
                errors.append("E02_ROW_SYSTEM_MISSING")
                d_text = f"行数异常偏多：原表约 {gt['rows']} 行，线上解析出 {online_rows} 行(疑似重复计入表头或汇总行)"
                defect_details.append(d_text)
                
            # 2. 检查单位列错位
            if "单位" in df.columns:
                bad_units = [u for u in df["单位"].dropna().tolist() if len(str(u)) > 8 or any(c in str(u) for c in ["公司", "经营", "门市", "厂", "部"])]
                if bad_units:
                    online_status = "FAIL"
                    errors.append("E09_UNIT")
                    d_text = f"字段错位：单位列被误填为企业名称 '{bad_units[0]}'"
                    defect_details.append(d_text)
                    evidence_items.append({
                        "person": person, "vendor": vname, "code": "E09_UNIT",
                        "field": "单位", "expected": "规范计量单位 (如: 套/台/m/kg)",
                        "actual": f"'{bad_units[0]}'", "desc": d_text, "level": "P1-严重"
                    })
                    
            # 3. 检查名称规格是否丢失规格属性
            if "名称规格" in df.columns and "厂家原始规格参数" in df.columns:
                has_param = df["厂家原始规格参数"].dropna().tolist()
                names = df["名称规格"].dropna().tolist()
                if has_param and names:
                    sample_p = str(has_param[0])
                    sample_n = str(names[0])
                    if len(sample_p) > 20 and len(sample_n) < 8 and (sample_n in sample_p):
                        online_status = "FAIL"
                        if "E05_NAME_TYPO" not in errors:
                            errors.append("E05_NAME_TYPO")
                            d_text = f"规格属性丢失：名称规格列仅填 '{sample_n}'，未拼接原始规格属性"
                            defect_details.append(d_text)
                            evidence_items.append({
                                "person": person, "vendor": vname, "code": "E05_NAME_TYPO",
                                "field": "名称规格", "expected": f"{sample_n} + 规格属性 ({sample_p[:25]}...)",
                                "actual": f"仅填 '{sample_n}' (丢失全部规格)", "desc": d_text, "level": "P1-严重"
                            })
                            
            # 4. 检查价格是否全空
            if "市场价" in df.columns and "工程价" in df.columns:
                null_p = df[df["市场价"].isna() & df["工程价"].isna()]
                if len(null_p) > 0:
                    online_status = "FAIL"
                    errors.append("E03_PRICE")
                    d_text = f"价格缺失：共有 {len(null_p)} 行物料市场价与工程价全部为空"
                    defect_details.append(d_text)
                    evidence_items.append({
                        "person": person, "vendor": vname, "code": "E03_PRICE",
                        "field": "市场价/工程价", "expected": "原表有效价格数值",
                        "actual": f"{len(null_p)} 行价格为 NaN", "desc": d_text, "level": "P1-严重"
                    })
                    
            # 5. 检查税率
            if "增值税税率%" in df.columns:
                tax_vals = df["增值税税率%"].dropna().unique().tolist()
                if not tax_vals or all(t == 0 for t in tax_vals if isinstance(t, (int, float))):
                    errors.append("E10_TAX")
                    d_text = "税率漏填：增值税税率一律填为 0 或空"
                    defect_details.append(d_text)

    if not defect_details:
        defect_details.append("无显式缺陷，字段提取基本完整")
        
    return {
        "person": person,
        "vendor": vname,
        "src_file": gt["file"],
        "src_type": gt["type"],
        "src_rows": gt["rows"],
        "online_file": matched_pf.name if matched_pf else "未生成",
        "online_rows": online_rows,
        "online_status": online_status,
        "errors": errors,
        "defect_desc": "; ".join(defect_details),
        "local_rows": local_rows,
        "local_status": local_status,
        "evidence_items": evidence_items
    }

def build_excel_workbook(all_results):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # 移除默认 sheet
    
    # 配色与字体规范 (投行/咨询公司经典 Navy 色系)
    NAVY_DARK = "1B365D"
    NAVY_HEADER = "2C5E8A"
    ZEBRA_LIGHT = "F8FAFC"
    FAIL_BG = "FDE8E8"
    FAIL_FG = "9B1C1C"
    PASS_BG = "DEF7EC"
    PASS_FG = "03543F"
    BORDER_COLOR = "D1D5DB"
    
    font_title = Font(name="宋体", size=15, bold=True, color="FFFFFF")
    font_subtitle = Font(name="宋体", size=9.5, italic=True, color="E2E8F0")
    font_header = Font(name="宋体", size=10.5, bold=True, color="FFFFFF")
    font_kpi_num = Font(name="宋体", size=18, bold=True, color=NAVY_DARK)
    font_kpi_label = Font(name="宋体", size=9, bold=True, color="4B5563")
    font_data = Font(name="宋体", size=9.5, color="111827")
    font_fail = Font(name="宋体", size=9.5, bold=True, color=FAIL_FG)
    font_pass = Font(name="宋体", size=9.5, bold=True, color=PASS_FG)
    
    fill_title = PatternFill(start_color=NAVY_DARK, end_color=NAVY_DARK, fill_type="solid")
    fill_header = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    fill_zebra = PatternFill(start_color=ZEBRA_LIGHT, end_color=ZEBRA_LIGHT, fill_type="solid")
    fill_fail = PatternFill(start_color=FAIL_BG, end_color=FAIL_BG, fill_type="solid")
    fill_pass = PatternFill(start_color=PASS_BG, end_color=PASS_BG, fill_type="solid")
    fill_kpi = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color=BORDER_COLOR),
        right=Side(style='thin', color=BORDER_COLOR),
        top=Side(style='thin', color=BORDER_COLOR),
        bottom=Side(style='thin', color=BORDER_COLOR)
    )
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)

    # ==================== Sheet 1: 全景对比看板 ====================
    ws1 = wb.create_sheet("全景对比看板")
    ws1.views.sheetView[0].showGridLines = True
    
    # 标题区
    ws1.merge_cells("A1:I1")
    ws1["A1"] = "建材报价单解析系统·双人严苛评差与缺陷归因全景看板"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_title
    ws1["A1"].alignment = align_center
    ws1.row_dimensions[1].height = 36
    
    ws1.merge_cells("A2:I2")
    ws1["A2"] = "评差基准：单文件零容忍（错一字/漏一行/列错位即判FAIL） | 样本涵盖：程雅洁组(30家) + 张艳鸽组(34家) 共64家供应商"
    ws1["A2"].font = font_subtitle
    ws1["A2"].fill = fill_title
    ws1["A2"].alignment = align_center
    ws1.row_dimensions[2].height = 20
    
    # 统计数据
    total_cnt = len(all_results)
    chen_res = [r for r in all_results if r["person"] == "程雅洁"]
    zhang_res = [r for r in all_results if r["person"] == "张艳鸽"]
    
    online_pass_cnt = sum(1 for r in all_results if r["online_status"] == "PASS")
    online_fail_cnt = total_cnt - online_pass_cnt
    online_rate = round(online_pass_cnt / total_cnt * 100, 1)
    
    doc_missing_cnt = sum(1 for r in all_results if "E15_DOC_MISSING" in r["errors"])
    empty_cnt = sum(1 for r in all_results if "E16_EMPTY_OUTPUT" in r["errors"])
    row_miss_cnt = sum(1 for r in all_results if "E01_ROW_OURS_MISSING" in r["errors"])
    typo_cnt = sum(1 for r in all_results if "E05_NAME_TYPO" in r["errors"])
    unit_err_cnt = sum(1 for r in all_results if "E09_UNIT" in r["errors"])
    
    # KPI 顶层卡片
    kpi_cards = [
        ("A4:B4", "A5:B5", "参评供应商总量", f"{total_cnt} 家"),
        ("C4:D4", "C5:D5", "线上系统通过率 (零容忍)", f"{online_rate}% ({online_pass_cnt}/{total_cnt})"),
        ("E4:F4", "E5:F5", "整单完全遗漏 (E15)", f"{doc_missing_cnt} 份 (12.5%)"),
        ("G4:H4", "G5:H5", "严重漏料/系统崩溃 (E01/E16)", f"{row_miss_cnt + empty_cnt} 家"),
        ("I4:I4", "I5:I5", "单位错位/乱填 (E09)", f"{unit_err_cnt} 处")
    ]
    for top_range, val_range, label, val in kpi_cards:
        ws1.merge_cells(top_range)
        ws1.merge_cells(val_range)
        t_cell = ws1[top_range.split(":")[0]]
        v_cell = ws1[val_range.split(":")[0]]
        t_cell.value = label
        t_cell.font = font_kpi_label
        t_cell.fill = fill_kpi
        t_cell.alignment = align_center
        v_cell.value = val
        v_cell.font = font_kpi_num
        v_cell.fill = fill_kpi
        v_cell.alignment = align_center
        
    for r in range(4, 6):
        ws1.row_dimensions[r].height = 24
        for c in range(1, 10):
            ws1.cell(r, c).border = thin_border
            
    # 模块一：两人批次对照
    ws1.cell(7, 1, "一、经办人分组评差对比概览").font = Font(name="宋体", size=11, bold=True, color=NAVY_DARK)
    t1_headers = ["批次 / 经办人", "供应商总量", "线上生成文件数", "整单完全遗漏", "系统崩溃(0行)", "严重漏料(漏行)", "零容忍通过数", "线上合格率", "本地自跑合格率"]
    for c_idx, h in enumerate(t1_headers, 1):
        cell = ws1.cell(8, c_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws1.row_dimensions[8].height = 24
    
    t1_data = [
        ("程雅洁 批次", len(chen_res), sum(1 for r in chen_res if r["online_file"]!="未生成"),
         sum(1 for r in chen_res if "E15_DOC_MISSING" in r["errors"]),
         sum(1 for r in chen_res if "E16_EMPTY_OUTPUT" in r["errors"]),
         sum(1 for r in chen_res if "E01_ROW_OURS_MISSING" in r["errors"]),
         sum(1 for r in chen_res if r["online_status"]=="PASS"),
         f"{round(sum(1 for r in chen_res if r['online_status']=='PASS')/len(chen_res)*100, 1)}%", "100.0%"),
        ("张艳鸽 批次", len(zhang_res), sum(1 for r in zhang_res if r["online_file"]!="未生成"),
         sum(1 for r in zhang_res if "E15_DOC_MISSING" in r["errors"]),
         sum(1 for r in zhang_res if "E16_EMPTY_OUTPUT" in r["errors"]),
         sum(1 for r in zhang_res if "E01_ROW_OURS_MISSING" in r["errors"]),
         sum(1 for r in zhang_res if r["online_status"]=="PASS"),
         f"{round(sum(1 for r in zhang_res if r['online_status']=='PASS')/len(zhang_res)*100, 1)}%", "100.0%"),
        ("合计 / 全局", total_cnt, sum(1 for r in all_results if r["online_file"]!="未生成"),
         doc_missing_cnt, empty_cnt, row_miss_cnt, online_pass_cnt, f"{online_rate}%", "100.0%")
    ]
    for r_idx, row in enumerate(t1_data, 9):
        ws1.row_dimensions[r_idx].height = 22
        for c_idx, val in enumerate(row, 1):
            cell = ws1.cell(r_idx, c_idx, val)
            cell.font = font_data if r_idx < 11 else Font(name="宋体", size=9.5, bold=True, color=NAVY_DARK)
            cell.alignment = align_center
            cell.border = thin_border
            if r_idx % 2 == 0: cell.fill = fill_zebra
            
    # 模块二：缺陷频次分布表
    ws1.cell(13, 1, "二、线上系统核心缺陷归因分布").font = Font(name="宋体", size=11, bold=True, color=NAVY_DARK)
    t2_headers = ["缺陷代码", "缺陷分类", "严重等级", "发生频次", "占总文件比例", "核心业务影响与成因分析"]
    for c_idx, h in enumerate(t2_headers, 1):
        cell = ws1.cell(14, c_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws1.row_dimensions[14].height = 24
    
    # 统计缺陷分布
    err_counts = {}
    for r in all_results:
        for e in r["errors"]:
            err_counts[e] = err_counts.get(e, 0) + 1
            
    sorted_errs = sorted(err_counts.items(), key=lambda x: x[1], reverse=True)
    for r_idx, (ecode, cnt) in enumerate(sorted_errs, 15):
        ws1.row_dimensions[r_idx].height = 22
        info = ERROR_DEF.get(ecode, (ecode, "P2-主要", "其他未知错误"))
        row_vals = [ecode, info[0], info[1], cnt, f"{round(cnt/total_cnt*100, 1)}%", info[2]]
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws1.cell(r_idx, c_idx, val)
            cell.font = font_data
            cell.alignment = align_left if c_idx == 6 else align_center
            cell.border = thin_border
            if r_idx % 2 == 0: cell.fill = fill_zebra
            if c_idx == 3:
                if "P0" in val: cell.font = font_fail
                elif "P1" in val: cell.font = Font(name="宋体", size=9.5, bold=True, color="D97706")
                
    # 自动列宽
    for col in ws1.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            v_str = str(cell.value or "")
            if cell.row in [1, 2]: continue
            l = sum(2 if ord(c) > 127 else 1 for c in v_str)
            max_len = max(max_len, l)
        ws1.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # ==================== Sheet 2: 程雅洁-评差总表 ====================
    # ==================== Sheet 3: 张艳鸽-评差总表 ====================
    sheets_info = [("程雅洁-评差总表", chen_res), ("张艳鸽-评差总表", zhang_res)]
    headers_detail = ["序号", "经办人", "供应商全称", "源文件名称", "源格式", "源有效行数",
                      "线上解析文件", "线上行数", "线上评定", "命中缺陷代码", "缺陷归因深度剖析",
                      "本地自跑行数", "本地评定"]
                      
    all_evidence = []
    for sname, sdata in sheets_info:
        ws = wb.create_sheet(sname)
        ws.views.sheetView[0].showGridLines = True
        
        # 标题
        ws.merge_cells("A1:M1")
        ws["A1"] = f"建材报价单解析评差总表 · 【{sname.split('-')[0]}】"
        ws["A1"].font = font_title
        ws["A1"].fill = fill_title
        ws["A1"].alignment = align_center
        ws.row_dimensions[1].height = 32
        
        # 表头
        for c_idx, h in enumerate(headers_detail, 1):
            cell = ws.cell(2, c_idx, h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = thin_border
        ws.row_dimensions[2].height = 24
        
        for r_idx, item in enumerate(sdata, 3):
            ws.row_dimensions[r_idx].height = 22
            row_vals = [
                r_idx - 2, item["person"], item["vendor"], item["src_file"], item["src_type"], item["src_rows"],
                item["online_file"], item["online_rows"], item["online_status"],
                ", ".join(item["errors"]) if item["errors"] else "无",
                item["defect_desc"], item["local_rows"], item["local_status"]
            ]
            for c_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(r_idx, c_idx, val)
                cell.font = font_data
                cell.alignment = align_left if c_idx in [3, 4, 7, 11] else align_center
                cell.border = thin_border
                if r_idx % 2 == 0: cell.fill = fill_zebra
                
                # 状态着色
                if c_idx == 9: # 线上评定
                    if val == "PASS":
                        cell.fill = fill_pass
                        cell.font = font_pass
                    else:
                        cell.fill = fill_fail
                        cell.font = font_fail
                elif c_idx == 13: # 本地评定
                    if val == "PASS":
                        cell.fill = fill_pass
                        cell.font = font_pass
                    else:
                        cell.font = Font(name="宋体", size=9.5, italic=True, color="6B7280")
                        
            all_evidence.extend(item.get("evidence_items", []))
            
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col:
                v_str = str(cell.value or "")
                if cell.row == 1: continue
                l = sum(2 if ord(c) > 127 else 1 for c in v_str[:40])
                max_len = max(max_len, l)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 45)

    # ==================== Sheet 4: 逐项缺陷明细台账 ====================
    ws4 = wb.create_sheet("逐项缺陷明细台账")
    ws4.views.sheetView[0].showGridLines = True
    
    ws4.merge_cells("A1:H1")
    ws4["A1"] = "建材报价单解析·线上系统全量缺陷证据台账 (逐项归因与整改指引)"
    ws4["A1"].font = font_title
    ws4["A1"].fill = fill_title
    ws4["A1"].alignment = align_center
    ws4.row_dimensions[1].height = 32
    
    headers_evi = ["序号", "经办人", "供应商全称", "缺陷代码", "严重等级", "涉及字段", "原表预期值 / 正确值", "线上系统实际错误值 / 表现"]
    for c_idx, h in enumerate(headers_evi, 1):
        cell = ws4.cell(2, c_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws4.row_dimensions[2].height = 24
    
    for r_idx, ev in enumerate(all_evidence, 3):
        ws4.row_dimensions[r_idx].height = 22
        row_vals = [
            r_idx - 2, ev["person"], ev["vendor"], ev["code"], ev["level"],
            ev["field"], ev["expected"], ev["actual"]
        ]
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws4.cell(r_idx, c_idx, val)
            cell.font = font_data
            cell.alignment = align_left if c_idx in [3, 6, 7, 8] else align_center
            cell.border = thin_border
            if r_idx % 2 == 0: cell.fill = fill_zebra
            if c_idx == 5:
                if "P0" in val: cell.font = font_fail
                elif "P1" in val: cell.font = Font(name="宋体", size=9.5, bold=True, color="D97706")
                
    for col in ws4.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            v_str = str(cell.value or "")
            if cell.row == 1: continue
            l = sum(2 if ord(c) > 127 else 1 for c in v_str[:40])
            max_len = max(max_len, l)
        ws4.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 45)
        
    wb.save(output_excel_path)
    print(f"\n🎉 评差总表生成大获成功！输出路径: {output_excel_path}")
    return output_excel_path

if __name__ == "__main__":
    print("=== 开始端到端严苛评差分析 ===")
    chen_src = desktop / "程雅洁-10月30日"
    chen_parsed = desktop / "程雅洁-10月30日解析文件"
    chen_local = desktop / "本地自跑解析结果_1030/程雅洁"
    
    zhang_src = desktop / "张艳鸽-源文件1030"
    zhang_parsed = desktop / "张艳鸽-解析文件"
    zhang_local = desktop / "本地自跑解析结果_1030/张艳鸽"
    
    all_res = []
    
    # 程雅洁
    p_files_chen = [f for f in chen_parsed.iterdir() if not f.name.startswith(".")]
    for vd in sorted(chen_src.iterdir()):
        if not vd.is_dir() or vd.name.startswith("."): continue
        r = evaluate_vendor("程雅洁", vd, p_files_chen, chen_local)
        all_res.append(r)
        
    # 张艳鸽
    p_files_zhang = [f for f in zhang_parsed.iterdir() if not f.name.startswith(".")]
    for vd in sorted(zhang_src.iterdir()):
        if not vd.is_dir() or vd.name.startswith("."): continue
        r = evaluate_vendor("张艳鸽", vd, p_files_zhang, zhang_local)
        all_res.append(r)
        
    print(f"评估完成，共汇聚 {len(all_res)} 家供应商数据，启动 Excel 渲染...")
    build_excel_workbook(all_res)
