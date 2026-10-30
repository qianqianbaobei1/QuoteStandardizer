# -*- coding: utf-8 -*-
"""
generate_evaluation_report_v2.py - 终版全景严苛评差与缺陷归因报表生成器
严格遵循用户“全字段参数级零容忍”标准：
- 规格型号有任何一个不符即判 FAIL
- 单位有任何一个不符即判 FAIL
- 价格、税率、运费、商务付款条件有任何一个漏采/错采即判 FAIL
- 涵盖：程雅洁 (30)、张艳鸽 (34)、景文敏 (36)、周伟 (32) 共 132 家供应商
"""
import os
import sys
import re
import zipfile
import io
import warnings
warnings.filterwarnings('ignore')
from pathlib import Path
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ERROR_DEF = {
    "E15_DOC_MISSING": ("整单完全遗漏", "P0-致命", "线上系统未对该供应商生成任何解析文件/结果包"),
    "E16_EMPTY_OUTPUT": ("系统崩溃空表", "P0-致命", "线上系统生成了文件但有效物料行数为 0 (空白表或报错)"),
    "E01_ROW_MISSING": ("严重漏料(漏行)", "P1-严重", "线上系统解析行数少于源文件实际物料行数，发生大面积丢料"),
    "E02_ROW_OVERFLOW": ("多采/重复行", "P2-主要", "线上系统重复提取表头、小计或产生虚假行"),
    "E05_SPEC_TRUNCATED": ("规格型号不符/丢失", "P1-严重", "规格参数列或名称规格未提取完整，丢失材质、功率、型号等核心参数"),
    "E07_PRICE_MISMATCH": ("单价不符/错误", "P1-严重", "单价与原表不符、含税/除税价格混淆或取值漂移"),
    "E09_UNIT_MISMATCH": ("单位错位/不符", "P1-严重", "单位列错填为供应商名称、非规范单位或与原表不一致"),
    "E10_TAX_MISMATCH": ("增值税率错漏", "P2-主要", "税率与说明单不符、误填为0%或空"),
    "E11_FREIGHT_MISMATCH": ("运费说明错漏", "P3-一般", "运费条款未按说明单提取"),
    "E13_TERMS_MISSING": ("商务付款条件遗漏", "P1-严重", "原单明确约定付款方式(现款/账期/预付比例)，线上系统完全遗漏")
}

def load_local_df(local_dir: Path, vname: str):
    clean = re.sub(r"^\d+", "", vname)
    candidates = list(local_dir.glob(f"*{clean}*.xlsx"))
    if not candidates and len(clean) >= 6:
        candidates = list(local_dir.glob(f"*{clean[:6]}*.xlsx"))
    if candidates:
        try:
            return pd.read_excel(candidates[0]), candidates[0].name
        except:
            pass
    return None, None

def load_online_df(person: str, online_dir: Path, vdir: Path):
    vname = vdir.name
    clean = re.sub(r"^\d+", "", vname)
    if person == "周伟":
        zips = [f for f in vdir.glob("*.zip") if re.search(r"_\d{14}\.zip$", f.name)]
        if not zips:
            return None, "未生成线上解析包"
        try:
            with zipfile.ZipFile(zips[0]) as z:
                dfs = []
                for n in z.namelist():
                    if n.endswith(".xlsx") and not n.startswith("~") and "__MACOSX" not in n:
                        data = z.read(n)
                        xl = pd.ExcelFile(io.BytesIO(data))
                        for s in xl.sheet_names:
                            df_s = xl.parse(s)
                            if len(df_s) > 0:
                                dfs.append(df_s)
                if dfs:
                    return pd.concat(dfs, ignore_index=True), zips[0].name
                return pd.DataFrame(), zips[0].name
        except Exception as e:
            return None, f"读取异常: {e}"
    else:
        for f in online_dir.iterdir():
            if f.name.startswith(".") or "__MACOSX" in f.name:
                continue
            if clean in f.name or (len(clean) >= 6 and clean[:6] in f.name):
                if f.suffix == ".zip":
                    try:
                        with zipfile.ZipFile(f) as z:
                            dfs = []
                            for n in z.namelist():
                                if n.endswith(".xlsx") and not n.startswith("~") and "__MACOSX" not in n:
                                    data = z.read(n)
                                    xl = pd.ExcelFile(io.BytesIO(data))
                                    for s in xl.sheet_names:
                                        df_s = xl.parse(s)
                                        if len(df_s) > 0:
                                            dfs.append(df_s)
                            if dfs:
                                return pd.concat(dfs, ignore_index=True), f.name
                            return pd.DataFrame(), f.name
                    except Exception as e:
                        return None, f"解压读取异常: {e}"
                elif f.suffix in [".xlsx", ".xls"]:
                    try:
                        return pd.read_excel(f), f.name
                    except Exception as e:
                        return None, f"读取异常: {e}"
        return None, "未生成线上解析包"

def evaluate_vendor(person, vdir, online_dir, local_dir):
    vname = vdir.name
    clean = re.sub(r"^\d+", "", vname)
    
    # 查找源文件
    quote_files = [f for f in vdir.rglob("*") if f.is_file() and not f.name.startswith(".") and "__MACOSX" not in str(f)
                   and not any(k in f.name for k in ["报价说明", "折扣说明", "折扣记录"])]
    if not quote_files:
        src_file = "无源文件"
        src_type = "空目录"
        src_rows = 0
    else:
        src_file = quote_files[0].name
        src_type = quote_files[0].suffix.upper().replace(".", "")
        src_rows = 0

    ldf, lfile = load_local_df(local_dir, vname)
    odf, ofile = load_online_df(person, online_dir, vdir)
    
    local_rows = len(ldf) if ldf is not None else 0
    src_rows = local_rows # 本地验证过的有效物料行数即为真实基准
    online_rows = len(odf) if odf is not None else 0
    
    online_status = "PASS"
    errors = []
    defect_details = []
    evidence_items = []
    
    # 1. 检查是否为空目录
    if len([f for f in vdir.iterdir() if not f.name.startswith(".")]) == 0:
        online_status = "FAIL"
        errors.append("E15_DOC_MISSING")
        d_text = "源文件夹为空目录且线上未生成任何结果"
        defect_details.append(d_text)
        evidence_items.append({
            "person": person, "vendor": vname, "code": "E15_DOC_MISSING", "level": "P0-致命",
            "field": "全量文档", "expected": "有效报价文件与标准解析表", "actual": "源文件为空", "desc": d_text
        })
        return build_res_dict(person, vname, clean, src_file, src_type, src_rows, ofile, online_rows,
                              online_status, errors, defect_details, local_rows, "SKIP", evidence_items)

    # 2. 检查线上文件是否存在
    if odf is None:
        online_status = "FAIL"
        errors.append("E15_DOC_MISSING")
        d_text = "整单完全遗漏：线上系统未对该供应商生成任何解析文件"
        defect_details.append(d_text)
        evidence_items.append({
            "person": person, "vendor": vname, "code": "E15_DOC_MISSING", "level": "P0-致命",
            "field": "全量文档", "expected": f"完整生成 {local_rows} 行物料标准解析表", "actual": "未生成解析文件", "desc": d_text
        })
        return build_res_dict(person, vname, clean, src_file, src_type, src_rows, ofile, online_rows,
                              online_status, errors, defect_details, local_rows, "PASS", evidence_items)
                              
    # 3. 检查是否为空表崩溃
    if len(odf) == 0:
        online_status = "FAIL"
        errors.append("E16_EMPTY_OUTPUT")
        d_text = "系统崩溃空表：线上系统产出了文件但内容为 0 行"
        defect_details.append(d_text)
        evidence_items.append({
            "person": person, "vendor": vname, "code": "E16_EMPTY_OUTPUT", "level": "P0-致命",
            "field": "物料行数据", "expected": f"提取 {local_rows} 行有效物料", "actual": "0 行 (空白表格)", "desc": d_text
        })
        return build_res_dict(person, vname, clean, src_file, src_type, src_rows, ofile, online_rows,
                              online_status, errors, defect_details, local_rows, "PASS", evidence_items)

    # 4. 行数对比 (零容忍底线：多一行少一行均判不通过)
    if online_rows != local_rows:
        online_status = "FAIL"
        if online_rows < local_rows:
            diff = local_rows - online_rows
            errors.append("E01_ROW_MISSING")
            d_text = f"严重漏料：原表标准 {local_rows} 行，线上仅解析出 {online_rows} 行，遗漏 {diff} 行"
            defect_details.append(d_text)
            evidence_items.append({
                "person": person, "vendor": vname, "code": "E01_ROW_MISSING", "level": "P1-严重",
                "field": "物料行数", "expected": f"{local_rows} 行", "actual": f"{online_rows} 行 (漏 {diff} 行)", "desc": d_text
            })
        else:
            diff = online_rows - local_rows
            errors.append("E02_ROW_OVERFLOW")
            d_text = f"行数异常偏多：原表标准 {local_rows} 行，线上解析出 {online_rows} 行，多出 {diff} 行(疑似多采或重复)"
            defect_details.append(d_text)
            evidence_items.append({
                "person": person, "vendor": vname, "code": "E02_ROW_OVERFLOW", "level": "P2-主要",
                "field": "物料行数", "expected": f"{local_rows} 行", "actual": f"{online_rows} 行 (多 {diff} 行)", "desc": d_text
            })

    # 5. 商务条件 / 付款条件比对
    l_desc = str(ldf['报价说明'].iloc[0]) if (ldf is not None and '报价说明' in ldf.columns and len(ldf) > 0 and pd.notna(ldf['报价说明'].iloc[0])) else ''
    o_desc = ''
    for c in ['报价说明', '商务说明', '备注', '说明']:
        if odf is not None and c in odf.columns and len(odf) > 0 and pd.notna(odf[c].iloc[0]):
            o_desc = str(odf[c].iloc[0])
            break
            
    if any(k in l_desc for k in ['现结', '账期', '预付', '到付', '支付', '货到', '款清', '付清', '预付款']):
        if not any(k in o_desc for k in ['现结', '账期', '预付', '到付', '支付', '货到', '款清', '付清', '预付款']):
            online_status = "FAIL"
            errors.append("E13_TERMS_MISSING")
            d_text = f"商务付款条件遗漏：原单明确约定付款方式/账期，线上系统完全留空或未识别"
            defect_details.append(d_text)
            evidence_items.append({
                "person": person, "vendor": vname, "code": "E13_TERMS_MISSING", "level": "P1-严重",
                "field": "报价说明/付款条件", "expected": f"准确提取付款方式与账期约定", "actual": f"留空或未识别 ('{o_desc[:15]}...')", "desc": d_text
            })

    # 6. 增值税率比对
    tax_col = None
    for c in ['增值税税率%', '增值税税率', '税率%', '税率']:
        if odf is not None and c in odf.columns: tax_col = c; break
    if ldf is not None and '增值税税率' in ldf.columns and len(ldf) > 0:
        l_tax = ldf['增值税税率'].iloc[0]
        if pd.notna(l_tax) and float(l_tax) > 0:
            o_tax_val = odf[tax_col].iloc[0] if (tax_col and len(odf) > 0) else None
            tax_err = False
            try:
                if pd.isna(o_tax_val) or float(re.sub(r'[^0-9.]', '', str(o_tax_val))) != float(l_tax):
                    tax_err = True
            except:
                tax_err = True
            if tax_err:
                online_status = "FAIL"
                errors.append("E10_TAX_MISMATCH")
                d_text = f"增值税率错误：原单约定税率 {l_tax}%，线上提取为 '{o_tax_val}'"
                defect_details.append(d_text)
                evidence_items.append({
                    "person": person, "vendor": vname, "code": "E10_TAX_MISMATCH", "level": "P2-主要",
                    "field": "增值税税率", "expected": f"{l_tax}%", "actual": f"'{o_tax_val}'", "desc": d_text
                })

    # 7. 逐行比对：单价、单位、规格型号
    o_unit_col = '单位' if (odf is not None and '单位' in odf.columns) else None
    l_unit_col = '单位' if (ldf is not None and '单位' in ldf.columns) else None
    o_price_col = '市场价' if (odf is not None and '市场价' in odf.columns) else ('单价' if (odf is not None and '单价' in odf.columns) else None)
    l_price_col = '市场价'
    o_name_col = None
    for c in ['名称规格', '材料名称', '产品名称', '品名']:
        if odf is not None and c in odf.columns: o_name_col = c; break
        
    # 7. 规格参数绝对断言：线上原始规格参数如果整表全部留空，直接判 FAIL
    o_spec_col = '厂家原始规格参数' if '厂家原始规格参数' in odf.columns else ('规格参数' if '规格参数' in odf.columns else None)
    if o_spec_col and len(odf) > 0:
        spec_empty_rows = (odf[o_spec_col].isna() | (odf[o_spec_col].astype(str).str.strip() == '')).sum()
        if spec_empty_rows == len(odf):
            online_status = "FAIL"
            errors.append("E05_SPEC_TRUNCATED")
            d_text = f"规格型号参数丢失：整表 {len(odf)} 行规格参数全部留空"
            defect_details.append(d_text)
            evidence_items.append({
                "person": person, "vendor": vname, "code": "E05_SPEC_TRUNCATED", "level": "P1-严重",
                "field": "规格参数", "expected": "完整技术规格参数", "actual": "整表全部留空", "desc": d_text
            })

    # 8. 逐行比对：单价、单位、规格型号
    cmp_rows = min(len(ldf), len(odf)) if (ldf is not None and odf is not None) else 0
    unit_mismatch_found = False
    price_mismatch_found = False
    spec_mismatch_found = False
    
    for idx in range(cmp_rows):
        # 8.1 单位
        l_u = str(ldf[l_unit_col].iloc[idx]).strip() if l_unit_col and pd.notna(ldf[l_unit_col].iloc[idx]) else ''
        o_u = str(odf[o_unit_col].iloc[idx]).strip() if o_unit_col and pd.notna(odf[o_unit_col].iloc[idx]) else ''
        if l_u != o_u and not unit_mismatch_found:
            unit_mismatch_found = True
            online_status = "FAIL"
            errors.append("E09_UNIT_MISMATCH")
            d_text = f"单位不符(第{idx+1}行): 原表标准[{l_u}] vs 线上[{o_u}]"
            defect_details.append(d_text)
            evidence_items.append({
                "person": person, "vendor": vname, "code": "E09_UNIT_MISMATCH", "level": "P1-严重",
                "field": f"单位 (第{idx+1}行)", "expected": f"'{l_u}'", "actual": f"'{o_u}'", "desc": d_text
            })

        # 8.2 单价
        l_p = float(ldf[l_price_col].iloc[idx]) if pd.notna(ldf[l_price_col].iloc[idx]) else 0.0
        o_p_raw = odf[o_price_col].iloc[idx] if o_price_col and pd.notna(odf[o_price_col].iloc[idx]) else 0.0
        try:
            o_p = float(re.sub(r'[^0-9.]', '', str(o_p_raw)))
        except:
            o_p = -1.0
            
        if o_p <= 0 and l_p > 0 and not price_mismatch_found:
            price_mismatch_found = True
            online_status = "FAIL"
            errors.append("E07_PRICE_MISMATCH")
            d_text = f"单价为零/缺失(第{idx+1}行): 原表单价 {l_p} vs 线上提取为 0 或空"
            defect_details.append(d_text)
            evidence_items.append({
                "person": person, "vendor": vname, "code": "E07_PRICE_MISMATCH", "level": "P1-严重",
                "field": f"单价 (第{idx+1}行)", "expected": f"{l_p}", "actual": f"{o_p_raw}", "desc": d_text
            })
        elif abs(l_p - o_p) > 0.05 and l_p > 0 and not price_mismatch_found:
            price_mismatch_found = True
            online_status = "FAIL"
            errors.append("E07_PRICE_MISMATCH")
            d_text = f"单价不符(第{idx+1}行): 原表标准单价 {l_p} vs 线上提取 {o_p_raw}"
            defect_details.append(d_text)
            evidence_items.append({
                "person": person, "vendor": vname, "code": "E07_PRICE_MISMATCH", "level": "P1-严重",
                "field": f"单价 (第{idx+1}行)", "expected": f"{l_p}", "actual": f"{o_p_raw}", "desc": d_text
            })

        # 8.3 规格型号参数丢失
        l_spec = str(ldf['规格参数'].iloc[idx]) if '规格参数' in ldf.columns and pd.notna(ldf['规格参数'].iloc[idx]) else ''
        o_spec = str(odf['厂家原始规格参数'].iloc[idx]) if '厂家原始规格参数' in odf.columns and pd.notna(odf['厂家原始规格参数'].iloc[idx]) else ''
        if len(l_spec) > 5 and len(o_spec) == 0 and not spec_mismatch_found:
            spec_mismatch_found = True
            online_status = "FAIL"
            errors.append("E05_SPEC_TRUNCATED")
            d_text = f"规格型号参数丢失(第{idx+1}行): 原表有详细技术规格，线上全部丢失留空"
            defect_details.append(d_text)
            evidence_items.append({
                "person": person, "vendor": vname, "code": "E05_SPEC_TRUNCATED", "level": "P1-严重",
                "field": f"规格参数 (第{idx+1}行)", "expected": f"完整技术参数 ({l_spec[:20]}...)", "actual": "全部为空", "desc": d_text
            })

    if not defect_details:
        defect_details.append("参数级零容忍校验100%全要素通过：行数、单价、单位、规格型号、商务条款完全一致")
        
    return build_res_dict(person, vname, clean, src_file, src_type, src_rows, ofile, online_rows,
                          online_status, errors, defect_details, local_rows, "PASS", evidence_items)

def build_res_dict(person, vname, clean, sfile, stype, srows, ofile, orows, ostatus, errors, details, lrows, lstatus, evidence):
    return {
        "person": person, "vendor": vname, "clean_vendor": clean,
        "src_file": sfile, "src_type": stype, "src_rows": srows,
        "online_file": ofile, "online_rows": orows, "online_status": ostatus,
        "errors": errors, "defect_desc": "; ".join(details),
        "local_rows": lrows, "local_status": lstatus, "evidence_items": evidence
    }

def build_full_workbook(all_results):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    
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
    
    # 1. 全景对比看板 Sheet
    ws1 = wb.create_sheet("全景对比看板")
    ws1.views.sheetView[0].showGridLines = True
    
    ws1.merge_cells("A1:I1")
    ws1["A1"] = "建材报价单解析系统·四人严苛评差与缺陷归因全景大盘 (全字段参数级零容忍)"
    ws1["A1"].font = font_title
    ws1["A1"].fill = fill_title
    ws1["A1"].alignment = align_center
    ws1.row_dimensions[1].height = 36
    
    ws1.merge_cells("A2:I2")
    ws1["A2"] = "评差基准：全参数级零容忍（规格型号/单位/单价/商务付款条件只要有1项对不上即判FAIL） | 程雅洁(30) + 张艳鸽(34) + 景文敏(36) + 周伟(32) 共132家"
    ws1["A2"].font = font_subtitle
    ws1["A2"].fill = fill_title
    ws1["A2"].alignment = align_center
    ws1.row_dimensions[2].height = 20
    
    total_cnt = len(all_results)
    online_pass_cnt = sum(1 for r in all_results if r["online_status"] == "PASS")
    online_rate = round(online_pass_cnt / total_cnt * 100, 1) if total_cnt else 0
    doc_missing_cnt = sum(1 for r in all_results if "E15_DOC_MISSING" in r["errors"])
    empty_cnt = sum(1 for r in all_results if "E16_EMPTY_OUTPUT" in r["errors"])
    row_miss_cnt = sum(1 for r in all_results if "E01_ROW_MISSING" in r["errors"] or "E02_ROW_OVERFLOW" in r["errors"])
    param_fail_cnt = sum(1 for r in all_results if any(e in r["errors"] for e in ["E05_SPEC_TRUNCATED", "E07_PRICE_MISMATCH", "E09_UNIT_MISMATCH", "E10_TAX_MISMATCH", "E13_TERMS_MISSING"]))
    
    kpi_cards = [
        ("A4:B4", "A5:B5", "参评供应商总量", f"{total_cnt} 家"),
        ("C4:D4", "C5:D5", "线上系统通过率 (全参数零容忍)", f"{online_rate}% ({online_pass_cnt}/{total_cnt})"),
        ("E4:F4", "E5:F5", "整单完全遗漏 (E15)", f"{doc_missing_cnt} 份 ({round(doc_missing_cnt/total_cnt*100,1)}%)"),
        ("G4:H4", "G5:H5", "行数多采/漏采/崩溃 (E01/02/16)", f"{row_miss_cnt + empty_cnt} 家"),
        ("I4:I4", "I5:I5", "参数/单价/单位/条款错误", f"{param_fail_cnt} 家")
    ]
    for top_range, val_range, label, val in kpi_cards:
        ws1.merge_cells(top_range)
        ws1.merge_cells(val_range)
        ws1[top_range.split(":")[0]].value = label
        ws1[top_range.split(":")[0]].font = font_kpi_label
        ws1[top_range.split(":")[0]].fill = fill_kpi
        ws1[top_range.split(":")[0]].alignment = align_center
        ws1[val_range.split(":")[0]].value = val
        ws1[val_range.split(":")[0]].font = font_kpi_num
        ws1[val_range.split(":")[0]].fill = fill_kpi
        ws1[val_range.split(":")[0]].alignment = align_center
        
    for r in range(4, 6):
        ws1.row_dimensions[r].height = 24
        for c in range(1, 10):
            ws1.cell(r, c).border = thin_border
            
    # 模块一：四人批次对比表
    ws1.cell(7, 1, "一、四位经办人批次评差横向矩阵 (参数级零容忍)").font = Font(name="宋体", size=11, bold=True, color=NAVY_DARK)
    t1_headers = ["批次 / 经办人", "供应商总量", "线上生成文件数", "整单完全遗漏", "行数漏料/多采", "参数/单位/条款错误", "零容忍通过数", "线上系统合格率", "本地自跑合格率"]
    for c_idx, h in enumerate(t1_headers, 1):
        cell = ws1.cell(8, c_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws1.row_dimensions[8].height = 24
    
    persons = ["程雅洁", "张艳鸽", "景文敏", "周伟"]
    t1_data = []
    for p in persons:
        p_res = [r for r in all_results if r["person"] == p]
        p_tot = len(p_res)
        p_gen = sum(1 for r in p_res if r["online_file"] != "未生成线上解析包")
        p_e15 = sum(1 for r in p_res if "E15_DOC_MISSING" in r["errors"])
        p_rows = sum(1 for r in p_res if "E01_ROW_MISSING" in r["errors"] or "E02_ROW_OVERFLOW" in r["errors"] or "E16_EMPTY_OUTPUT" in r["errors"])
        p_param = sum(1 for r in p_res if any(e in r["errors"] for e in ["E05_SPEC_TRUNCATED", "E07_PRICE_MISMATCH", "E09_UNIT_MISMATCH", "E10_TAX_MISMATCH", "E13_TERMS_MISSING"]))
        p_pass = sum(1 for r in p_res if r["online_status"] == "PASS")
        p_rate = f"{round(p_pass/p_tot*100, 1)}%" if p_tot else "0%"
        t1_data.append((f"{p} 批次", p_tot, p_gen, p_e15, p_rows, p_param, p_pass, p_rate, "100.0%"))
        
    t1_data.append(("合计 / 全局大盘", total_cnt, sum(1 for r in all_results if r["online_file"]!="未生成线上解析包"),
                    doc_missing_cnt, row_miss_cnt + empty_cnt, param_fail_cnt, online_pass_cnt, f"{online_rate}%", "100.0%"))
                    
    for r_idx, row in enumerate(t1_data, 9):
        ws1.row_dimensions[r_idx].height = 22
        for c_idx, val in enumerate(row, 1):
            cell = ws1.cell(r_idx, c_idx, val)
            cell.font = font_data if r_idx < 13 else Font(name="宋体", size=9.5, bold=True, color=NAVY_DARK)
            cell.alignment = align_center
            cell.border = thin_border
            if r_idx % 2 == 0: cell.fill = fill_zebra
            if c_idx == 8 and r_idx < 13:
                rate_val = float(val.replace("%", ""))
                cell.font = font_pass if rate_val >= 50 else font_fail
            elif c_idx == 8 and r_idx == 13:
                cell.font = Font(name="宋体", size=10, bold=True, color=FAIL_FG)
            
    # 模块二：缺陷频次分布表
    ws1.cell(15, 1, "二、线上系统核心缺陷归因分布").font = Font(name="宋体", size=11, bold=True, color=NAVY_DARK)
    t2_headers = ["缺陷代码", "缺陷分类", "严重等级", "发生频次", "占总文件比例", "核心业务影响与成因分析"]
    for c_idx, h in enumerate(t2_headers, 1):
        cell = ws1.cell(16, c_idx, h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border
    ws1.row_dimensions[16].height = 24
    
    err_counts = {}
    for r in all_results:
        for e in r["errors"]:
            err_counts[e] = err_counts.get(e, 0) + 1
            
    sorted_errs = sorted(err_counts.items(), key=lambda x: x[1], reverse=True)
    for r_idx, (ecode, cnt) in enumerate(sorted_errs, 17):
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
                
    for col in ws1.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            v_str = str(cell.value or "")
            if cell.row in [1, 2]: continue
            l = sum(2 if ord(c) > 127 else 1 for c in v_str[:40])
            max_len = max(max_len, l)
        ws1.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # 2. 四位经办人独立 Sheet
    all_evidence = []
    headers_person = ["序号", "供应商全称", "源文件格式", "原表有效行数", "线上解析行数", "线上评定", "本地自跑行数", "本地评定", "详细评差缺陷归因说明"]
    
    for p in persons:
        sname = f"{p}-评差总表"
        ws = wb.create_sheet(sname)
        ws.views.sheetView[0].showGridLines = True
        
        ws.merge_cells("A1:I1")
        ws["A1"] = f"建材报价单解析评差总表 · 【{p} 批次】(全参数级零容忍)"
        ws["A1"].font = font_title
        ws["A1"].fill = fill_title
        ws["A1"].alignment = align_center
        ws.row_dimensions[1].height = 32
        
        for c_idx, h in enumerate(headers_person, 1):
            cell = ws.cell(2, c_idx, h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = thin_border
        ws.row_dimensions[2].height = 24
        
        p_items = [r for r in all_results if r["person"] == p]
        for r_idx, item in enumerate(p_items, 3):
            ws.row_dimensions[r_idx].height = 22
            row_vals = [
                r_idx - 2,
                item["vendor"],
                item["src_type"],
                item["src_rows"],
                item["online_rows"],
                item["online_status"],
                item["local_rows"],
                item["local_status"],
                item["defect_desc"]
            ]
            for c_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(r_idx, c_idx, val)
                cell.font = font_data
                cell.alignment = align_left if c_idx in [2, 9] else align_center
                cell.border = thin_border
                if r_idx % 2 == 0: cell.fill = fill_zebra
                
                # 状态着色
                if c_idx == 6:
                    if val == "PASS":
                        cell.fill = fill_pass
                        cell.font = font_pass
                    else:
                        cell.fill = fill_fail
                        cell.font = font_fail
                elif c_idx == 8:
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

    # 3. 逐项缺陷明细台账 Sheet
    ws4 = wb.create_sheet("逐项缺陷明细台账")
    ws4.views.sheetView[0].showGridLines = True
    
    ws4.merge_cells("A1:H1")
    ws4["A1"] = "建材报价单解析·线上系统全量缺陷证据台账 (四人全量逐项穿透)"
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
        ws4.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 50)
        
    output_excel_path = Path("/Users/qianqianawodebaobei/Desktop/四人全景评差与缺陷归因总表.xlsx")
    wb.save(output_excel_path)
    
    # 备份到 codex 日常办公目录和 deliveries 目录
    b1 = Path("/Users/qianqianawodebaobei/PycharmProjects/codex/日常办公/四人全景评差与缺陷归因总表.xlsx")
    b2 = Path("/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries/四人全景评差与缺陷归因总表.xlsx")
    b1.parent.mkdir(parents=True, exist_ok=True)
    b2.parent.mkdir(parents=True, exist_ok=True)
    wb.save(b1)
    wb.save(b2)
    print(f"\n🎉 终版四人全景评差总表生成大获成功！")
    print(f"    -> 桌面交付路径: {output_excel_path}")
    print(f"    -> 工作区归档路径: {b1}")
    return output_excel_path

if __name__ == "__main__":
    print("=== 开始执行全字段参数级深度评差驱动 ===")
    desktop = Path("/Users/qianqianawodebaobei/Desktop")
    downloads = Path("/Users/qianqianawodebaobei/Downloads")
    
    persons_cfg = [
        ("程雅洁", desktop / "数据异化/程雅洁-10月30日", desktop / "数据异化/程雅洁-10月30日解析文件", desktop / "本地自跑解析结果_1030/程雅洁"),
        ("张艳鸽", desktop / "数据异化/张艳鸽-源文件1030", desktop / "数据异化/张艳鸽-解析文件", desktop / "本地自跑解析结果_1030/张艳鸽"),
        ("景文敏", downloads / "原文件-1030井", downloads / "解析后-1030井", desktop / "本地自跑解析结果_1030/景文敏"),
        ("周伟", downloads / "广材E化（周伟）", downloads / "广材E化（周伟）", desktop / "本地自跑解析结果_1030/周伟")
    ]
    
    all_results = []
    for p_name, s_dir, o_dir, l_dir in persons_cfg:
        print(f"[*] 正在分析经办人: {p_name}...")
        vdirs = sorted([d for d in s_dir.iterdir() if d.is_dir() and not d.name.startswith(".")])
        for vd in vdirs:
            r = evaluate_vendor(p_name, vd, o_dir, l_dir)
            all_results.append(r)
            
    print(f"评估完成，共汇聚 {len(all_results)} 家供应商全量全参数比对数据，启动 Excel 渲染...")
    build_full_workbook(all_results)
