# -*- coding: utf-8 -*-
"""
execute_clean_batch.py - 严格单 Sheet 纯净批量解析引擎
"""
import os
import re
import sys
import time
from pathlib import Path
import openpyxl

from run_batch_standardize import find_primary_quote_file, run_single_vendor
from zip_standardizer_pipeline import ZipQuoteStandardizerPipeline
from style_enhancer import apply_enterprise_style
from llm_auditor import LLMAuditor

# 源目录不再写死在代码里：命令行第一个参数 > 环境变量 QS_SOURCE_ROOT > 本机默认
DEFAULT_SRC_ROOT = '/Users/qianqianawodebaobei/Desktop/数据异化'
src_root = Path(sys.argv[1] if len(sys.argv) > 1 else os.getenv('QS_SOURCE_ROOT', DEFAULT_SRC_ROOT))
output_dir = Path('deliveries')
output_dir.mkdir(parents=True, exist_ok=True)

# 1. 扫描已完成名单
parsed_vendors = set()
for f in src_root.glob("*-标准化表.xlsx"):
    m = re.match(r"^([^\-_]+)", f.name)
    if m: parsed_vendors.add(m.group(1))

for parsed_dir_name in ['张艳鸽-解析文件', '程雅洁-10月30日解析文件']:
    pdir = src_root / parsed_dir_name
    if pdir.exists():
        for f in pdir.glob("*"):
            if f.is_file() and not f.name.startswith('.'):
                m = re.match(r"^([^\d_\-]+)", f.name)
                if m: parsed_vendors.add(m.group(1))

source_dirs = ['张艳鸽-源文件1030', '程雅洁-10月30日']
todo_tasks = []

for sdir_name in source_dirs:
    sdir = src_root / sdir_name
    if not sdir.exists(): continue
    for vpath in sorted(sdir.iterdir()):
        if not vpath.is_dir() or vpath.name.startswith('.'): continue
        vname = vpath.name
        v_clean = re.sub(r'[\(\)（）\s]', '', vname)
        already_done = any(v_clean in re.sub(r'[\(\)（）\s]', '', pv) or re.sub(r'[\(\)（）\s]', '', pv) in v_clean for pv in parsed_vendors)
        if not already_done:
            todo_tasks.append((sdir_name, vname, vpath))

print(f"==================== 正式批量解析启动 (待处理 {len(todo_tasks)} 家) ====================")

pipeline_zip = ZipQuoteStandardizerPipeline(str(output_dir))
auditor = LLMAuditor()

results = []

for idx, (batch_name, vendor_name, vendor_path) in enumerate(todo_tasks, 1):
    print(f"\n[{idx}/{len(todo_tasks)}] 正在解析: {vendor_name} ({batch_name})...")
    start_time = time.time()
    
    primary_file = find_primary_quote_file(vendor_path)
    if not primary_file:
        print(f"  [Error] 未找到主报价文件")
        results.append((vendor_name, "FAIL", "未找到主报价文件", 0))
        continue
        
    print(f"  -> 主报价文件: {primary_file.name}")
    ext = primary_file.suffix.lower()
    
    out_xlsx = None
    try:
        if ext == '.zip':
            out_xlsx = pipeline_zip.process_zip(str(primary_file), vendor_name)
        elif ext in ['.xlsx', '.xls', '.pdf', '.png', '.jpg', '.jpeg']:
            res_dict = run_single_vendor(vendor_path, output_dir)
            out_xlsx = res_dict.get('output_file')
            # 完整性校验未过就不得算 PASS：run_single_vendor 的 failures 是唯一通行证
            if res_dict.get('failures'):
                results.append((vendor_name, "FAIL",
                                f"完整性校验未通过 {len(res_dict['failures'])} 项: {res_dict['failures'][0]}", 0))
                continue
            
        if out_xlsx and os.path.exists(out_xlsx):
            # 严格后处理：保证单 Sheet，应用美学样式，自检
            wb = openpyxl.load_workbook(out_xlsx)
            # 只留标准化报价
            for s in list(wb.sheetnames):
                if s != '标准化报价' and s != wb.sheetnames[0]:
                    wb.remove(wb[s])
            ws = wb.worksheets[0]
            if ws.title != '标准化报价':
                ws.title = '标准化报价'
                
            # 【重要】强制大模型深度业务逻辑终审与就地自愈回填
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            auditor.verify_and_auto_heal(ws, headers)
            
            # 美学样式刷新
            apply_enterprise_style(ws)
            wb.save(out_xlsx)
            
            # 统计
            total_items = ws.max_row - 1
            duration = round(time.time() - start_time, 2)
            print(f"  [✓ PASS] 成功生成: {out_xlsx} (单Sheet, 共 {total_items} 行, 耗时 {duration}s)")
            results.append((vendor_name, "PASS", out_xlsx, total_items))
        else:
            results.append((vendor_name, "FAIL", "未生成有效Excel", 0))
    except Exception as e:
        print(f"  [✗ EXCEPTION] 处理失败: {e}")
        results.append((vendor_name, "FAIL", str(e), 0))

print("\n==================== 最终批量完成汇总 ====================")
for vname, status, info, cnt in results:
    print(f"  {vname}: {status} | 行数: {cnt} | {info}")
