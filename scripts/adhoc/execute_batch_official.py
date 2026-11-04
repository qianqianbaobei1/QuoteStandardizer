# -*- coding: utf-8 -*-
"""
execute_batch_official.py - 官方全量多模态自动化批量解析与自我核对执行脚本
"""
import _adhoc_paths  # noqa: F401  把项目根目录加入 sys.path
import os
import re
import sys
import time
import zipfile
from pathlib import Path
import openpyxl

from run_batch_standardize import find_primary_quote_file, run_single_vendor, verify_delivery_file
from zip_standardizer_pipeline import ZipQuoteStandardizerPipeline
from style_enhancer import apply_enterprise_style
from llm_auditor import LLMAuditor

# 源目录参数化：命令行第一个参数 > 环境变量 QS_SOURCE_ROOT
src_root = Path(sys.argv[1] if len(sys.argv) > 1 else os.getenv('QS_SOURCE_ROOT', ''))
if not str(src_root).strip() or str(src_root) in ('.', './') or not src_root.is_dir():
    print('用法: python3 execute_batch_official.py <源目录>   (或设置 QS_SOURCE_ROOT)')
    raise SystemExit(1)
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

print(f"==================== 开始正式批量解析 (共 {len(todo_tasks)} 家) ====================")

pipeline_zip = ZipQuoteStandardizerPipeline(str(output_dir))
auditor = LLMAuditor()

results = []

for idx, (batch_name, vendor_name, vendor_path) in enumerate(todo_tasks, 1):
    print(f"\n[{idx}/{len(todo_tasks)}] 正在处理: {vendor_name} ({batch_name})...")
    start_time = time.time()
    
    # 查找主文件
    primary_file = find_primary_quote_file(vendor_path)
    if not primary_file:
        print(f"  [Error] 未找到主报价文件")
        results.append((vendor_name, "FAIL", "未找到主报价文件", 0))
        continue
        
    print(f"  -> 主报价文件: {primary_file.name}")
    ext = primary_file.suffix.lower()
    
    out_xlsx = None
    try:
        # 如果是 ZIP 压缩包（可能含图片与多表数据）
        if ext == '.zip':
            out_xlsx = pipeline_zip.process_zip(str(primary_file), vendor_name)
        elif ext in ['.xlsx', '.xls', '.pdf', '.png', '.jpg', '.jpeg']:
            res_dict = run_single_vendor(vendor_path, output_dir)
            out_xlsx = res_dict.get('output_file')
            
        # ZIP 等不经过 main 校验的链路，也要独立体检后才算产出
        if out_xlsx and os.path.exists(out_xlsx):
            delivery_errs = verify_delivery_file(out_xlsx)
            if delivery_errs:
                results.append((vendor_name, "FAIL",
                                f"交付件体检未通过 {len(delivery_errs)} 项: {delivery_errs[0]}", 0))
                continue
            # 执行商业级投行排版与大模型自我核对
            wb = openpyxl.load_workbook(out_xlsx)
            for ws in wb.worksheets:
                auditor.verify_and_auto_heal(ws, [ws.cell(1, c).value for c in range(1, ws.max_column+1)])
                apply_enterprise_style(ws)
            wb.save(out_xlsx)
            
            # 计算行数
            wb_check = openpyxl.load_workbook(out_xlsx, data_only=True)
            total_items = sum(ws.max_row - 1 for ws in wb_check.worksheets)
            duration = round(time.time() - start_time, 2)
            print(f"  [✓ PASS] 成功生成: {out_xlsx} (共 {total_items} 行, 耗时 {duration}s)")
            results.append((vendor_name, "PASS", out_xlsx, total_items))
        else:
            results.append((vendor_name, "FAIL", "未生成有效Excel", 0))
    except Exception as e:
        print(f"  [✗ EXCEPTION] 处理失败: {e}")
        results.append((vendor_name, "FAIL", str(e), 0))

print("\n==================== 最终批量完成汇总 ====================")
for vname, status, info, cnt in results:
    print(f"  {vname}: {status} | 行数: {cnt} | {info}")
