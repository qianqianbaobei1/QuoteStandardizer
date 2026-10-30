# -*- coding: utf-8 -*-
"""
run_zhou.py - 并发独立执行【周伟】批次原文件端到端解析落盘
"""
import os
import sys
import time
import re
from pathlib import Path

from main import process_excel, process_pdf
from terms_extractor import find_terms_image, TermsExtractor
from config import DEFAULT_TEMPLATE_PATH
from run_batch_standardize import extract_zip_if_needed, find_primary_quote_file, run_single_vendor

downloads = Path("/Users/qianqianawodebaobei/Downloads")
desktop = Path("/Users/qianqianawodebaobei/Desktop")
zhou_dir = downloads / "广材E化（周伟）"
out_zhou = desktop / "本地自跑解析结果_1030/周伟"
out_zhou.mkdir(parents=True, exist_ok=True)

def run_batch_zhou():
    print("\n==================== 开始独立并发自跑: 周伟 ====================")
    vdirs = sorted([d for d in zhou_dir.iterdir() if d.is_dir() and not d.name.startswith(".")])
    for idx, vd in enumerate(vdirs, 1):
        raw_files = [f for f in vd.iterdir() if f.is_file() and not f.name.startswith(".") and not re.search(r"_\d{14}\.zip$", f.name)]
        if not raw_files:
            print(f"\n[周伟 {idx}/{len(vdirs)}] 跳过空目录: {vd.name}")
            continue
            
        vname = vd.name
        clean_vname = re.sub(r"^\d+", "", vname)
        existing = list(out_zhou.glob(f"*{clean_vname}*-标准解析.xlsx"))
        if not existing:
            existing = list(out_zhou.glob(f"*{clean_vname[:6]}*-标准解析.xlsx" if len(clean_vname)>=6 else "____"))
        if existing and existing[0].exists() and existing[0].stat().st_size > 1000:
            print(f"\n[周伟 {idx}/{len(vdirs)}] 已存在，跳过: {existing[0].name}")
            continue

        print(f"\n[周伟 {idx}/{len(vdirs)}] 处理: {vd.name}...")
        r = run_single_vendor(vd, out_zhou)
        print(f"    -> 结果: {r['status']} | 物料: {r['item_count']} 行 | 耗时: {r['duration']}s | 文件: {r['output_file']}")

if __name__ == "__main__":
    run_batch_zhou()

