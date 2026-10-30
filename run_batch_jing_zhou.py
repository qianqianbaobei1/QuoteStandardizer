# -*- coding: utf-8 -*-
"""
run_batch_jing_zhou.py
针对下载文件夹中【景文敏】与【周伟】两批原文件执行本地端到端解析落盘
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

jing_src = downloads / "原文件-1030井"
zhou_dir = downloads / "广材E化（周伟）"

out_jing = desktop / "本地自跑解析结果_1030/景文敏"
out_zhou = desktop / "本地自跑解析结果_1030/周伟"

out_jing.mkdir(parents=True, exist_ok=True)
out_zhou.mkdir(parents=True, exist_ok=True)

def run_batch_jing():
    print("\n==================== 开始批量自跑: 景文敏 ====================")
    vdirs = sorted([d for d in jing_src.iterdir() if d.is_dir() and not d.name.startswith(".")])
    for idx, vd in enumerate(vdirs, 1):
        print(f"\n[景文敏 {idx}/{len(vdirs)}] 处理: {vd.name}...")
        r = run_single_vendor(vd, out_jing)
        print(f"    -> 结果: {r['status']} | 物料: {r['item_count']} 行 | 耗时: {r['duration']}s | 文件: {r['output_file']}")

def run_batch_zhou():
    print("\n==================== 开始批量自跑: 周伟 ====================")
    vdirs = sorted([d for d in zhou_dir.iterdir() if d.is_dir() and not d.name.startswith(".")])
    for idx, vd in enumerate(vdirs, 1):
        # 排除空目录
        raw_files = [f for f in vd.iterdir() if f.is_file() and not f.name.startswith(".") and not re.search(r"_\d{14}\.zip$", f.name)]
        if not raw_files:
            print(f"\n[周伟 {idx}/{len(vdirs)}] 跳过空目录: {vd.name}")
            continue
        print(f"\n[周伟 {idx}/{len(vdirs)}] 处理: {vd.name}...")
        r = run_single_vendor(vd, out_zhou)
        print(f"    -> 结果: {r['status']} | 物料: {r['item_count']} 行 | 耗时: {r['duration']}s | 文件: {r['output_file']}")

if __name__ == "__main__":
    run_batch_jing()
    run_batch_zhou()

