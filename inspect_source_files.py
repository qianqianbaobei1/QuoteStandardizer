import os, sys, zipfile, io
from pathlib import Path
import pandas as pd
import pypdf

desktop = Path("/Users/qianqianawodebaobei/Desktop")
batches = [
    ("程雅洁", desktop / "程雅洁-10月30日"),
    ("张艳鸽", desktop / "张艳鸽-源文件1030")
]

def find_quote_files(vendor_dir):
    """动态识别供应商目录下的核心报价单文件，排除报价说明和折扣记录"""
    candidates = []
    for p in vendor_dir.rglob("*"):
        if not p.is_file() or p.name.startswith(".") or "__MACOSX" in str(p):
            continue
        name_lower = p.name.lower()
        # 排除报价说明、折扣记录、折扣说明
        if any(k in p.name for k in ["报价说明", "折扣说明", "折扣记录", "说明"]):
            continue
        # 查找包含报价、清单、明细、价格、或者核心数据格式
        if p.suffix.lower() in [".xlsx", ".xls", ".pdf", ".png", ".jpg", ".jpeg", ".zip"]:
            candidates.append(p)
    return candidates

for person, root in batches:
    print(f"\n==================== {person} 核心报价单定位 ====================")
    for vdir in sorted(root.iterdir()):
        if not vdir.is_dir() or vdir.name.startswith("."):
            continue
        cands = find_quote_files(vdir)
        c_str = ", ".join([f"{c.name} ({c.suffix}, {c.stat().st_size}B)" for c in cands])
        print(f"[{vdir.name}] -> {len(cands)} 个核心文件: {c_str}")

