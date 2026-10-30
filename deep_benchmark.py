import os, sys, re, io, zipfile, json
from pathlib import Path
import pandas as pd
import pypdf

desktop = Path("/Users/qianqianawodebaobei/Desktop")

batches = [
    {
        "person": "程雅洁",
        "src_root": desktop / "程雅洁-10月30日",
        "parsed_root": desktop / "程雅洁-10月30日解析文件"
    },
    {
        "person": "张艳鸽",
        "src_root": desktop / "张艳鸽-源文件1030",
        "parsed_root": desktop / "张艳鸽-解析文件"
    }
]

def load_parsed_df(parsed_path):
    if not parsed_path or not parsed_path.exists():
        return None
    if parsed_path.suffix == ".xlsx":
        try:
            return pd.read_excel(parsed_path)
        except Exception as e:
            return None
    elif parsed_path.suffix == ".zip":
        try:
            with zipfile.ZipFile(parsed_path) as z:
                for name in z.namelist():
                    if name.endswith(".xlsx"):
                        return pd.read_excel(io.BytesIO(z.read(name)))
        except Exception as e:
            return None
    return None

def find_matched_parsed_file(vendor_name, parsed_files):
    for f in parsed_files:
        if f.name.startswith(vendor_name):
            return f
    # 模糊匹配：核心商号匹配
    for f in parsed_files:
        if any(token in f.name for token in [vendor_name[:6], vendor_name[:8]] if len(token) >= 4):
            return f
    return None

def find_source_quote_files(vendor_dir):
    cands = []
    for p in vendor_dir.rglob("*"):
        if not p.is_file() or p.name.startswith(".") or "__MACOSX" in str(p):
            continue
        if any(k in p.name for k in ["报价说明", "折扣说明", "折扣记录", "说明"]):
            continue
        ext = p.suffix.lower()
        if ext in [".xlsx", ".xls", ".pdf", ".png", ".jpg", ".jpeg", ".zip"]:
            cands.append(p)
    return cands

def find_terms_files(vendor_dir):
    terms = []
    for p in vendor_dir.rglob("*"):
        if not p.is_file() or p.name.startswith(".") or "__MACOSX" in str(p):
            continue
        if any(k in p.name for k in ["报价说明", "折扣说明", "折扣记录", "说明"]):
            terms.append(p)
    return terms

print("Benchmark preparation loaded successfully.")

