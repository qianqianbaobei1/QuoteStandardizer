import os
import zipfile
import pandas as pd
from pathlib import Path

desktop = Path("/Users/qianqianawodebaobei/Desktop")
folders = {
    "程雅洁_源": desktop / "程雅洁-10月30日",
    "程雅洁_解析": desktop / "程雅洁-10月30日解析文件",
    "张艳鸽_源": desktop / "张艳鸽-源文件1030",
    "张艳鸽_解析": desktop / "张艳鸽-解析文件"
}

for name, p in folders.items():
    print(f"=== {name}: {p} (exists={p.exists()}) ===")
    if not p.exists():
        continue
    items = sorted([x for x in p.iterdir() if not x.name.startswith(".")])
    print(f"Total count: {len(items)}")
    for item in items[:5]:
        print(f"  - {item.name} ({'dir' if item.is_dir() else 'file'}, size={item.stat().st_size if item.is_file() else 'N/A'})")
    if len(items) > 5:
        print(f"  ... and {len(items)-5} more")

print("\n--- Inspecting sample parsed files from 程雅洁_解析 ---")
for f in (folders["程雅洁_解析"]).glob("*"):
    if f.name.startswith("."): continue
    if f.suffix == ".zip":
        with zipfile.ZipFile(f, 'r') as z:
            namelist = z.namelist()
            print(f"Zip: {f.name} -> {namelist}")
    elif f.suffix == ".xlsx":
        print(f"Excel: {f.name}")
    break


