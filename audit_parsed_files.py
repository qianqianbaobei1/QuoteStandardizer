import os, io, zipfile
import pandas as pd
from pathlib import Path

desktop = Path("/Users/qianqianawodebaobei/Desktop")
batches = [
    ("程雅洁", desktop / "程雅洁-10月30日解析文件"),
    ("张艳鸽", desktop / "张艳鸽-解析文件")
]

for person, root in batches:
    print(f"\n==================== {person} 解析结果盘点 ====================")
    files = sorted([f for f in root.iterdir() if not f.name.startswith(".")])
    print(f"总解析文件数: {len(files)}")
    
    total_rows = 0
    zero_row_files = []
    
    for f in files:
        df = None
        if f.suffix == ".xlsx":
            try:
                df = pd.read_excel(f)
            except Exception as e:
                print(f"读取 Excel 失败 {f.name}: {e}")
        elif f.suffix == ".zip":
            try:
                with zipfile.ZipFile(f) as z:
                    for name in z.namelist():
                        if name.endswith(".xlsx"):
                            df = pd.read_excel(io.BytesIO(z.read(name)))
                            break
            except Exception as e:
                print(f"读取 Zip 失败 {f.name}: {e}")
        
        if df is not None:
            rows = len(df)
            total_rows += rows
            if rows == 0:
                zero_row_files.append(f.name)
            print(f"  [{f.name[:30]}] -> {rows} 行, 列数: {len(df.columns)}")
        else:
            print(f"  [{f.name[:30]}] -> 未找到/未能解析出 Excel")
            zero_row_files.append(f.name)
            
    print(f"总结: 总行数={total_rows}, 0行/解析失败文件数={len(zero_row_files)}")

