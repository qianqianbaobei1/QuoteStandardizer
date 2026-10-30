import os, sys, re, io, zipfile
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

def load_parsed(fpath):
    if not fpath: return None
    if fpath.suffix == ".xlsx":
        try:
            return pd.read_excel(fpath)
        except Exception as e:
            return None
    elif fpath.suffix == ".zip":
        try:
            with zipfile.ZipFile(fpath) as z:
                for name in z.namelist():
                    if name.endswith(".xlsx"):
                        return pd.read_excel(io.BytesIO(z.read(name)))
        except Exception as e:
            return None
    return None

def analyze_vendor(person, vdir, parsed_files):
    vname = vdir.name
    # 查找匹配的解析文件
    parsed_f = None
    for pf in parsed_files:
        if pf.name.startswith(vname):
            parsed_f = pf
            break
            
    # 找源报价单
    src_cands = []
    terms_cands = []
    for p in vdir.rglob("*"):
        if not p.is_file() or p.name.startswith(".") or "__MACOSX" in str(p):
            continue
        if any(k in p.name for k in ["报价说明", "折扣说明", "折扣记录", "说明"]):
            terms_cands.append(p)
        elif p.suffix.lower() in [".xlsx", ".xls", ".pdf", ".png", ".jpg", ".jpeg", ".zip"]:
            src_cands.append(p)
            
    parsed_df = load_parsed(parsed_f) if parsed_f else None
    parsed_rows = len(parsed_df) if parsed_df is not None else 0
    parsed_cols = len(parsed_df.columns) if parsed_df is not None else 0
    
    # 探查源文件详细信息
    src_details = []
    estimated_src_rows = None
    
    for sc in src_cands:
        ext = sc.suffix.lower()
        if ext == ".xlsx":
            try:
                xl = pd.ExcelFile(sc)
                rows_in_sheets = []
                for s in xl.sheet_names:
                    sdf = xl.parse(s, header=None)
                    # 粗略估计有效物料行：至少有数值价格的行
                    val_rows = 0
                    for _, r in sdf.iterrows():
                        # 判断是否含有价格类的数字
                        has_num = any(re.search(r'\d+(\.\d+)?', str(x)) for x in r if pd.notna(x))
                        if has_num:
                            val_rows += 1
                    rows_in_sheets.append(f"{s}:{len(sdf)}行(估物料~{val_rows})")
                    if estimated_src_rows is None:
                        estimated_src_rows = val_rows
                    else:
                        estimated_src_rows = max(estimated_src_rows, val_rows)
                src_details.append(f"Excel[{sc.name}]({', '.join(rows_in_sheets)})")
            except Exception as e:
                src_details.append(f"Excel[{sc.name}](read err: {e})")
        elif ext == ".pdf":
            try:
                reader = pypdf.PdfReader(str(sc))
                pages = len(reader.pages)
                src_details.append(f"PDF[{sc.name}]({pages}页)")
            except Exception as e:
                src_details.append(f"PDF[{sc.name}](err: {e})")
        elif ext in [".png", ".jpg", ".jpeg"]:
            src_details.append(f"IMG[{sc.name}]({sc.stat().st_size}B)")
        elif ext == ".zip":
            src_details.append(f"ZIP[{sc.name}]({sc.stat().st_size}B)")
            
    terms_info = [t.name for t in terms_cands]
    
    return {
        "person": person,
        "vendor": vname,
        "has_parsed": parsed_f is not None,
        "parsed_file": parsed_f.name if parsed_f else None,
        "parsed_rows": parsed_rows,
        "parsed_cols": parsed_cols,
        "src_files": "; ".join(src_details),
        "terms_files": "; ".join(terms_info),
        "parsed_df": parsed_df
    }

results = []
for b in batches:
    parsed_files = [f for f in b["parsed_root"].iterdir() if not f.name.startswith(".")]
    for vdir in sorted(b["src_root"].iterdir()):
        if not vdir.is_dir() or vdir.name.startswith("."): continue
        res = analyze_vendor(b["person"], vdir, parsed_files)
        results.append(res)

print(f"Total analyzed vendors: {len(results)}")
for r in results[:10]:
    print(f"[{r['person']}] {r['vendor']}: 线上解析={r['has_parsed']} ({r['parsed_rows']}行) | 源文件: {r['src_files'][:50]}")

