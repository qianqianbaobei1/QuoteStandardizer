import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

# -*- coding: utf-8 -*-
"""
backfill_delivery_images.py - 给已生成的交付件回填材料配图（不发请求，只吃配图缓存）

背景：配图服务改为优先用“原件品名”检索后，品名过长/带型号的物料（筒灯、灯管、NLED9203SA 筒灯）
在配图接口里查不到图，覆盖率从 29.5% 掉到 2.9%。这里按“类别编码 → 标准类别名 → 配图缓存”
给空着的图片列补上（和 material_illustration_service 里新增的类别兜底逻辑同源）。
"""
import glob
import json
import os

import openpyxl
from openpyxl.styles import Font

from config import CATEGORY_CODE_MAP

BASE = os.path.dirname(os.path.abspath(__file__))
CACHE = os.path.join(BASE, "material_illustration_cache.json")
DELIVERIES = os.path.join(BASE, "deliveries")

code_to_name = {}
for _name, _code in CATEGORY_CODE_MAP.items():
    code_to_name.setdefault(_code, _name)


def main():
    with open(CACHE, encoding="utf-8") as f:
        cache = json.load(f)
    assert code_to_name.get("091709") == "挡烟垂壁", "类别编码反查表没建好"

    link_font = Font(name="宋体", size=10, color="0563C1", underline="single")
    files = sorted(glob.glob(os.path.join(DELIVERIES, "*-标准解析.xlsx")))
    total_empty = total_filled = 0
    for path in files:
        try:
            wb = openpyxl.load_workbook(path)
        except Exception as e:
            print(f"[skip] {os.path.basename(path)}: {e}")
            continue
        ws = wb["标准化报价"] if "标准化报价" in wb.sheetnames else wb.worksheets[0]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        if "图片" not in headers or "类别编码" not in headers:
            continue
        c_img = headers.index("图片") + 1
        c_folder = headers.index("图片文件夹") + 1 if "图片文件夹" in headers else None
        c_code = headers.index("类别编码") + 1
        filled = 0
        for r in range(2, ws.max_row + 1):
            if c_folder and ws.cell(r, c_folder).value:
                continue
            if ws.cell(r, c_img).value:
                continue
            total_empty += 1
            code = str(ws.cell(r, c_code).value or "").strip()
            hit = cache.get(code_to_name.get(code, "")) or {}
            url = hit.get("pictureUrl") or ""
            if not url:
                continue
            cell = ws.cell(r, c_img)
            cell.value = url
            cell.hyperlink = url
            cell.font = link_font
            if c_folder and hit.get("folderName"):
                ws.cell(r, c_folder).value = hit["folderName"]
            filled += 1
            total_filled += 1
        if filled:
            wb.save(path)
            print(f"[{os.path.basename(path)}] 回填配图 {filled} 行")
    print(f"\n合计：空图行 {total_empty}，回填 {total_filled}")


if __name__ == "__main__":
    main()
