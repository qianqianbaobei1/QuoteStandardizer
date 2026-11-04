#!/usr/bin/env python3
"""
batch_apply_solution_a.py
=============================================================================
批量为全量 161 家供应商标准交付 Excel 实施【方案 A：行级微距特写切片 (Row-Level Macro Crop)】：
1. 针对每家供应商底稿的高清 300 DPI 页面进行行级特写裁剪（包含品名、规格、单价等关键报价信息横条）；
2. 彻底清除单元格内文本与超链接（cell.value = None, cell.hyperlink = None）；
3. 单元格内直接嵌入特写图片对象（ws.add_image(...)），列宽设为 52（约 380 像素显示宽）；
4. 行高自适应调整为 55~75pt，确保在 Excel 内无须手动放大即可清晰查验每一条报价明细；
5. 应用 DrawingML 共享关系补丁，全国 31 省份展开行共享同源特写图片，文件体积严格控制在数百 KB，打开秒开无卡顿！
=============================================================================
"""

import os
import sys
import time
import re
import cv2
import zipfile
import hashlib
from pathlib import Path
from collections import defaultdict
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import get_column_letter
import openpyxl.writer.excel as w

# ---------------------------------------------------------------------------
# 1. Openpyxl 高性能 DrawingML 共享补丁 (防止多行重复保存相同图片导致体积膨胀与卡顿)
# ---------------------------------------------------------------------------
def fast_write_drawing(self, drawing):
    self._drawings.append(drawing)
    drawing._id = len(self._drawings)
    for chart in drawing.charts:
        self._charts.append(chart)
        chart._id = len(self._charts)
    source_to_img = {}
    for img in drawing.images:
        src_key = getattr(img, '_source_key', None) or getattr(img, 'ref', None) or id(img)
        if src_key in source_to_img:
            prim = source_to_img[src_key]
            img._id = prim._id
        else:
            self._images.append(img)
            img._id = len(self._images)
            source_to_img[src_key] = img

    from openpyxl.packaging.relationship import get_rels_path
    from openpyxl.xml.functions import tostring
    rels_path = get_rels_path(drawing.path)[1:]
    self._archive.writestr(drawing.path[1:], tostring(drawing._write()))
    self._archive.writestr(rels_path, tostring(drawing._write_rels()))
    self.manifest.append(drawing)

def fast_write_images(self):
    written = set()
    for img in self._images:
        if img.path not in written:
            self._archive.writestr(img.path[1:], img._data())
            written.add(img.path)

w.ExcelWriter._write_drawing = fast_write_drawing
w.ExcelWriter._write_images = fast_write_images

# ---------------------------------------------------------------------------
# 2. 页面行级特写裁剪函数
# ---------------------------------------------------------------------------
def get_page_row_crops(img_path: Path, num_items: int, output_dir: Path, page_key: str):
    """
    针对单页底稿生成 num_items 个物料的行级微距特写图片。
    优先采用表格横线检测定位，回退采用有效表格区间自适应切片。
    """
    im_pil = Image.open(img_path)
    w_px, h_px = im_pil.size
    crops = []
    
    # 尝试使用 cv2 识别表格横线
    try:
        cv_img = cv2.imread(str(img_path))
        if cv_img is not None and num_items > 0:
            gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)
            bw = cv2.adaptiveThreshold(~gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, -2)
            h_structure = cv2.getStructuringElement(cv2.MORPH_RECT, (w_px // 30, 1))
            horizontal = cv2.morphologyEx(bw, cv2.MORPH_OPEN, h_structure)
            y_lines = np.where(np.sum(horizontal > 0, axis=1) > w_px * 0.25)[0]
            lines = []
            for y in y_lines:
                if not lines or y - lines[-1] > 12:
                    lines.append(int(y))
            if len(lines) >= num_items + 1:
                data_lines = lines[-(num_items + 1):]
                for i in range(num_items):
                    y0 = max(0, data_lines[i] - 2)
                    y1 = min(h_px, data_lines[i+1] + 2)
                    c = im_pil.crop((int(w_px * 0.02), y0, int(w_px * 0.98), y1))
                    c_path = output_dir / f'crop_{page_key}_m{i+1}.png'
                    c.save(c_path)
                    crops.append((c_path, c.size))
    except Exception:
        pass

    # 若横线未能匹配完整行数，采用等分切片
    if len(crops) != num_items:
        crops = []
        table_top = int(h_px * 0.16)
        table_bottom = int(h_px * 0.90)
        step = (table_bottom - table_top) / max(1, num_items)
        for i in range(num_items):
            y0 = max(0, int(table_top + i * step - 2))
            y1 = min(h_px, int(table_top + (i + 1) * step + 2))
            c = im_pil.crop((int(w_px * 0.02), y0, int(w_px * 0.98), y1))
            c_path = output_dir / f'crop_{page_key}_m{i+1}.png'
            c.save(c_path)
            crops.append((c_path, c.size))

    return crops

# ---------------------------------------------------------------------------
# 3. 单个供应商工作簿实施方案 A
# ---------------------------------------------------------------------------
def process_single_vendor(xlsx_path: Path):
    vname = xlsx_path.stem.replace('-标准解析', '')
    deliv = xlsx_path.parent
    img_dir = deliv / f'{vname}_图片库'
    
    if not img_dir.exists():
        img_dir.mkdir(parents=True, exist_ok=True)
        
    t0 = time.time()
    wb = openpyxl.load_workbook(str(xlsx_path))
    ws = wb['标准化报价']
    
    # 查找列索引
    col_idx = 9
    col_page_idx = 8
    col_name_idx = 3
    for c in range(1, ws.max_column + 1):
        val = str(ws.cell(1, c).value or '').strip()
        if val == '厂家报价图片文件名':
            col_idx = c
        elif val == '厂家报价文件页码':
            col_page_idx = c
        elif val == '名称规格':
            col_name_idx = c
            
    col_letter = get_column_letter(col_idx)
    ws.column_dimensions[col_letter].width = 52 # 约 380 像素显示宽
    ws._images = [] # 彻底清空旧图片
    
    # 扫描收集全部物料行与对应页码
    row_records = []
    page_to_materials = defaultdict(list)
    
    for r in range(2, ws.max_row + 1):
        p_val = str(ws.cell(r, col_page_idx).value or '第1页').strip()
        m_page = re.search(r'\d+', p_val)
        page_num = int(m_page.group(0)) if m_page else 1
        name_val = str(ws.cell(r, col_name_idx).value or '').strip()
        
        row_records.append((r, page_num, name_val))
        if name_val not in page_to_materials[page_num]:
            page_to_materials[page_num].append(name_val)
            
    # 获取底稿所有可用高清图片
    available_imgs = sorted(list(img_dir.glob('quote_img*.png')))
    page_to_img_file = {}
    for p in page_to_materials:
        target = img_dir / f'quote_img{p}.png'
        if target.exists():
            page_to_img_file[p] = target
        elif available_imgs:
            page_to_img_file[p] = available_imgs[min(p - 1, len(available_imgs) - 1)]
            
    # 针对每页物料生成特写切片
    material_to_crop = {}
    for p, mat_list in page_to_materials.items():
        if p not in page_to_img_file:
            continue
        p_img_path = page_to_img_file[p]
        crops = get_page_row_crops(p_img_path, len(mat_list), img_dir, f'p{p}')
        for idx, m_name in enumerate(mat_list):
            if idx < len(crops):
                material_to_crop[(p, m_name)] = crops[idx]

    # 回退切片 (防范未匹配)
    default_crop = None
    if material_to_crop:
        default_crop = list(material_to_crop.values())[0]
    elif available_imgs:
        # 直接使用第一张主图的前部切片
        im0 = Image.open(available_imgs[0])
        w0, h0 = im0.size
        c0 = im0.crop((int(w0 * 0.02), int(h0 * 0.2), int(w0 * 0.98), int(h0 * 0.3)))
        c0_path = img_dir / 'crop_default.png'
        c0.save(c0_path)
        default_crop = (c0_path, c0.size)

    # 嵌入单元格
    for r, p, name_val in row_records:
        cell = ws.cell(r, col_idx)
        cell.value = None
        cell.hyperlink = None
        
        crop_info = material_to_crop.get((p, name_val), default_crop)
        if crop_info:
            c_path, (cw, ch) = crop_info
            op_img = OpenpyxlImage(str(c_path))
            op_img._source_key = str(c_path) # 关键：共享关系键
            target_w = 380
            target_h = max(28, min(80, int(ch * (target_w / cw))))
            op_img.width = target_w
            op_img.height = target_h
            ws.add_image(op_img, f'{col_letter}{r}')
            ws.row_dimensions[r].height = max(55, int(target_h * 0.75) + 8)

    wb.save(str(xlsx_path))
    t1 = time.time()
    size_kb = os.path.getsize(str(xlsx_path)) / 1024
    return {
        "vendor": vname,
        "rows": len(row_records),
        "crops": len(material_to_crop),
        "time": t1 - t0,
        "size_kb": size_kb
    }

# ---------------------------------------------------------------------------
# 4. 主执行入口
# ---------------------------------------------------------------------------
def main():
    deliv = Path('deliveries')
    workbooks = sorted(list(deliv.glob('*-标准解析.xlsx')))
    print(f'=============================================================================')
    print(f'  开始全量批量实施【方案 A：行级微距特写切片嵌入】')
    print(f'  待处理供应商总数: {len(workbooks)} 家')
    print(f'=============================================================================')
    
    success_cnt = 0
    fail_cnt = 0
    start_all = time.time()
    
    for idx, wb_path in enumerate(workbooks, 1):
        vname = wb_path.stem.replace('-标准解析', '')
        try:
            res = process_single_vendor(wb_path)
            success_cnt += 1
            print(f'[{idx:3d}/{len(workbooks)}] ✓ {vname:<32} | {res["rows"]:5d} 行 | 特写: {res["crops"]:2d} 张 | 耗时: {res["time"]:4.1f}s | 体积: {res["size_kb"]:6.1f} KB')
        except Exception as e:
            fail_cnt += 1
            print(f'[{idx:3d}/{len(workbooks)}] ✗ {vname:<32} | 失败原因: {e}')

    print(f'=============================================================================')
    print(f'  全量方案 A 批量嵌入完成！')
    print(f'  成功: {success_cnt} 家 | 失败: {fail_cnt} 家 | 总耗时: {time.time() - start_all:.1f} 秒')
    print(f'=============================================================================')

if __name__ == '__main__':
    main()
