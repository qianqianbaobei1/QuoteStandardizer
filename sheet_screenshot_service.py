# -*- coding: utf-8 -*-
"""
sheet_screenshot_service.py - 报价底稿原件全景超清截图与溯源渲染引擎 (300DPI/3000px Ultra-HD)
核心升级：
1. PDF 原件：采用 PyMuPDF (fitz) 以 300 DPI 印刷级分辨率全幅渲染，文字与公章纤毫毕现 (quote_img{页码}.png)
2. Excel 原件：采用 macOS 原生矢量级 QuickLook 引擎以 3000px 超清规格栅格化输出 (quote_img_{sheet名}.png)
3. 视网膜双轨机制：同时生成 3000px/300DPI 原始母图 与 800px 视网膜抗锯齿微缩图 (thumb_retina_*.jpg)
4. 输出映射字典：提供 {页码/Sheet名: 截图文件名}，供 builder 和流水线准确填入「厂家报价图片文件名」及穿透超链接
"""
import os
import re
import shutil
import tempfile
import subprocess
from typing import Dict, Any, List
import openpyxl
from PIL import Image, ImageDraw, ImageFont

from config import SKIP_HIDDEN_CONTENT

def safe_filename(name: str) -> str:
    """清理文件名中的非法字符"""
    s = re.sub(r'[\/\\:\*\?\"<>\|]', '_', str(name or '')).strip()
    return s or 'sheet'

def generate_retina_thumbnail(master_img_path: str, output_thumb_path: str, target_width: int = 800) -> bool:
    """
    为 3000px 巨幅母图生成 800px 视网膜级抗锯齿微缩图，
    专供 Excel 单元格内嵌使用，既不失真模糊，又保障 Excel 运行轻快。
    """
    if not os.path.exists(master_img_path) or os.path.getsize(master_img_path) == 0:
        return False
    try:
        with Image.open(master_img_path) as im:
            w, h = im.size
            if w <= target_width:
                shutil.copy2(master_img_path, output_thumb_path)
                return True
            target_h = int(h * (target_width / w))
            im_rgb = im.convert('RGB')
            # 使用高质量 Lanczos 滤波器抗锯齿重采样
            im_resized = im_rgb.resize((target_width, target_h), Image.Resampling.LANCZOS)
            im_resized.save(output_thumb_path, 'JPEG', quality=95, optimize=True)
            return True
    except Exception as e:
        print(f"[SheetScreenshot] 生成视网膜微缩图失败 {master_img_path}: {e}")
        return False

def render_fallback_sheet_image(ws, target_path: str, max_r: int = 100, max_c: int = 20):
    """
    当系统原生渲染器不可用时的高清网格绘制兜底引擎 (2400px 超清排版)
    """
    try:
        cell_w, cell_h = 160, 40
        img_w = min(max_c * cell_w, 2400)
        img_h = max_r * cell_h + 80
        img = Image.new("RGB", (img_w, img_h), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)

        # 尝试加载中文字体
        font = None
        for font_path in [
            "/System/Library/Fonts/PingFang.ttc",
            "/Library/Fonts/Songti.ttc",
            "/System/Library/Fonts/STHeiti Light.ttc",
            "/Library/Fonts/Arial Unicode.ttf"
        ]:
            if os.path.exists(font_path):
                try:
                    font = ImageFont.truetype(font_path, 16)
                    break
                except Exception:
                    continue
        if font is None:
            font = ImageFont.load_default()

        # 绘制表头与网格
        for r in range(1, max_r + 1):
            for c in range(1, max_c + 1):
                val = str(ws.cell(r, c).value or '')
                x1 = (c - 1) * cell_w
                y1 = (r - 1) * cell_h + 40
                x2 = x1 + cell_w
                y2 = y1 + cell_h
                fill_color = (240, 244, 248) if r == 1 else (255, 255, 255)
                draw.rectangle([x1, y1, x2, y2], fill=fill_color, outline=(210, 215, 220))
                if val:
                    draw.text((x1 + 6, y1 + 10), val[:14], fill=(0, 0, 0), font=font)

        img.save(target_path, "PNG")
        return True
    except Exception as e:
        print(f"[SheetScreenshot] 绘制备用表格失败: {e}")
        return False

def capture_excel_sheets(excel_path: str, output_dir: str) -> Dict[str, str]:
    """
    为 Excel 报价单的每一个 Sheet 生成 3000px 超清全景底稿原图，
    保存在 output_dir 中，返回 {sheet_name: quote_img_xxx.png} 映射。
    """
    os.makedirs(output_dir, exist_ok=True)
    if not os.path.exists(excel_path):
        return {}

    mapping = {}
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception as e:
        print(f"[SheetScreenshot] 无法加载工作簿 {excel_path}: {e}")
        return {}

    for idx, sname in enumerate(wb.sheetnames, 1):
        ws = wb[sname]
        if SKIP_HIDDEN_CONTENT and ws.sheet_state != 'visible':
            continue    # 隐藏口径: 隐藏工作表不产出底稿图, 也不进映射
        safe_s = safe_filename(sname)
        target_fn = f"quote_img_{safe_s}.png"
        target_path = os.path.join(output_dir, target_fn)
        thumb_path = os.path.join(output_dir, f"thumb_retina_{safe_s}.jpg")

        # 若已存在且尺寸大于 50KB 则视为有效超清图复用
        if os.path.exists(target_path) and os.path.getsize(target_path) > 50 * 1024:
            mapping[sname] = target_fn
            if not os.path.exists(thumb_path):
                generate_retina_thumbnail(target_path, thumb_path)
            continue

        max_r = min(ws.max_row or 1, 300)
        max_c = min(ws.max_column or 1, 35)

        with tempfile.TemporaryDirectory() as tmpdir:
            temp_wb = openpyxl.Workbook()
            temp_ws = temp_wb.active
            temp_ws.title = str(sname)[:31]
            temp_ws.views.sheetView[0].showGridLines = True

            for r in range(1, max_r + 1):
                for c in range(1, max_c + 1):
                    temp_ws.cell(r, c, ws.cell(r, c).value)

            temp_file = os.path.join(tmpdir, f"sheet_{idx}.xlsx")
            temp_wb.save(temp_file)

            success = False
            # 优先调用 macOS 原生 QuickLook 以 3000px 渲染
            try:
                res = subprocess.run(
                    ['qlmanage', '-t', '-s', '3000', '-o', tmpdir, temp_file],
                    capture_output=True, text=True, timeout=20
                )
                if res.returncode == 0:
                    pngs = [f for f in os.listdir(tmpdir) if f.endswith('.png')]
                    if pngs:
                        src_png = os.path.join(tmpdir, pngs[0])
                        shutil.copy2(src_png, target_path)
                        generate_retina_thumbnail(target_path, thumb_path)
                        mapping[sname] = target_fn
                        success = True
            except Exception as e:
                print(f"[SheetScreenshot] qlmanage 渲染异常 [{sname}]: {e}")

            # 兜底：若 qlmanage 失败，使用高清晰度网格引擎绘制
            if not success:
                render_fallback_sheet_image(ws, target_path, max_r=max_r, max_c=max_c)
                generate_retina_thumbnail(target_path, thumb_path)
                mapping[sname] = target_fn

    return mapping

def capture_pdf_pages(pdf_path: str, output_dir: str, dpi: int = 300) -> Dict[str, str]:
    """
    为 PDF 报价单的每一页以 300 DPI 印刷级分辨率生成超清底稿截图与视网膜微缩图，
    返回 {"第1页": "quote_img1.png", "1": "quote_img1.png", 1: "quote_img1.png"}
    """
    os.makedirs(output_dir, exist_ok=True)
    if not os.path.exists(pdf_path):
        return {}

    mapping = {}
    try:
        import fitz
        with fitz.open(pdf_path) as doc:
            for idx, page in enumerate(doc, 1):
                fn = f"quote_img{idx}.png"
                dst = os.path.join(output_dir, fn)
                thumb_fn = f"thumb_retina_quote_img{idx}.jpg"
                thumb_dst = os.path.join(output_dir, thumb_fn)

                if not os.path.exists(dst) or os.path.getsize(dst) < 50 * 1024:
                    pix = page.get_pixmap(dpi=dpi)
                    pix.save(dst)
                if not os.path.exists(thumb_dst):
                    generate_retina_thumbnail(dst, thumb_dst)

                mapping[f"第{idx}页"] = fn
                mapping[str(idx)] = fn
                mapping[idx] = fn
    except Exception as e:
        print(f"[SheetScreenshot] PDF 300DPI 截图生成失败 {pdf_path}: {e}")
    return mapping

def ensure_quote_screenshots(source_path: str, output_dir: str) -> Dict[Any, str]:
    """
    通用入口：自动根据源文件类型（Excel / PDF / 单张图片 / 压缩包）
    生成全量 3000px/300DPI 原件底稿截图与视网膜微缩图，并返回页码/Sheet -> 截图文件名映射。
    """
    if not source_path or not os.path.exists(source_path):
        return {}

    ext = os.path.splitext(source_path)[1].lower()
    if ext in ('.xlsx', '.xls'):
        return capture_excel_sheets(source_path, output_dir)
    elif ext == '.pdf':
        return capture_pdf_pages(source_path, output_dir, dpi=300)
    elif ext in ('.png', '.jpg', '.jpeg', '.bmp'):
        os.makedirs(output_dir, exist_ok=True)
        fn = "quote_img1.png"
        dst = os.path.join(output_dir, fn)
        thumb_dst = os.path.join(output_dir, "thumb_retina_quote_img1.jpg")
        shutil.copy2(source_path, dst)
        generate_retina_thumbnail(dst, thumb_dst)
        return {"1": fn, "第1页": fn, 1: fn, os.path.basename(source_path): fn}
    return {}
