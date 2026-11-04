# -*- coding: utf-8 -*-
"""
rebuild_usg_star_dual_sheets.py - 星牌优时吉建筑材料有限公司 双 Sheet 标准解析构建脚本
核心结构：
1. Sheet 1: 【客户原始解析】 - 原封不动放客户原本解析的 40 列内容与原始单元格数据
2. Sheet 2: 【标准化报价】 - 重新解析纠错版（规范分类 A123501 矿棉板、配图直链回填、底稿图内嵌与超链接、真实商务条款）
"""

import os
import io
import re
import zipfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage

from builder import StandardizedWorkbookBuilder, STANDARD_HEADERS
from models import CanonicalMaterialItem

DELIVERY_DIR = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries"
OUT_XLSX = os.path.join(DELIVERY_DIR, "星牌优时吉建筑材料有限公司-标准解析.xlsx")
IMG_DIR = os.path.join(DELIVERY_DIR, "星牌优时吉建筑材料有限公司_图片库")

ZIP_SOURCE = "/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案/2_张艳鸽（34家）/2_三方线上解析/星牌优时吉建筑材料有限公司1_20261030163511.zip"

PIC_URL = "https://gcj-pub.oss-cn-beijing.aliyuncs.com/opcrm/pic_pictures/m0000984/b4e39aeddb3ed18a.jpg"
FOLDER_NAME = "m0000984"

def parse_specs_from_raw(raw_spec: str, full_name: str) -> dict:
    """结构化解析规格与系列型号"""
    spec_dict = {}
    clean_s = re.sub(r'\s+', '', str(raw_spec or ''))
    m_dim = re.search(r'(\d+)\*(\d+)\*(\d+(?:\.\d+)?)mm?', clean_s)
    if m_dim:
        spec_dict["长度(mm)"] = m_dim.group(1)
        spec_dict["宽度(mm)"] = m_dim.group(2)
        spec_dict["厚度(mm)"] = m_dim.group(3)
        spec_dict["规格"] = f"{m_dim.group(1)}*{m_dim.group(2)}*{m_dim.group(3)}mm"
    else:
        m_th = re.search(r'(\d+(?:\.\d+)?)mm', clean_s)
        if m_th:
            spec_dict["厚度(mm)"] = m_th.group(1)
            
    series_name = full_name
    for prefix in ["矿物棉装饰吸声板", "矿棉板", "吸声板"]:
        series_name = series_name.replace(prefix, "")
    series_name = series_name.strip()
    if series_name:
        spec_dict["系列"] = series_name
        
    return spec_dict

def get_page_and_img_for_row(row_idx: int, cur_page, cur_img):
    """补齐换页处缺失的页码和图片名 (每页约 5~6 条)"""
    if cur_page and cur_img:
        return str(cur_page).strip(), str(cur_img).strip()
        
    mapping = {
        6: (1, "quote_img1.png"),
        12: (2, "quote_img2.png"),
        18: (3, "quote_img3.png"),
        23: (4, "quote_img4.png"),
        28: (5, "quote_img5.png"),
        34: (6, "quote_img6.png"),
        39: (7, "quote_img7.png"),
        45: (8, "quote_img8.png"),
        51: (9, "quote_img9.png"),
        57: (10, "quote_img10.png"),
        63: (11, "quote_img11.png")
    }
    if row_idx in mapping:
        p, img = mapping[row_idx]
        return str(p), img
        
    calc_p = min(12, max(1, (row_idx - 2) // 5 + 1))
    return str(calc_p), f"quote_img{calc_p}.png"

def run():
    print(f"[*] 读取客户原始解析压缩包: {ZIP_SOURCE}")
    with zipfile.ZipFile(ZIP_SOURCE, 'r') as z:
        excel_bytes = z.read("报价数据.xlsx")
        
    wb_client = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
    ws_client_src = wb_client.active
    
    headers_client = [ws_client_src.cell(1, c).value for c in range(1, ws_client_src.max_column + 1)]
    
    canonical_items = []
    supplier = "星牌优时吉建筑材料有限公司"
    brand = "星牌优时吉"
    id_prefix = "XPY"
    seq = 10
    
    terms_info = {
        'quote_desc': '当前提供的报价要求现款现结',
        'freight_desc': '13米整车：200km内含运；超200km，每100公里按产品加价1.5元，具体根据实际采购情况商定',
        'is_freight_included': '包含',
        'is_tax_included': '含税',
        'tax_rate': 13,
        'tax_desc': '增值税税率13%',
        'discount_ratio': 1.0,
        'discount_series': '全系列无折扣，具体以实际采购量为准',
        'quote_type': '出厂价'
    }

    print(f"[*] 解析客户数据并转换为标准 Canonical 模型 (共 {ws_client_src.max_row - 1} 条)...")
    for r in range(2, ws_client_src.max_row + 1):
        name = str(ws_client_src.cell(r, headers_client.index('名称规格') + 1).value or '').strip()
        raw_spec = str(ws_client_src.cell(r, headers_client.index('厂家原始规格参数') + 1).value or '').strip()
        price_val = ws_client_src.cell(r, headers_client.index('市场价') + 1).value
        try:
            f_price = float(price_val)
        except:
            f_price = 0.0
            
        unit_val = str(ws_client_src.cell(r, headers_client.index('单位') + 1).value or 'm²').strip()
        p_val = ws_client_src.cell(r, headers_client.index('厂家报价文件页码') + 1).value
        img_val = ws_client_src.cell(r, headers_client.index('厂家报价图片文件名') + 1).value
        
        page_str, img_fn = get_page_and_img_for_row(r, p_val, img_val)
        struct_specs = parse_specs_from_raw(raw_spec, name)
        
        item_id = f"{id_prefix}{seq:05d}"
        seq += 10
        
        item = CanonicalMaterialItem(
            quote_id=item_id,
            supplier=supplier,
            brand=brand,
            category_name="矿棉板",
            category_code="A123501",
            structured_specs=struct_specs,
            raw_spec=raw_spec,
            file_page=f"第{page_str}页",
            quote_image_filename=img_fn,
            image_filenames=[img_fn],
            unit=unit_val,
            market_price=f_price,
            project_price=f_price,
            discount_ratio=terms_info['discount_ratio'],
            discount_series=terms_info['discount_series'],
            is_freight_included=terms_info['is_freight_included'],
            freight_desc=terms_info['freight_desc'],
            is_tax_included=terms_info['is_tax_included'],
            tax_rate=terms_info['tax_rate'],
            tax_desc=terms_info['tax_desc'],
            quote_desc=terms_info['quote_desc'],
            quote_type=terms_info['quote_type'],
            province="全国",
            city="全国",
            source_ref=f"原单!R{r}",
            raw_price=f_price
        )
        item.picture_url = PIC_URL
        item.picture_urls = [PIC_URL]
        item.illustration_folder = FOLDER_NAME
        canonical_items.append(item)

    # 2. 创建全新的双 Sheet 工作簿
    print(f"[*] 初始化双 Sheet 工作簿...")
    wb_out = openpyxl.Workbook()
    
    # ----------------------------------------------------
    # Sheet 1: 【客户原始解析】 (按标准表头原封不动搬入)
    # ----------------------------------------------------
    ws_client_dst = wb_out.active
    ws_client_dst.title = "客户原始解析"
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    body_font = Font(name='宋体', size=10)
    alt_fill = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')
    white_fill = PatternFill(fill_type=None)
    
    print(f"[*] 写入 Sheet 1: 【客户原始解析】(原样 40 列)...")
    for r in range(1, ws_client_src.max_row + 1):
        ws_client_dst.row_dimensions[r].height = ws_client_src.row_dimensions[r].height or 22
        for c in range(1, ws_client_src.max_column + 1):
            src_val = ws_client_src.cell(r, c).value
            dst_cell = ws_client_dst.cell(r, c, src_val)
            
            if r == 1:
                dst_cell.font = Font(name='宋体', size=10, bold=True, color='FFFFFF')
                dst_cell.fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
                dst_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            else:
                dst_cell.font = body_font
                dst_cell.alignment = Alignment(horizontal='left', vertical='center')
                dst_cell.fill = alt_fill if (r % 2 == 0) else white_fill
            dst_cell.border = thin_border

    # ----------------------------------------------------
    # Sheet 2: 【标准化报价】 (规范分类、配图直链、底稿图片内嵌与超链接)
    # ----------------------------------------------------
    print(f"[*] 写入 Sheet 2: 【标准化报价】(规范 33 列企业级呈现)...")
    ws_standard = wb_out.create_sheet(title="标准化报价")
    
    builder = StandardizedWorkbookBuilder(None)
    std_headers = builder.headers
    
    # 写入标准表头
    ws_standard.row_dimensions[1].height = 36
    for c_idx, h in enumerate(std_headers, 1):
        cell = ws_standard.cell(row=1, column=c_idx, value=h)
        cell.font = Font(name='宋体', size=10, bold=True, color='FFFFFF')
        cell.fill = PatternFill(start_color='0B192C', end_color='0B192C', fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = thin_border
        
    os.makedirs(IMG_DIR, exist_ok=True)
    quote_img_col_idx = std_headers.index('厂家报价图片文件名') + 1

    total_imgs_embedded = 0
    center_align = Alignment(horizontal='center', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')
    multiline_left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    for row_idx, item in enumerate(canonical_items, start=2):
        row_fill = alt_fill if (row_idx % 2 == 0) else white_fill
        
        lines = []
        if item.structured_specs:
            for k, v in item.structured_specs.items():
                if v: lines.append(f"{k}:{v}")
        multiline_specs = "\n".join(lines) if lines else (item.raw_spec or '')
        line_count = max(1, multiline_specs.count('\n') + 1)
        
        row_vals = builder._row_values(
            item, row_idx - 1, multiline_specs,
            item.quote_image_filename, FOLDER_NAME, item.quote_desc
        )
        row_vals.update(builder._dynamic_spec_values(item, row_vals))
        
        img_col_counter = 1
        for c_idx, header in enumerate(std_headers, 1):
            if header == '图片':
                val = row_vals.get(f'_img{img_col_counter}', '')
                img_col_counter += 1
            elif header == '厂家报价图片文件名':
                val = row_vals.get('厂家报价图片文件名', '')
            else:
                val = row_vals.get(header, '')
                
            cell = ws_standard.cell(row=row_idx, column=c_idx, value=val)
            cell.font = body_font
            cell.fill = row_fill
            cell.border = thin_border
            
            if header in ['报价ID', '类别编码', '品牌', '厂家报价文件页码', '顺序号', '单位', '工程价折扣比例', '增值税税率', '是否含增值税', '是否含运费', '价格填报类型', '厂家报价图片文件名']:
                cell.alignment = center_align
            elif header in ['市场价', '工程价', '折扣下限', '折扣上限']:
                cell.alignment = right_align
                if isinstance(val, (int, float)):
                    cell.number_format = '0.00'
            else:
                cell.alignment = multiline_left_align
                
            # 图片直链超链接
            if header == '图片' and isinstance(val, str) and val.startswith('http'):
                cell.hyperlink = val
                cell.font = Font(name='宋体', size=10, color='0563C1', underline='single')

        # 内嵌底稿截图
        calc_height = max(38, line_count * 18 + 12)
        target_img_fn = item.quote_image_filename
        if target_img_fn and IMG_DIR:
            src_img = os.path.join(IMG_DIR, target_img_fn)
            if os.path.exists(src_img):
                try:
                    thumb_fn = f"thumb_retina_{target_img_fn}.jpg"
                    thumb_path = os.path.join(IMG_DIR, thumb_fn)
                    if not os.path.exists(thumb_path):
                        with PILImage.open(src_img) as pim:
                            pim.thumbnail((300, 180))
                            pim.convert("RGB").save(thumb_path, "JPEG", quality=85)
                            
                    cell_ref = ws_standard.cell(row=row_idx, column=quote_img_col_idx).coordinate
                    ws_standard.row_dimensions[row_idx].height = 42
                    
                    xl_img = OpenpyxlImage(thumb_path)
                    xl_img.width = 75
                    xl_img.height = 40
                    ws_standard.add_image(xl_img, cell_ref)
                    total_imgs_embedded += 1
                    
                    cell = ws_standard.cell(row=row_idx, column=quote_img_col_idx)
                    cell.value = target_img_fn
                    cell.hyperlink = f"星牌优时吉建筑材料有限公司_图片库/{target_img_fn}"
                    cell.font = Font(name='宋体', size=10, color='0563C1', underline='single')
                except Exception as e:
                    pass
        else:
            ws_standard.row_dimensions[row_idx].height = calc_height

    # 列宽适配
    for col in ws_client_dst.columns:
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        max_len = max(len(str(c.value or '')) for c in col[:15])
        ws_client_dst.column_dimensions[col_letter].width = max(12, min(max_len * 1.5 + 4, 35))

    for col in ws_standard.columns:
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws_standard.column_dimensions[col_letter].width = max(14, min(35, ws_standard.column_dimensions[col_letter].width or 16))

    print(f"[*] 保存双 Sheet 成果至: {OUT_XLSX}")
    wb_out.save(OUT_XLSX)
    print(f"[✓] 成功生成文件: {OUT_XLSX} (大小: {os.path.getsize(OUT_XLSX)} 字节)")

if __name__ == '__main__':
    run()
