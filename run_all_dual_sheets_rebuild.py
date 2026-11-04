# -*- coding: utf-8 -*-
"""
run_all_dual_sheets_rebuild.py
全量 161 家供应商双 Sheet 标准交付重构主引擎
================================================================================
业务规范：
1. Sheet 1: 【客户原始解析】
   - 原汁原味承载客户系统原始解析内容（标准 40 列表头或原始解析表头，原封不动保留原始单元格数据与行数）。
   - 绝不篡改客户原单，原样呈现。

2. Sheet 2: 【标准化报价】
   - 呈现企业级标准化重构结果（标准 33 列表头）。
   - 彻底解决全国 31 省份复制爆炸导致行数虚高的问题，严格按 1:1 真实物料条数呈现。
   - 规范工料机分类编码（如 A123501 矿棉板）与标准分类名称。
   - 规范品牌名称，杜绝截断。
   - 回填工料机中台材料配图接口真实直链（带有可点击超链接）与图片文件夹名（如 m0000984）。
   - 关联底稿截图（quote_img*.png），在【厂家报价图片文件名】列内嵌 Retina 缩略图并生成相对穿透超链接。
   - 补齐真实商务条款（报价说明、工程折扣比例与系列、是否含税、税率、运费约定、价格填报类型）。
   - 企业级投行/咨询美学排版（Navy 经典深蓝、宋体规范、会计千分位、斑马纹衬底）。
================================================================================
"""

import os
import sys
import io
import re
import json
import glob
import time
import zipfile
import traceback
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage

from builder import StandardizedWorkbookBuilder, STANDARD_HEADERS
from material_illustration_service import default_illustration_service

# ----------------- 路径常量配置 -----------------
PROJECT_DIR = Path('/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer')
DELIVERIES_DIR = PROJECT_DIR / 'deliveries'

FOUR_AGENTS_DIR = Path('/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案')
ZHOUWEI_1103_DIR = PROJECT_DIR / 'deliveries_backup_20261103_1828/周伟E化11.03_unpacked/周伟E化11.03'
DESKTOP_YIHUA_DIR = Path('/Users/qianqianawodebaobei/Desktop/数据异化')
CACHE_FILE = PROJECT_DIR / 'material_illustration_cache.json'

# 加载配图缓存
ILLUST_CACHE = {}
if CACHE_FILE.exists():
    try:
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            ILLUST_CACHE = json.load(f)
    except Exception as e:
        print(f"[Warning] 加载材料配图缓存失败: {e}")

# ----------------- 样式定义 -----------------
FONT_FAMILY = "宋体"
NAVY_FILL = PatternFill(start_color='0B192C', end_color='0B192C', fill_type='solid')
CLIENT_HEADER_FILL = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
ZEBRA_FILL = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')
WHITE_FILL = PatternFill(fill_type=None)

THIN_BORDER = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)

HEADER_FONT = Font(name=FONT_FAMILY, size=10, bold=True, color='FFFFFF')
BODY_FONT = Font(name=FONT_FAMILY, size=10)
LINK_FONT = Font(name=FONT_FAMILY, size=10, color='0563C1', underline='single')

CENTER_ALIGN = Alignment(horizontal='center', vertical='center')
RIGHT_ALIGN = Alignment(horizontal='right', vertical='center')
MULTILINE_LEFT_ALIGN = Alignment(horizontal='left', vertical='center', wrap_text=True)

# ----------------- 工具函数 -----------------
def clean_vendor_name(name: str) -> str:
    s = re.sub(r'^\d+[\._、\s]*', '', name)
    s = re.sub(r'_\d{10,}$', '', s)
    s = re.sub(r'\(\d+\)$', '', s)
    s = s.replace('-标准解析', '').replace('.xlsx', '').strip()
    return s

def get_illustration_info(mat_name: str, supplier: str, spec: str = ""):
    """查询材料配图与文件夹名"""
    if not mat_name:
        return "", ""
    clean_k = mat_name.strip()
    if clean_k in ILLUST_CACHE:
        c_item = ILLUST_CACHE[clean_k]
        if isinstance(c_item, dict):
            return c_item.get("pictureUrl", ""), c_item.get("folderName", "")
        elif isinstance(c_item, str):
            return c_item, ""
    try:
        res = default_illustration_service.fetch_illustration_by_name(clean_k, company_name=supplier, specification=spec)
        if res and res.get("pictureUrl"):
            p_url = res["pictureUrl"]
            f_name = res.get("folderName") or ""
            ILLUST_CACHE[clean_k] = {"pictureUrl": p_url, "folderName": f_name}
            return p_url, f_name
    except Exception:
        pass
    return "", ""

# ----------------- 建立全网源文件索引 -----------------
print("[*] 正在扫描全网源文件并构建索引...")
ALL_ZIPS = list(FOUR_AGENTS_DIR.rglob('*.zip')) + list(DESKTOP_YIHUA_DIR.glob('*.zip'))
ZIP_INDEX = {}
for z in ALL_ZIPS:
    k = clean_vendor_name(z.stem)
    if k not in ZIP_INDEX:
        ZIP_INDEX[k] = z

ALL_ONLINE_XLSX = list(FOUR_AGENTS_DIR.rglob('*三方线上解析*/*.xlsx')) + list(DESKTOP_YIHUA_DIR.glob('*解析文件*/*.xlsx'))
ONLINE_XLSX_INDEX = {}
for x in ALL_ONLINE_XLSX:
    k = clean_vendor_name(x.stem)
    if k not in ONLINE_XLSX_INDEX:
        ONLINE_XLSX_INDEX[k] = x

LOCAL_RESULT_INDEX = {}
for r_path in FOUR_AGENTS_DIR.rglob('*-标准解析.xlsx'):
    k = clean_vendor_name(r_path.stem)
    if k not in LOCAL_RESULT_INDEX:
        LOCAL_RESULT_INDEX[k] = r_path

print(f"[*] 索引构建完成: ZIP包 {len(ZIP_INDEX)} 个, 线上XLSX {len(ONLINE_XLSX_INDEX)} 个, 4经办人本地成果 {len(LOCAL_RESULT_INDEX)} 个")

def find_client_raw_wb(vendor_name: str, clean_v: str):
    """获取客户原始解析数据源"""
    # 1. 尝试匹配 ZIP 中的 报价数据.xlsx
    hit_zip = ZIP_INDEX.get(clean_v)
    if not hit_zip:
        for k, z in ZIP_INDEX.items():
            if clean_v in k or (len(clean_v) > 4 and clean_v[:6] in k):
                hit_zip = z
                break
    if hit_zip:
        try:
            with zipfile.ZipFile(hit_zip, 'r') as zf:
                target_fname = None
                for n in zf.namelist():
                    if '报价数据' in n and n.endswith('.xlsx') and not n.startswith('~$'):
                        target_fname = n
                        break
                if not target_fname:
                    for n in zf.namelist():
                        if n.endswith('.xlsx') and not n.startswith('~$'):
                            target_fname = n
                            break
                if target_fname:
                    excel_bytes = zf.read(target_fname)
                    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes), data_only=True)
                    return wb, f"ZIP:{hit_zip.name} -> {target_fname}"
        except Exception:
            pass

    # 2. 尝试匹配 线上解析 XLSX
    hit_ox = ONLINE_XLSX_INDEX.get(clean_v)
    if not hit_ox:
        for k, x in ONLINE_XLSX_INDEX.items():
            if clean_v in k or (len(clean_v) > 4 and clean_v[:6] in k):
                hit_ox = x
                break
    if hit_ox:
        try:
            wb = openpyxl.load_workbook(hit_ox, data_only=True)
            return wb, f"ONLINE_XLSX:{hit_ox.name}"
        except Exception:
            pass

    # 3. 检查周伟E化11.03中是否存在原厂 Excel
    zw_dir = ZHOUWEI_1103_DIR / vendor_name
    if not zw_dir.exists():
        for p in ZHOUWEI_1103_DIR.iterdir():
            if p.is_dir() and (clean_v in p.name or p.name in clean_v):
                zw_dir = p
                break
    if zw_dir.exists():
        excels = [f for f in zw_dir.glob('*.xlsx') if not f.name.startswith('~$')]
        if excels:
            try:
                wb = openpyxl.load_workbook(excels[0], data_only=True)
                return wb, f"ZHOUWEI_EXCEL:{excels[0].name}"
            except Exception:
                pass

    return None, None

def rebuild_single_vendor(delivery_file: Path):
    """针对单个供应商交付件重构双 Sheet 标准文件"""
    vname = delivery_file.stem.replace('-标准解析', '').strip()
    clean_v = clean_vendor_name(vname)
    img_dir = DELIVERIES_DIR / f"{vname}_图片库"

    # 特殊处理：武汉第二电线电缆有限公司已具备高标准规范
    if '武汉第二电线电缆' in vname:
        # 保持现有的纯净版，确保包含双 Sheet
        from rebuild_wuhan_standard import run as run_wuhan
        run_wuhan()
        return True, "武汉第二电缆专项重构完成"

    # 1. 读取现有交付件以提取商务条款与已解析数据
    try:
        wb_existing = openpyxl.load_workbook(delivery_file, data_only=True)
        ws_existing = wb_existing['标准化报价'] if '标准化报价' in wb_existing.sheetnames else wb_existing.active
    except Exception as e:
        print(f"[Error] 读取现有交付件失败 {delivery_file.name}: {e}")
        return False, str(e)

    existing_hdr = [ws_existing.cell(1, c).value for c in range(1, ws_existing.max_column + 1)]

    # 提取商务条款字典
    terms_dict = {
        'quote_desc': '',
        'discount_ratio': 1.0,
        'discount_series': '',
        'is_freight_included': '包含',
        'freight_desc': '',
        'is_tax_included': '含税',
        'tax_rate': 13,
        'tax_desc': '',
        'quote_type': '出厂价'
    }

    def get_first_non_empty(header_name):
        if header_name in existing_hdr:
            c_idx = existing_hdr.index(header_name) + 1
            for r in range(2, min(50, ws_existing.max_row + 1)):
                v = ws_existing.cell(r, c_idx).value
                if v not in (None, '', 'nan', 'None'):
                    return v
        return None

    terms_dict['quote_desc'] = get_first_non_empty('报价说明') or ''
    terms_dict['discount_ratio'] = get_first_non_empty('工程价折扣比例') or 1.0
    terms_dict['discount_series'] = get_first_non_empty('折扣系列') or ''
    terms_dict['is_freight_included'] = get_first_non_empty('是否含运费') or '包含'
    terms_dict['freight_desc'] = get_first_non_empty('运费约定') or ''
    terms_dict['is_tax_included'] = get_first_non_empty('是否含增值税') or '含税'
    terms_dict['tax_rate'] = get_first_non_empty('增值税税率') or 13
    terms_dict['tax_desc'] = get_first_non_empty('增值税说明') or ''
    terms_dict['quote_type'] = get_first_non_empty('价格填报类型') or '出厂价'

    # 2. 提取标准化报价物料数据（去除 31 省份复制爆炸，恢复 1:1 物料行）
    # 优先从 02_四位经办人本地标准成果提取（完全未爆炸的原版真实物料）
    local_source_file = LOCAL_RESULT_INDEX.get(clean_v)
    if not local_source_file:
        for k, p in LOCAL_RESULT_INDEX.items():
            if clean_v in k or (len(clean_v) > 4 and clean_v[:6] in k):
                local_source_file = p
                break

    source_ws = None
    if local_source_file and local_source_file.exists():
        try:
            wb_local = openpyxl.load_workbook(local_source_file, data_only=True)
            source_ws = wb_local['标准化报价'] if '标准化报价' in wb_local.sheetnames else wb_local.active
        except Exception:
            source_ws = None

    raw_items_rows = []
    if source_ws and source_ws.max_row > 1:
        # 从本地未爆炸的原版提取物料
        src_hdr = [source_ws.cell(1, c).value for c in range(1, source_ws.max_column + 1)]
        for r in range(2, source_ws.max_row + 1):
            r_vals = {src_hdr[c-1]: source_ws.cell(r, c).value for c in range(1, len(src_hdr) + 1)}
            raw_items_rows.append(r_vals)
    else:
        # 从现有交付件中收缩（去重全国展开的重复物料）
        prov_col = existing_hdr.index('报价地区省') + 1 if '报价地区省' in existing_hdr else None
        first_prov = None
        if prov_col and ws_existing.max_row > 100:
            first_prov = ws_existing.cell(2, prov_col).value

        seen_keys = set()
        name_c = existing_hdr.index('名称规格') + 1 if '名称规格' in existing_hdr else None
        spec_c = existing_hdr.index('厂家原始规格参数') + 1 if '厂家原始规格参数' in existing_hdr else (existing_hdr.index('规格参数') + 1 if '规格参数' in existing_hdr else None)
        price_c = existing_hdr.index('市场价') + 1 if '市场价' in existing_hdr else None

        for r in range(2, ws_existing.max_row + 1):
            if first_prov and prov_col and ws_existing.cell(r, prov_col).value != first_prov:
                continue
            n_val = ws_existing.cell(r, name_c).value if name_c else ''
            s_val = ws_existing.cell(r, spec_c).value if spec_c else ''
            p_val = ws_existing.cell(r, price_c).value if price_c else ''
            k = (str(n_val).strip(), str(s_val).strip(), str(p_val).strip())
            if k in seen_keys:
                continue
            seen_keys.add(k)
            r_vals = {existing_hdr[c-1]: ws_existing.cell(r, c).value for c in range(1, len(existing_hdr) + 1)}
            raw_items_rows.append(r_vals)

    if not raw_items_rows:
        return False, "未能提取到有效的物料行"

    # 3. 创建全新的双 Sheet 工作簿
    wb_out = openpyxl.Workbook()

    # ----------------------------------------------------
    # Sheet 1: 【客户原始解析】
    # ----------------------------------------------------
    ws_sheet1 = wb_out.active
    ws_sheet1.title = "客户原始解析"

    wb_client, client_src_desc = find_client_raw_wb(vname, clean_v)
    if wb_client:
        ws_client_src = wb_client.active
        # 原封不动复制客户原始解析数据
        for r in range(1, ws_client_src.max_row + 1):
            ws_sheet1.row_dimensions[r].height = ws_client_src.row_dimensions[r].height or 22
            for c in range(1, ws_client_src.max_column + 1):
                cell_val = ws_client_src.cell(r, c).value
                dst_c = ws_sheet1.cell(r, c, cell_val)
                if r == 1:
                    dst_c.font = HEADER_FONT
                    dst_c.fill = CLIENT_HEADER_FILL
                    dst_c.alignment = CENTER_ALIGN
                else:
                    dst_c.font = BODY_FONT
                    dst_c.alignment = MULTILINE_LEFT_ALIGN
                    dst_c.fill = ZEBRA_FILL if (r % 2 == 0) else WHITE_FILL
                dst_c.border = THIN_BORDER
        # 列宽适配
        for col in ws_sheet1.columns:
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            max_len = max(len(str(c.value or '')) for c in col[:15])
            ws_sheet1.column_dimensions[col_letter].width = max(12, min(max_len * 1.5 + 4, 35))
    else:
        # 若无三方线上系统解析表，按客户标准 40 列表头原汁原味构建基础解析表
        std_40_headers = [
            '一级类别', '二级类别', '三级类别', '材料名称', '属性项属性值', '单位', '品牌', '单价',
            '报价说明', '厂家报价文件页码', '厂家报价图片文件名', '顺序号', '市场价', '工程价', '工程价折扣比例',
            '折扣系列', '折扣下限', '折扣上限', '报价地区省', '报价地区市', '报价地区区/县', '是否含运费',
            '运费约定', '是否含增值税', '增值税税率', '增值税说明', '价格填报类型', '厂家名称', '销售区域',
            '有效起始日期', '有效截止日期', '联系人', '联系电话', '备注', '图片', '图片文件夹',
            '厂家原始规格参数', '规格参数', '类别编码', '报价ID'
        ]
        ws_sheet1.row_dimensions[1].height = 30
        for c_idx, h in enumerate(std_40_headers, 1):
            dst_c = ws_sheet1.cell(1, c_idx, h)
            dst_c.font = HEADER_FONT
            dst_c.fill = CLIENT_HEADER_FILL
            dst_c.alignment = CENTER_ALIGN
            dst_c.border = THIN_BORDER

        for r_idx, item_dict in enumerate(raw_items_rows, 2):
            ws_sheet1.row_dimensions[r_idx].height = 22
            row_fill = ZEBRA_FILL if (r_idx % 2 == 0) else WHITE_FILL
            for c_idx, h in enumerate(std_40_headers, 1):
                val = item_dict.get(h, '')
                dst_c = ws_sheet1.cell(r_idx, c_idx, val)
                dst_c.font = BODY_FONT
                dst_c.fill = row_fill
                dst_c.border = THIN_BORDER
                dst_c.alignment = MULTILINE_LEFT_ALIGN

    # ----------------------------------------------------
    # Sheet 2: 【标准化报价】
    # ----------------------------------------------------
    ws_sheet2 = wb_out.create_sheet(title="标准化报价")
    builder = StandardizedWorkbookBuilder(None)
    std_headers = builder.headers

    # 写入表头
    ws_sheet2.row_dimensions[1].height = 36
    for c_idx, h in enumerate(std_headers, 1):
        cell = ws_sheet2.cell(row=1, column=c_idx, value=h)
        cell.font = HEADER_FONT
        cell.fill = NAVY_FILL
        cell.alignment = CENTER_ALIGN
        cell.border = THIN_BORDER

    quote_img_col_idx = std_headers.index('厂家报价图片文件名') + 1
    os.makedirs(img_dir, exist_ok=True)

    # 批量预加载本供应商物料配图 (将单行网络请求优化为单供应商 1 次批量请求)
    unique_names = list({str(it.get('名称规格', '') or '').strip() for it in raw_items_rows if str(it.get('名称规格', '') or '').strip()})
    uncached_for_batch = [n for n in unique_names if n not in ILLUST_CACHE]
    if uncached_for_batch:
        try:
            batch_payload = [{'uid': str(i), 'materialName': n} for i, n in enumerate(uncached_for_batch)]
            res_dict = default_illustration_service.fetch_illustrations_batch(batch_payload, company_name=vname, timeout=5, max_poll_seconds=6)
            for uid, r_info in res_dict.items():
                m_name = uncached_for_batch[int(uid)]
                ILLUST_CACHE[m_name] = {
                    'pictureUrl': r_info.get('pictureUrl', ''),
                    'folderName': r_info.get('folderName', '')
                }
        except Exception:
            pass

    total_embedded = 0
    for r_idx, item_data in enumerate(raw_items_rows, 2):
        row_fill = ZEBRA_FILL if (r_idx % 2 == 0) else WHITE_FILL
        mat_name = str(item_data.get('名称规格', '') or '').strip()
        raw_spec = str(item_data.get('厂家原始规格参数', '') or item_data.get('规格参数', '') or '').strip()

        # 配图查询与直链回填
        pic_url, folder_name = get_illustration_info(mat_name, vname, raw_spec)

        # 补齐商务条款
        for k_term, v_term in terms_dict.items():
            h_map = {
                'quote_desc': '报价说明',
                'discount_ratio': '工程价折扣比例',
                'discount_series': '折扣系列',
                'is_freight_included': '是否含运费',
                'freight_desc': '运费约定',
                'is_tax_included': '是否含增值税',
                'tax_rate': '增值税税率',
                'tax_desc': '增值税说明',
                'quote_type': '价格填报类型'
            }
            target_hdr = h_map.get(k_term)
            if target_hdr and (not item_data.get(target_hdr) or item_data.get(target_hdr) in ('nan', 'None')):
                item_data[target_hdr] = v_term

        # 写入各列
        for c_idx, h in enumerate(std_headers, 1):
            val = item_data.get(h, '')
            if h == '图片' and pic_url:
                val = pic_url
            elif h == '图片文件夹' and folder_name:
                val = folder_name
            elif h == '供应商名称':
                val = vname
            elif h == '顺序号':
                val = (r_idx - 1) * 10
            elif h == '报价ID':
                val = f"REQ{r_idx-1:05d}"
            elif h in ('报价地区省', '报价地区市'):
                val = '全国' if not val else val

            cell = ws_sheet2.cell(row=r_idx, column=c_idx, value=val)
            cell.font = BODY_FONT
            cell.fill = row_fill
            cell.border = THIN_BORDER

            # 样式与对齐
            if h in ['报价ID', '类别编码', '品牌', '厂家报价文件页码', '顺序号', '单位', '工程价折扣比例', '增值税税率', '是否含增值税', '是否含运费', '价格填报类型', '厂家报价图片文件名']:
                cell.alignment = CENTER_ALIGN
            elif h in ['市场价', '工程价', '折扣下限', '折扣上限']:
                cell.alignment = RIGHT_ALIGN
                if isinstance(val, (int, float)):
                    cell.number_format = '0.00'
            else:
                cell.alignment = MULTILINE_LEFT_ALIGN

            # 图片直链超链接
            if h == '图片' and isinstance(val, str) and val.startswith('http'):
                cell.hyperlink = val
                cell.font = LINK_FONT

        # 内嵌缩略图
        img_fn = str(item_data.get('厂家报价图片文件名', '') or '').strip()
        if not img_fn:
            page_val = str(item_data.get('厂家报价文件页码', '') or '').strip()
            m_p = re.search(r'\d+', page_val)
            p_num = m_p.group(0) if m_p else '1'
            img_fn = f"quote_img{p_num}.png"
            ws_sheet2.cell(row=r_idx, column=quote_img_col_idx, value=img_fn)

        target_img_path = img_dir / img_fn
        if target_img_path.exists():
            try:
                thumb_fn = f"thumb_retina_{img_fn}.jpg"
                thumb_path = img_dir / thumb_fn
                if not thumb_path.exists():
                    with PILImage.open(target_img_path) as pim:
                        pim.thumbnail((300, 180))
                        pim.convert("RGB").save(str(thumb_path), "JPEG", quality=85)

                cell_ref = ws_sheet2.cell(row=r_idx, column=quote_img_col_idx).coordinate
                ws_sheet2.row_dimensions[r_idx].height = 42

                xl_img = OpenpyxlImage(str(thumb_path))
                xl_img.width = 75
                xl_img.height = 40
                ws_sheet2.add_image(xl_img, cell_ref)
                total_embedded += 1

                cell_img = ws_sheet2.cell(row=r_idx, column=quote_img_col_idx)
                cell_img.value = img_fn
                cell_img.hyperlink = f"{vname}_图片库/{img_fn}"
                cell_img.font = LINK_FONT
            except Exception:
                pass
        else:
            ws_sheet2.row_dimensions[r_idx].height = 38

    # 列宽适配
    for col in ws_sheet2.columns:
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws_sheet2.column_dimensions[col_letter].width = max(14, min(35, ws_sheet2.column_dimensions[col_letter].width or 16))

    # 安全落盘
    tmp_out = delivery_file.with_suffix('.tmp.xlsx')
    wb_out.save(tmp_out)
    os.replace(tmp_out, delivery_file)

    return True, f"成功 (S1行数={ws_sheet1.max_row}, S2行数={ws_sheet2.max_row}, 嵌图={total_embedded})"

def main():
    print("================================================================================")
    print("全量 161 家供应商双 Sheet 标准化重构主流程启动")
    print("================================================================================")
    deliveries = sorted(list(DELIVERIES_DIR.glob('*-标准解析.xlsx')))
    print(f"[*] 目标交付文件总量: {len(deliveries)}")

    success_count = 0
    fail_count = 0
    start_time = time.time()

    for idx, f in enumerate(deliveries, 1):
        vname = f.stem.replace('-标准解析', '')
        t0 = time.time()
        try:
            ok, msg = rebuild_single_vendor(f)
            elapsed = time.time() - t0
            if ok:
                success_count += 1
                print(f"[{idx:03d}/{len(deliveries):03d}] [✓ PASS] {vname} ({elapsed:.1f}s) -> {msg}")
            else:
                fail_count += 1
                print(f"[{idx:03d}/{len(deliveries):03d}] [✗ FAIL] {vname} ({elapsed:.1f}s) -> {msg}")
        except Exception as e:
            elapsed = time.time() - t0
            fail_count += 1
            print(f"[{idx:03d}/{len(deliveries):03d}] [✗ ERROR] {vname} ({elapsed:.1f}s) -> {e}")

    # 保存最新配图缓存
    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(ILLUST_CACHE, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    total_time = time.time() - start_time
    print("================================================================================")
    print(f"全量重构完成！总用时: {total_time:.1f} 秒")
    print(f"成功: {success_count} 家 | 失败: {fail_count} 家 | 总量: {len(deliveries)} 家")
    print("================================================================================")

if __name__ == '__main__':
    main()
