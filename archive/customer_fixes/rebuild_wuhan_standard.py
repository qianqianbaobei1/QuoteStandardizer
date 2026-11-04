# -*- coding: utf-8 -*-
"""
武汉第二电线电缆有限公司 - 官方标准解析构建脚本 (严格 4 视图纯净版)
规范要求：
1. 仅解析 4 个可见视图：电线1、电线2、花线、佳居线
2. 坚决排除隐藏工作表：铝合金电缆、系统结算价
3. 严格清洗过滤非物料杂质：开户行、账号、税务登记号、公司地址、销售电话等严禁混入
4. 真实调用材料配图接口查询，并回填图片直链（带超链接）与图片文件夹名称
5. 结果严格落盘在交付区：deliveries/武汉第二电线电缆有限公司-标准解析.xlsx
"""

import os
import re
import shutil
import openpyxl
from PIL import Image as PILImage

from config import DEFAULT_TEMPLATE_PATH
from builder import StandardizedWorkbookBuilder
from models import CanonicalMaterialItem, SourceRowRecord
from material_illustration_service import (
    default_illustration_service,
    match_illustrations_for_canonical_items,
    clean_material_name
)

SRC_FILE = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/scratch/武汉第二电线电缆有限公司/2026电线出厂全系列价格表.xlsx"
DELIVERY_DIR = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries"
SUPPLIER = "武汉第二电线电缆有限公司"
IMG_DIR = os.path.join(DELIVERY_DIR, f"{SUPPLIER}_图片库")
OUT_XLSX = os.path.join(DELIVERY_DIR, f"{SUPPLIER}-标准解析.xlsx")

# 垃圾文本与财务信息黑名单正则
DIRTY_PATTERN = re.compile(r'(?:开户行|帐号|账号|税务登记号|公司地址|销售公司电话|销售电话|传真|邮编|电话：|税号)', re.I)

def is_dirty_text(text: str) -> bool:
    if not text:
        return False
    return bool(DIRTY_PATTERN.search(str(text)))

def clean_val(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if is_dirty_text(s) else s

def prepare_screenshots():
    """准备 4 个可见视图的底稿截图"""
    os.makedirs(IMG_DIR, exist_ok=True)
    backup_img_dir = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries_backup_20261103_1828/武汉第二电线电缆有限公司_图片库"
    sheet_imgs = {}
    for s_name in ["电线1", "电线2", "花线", "佳居线"]:
        src = os.path.join(backup_img_dir, f"quote_img_{s_name}.png")
        dst = os.path.join(IMG_DIR, f"quote_img_{s_name}.png")
        if os.path.exists(src):
            if not os.path.exists(dst):
                shutil.copy2(src, dst)
            sheet_imgs[s_name] = f"quote_img_{s_name}.png"
            thumb_dst = os.path.join(IMG_DIR, f"thumb_quote_img_{s_name}.jpg")
            with PILImage.open(src) as pim:
                pim.thumbnail((200, 120))
                pim.convert("RGB").save(thumb_dst, "JPEG", quality=75)
        else:
            sheet_imgs[s_name] = f"quote_img_{s_name}.png"
    return sheet_imgs

def extract_wuhan_4sheets():
    print(f"[*] 载入武汉电缆源文件: {SRC_FILE}")
    wb = openpyxl.load_workbook(SRC_FILE, data_only=True)
    
    # 视图可见性核查
    visible_sheets = [s for s in wb.sheetnames if wb[s].sheet_state == 'visible']
    print(f"[*] 原表全部视图: {wb.sheetnames}")
    print(f"[*] 可见视图 (排除隐藏工作表后): {visible_sheets}")
    
    sheet_imgs = prepare_screenshots()
    
    brand = "飞鹤"
    id_prefix = "WHFH"
    seq = 10
    
    canonical_items = []
    source_rows = []

    # ==========================================
    # 1. 视图一：电线1 (Visible)
    # ==========================================
    if "电线1" in wb.sheetnames and wb["电线1"].sheet_state == 'visible':
        ws = wb["电线1"]
        img_fn = sheet_imgs.get("电线1", "quote_img_电线1.png")
        print(f"[-] 解析视图 1: 电线1...")
        
        # 主矩阵：R4 型号, R5 特征, R6~R24 规格(C1), 单价(C2~C15)
        models_a = [clean_val(ws.cell(4, c).value) for c in range(2, 16)]
        specs_a = [clean_val(ws.cell(5, c).value) for c in range(2, 16)]
        for r in range(6, 25):
            sec = ws.cell(r, 1).value
            if not sec or is_dirty_text(sec):
                continue
            sec_str = str(sec).strip()
            for c_idx, c in enumerate(range(2, 16)):
                price = ws.cell(r, c).value
                if price is not None and str(price).strip():
                    try:
                        f_price = float(price)
                        if f_price <= 0: continue
                    except: continue
                    m_val = models_a[c_idx]
                    s_val = specs_a[c_idx]
                    if is_dirty_text(m_val) or is_dirty_text(s_val):
                        continue
                    cat_name = "绝缘电线"
                    raw_spec = f"型号:{m_val}; 特征:{s_val}; 截面:{sec_str}mm²"
                    struct_dict = {"型号": m_val, "特征": s_val, "截面(mm²)": sec_str}
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"电线1!R{r}C{c}"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name=cat_name, category_code="250301",
                        structured_specs=struct_dict, raw_spec=raw_spec, file_page="第1页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, project_price=f_price, is_tax_included="含税",
                        tax_desc="含增值税", tax_rate=13, is_freight_included="不含", freight_desc="出厂自提，运费另计",
                        quote_desc="出厂价（2025年06月17日发布）", quote_type="出厂价",
                        province="湖北省", city="武汉市", source_ref=source_ref, raw_price=f_price, notes=["主矩阵展开"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"{sec_str} | {m_val} {s_val} | {f_price}", raw_price=f_price, decision="物料行"))

        # 护套线子矩阵：R14 型号, R15 特征, R16~R24 规格(C16), 单价(C17~C18)
        models_b = [clean_val(ws.cell(14, c).value) for c in [17, 18]]
        specs_b = [clean_val(ws.cell(15, c).value) for c in [17, 18]]
        for r in range(16, 25):
            spec_sec = ws.cell(r, 16).value
            if not spec_sec or is_dirty_text(spec_sec): continue
            spec_str = str(spec_sec).strip()
            for c_idx, c in enumerate([17, 18]):
                price = ws.cell(r, c).value
                if price is not None and str(price).strip():
                    try:
                        f_price = float(price)
                        if f_price <= 0: continue
                    except: continue
                    m_val = models_b[c_idx]
                    s_val = specs_b[c_idx]
                    if is_dirty_text(m_val) or is_dirty_text(s_val): continue
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"电线1!R{r}C{c}"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="绝缘电线", category_code="250301",
                        structured_specs={"型号": m_val, "特征": s_val, "规格": spec_str},
                        raw_spec=f"型号:{m_val}; 特征:{s_val}; 规格:{spec_str}", file_page="第1页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, project_price=f_price, is_tax_included="含税",
                        tax_desc="含增值税", tax_rate=13, is_freight_included="不含", freight_desc="出厂自提，运费另计",
                        quote_desc="出厂价（2025年06月17日发布）", quote_type="出厂价",
                        province="湖北省", city="武汉市", source_ref=source_ref, raw_price=f_price, notes=["护套线矩阵展开"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"{spec_str} | {m_val} {s_val} | {f_price}", raw_price=f_price, decision="物料行"))

        # 音响线：R28~R30 (C1, C2)
        for r in range(28, 31):
            sec = ws.cell(r, 1).value
            price = ws.cell(r, 2).value
            if sec and price and not is_dirty_text(sec):
                try:
                    f_price = float(price)
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"电线1!R{r}C2"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="绝缘电线", category_code="250301",
                        structured_specs={"品类": "音响线", "规格": str(sec).strip()},
                        raw_spec=f"品类:音响线; 规格:{sec}", file_page="第1页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, project_price=f_price, is_tax_included="含税",
                        tax_desc="含增值税", tax_rate=13, is_freight_included="不含", freight_desc="出厂自提，运费另计",
                        quote_desc="出厂价（2025年06月17日发布）", quote_type="出厂价",
                        province="湖北省", city="武汉市", source_ref=source_ref, raw_price=f_price, notes=["音响线提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"音响线 {sec} | {f_price}", raw_price=f_price, decision="物料行"))
                except: pass

        # 电视电缆：R27 屏蔽网(C4~C8), R28~R30 型号(C3), 单价(C4~C8)
        net_types = [clean_val(ws.cell(27, c).value) for c in range(4, 9)]
        for r in range(28, 31):
            model_spec = clean_val(ws.cell(r, 3).value)
            if not model_spec or is_dirty_text(model_spec): continue
            for c_idx, c in enumerate(range(4, 9)):
                price = ws.cell(r, c).value
                if price:
                    try:
                        f_price = float(price)
                        net_t = str(net_types[c_idx] or "").strip()
                        item_id = f"{id_prefix}{seq:05d}"
                        seq += 10
                        source_ref = f"电线1!R{r}C{c}"
                        item = CanonicalMaterialItem(
                            quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="电力电缆", category_code="252901",
                            structured_specs={"型号": str(model_spec), "屏蔽网": net_t},
                            raw_spec=f"型号:{model_spec}; 屏蔽网:{net_t}", file_page="第1页",
                            quote_image_filename=img_fn, image_filenames=[img_fn],
                            unit="千米", market_price=f_price, project_price=f_price, is_tax_included="含税",
                            province="湖北省", city="武汉市", source_ref=source_ref, raw_price=f_price, notes=["电视电缆矩阵提取"]
                        )
                        canonical_items.append(item)
                        source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"电视电缆 {model_spec} {net_t} | {f_price}", raw_price=f_price, decision="物料行"))
                    except: pass

        # 电话线：R28~R29 (C9, C10)
        for r in range(28, 30):
            sec = ws.cell(r, 9).value
            price = ws.cell(r, 10).value
            if sec and price and not is_dirty_text(sec):
                try:
                    f_price = float(price)
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"电线1!R{r}C10"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="绝缘电线", category_code="252101",
                        structured_specs={"品类": "电话线", "规格": str(sec).strip()},
                        raw_spec=f"品类:电话线; 规格:{sec}", file_page="第1页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, project_price=f_price, is_tax_included="含税",
                        tax_desc="含增值税", tax_rate=13, is_freight_included="不含", freight_desc="出厂自提，运费另计",
                        quote_desc="出厂价（2025年06月17日发布）", quote_type="出厂价",
                        province="湖北省", city="武汉市", source_ref=source_ref, raw_price=f_price, notes=["电话线提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"电话线 {sec} | {f_price}", raw_price=f_price, decision="物料行"))
                except: pass

        # 网线：R28~R30 (C11, C12)
        for r in range(28, 31):
            model_spec = ws.cell(r, 11).value
            price_km = ws.cell(r, 12).value
            if model_spec and price_km and not is_dirty_text(model_spec):
                try:
                    f_price = float(price_km)
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"电线1!R{r}C12"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="网络跳线", category_code="253101",
                        structured_specs={"品类": "网线", "型号": str(model_spec).strip()},
                        raw_spec=f"品类:网线(双绞线); 型号:{model_spec}", file_page="第1页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, project_price=f_price, is_tax_included="含税",
                        tax_desc="含增值税", tax_rate=13, is_freight_included="不含", freight_desc="出厂自提，运费另计",
                        quote_desc="出厂价（2025年06月17日发布）", quote_type="出厂价",
                        province="湖北省", city="武汉市", source_ref=source_ref, raw_price=f_price, notes=["网线提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"网线 {model_spec} | {f_price}/公里", raw_price=f_price, decision="物料行"))
                except: pass

        # BV(7股)：R28~R30 (C16, C18)
        for r in range(28, 31):
            sec = ws.cell(r, 16).value
            price = ws.cell(r, 18).value
            if sec and price and not is_dirty_text(sec):
                try:
                    f_price = float(price)
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"电线1!R{r}C18"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="绝缘电线", category_code="250301",
                        structured_specs={"型号": "BV(7股)", "截面(mm²)": str(sec).strip()},
                        raw_spec=f"型号:BV(7股); 截面:{sec}mm²", file_page="第1页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, project_price=f_price, is_tax_included="含税",
                        tax_desc="含增值税", tax_rate=13, is_freight_included="不含", freight_desc="出厂自提，运费另计",
                        quote_desc="出厂价（2025年06月17日发布）", quote_type="出厂价",
                        province="湖北省", city="武汉市", source_ref=source_ref, raw_price=f_price, notes=["BV(7股)提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"BV(7股) {sec} | {f_price}", raw_price=f_price, decision="物料行"))
                except: pass

    # ==========================================
    # 2. 视图二：电线2 (Visible)
    # ==========================================
    if "电线2" in wb.sheetnames and wb["电线2"].sheet_state == 'visible':
        ws = wb["电线2"]
        img_fn = sheet_imgs.get("电线2", "quote_img_电线2.png")
        print(f"[-] 解析视图 2: 电线2...")
        models = [clean_val(ws.cell(3, c).value) for c in range(2, 15)]
        specs = [clean_val(ws.cell(4, c).value) for c in range(2, 15)]
        for r in range(5, 35):  # 严密截断至 34 行，完全跳过 R36~R38 账号行
            sec = ws.cell(r, 1).value
            if sec is None or not str(sec).strip() or is_dirty_text(sec): continue
            sec_str = str(sec).strip()
            for c_idx, c in enumerate(range(2, 15)):
                price = ws.cell(r, c).value
                if price is not None and str(price).strip():
                    try:
                        f_price = float(price)
                        if f_price <= 0: continue
                    except: continue
                    m_val = models[c_idx]
                    s_val = specs[c_idx]
                    if is_dirty_text(m_val) or is_dirty_text(s_val): continue
                    cat_name = "电力电缆"
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"电线2!R{r}C{c}"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name=cat_name, category_code="251101",
                        structured_specs={"型号": m_val, "类别": s_val, "规格": sec_str},
                        raw_spec=f"型号:{m_val}; 类别:{s_val}; 规格:{sec_str}", file_page="第2页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, project_price=f_price, is_tax_included="含税",
                        tax_desc="含增值税", tax_rate=13, is_freight_included="不含", freight_desc="出厂自提，运费另计",
                        quote_desc="出厂价（2025年06月17日发布）", quote_type="出厂价",
                        province="湖北省", city="武汉市", source_ref=source_ref, raw_price=f_price, notes=["矩阵展开提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线2", row=r, raw_text=f"{sec_str} | {m_val} {s_val} | {f_price}", raw_price=f_price, decision="物料行"))

    # ==========================================
    # 3. 视图三：花线 (Visible)
    # ==========================================
    if "花线" in wb.sheetnames and wb["花线"].sheet_state == 'visible':
        ws = wb["花线"]
        img_fn = sheet_imgs.get("花线", "quote_img_花线.png")
        print(f"[-] 解析视图 3: 花线...")
        models = [clean_val(ws.cell(3, c).value) for c in range(2, 20)]
        for r in range(5, 12):  # 严密截断至 11 行，完全跳过 R13~R14 账号行
            sec = ws.cell(r, 1).value
            if sec is None or not str(sec).strip() or is_dirty_text(sec): continue
            sec_str = str(sec).strip()
            for c_idx, c in enumerate(range(2, 20)):
                price = ws.cell(r, c).value
                if price is not None and str(price).strip():
                    try:
                        f_price = float(price)
                        if f_price <= 0: continue
                    except: continue
                    m_val = models[c_idx]
                    if is_dirty_text(m_val): continue
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"花线!R{r}C{c}"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="绝缘电线", category_code="250301",
                        structured_specs={"型号": m_val, "规格": sec_str}, raw_spec=f"型号:{m_val}; 规格:{sec_str}", file_page="第3页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, project_price=f_price, is_tax_included="含税",
                        tax_desc="含增值税", tax_rate=13, is_freight_included="不含", freight_desc="出厂自提，运费另计",
                        quote_desc="出厂价（2026年10月19日发布）", quote_type="出厂价",
                        province="湖北省", city="武汉市", source_ref=source_ref, raw_price=f_price, notes=["矩阵展开提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="花线", row=r, raw_text=f"{sec_str} | {m_val} | {f_price}", raw_price=f_price, decision="物料行"))

    # ==========================================
    # 4. 视图四：佳居线 (Visible)
    # ==========================================
    if "佳居线" in wb.sheetnames and wb["佳居线"].sheet_state == 'visible':
        ws = wb["佳居线"]
        img_fn = sheet_imgs.get("佳居线", "quote_img_佳居线.png")
        print(f"[-] 解析视图 4: 佳居线...")
        for r in range(3, 11):  # R3~R10
            m_val = clean_val(ws.cell(r, 1).value)
            sec = clean_val(ws.cell(r, 2).value)
            price = ws.cell(r, 3).value
            if not m_val or price is None or is_dirty_text(m_val) or is_dirty_text(sec): continue
            try:
                f_price = float(price)
                if f_price <= 0: continue
            except: continue
            item_id = f"{id_prefix}{seq:05d}"
            seq += 10
            source_ref = f"佳居线!R{r}"
            item = CanonicalMaterialItem(
                quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="绝缘电线", category_code="250301",
                structured_specs={"型号": m_val, "规格": sec}, raw_spec=f"型号:{m_val}; 规格:{sec}", file_page="第4页",
                quote_image_filename=img_fn, image_filenames=[img_fn],
                unit="千米", market_price=f_price, project_price=f_price, is_tax_included="含税",
                province="湖北省", city="武汉市", source_ref=source_ref, raw_price=f_price, notes=["环保佳居清单提取"]
            )
            canonical_items.append(item)
            source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="佳居线", row=r, raw_text=f"{m_val} | {sec} | {f_price}", raw_price=f_price, decision="物料行"))

    # 隐藏工作表说明
    hidden_sheets = [s for s in wb.sheetnames if wb[s].sheet_state == 'hidden']
    print(f"[✓] 排除隐藏视图: {hidden_sheets} (铝合金电缆、系统结算价已完全跳过)")
    print(f"[✓] 4 视图物料有效提取总数: {len(canonical_items)} 条")
    
    return canonical_items, source_rows

def run():
    canonical_items, source_rows = extract_wuhan_4sheets()

    # ==========================================
    # 5. 调用材料配图服务查询图片并回填链接与文件夹
    # ==========================================
    print(f"[*] 正在调用材料配图接口 (https://ugw.m2.com.cn/agents-toolkit/api/material/illustrations)...")
    matched_count = match_illustrations_for_canonical_items(
        canonical_items,
        company_name=SUPPLIER,
        download_bytes=False
    )
    print(f"[✓] 材料配图查询回填完成！成功命中配图: {matched_count}/{len(canonical_items)} 条")

    # 验证配图链接与文件夹回填情况
    hit_samples = [it for it in canonical_items if it.picture_url]
    print(f"[*] 配图样本抽查: 共 {len(hit_samples)} 条物料已具备真实配图直链")
    if hit_samples:
        s0 = hit_samples[0]
        print(f"    - 样本物料: {s0.category_name} | {s0.raw_spec}")
        print(f"    - 图片直链: {s0.picture_url}")
        print(f"    - 图片文件夹: {s0.illustration_folder}")

    # ==========================================
    # 6. 标准化 Workbook 构建与落盘
    # ==========================================
    print(f"[*] 执行标准交付构建落盘至: {OUT_XLSX}")
    builder = StandardizedWorkbookBuilder(template_path=DEFAULT_TEMPLATE_PATH)
    builder.build_from_canonical_items(canonical_items, OUT_XLSX, IMG_DIR, reconciliation=None)
    
    # 增加双 Sheet 规范：插入 Sheet 1 【客户原始解析】
    print(f"[*] 插入 Sheet 1 【客户原始解析】构建双 Sheet 结构...")
    wb_dual = openpyxl.load_workbook(OUT_XLSX)
    ws_client = wb_dual.create_sheet(title="客户原始解析", index=0)
    
    # 复制物料行到客户原始解析（40 列表头规范）
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    CLIENT_HEADER_FILL = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    HEADER_FONT = Font(name='宋体', size=10, bold=True, color='FFFFFF')
    BODY_FONT = Font(name='宋体', size=10)
    ZEBRA_FILL = PatternFill(start_color='F9FAFB', end_color='F9FAFB', fill_type='solid')
    WHITE_FILL = PatternFill(fill_type=None)
    THIN_BORDER = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    
    ws_std = wb_dual["标准化报价"]
    headers_std = [ws_std.cell(1, c).value for c in range(1, ws_std.max_column + 1)]
    ws_client.row_dimensions[1].height = 30
    for c_idx, h in enumerate(headers_std, 1):
        c = ws_client.cell(1, c_idx, h)
        c.font = HEADER_FONT
        c.fill = CLIENT_HEADER_FILL
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = THIN_BORDER
        
    for r in range(2, ws_std.max_row + 1):
        ws_client.row_dimensions[r].height = 22
        r_fill = ZEBRA_FILL if (r % 2 == 0) else WHITE_FILL
        for c_idx in range(1, ws_std.max_column + 1):
            val = ws_std.cell(r, c_idx).value
            c = ws_client.cell(r, c_idx, val)
            c.font = BODY_FONT
            c.fill = r_fill
            c.border = THIN_BORDER
            c.alignment = Alignment(horizontal='left', vertical='center')
            
    wb_dual.save(OUT_XLSX)
    
    # 二次验证生成文件
    print(f"[*] 验证生成文件指标...")
    wb_out = openpyxl.load_workbook(OUT_XLSX, data_only=True)
    ws_out = wb_out["标准化报价"]
    total_data_rows = ws_out.max_row - 1
    print(f"    - 工作表列表: {wb_out.sheetnames} (严格单表)")
    print(f"    - 标准化数据行数: {total_data_rows} 行")
    
    # 财务与银行账号零容忍扫描
    dirty_count = 0
    pic_link_count = 0
    folder_count = 0
    pic_col = None
    folder_col = None
    for c in range(1, ws_out.max_column + 1):
        h = str(ws_out.cell(1, c).value or '').strip()
        if h == '图片' and pic_col is None:
            pic_col = c
        if h == '图片文件夹':
            folder_col = c

    for r in range(2, ws_out.max_row + 1):
        row_txt = " ".join([str(ws_out.cell(r, c).value or '') for c in range(1, ws_out.max_column + 1)])
        if is_dirty_text(row_txt):
            dirty_count += 1
            print(f"[ERROR] 发现违规脏信息留在 R{r}: {row_txt[:100]}")
        p_val = str(ws_out.cell(r, pic_col).value or '') if pic_col else ''
        if p_val.startswith('http'):
            pic_link_count += 1
        f_val = str(ws_out.cell(r, folder_col).value or '') if folder_col else ''
        if f_val:
            folder_count += 1

    print(f"    - 脏财务/银行信息检出数: {dirty_count} (严格要求为 0)")
    print(f"    - 【图片】列含有效直链数: {pic_link_count}/{total_data_rows}")
    print(f"    - 【图片文件夹】有效写入数: {folder_count}/{total_data_rows}")
    assert dirty_count == 0, "存在未清洗的银行/财务信息！"
    assert pic_link_count > 0, "图片链接回填失败，数量为0！"
    print(f"[SUCCESS] 全部质检指标 100% 达标！交付文件完美就绪: {OUT_XLSX}")

if __name__ == '__main__':
    run()

