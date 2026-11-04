from _adhoc_paths import DELIVERIES_DIR, wuhan_source_xlsx
import os
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from PIL import Image as PILImage
from config import DEFAULT_TEMPLATE_PATH
from builder import StandardizedWorkbookBuilder
from models import CanonicalMaterialItem, SourceRowRecord

INPUT_FILE = wuhan_source_xlsx()
DELIVERY_DIR = DELIVERIES_DIR
SUPPLIER = "武汉第二电线电缆有限公司"
IMG_DIR = os.path.join(DELIVERY_DIR, f"{SUPPLIER}_图片库")

def clean_val(v):
    if v is None:
        return ""
    return str(v).strip()

def run():
    print(f"[*] 生成轻量高性能标准工作簿（矩阵页全量内嵌底稿图，5.8万行清单保留页码与配图）...")
    wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)
    brand = "飞鹤"
    id_prefix = "WHFH"
    seq = 10
    
    canonical_items = []
    source_rows = []

    # Sheet: 电线1
    if "电线1" in wb.sheetnames:
        ws = wb["电线1"]
        img_fn = "quote_img_电线1.png"
        models_a = [clean_val(ws.cell(4, c).value) for c in range(2, 16)]
        specs_a = [clean_val(ws.cell(5, c).value) for c in range(2, 16)]
        for r in range(6, 25):
            sec = ws.cell(r, 1).value
            if not sec or not str(sec).strip(): continue
            for c_idx, c in enumerate(range(2, 16)):
                price = ws.cell(r, c).value
                if price is not None and str(price).strip():
                    try:
                        f_price = float(price)
                        if f_price <= 0: continue
                    except: continue
                    m_val = models_a[c_idx]
                    s_val = specs_a[c_idx]
                    cat_name = "电线" if "线" in s_val else "绝缘电线"
                    raw_spec = f"型号:{m_val}; 特征:{s_val}; 截面:{sec}mm²"
                    struct_dict = {"型号": m_val, "特征": s_val, "截面(mm²)": str(sec)}
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"电线1!R{r}C{c}"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name=cat_name, category_code="170101",
                        structured_specs=struct_dict, raw_spec=raw_spec, file_page="第1页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                        source_ref=source_ref, raw_price=f_price, notes=["主矩阵提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"{sec} | {m_val} {s_val} | {f_price}", raw_price=f_price, decision="物料行"))

        # 护套线
        models_b = [clean_val(ws.cell(14, c).value) for c in [17, 18]]
        specs_b = [clean_val(ws.cell(15, c).value) for c in [17, 18]]
        for r in range(16, 25):
            spec_sec = ws.cell(r, 16).value
            if not spec_sec: continue
            for c_idx, c in enumerate([17, 18]):
                price = ws.cell(r, c).value
                if price is not None and str(price).strip():
                    try: f_price = float(price)
                    except: continue
                    m_val = models_b[c_idx]
                    s_val = specs_b[c_idx]
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"电线1!R{r}C{c}"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="护套绝缘电线", category_code="170101",
                        structured_specs={"型号": m_val, "特征": s_val, "规格": str(spec_sec)},
                        raw_spec=f"型号:{m_val}; 特征:{s_val}; 规格:{spec_sec}", file_page="第1页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                        source_ref=source_ref, raw_price=f_price, notes=["护套线矩阵提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"{spec_sec} | {m_val} {s_val} | {f_price}", raw_price=f_price, decision="物料行"))

        # 音响线
        for r in range(28, 31):
            sec = ws.cell(r, 1).value
            price = ws.cell(r, 2).value
            if sec and price:
                f_price = float(price)
                item_id = f"{id_prefix}{seq:05d}"
                seq += 10
                source_ref = f"电线1!R{r}C2"
                item = CanonicalMaterialItem(
                    quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="音响线", category_code="170101",
                    structured_specs={"品类": "音响线", "规格": str(sec)}, raw_spec=f"品类:音响线; 规格:{sec}", file_page="第1页",
                    quote_image_filename=img_fn, image_filenames=[img_fn],
                    unit="千米", market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                    source_ref=source_ref, raw_price=f_price, notes=["音响线提取"]
                )
                canonical_items.append(item)
                source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"音响线 {sec} | {f_price}", raw_price=f_price, decision="物料行"))

        # 电视电缆
        net_types = [ws.cell(27, c).value for c in range(4, 9)]
        for r in range(28, 31):
            model_spec = ws.cell(r, 3).value
            if not model_spec: continue
            for c_idx, c in enumerate(range(4, 9)):
                price = ws.cell(r, c).value
                if price:
                    f_price = float(price)
                    net_t = str(net_types[c_idx] or "").strip()
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"电线1!R{r}C{c}"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="电视电缆(同轴电缆)", category_code="170101",
                        structured_specs={"型号": str(model_spec), "屏蔽网": net_t}, raw_spec=f"型号:{model_spec}; 屏蔽网:{net_t}", file_page="第1页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                        source_ref=source_ref, raw_price=f_price, notes=["电视电缆矩阵提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"电视电缆 {model_spec} {net_t} | {f_price}", raw_price=f_price, decision="物料行"))

        # 电话线
        for r in range(28, 30):
            sec = ws.cell(r, 9).value
            price = ws.cell(r, 10).value
            if sec and price:
                f_price = float(price)
                item_id = f"{id_prefix}{seq:05d}"
                seq += 10
                source_ref = f"电线1!R{r}C10"
                item = CanonicalMaterialItem(
                    quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="电话线", category_code="170101",
                    structured_specs={"品类": "电话线", "规格": str(sec)}, raw_spec=f"品类:电话线; 规格:{sec}", file_page="第1页",
                    quote_image_filename=img_fn, image_filenames=[img_fn],
                    unit="千米", market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                    source_ref=source_ref, raw_price=f_price, notes=["电话线提取"]
                )
                canonical_items.append(item)
                source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"电话线 {sec} | {f_price}", raw_price=f_price, decision="物料行"))

        # 网线
        for r in range(28, 31):
            model_spec = ws.cell(r, 11).value
            price_km = ws.cell(r, 12).value
            if model_spec and price_km:
                f_price = float(price_km)
                item_id = f"{id_prefix}{seq:05d}"
                seq += 10
                source_ref = f"电线1!R{r}C12"
                item = CanonicalMaterialItem(
                    quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="网络电缆(双绞线)", category_code="170101",
                    structured_specs={"品类": "网线", "型号": str(model_spec)}, raw_spec=f"品类:网线(双绞线); 型号:{model_spec}", file_page="第1页",
                    quote_image_filename=img_fn, image_filenames=[img_fn],
                    unit="千米", market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                    source_ref=source_ref, raw_price=f_price, notes=["网线提取"]
                )
                canonical_items.append(item)
                source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"网线 {model_spec} | {f_price}/公里", raw_price=f_price, decision="物料行"))

        # BV(7股)
        for r in range(28, 31):
            sec = ws.cell(r, 16).value
            price = ws.cell(r, 18).value
            if sec and price:
                f_price = float(price)
                item_id = f"{id_prefix}{seq:05d}"
                seq += 10
                source_ref = f"电线1!R{r}C18"
                item = CanonicalMaterialItem(
                    quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="绝缘电线", category_code="170101",
                    structured_specs={"型号": "BV(7股)", "截面(mm²)": str(sec)}, raw_spec=f"型号:BV(7股); 截面:{sec}mm²", file_page="第1页",
                    quote_image_filename=img_fn, image_filenames=[img_fn],
                    unit="千米", market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                    source_ref=source_ref, raw_price=f_price, notes=["BV(7股)提取"]
                )
                canonical_items.append(item)
                source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线1", row=r, raw_text=f"BV(7股) {sec} | {f_price}", raw_price=f_price, decision="物料行"))

    # Sheet: 电线2
    if "电线2" in wb.sheetnames:
        ws = wb["电线2"]
        img_fn = "quote_img_电线2.png"
        models = [clean_val(ws.cell(3, c).value) for c in range(2, 15)]
        specs = [clean_val(ws.cell(4, c).value) for c in range(2, 15)]
        for r in range(5, 35):
            sec = ws.cell(r, 1).value
            if sec is None or not str(sec).strip(): continue
            for c_idx, c in enumerate(range(2, 15)):
                price = ws.cell(r, c).value
                if price is not None and str(price).strip():
                    try: f_price = float(price)
                    except: continue
                    m_val = models[c_idx]
                    s_val = specs[c_idx]
                    cat_name = "控制电缆" if "控制" in s_val else ("电力电缆" if "电缆" in s_val else "绝缘软线")
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"电线2!R{r}C{c}"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name=cat_name, category_code="170101",
                        structured_specs={"型号": m_val, "类别": s_val, "规格": str(sec)}, raw_spec=f"型号:{m_val}; 类别:{s_val}; 规格:{sec}", file_page="第2页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                        source_ref=source_ref, raw_price=f_price, notes=["矩阵展开提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="电线2", row=r, raw_text=f"{sec} | {m_val} {s_val} | {f_price}", raw_price=f_price, decision="物料行"))

    # Sheet: 花线
    if "花线" in wb.sheetnames:
        ws = wb["花线"]
        img_fn = "quote_img_花线.png"
        models = [clean_val(ws.cell(3, c).value) for c in range(2, 20)]
        for r in range(5, 12):
            sec = ws.cell(r, 1).value
            if sec is None or not str(sec).strip(): continue
            for c_idx, c in enumerate(range(2, 20)):
                price = ws.cell(r, c).value
                if price is not None and str(price).strip():
                    try: f_price = float(price)
                    except: continue
                    m_val = models[c_idx]
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"花线!R{r}C{c}"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="连接用软线(花线)", category_code="170101",
                        structured_specs={"型号": m_val, "规格": str(sec)}, raw_spec=f"型号:{m_val}; 规格:{sec}", file_page="第3页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                        source_ref=source_ref, raw_price=f_price, notes=["矩阵展开提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="花线", row=r, raw_text=f"{sec} | {m_val} | {f_price}", raw_price=f_price, decision="物料行"))

    # Sheet: 佳居线
    if "佳居线" in wb.sheetnames:
        ws = wb["佳居线"]
        img_fn = "quote_img_佳居线.png"
        for r in range(3, ws.max_row + 1):
            m_val = clean_val(ws.cell(r, 1).value)
            sec = clean_val(ws.cell(r, 2).value)
            price = ws.cell(r, 3).value
            if not m_val or price is None: continue
            try: f_price = float(price)
            except: continue
            item_id = f"{id_prefix}{seq:05d}"
            seq += 10
            source_ref = f"佳居线!R{r}"
            item = CanonicalMaterialItem(
                quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="环保佳居电线", category_code="170101",
                structured_specs={"型号": m_val, "规格": str(sec)}, raw_spec=f"型号:{m_val}; 规格:{sec}", file_page="第4页",
                quote_image_filename=img_fn, image_filenames=[img_fn],
                unit="千米", market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                source_ref=source_ref, raw_price=f_price, notes=["清单解析"]
            )
            canonical_items.append(item)
            source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="佳居线", row=r, raw_text=f"{m_val} | {sec} | {f_price}", raw_price=f_price, decision="物料行"))

    # Sheet: 铝合金电缆
    if "铝合金电缆" in wb.sheetnames:
        ws = wb["铝合金电缆"]
        img_fn = "quote_img_铝合金电缆.png"
        cores = [ws.cell(4, c).value for c in range(2, 18)]
        cur_core = ""
        for idx in range(len(cores)):
            if cores[idx] is not None and str(cores[idx]).strip():
                cur_core = str(cores[idx]).strip()
            cores[idx] = cur_core
        models = [clean_val(ws.cell(6, c).value) for c in range(2, 18)]
        for r in range(8, 25):
            sec = ws.cell(r, 1).value
            if sec is None or not str(sec).strip(): continue
            for c_idx, c in enumerate(range(2, 18)):
                price = ws.cell(r, c).value
                if price is not None and str(price).strip():
                    try: f_price = float(price)
                    except: continue
                    c_val = cores[c_idx]
                    m_val = models[c_idx]
                    item_id = f"{id_prefix}{seq:05d}"
                    seq += 10
                    source_ref = f"铝合金电缆!R{r}C{c}"
                    item = CanonicalMaterialItem(
                        quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name="铝合金电力电缆", category_code="170101",
                        structured_specs={"芯数": c_val, "型号": m_val, "规格(mm²)": str(sec)}, raw_spec=f"芯数:{c_val}; 型号:{m_val}; 规格:{sec}mm²", file_page="第5页",
                        quote_image_filename=img_fn, image_filenames=[img_fn],
                        unit="千米", market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                        source_ref=source_ref, raw_price=f_price, notes=["矩阵展开提取"]
                    )
                    canonical_items.append(item)
                    source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="铝合金电缆", row=r, raw_text=f"{sec} | {c_val} {m_val} | {f_price}", raw_price=f_price, decision="物料行"))

    # 标记已完成内嵌的前 5 个核心矩阵物料数量
    core_matrix_count = len(canonical_items)
    print(f"[*] 前 5 张核心图纸物料共 {core_matrix_count} 条，将逐行 100% 嵌入底稿截图！")

    # Sheet: 系统结算价（5.8万行流式导入）
    if "系统结算价" in wb.sheetnames:
        ws = wb["系统结算价"]
        print(f"[-] 高速流式导入 Sheet: 系统结算价...")
        for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not row or not any(row): continue
            code, name, unit, spec, model, color, price = row[0], row[1], row[2], row[3], row[4], row[5], row[6]
            if not name or price is None: continue
            try: f_price = float(price)
            except: continue
            unit_str = str(unit or "千米").strip()
            name_str = str(name).strip()
            spec_str = str(spec or "").strip()
            model_str = str(model or "").strip()
            color_str = str(color or "").strip()
            raw_parts = []
            struct_dict = {}
            if model_str:
                raw_parts.append(f"型号:{model_str}")
                struct_dict["型号"] = model_str
            if spec_str:
                raw_parts.append(f"规格:{spec_str}")
                struct_dict["规格"] = spec_str
            if color_str:
                raw_parts.append(f"颜色:{color_str}")
                struct_dict["颜色"] = color_str
            raw_spec = "; ".join(raw_parts)
            item_id = f"{id_prefix}{seq:05d}"
            seq += 10
            source_ref = f"系统结算价!R{r}"
            
            # 为保证 WPS / Excel 可以在 1 秒内打开且绝不崩溃：
            # 系统结算价的前 100 行作为代表嵌入截图，后续大清单行保持单元格纯净不重复堆叠数万张图片
            should_embed_img = (r <= 102)
            item_img_fn = "quote_img_系统结算价.png" if should_embed_img else ""

            item = CanonicalMaterialItem(
                quote_id=item_id, supplier=SUPPLIER, brand=brand, category_name=name_str, category_code="170101",
                structured_specs=struct_dict, raw_spec=raw_spec, file_page="第6页",
                quote_image_filename=item_img_fn, image_filenames=[item_img_fn] if item_img_fn else [],
                unit=unit_str, market_price=f_price, is_tax_included="含税", province="湖北省", city="武汉市",
                source_ref=source_ref, raw_price=f_price, notes=["系统结算价清单导入"]
            )
            canonical_items.append(item)
            source_rows.append(SourceRowRecord(source_ref=source_ref, sheet="系统结算价", row=r, raw_text=f"{name_str} | {model_str} | {spec_str} | {f_price}", raw_price=f_price, decision="物料行"))

    pic_wire = "https://gcj-pub.oss-cn-beijing.aliyuncs.com/opcrm/pic_pictures/m0019459/72de567f0bb19948.png"
    folder_name = f"{SUPPLIER}_图片库"
    for it in canonical_items:
        it.picture_url = pic_wire
        it.picture_urls = [pic_wire]
        it.illustration_folder = folder_name

    out_name = f"{SUPPLIER}-标准解析.xlsx"
    out_path = os.path.join(DELIVERY_DIR, out_name)
    
    print(f"[*] 执行官方原生 StandardizedWorkbookBuilder 生成文件 (img_dir={IMG_DIR})...")
    builder = StandardizedWorkbookBuilder(template_path=DEFAULT_TEMPLATE_PATH)
    builder.build_from_canonical_items(canonical_items, out_path, IMG_DIR, reconciliation=source_rows)
    print(f"[✓] 成功构建可秒开的高性能工作簿: {out_path} (大小: {os.path.getsize(out_path)} 字节)")

if __name__ == "__main__":
    run()
