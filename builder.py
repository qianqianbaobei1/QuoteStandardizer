# -*- coding: utf-8 -*-
"""
更新标准 Excel 生成与构建器
1. 强制启用单元格内 \\n 多行结构化排版
2. 自动根据结构化行数与是否有图片动态计算最佳行高
3. 单元格内边距与自动换行优化 (wrap_text=True)
4. 保证后续任意厂商数据持久化时均严格遵循此多行规范
"""
import io
import os
import re
import openpyxl
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from PIL import Image as PILImage
from models import (CanonicalMaterialItem, SourceRowRecord, GROUP_ADVICE, SEVERITY_ORDER, classify_note, note_field,
                    summarize_audit)

# 未配置模板时使用的内置标准表头 (GenOffice 33 列)
# 用户确认唯一标准 33 列表头规范
STANDARD_HEADERS = [
    '报价ID', '供应商名称', '名称规格', '类别编码', '品牌', '规格参数', '厂家原始规格参数', '厂家报价文件页码',
    '厂家报价图片文件名', '顺序号', '报价地区省', '报价地区市', '单位', '市场价', '工程价', '工程价折扣比例',
    '折扣下限', '折扣上限', '折扣系列', '运费说明', '是否含运费', '税费说明', '增值税税率', '运费结算方式',
    '是否含增值税', '报价说明', '价格填报类型', '图片文件夹', '图片', '图片', '图片', '图片', '图片'
]

# 汇总主表【标准化报价】基础列规范（完整保留 33 列全要素，并衔接工程与时间字段）
MAIN_PRE_ATTR_HEADERS = [
    '报价ID', '供应商名称', '名称规格', '类别编码', '品牌', '规格参数', '厂家原始规格参数', '厂家报价文件页码',
    '厂家报价图片文件名', '顺序号', '报价地区省', '报价地区市', '单位', '市场价', '工程价', '工程价折扣比例',
    '折扣下限', '折扣上限', '折扣系列', '运费说明', '是否含运费', '税费说明', '增值税税率', '运费结算方式',
    '是否含增值税', '报价说明', '价格填报类型',
    '项目类型', '工程用量', '报价时间'
]

# 标杆线上系统基础 27 列 (AA 列为 报价时间 锚点)
BENCHMARK_PRE_ATTR_HEADERS = [
    '报价id', '名称规格', '品牌', '厂家原始规格参数', '厂家报价文件页码', '厂家报价图片文件名',
    '顺序号', '报价地区省', '报价地区市', '单位', '市场价', '工程价', '工程价折扣比例',
    '折扣下限', '折扣上限', '折扣系列', '是否含运费', '运费说明', '是否含增值税', '税费说明',
    '增值税税率%', '运费结算方式', '报价说明', '价格填报类型',
    '项目类型', '工程用量', '报价时间'
]
BENCHMARK_TAIL_HEADERS = ['图片文件夹', '图片', '图片', '图片', '图片', '图片']

def collect_dynamic_attribute_headers(items: list) -> list:
    """
    根据本批物料的分类与提取参数，召回广材官方三级品类标准属性，并动态编排一属性一列的表头序列
    """
    from category_attribute_service import get_category_standard_attr_names
    ordered_attrs = []
    seen = set()

    # 1. 优先按官方 Schema 推荐顺序召回标准属性
    for item in items:
        cat_key = getattr(item, 'category_code', '') or getattr(item, 'category_name', '')
        if cat_key:
            std_names = get_category_standard_attr_names(cat_key)
            for name in std_names:
                n_clean = str(name).strip()
                if n_clean and n_clean not in seen and not any(f in n_clean for f in FORBIDDEN_SPEC_KEYS):
                    seen.add(n_clean)
                    ordered_attrs.append(n_clean)

    # 2. 补齐物料中已成功提取但未在首批 Schema 中的物理规格键
    for item in items:
        specs = getattr(item, 'structured_specs', {}) or {}
        for k in specs.keys():
            k_clean = str(k).strip()
            if k_clean and k_clean not in seen and not any(f in k_clean for f in FORBIDDEN_SPEC_KEYS):
                seen.add(k_clean)
                ordered_attrs.append(k_clean)

    # 3. 排序微调：确保 '品种' 排在属性最前，'产品描述' 或 '详细描述' 排在属性末尾
    if '品种' in ordered_attrs:
        ordered_attrs.remove('品种')
        ordered_attrs.insert(0, '品种')
    if '产品描述' in ordered_attrs:
        ordered_attrs.remove('产品描述')
        ordered_attrs.append('产品描述')

    return ordered_attrs

CENTER_HEADERS = {'报价ID', '报价id', '类别编码', '顺序号', '报价地区省', '报价地区市', '单位', '工程价折扣比例', '折扣下限',
                  '折扣上限', '折扣系列', '是否含运费', '增值税税率', '增值税税率%', '运费结算方式', '是否含增值税',
                  '价格填报类型', '项目类型', '报价时间'}
PRICE_HEADERS = {'市场价', '工程价'}
# 规格参数里不许出现的“参数名”关键词（基价/调价属于商务信息，产品编码/物料编码属于内部代码，严禁混入物理规格）
FORBIDDEN_SPEC_KEYS = ('表头', '表标题', '基价', '单价', '总价', '金额', '指导价', '调价',
                       '产品编码', '物料编码', '商品编码', '内部编码', '条形码', '编码', '条码', '代码',
                       '税号', '账号', '帐号', '开户行', '一级类别', '二级类别', '三级类别', '材料名称', '属性项属性值',
                       '物料名称', '类别名称', '分类编码', '原始描述', '列标题', '行标题')

def safe_excel_val(val):
    """确保单元格内容不含非法控制字符，统一换行符，杜绝乱码与损坏"""
    if val is None:
        return ""
    if isinstance(val, str):
        # 过滤无法写入 Excel 的底层 ASCII 控制字符，杜绝 IllegalCharacterError
        val = ILLEGAL_CHARACTERS_RE.sub('', val)
        # 统一规范换行符
        val = val.replace('\r\n', '\n').replace('\r', '\n')
    return val

class StandardizedWorkbookBuilder:
    def __init__(self, template_path=None, auto_expand_attributes=False, split_by_category=False):
        self.template_path = template_path
        self.auto_expand_attributes = auto_expand_attributes
        self.split_by_category = split_by_category
        self._load_template()

    def _load_template(self):
        if self.template_path and os.path.exists(self.template_path):
            wb_tpl = openpyxl.load_workbook(self.template_path)
            ws_tpl = wb_tpl.active
            self.headers = [ws_tpl.cell(1, c).value or '' for c in range(1, ws_tpl.max_column + 1)]
            # 去掉末尾的空表头列
            while self.headers and not self.headers[-1]:
                self.headers.pop()
        else:
            if self.template_path:
                print(f"[Builder Warning] 模板不存在: {self.template_path}，使用内置标准表头。")
            self.headers = list(STANDARD_HEADERS)
        missing = [h for h in STANDARD_HEADERS if h not in self.headers]
        if missing:
            print(f"[Builder Warning] 模板缺少标准列: {missing}，这些字段不会输出。")
        dynamic = [h for h in self.headers if h and h not in STANDARD_HEADERS]
        if dynamic:
            print(f"[*] 模板动态属性列（按表头名从结构化参数取值）: {dynamic}")

    @staticmethod
    def _norm_header(text):
        """表头/参数名归一：去掉括号里的单位说明与空白标点，厚度(mm) 与 厚度 视为同一列"""
        s = re.sub(r'[（(][^）)]*[）)]', '', str(text or ''))
        return re.sub(r'[\s:：*()（）\-_/]', '', s).lower()

    def _dynamic_spec_values(self, item, known_headers):
        """
        模板里的动态属性列 (品种 / 厚度(mm) / 胎基材料 / 覆面材料 / 是否曝露使用 / 容量 …)
        按表头名到结构化参数里取值 —— 不看列位置，因此同一套代码能同时匹配
        112101 (AD=胎基材料) 与 112115 (AD=是否曝露使用) 这类列序不同的模板。
        取不到的列留空，不猜。
        """
        out = {}
        specs = {k: v for k, v in (item.structured_specs or {}).items() if v not in (None, '')}
        index = {}
        for k, v in specs.items():
            index.setdefault(self._norm_header(k), v)
        for header in self.headers:
            if not header or header in known_headers:
                continue
            if any(k in header for k in FORBIDDEN_SPEC_KEYS):
                continue
            key = self._norm_header(header)
            val = index.get(key)
            if val is None and len(key) >= 2:
                # 只认两种情况：参数名是表头的一部分，或参数名以表头开头（容量 ↔ 容量kVA）。
                # 不允许反向前缀（表头“材料”不能命中参数“胎基材料”），否则会静默填错值
                for sk, sv in index.items():
                    if any(k in sk for k in FORBIDDEN_SPEC_KEYS):
                        continue
                    if len(sk) >= 2 and (sk in key or sk.startswith(key)):
                        val = sv
                        break
            if val is not None:
                out[header] = val
        return out

    @staticmethod
    def _row_values(item, seq, formatted_specs, first_img_name, img_folder_name, final_quote_desc):
        """严格按标准 33 列表头字典映射，绝不错列，绝不虚构内容！"""
        # 格式化税率
        tax_val = item.tax_rate if item.tax_rate is not None else ''

        # 图片列映射（5列“图片”）
        # 仅使用材料配图接口返回的在线图片链接！未命中的保持为空，绝不多加、不编造任何假内容，绝不把底稿图作为配图写入！
        pic_urls = getattr(item, 'picture_urls', []) or ([item.picture_url] if getattr(item, 'picture_url', None) else [])
        img1 = pic_urls[0] if len(pic_urls) > 0 else ''
        img2 = pic_urls[1] if len(pic_urls) > 1 else ''
        img3 = pic_urls[2] if len(pic_urls) > 2 else ''
        img4 = pic_urls[3] if len(pic_urls) > 3 else ''
        img5 = pic_urls[4] if len(pic_urls) > 4 else ''

        # 厂家报价图片文件名：记录底稿截图文件名
        quote_img_file = getattr(item, 'quote_image_filename', '') or ''
        if not quote_img_file:
            if hasattr(item, 'image_filenames') and item.image_filenames:
                for fn in item.image_filenames:
                    if str(fn).lower().startswith('quote_img'):
                        quote_img_file = fn
                        break
        if not quote_img_file and first_img_name and str(first_img_name).lower().startswith('quote_img'):
            quote_img_file = first_img_name
        if not quote_img_file and item.file_page:
            p_str = str(item.file_page).strip()
            p_m = re.search(r'\d+', p_str)
            if p_str.startswith('第') and p_str.endswith('页') and p_m:
                quote_img_file = f"quote_img{p_m.group(0)}.png"
            elif p_str.isdigit():
                quote_img_file = f"quote_img{p_str}.png"
            else:
                clean_s = re.sub(r'[\/\\:\*\?\"<>\|]', '_', p_str).strip()
                quote_img_file = f"quote_img_{clean_s}.png"

        # 图片文件夹：来自配图接口返回的目录名（如 m0000578），未命中则留空
        folder_val = getattr(item, 'illustration_folder', '') or ''

        # 彻底清洗厂家原始规格参数：剥除混入的产品编码/物料编码等内部代码与非物理参数
        clean_raw_spec = re.sub(r'(?:^|[;；\s])(?:产品编码|物料编码|商品编码|内部编码|条形码|条码|代码|三级类别编码|二级类别|一级类别|材料名称|属性项属性值)\s*[:：][^;；\n]*', '', item.raw_spec or '')
        clean_raw_spec = re.sub(r'^[;；\s]+|[;；\s]+$', '', clean_raw_spec).strip()

        return {
            '报价ID': str(item.quote_id),
            '报价id': str(item.quote_id),
            '供应商名称': item.supplier or '',
            '名称规格': item.category_name or '',
            '类别编码': item.category_code or '',
            '品牌': item.brand or '',
            '规格参数': formatted_specs or '',
            '厂家原始规格参数': clean_raw_spec,
            '厂家报价文件页码': item.file_page or '',
            '厂家报价图片文件名': quote_img_file,
            '_quote_img_file': quote_img_file,
            '顺序号': seq,
            '报价地区省': item.province or '',
            '报价地区市': item.city or '',
            '单位': item.unit or '',
            '市场价': item.market_price,
            '工程价': item.market_price,  # 严格遵循红线规范：工程价保持原表原价绝对保真，折扣独立记录在折扣比例列，绝不直接写折后价！
            '工程价折扣比例': item.discount_ratio if item.discount_ratio is not None else None,
            '折扣下限': item.discount_min if item.discount_min is not None else item.discount_ratio,
            '折扣上限': item.discount_max if item.discount_max is not None else item.discount_ratio,
            '折扣系列': item.discount_series or '',
            '运费说明': item.freight_desc or '',
            '是否含运费': item.is_freight_included or '',
            '税费说明': item.tax_desc or '',
            '增值税税率': tax_val,
            '增值税税率%': tax_val,
            '运费结算方式': item.freight_settle or '',
            '是否含增值税': item.is_tax_included or '',
            '报价说明': final_quote_desc or '',
            '价格填报类型': item.quote_type or '',
            '项目类型': getattr(item, 'project_type', '') or '',
            '工程用量': getattr(item, 'project_volume', '') or '',
            '报价时间': getattr(item, 'quote_date', '') or '',
            '图片文件夹': folder_val,
            '图片文件': folder_val,
            '_img1': img1,
            '_img2': img2,
            '_img3': img3,
            '_img4': img4,
            '_img5': img5,
        }

    def _render_material_sheet(self, ws_out, items: list, headers: list, dynamic_attr_columns: list,
                               img_dir: str, _quote_img_cache: dict, thin_border):
        """
        渲染单个物料工作表：表头样式、斑马纹行、单元格格式化、缩略图内嵌与列宽自适应
        返回: (embedded_img_count, list_of_uncertain_notes)
        """
        ws_out.views.sheetView[0].showGridLines = True
        ws_out.row_dimensions[1].height = 36

        header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        header_font = Font(name='宋体', size=11, bold=True, color='FFFFFF')
        header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

        header_time_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')  # 浅绿
        header_time_font = Font(name='宋体', size=11, bold=True, color='385723')

        header_attr_fill = PatternFill(start_color='C00000', end_color='C00000', fill_type='solid')  # 橙红
        header_attr_font = Font(name='宋体', size=11, bold=True, color='FFFFFF')

        header_desc_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
        header_desc_font = Font(name='宋体', size=11, bold=True, color='385723')

        for col_idx, h in enumerate(headers, 1):
            c = ws_out.cell(row=1, column=col_idx, value=h)
            c.border = thin_border
            c.alignment = header_align
            if h == '报价时间':
                c.fill = header_time_fill
                c.font = header_time_font
            elif h == '产品描述':
                c.fill = header_desc_fill
                c.font = header_desc_font
            elif h in dynamic_attr_columns:
                c.fill = header_attr_fill
                c.font = header_attr_font
            else:
                c.fill = header_fill
                c.font = header_font

        body_font = Font(name='宋体', size=10)
        warning_font = Font(name='宋体', size=10, color='C00000', bold=True)
        center_align = Alignment(horizontal='center', vertical='center')
        multiline_left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
        right_align = Alignment(horizontal='right', vertical='center')

        alt_fill = PatternFill(start_color='F9FAFC', end_color='F9FAFC', fill_type='solid')
        white_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        warning_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')

        seq = 10
        row_num = 2
        total_imgs_embedded = 0
        uncertain_items_for_audit = []

        quote_img_col_idx = None
        for c_idx, h in enumerate(headers, 1):
            if h == '厂家报价图片文件名':
                quote_img_col_idx = c_idx
                break

        for item in items:
            formatted_multiline_specs = item.format_multiline_specs()
            clean_lines = []
            for l in formatted_multiline_specs.split('\n'):
                spec_key = l.split(':', 1)[0].strip()
                if any(k in spec_key for k in FORBIDDEN_SPEC_KEYS):
                    continue
                clean_lines.append(l)
            formatted_multiline_specs = '\n'.join(clean_lines) if clean_lines else ''
            line_count = max(1, formatted_multiline_specs.count('\n') + 1)

            saved_images = []
            if item.image_bytes_list and img_dir:
                os.makedirs(img_dir, exist_ok=True)
                for idx, img_bytes in enumerate(item.image_bytes_list):
                    ext = self._detect_ext(img_bytes)
                    safe_name = re.sub(r'[\/\\:\*\?\"<>\|]', '_', item.category_name)
                    suffix = f'_{idx+1}' if len(item.image_bytes_list) > 1 else ''
                    img_fn = f"{item.quote_id}_{safe_name}{suffix}{ext}"
                    img_dst = os.path.join(img_dir, img_fn)
                    if not os.path.exists(img_dst):
                        with open(img_dst, 'wb') as f_out:
                            f_out.write(img_bytes)
                    saved_images.append((img_fn, img_dst))

            first_img_name = saved_images[0][0] if saved_images else (
                item.image_filenames[0] if (hasattr(item, 'image_filenames') and item.image_filenames) else '')
            has_img_ref = bool(
                saved_images or first_img_name or (hasattr(item, 'image_filenames') and item.image_filenames) or getattr(
                    item, 'quote_image_filename', ''))
            img_folder_name = os.path.basename(img_dir) if (has_img_ref and img_dir) else ''
            final_quote_desc = (item.quote_desc or '').strip()

            values = self._row_values(item, seq, formatted_multiline_specs, first_img_name,
                                      img_folder_name, final_quote_desc)
            values.update(self._dynamic_spec_values(item, values))
            needs_human = item.worst_severity() in ('致命', '严重')
            row_fill = alt_fill if (row_num % 2 == 0) else white_fill

            img_col_counter = 1
            for c_idx, header in enumerate(headers, 1):
                if header == '图片':
                    val = values.get(f'_img{img_col_counter}', '')
                    img_col_counter += 1
                elif header == '厂家报价图片文件名':
                    val = values.get('厂家报价图片文件名', '')
                elif header in ('图片文件夹', '图片文件'):
                    val = values.get('图片文件夹') or values.get('图片文件', '')
                else:
                    val = values.get(header, '')
                val = safe_excel_val(val)
                cell = ws_out.cell(row=row_num, column=c_idx, value=val)
                cell.font = body_font
                cell.fill = row_fill
                cell.border = thin_border

                if header in CENTER_HEADERS or header == '厂家报价图片文件名':
                    cell.alignment = center_align
                elif header in PRICE_HEADERS:
                    cell.alignment = right_align
                    if isinstance(val, (int, float)):
                        cell.number_format = '0.000'
                else:
                    cell.alignment = multiline_left_align

                if header == '图片' and isinstance(val, str) and (
                        val.startswith('http://') or val.startswith('https://')):
                    cell.hyperlink = val
                    cell.font = Font(name='宋体', size=10, color='0563C1', underline='single')

                if needs_human:
                    cell.fill = warning_fill
                    cell.font = warning_font
                    if header == '市场价':
                        cell.comment = Comment(f"【{item.worst_severity()}】{item.uncertainty_note}", "造价审计员")

            calc_text_height = max(30, line_count * 18 + 10)
            has_embedded_quote_img = False

            target_quote_img_fn = values.get('_quote_img_file') or getattr(item, 'quote_image_filename', '')
            if quote_img_col_idx and target_quote_img_fn and img_dir:
                if target_quote_img_fn not in _quote_img_cache:
                    img_dst = os.path.join(img_dir, target_quote_img_fn)
                    if not os.path.exists(img_dst) and os.path.exists(img_dir):
                        for c_fn in os.listdir(img_dir):
                            if c_fn.lower() == target_quote_img_fn.lower():
                                img_dst = os.path.join(img_dir, c_fn)
                                break
                    if os.path.exists(img_dst) and os.path.getsize(img_dst) > 0:
                        retina_thumb_fn = f"thumb_retina_{os.path.splitext(os.path.basename(img_dst))[0]}.jpg"
                        retina_thumb_dst = os.path.join(img_dir, retina_thumb_fn)
                        if not os.path.exists(retina_thumb_dst) or os.path.getsize(retina_thumb_dst) == 0:
                            try:
                                with PILImage.open(img_dst) as pil_orig:
                                    im_rgb = pil_orig.convert('RGB')
                                    im_rgb.thumbnail((800, 600), PILImage.Resampling.LANCZOS)
                                    im_rgb.save(retina_thumb_dst, 'JPEG', quality=95, optimize=True)
                            except Exception:
                                pass
                        w, h = 120, 68
                        if os.path.exists(retina_thumb_dst):
                            try:
                                with PILImage.open(retina_thumb_dst) as pil_img:
                                    w, h = pil_img.size
                                    target_h = 68
                                    target_w = int(w * (target_h / h))
                                    if target_w > 120:
                                        target_w = 120
                                        target_h = int(h * (target_w / w))
                                    w, h = target_w, target_h
                            except Exception:
                                pass
                        rel_link = f"{img_folder_name}/{os.path.basename(img_dst)}" if img_folder_name else os.path.basename(
                            img_dst)
                        _quote_img_cache[target_quote_img_fn] = (img_dst, retina_thumb_dst, w, h, rel_link)
                    else:
                        _quote_img_cache[target_quote_img_fn] = None

                cached_info = _quote_img_cache.get(target_quote_img_fn)
                if cached_info:
                    img_dst, retina_thumb_dst, target_w, target_h, rel_link = cached_info
                    if total_imgs_embedded < 50 and os.path.exists(retina_thumb_dst):
                        try:
                            op_img = OpenpyxlImage(retina_thumb_dst)
                            op_img.width = target_w
                            op_img.height = target_h
                            col_letter = get_column_letter(quote_img_col_idx)
                            ws_out.add_image(op_img, f"{col_letter}{row_num}")
                            total_imgs_embedded += 1
                            has_embedded_quote_img = True
                        except Exception:
                            pass
                    cell_quote = ws_out.cell(row=row_num, column=quote_img_col_idx)
                    cell_quote.value = target_quote_img_fn
                    cell_quote.hyperlink = rel_link
                    cell_quote.font = Font(name='宋体', size=9, color='0563C1', underline='single')

            if has_embedded_quote_img:
                ws_out.row_dimensions[row_num].height = max(68, calc_text_height)
            else:
                ws_out.row_dimensions[row_num].height = calc_text_height

            for note in item.note_list():
                severity, group = classify_note(note)
                uncertain_items_for_audit.append({
                    "quote_id": item.quote_id,
                    "severity": severity,
                    "group": group,
                    "field": note_field(note, headers),
                    "model": item.structured_specs.get("型号") or item.category_name,
                    "source_ref": item.source_ref,
                    "page": item.file_page,
                    "price": item.market_price,
                    "note": note
                })

            row_num += 1
            seq += 10

        for col_idx, header_val in enumerate(headers, 1):
            col_letter = get_column_letter(col_idx)
            header_val = str(header_val or '')

            if header_val == '厂家报价图片文件名':
                ws_out.column_dimensions[col_letter].width = 18
            elif header_val == '图片' or any(k in header_val for k in ['图片1', '图片2', '图片3', '图片4', '图片5', '报价图片']):
                ws_out.column_dimensions[col_letter].width = 24
            elif '厂家原始规格参数' in header_val:
                ws_out.column_dimensions[col_letter].width = 42
            elif '规格参数' in header_val:
                ws_out.column_dimensions[col_letter].width = 48
            elif any(k in header_val for k in ['供应商名称', '报价说明', '运费说明', '税费说明']):
                ws_out.column_dimensions[col_letter].width = 32
            elif any(k in header_val for k in ['名称规格', '材料名称', '品名']):
                ws_out.column_dimensions[col_letter].width = 24
            elif any(k in header_val for k in ['品牌', '图片文件夹', '报价ID']):
                ws_out.column_dimensions[col_letter].width = 16
            elif any(k in header_val for k in ['市场价', '工程价', '单价', '金额']):
                ws_out.column_dimensions[col_letter].width = 15
            else:
                ws_out.column_dimensions[col_letter].width = 13

        return total_imgs_embedded, uncertain_items_for_audit

    def build_from_canonical_items(self, items: list, output_xlsx_path: str, img_dir: str, reconciliation: list = None,
                                   structure_log: list = None):
        """
        基于 CanonicalMaterialItem 统一物料模型进行企业级输出构建：
        1. 保留【客户原始解析】原表全量对账页
        2. 保留【标准化报价】全量汇总主表
        3. 对于不同属性值的材料（不同三级分类编码），自动按分类拆分独立 Sheet 保存，每个 Sheet 专属展开各自的标准属性列
        """
        if not items:
            raise ValueError(f"物料列表为空，严禁生成或输出空的报价表格文件: {output_xlsx_path}")

        # 若启用自适应动态属性列且未传特定静态模板，按物料标准分类自动扩展一属性一列表头
        active_headers = list(self.headers)
        dynamic_attr_columns = []
        if self.auto_expand_attributes and not self.template_path:
            dynamic_attr_columns = collect_dynamic_attribute_headers(items)
            # 汇总主表【标准化报价】：完整保留 33 列全要素（供应商名称、类别编码、规格参数等）+ 报价时间 + 动态属性列
            active_headers = list(MAIN_PRE_ATTR_HEADERS) + dynamic_attr_columns + list(BENCHMARK_TAIL_HEADERS)
            self.headers = active_headers

        # 图片目录按需延迟创建，杜绝在磁盘生成无用的空图片库文件夹
        has_any_image = any(len(it.image_bytes_list) > 0 for it in items)
        if has_any_image:
            os.makedirs(img_dir, exist_ok=True)

        wb_out = openpyxl.Workbook()
        ws_out = wb_out.active
        ws_out.title = '标准化报价'

        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        _quote_img_cache = {}

        # 1. 渲染全量物料主表【标准化报价】
        total_imgs_embedded, uncertain_items_for_audit = self._render_material_sheet(
            ws_out, items, self.headers, dynamic_attr_columns, img_dir, _quote_img_cache, thin_border
        )

        # 2. 对于不同属性值的材料：按三级分类编码（或品名）拆分多个专属 Sheet
        category_groups = {}
        for it in items:
            code = (getattr(it, 'category_code', '') or '').strip()
            if not code:
                code = (getattr(it, 'category_name', '') or '').strip() or '未分类'
            category_groups.setdefault(code, []).append(it)

        category_sheet_names = []
        # 只要包含多种不同属性分类材料，或显式要求按品类分表时，自动生成独立分类 Sheet
        if len(category_groups) > 1 or self.split_by_category:
            for cat_code, cat_items in category_groups.items():
                safe_title = re.sub(r'[\/\\:\*\?\[\]]', '_', str(cat_code)).strip() or '分类明细'
                safe_title = safe_title[:31]
                if safe_title in ('客户原始解析', '标准化报价'):
                    safe_title = f"{safe_title}_明细"[:31]

                base_title = safe_title
                cnt = 1
                while safe_title in wb_out.sheetnames:
                    safe_title = f"{base_title[:28]}_{cnt}"
                    cnt += 1

                ws_cat = wb_out.create_sheet(title=safe_title)
                # 分类 Sheet 专有属性列：只展开属于该分类的标准属性 Schema 与提取到的属性
                cat_dyn_cols = collect_dynamic_attribute_headers(cat_items) if (
                            self.auto_expand_attributes and not self.template_path) else []
                cat_headers = (list(BENCHMARK_PRE_ATTR_HEADERS) + cat_dyn_cols + list(
                    BENCHMARK_TAIL_HEADERS)) if cat_dyn_cols else list(self.headers)
                cat_imgs, _ = self._render_material_sheet(
                    ws_cat, cat_items, cat_headers, cat_dyn_cols, img_dir, _quote_img_cache, thin_border
                )
                total_imgs_embedded += cat_imgs
                category_sheet_names.append(safe_title)

        # 真正兑现双 Sheet 标准交付架构：生成 Sheet 1【客户原始解析】
        ws_raw = wb_out.create_sheet(title='客户原始解析', index=0)
        ws_raw.views.sheetView[0].showGridLines = True
        raw_headers = ['序号', '工作表/页码', '原表行标识', '原表文本内容', '原表单价', '解析判定', '说明备注']
        ws_raw.row_dimensions[1].height = 32

        body_font = Font(name='宋体', size=10)
        center_align = Alignment(horizontal='center', vertical='center')
        multiline_left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
        right_align = Alignment(horizontal='right', vertical='center')
        alt_fill = PatternFill(start_color='F9FAFC', end_color='F9FAFC', fill_type='solid')
        white_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

        raw_header_fill = PatternFill(start_color='203764', end_color='203764', fill_type='solid')
        raw_header_font = Font(name='宋体', size=11, bold=True, color='FFFFFF')
        for c_idx, rh in enumerate(raw_headers, 1):
            c = ws_raw.cell(row=1, column=c_idx, value=rh)
            c.fill = raw_header_fill
            c.font = raw_header_font
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = thin_border

        raw_row_idx = 2
        rec_list = reconciliation or []
        if not rec_list:
            for it in items:
                rec_list.append(SourceRowRecord(
                    source_ref=getattr(it, 'source_ref', '') or f"R{raw_row_idx}",
                    sheet=getattr(it, 'file_page', '') or '原表',
                    row=raw_row_idx,
                    raw_text=getattr(it, 'raw_full_line', '') or getattr(it, 'raw_spec', ''),
                    raw_price=getattr(it, 'market_price', 0.0),
                    decision='item'
                ))

        for idx, rec in enumerate(rec_list, 1):
            r_fill = alt_fill if (raw_row_idx % 2 == 0) else white_fill
            s_ref = getattr(rec, 'source_ref', '')
            s_sheet = getattr(rec, 'sheet', '')
            s_text = getattr(rec, 'raw_text', '')
            s_price = getattr(rec, 'raw_price', None)
            s_dec = getattr(rec, 'decision', 'item')
            s_note = getattr(rec, 'issue', '') or ('' if s_dec == 'item' else s_dec)

            row_data = [idx, s_sheet, s_ref, safe_excel_val(s_text), s_price, s_dec, safe_excel_val(s_note)]
            for col_i, val in enumerate(row_data, 1):
                cell = ws_raw.cell(row=raw_row_idx, column=col_i, value=val)
                cell.font = body_font
                cell.fill = r_fill
                cell.border = thin_border
                if col_i in (1, 2, 3, 6):
                    cell.alignment = center_align
                elif col_i == 5 and isinstance(val, (int, float)):
                    cell.alignment = right_align
                    cell.number_format = '0.00'
                else:
                    cell.alignment = multiline_left_align
            raw_row_idx += 1

        ws_raw.column_dimensions['A'].width = 8
        ws_raw.column_dimensions['B'].width = 14
        ws_raw.column_dimensions['C'].width = 16
        ws_raw.column_dimensions['D'].width = 65
        ws_raw.column_dimensions['E'].width = 14
        ws_raw.column_dimensions['F'].width = 18
        ws_raw.column_dimensions['G'].width = 30

        # 交付规范：保留【客户原始解析】、【标准化报价】与分类专属 Sheet，清除多余的调试/临时对账 Sheet
        allowed_sheets = {'客户原始解析', '标准化报价'} | set(category_sheet_names)
        for s in list(wb_out.sheetnames):
            if s not in allowed_sheets:
                wb_out.remove(wb_out[s])

        wb_out.save(output_xlsx_path)

        # 垃圾文件防护：若全表无任何图片生成，自动清除空的临时图片目录，保持交付区清爽整洁
        if os.path.exists(img_dir) and not os.listdir(img_dir):
            try:
                os.rmdir(img_dir)
            except Exception:
                pass

        print(f"[Builder Success] 规范构建生成完成: {output_xlsx_path}, 共内嵌 {total_imgs_embedded} 张图片。")

    @staticmethod
    def _detect_ext(img_bytes: bytes) -> str:
        """按真实格式确定图片扩展名 (原逻辑一律存 .png，jpeg/emf 会被错误命名)"""
        try:
            with PILImage.open(io.BytesIO(img_bytes)) as im:
                fmt = (im.format or 'png').lower()
                return '.jpg' if fmt in ('jpeg', 'mpo') else f'.{fmt}'
        except Exception:
            if img_bytes[:4] == b'\x01\x00\x00\x00':
                return '.emf'
            return '.bin'

    @staticmethod
    def _build_reconciliation_sheet(wb_out, recon: list, thin_border):
        """原表 ↔ 标准化结果逐行对账页 (含数学勾稽)"""
        ws = wb_out.create_sheet(title='原表对账')
        total = sum(1 for r in recon if r['status'] != '跳过')
        bad = sum(1 for r in recon if r['status'].startswith('异常'))
        skipped = sum(1 for r in recon if r['status'] == '跳过')
        ws.cell(row=1, column=1, value=f"原表物理行 {len(recon)} 行 | 物料输出 {total} 行 | 跳过 {skipped} 行 | 对账异常 {bad} 行").font = \
            Font(name='宋体', size=11, bold=True, color='C00000' if bad else '385623')

        headers = ['原表位置', '对账结果', '数学勾稽', '报价ID', '原表单价', '输出单价', '说明/跳过原因', '原表整行内容']
        hfill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=2, column=c, value=h)
            cell.fill = hfill
            cell.font = Font(name='宋体', size=11, bold=True, color='FFFFFF')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border

        red = Font(name='宋体', size=10, color='C00000', bold=True)
        grey = Font(name='宋体', size=10, color='808080')
        normal = Font(name='宋体', size=10)
        # 异常排在最前
        order = {'异常': 0, '一致(有待复核项)': 1, '一致': 2, '跳过': 3}
        rows = sorted(recon, key=lambda r: order.get(r['status'].split(':')[0], order.get(r['status'], 1)))
        for i, r in enumerate(rows, 3):
            vals = [r['source_ref'], r['status'], r.get('math_audit', ''), r['quote_id'], r['raw_price'], r['out_price'], r['detail'], r['raw_text']]
            font = red if r['status'].startswith('异常') else (grey if r['status'] == '跳过' else normal)
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=i, column=c, value=safe_excel_val(v))
                cell.font = font
                cell.border = thin_border
                cell.alignment = Alignment(vertical='center', wrap_text=c in (7, 8))
                if c in (5, 6) and isinstance(v, (int, float)):
                    cell.number_format = '0.000'
        for col, w in zip('ABCDEFGH', [20, 24, 28, 14, 12, 12, 45, 65]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = 'A3'

    @classmethod
    def build_integrity_certificate(cls, report_path: str, recon: list, summary: dict, supplier: str, source_file: str):
        """
        生成独立的《防欺瞒体检报告与对账白皮书》(投行深蓝美学风格)
        """
        wb = openpyxl.Workbook()
        ws_cert = wb.active
        ws_cert.title = '防欺瞒体检指标概览'
        ws_cert.views.sheetView[0].showGridLines = True

        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
        )
        navy_header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        header_font = Font(name='宋体', size=11, bold=True, color='FFFFFF')
        kpi_title_font = Font(name='宋体', size=12, bold=True, color='1F4E78')
        kpi_val_font = Font(name='Arial', size=20, bold=True, color='385623' if summary.get('is_zero_defect') else 'C00000')

        ws_cert.merge_cells('A1:G1')
        t_cell = ws_cert.cell(row=1, column=1, value=f"【零信任防欺瞒体检证书】{supplier} 报价单工程解析")
        t_cell.font = Font(name='宋体', size=14, bold=True, color='1F4E78')
        t_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws_cert.row_dimensions[1].height = 40

        ws_cert.merge_cells('A2:G2')
        sub_cell = ws_cert.cell(row=2, column=1, value=f"原始文件: {os.path.basename(source_file)} | 自动免检状态: {'【通过·零缺陷免检入库】' if summary.get('is_zero_defect') else '【存在异常·需红圈复核】'}")
        sub_cell.font = Font(name='宋体', size=10, color='595959')
        sub_cell.alignment = Alignment(horizontal='center', vertical='center')
        ws_cert.row_dimensions[2].height = 24

        # 指标卡片
        kpis = [
            ("物理行守恒率", f"{summary.get('row_conservation_rate', 0):.1f}%", "物料行+跳过行 / 原表总物理扫描行"),
            ("对账异常总数", f"{summary.get('anomalies_count', 0)} 项", "漏行/多行/价格窜改/单向漏检"),
            ("数学勾稽自洽率", f"{summary.get('math_pass_rate', 100):.1f}%", f"参检 {summary.get('math_audited_count', 0)} 行 / 自洽 {summary.get('math_passed_count', 0)} 行"),
            ("原表物理总行数", f"{summary.get('total_source_rows', 0)} 行", f"物料输出 {summary.get('matched_items', 0)} | 合法跳过 {summary.get('skipped_rows', 0)}")
        ]

        ws_cert.row_dimensions[4].height = 26
        ws_cert.row_dimensions[5].height = 36
        ws_cert.row_dimensions[6].height = 20

        for idx, (title, val, desc) in enumerate(kpis):
            col_st = idx * 2 + 1
            col_ed = col_st + 1 if idx < 3 else col_st
            if col_st < 7:
                ws_cert.merge_cells(start_row=4, start_column=col_st, end_row=4, end_column=col_ed)
                ws_cert.merge_cells(start_row=5, start_column=col_st, end_row=5, end_column=col_ed)
                ws_cert.merge_cells(start_row=6, start_column=col_st, end_row=6, end_column=col_ed)
            c1 = ws_cert.cell(row=4, column=col_st, value=title)
            c1.font = kpi_title_font
            c1.alignment = Alignment(horizontal='center', vertical='center')
            c2 = ws_cert.cell(row=5, column=col_st, value=val)
            c2.font = kpi_val_font
            c2.alignment = Alignment(horizontal='center', vertical='center')
            c3 = ws_cert.cell(row=6, column=col_st, value=desc)
            c3.font = Font(name='宋体', size=9, color='808080')
            c3.alignment = Alignment(horizontal='center', vertical='center')

        # 对账明细副表
        cls._build_reconciliation_sheet(wb, recon, thin_border)
        wb.save(report_path)


    @staticmethod
    def _build_structure_sheet(wb_out, log: list, thin_border):
        """表格结构 / 行类型：现有做法与 Jev 的判定对比 (试点)。分歧排在最前，需以原件判定谁对"""
        ws = wb_out.create_sheet(title='结构判定对比')

        def same(r):
            if not isinstance(r, dict):
                return True
            return str(r.get('current')) == str(r.get('jev'))

        diff = [r for r in log if isinstance(r, dict) and not same(r)]
        ws.cell(row=1, column=1, value=f"判定项 {len(log)} 个 | 分歧 {len(diff)} 个 (分歧项需对照原件判定哪一方正确)").font = \
            Font(name='宋体', size=11, bold=True, color='C00000' if diff else '385623')
        headers = ['Sheet', '判定项', '现有做法', 'Jev', 'Jev 置信度', '是否一致', '输出采用', '行内容', '人工判定 (对照原件填写)']
        hfill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=2, column=c, value=h)
            cell.fill = hfill
            cell.font = Font(name='宋体', size=11, bold=True, color='FFFFFF')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
        red = Font(name='宋体', size=10, color='C00000', bold=True)
        normal = Font(name='宋体', size=10)
        rows = sorted(log, key=lambda r: same(r))
        for i, r in enumerate(rows, 3):
            vals = [r['sheet'], r['item'], r.get('current'), r.get('jev'),
                    round(r['conf'], 3) if isinstance(r.get('conf'), (int, float)) else '',
                    '一致' if same(r) else '分歧', r.get('adopted', ''), r.get('row_text', ''), '']
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=i, column=c, value=safe_excel_val('' if v is None else v))
                cell.font = normal if same(r) else red
                cell.border = thin_border
                cell.alignment = Alignment(vertical='center', wrap_text=c == 8)
        for col, w in zip('ABCDEFGHI', [14, 26, 14, 14, 11, 9, 10, 70, 22]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = 'A3'


def build_from_canonical_items(items: list, output_xlsx_path: str, image_dir: str = None,
                               reconciliation: list = None, structure_log: list = None,
                               auto_expand_attributes: bool = True, split_by_category: bool = True,
                               template_path: str = None, **kwargs):
    """便捷模块级函数，包装 StandardizedWorkbookBuilder"""
    builder = StandardizedWorkbookBuilder(
        template_path=template_path,
        auto_expand_attributes=auto_expand_attributes,
        split_by_category=split_by_category
    )
    return builder.build_from_canonical_items(
        items=items,
        output_xlsx_path=output_xlsx_path,
        img_dir=image_dir or "",
        reconciliation=reconciliation,
        structure_log=structure_log
    )

