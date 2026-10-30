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
from models import CanonicalMaterialItem

# 未配置模板时使用的内置标准表头 (GenOffice 33 列)
STANDARD_HEADERS = [
    '报价ID', '供应商名称', '名称规格', '类别编码', '品牌', '规格参数', '厂家原始规格参数', '厂家报价文件页码',
    '厂家报价图片文件名', '顺序号', '报价地区省', '报价地区市', '单位', '市场价', '工程价', '工程价折扣比例',
    '折扣下限', '折扣上限', '折扣系列', '运费说明', '是否含运费', '税费说明', '增值税税率', '运费结算方式',
    '是否含增值税', '报价说明', '价格填报类型', '图片文件夹', '图片', '图片', '图片', '图片', '图片'
]
CENTER_HEADERS = {'报价ID', '类别编码', '顺序号', '报价地区省', '报价地区市', '单位', '工程价折扣比例', '折扣下限',
                  '折扣上限', '折扣系列', '是否含运费', '增值税税率', '运费结算方式', '是否含增值税', '价格填报类型'}
PRICE_HEADERS = {'市场价', '工程价'}

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
    def __init__(self, template_path):
        self.template_path = template_path
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

    @staticmethod
    def _row_values(item, seq, formatted_specs, first_img_name, img_folder_name, final_quote_desc):
        """按表头名称组织一行数据 (不依赖列位置，模板列顺序变化也不会错位)"""
        return {
            '报价ID': item.quote_id,
            '供应商名称': item.supplier,
            '名称规格': item.category_name,
            '类别编码': item.category_code,
            '品牌': item.brand,
            '规格参数': formatted_specs,
            '厂家原始规格参数': item.raw_spec,
            '厂家报价文件页码': item.file_page,
            '厂家报价图片文件名': first_img_name,
            '顺序号': seq,
            '报价地区省': item.province,
            '报价地区市': item.city,
            '单位': item.unit,
            '市场价': item.market_price,
            '工程价': item.project_price,
            '工程价折扣比例': item.discount_ratio,
            '折扣下限': item.discount_min if item.discount_min is not None else item.discount_ratio,
            '折扣上限': item.discount_max if item.discount_max is not None else item.discount_ratio,
            '折扣系列': item.discount_series,
            '运费说明': item.freight_desc,
            '是否含运费': item.is_freight_included,
            '税费说明': item.tax_desc,
            '增值税税率': item.tax_rate,
            '运费结算方式': item.freight_settle,
            '是否含增值税': item.is_tax_included,
            '报价说明': final_quote_desc,
            '价格填报类型': item.quote_type,
            '图片文件夹': img_folder_name,
        }

    def build_from_canonical_items(self, items: list, output_xlsx_path: str, img_dir: str, reconciliation: list = None,
                                   structure_log: list = None):
        """
        基于 CanonicalMaterialItem 统一物料模型进行企业级输出构建
        """
        # 图片目录按需延迟创建，杜绝在磁盘生成无用的空图片库文件夹
        has_any_image = any(len(it.image_bytes_list) > 0 for it in items)
        if has_any_image:
            os.makedirs(img_dir, exist_ok=True)

        wb_out = openpyxl.Workbook()
        ws_out = wb_out.active
        ws_out.title = '标准化报价'
        ws_out.views.sheetView[0].showGridLines = True

        # 表头：深蓝底白字，全宋体
        header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
        header_font = Font(name='宋体', size=11, bold=True, color='FFFFFF')
        header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )

        ws_out.row_dimensions[1].height = 36

        for col_idx, h in enumerate(self.headers, 1):
            c = ws_out.cell(row=1, column=col_idx, value=h)
            c.fill = header_fill
            c.font = header_font
            c.alignment = header_align
            c.border = thin_border

        # 正文严格全部统一为【宋体】
        body_font = Font(name='宋体', size=10)
        warning_font = Font(name='宋体', size=10, color='C00000', bold=True) # 红色宋体警示
        center_align = Alignment(horizontal='center', vertical='center')
        multiline_left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
        right_align = Alignment(horizontal='right', vertical='center')
        
        alt_fill = PatternFill(start_color='F9FAFC', end_color='F9FAFC', fill_type='solid')
        white_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        warning_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')

        seq = 10
        row_num = 2
        total_imgs_embedded = 0
        uncertain_items_for_audit = [] # 记录需在独立sheet中以红色字体展示的疑义与失败项

        for item in items:
            # 格式化多行结构化技术规格
            formatted_multiline_specs = item.format_multiline_specs()
            line_count = formatted_multiline_specs.count('\n') + 1

            # 落地图片资产 (按需写入)
            saved_images = []
            if item.image_bytes_list:
                os.makedirs(img_dir, exist_ok=True)
                for idx, img_bytes in enumerate(item.image_bytes_list):
                    ext = self._detect_ext(img_bytes)
                    safe_name = re.sub(r'[\/\\:\*\?\"<>\|]', '_', item.category_name)
                    suffix = f'_{idx+1}' if len(item.image_bytes_list) > 1 else ''
                    img_fn = f"{item.quote_id}_{safe_name}{suffix}{ext}"
                    img_dst = os.path.join(img_dir, img_fn)
                    with open(img_dst, 'wb') as f_out:
                        f_out.write(img_bytes)
                    saved_images.append((img_fn, img_dst))

            first_img_name = saved_images[0][0] if saved_images else (item.image_filenames[0] if item.image_filenames else '无')
            img_folder_name = os.path.basename(img_dir) if (saved_images or item.image_filenames) else ''

            # 诚实性设计：如果存在印章遮挡或模糊，显式在报价说明中留痕，严禁大模型擅自脑补
            final_quote_desc = f"【⚠️待人工复核】{item.uncertainty_note}；{item.quote_desc}" if item.uncertainty_note else item.quote_desc

            values = self._row_values(item, seq, formatted_multiline_specs, first_img_name,
                                      img_folder_name, final_quote_desc)
            row_fill = alt_fill if (row_num % 2 == 0) else white_fill

            for c_idx, header in enumerate(self.headers, 1):
                val = safe_excel_val(values.get(header, ''))  # 安全清洗非法字符，保障无损坏
                cell = ws_out.cell(row=row_num, column=c_idx, value=val)
                cell.font = body_font
                cell.fill = row_fill
                cell.border = thin_border

                # 对齐与格式控制
                if header in CENTER_HEADERS:
                    cell.alignment = center_align
                elif header in PRICE_HEADERS:
                    cell.alignment = right_align
                    if isinstance(val, (int, float)):
                        cell.number_format = '#,##0.00'
                    # 置信度不足 / 对账异常：浅黄高亮、红色宋体加粗并在市场价插入批注
                    if item.uncertainty_note:
                        cell.fill = warning_fill
                        cell.font = warning_font
                        if header == '市场价':
                            cell.comment = Comment(f"【待人工复核】{item.uncertainty_note}", "造价审计员")
                else:
                    cell.alignment = multiline_left_align

            # 动态计算最适行高:
            # 基础文本每行约 16pt，有图片时最少 68pt，多行文本时根据行数自适应提升
            calc_text_height = max(30, line_count * 18 + 10)
            img_cols = [i for i, h in enumerate(self.headers, 1) if h == '图片']
            if saved_images and img_cols:
                ws_out.row_dimensions[row_num].height = max(68, calc_text_height)
                for (fn, dst), col_target in zip(saved_images, img_cols):
                    col_letter = get_column_letter(col_target)
                    try:
                        with PILImage.open(dst) as pil_img:
                            w, h = pil_img.size
                            target_h = 70
                            target_w = int(w * (target_h / h))
                            if target_w > 105:
                                target_w = 105
                                target_h = int(h * (target_w / w))

                        op_img = OpenpyxlImage(dst)
                        op_img.width = target_w
                        op_img.height = target_h
                        ws_out.add_image(op_img, f"{col_letter}{row_num}")
                        total_imgs_embedded += 1
                    except Exception as e:
                        print(f"Error embedding {dst}: {e}")
            else:
                ws_out.row_dimensions[row_num].height = calc_text_height

            # 收集审计疑义与失败项
            if item.uncertainty_note:
                uncertain_items_for_audit.append({
                    "quote_id": item.quote_id,
                    "model": item.structured_specs.get("型号") or item.category_name,
                    "source_ref": item.source_ref,
                    "page": item.file_page,
                    "price": item.market_price,
                    "uncertainty_note": item.uncertainty_note
                })

            row_num += 1
            seq += 10

        # 列宽按表头名称设置
        for col in ws_out.columns:
            col_idx = col[0].column
            col_letter = get_column_letter(col_idx)
            header_val = str(ws_out.cell(1, col_idx).value or '')

            if header_val == '图片' or any(k in header_val for k in ['图片1', '图片2', '图片3', '图片4', '图片5', '报价图片']):
                ws_out.column_dimensions[col_letter].width = 18
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

        # 构建专属审计 Sheet: 【解析疑义与失败说明】(全部宋体，红色字体明确说明)
        ws_audit = wb_out.create_sheet(title='疑义与失败审计说明')
        ws_audit.views.sheetView[0].showGridLines = True
        
        audit_headers = ['序号', '报价ID', '品名/型号规格', '原表位置', '输出单价(元)', '审计警示/失败与疑义原因 (红色重点说明)', '建议处理方案']
        ws_audit.row_dimensions[1].height = 34
        
        audit_header_fill = PatternFill(start_color='8B0000', end_color='8B0000', fill_type='solid')
        audit_header_font = Font(name='宋体', size=11, bold=True, color='FFFFFF')
        
        for c_idx, h_text in enumerate(audit_headers, 1):
            c = ws_audit.cell(row=1, column=c_idx, value=h_text)
            c.fill = audit_header_fill
            c.font = audit_header_font
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            c.border = thin_border
        
        red_audit_font = Font(name='宋体', size=10, color='FF0000', bold=True) # 严格红色宋体
        normal_audit_font = Font(name='宋体', size=10)
        
        if not uncertain_items_for_audit:
            c = ws_audit.cell(row=2, column=1, value="本次解析未发现待复核项 (原表对账一致、全部经大模型清洗、编码均命中标准库)。")
            c.font = normal_audit_font
            ws_audit.merge_cells("A2:G2")
            ws_audit.row_dimensions[2].height = 28
        else:
            for a_idx, a_item in enumerate(uncertain_items_for_audit, start=1):
                a_row = a_idx + 1
                ws_audit.row_dimensions[a_row].height = 30
                
                row_vals = [
                    a_idx,
                    a_item['quote_id'],
                    a_item['model'],
                    a_item['source_ref'] or a_item['page'],
                    a_item['price'],
                    a_item['uncertainty_note'],
                    "对照原表对应位置人工核验；印章遮挡/模糊项向供应商索取清晰底单"
                ]
                
                for col_i, val in enumerate(row_vals, 1):
                    cell = ws_audit.cell(row=a_row, column=col_i, value=val)
                    cell.border = thin_border
                    if col_i in [1, 2, 4]:
                        cell.alignment = center_align
                        cell.font = normal_audit_font
                    elif col_i == 5:
                        cell.alignment = right_align
                        cell.font = red_audit_font
                        cell.number_format = '#,##0.00'
                    elif col_i == 6:
                        cell.alignment = multiline_left_align
                        cell.font = red_audit_font # 严格红色宋体加粗详细说明!
                    else:
                        cell.alignment = multiline_left_align
                        cell.font = normal_audit_font
        
        # 审计 Sheet 列宽自适应
        ws_audit.column_dimensions['A'].width = 8
        ws_audit.column_dimensions['B'].width = 14
        ws_audit.column_dimensions['C'].width = 28
        ws_audit.column_dimensions['D'].width = 22
        ws_audit.column_dimensions['E'].width = 18
        ws_audit.column_dimensions['F'].width = 65
        ws_audit.column_dimensions['G'].width = 45

        if reconciliation is not None:
            self._build_reconciliation_sheet(wb_out, reconciliation, thin_border)
        if structure_log:
            self._build_structure_sheet(wb_out, structure_log, thin_border)

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
        """原表 ↔ 标准化结果逐行对账页"""
        ws = wb_out.create_sheet(title='原表对账')
        total = sum(1 for r in recon if r['status'] != '跳过')
        bad = sum(1 for r in recon if r['status'].startswith('异常'))
        skipped = sum(1 for r in recon if r['status'] == '跳过')
        ws.cell(row=1, column=1, value=f"原表物料行 {total} 行 | 跳过 {skipped} 行 | 对账异常 {bad} 行").font = \
            Font(name='宋体', size=11, bold=True, color='C00000' if bad else '385623')

        headers = ['原表位置', '对账结果', '报价ID', '原表单价', '输出单价', '说明/跳过原因', '原表整行内容']
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
            vals = [r['source_ref'], r['status'], r['quote_id'], r['raw_price'], r['out_price'], r['detail'], r['raw_text']]
            font = red if r['status'].startswith('异常') else (grey if r['status'] == '跳过' else normal)
            for c, v in enumerate(vals, 1):
                cell = ws.cell(row=i, column=c, value=safe_excel_val(v))
                cell.font = font
                cell.border = thin_border
                cell.alignment = Alignment(vertical='center', wrap_text=c in (6, 7))
                if c in (4, 5) and isinstance(v, (int, float)):
                    cell.number_format = '#,##0.00'
        for col, w in zip('ABCDEFG', [22, 26, 14, 12, 12, 50, 70]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = 'A3'

    @staticmethod
    def _build_structure_sheet(wb_out, log: list, thin_border):
        """表格结构 / 行类型：现有做法与 Jev 的判定对比 (试点)。分歧排在最前，需以原件判定谁对"""
        ws = wb_out.create_sheet(title='结构判定对比')

        def same(r):
            return str(r.get('current')) == str(r.get('jev'))

        diff = [r for r in log if not same(r)]
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
