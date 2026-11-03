# -*- coding: utf-8 -*-
"""
style_enhancer.py - 系统级表格企业级美学与结构化排版引擎
严格执行：
1. 表头：深蓝底（#1F4E78）白字、宋体加粗、高 36pt、居中垂直对齐、强制换行 (wrap_text=True)
2. 动态规格列（规格参数/厂家原始规格参数/技术描述等）：
   - 将分号分隔的字符串（品种:xxx;牌号:yyy;规格:zzz）转换为结构化换行展示（每行一个“键: 值”）
   - 自动开启 wrap_text=True，靠左垂直居中，自适应行高
3. 数据行：
   - 全局细边框（#D9D9D9）
   - 隔行浅灰浅蓝底纹（斑马纹 #F9FAFC / #FFFFFF）
   - 居中列：报价ID、页码、序号、省、市、单位、折扣、税率、是否含运/税等
   - 右对齐列：市场价、工程价等金额列，标准千分位会计格式 #,##0.00
   - 左对齐列：规格参数、物料名称、品牌、说明等
4. 图片列：
   - 单元格文字彻底清空（value=None）
   - 行高规范 95pt，图片缩放居中嵌入
   - 冻结首行，开启网格线显示 (showGridLines=True)
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re

CENTER_HEADERS = {
    '报价id', '报价ID', '类别编码', '厂家报价文件页码', '顺序号', '报价地区省', '报价地区市',
    '单位', '工程价折扣比例', '折扣下限', '折扣上限', '折扣系列', '是否含运费', '是否含增值税',
    '增值税税率%', '增值税税率', '运费结算方式', '价格填报类型', '项目类型'
}

PRICE_HEADERS = {'市场价', '工程价', '单价', '出厂价', '含税价', '面价', '金额'}

SPEC_HEADERS = {'厂家原始规格参数', '规格参数', '属性项属性值', '详细描述', '技术参数'}

def format_specs_to_multiline(val_str: str) -> str:
    """将分号/逗号分隔的键值参数转换为美观的换行结构化展示"""
    if not val_str:
        return ""
    # 分割键值项
    items = re.split(r'[;\n；]', str(val_str))
    cleaned = []
    for it in items:
        it = it.strip()
        if not it:
            continue
        # 统一冒号间距
        if ':' in it:
            k, v = it.split(':', 1)
            cleaned.append(f"{k.strip()}: {v.strip()}")
        elif '：' in it:
            k, v = it.split('：', 1)
            cleaned.append(f"{k.strip()}: {v.strip()}")
        else:
            cleaned.append(it)
    return "\n".join(cleaned)

def apply_enterprise_style(ws):
    """为 openpyxl 的 Worksheet 应用完整的商业级样式"""
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = 'A2'  # 冻结首行

    header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    header_font = Font(name='宋体', size=11, bold=True, color='FFFFFF')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    ws.row_dimensions[1].height = 36

    headers = [ws.cell(1, c).value or '' for c in range(1, ws.max_column + 1)]

    # 1. 设置表头样式
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border

    # 2. 正文样式
    body_font = Font(name='宋体', size=10)
    alt_fill = PatternFill(start_color='F9FAFC', end_color='F9FAFC', fill_type='solid')
    white_fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')

    center_align = Alignment(horizontal='center', vertical='center')
    left_align_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)
    right_align = Alignment(horizontal='right', vertical='center')

    for r in range(2, ws.max_row + 1):
        row_fill = alt_fill if (r % 2 == 0) else white_fill
        max_lines = 1

        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=r, column=c_idx)
            cell.font = body_font
            cell.fill = row_fill
            cell.border = thin_border

            # 规格参数多行化
            if h in SPEC_HEADERS and cell.value:
                multiline_val = format_specs_to_multiline(str(cell.value))
                cell.value = multiline_val
                lines = multiline_val.count('\n') + 1
                if lines > max_lines:
                    max_lines = lines

            # 对齐与数字格式
            if h in CENTER_HEADERS:
                cell.alignment = center_align
            elif h in PRICE_HEADERS:
                cell.alignment = right_align
                try:
                    if cell.value is not None and str(cell.value).strip() != '':
                        num_val = float(str(cell.value).replace(',', ''))
                        cell.value = num_val
                        cell.number_format = '#,##0.00'
                except Exception:
                    pass
            elif h == '图片':
                cell.value = None
            else:
                cell.alignment = left_align_wrap

        # 动态计算最合适行高
        # 若本行有内嵌图片，行高保持 95pt；否则按多行文本计算自适应行高
        if ws.row_dimensions[r].height is None or ws.row_dimensions[r].height < 90:
            ws.row_dimensions[r].height = max(32, max_lines * 18 + 10)

    # 3. 列宽自适应设置
    for col in ws.columns:
        col_idx = col[0].column
        col_letter = get_column_letter(col_idx)
        header_val = str(ws.cell(1, col_idx).value or '')

        if header_val == '图片':
            ws.column_dimensions[col_letter].width = 24
        elif '规格参数' in header_val or '厂家原始规格参数' in header_val:
            ws.column_dimensions[col_letter].width = 38
        elif any(k in header_val for k in ['说明', '描述', '运费说明', '税费说明', '报价说明']):
            ws.column_dimensions[col_letter].width = 32
        elif any(k in header_val for k in ['名称规格', '品名', '材料名称']):
            ws.column_dimensions[col_letter].width = 22
        elif any(k in header_val for k in ['品牌', '图片文件夹', '报价id', '报价ID']):
            ws.column_dimensions[col_letter].width = 16
        elif any(k in header_val for k in ['市场价', '工程价', '单价']):
            ws.column_dimensions[col_letter].width = 14
        else:
            ws.column_dimensions[col_letter].width = 12

print("Style enhancer defined.")
