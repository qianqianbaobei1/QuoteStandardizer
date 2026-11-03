# -*- coding: utf-8 -*-
"""
style_enhancer.py - 投行级现代商业美学重构引擎 (Enterprise Aesthetic Architecture)
核心升级：
1. 字体升级：采用现代专业商用排版字体（微软雅黑 / Arial 组合），告别传统生硬宋体
2. 配色系统（投行高级感色盘）：
   - 标准业务表头：深海蓝 (#1E3A8A)，纯白加粗，居中对齐，底部加深色分割线 (#0F172A)
   - 动态属性表头：尊贵科技蓝 (#2563EB)，色彩分区明显，突出物料动态规格
   - 背景交替斑马纹：纯白 (#FFFFFF) 与 精致极浅灰蓝 (#F8FAFC) 柔和交替，护眼且层次分明
   - 边框系统：极精细柔和灰 (#E2E8F0)，杜绝生硬黑框
3. 动态规格参数格式化：
   - 智能解析分号或键值串，转化为清晰利落的换行卡片格式（如 "品种: 镀锌槽钢\n牌号: Q235B\n规格: 10#"）
   - 自动开启 wrap_text=True，靠左垂直居中排版，行高舒展（100pt）
4. 数据列对齐与数字格式：
   - 金额列（市场价、工程价）：右对齐，专业加粗，严格统一为人民币会计币种格式 ¥#,##0.00
   - 标识列（单位、序号、页码、税率、状态等）：严格居中对齐
   - 文本列（名称、品牌、说明）：左对齐并垂直居中
5. 图片列纯净展示：
   - 单元格清除底层文件名文字，行高 100pt，高清图片自适应缩放无缝嵌入
6. 冻结首行 (A2) + 显示系统级网格线
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re

CENTER_HEADERS = {
    '报价id', '报价ID', '类别编码', '厂家报价文件页码', '顺序号', '报价地区省', '报价地区市',
    '单位', '工程价折扣比例', '折扣下限', '折扣上限', '折扣系列', '是否含运费', '是否含增值税',
    '增值税税率%', '增值税税率', '运费结算方式', '价格填报类型', '项目类型', '品种', '牌号', '规格'
}

PRICE_HEADERS = {'市场价', '工程价', '单价', '出厂价', '含税价', '面价', '金额'}
SPEC_HEADERS = {'厂家原始规格参数', '规格参数', '属性项属性值', '详细描述', '技术参数'}

def format_specs_to_multiline(val_str: str) -> str:
    """将分号/逗号分隔的键值参数转换为美观的换行结构化展示"""
    if not val_str:
        return ""
    items = re.split(r'[;\n；]', str(val_str))
    cleaned = []
    for it in items:
        it = it.strip()
        if not it:
            continue
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
    """为 openpyxl 的 Worksheet 应用完整的商业级现代美学样式"""
    ws.views.sheetView[0].showGridLines = True
    ws.freeze_panes = 'A2'

    # 1. 颜色与字体定义
    header_font = Font(name='微软雅黑', size=10.5, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid') # 投行深海蓝
    attr_header_fill = PatternFill(start_color='2563EB', end_color='2563EB', fill_type='solid') # 靓丽科技蓝
    
    body_font = Font(name='微软雅黑', size=9.5, color='1E293B') # 炭黑主文本
    price_font = Font(name='Arial', size=10, bold=True, color='0F172A') # 专业数字字体
    spec_font = Font(name='微软雅黑', size=9, color='334155') # 规格参数灰蓝

    zebra_even = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
    zebra_odd = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')

    border_color = 'E2E8F0'
    thin_border = Border(
        left=Side(style='thin', color=border_color),
        right=Side(style='thin', color=border_color),
        top=Side(style='thin', color=border_color),
        bottom=Side(style='thin', color=border_color)
    )
    header_border = Border(
        left=Side(style='thin', color='1E3A8A'),
        right=Side(style='thin', color='1E3A8A'),
        top=Side(style='thin', color='1E3A8A'),
        bottom=Side(style='medium', color='0F172A')
    )

    ws.row_dimensions[1].height = 38
    headers = [ws.cell(1, c).value or '' for c in range(1, ws.max_column + 1)]
    
    # 定位动态属性分区（在报价时间之后的业务属性）
    time_col_idx = 999
    for idx, h in enumerate(headers, 1):
        if '报价时间' in str(h):
            time_col_idx = idx
            break

    # 2. 表头美化
    for c_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx)
        is_attr = c_idx > time_col_idx and h not in ['图片文件夹', '图片']
        cell.fill = attr_header_fill if is_attr else header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = header_border

    # 3. 正文数据行渲染
    center_align = Alignment(horizontal='center', vertical='center')
    left_align_wrap = Alignment(horizontal='left', vertical='center', wrap_text=True)
    right_align = Alignment(horizontal='right', vertical='center')

    for r in range(2, ws.max_row + 1):
        base_fill = zebra_odd if (r % 2 == 1) else zebra_even
        
        for c_idx, h in enumerate(headers, 1):
            cell = ws.cell(row=r, column=c_idx)
            cell.border = thin_border
            cell.font = body_font
            cell.fill = base_fill

            # 规格参数多行结构化
            if h in SPEC_HEADERS:
                cell.font = spec_font
                cell.alignment = left_align_wrap
                if cell.value:
                    cell.value = format_specs_to_multiline(str(cell.value))
            elif h in PRICE_HEADERS:
                cell.font = price_font
                cell.alignment = right_align
                try:
                    if cell.value is not None and str(cell.value).strip() != '':
                        num_val = float(str(cell.value).replace('¥', '').replace(',', ''))
                        cell.value = num_val
                        cell.number_format = '¥#,##0.00'
                except Exception:
                    pass
            elif h in CENTER_HEADERS:
                cell.alignment = center_align
            elif h == '图片':
                cell.value = None
            else:
                cell.alignment = left_align_wrap

        ws.row_dimensions[r].height = 100

    # 4. 列宽自适应比例布局
    for col in ws.columns:
        c_idx = col[0].column
        c_letter = get_column_letter(c_idx)
        header = str(ws.cell(1, c_idx).value or '').strip()

        if header == '图片':
            ws.column_dimensions[c_letter].width = 24
        elif '规格参数' in header or '厂家原始规格参数' in header:
            ws.column_dimensions[c_letter].width = 38
        elif any(k in header for k in ['说明', '运费说明', '税费说明', '报价说明']):
            ws.column_dimensions[c_letter].width = 30
        elif any(k in header for k in ['名称规格', '材料名称']):
            ws.column_dimensions[c_letter].width = 22
        elif any(k in header for k in ['品牌', '图片文件夹', '报价id', '报价ID', '厂家报价图片文件名']):
            ws.column_dimensions[c_letter].width = 18
        elif any(k in header for k in ['市场价', '工程价']):
            ws.column_dimensions[c_letter].width = 16
        elif any(k in header for k in ['品种', '牌号', '规格', '材质及表面处理']):
            ws.column_dimensions[c_letter].width = 16
        elif any(k in header for k in ['宽度', '高度', '厚度']):
            ws.column_dimensions[c_letter].width = 14
        elif any(k in header for k in ['单位', '顺序号', '页码', '省', '市', '税率']):
            ws.column_dimensions[c_letter].width = 10
        else:
            ws.column_dimensions[c_letter].width = 14
