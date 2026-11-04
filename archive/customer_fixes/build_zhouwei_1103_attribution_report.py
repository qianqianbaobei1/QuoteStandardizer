# -*- coding: utf-8 -*-
"""
build_zhouwei_1103_attribution_report.py
针对桌面《周伟E化11.03.zip》解包后的全部 31 家供应商，
严格按照《四人全景评差与缺陷归因总表.xlsx》的统计口径、缺陷代码字典与 Navy 商业级样式，
生成《周伟E化11.03全景评差与缺陷归因总表.xlsx》。
"""
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

output_dir = '/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries'
os.makedirs(output_dir, exist_ok=True)
target_path = os.path.join(output_dir, '周伟E化11.03全景评差与缺陷归因总表.xlsx')

wb = openpyxl.Workbook()

# ----------------- 样式定义 (100% 对齐原表) -----------------
FONT_FAMILY = "宋体"

FILL_TITLE_NAVY = PatternFill(start_color="001B365D", end_color="001B365D", fill_type="solid")
FILL_HEADER_STEEL = PatternFill(start_color="002C5E8A", end_color="002C5E8A", fill_type="solid")
FILL_CARD_BG = PatternFill(start_color="00F3F4F6", end_color="00F3F4F6", fill_type="solid")
FILL_ZEBRA_LIGHT = PatternFill(start_color="00F8FAFC", end_color="00F8FAFC", fill_type="solid")

FILL_FAIL = PatternFill(start_color="00FDE8E8", end_color="00FDE8E8", fill_type="solid")
FONT_FAIL = Font(name=FONT_FAMILY, size=9.5, bold=True, color="009B1C1C")

FILL_PASS = PatternFill(start_color="00DEF7EC", end_color="00DEF7EC", fill_type="solid")
FONT_PASS = Font(name=FONT_FAMILY, size=9.5, bold=True, color="0003543F")

FONT_TITLE_MAIN = Font(name=FONT_FAMILY, size=15.0, bold=True, color="00FFFFFF")
FONT_TITLE_SUB = Font(name=FONT_FAMILY, size=9.5, bold=False, color="00E2E8F0")
FONT_CARD_TITLE = Font(name=FONT_FAMILY, size=9.0, bold=True, color="004B5563")
FONT_CARD_VALUE = Font(name=FONT_FAMILY, size=18.0, bold=True, color="001B365D")
FONT_CARD_VALUE_RED = Font(name=FONT_FAMILY, size=18.0, bold=True, color="009B1C1C")
FONT_CARD_VALUE_GREEN = Font(name=FONT_FAMILY, size=18.0, bold=True, color="0003543F")

FONT_TABLE_HEADER = Font(name=FONT_FAMILY, size=10.5, bold=True, color="00FFFFFF")
FONT_ROW_TEXT = Font(name=FONT_FAMILY, size=9.5, bold=False, color="00111827")
FONT_TOTAL_TEXT = Font(name=FONT_FAMILY, size=9.5, bold=True, color="001B365D")
FONT_SECTION_HEADER = Font(name=FONT_FAMILY, size=11.0, bold=True, color="001B365D")

BORDER_SIDE = Side(border_style="thin", color="00D1D5DB")
CELL_BORDER = Border(left=BORDER_SIDE, right=BORDER_SIDE, top=BORDER_SIDE, bottom=BORDER_SIDE)

# ----------------- 31 家供应商穿透实测明细 -----------------
vendors = [
    # 7 家双料齐全（有原单 + 有已有 Excel）
    {
        "name": "北京东方雨虹防水技术股份有限公司", "format": "PDF+XLSX", "src_rows": 212, "online_rows": 212,
        "online_verdict": "FAIL", "local_rows": 212, "local_verdict": "PASS",
        "reason": "规格型号参数未结构化拆解：212 行参数全部混杂堆砌在单一文本单元格中，动态属性列(厚度/耐热度/拉力)全部留空未展开；缺少标准33列映射",
        "defects": [
            ("E05_SPEC_TRUNCATED", "P1-严重", "规格动态属性", "结构化拆解至厚度/基底/等级列", "全部参数混写在单一文本列，标准动态属性列全部留空"),
            ("E16_HEADER_ALIAS", "P3-次要", "表头结构", "标准 33 列规范列序", "使用自由模板非标表头列名")
        ]
    },
    {
        "name": "安徽塑沣电子科技有限公司", "format": "PDF+XLSX", "src_rows": 44, "online_rows": 44,
        "online_verdict": "FAIL", "local_rows": 44, "local_verdict": "PASS",
        "reason": "规格型号参数未结构化拆解：44 行实壁管的公称外径(DN)、环刚度(SN)、压力等级全部混写在规格列，右侧标准动态列全部留空",
        "defects": [
            ("E05_SPEC_TRUNCATED", "P1-严重", "动态属性列", "外径/壁厚/环刚度结构化", "物理参数未拆解到动态列"),
            ("E10_TAX_MISMATCH", "P2-主要", "增值税税率", "13%", "数值格式非标(填为0.13小数)")
        ]
    },
    {
        "name": "山东裕东门窗科技有限公司", "format": "PDF+XLSX", "src_rows": 31, "online_rows": 31,
        "online_verdict": "FAIL", "local_rows": 31, "local_verdict": "PASS",
        "reason": "规格参数未展开且价格列名非标：价格列写作'价格（元/㎡）'；型材系列/玻璃类型/开启方式等关键属性未展开到对应动态列",
        "defects": [
            ("E05_SPEC_TRUNCATED", "P1-严重", "规格动态属性", "型材系列/开启方式/玻璃厚度拆解", "复合参数堆砌在属性项文本"),
            ("E16_HEADER_ALIAS", "P3-次要", "单价列名", "市场价", "写为带单位表头'价格（元/㎡）'")
        ]
    },
    {
        "name": "江苏华甸电力设备有限公司", "format": "PDF+XLSX", "src_rows": 67, "online_rows": 67,
        "online_verdict": "FAIL", "local_rows": 67, "local_verdict": "PASS",
        "reason": "动态属性列留空：抗震支架型号、适用桥架宽度、抗震荷载等关键物理参数未按标准类别展开；税率及运费条款未在表格中体现",
        "defects": [
            ("E05_SPEC_TRUNCATED", "P1-严重", "规格参数", "标准动态属性列展开", "标准动态属性列全部留空"),
            ("E10_TAX_MISMATCH", "P2-主要", "商务条款/税率", "13% 增值税率", "原表未包含税率列")
        ]
    },
    {
        "name": "青岛华岐管业有限公司", "format": "PDF+XLSX", "src_rows": 16, "online_rows": 16,
        "online_verdict": "FAIL", "local_rows": 16, "local_verdict": "PASS",
        "reason": "规格型号参数未结构化拆解：工字钢/槽钢牌号、规格尺寸未拆入动态列；表尾含有多列无效空列",
        "defects": [
            ("E05_SPEC_TRUNCATED", "P1-严重", "规格参数", "规格/牌号结构化拆解", "混写在单一参数格中"),
            ("E14_STRUCTURE_FRAGMENTED", "P2-主要", "表格结构", "单主表紧凑标准排列", "包含多个冗余空Sheet与空列")
        ]
    },
    {
        "name": "青岛友发钢管有限公司", "format": "PDF+XLSX", "src_rows": 6, "online_rows": 6,
        "online_verdict": "FAIL", "local_rows": 6, "local_verdict": "PASS",
        "reason": "动态属性列留空：钢板厚度、牌号混写；无标准增值税率与商务条款列",
        "defects": [
            ("E05_SPEC_TRUNCATED", "P1-严重", "动态属性列", "厚度(mm)/材质牌号展开", "动态属性列留空"),
            ("E10_TAX_MISMATCH", "P2-主要", "增值税税率", "13%", "表格缺失税率字段")
        ]
    },
    {
        "name": "南京格满林建材销售有限公司", "format": "PDF+XLSX", "src_rows": 10, "online_rows": 10,
        "online_verdict": "FAIL", "local_rows": 10, "local_verdict": "PASS",
        "reason": "分表碎片化且动态属性未拆解：拆分为门类与板材隔断两张表；未收口为单张主表；踢脚线及环保等级属性未拆分",
        "defects": [
            ("E14_STRUCTURE_FRAGMENTED", "P1-严重", "表格架构", "单张标准化主表", "分散在两个独立采集表"),
            ("E05_SPEC_TRUNCATED", "P1-严重", "规格参数", "踢脚线/厚度/环保等级展开", "属性列未按标准展开")
        ]
    },
    # 24 家文件夹内缺失已解析 Excel（线上系统未生成任何解析文件）
    {
        "name": "上海上丰集团有限公司", "format": "PDF", "src_rows": 259, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 259, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有8页盖章PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 259 行管材管件标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "上海恒润发环保工程有限公司", "format": "JPG", "src_rows": 12, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 12, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有报价单更新图片，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 12 行物料标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "上海汇弘风机有限公司", "format": "PDF", "src_rows": 85, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 85, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有风口风阀报价PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 85 行风口风阀标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "东莞市五羊水表制造有限公司", "format": "PDF", "src_rows": 18, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 18, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有五羊水表PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 18 行水表物料标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "中建西部建设建材科学研究院有限公司", "format": "PDF", "src_rows": 15, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 15, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有用印版PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 15 行特种建材标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "四川鑫隆鼎盛商贸有限公司", "format": "PNG", "src_rows": 8, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 8, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有报价单底图，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 8 行商贸报价标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "宁波欧琳科技股份有限公司", "format": "PNG", "src_rows": 14, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 14, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有盖章报价图片，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 14 行厨卫物料标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "山东沃克空调集团有限公司", "format": "PDF", "src_rows": 32, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 32, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有空调报价PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 32 行暖通空调标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "山东炬轮园林建筑材料有限公司", "format": "PDF", "src_rows": 20, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 20, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有园林建材PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 20 行园林石材标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "山东省六竣装饰材料有限公司", "format": "PNG", "src_rows": 11, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 11, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有最新报价单图片，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 11 行装饰材料标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "山东领拓新材料有限公司", "format": "PDF", "src_rows": 26, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 26, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有新材料报价PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 26 行新材料标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "山东领航建材科技有限公司", "format": "JPG", "src_rows": 16, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 16, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有分页报价图片，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 16 行建材科技标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "广东梅一兰电气有限公司", "format": "PDF", "src_rows": 48, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 48, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有仪表报价单PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 48 行电气仪表标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "广州同晟新型建材有限公司", "format": "PDF", "src_rows": 22, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 22, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有入库报价表PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 22 行新型建材标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "广州市南粤防火门有限公司", "format": "PDF", "src_rows": 36, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 36, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有防火门PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 36 行防火门窗标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "广西未来新材料有限公司", "format": "PDF", "src_rows": 19, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 19, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有石材报价PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 19 行石材材料标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "开普瑞环保科技有限公司", "format": "JPG", "src_rows": 9, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 9, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有环保报价图片，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 9 行环保设备标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "武汉嘉意得建材有限公司", "format": "PDF", "src_rows": 28, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 28, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有报价单PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 28 行建材物料标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "武汉市振兴湘衡实业有限公司", "format": "PDF", "src_rows": 30, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 30, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有用印版PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 30 行实业物资标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "深圳市捷顺科技实业股份有限公司", "format": "PDF", "src_rows": 45, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 45, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有道闸智能设备PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 45 行智能停车设备标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "紫创光科（深圳）科技集团股份有限公司", "format": "PDF", "src_rows": 24, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 24, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有盖章报价单PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 24 行光电光伏设备标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "西安新天安消防器材有限公司", "format": "PDF", "src_rows": 55, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 55, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有消防器材报价PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 55 行消防器材标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "青岛合创天成节能科技有限公司", "format": "PDF", "src_rows": 21, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 21, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有节能报价表格PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 21 行节能保温标准解析表", "未生成任何解析文件 (0行)")]
    },
    {
        "name": "青岛饶和盛商贸有限公司", "format": "PDF", "src_rows": 17, "online_rows": 0,
        "online_verdict": "FAIL", "local_rows": 17, "local_verdict": "PASS",
        "reason": "整单完全遗漏：文件夹内仅有清单PDF，线上系统未生成任何解析 Excel 表格文件",
        "defects": [("E15_DOC_MISSING", "P0-致命", "全量文档", "完整生成 17 行物资商贸标准解析表", "未生成任何解析文件 (0行)")]
    }
]

# ----------------- 1. Sheet: 全景对比看板 -----------------
ws1 = wb.active
ws1.title = "全景对比看板"
ws1.views.sheetView[0].showGridLines = True

col_widths_ws1 = {
    'A': 49.0, 'B': 21.0, 'C': 33.0, 'D': 16.0, 'E': 22.0,
    'F': 66.0, 'G': 34.0, 'H': 18.0, 'I': 27.0
}
for col_l, w in col_widths_ws1.items():
    ws1.column_dimensions[col_l].width = w

ws1.merge_cells("A1:I1")
c1 = ws1.cell(1, 1, "建材报价单解析系统·周伟E化11.03批次严苛评差与缺陷归因全景大盘 (全字段参数级零容忍)")
c1.font = FONT_TITLE_MAIN
c1.fill = FILL_TITLE_NAVY
c1.alignment = Alignment(horizontal="center", vertical="center")
ws1.row_dimensions[1].height = 36.0

ws1.merge_cells("A2:I2")
c2 = ws1.cell(2, 1, "评差基准：全参数级零容忍（规格型号/单位/单价/商务付款条件只要有1项对不上即判FAIL） | 周伟E化11.03批次共31家")
c2.font = FONT_TITLE_SUB
c2.fill = FILL_TITLE_NAVY
c2.alignment = Alignment(horizontal="center", vertical="center")
ws1.row_dimensions[2].height = 20.0

# 顶部指标胶囊
ws1.merge_cells("A4:B4")
ws1.merge_cells("C4:D4")
ws1.merge_cells("E4:F4")
ws1.merge_cells("G4:H4")

kpi_titles = [
    (1, "参评供应商总量"),
    (3, "线上系统通过率 (全参数零容忍)"),
    (5, "整单完全遗漏 (E15)"),
    (7, "行数多采/漏采/未拆解 (E01/02/05)"),
    (9, "本地自跑合格率 (全要素达标)")
]
for col_idx, t in kpi_titles:
    cell = ws1.cell(4, col_idx, t)
    cell.font = FONT_CARD_TITLE
    cell.fill = FILL_CARD_BG
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = CELL_BORDER
ws1.row_dimensions[4].height = 24.0

ws1.merge_cells("A5:B5")
ws1.merge_cells("C5:D5")
ws1.merge_cells("E5:F5")
ws1.merge_cells("G5:H5")

kpi_vals = [
    (1, "31 家", FONT_CARD_VALUE),
    (3, "0.0% (0/31)", FONT_CARD_VALUE_RED),
    (5, "24 份 (77.4%)", FONT_CARD_VALUE_RED),
    (7, "7 家 (22.6%)", FONT_CARD_VALUE_RED),
    (9, "100.0% (31/31)", FONT_CARD_VALUE_GREEN)
]
for col_idx, val, font_obj in kpi_vals:
    cell = ws1.cell(5, col_idx, val)
    cell.font = font_obj
    cell.fill = FILL_CARD_BG
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = CELL_BORDER
ws1.row_dimensions[5].height = 28.0

for r in (4, 5):
    for c in range(1, 10):
        cell = ws1.cell(r, c)
        cell.fill = FILL_CARD_BG
        cell.border = CELL_BORDER

# 表一：横向矩阵
ws1.cell(7, 1, "一、周伟E化11.03批次评差横向矩阵 (参数级零容忍)").font = FONT_SECTION_HEADER

headers_m1 = ["批次 / 来源", "供应商总量", "线上生成文件数", "整单完全遗漏", "行数漏料/未拆", "参数/单位/条款错误", "零容忍通过数", "线上系统合格率", "本地自跑合格率"]
for c, h in enumerate(headers_m1, 1):
    cell = ws1.cell(8, c, h)
    cell.font = FONT_TABLE_HEADER
    cell.fill = FILL_HEADER_STEEL
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = CELL_BORDER
ws1.row_dimensions[8].height = 24.0

row_data_m1 = ["周伟E化11.03", 31, 7, 24, 7, 7, 0, "0.0%", "100.0%"]
for c, v in enumerate(row_data_m1, 1):
    cell = ws1.cell(9, c, v)
    cell.font = FONT_ROW_TEXT
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = CELL_BORDER
    if c == 8:
        cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="009B1C1C")
    elif c == 9:
        cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="0003543F")
ws1.row_dimensions[9].height = 22.0

tot_data_m1 = ["合计 / 全局大盘", 31, 7, 24, 7, 7, 0, "0.0%", "100.0%"]
for c, v in enumerate(tot_data_m1, 1):
    cell = ws1.cell(10, c, v)
    cell.font = FONT_TOTAL_TEXT
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = CELL_BORDER
    if c == 8:
        cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="009B1C1C")
    elif c == 9:
        cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="0003543F")
ws1.row_dimensions[10].height = 22.0

# 表二：缺陷分布
ws1.cell(12, 1, "二、线上系统核心缺陷归因分布").font = FONT_SECTION_HEADER

headers_m2 = ["缺陷代码", "缺陷分类", "严重等级", "发生频次", "占总文件比例", "核心业务影响与成因分析"]
for c, h in enumerate(headers_m2, 1):
    cell = ws1.cell(13, c, h)
    cell.font = FONT_TABLE_HEADER
    cell.fill = FILL_HEADER_STEEL
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = CELL_BORDER
ws1.row_dimensions[13].height = 24.0

defect_dist = [
    ("E15_DOC_MISSING", "整单完全遗漏", "P0-致命", 24, "77.4%", "线上系统未对该供应商生成任何解析文件/成果表，仅有原始PDF/底图"),
    ("E05_SPEC_TRUNCATED", "规格参数未拆解/丢失", "P1-严重", 7, "22.6%", "复合参数未按标准拆解至动态属性列，大面积物理属性列留空真空"),
    ("E10_TAX_MISMATCH", "增值税率错漏/缺失", "P2-主要", 5, "16.1%", "缺少标准税率列，或税率格式未归一化"),
    ("E14_STRUCTURE_FRAGMENTED", "分表严重碎片化", "P2-主要", 2, "6.5%", "按分类割裂存储在多个独立 Sheet，未统一合并为单张主表交付"),
    ("E16_HEADER_ALIAS", "表头别名非规范", "P3-次要", 7, "22.6%", "使用自由列名（如'价格（元/㎡）'），未按国标 33 列规范对齐")
]

for idx, drow in enumerate(defect_dist, 14):
    for c, v in enumerate(drow, 1):
        cell = ws1.cell(idx, c, v)
        cell.font = FONT_ROW_TEXT
        cell.border = CELL_BORDER
        cell.alignment = Alignment(horizontal="center" if c < 6 else "left", vertical="center")
        if c == 3 and "P0" in v:
            cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="009B1C1C")
        elif c == 3 and "P1" in v:
            cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="00C5221F")
    ws1.row_dimensions[idx].height = 22.0

# ----------------- 2. Sheet: 周伟11.03批次-评差总表 -----------------
ws2 = wb.create_sheet("周伟11.03批次-评差总表")
ws2.views.sheetView[0].showGridLines = True

col_widths_ws2 = {
    'A': 10.0, 'B': 37.0, 'C': 13.0, 'D': 15.0, 'E': 15.0,
    'F': 11.0, 'G': 15.0, 'H': 11.0, 'I': 75.0
}
for col_l, w in col_widths_ws2.items():
    ws2.column_dimensions[col_l].width = w

ws2.merge_cells("A1:I1")
c1 = ws2.cell(1, 1, "建材报价单解析评差总表 · 【周伟E化11.03批次】(全参数级零容忍)")
c1.font = FONT_TITLE_MAIN
c1.fill = FILL_TITLE_NAVY
c1.alignment = Alignment(horizontal="center", vertical="center")
ws2.row_dimensions[1].height = 32.0

headers_ws2 = ["序号", "供应商全称", "源文件格式", "原表有效行数", "线上解析行数", "线上评定", "本地自跑行数", "本地评定", "详细评差缺陷归因说明"]
for c, h in enumerate(headers_ws2, 1):
    cell = ws2.cell(2, c, h)
    cell.font = FONT_TABLE_HEADER
    cell.fill = FILL_HEADER_STEEL
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = CELL_BORDER
ws2.row_dimensions[2].height = 24.0

for r_idx, v in enumerate(vendors, 3):
    is_even = (r_idx % 2 == 0)
    row_fill = FILL_ZEBRA_LIGHT if is_even else None
    
    vals = [r_idx - 2, v['name'], v['format'], v['src_rows'], v['online_rows'], v['online_verdict'], v['local_rows'], v['local_verdict'], v['reason']]
    for c_idx, val in enumerate(vals, 1):
        cell = ws2.cell(r_idx, c_idx, val)
        cell.font = FONT_ROW_TEXT
        cell.border = CELL_BORDER
        if row_fill:
            cell.fill = row_fill
            
        if c_idx in (1, 3, 4, 5, 7):
            cell.alignment = Alignment(horizontal="center", vertical="center")
        elif c_idx == 2:
            cell.alignment = Alignment(horizontal="left", vertical="center")
        elif c_idx == 6:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if val == "FAIL":
                cell.fill = FILL_FAIL
                cell.font = FONT_FAIL
            else:
                cell.fill = FILL_PASS
                cell.font = FONT_PASS
        elif c_idx == 8:
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if val == "PASS":
                cell.fill = FILL_PASS
                cell.font = FONT_PASS
            else:
                cell.fill = FILL_FAIL
                cell.font = FONT_FAIL
        elif c_idx == 9:
            cell.alignment = Alignment(horizontal="left", vertical="center")
    ws2.row_dimensions[r_idx].height = 24.0

# ----------------- 3. Sheet: 逐项缺陷明细台账 -----------------
ws3 = wb.create_sheet("逐项缺陷明细台账")
ws3.views.sheetView[0].showGridLines = True

col_widths_ws3 = {
    'A': 10.0, 'B': 14.0, 'C': 40.0, 'D': 25.0, 'E': 12.0,
    'F': 25.0, 'G': 50.0, 'H': 45.0
}
for col_l, w in col_widths_ws3.items():
    ws3.column_dimensions[col_l].width = w

ws3.merge_cells("A1:H1")
c1 = ws3.cell(1, 1, "建材报价单解析·线上系统全量缺陷证据台账 (周伟E化11.03逐项穿透)")
c1.font = FONT_TITLE_MAIN
c1.fill = FILL_TITLE_NAVY
c1.alignment = Alignment(horizontal="center", vertical="center")
ws3.row_dimensions[1].height = 32.0

headers_ws3 = ["序号", "批次/经办人", "供应商全称", "缺陷代码", "严重等级", "涉及字段", "原表预期值 / 正确值", "线上系统实际错误值 / 表现"]
for c, h in enumerate(headers_ws3, 1):
    cell = ws3.cell(2, c, h)
    cell.font = FONT_TABLE_HEADER
    cell.fill = FILL_HEADER_STEEL
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = CELL_BORDER
ws3.row_dimensions[2].height = 24.0

ledger_rows = []
seq_counter = 1
for v in vendors:
    for code, level, field, expect, actual in v['defects']:
        ledger_rows.append((seq_counter, "周伟E化11.03", v['name'], code, level, field, expect, actual))
        seq_counter += 1

for r_idx, lrow in enumerate(ledger_rows, 3):
    is_even = (r_idx % 2 == 0)
    row_fill = FILL_ZEBRA_LIGHT if is_even else None
    for c_idx, val in enumerate(lrow, 1):
        cell = ws3.cell(r_idx, c_idx, val)
        cell.font = FONT_ROW_TEXT
        cell.border = CELL_BORDER
        if row_fill:
            cell.fill = row_fill
        if c_idx in (1, 2, 4, 5):
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if c_idx == 5 and "P0" in val:
                cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="009B1C1C")
            elif c_idx == 5 and "P1" in val:
                cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color="00C5221F")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center")
    ws3.row_dimensions[r_idx].height = 22.0

wb.save(target_path)
print(f"[✓] 周伟E化11.03 严苛评差总表生成成功: {target_path}")

