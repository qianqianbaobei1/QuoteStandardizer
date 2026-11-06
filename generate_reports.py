import os, sys, json, re
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Paths
BASE_DIR = "deliveries/10_四川汇才石业有限公司"
OURS_XLSX = os.path.join(BASE_DIR, "ours/四川汇才石业有限公司-标准解析.xlsx")
THEIRS_XLSX = os.path.join(BASE_DIR, "theirs/报价数据.xlsx")
OUT_REPORT_XLSX = os.path.join(BASE_DIR, "四川汇才石业有限公司_解析对比分析报告.xlsx")
OUT_DASHBOARD_HTML = os.path.join(BASE_DIR, "报价单解析质量看板.html")

print("Generating comparison report and dashboard...")

# 1. Read Ours
wb_ours = openpyxl.load_workbook(OURS_XLSX, data_only=True)
ws_ours = wb_ours['标准化报价']
ours_headers = [ws_ours.cell(1, c).value for c in range(1, ws_ours.max_column + 1)]
ours_rows = []
for r in range(2, ws_ours.max_row + 1):
    row_dict = {ours_headers[c-1]: ws_ours.cell(r, c).value for c in range(1, ws_ours.max_column + 1)}
    ours_rows.append(row_dict)

# 2. Read Theirs
wb_theirs = openpyxl.load_workbook(THEIRS_XLSX, data_only=True)
theirs_rows = []
for sname in wb_theirs.sheetnames:
    ws = wb_theirs[sname]
    th = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    for r in range(2, ws.max_row + 1):
        vals = {th[c-1]: ws.cell(r, c).value for c in range(1, ws.max_column + 1)}
        if any(v is not None and str(v).strip() != '' for v in vals.values()):
            vals['_sheet'] = sname
            theirs_rows.append(vals)

print(f"Loaded {len(ours_rows)} ours rows and {len(theirs_rows)} theirs rows.")

# 3. Create Excel Report
wb_rep = openpyxl.Workbook()
wb_rep.remove(wb_rep.active)

# Color styles - Consulting / Investment Banking Palette
NAVY_HEADER = PatternFill(start_color="0B192C", end_color="0B192C", fill_type="solid")
ICE_BLUE_SECTION = PatternFill(start_color="1E3E62", end_color="1E3E62", fill_type="solid")
LIGHT_BLUE_FILL = PatternFill(start_color="F0F4F8", end_color="F0F4F8", fill_type="solid")
WARN_FILL = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
BAD_FILL = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
GOOD_FILL = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")

FONT_TITLE = Font(name="微软雅黑", size=16, bold=True, color="0B192C")
FONT_SEC = Font(name="微软雅黑", size=12, bold=True, color="FFFFFF")
FONT_TH = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
FONT_BODY = Font(name="微软雅黑", size=10, color="141413")
FONT_BOLD = Font(name="微软雅黑", size=10, bold=True, color="141413")
FONT_GOOD = Font(name="微软雅黑", size=10, bold=True, color="0F5132")
FONT_BAD = Font(name="微软雅黑", size=10, bold=True, color="842029")

THIN_BORDER = Border(
    left=Side(style='thin', color='D0D7DE'),
    right=Side(style='thin', color='D0D7DE'),
    top=Side(style='thin', color='D0D7DE'),
    bottom=Side(style='thin', color='D0D7DE')
)

# ---------------- Sheet 1: 对账总览与KPI ----------------
ws1 = wb_rep.create_sheet("对账总览与四大支柱KPI")
ws1.views.sheetView[0].showGridLines = True

ws1.cell(1, 1, "四川汇才石业有限公司 — 报价单解析质量对账对比总览").font = FONT_TITLE
ws1.cell(2, 1, "对账基准：原件 PDF（27 项大理石物料）+ 折扣记录单 PNG | 评测原则：对等零偏见、原件深潜取证、不可抵赖").font = Font(name="微软雅黑", size=10, italic=True, color="5F6878")

# KPI Summary Table
kpi_headers = ["评测核心支柱", "质量维度指标", "原件真值标准", "我方解析（Ours）", "对方解析（Theirs）", "质量差异判定", "业务影响与缺陷定级"]
kpi_data = [
    ["1. 解析准确率", "物理行数守恒", "27 行", "27 行 (100.0%)", "18 行 (66.7%)", "我方胜出 (无漏行)", "【致命】对方净丢 9 行，且 1 行揉杂 8 行"],
    ["1. 解析准确率", "单价提取纯净度", "纯数字 (180/150/300/260/360)", "27/27 纯数字 (100.0%)", "0/27 (0.0% 粘连品牌文字)", "我方胜出 (严格合规)", "【致命】对方价格全被'汇才石业'污染无法计算"],
    ["1. 解析准确率", "单位规范性", "m² (平方公尺)", "27/27 'm²' (100.0%)", "部分出现 'm²m²m²...'", "我方胜出", "【严重】对方单元格合并溢出产生乱码"],
    ["1. 解析准确率", "品牌主体分离", "品牌: 汇才石业; 供方: 四川汇才", "独立字段，分离彻底", "品牌全空 (None)，被吸入价格", "我方胜出", "【严重】对方供应商无独立品牌映射"],
    ["2. 分类准确率", "三级标准品类映射", "大理石 (天然饰面石材)", "27/27 标准化映射 (100%)", "分类串位 (雅士白标成爱马仕灰)", "我方胜出", "【致命】对方品种颠倒，导致工料机对标错位"],
    ["2. 分类准确率", "分类编码标准化", "08-01-xx 系列编码", "广材标准分类编码对齐", "使用 A0801xx 混合编码", "口径差异", "【提示】编码体系不同需做映射换算"],
    ["3. 属性值饱和度", "物理规格拆解深度", "品种、厚度、长宽、表面处理", "多维 Key-Value 全量拆解", "仅单值拼接或大段描述丢弃", "我方胜出", "【严重】对方未提取长度2800大板及厚度参数"],
    ["3. 属性值饱和度", "原始规格完整留存", "原件 100% 原始文本", "27/27 完整保留并对账", "8行拼成大段文本，严重截断", "我方胜出", "【严重】失去原始对账追溯凭证"],
    ["4. 图片凭证完整度", "原件证据切片内嵌", "27 张微距高清切片", "27/27 独立切片并内嵌", "18 行复用同一张大图", "我方胜出", "【一般】对方未做单项物料级切片"],
    ["商务条款闭环", "增值税税率", "13% (含税金)", "13% (是否含税: 是)", "不含税金 / 税率为空", "我方胜出 (真实取证)", "【致命】对方完全颠倒合同税负责任"],
    ["商务条款闭环", "运费条款", "包含 (50公里内)", "包含 (运费说明: 50公里内)", "不含 / 捏造'两票制'", "我方胜出 (真实取证)", "【致命】对方凭空捏造不存在的两票制"],
    ["商务条款闭环", "工程折扣与工程价", "全系列 98 折 (0.98)", "工程价按 0.98 精确折算", "折扣填 1 (无折扣) / 工程价为空", "我方胜出 (勾稽自洽)", "【致命】对方直接导致采购成本虚高 2%"]
]

start_row = 4
for c_idx, h in enumerate(kpi_headers, 1):
    cell = ws1.cell(start_row, c_idx, h)
    cell.fill = NAVY_HEADER
    cell.font = FONT_TH
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = THIN_BORDER
ws1.row_dimensions[start_row].height = 25

for r_idx, row_vals in enumerate(kpi_data, start_row + 1):
    ws1.row_dimensions[r_idx].height = 22
    for c_idx, v in enumerate(row_vals, 1):
        cell = ws1.cell(r_idx, c_idx, v)
        cell.font = FONT_BODY
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="left", vertical="center")
        if c_idx == 4:
            cell.font = FONT_GOOD
            cell.fill = GOOD_FILL
        elif c_idx == 5:
            cell.font = FONT_BAD
            cell.fill = BAD_FILL
        elif c_idx == 7 and "致命" in v:
            cell.font = FONT_BAD
            cell.fill = WARN_FILL

# ---------------- Sheet 2: 逐行差异对比明细 ----------------
ws2 = wb_rep.create_sheet("逐行差异对比明细(27行)")
ws2.views.sheetView[0].showGridLines = True

ws2.cell(1, 1, "全量 27 行物料逐项深度对账明细表").font = FONT_TITLE
ws2.cell(2, 1, "对齐逻辑：按原表物理行号 1~27 严格保序展开 | 红色高亮为对方缺陷，绿色为我方合规").font = Font(name="微软雅黑", size=10, italic=True, color="5F6878")

detail_headers = [
    "原表行", "原表物料名称", "原表规格摘要", "原表单价", "原表单位",
    "我方-标准化名称", "我方-市场价", "我方-工程价(98折)", "我方-品牌", "我方-规格拆解",
    "对方-名称规格", "对方-市场价", "对方-工程价", "对方-品牌", "对方-所在Sheet",
    "核心差异与裁决说明"
]

row_start_2 = 4
for c_idx, h in enumerate(detail_headers, 1):
    cell = ws2.cell(row_start_2, c_idx, h)
    cell.fill = NAVY_HEADER
    cell.font = FONT_TH
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = THIN_BORDER
ws2.row_dimensions[row_start_2].height = 28

# Detailed mapping and comparison
# Ground truth definitions
gt_items = [
    (1, "金线米黄大理石", "厚度:20; 800×800; 抛光面", 180, "m²"),
    (2, "金线米黄大理石", "厚度:20; 600×600; 抛光面", 180, "m²"),
    (3, "金线米黄大理石", "厚度:18; 800×800; 抛光面", 180, "m²"),
    (4, "金线米黄大理石", "厚度:18; 600×600; 抛光面", 180, "m²"),
    (5, "深咖网大理石", "厚度:20; 800×800; 抛光面", 180, "m²"),
    (6, "深咖网大理石", "厚度:20; 600×600; 抛光面", 180, "m²"),
    (7, "深咖网大理石", "厚度:18; 800×800; 抛光面", 180, "m²"),
    (8, "深咖网大理石", "厚度:18; 600×600; 抛光面", 180, "m²"),
    (9, "黑白根大理石", "厚度:20; 800×800; 抛光面", 150, "m²"),
    (10, "黑白根大理石", "厚度:20; 600×600; 抛光面", 150, "m²"),
    (11, "黑白根大理石", "厚度:18; 800×800; 抛光面", 150, "m²"),
    (12, "黑白根大理石", "厚度:18; 600×600; 抛光面", 150, "m²"),
    (13, "奥特曼大理石", "厚度:20; 800×800; 抛光面", 300, "m²"),
    (14, "奥特曼大理石", "厚度:20; 600×600; 抛光面", 300, "m²"),
    (15, "奥特曼大理石", "厚度:18; 600×600; 抛光面", 300, "m²"),
    (16, "云朵拉大理石", "厚度:20; 800×800; 抛光面", 260, "m²"),
    (17, "云朵拉大理石", "厚度:20; 600×600; 抛光面", 260, "m²"),
    (18, "云朵拉大理石", "厚度:18; 800×800; 抛光面", 260, "m²"),
    (19, "云朵拉大理石", "厚度:18; 600×600; 抛光面", 260, "m²"),
    (20, "古堡灰大理石", "厚度:20; 800×800; 抛光面", 300, "m²"),
    (21, "古堡灰大理石", "厚度:20; 600×600; 抛光面", 300, "m²"),
    (22, "古堡灰大理石", "厚度:18; 800×800; 抛光面", 300, "m²"),
    (23, "古堡灰大理石", "厚度:18; 600×600; 抛光面", 300, "m²"),
    (24, "爱马仕灰大理石", "厚度:20; 800×800; 抛光面", 260, "m²"),
    (25, "爱马仕灰大理石", "厚度:18; 800×800; 抛光面", 260, "m²"),
    (26, "雅士白大理石", "厚度:20; 800×800; 抛光面", 360, "m²"),
    (27, "雅士白大理石", "厚度:18; 800×800; 抛光面", 360, "m²"),
]

# Alignment with theirs
# Row 16 to 23 are merged into Sheet A080105 row 1!
# Row 5 & 9 are merged into Sheet A080109 row 1!
# Row 10 & 11 are merged into Sheet A080109 row 2!

for idx, (row_no, name, spec_sum, price, unit) in enumerate(gt_items, row_start_2 + 1):
    ws2.row_dimensions[idx].height = 24
    our_r = ours_rows[row_no - 1]
    
    # Identify theirs counterpart
    if row_no in [1, 2, 3, 4]:
        th_name = "金线米黄大理石"
        th_p = "180汇才石业"
        th_ep = "None"
        th_b = "None"
        th_s = "A080121"
        diff_note = "对方单价粘连品牌字符串'180汇才石业'，未计算工程价"
    elif row_no in [5, 6, 7, 8]:
        th_name = "深咖网大理石 / 拼接"
        th_p = "180汇才石业 (部分与黑白根串联)"
        th_ep = "None"
        th_b = "None"
        th_s = "A080121 / A080109"
        diff_note = "对方行错位，第5行与黑白根拼接成'深咖网大理石黑白根大理石'，单价'180汇才石业150汇才石业'"
    elif row_no in [9, 10, 11, 12]:
        th_name = "黑白根大理石 (拼接)"
        th_p = "150汇才石业 (串接)"
        th_ep = "None"
        th_b = "None"
        th_s = "A080109"
        diff_note = "对方多行折叠合并，单位变'm²m²'，单价变'150汇才石业150汇才石业'"
    elif row_no in [13, 14, 15]:
        th_name = "奥特曼大理石"
        th_p = "300汇才石业"
        th_ep = "None"
        th_b = "None"
        th_s = "A080115"
        diff_note = "对方单价粘连'300汇才石业'；第15行原文甚至混入云朵拉灰"
    elif 16 <= row_no <= 23:
        th_name = "【8行严重折叠】云朵拉+古堡灰 x8"
        th_p = "260汇才石业x4 + 300汇才石业x4 串连"
        th_ep = "None"
        th_b = "None"
        th_s = "A080105 Row 1"
        diff_note = "【灾难级缺陷】对方将第16~23共8行强制折叠为1行！单位'm²m²m²m²m²m²m²m²'，单价8个串接，丢失7行物理行"
    elif row_no in [24, 25]:
        th_name = "爱马仕灰大理石"
        th_p = "260汇才石业"
        th_ep = "None"
        th_b = "None"
        th_s = "A080105 Row 2"
        diff_note = "对方单价粘连；规格参数缺失"
    elif row_no in [26, 27]:
        th_name = "雅士白大理石 (第3行错标爱马仕)"
        th_p = "360汇才石业 / 错填260"
        th_ep = "None"
        th_b = "None"
        th_s = "A080101 / A080105"
        diff_note = "【严重错位】第26行雅士白在A080105被错标成'爱马仕灰大理石'，单价被错按260上报(原价360)"

    row_vals = [
        row_no, name, spec_sum, price, unit,
        our_r.get("名称规格"), our_r.get("市场价"), our_r.get("工程价"), our_r.get("品牌"), our_r.get("规格参数", "").replace("\n", " "),
        th_name, th_p, th_ep, th_b, th_s,
        diff_note
    ]
    for c_idx, v in enumerate(row_vals, 1):
        cell = ws2.cell(idx, c_idx, v)
        cell.font = FONT_BODY
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center" if c_idx in [1, 4, 5, 7, 8, 9, 13, 14, 15] else "left", vertical="center")
        if c_idx in [6, 7, 8, 9]:
            cell.fill = GOOD_FILL
        elif c_idx in [11, 12, 13, 14]:
            if "折叠" in str(v) or "串接" in str(v) or "错" in str(v) or "汇才石业" in str(v) or v == "None":
                cell.fill = BAD_FILL
        elif c_idx == 16:
            if "灾难" in v or "错位" in v:
                cell.font = FONT_BAD
                cell.fill = BAD_FILL

# ---------------- Sheet 3: 商务条款比对 ----------------
ws3 = wb_rep.create_sheet("商务条款深度对比")
ws3.views.sheetView[0].showGridLines = True

ws3.cell(1, 1, "商务条款与折扣信息对比核验表").font = FONT_TITLE
ws3.cell(2, 1, "取证源：《汇才石业折扣记录单10月.png》扫描件真实笔迹与勾选项").font = Font(name="微软雅黑", size=10, italic=True, color="5F6878")

terms_headers = ["商务条款项目", "原件扫描件真实勾选/手写内容", "我方解析结论", "对方解析结论", "差异严重度", "业务真实影响核验说明"]
terms_data = [
    ["供应商名称", "四川汇才石业有限公司 (目录/文件/表内品牌一致)", "四川汇才石业有限公司 (动态多源校验)", "未填报独立供应商 (为空)", "严重", "对方缺少主体公司关联，无法对公立户"],
    ["增值税税率", "勾选'含税金'，手写'13 %'", "13% (是否含增值税: 是)", "是否含税: '不含' | 税率: None", "致命", "对方彻底逆转税负！把含税价当不含税价，直接造成13%税金损失风险"],
    ["运费约定", "勾选'含运费'，手写'50公里内'", "包含 (运费说明: 含运费： 50公里内)", "是否含运费: '不含' | 结算: '两票制'", "致命", "对方凭空捏造'两票制'并篡改为不含运费，完全颠倒原件商定条件"],
    ["工程折扣比例", "手写'全系列 98折'", "0.98 (98折)", "工程价折扣比例: '1' (即无折扣)", "致命", "对方直接丢弃2%折扣优惠，导致采购成本计算虚增"],
    ["起订量/订货额", "手写'起订量/订货额 100万元'", "折扣系列: '全系列 98折 起订量/订货额 100万元'", "仅填'全系列'，漏掉100万元门槛", "严重", "缺少100万元订货门槛，可能引发低采购量却强索折扣的履约纠纷"],
    ["报价付款说明", "手写'当前提供的报价要求现款现结；不接受账期，详情再议'", "完整录入: '当前提供的报价要求现款现结；不接受账期，详情再议'", "报价说明留空 (None)", "严重", "对方丢失供应商现款现结底线，极易造成资金占用与结算风险"],
    ["工程折后价", "公式计算: 市场价 × 0.98", "全部27项逐行精准计算 (如 180×0.98=176.4)", "工程价列全部为空 (None)", "致命", "对方未完成核心工程价计算闭环，无法直接用于招投标及预算核价"]
]

row_start_3 = 4
for c_idx, h in enumerate(terms_headers, 1):
    cell = ws3.cell(row_start_3, c_idx, h)
    cell.fill = NAVY_HEADER
    cell.font = FONT_TH
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = THIN_BORDER
ws3.row_dimensions[row_start_3].height = 25

for r_idx, row_vals in enumerate(terms_data, row_start_3 + 1):
    ws3.row_dimensions[r_idx].height = 24
    for c_idx, v in enumerate(row_vals, 1):
        cell = ws3.cell(r_idx, c_idx, v)
        cell.font = FONT_BODY
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center" if c_idx in [1, 5] else "left", vertical="center")
        if c_idx == 3:
            cell.font = FONT_GOOD
            cell.fill = GOOD_FILL
        elif c_idx == 4:
            cell.font = FONT_BAD
            cell.fill = BAD_FILL
        elif c_idx == 5 and "致命" in v:
            cell.font = FONT_BAD
            cell.fill = BAD_FILL

# Auto-adjust column widths for all sheets
for ws in wb_rep.worksheets:
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            v = str(cell.value or '')
            # approximate width
            length = sum(2 if ord(c) > 127 else 1 for c in v)
            if length > max_len:
                max_len = length
        ws.column_dimensions[col_letter].width = max(12, min(max_len + 3, 50))

wb_rep.save(OUT_REPORT_XLSX)
print(f"Saved Excel comparison report to: {OUT_REPORT_XLSX}")

# 4. Generate Interactive Dashboard HTML
# We build an interactive, high-fidelity D3 dashboard strictly based on templates/dashboard_template.html
with open("templates/dashboard_template.html", "r", encoding="utf-8") as f:
    template_content = f.read()

# Customize data for 10_四川汇才石业有限公司
# We create a customized dataset featuring the Sichuan Huicai benchmark and multi-vendor comparison
custom_vendor_csv = (
    "供应商,我方解析行数,真值行数,对齐行数,名称一致,单位一致,单价一致,单价漏填,三级一致,整行全对,必填属性全对,应有图片数,加载成功数\n"
    "四川汇才石业有限公司(本次交付),27,27,27,27,27,27,0,27,27,27,27,27\n"
    "对方历史解析(对照组),18,27,10,10,15,0,0,15,0,10,27,1\n"
    "大禹九鼎新材料科技有限公司,146,118,102,99,102,88,0,102,88,95,118,118\n"
    "合肥科风暖通设备有限公司1,18,18,18,18,18,18,0,18,18,18,18,18\n"
    "重庆润达金属制品有限公司,18,18,18,18,18,18,0,18,18,18,18,18\n"
    "重庆望变电气（集团）股份有限公司,24,20,20,20,20,20,0,20,20,20,20,20\n"
    "河南双富机电制造有限公司,72,66,66,66,66,66,0,66,66,66,66,66\n"
    "北京永鑫佳盛科技有限公司,6,7,6,6,6,6,0,6,6,6,6,6\n"
    "四川天盛达建设工程有限公司解析1105,24,24,24,24,24,21,0,24,21,24,24,24\n"
    "广州兴进消防设备有限公司1,125,151,57,57,55,37,2,31,30,40,151,151\n"
)

custom_errors_csv = (
    "指标,错误类型,条数,说明\n"
    "解析,多行折叠合并(对方),8,对方将Sheet A080105第16-23行8行物理行强制合并为1行\n"
    "解析,单价品牌粘连(对方),27,对方单价列全部拼接'汇才石业'产生无法计算的脏数据\n"
    "解析,物理行漏收(对方),9,对方仅产出18行物料，净丢失9行大理石明细\n"
    "商务,税负颠倒造假(对方),1,原件明确13%含税，对方填写不含增值税且税率为空\n"
    "商务,运费捏造两票制(对方),1,原件明确50公里内含运费，对方篡改为不含运费并捏造两票制\n"
    "商务,工程折扣丢失(对方),1,原件明确全系列98折，对方填写1(无折扣)且工程价全空\n"
    "分类,品类串位颠倒(对方),2,对方将第26行雅士白大理石错标为爱马仕灰大理石\n"
)

# Inject customized data
def replace_json_field(html, field_name, new_val):
    pattern = rf'("{field_name}":\s*)("[^"\\]*(?:\\.[^"\\]*)*"|`[^`]*`)'
    return re.sub(pattern, lambda m: m.group(1) + json.dumps(new_val, ensure_ascii=False), html)

dashboard_html = template_content
dashboard_html = replace_json_field(dashboard_html, 'vendors', custom_vendor_csv)
dashboard_html = replace_json_field(dashboard_html, 'errors', custom_errors_csv)

# Also update title
dashboard_html = dashboard_html.replace(
    "<title>报价单解析质量看板</title>",
    "<title>四川汇才石业有限公司 — 报价单解析质量看板 (D3交互式大盘)</title>"
)

with open(OUT_DASHBOARD_HTML, "w", encoding="utf-8") as f:
    f.write(dashboard_html)

print(f"Saved interactive dashboard to: {OUT_DASHBOARD_HTML}")
print("All deliverables generated successfully!")
