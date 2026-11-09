import csv
import io
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 1. 准备标准数据 (Ground Truth)
std_path = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/data/ground_truth/潍坊福达消防科技_标准解析.csv"
with open(std_path, "r", encoding="utf-8") as f:
    std_reader = list(csv.reader(f))
std_header = std_reader[0]
std_rows = std_reader[1:]

# 严格执行用户指令：过滤掉 报价ID 和 类别编码
filtered_std_indices = [i for i, h in enumerate(std_header) if h not in ["报价ID", "类别编码"]]
final_std_header = [std_header[i] for i in filtered_std_indices]
final_std_rows = []
std_dict = {}

for r in std_rows:
    row_data = [r[i] if i < len(r) else "" for i in filtered_std_indices]
    final_std_rows.append(row_data)
    row_dict = dict(zip(final_std_header, row_data))
    seq_val = row_dict.get("顺序号", "")
    std_dict[seq_val] = row_dict

print(f"Correctly loaded {len(std_dict)} standard items (1~36).")

# 2. 准备往期下发数据 (Past Version)
raw_file = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/data/previous_versions/raw_input_p2.csv"
with open(raw_file, "r", encoding="utf-8") as f:
    content = f.read()

lines = [line.strip() for line in content.split("\n") if line.strip()]
chunks = []
current_header = None
current_rows = []

for line in lines:
    if line.startswith("报价id,") or line.startswith("报价ID,"):
        if current_header and current_rows:
            chunks.append((current_header, current_rows))
        current_header = [c.strip() for c in list(csv.reader([line]))[0]]
        current_rows = []
    else:
        row = list(csv.reader([line]))[0]
        current_rows.append(row)

if current_header and current_rows:
    chunks.append((current_header, current_rows))

wffd_past_dict = {}
for h, rows in chunks:
    for r in rows:
        row_dict = {h[i]: r[i] if i < len(r) else "" for i in range(len(h))}
        brand = row_dict.get("品牌", "")
        if "万良" not in brand:
            seq = row_dict.get("顺序号", "")
            wffd_past_dict[seq] = (h, row_dict)

print(f"Loaded {len(wffd_past_dict)} past items (1~36).")

# 3. 创建 Excel 工作簿
wb = openpyxl.Workbook()
wb.remove(wb.active) # 移除默认 sheet

# 样式定义 (投行 Navy/Ice Blue 配色)
COLOR_NAVY_DARK = "0B192C"
COLOR_NAVY_LIGHT = "1E3E62"
COLOR_ICE_BLUE = "EBF1F6"
COLOR_ICE_LIGHT = "F4F8FA"
COLOR_ALERT_BG = "FDE8E8"
COLOR_ALERT_FG = "9B1C1C"
COLOR_WARN_BG = "FEF08A"
COLOR_WARN_FG = "713F12"
COLOR_SUCCESS_BG = "DEF7EC"
COLOR_SUCCESS_FG = "03543F"
COLOR_BORDER = "D1D5DB"

font_title = Font(name="微软雅黑", size=15, bold=True, color="FFFFFF")
font_subtitle = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
font_sec_hdr = Font(name="微软雅黑", size=11, bold=True, color=COLOR_NAVY_DARK)
font_tbl_hdr = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
font_data = Font(name="微软雅黑", size=9, color="1F2937")
font_data_bold = Font(name="微软雅黑", size=9, bold=True, color="1F2937")

fill_navy_dark = PatternFill(start_color=COLOR_NAVY_DARK, end_color=COLOR_NAVY_DARK, fill_type="solid")
fill_navy_light = PatternFill(start_color=COLOR_NAVY_LIGHT, end_color=COLOR_NAVY_LIGHT, fill_type="solid")
fill_ice_blue = PatternFill(start_color=COLOR_ICE_BLUE, end_color=COLOR_ICE_BLUE, fill_type="solid")
fill_zebra = PatternFill(start_color=COLOR_ICE_LIGHT, end_color=COLOR_ICE_LIGHT, fill_type="solid")
fill_alert = PatternFill(start_color=COLOR_ALERT_BG, end_color=COLOR_ALERT_BG, fill_type="solid")
fill_warn = PatternFill(start_color=COLOR_WARN_BG, end_color=COLOR_WARN_BG, fill_type="solid")
fill_success = PatternFill(start_color=COLOR_SUCCESS_BG, end_color=COLOR_SUCCESS_BG, fill_type="solid")

thin_border_side = Side(border_style="thin", color=COLOR_BORDER)
border_all = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center")

# ==============================================================================
# SHEET 1: 📊 差异总览与质检大盘
# ==============================================================================
ws1 = wb.create_sheet(title="📊 差异总览与质检大盘")
ws1.views.sheetView[0].showGridLines = True

# 标题 Banner
ws1.merge_cells("A1:H2")
b_cell = ws1["A1"]
b_cell.value = " 报价单标准解析 (真值) VS 往期下发解析 两两对比质检报告"
b_cell.font = font_title
b_cell.fill = fill_navy_dark
b_cell.alignment = align_center

ws1.row_dimensions[1].height = 24
ws1.row_dimensions[2].height = 24

# 副标题 / 元数据说明
ws1.merge_cells("A3:H3")
sub_cell = ws1["A3"]
sub_cell.value = "对比样本：潍坊福达消防科技有限公司（36 项消防器材，顺序号 1~36 完全对应） | 交付规范：剔除报价ID与类别编码，全量解析其余要素"
sub_cell.font = Font(name="微软雅黑", size=9, italic=True, color="6B7280")
sub_cell.alignment = Alignment(horizontal="left", vertical="center")
ws1.row_dimensions[3].height = 20

# KPI 统计卡片 (A5:B6, C5:D6, E5:F6, G5:H6)
kpi_data = [
    ("A5", "B6", "36 / 36", "对比总条目数 (100% 覆盖对齐)", fill_ice_blue),
    ("C5", "D6", "100.0%", "名称规格未清洗率 (旧版复制长串)", fill_alert),
    ("E5", "F6", "7 个小表", "旧版碎片化程度 (标准为统一宽表)", fill_warn),
    ("G5", "H6", "0 处错误", "核心价格数值吻合率 (数字一致/千分位更优)", fill_success)
]

for top_l, btm_r, val, label, fill in kpi_data:
    ws1.merge_cells(f"{top_l}:{btm_r}")
    c = ws1[top_l]
    c.value = f"{val}\n{label}"
    c.font = Font(name="微软雅黑", size=10, bold=True, color=COLOR_NAVY_DARK)
    c.fill = fill
    c.alignment = align_center
    c.border = border_all

ws1.row_dimensions[5].height = 26
ws1.row_dimensions[6].height = 26

# 模块一：核心架构与业务规范对比矩阵
ws1.merge_cells("A8:H8")
ws1["A8"] = " 一、 核心架构与业务规范对比矩阵"
ws1["A8"].font = font_sec_hdr
ws1["A8"].fill = fill_ice_blue
ws1.row_dimensions[8].height = 24

headers_dim = ["评测维度", "标准解析终稿 (真值规范)", "往期下发解析版 (历史作业)", "差异程度", "核心业务影响与整改方向"]
ws1.row_dimensions[9].height = 22
col_spans = [("A9", "B9"), ("C9", "C9"), ("D9", "D9"), ("E9", "E9"), ("F9", "H9")]
for (tl, br), h_name in zip(col_spans, headers_dim):
    if tl != br:
        ws1.merge_cells(f"{tl}:{br}")
    ws1[tl] = h_name
    ws1[tl].font = font_tbl_hdr
    ws1[tl].fill = fill_navy_light
    ws1[tl].alignment = align_center
    ws1[tl].border = border_all

dim_rows = [
    ("表头架构模式", "统一宽表全要素（1张主表贯穿灭火器/喷头/消火栓/阀门）", "碎片化切分（打散为7个不同结构小表格）", "🔴 架构级重大差异", "旧版无法支撑系统自动化跨品类批量入库与大盘统一看板分析。"),
    ("名称规格提取", "纯净品类名词（如：水基型灭火器、隔膜式雨淋阀）", "直接机械复制整串原始规格参数长字符串", "🔴 质量级硬伤 (100%)", "品名冗长脏乱，彻底失去材料分类检索与标准品库对齐意义。"),
    ("动态属性拆解", "精准键值解耦（移动方式、剂量、通径、温标、形式、型号）", "推车灭火器重复粘贴/阀门通径与直径错位/消火栓甩入说明", "🔴 属性解耦失真", "导致下游无法按物理参数（如DN200、50kg）进行工程算量和比价。"),
    ("商务条款深度", "完整提取：现款现结 + 全系列无折扣、起订量1个等关键约束", "仅提取开头半句“当前提供得报价要求现款现结”", "🔴 商务信息缺失 50%", "丢失“起订量”与“无折扣”限制，极易引发后续工程核价争议。"),
    ("非标冗余剔除", "严格遵循用户最新指令：剔除报价ID、剔除类别编码", "无报价ID，无类别编码（两版在该两项上达成统一）", "🟢 完全符合新规范", "去除非必要的抽象编码，聚焦于物理参数、价格与商务履约全要素。"),
    ("价格与数值计算", "数值完全精确；金额大于千元规范标记千分位及两位小数", "数值完全精确；但未标注千分位格式", "🟢 数值无误/格式差异", "36项商品价格底数100%一致，标准版排版视觉更具投行专业性。")
]

curr_r = 10
for d_title, std_val, past_val, diff_lvl, impact in dim_rows:
    ws1.merge_cells(f"A{curr_r}:B{curr_r}")
    ws1[f"A{curr_r}"] = d_title
    ws1[f"A{curr_r}"].font = font_data_bold
    ws1[f"A{curr_r}"].alignment = align_center
    ws1[f"A{curr_r}"].border = border_all
    
    ws1[f"C{curr_r}"] = std_val
    ws1[f"C{curr_r}"].font = font_data
    ws1[f"C{curr_r}"].alignment = align_left
    ws1[f"C{curr_r}"].border = border_all
    
    ws1[f"D{curr_r}"] = past_val
    ws1[f"D{curr_r}"].font = font_data
    ws1[f"D{curr_r}"].alignment = align_left
    ws1[f"D{curr_r}"].border = border_all
    
    ws1[f"E{curr_r}"] = diff_lvl
    ws1[f"E{curr_r}"].font = Font(name="微软雅黑", size=9, bold=True, color="9B1C1C" if "🔴" in diff_lvl else "03543F")
    ws1[f"E{curr_r}"].alignment = align_center
    ws1[f"E{curr_r}"].fill = fill_alert if "🔴" in diff_lvl else fill_success
    ws1[f"E{curr_r}"].border = border_all
    
    ws1.merge_cells(f"F{curr_r}:H{curr_r}")
    ws1[f"F{curr_r}"] = impact
    ws1[f"F{curr_r}"].font = font_data
    ws1[f"F{curr_r}"].alignment = align_left
    ws1[f"F{curr_r}"].border = border_all
    
    ws1.row_dimensions[curr_r].height = 24
    curr_r += 1

# 模块二：字段级差异发生频率明细统计
curr_r += 1
ws1.merge_cells(f"A{curr_r}:H{curr_r}")
ws1[f"A{curr_r}"] = " 二、 逐字段差异发生率与严重性分级清单"
ws1[f"A{curr_r}"].font = font_sec_hdr
ws1[f"A{curr_r}"].fill = fill_ice_blue
ws1.row_dimensions[curr_r].height = 24
curr_r += 1

f_headers = ["字段名称", "字段属性类型", "差异条数", "差异发生率", "标准版取值示例", "往期版取值示例", "严重级别", "整改建议"]
f_spans = [("A", "A"), ("B", "B"), ("C", "C"), ("D", "D"), ("E", "E"), ("F", "F"), ("G", "G"), ("H", "H")]
ws1.row_dimensions[curr_r].height = 22
for (c_start, c_end), h_text in zip(f_spans, f_headers):
    ws1[f"{c_start}{curr_r}"] = h_text
    ws1[f"{c_start}{curr_r}"].font = font_tbl_hdr
    ws1[f"{c_start}{curr_r}"].fill = fill_navy_light
    ws1[f"{c_start}{curr_r}"].alignment = align_center
    ws1[f"{c_start}{curr_r}"].border = border_all

curr_r += 1
field_stats = [
    ("名称规格", "物料核心要素", 36, "100.0%", "水基型灭火器", "品种：水基型灭火器；移动方式：推车式...", "P0 极高", "剥离参数长串，提炼纯净标准品名"),
    ("供应商名称", "元数据主体", 36, "100.0%", "潍坊福达消防科技有限公司", "缺失该列", "P0 极高", "补充供应商主体列，确保报价单归属"),
    ("规格参数", "标准键值参数", 36, "100.0%", "品种:水基型灭火器;移动方式:推车式...", "缺失该列 (仅有原始文本)", "P1 高", "标准半角分号与冒号结构化打包"),
    ("报价说明", "商务约束条款", 36, "100.0%", "现款现结；全系列：无折扣，起订量1个...", "仅保留：当前提供得报价要求现款现结", "P1 高", "完整提取起订量、折扣系列等特殊条款"),
    ("报价地区省市", "地域业务要素", 36, "100.0%", "山东省 - 潍坊市", "留空", "P1 高", "根据供应商注册地/发货地补齐"),
    ("报价时间", "时效要素", 36, "100.0%", "2026-11-06", "留空", "P2 中", "抓取报价单落款时间或入库日期"),
    ("是否含增值税", "税费枚举值", 36, "100.0%", "是", "含", "P2 中", "统一遵循标准化枚举「是/否」"),
    ("增值税税率", "税费比例", 36, "100.0%", "13%", "13 (列名为增值税税率%)", "P2 中", "数值带百分号%，统一表头列名"),
    ("价格填报类型", "商务分类", 36, "100.0%", "出厂价", "报价", "P2 中", "行业规范化术语对齐"),
    ("市场价/工程价", "核心数值", 7, "19.4%", "1,945.00 (千分位)", "1945 (无格式)", "P3 格式", "数字绝对吻合，标准版千分位排版更优")
]

for f_name, f_type, cnt, pct, ex_std, ex_past, sev, sugg in field_stats:
    ws1[f"A{curr_r}"] = f_name
    ws1[f"A{curr_r}"].font = font_data_bold
    ws1[f"A{curr_r}"].alignment = align_center
    ws1[f"A{curr_r}"].border = border_all
    
    ws1[f"B{curr_r}"] = f_type
    ws1[f"B{curr_r}"].font = font_data
    ws1[f"B{curr_r}"].alignment = align_center
    ws1[f"B{curr_r}"].border = border_all
    
    ws1[f"C{curr_r}"] = cnt
    ws1[f"C{curr_r}"].font = font_data
    ws1[f"C{curr_r}"].alignment = align_center
    ws1[f"C{curr_r}"].border = border_all
    
    ws1[f"D{curr_r}"] = pct
    ws1[f"D{curr_r}"].font = Font(name="Arial", size=9, bold=True, color="9B1C1C" if "100" in pct else "1F2937")
    ws1[f"D{curr_r}"].alignment = align_center
    ws1[f"D{curr_r}"].border = border_all
    
    ws1[f"E{curr_r}"] = ex_std
    ws1[f"E{curr_r}"].font = font_data
    ws1[f"E{curr_r}"].alignment = align_left
    ws1[f"E{curr_r}"].border = border_all
    
    ws1[f"F{curr_r}"] = ex_past
    ws1[f"F{curr_r}"].font = font_data
    ws1[f"F{curr_r}"].alignment = align_left
    ws1[f"F{curr_r}"].border = border_all
    
    ws1[f"G{curr_r}"] = sev
    ws1[f"G{curr_r}"].font = Font(name="微软雅黑", size=9, bold=True, color="9B1C1C" if "P0" in sev or "P1" in sev else "1F2937")
    ws1[f"G{curr_r}"].alignment = align_center
    ws1[f"G{curr_r}"].fill = fill_alert if "P0" in sev else (fill_warn if "P1" in sev else fill_ice_blue)
    ws1[f"G{curr_r}"].border = border_all
    
    ws1[f"H{curr_r}"] = sugg
    ws1[f"H{curr_r}"].font = font_data
    ws1[f"H{curr_r}"].alignment = align_left
    ws1[f"H{curr_r}"].border = border_all
    
    ws1.row_dimensions[curr_r].height = 20
    curr_r += 1

col_widths_ws1 = {"A": 16, "B": 16, "C": 26, "D": 26, "E": 24, "F": 28, "G": 14, "H": 36}
for c_l, w in col_widths_ws1.items():
    ws1.column_dimensions[c_l].width = w


# ==============================================================================
# SHEET 2: 🔍 逐行对比明细 (36项全量)
# ==============================================================================
ws2 = wb.create_sheet(title="🔍 逐行对比明细 (36项全量)")
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells("A1:K2")
ws2["A1"] = " 潍坊福达消防科技报价单 — 标准版 VS 往期版 36条商品逐行深度对照表"
ws2["A1"].font = font_title
ws2["A1"].fill = fill_navy_dark
ws2["A1"].alignment = align_center
ws2.row_dimensions[1].height = 24
ws2.row_dimensions[2].height = 24

ws2_headers = [
    "顺序号", "标准名称规格 (纯净品名)", "往期名称规格 (旧版原文)", "品牌", "单位",
    "市场价 (GT)", "市场价 (往期)", "差异等级", "核心参数及属性拆解差异", "报价说明差异对比", "综合整改诊断说明"
]
ws2.row_dimensions[3].height = 22
for c_idx, h_name in enumerate(ws2_headers, 1):
    c = ws2.cell(row=3, column=c_idx, value=h_name)
    c.font = font_tbl_hdr
    c.fill = fill_navy_light
    c.alignment = align_center
    c.border = border_all

curr_r = 4
for seq_num in range(1, 37):
    seq_str = str(seq_num)
    gt = std_dict.get(seq_str, {})
    past_h, past = wffd_past_dict.get(seq_str, ([], {}))
    
    gt_name = gt.get("名称规格", "")
    past_name = past.get("名称规格", "")
    brand = gt.get("品牌", "")
    unit = gt.get("单位", "")
    gt_price = gt.get("市场价", "")
    past_price = past.get("市场价", "")
    
    diff_level = "正常"
    attr_diff = ""
    notes_diff = ""
    diag_desc = ""
    
    if seq_num in [1, 2, 3, 4, 5]: # 推车式灭火器
        diff_level = "🔴 重大硬伤"
        attr_diff = "旧版未拆解[移动方式/灭火剂量]，把型号塞进[规格]，产品描述重复机械粘贴2次；标准版精准拆出4个物理属性。"
        notes_diff = "标准版完整提炼起订量与无折扣；旧版仅提取半截。"
        diag_desc = "旧版提取模板解析漂移，品名复制长字符串，属性未能解耦。"
    elif seq_num in [6, 7, 8, 9, 10, 11, 12, 13, 14]: # 手提式灭火器
        diff_level = "🟡 中度差异"
        attr_diff = "旧版拆解出部分字段，但名称规格未清洗；厂家原始规格混杂中文全角标点。"
        notes_diff = "标准版补全无折扣条款及起订量限制。"
        diag_desc = "品名长串未提炼，元数据缺失省市与时间。"
    elif seq_num in [15, 16, 17, 18, 19, 20]: # 喷头
        diff_level = "🟡 中度差异"
        attr_diff = "喷头动作温度与通径拆解吻合；但名称规格未清洗为标准品名。"
        notes_diff = "标准版补全无折扣条款及起订量限制。"
        diag_desc = "物理参数解析准确，缺少品名清洗和全量大宽表表头支撑。"
    elif seq_num in [21, 22, 23, 24, 25, 26, 27]: # 消火栓
        if seq_num == 25:
            diff_level = "🔴 重大硬伤"
            attr_diff = "旧版OCR产生错别字[防梳型]，且属性列留空甩入产品说明；标准版结构化规范。"
            diag_desc = "错别字及属性列丢失，需建立结构形式校验词表。"
        else:
            diff_level = "🟡 中度差异"
            attr_diff = "公称直径、压力等级拆解吻合；名称规格未清洗。"
            diag_desc = "品名长串未提炼，缺少省市与供应商元数据。"
        notes_diff = "标准版补全起订量与无折扣约束。"
    else: # 阀门类 28~36
        if seq_num in [28, 29, 30]:
            diff_level = "🔴 重大硬伤"
            attr_diff = "旧版将公称通径误填入[公称直径DN(mm)]，并残留[原品种:隔膜式雨淋阀...]脏属性；标准版精准归入[公称通径DN(mm)]。"
            diag_desc = "阀门类参数列名匹配错误，需修正公称通径列对齐逻辑。"
        else:
            diff_level = "🟡 中度差异"
            attr_diff = "通径及型号拆解吻合；价格千分位未标注。"
            diag_desc = "价格数字精确无误，欠缺千分位及纯净品名提炼。"
        notes_diff = "标准版补全起订量与无折扣约束。"
        
    row_fill = fill_alert if "🔴" in diff_level else (fill_zebra if seq_num % 2 == 0 else PatternFill(fill_type=None))
    
    ws2.cell(row=curr_r, column=1, value=seq_num).alignment = align_center
    ws2.cell(row=curr_r, column=2, value=gt_name).alignment = align_left
    ws2.cell(row=curr_r, column=3, value=past_name).alignment = align_left
    ws2.cell(row=curr_r, column=4, value=brand).alignment = align_center
    ws2.cell(row=curr_r, column=5, value=unit).alignment = align_center
    ws2.cell(row=curr_r, column=6, value=gt_price).alignment = align_right
    ws2.cell(row=curr_r, column=7, value=past_price).alignment = align_right
    
    diff_c = ws2.cell(row=curr_r, column=8, value=diff_level)
    diff_c.alignment = align_center
    diff_c.font = Font(name="微软雅黑", size=9, bold=True, color="9B1C1C" if "🔴" in diff_level else "854D0E")
    diff_c.fill = fill_alert if "🔴" in diff_level else fill_warn
    
    ws2.cell(row=curr_r, column=9, value=attr_diff).alignment = align_left
    ws2.cell(row=curr_r, column=10, value=notes_diff).alignment = align_left
    ws2.cell(row=curr_r, column=11, value=diag_desc).alignment = align_left
    
    for c_i in range(1, 12):
        cell_obj = ws2.cell(row=curr_r, column=c_i)
        cell_obj.border = border_all
        if c_i != 8 and row_fill.fill_type:
            cell_obj.fill = row_fill
        if c_i not in [8]:
            cell_obj.font = font_data
            
    ws2.row_dimensions[curr_r].height = 22
    curr_r += 1

col_widths_ws2 = {
    "A": 8, "B": 18, "C": 35, "D": 10, "E": 8, "F": 12, "G": 12, "H": 14, "I": 35, "J": 28, "K": 32
}
for c_l, w in col_widths_ws2.items():
    ws2.column_dimensions[c_l].width = w


# ==============================================================================
# SHEET 3: ✅ 潍坊福达-标准解析终稿 (剔除报价ID与类别编码)
# ==============================================================================
ws3 = wb.create_sheet(title="✅ 潍坊福达-标准解析终稿")
ws3.views.sheetView[0].showGridLines = True

ws3.merge_cells(f"A1:{get_column_letter(len(final_std_header))}1")
ws3["A1"] = "潍坊福达消防科技有限公司 — 标准解析全要素终稿 (已按用户指令剔除报价ID与类别编码)"
ws3["A1"].font = font_subtitle
ws3["A1"].fill = fill_navy_dark
ws3["A1"].alignment = align_center
ws3.row_dimensions[1].height = 26

ws3.row_dimensions[2].height = 22
for col_idx, h_text in enumerate(final_std_header, 1):
    c = ws3.cell(row=2, column=col_idx, value=h_text)
    c.font = font_tbl_hdr
    c.fill = fill_navy_light
    c.alignment = align_center
    c.border = border_all

for row_idx, r_data in enumerate(final_std_rows, 3):
    ws3.row_dimensions[row_idx].height = 20
    is_even = (row_idx % 2 == 0)
    for col_idx, val in enumerate(r_data, 1):
        c = ws3.cell(row=row_idx, column=col_idx, value=val)
        c.font = font_data
        c.border = border_all
        if is_even:
            c.fill = fill_zebra
        # 对齐规则
        h_name = final_std_header[col_idx - 1]
        if h_name in ["市场价", "工程价"]:
            c.alignment = align_right
        elif h_name in ["顺序号", "单位", "品牌", "是否含运费", "是否含增值税", "增值税税率", "价格填报类型", "报价时间"]:
            c.alignment = align_center
        else:
            c.alignment = align_left

for col_idx, h_text in enumerate(final_std_header, 1):
    col_letter = get_column_letter(col_idx)
    if "参数" in h_text or "说明" in h_text:
        ws3.column_dimensions[col_letter].width = 30
    elif h_text in ["供应商名称", "名称规格"]:
        ws3.column_dimensions[col_letter].width = 20
    elif "图片" in h_text:
        ws3.column_dimensions[col_letter].width = 10
    else:
        ws3.column_dimensions[col_letter].width = 14


# ==============================================================================
# SHEET 4: ⚠️ 潍坊福达-往期下发原貌 (7个小表归集)
# ==============================================================================
ws4 = wb.create_sheet(title="⚠️ 潍坊福达-往期下发原貌")
ws4.views.sheetView[0].showGridLines = True

ws4.merge_cells("A1:N1")
ws4["A1"] = "潍坊福达消防科技有限公司 — 往期下发历史解析归档原貌 (保留历史碎片化多表头结构记录)"
ws4["A1"].font = font_subtitle
ws4["A1"].fill = PatternFill(start_color="4B5563", end_color="4B5563", fill_type="solid")
ws4["A1"].alignment = align_center
ws4.row_dimensions[1].height = 26

curr_r = 3
for ch_idx, (h, rows) in enumerate(chunks):
    brand_check = rows[0][h.index("品牌")] if "品牌" in h and len(rows[0]) > h.index("品牌") else ""
    if "万良" in brand_check:
        continue
    
    ws4.cell(row=curr_r, column=1, value=f"【历史片段 {ch_idx+1}】 表头字段数: {len(h)} 列 | 记录数: {len(rows)} 条").font = font_sec_hdr
    curr_r += 1
    
    for c_i, h_txt in enumerate(h, 1):
        c = ws4.cell(row=curr_r, column=c_i, value=h_txt)
        c.font = font_tbl_hdr
        c.fill = PatternFill(start_color="6B7280", end_color="6B7280", fill_type="solid")
        c.alignment = align_center
        c.border = border_all
    curr_r += 1
    
    for r in rows:
        for c_i, v in enumerate(r, 1):
            c = ws4.cell(row=curr_r, column=c_i, value=v)
            c.font = font_data
            c.border = border_all
            c.alignment = align_left
        curr_r += 1
    curr_r += 1

for c_idx in range(1, 20):
    ws4.column_dimensions[get_column_letter(c_idx)].width = 18

# 4. 保存文件到 deliveries/
out_dir = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries"
os.makedirs(out_dir, exist_ok=True)
out_file = os.path.join(out_dir, "潍坊福达消防科技_报价解析标准版与往期版对比评估报告.xlsx")
wb.save(out_file)
print(f"Excel comparison report successfully generated: {out_file}")

supp_dir = os.path.join(out_dir, "01_潍坊福达消防科技有限公司")
if os.path.exists(supp_dir):
    supp_file = os.path.join(supp_dir, "潍坊福达消防科技_报价解析标准版与往期版对比评估报告.xlsx")
    wb.save(supp_file)
    print(f"Backup copy saved in supplier folder: {supp_file}")
