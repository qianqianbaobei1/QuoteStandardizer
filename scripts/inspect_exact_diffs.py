import csv
import pandas as pd

std_path = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/data/ground_truth/潍坊福达消防科技_标准解析.csv"
std_df = pd.read_csv(std_path, dtype=str).fillna("")
std_dict = {row['顺序号']: row for _, row in std_df.iterrows()}

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

wffd_past_rows = {}
for header, rows in chunks:
    for r in rows:
        row_dict = {header[i]: r[i] if i < len(r) else "" for i in range(len(header))}
        brand = row_dict.get("品牌", "")
        if "万良" not in brand:
            seq = row_dict.get("顺序号", "")
            wffd_past_rows[seq] = (header, row_dict)

print("=== 1. 价格差异详细检查 ===")
for seq in sorted(wffd_past_rows.keys(), key=lambda x: int(x)):
    gt = std_dict[seq]
    past_h, past = wffd_past_rows[seq]
    gt_p = gt.get("市场价", "").replace('"', '').replace(',', '').strip()
    past_p = past.get("市场价", "").replace('"', '').replace(',', '').strip()
    if gt_p != past_p:
        print(f"Seq {seq} [{gt.get('名称规格')}] GT市场价: '{gt.get('市场价')}' vs Past市场价: '{past.get('市场价')}'")

print("\n=== 2. 字段名与核心基础字段差异对比（抽样 Seq 1, 6, 21, 28） ===")
for seq in ["1", "6", "21", "28"]:
    gt = std_dict[seq]
    past_h, past = wffd_past_rows[seq]
    print(f"\n--- Seq {seq} ({gt.get('名称规格')}) ---")
    keys_to_compare = [
        ("报价ID", "报价id"),
        ("供应商名称", "供应商名称"),
        ("类别编码", "类别编码"),
        ("名称规格", "名称规格"),
        ("规格参数", "规格参数"),
        ("厂家原始规格参数", "厂家原始规格参数"),
        ("报价地区省", "报价地区省"),
        ("报价地区市", "报价地区市"),
        ("是否含增值税", "是否含增值税"),
        ("税费说明", "税费说明"),
        ("增值税税率", "增值税税率%"),
        ("报价说明", "报价说明"),
        ("价格填报类型", "价格填报类型"),
        ("报价时间", "报价时间"),
        ("运费说明", "运费说明")
    ]
    for gt_k, past_k in keys_to_compare:
        gt_v = gt.get(gt_k, "<缺失列>")
        past_v = past.get(past_k, "<缺失列>")
        diff_flag = " != " if gt_v != past_v else " == "
        print(f"  {gt_k} vs {past_k}: GT='{gt_v}' {diff_flag} Past='{past_v}'")

print("\n=== 3. 动态属性拆解列与值对比（抽样 Seq 1, 15, 21, 28） ===")
for seq in ["1", "15", "21", "28"]:
    gt = std_dict[seq]
    past_h, past = wffd_past_rows[seq]
    print(f"\n--- Seq {seq} 属性对比 ---")
    # 提取两者非空且不属于公共字段的属性
    common_cols = [
        "报价ID", "报价id", "供应商名称", "类别编码", "品牌", "规格参数", "厂家原始规格参数",
        "厂家报价文件页码", "厂家报价图片文件名", "顺序号", "报价地区省", "报价地区市", "单位",
        "市场价", "工程价", "工程价折扣比例", "折扣下限", "折扣上限", "折扣系列", "运费说明",
        "是否含运费", "税费说明", "增值税税率", "增值税税率%", "运费结算方式", "是否含增值税",
        "报价说明", "价格填报类型", "项目类型", "工程用量", "报价时间", "图片", "图片文件夹"
    ]
    gt_attrs = {k: v for k, v in gt.items() if k not in common_cols and v != ""}
    past_attrs = {k: v for k, v in past.items() if k not in common_cols and v != ""}
    print(f"  GT 属性:   {gt_attrs}")
    print(f"  Past 属性: {past_attrs}")

