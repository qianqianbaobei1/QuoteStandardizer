import csv
import io
import re
import pandas as pd

# 1. 读取标准版 (Ground Truth)
std_path = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/data/ground_truth/潍坊福达消防科技_标准解析.csv"
std_df = pd.read_csv(std_path, dtype=str).fillna("")
std_dict = {row['顺序号']: row for _, row in std_df.iterrows()}

# 2. 读取第2条消息的文本并分段解析
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

print(f"Total chunks found: {len(chunks)}")

wffd_past_rows = {}
wanliang_past_rows = []

for idx, (header, rows) in enumerate(chunks):
    # 判断是万良还是潍坊福达
    first_row = rows[0]
    # header mapping
    for r in rows:
        row_dict = {header[i]: r[i] if i < len(r) else "" for i in range(len(header))}
        brand = row_dict.get("品牌", "")
        if "万良" in brand:
            wanliang_past_rows.append((header, row_dict))
        else:
            seq = row_dict.get("顺序号", "")
            wffd_past_rows[seq] = (header, row_dict)

print(f"Parsed past WFFD rows: {len(wffd_past_rows)} (Sequence range: {min(map(int, wffd_past_rows.keys()))} ~ {max(map(int, wffd_past_rows.keys()))})")
print(f"Parsed past Wanliang rows: {len(wanliang_past_rows)}")

# 3. 两两详细对比：标准版 WFFD vs 往期版 WFFD
print("\n" + "="*80)
print("=== 详细字段对比：标准真值版 (GT) VS 往期解析版 (Past) ===")
print("="*80)

field_diff_counts = {}
row_diff_details = []

all_seqs = sorted(list(std_dict.keys()), key=lambda x: int(x))

for seq in all_seqs:
    gt_row = std_dict[seq]
    if seq not in wffd_past_rows:
        print(f"Seq {seq} missing in past rows!")
        continue
    past_header, past_row = wffd_past_rows[seq]
    
    diffs_in_row = {}
    
    # 对比公共/标准核心字段
    # 1. 报价ID
    if gt_row.get("报价ID", "") != past_row.get("报价id", ""):
        diffs_in_row["报价ID"] = {"GT": gt_row.get("报价ID"), "Past": past_row.get("报价id")}
        field_diff_counts["报价ID"] = field_diff_counts.get("报价ID", 0) + 1
        
    # 2. 供应商名称
    gt_supplier = gt_row.get("供应商名称", "")
    past_supplier = past_row.get("供应商名称", "")
    if gt_supplier != past_supplier:
        diffs_in_row["供应商名称"] = {"GT": gt_supplier, "Past": past_supplier}
        field_diff_counts["供应商名称"] = field_diff_counts.get("供应商名称", 0) + 1
        
    # 3. 类别编码
    gt_cat = gt_row.get("类别编码", "")
    past_cat = past_row.get("类别编码", "")
    if gt_cat != past_cat:
        diffs_in_row["类别编码"] = {"GT": gt_cat, "Past": past_cat}
        field_diff_counts["类别编码"] = field_diff_counts.get("类别编码", 0) + 1
        
    # 4. 名称规格
    gt_name = gt_row.get("名称规格", "")
    past_name = past_row.get("名称规格", "")
    if gt_name != past_name:
        diffs_in_row["名称规格"] = {"GT": gt_name, "Past": past_name}
        field_diff_counts["名称规格"] = field_diff_counts.get("名称规格", 0) + 1
        
    # 5. 规格参数
    gt_spec = gt_row.get("规格参数", "")
    past_spec = past_row.get("规格参数", "")
    if gt_spec != past_spec:
        diffs_in_row["规格参数"] = {"GT": gt_spec, "Past": past_spec}
        field_diff_counts["规格参数"] = field_diff_counts.get("规格参数", 0) + 1
        
    # 6. 厂家原始规格参数
    gt_raw_spec = gt_row.get("厂家原始规格参数", "")
    past_raw_spec = past_row.get("厂家原始规格参数", "")
    if gt_raw_spec != past_raw_spec:
        diffs_in_row["厂家原始规格参数"] = {"GT": gt_raw_spec, "Past": past_raw_spec}
        field_diff_counts["厂家原始规格参数"] = field_diff_counts.get("厂家原始规格参数", 0) + 1
        
    # 7. 省市
    if gt_row.get("报价地区省", "") != past_row.get("报价地区省", "") or gt_row.get("报价地区市", "") != past_row.get("报价地区市", ""):
        diffs_in_row["报价地区省市"] = {"GT": f"{gt_row.get('报价地区省')}-{gt_row.get('报价地区市')}", "Past": f"{past_row.get('报价地区省')}-{past_row.get('报价地区市')}"}
        field_diff_counts["报价地区省市"] = field_diff_counts.get("报价地区省市", 0) + 1
        
    # 8. 价格与千分位
    gt_price = gt_row.get("市场价", "").replace(",", "")
    past_price = past_row.get("市场价", "").replace(",", "")
    if gt_price != past_price:
        diffs_in_row["价格数值"] = {"GT": gt_price, "Past": past_price}
        field_diff_counts["价格数值"] = field_diff_counts.get("价格数值", 0) + 1
    elif gt_row.get("市场价", "") != past_row.get("市场价", ""):
        diffs_in_row["价格千分位格式"] = {"GT": gt_row.get("市场价"), "Past": past_row.get("市场价")}
        field_diff_counts["价格千分位格式"] = field_diff_counts.get("价格千分位格式", 0) + 1
        
    # 9. 是否含增值税
    gt_vat_inc = gt_row.get("是否含增值税", "")
    past_vat_inc = past_row.get("是否含增值税", "")
    if gt_vat_inc != past_vat_inc:
        diffs_in_row["是否含增值税"] = {"GT": gt_vat_inc, "Past": past_vat_inc}
        field_diff_counts["是否含增值税"] = field_diff_counts.get("是否含增值税", 0) + 1
        
    # 10. 增值税税率
    gt_rate = gt_row.get("增值税税率", "")
    past_rate = past_row.get("增值税税率%", "") or past_row.get("增值税税率", "")
    if gt_rate.replace("%","") != past_rate.replace("%",""):
        diffs_in_row["增值税税率"] = {"GT": gt_rate, "Past": past_rate}
        field_diff_counts["增值税税率"] = field_diff_counts.get("增值税税率", 0) + 1
        
    # 11. 报价说明
    gt_notes = gt_row.get("报价说明", "")
    past_notes = past_row.get("报价说明", "")
    if gt_notes != past_notes:
        diffs_in_row["报价说明"] = {"GT": gt_notes, "Past": past_notes}
        field_diff_counts["报价说明"] = field_diff_counts.get("报价说明", 0) + 1
        
    # 12. 价格填报类型
    gt_ptype = gt_row.get("价格填报类型", "")
    past_ptype = past_row.get("价格填报类型", "")
    if gt_ptype != past_ptype:
        diffs_in_row["价格填报类型"] = {"GT": gt_ptype, "Past": past_ptype}
        field_diff_counts["价格填报类型"] = field_diff_counts.get("价格填报类型", 0) + 1
        
    # 13. 报价时间
    gt_time = gt_row.get("报价时间", "")
    past_time = past_row.get("报价时间", "")
    if gt_time != past_time:
        diffs_in_row["报价时间"] = {"GT": gt_time, "Past": past_time}
        field_diff_counts["报价时间"] = field_diff_counts.get("报价时间", 0) + 1
        
    # 14. 动态属性差异
    # 比较属性拆解
    past_dynamic = {k: v for k, v in past_row.items() if k not in [
        "报价id", "名称规格", "品牌", "厂家原始规格参数", "厂家报价文件页码", "厂家报价图片文件名",
        "顺序号", "报价地区省", "报价地区市", "单位", "市场价", "工程价", "工程价折扣比例",
        "折扣下限", "折扣上限", "折扣系列", "是否含运费", "运费说明", "是否含增值税",
        "税费说明", "增值税税率%", "运费结算方式", "报价说明", "价格填报类型", "项目类型",
        "工程用量", "报价时间", "图片文件夹", "图片"
    ] and v != ""}
    
    diffs_in_row["past_dynamic_keys"] = list(past_dynamic.keys())
    
    row_diff_details.append((seq, gt_row.get("名称规格"), diffs_in_row))

print(f"Field difference frequencies across 36 rows:")
for k, v in sorted(field_diff_counts.items(), key=lambda x: -x[1]):
    print(f"  - {k}: {v}/36 发生差异 ({v/36*100:.1f}%)")

