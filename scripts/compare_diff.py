import csv
import io
import json
import pandas as pd

# 读取标准版潍坊福达
std_file = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/data/ground_truth/潍坊福达消防科技_标准解析.csv"
std_df = pd.read_csv(std_file, dtype=str)

print(f"Standard WFFD rows: {len(std_df)}, columns: {len(std_df.columns)}")
