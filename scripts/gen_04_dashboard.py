# -*- coding: utf-8 -*-
import os
import sys
import re
import json
from pathlib import Path

BASE_DIR = Path('/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer')
TEMPLATE_PATH = BASE_DIR / 'templates' / 'dashboard_template.html'
OUT_HTML = BASE_DIR / 'deliveries' / '04_天兴消防器材集团有限公司' / '报价单解析质量看板.html'

with open(TEMPLATE_PATH, 'r', encoding='utf-8') as f:
    template_content = f.read()

vendor_csv = (
    "供应商,我方解析行数,真值行数,对齐行数,名称一致,单位一致,单价一致,单价漏填,三级一致,整行全对,必填属性全对,应有图片数,加载成功数\n"
    "天兴消防器材集团有限公司(本次交付),16,16,16,16,16,16,0,16,16,16,16,16\n"
    "线上系统历史解析(对照组),16,16,16,12,16,16,0,16,12,14,16,16\n"
)

errors_csv = (
    "指标,错误类型,条数,说明\n"
    "商务,税负误判不含税(线上),16,原单明确含增值税13%，线上系统16行全部误填为不含税\n"
    "解析,OCR错别字(线上),4,线上将下垂识别为下重、直立边墙识别为直立边填、洒水识别为酒水\n"
    "商务,工程折扣缺失(线上),1,原件明确现款现结8.9折(满10000只)，线上系统折扣未联动\n"
    "缓存,历史跨样本缓存污染(已根除),1,历史测试原料药缓存Key碰撞天兴图片，已清空并升级彩色双通道校验\n"
)

def replace_json_field(html, field_name, new_val):
    pattern = rf'("{field_name}":\s*)("[^"\\]*(?:\\.[^"\\]*)*"|`[^`]*`)'
    return re.sub(pattern, lambda m: m.group(1) + json.dumps(new_val, ensure_ascii=False), html)

dashboard_html = template_content
dashboard_html = replace_json_field(dashboard_html, 'vendors', vendor_csv)
dashboard_html = replace_json_field(dashboard_html, 'errors', errors_csv)
dashboard_html = dashboard_html.replace(
    "<title>报价单解析质量看板</title>",
    "<title>天兴消防器材集团有限公司 — 报价单解析质量看板 (D3交互式大盘)</title>"
)

with open(OUT_HTML, 'w', encoding='utf-8') as f:
    f.write(dashboard_html)
print(f"Generated: {OUT_HTML}, Size: {OUT_HTML.stat().st_size}")
