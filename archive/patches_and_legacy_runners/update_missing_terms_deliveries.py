# -*- coding: utf-8 -*-
"""
update_missing_terms_deliveries.py - 针对报价说明/折扣说明缺失文件的专项高精度追回工具
"""

import os
import re
from pathlib import Path
from typing import Dict, Any, List
import openpyxl
from openpyxl.styles import Font, Alignment

from local_terms_engine import extract_terms_from_image_local

DELIVERY_DIR = Path('/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries')

TARGET_VENDORS = [
    '中建西部建设建材科学研究院有限公司',
    '北京启达世纪装饰材料有限公司',
    '宁波欧琳科技股份有限公司',
    '山东拜尔建材有限公司',
    '山东领拓新材料有限公司',
    '广东平安消防实业有限公司',
    '广东雷士照明科技有限公司',
    '广州同晟新型建材有限公司',
    '敏华电器（广州）有限公司',
    '武汉第二电线电缆有限公司',
    '深圳市麦驰智慧技术有限公司',
    '营口新山鹰报警设备有限公司'
]

# 预定义的真实文件附件路径映射，确保 100% 命中
MANUAL_ATTACHMENTS = {
    '中建西部建设建材科学研究院有限公司': [
        '/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries_backup_20261103_1828/周伟E化11.03_unpacked/周伟E化11.03/中建西部建设建材科学研究院有限公司/报价说明+折扣记录.png'
    ],
    '北京启达世纪装饰材料有限公司': [
        '/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案/2_张艳鸽（34家）/1_原始报价文件/北京启达世纪装饰材料有限公司/报价说明新的.png',
        '/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案/2_张艳鸽（34家）/1_原始报价文件/北京启达世纪装饰材料有限公司/折扣记录新的.png'
    ],
    '宁波欧琳科技股份有限公司': [
        '/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries_backup_20261103_1828/周伟E化11.03_unpacked/周伟E化11.03/宁波欧琳科技股份有限公司/报价说明和折扣记录.jpg'
    ],
    '山东拜尔建材有限公司': [
        '/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案/2_张艳鸽（34家）/1_原始报价文件/山东拜尔建材有限公司/报价说明-山东拜尔-20260917_01.jpg',
        '/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案/2_张艳鸽（34家）/1_原始报价文件/山东拜尔建材有限公司/折扣说明-山东拜尔-20260917_01.jpg'
    ],
    '山东领拓新材料有限公司': [
        '/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries_backup_20261103_1828/周伟E化11.03_unpacked/周伟E化11.03/山东领拓新材料有限公司/折扣说明10.28.png'
    ],
    '广东平安消防实业有限公司': [
        '/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案/3_景文敏（36家）/1_原始报价文件/11广东平安消防实业有限公司/平安报价说明+折扣记录.png'
    ],
    '广东雷士照明科技有限公司': [
        '/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案/1_程雅洁（30家）/1_原始报价文件/广东雷士照明科技有限公司/报价说明+折扣记录.png'
    ],
    '广州同晟新型建材有限公司': [
        '/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries_backup_20261103_1828/周伟E化11.03_unpacked/周伟E化11.03/广州同晟新型建材有限公司/报价说明+折扣记录.png'
    ],
    '敏华电器（广州）有限公司': [
        '/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案/2_张艳鸽（34家）/1_原始报价文件/敏华电器（广州）有限公司/报价说明和折扣记录.jpg'
    ],
    '深圳市麦驰智慧技术有限公司': [
        '/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案/1_程雅洁（30家）/1_原始报价文件/深圳市麦驰智慧技术有限公司/报价说明+折扣记录 (3).png'
    ],
    '营口新山鹰报警设备有限公司': [
        '/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案/3_景文敏（36家）/1_原始报价文件/21营口新山鹰报警设备有限公司/报价说明26.9.28.png',
        '/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案/3_景文敏（36家）/1_原始报价文件/21营口新山鹰报警设备有限公司/折扣说明26.9.28.png'
    ]
}

def get_terms_for_target(name: str) -> Dict[str, Any]:
    if name == '武汉第二电线电缆有限公司':
        return {
            'quote_desc': '出厂价（2025年06月17日发布；花线2026年10月19日发布）',
            'quote_type': '出厂价',
            'is_tax_included': '含税',
            'is_freight_included': '不含',
            'freight_desc': '出厂自提或运费另计'
        }
        
    imgs = MANUAL_ATTACHMENTS.get(name, [])
    merged = {}
    for img_p in imgs:
        if os.path.exists(img_p):
            t = extract_terms_from_image_local(img_p)
            for k, v in t.items():
                if v and not merged.get(k):
                    merged[k] = v
                elif v and k in ('quote_desc', 'freight_desc', 'discount_series'):
                    if v not in str(merged.get(k, '')):
                        merged[k] = f"{merged[k]}；{v}".strip('；')
    return merged

def update_file(name: str):
    xlsx_path = DELIVERY_DIR / f"{name}-标准解析.xlsx"
    if not xlsx_path.exists():
        print(f"[Skip] 文件不存在: {xlsx_path.name}")
        return

    terms = get_terms_for_target(name)
    if not terms:
        print(f"[Warning] 未解析到有效条款: {name}")
        return

    print(f"[*] 正在更新: {name}...")
    wb = openpyxl.load_workbook(str(xlsx_path))
    if '标准化报价' not in wb.sheetnames:
        return
    ws = wb['标准化报价']
    headers = [str(ws.cell(1, c).value or '').strip() for c in range(1, ws.max_column + 1)]
    
    col_map = {f: headers.index(f) + 1 for f in [
        '工程价折扣比例', '折扣下限', '折扣上限', '折扣系列',
        '运费说明', '是否含运费', '税费说明', '增值税税率',
        '运费结算方式', '是否含增值税', '报价说明', '价格填报类型'
    ] if f in headers}

    body_font = Font(name='宋体', size=10)
    multiline_left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

    for r in range(2, ws.max_row + 1):
        if '报价说明' in col_map and terms.get('quote_desc'):
            cell = ws.cell(r, col_map['报价说明'], terms['quote_desc'])
            cell.font = body_font
            cell.alignment = multiline_left_align
            
        if '工程价折扣比例' in col_map and terms.get('discount_ratio') is not None:
            cur_d = ws.cell(r, col_map['工程价折扣比例']).value
            if cur_d in (None, ''):
                ws.cell(r, col_map['工程价折扣比例'], terms['discount_ratio'])
                
        if '折扣系列' in col_map and terms.get('discount_series'):
            cell = ws.cell(r, col_map['折扣系列'], terms['discount_series'])
            cell.font = body_font
            cell.alignment = multiline_left_align

        if '是否含运费' in col_map and terms.get('is_freight_included'):
            ws.cell(r, col_map['是否含运费'], terms['is_freight_included'])
            
        if '运费说明' in col_map and terms.get('freight_desc'):
            cell = ws.cell(r, col_map['运费说明'], terms['freight_desc'])
            cell.font = body_font
            cell.alignment = multiline_left_align

        if '是否含增值税' in col_map and terms.get('is_tax_included'):
            ws.cell(r, col_map['是否含增值税'], terms['is_tax_included'])
            
        if '增值税税率' in col_map and terms.get('tax_rate') is not None:
            ws.cell(r, col_map['增值税税率'], terms['tax_rate'])
            
        if '税费说明' in col_map and terms.get('tax_desc'):
            cell = ws.cell(r, col_map['税费说明'], terms['tax_desc'])
            cell.font = body_font
            cell.alignment = multiline_left_align

        if '价格填报类型' in col_map and terms.get('quote_type'):
            cell = ws.cell(r, col_map['价格填报类型'], terms['quote_type'])
            cell.font = body_font

    wb.save(str(xlsx_path))
    wb.close()
    print(f"[✓ 成功更新落盘] {name}")
    print(f"    - 报价说明: {terms.get('quote_desc')}")
    print(f"    - 工程折扣: {terms.get('discount_ratio')} ({terms.get('discount_series')})")
    print(f"    - 运费说明: {terms.get('freight_desc')}")

def run():
    print(f"[*] 开始专项追回 {len(TARGET_VENDORS)} 家供应商原始材料中的报价说明与折扣说明...")
    for v in TARGET_VENDORS:
        update_file(v)
    print(f"\n[SUCCESS] 专项追回与回填全部完成！")

if __name__ == '__main__':
    run()

