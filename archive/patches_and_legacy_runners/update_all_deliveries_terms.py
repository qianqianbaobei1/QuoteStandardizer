# -*- coding: utf-8 -*-
"""
update_all_deliveries_terms.py - 全域交付文件商务条款、报价说明与折扣说明深度追回与回填工具 (极速缓存版)
"""

import os
import re
import json
from pathlib import Path
from typing import Optional, List, Dict, Any
import openpyxl
from openpyxl.styles import Font, Alignment

from local_terms_engine import extract_terms_from_image_local

DELIVERY_DIR = Path('/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries')
SOURCE_ROOTS = [
    Path('/Users/qianqianawodebaobei/Desktop/数据异化/四人建材报价标准解析与评差全量汇总/02_四位经办人分类完整档案'),
    Path('/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries_backup_20261103_1828/周伟E化11.03_unpacked/周伟E化11.03'),
    Path('/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/scratch')
]

TERMS_KEYWORDS = ['报价说明', '折扣说明', '折扣记录', '折扣单', '商务条款', '说明']

# 预构建供应商目录索引
VENDOR_DIR_CACHE: Dict[str, Path] = {}

def build_dir_cache():
    print("[*] 正在预构建全域供应商源目录索引...")
    for root in SOURCE_ROOTS:
        if not root.exists():
            continue
        for d in root.rglob('*'):
            if d.is_dir() and not d.name.startswith('.') and not '__MACOSX' in str(d):
                clean_name = re.sub(r'^\d+[\._、\s]*', '', d.name).strip()
                if clean_name and clean_name not in VENDOR_DIR_CACHE:
                    VENDOR_DIR_CACHE[clean_name] = d
    print(f"[✓] 索引构建完成，共捕获 {len(VENDOR_DIR_CACHE)} 个候选供应商目录。")

def find_vendor_source_dir(vendor_name: str) -> Optional[Path]:
    clean_name = re.sub(r'^\d+[\._、\s]*', '', vendor_name).strip()
    clean_name = clean_name.replace('-标准解析', '').replace('.xlsx', '').strip()
    
    if clean_name in VENDOR_DIR_CACHE:
        return VENDOR_DIR_CACHE[clean_name]
    for k, v in VENDOR_DIR_CACHE.items():
        if clean_name in k or k in clean_name:
            return v
    return None

def find_terms_images_in_dir(vendor_dir: Path) -> List[Path]:
    found = []
    try:
        for f in vendor_dir.iterdir():
            if f.is_file() and f.suffix.lower() in ('.png', '.jpg', '.jpeg'):
                if any(k in f.name for k in TERMS_KEYWORDS):
                    found.append(f)
    except Exception:
        pass
    return found

def get_terms_for_vendor(vendor_name: str) -> Dict[str, Any]:
    vdir = find_vendor_source_dir(vendor_name)
    merged_terms = {}
    
    if vdir:
        img_paths = find_terms_images_in_dir(vdir)
        for img in img_paths:
            terms = extract_terms_from_image_local(str(img))
            if terms:
                for k, v in terms.items():
                    if v and not merged_terms.get(k):
                        merged_terms[k] = v
                    elif v and k in ('quote_desc', 'freight_desc', 'discount_series'):
                        if v not in str(merged_terms.get(k, '')):
                            merged_terms[k] = f"{merged_terms[k]}；{v}".strip('；')

    # 特殊纯 Excel 原单处理（如武汉第二电线电缆）
    if '武汉第二电线电缆' in vendor_name:
        merged_terms.update({
            'quote_desc': '出厂价（2025年06月17日发布；花线2026年10月19日发布）',
            'quote_type': '出厂价',
            'is_tax_included': '含税',
            'is_freight_included': '不含',
            'freight_desc': '出厂自提或运费另计'
        })
        
    return merged_terms

def update_delivery_xlsx(xlsx_path: Path, terms: Dict[str, Any]) -> bool:
    if not terms:
        return False
        
    try:
        wb = openpyxl.load_workbook(str(xlsx_path))
        if '标准化报价' not in wb.sheetnames:
            return False
        ws = wb['标准化报价']
        
        headers = [str(ws.cell(1, c).value or '').strip() for c in range(1, ws.max_column + 1)]
        
        col_map = {}
        target_fields = [
            '工程价折扣比例', '折扣下限', '折扣上限', '折扣系列',
            '运费说明', '是否含运费', '税费说明', '增值税税率',
            '运费结算方式', '是否含增值税', '报价说明', '价格填报类型'
        ]
        for f in target_fields:
            if f in headers:
                col_map[f] = headers.index(f) + 1

        updated_any = False
        body_font = Font(name='宋体', size=10)
        center_align = Alignment(horizontal='center', vertical='center')
        multiline_left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

        for r in range(2, ws.max_row + 1):
            # 1. 报价说明
            if '报价说明' in col_map and terms.get('quote_desc'):
                cur_val = str(ws.cell(r, col_map['报价说明']).value or '').strip()
                if not cur_val or cur_val in ('None', ''):
                    cell = ws.cell(r, col_map['报价说明'], terms['quote_desc'])
                    cell.font = body_font
                    cell.alignment = multiline_left_align
                    updated_any = True

            # 2. 折扣与折扣系列
            if '工程价折扣比例' in col_map and terms.get('discount_ratio') is not None:
                cur_val = ws.cell(r, col_map['工程价折扣比例']).value
                if cur_val in (None, ''):
                    ws.cell(r, col_map['工程价折扣比例'], terms['discount_ratio'])
                    updated_any = True
            if '折扣系列' in col_map and terms.get('discount_series'):
                cur_val = str(ws.cell(r, col_map['折扣系列']).value or '').strip()
                if not cur_val:
                    cell = ws.cell(r, col_map['折扣系列'], terms['discount_series'])
                    cell.font = body_font
                    cell.alignment = multiline_left_align
                    updated_any = True

            # 3. 运费与税费
            if '是否含运费' in col_map and terms.get('is_freight_included'):
                cur_val = str(ws.cell(r, col_map['是否含运费']).value or '').strip()
                if not cur_val:
                    ws.cell(r, col_map['是否含运费'], terms['is_freight_included'])
                    updated_any = True
            if '运费说明' in col_map and terms.get('freight_desc'):
                cur_val = str(ws.cell(r, col_map['运费说明']).value or '').strip()
                if not cur_val:
                    cell = ws.cell(r, col_map['运费说明'], terms['freight_desc'])
                    cell.font = body_font
                    cell.alignment = multiline_left_align
                    updated_any = True

            if '是否含增值税' in col_map and terms.get('is_tax_included'):
                cur_val = str(ws.cell(r, col_map['是否含增值税']).value or '').strip()
                if not cur_val:
                    ws.cell(r, col_map['是否含增值税'], terms['is_tax_included'])
                    updated_any = True
            if '增值税税率' in col_map and terms.get('tax_rate') is not None:
                cur_val = ws.cell(r, col_map['增值税税率']).value
                if cur_val in (None, ''):
                    ws.cell(r, col_map['增值税税率'], terms['tax_rate'])
                    updated_any = True
            if '税费说明' in col_map and terms.get('tax_desc'):
                cur_val = str(ws.cell(r, col_map['税费说明']).value or '').strip()
                if not cur_val:
                    ws.cell(r, col_map['税费说明'], terms['tax_desc'])
                    updated_any = True

            # 4. 价格填报类型
            if '价格填报类型' in col_map and terms.get('quote_type'):
                cur_val = str(ws.cell(r, col_map['价格填报类型']).value or '').strip()
                if not cur_val:
                    ws.cell(r, col_map['价格填报类型'], terms['quote_type'])
                    updated_any = True

        if updated_any:
            wb.save(str(xlsx_path))
            return True
    except Exception as e:
        print(f"[Update Error] {xlsx_path.name}: {e}")
    return False

def run():
    build_dir_cache()
    xlsx_files = sorted([f for f in DELIVERY_DIR.glob('*-标准解析.xlsx')])
    print(f"[*] 开始检查并更新 deliveries/ 下的 {len(xlsx_files)} 份交付文件...")

    success_count = 0
    updated_files = []

    for f in xlsx_files:
        vname = f.stem.replace('-标准解析', '')
        terms = get_terms_for_vendor(vname)
        if terms:
            updated = update_delivery_xlsx(f, terms)
            if updated:
                success_count += 1
                updated_files.append((vname, terms))
                print(f"[✓ 已追回并更新] {vname}:")
                if terms.get('quote_desc'):
                    print(f"    - 报价说明: {terms['quote_desc']}")
                if terms.get('discount_ratio') or terms.get('discount_series'):
                    print(f"    - 工程折扣: {terms.get('discount_ratio')} ({terms.get('discount_series')})")
                if terms.get('freight_desc'):
                    print(f"    - 运费说明: {terms.get('freight_desc')}")

    print(f"\n[SUCCESS] 全域商务条款与报价说明追回完成！成功回填更新 {success_count} 份文件。")

if __name__ == '__main__':
    run()
