# -*- coding: utf-8 -*-
"""
Excel 底层资产与通用结构提取器 (深度强化版)
核心防线：
1. 二维合并单元格 (Merged Cells) 空间拓扑穿透向下/向右填充 (Forward Fill)
   避免合并区域第 2~N 行读取为空
2. 文本与公差精度保护 (禁止将 '1.0' 强转为 1)
3. 稳健数值价格清洗 (过滤货币符号与千分位)
4. WPS DISPIMG 与 Drawing 浮动图底层解包与空间坐标锚定
"""
import os
import re
import posixpath
import zipfile
from decimal import Decimal
import xml.etree.ElementTree as ET
import openpyxl

def unmerge_and_fill_cells(ws):
    """
    预处理工作表：将所有合并单元格拆解，并将左上角真实值向下、向右穿透填充。
    """
    merged_ranges = list(ws.merged_cells.ranges)
    for merged_range in merged_ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        top_left_value = ws.cell(row=min_row, column=min_col).value
        # 释放合并，并把值赋予矩形区域内的所有单元格
        ws.unmerge_cells(str(merged_range))
        if top_left_value is not None:
            for r in range(min_row, max_row + 1):
                for c in range(min_col, max_col + 1):
                    ws.cell(row=r, column=c, value=top_left_value)

def safe_clean_price(val) -> float:
    """
    通用稳健价格清洗转换：
    处理 ¥, $, 千分位逗号，提取首个连续的有效浮点数
    """
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).replace(',', '').replace('¥', '').replace('$', '').strip()
    m = re.search(r'(\d+(?:\.\d+)?)', s)
    if m:
        try:
            return float(m.group(1))
        except ValueError:
            return 0.0
    return 0.0

def _decimals_from_format(number_format) -> int:
    """从 Excel 数字格式 (如 '0.00'、'#,##0.0') 中读出显示的小数位数；无法判断返回 -1"""
    if not number_format or number_format == 'General':
        return -1
    m = re.search(r'0\.(0+)', number_format)
    if m:
        return len(m.group(1))
    if re.search(r'[0#]', number_format):
        return 0
    return -1

def safe_cell_str(val, number_format=None) -> str:
    """
    保留原始精度的字符串清洗：
    - 有数字格式时按 Excel 显示的小数位输出 (格式 '0.0' 的 1 → '1.0')
    - 无格式信息时使用最短无损表示 (1.0 → '1.0'，0.00001 → '0.00001'，不做四舍五入)
    """
    if val is None:
        return ""
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, float):
        decimals = _decimals_from_format(number_format)
        if decimals >= 0:
            return f"{val:.{decimals}f}"
        if number_format == 'General' and val.is_integer():
            return str(int(val))
        return format(Decimal(repr(val)), 'f')
    return str(val).replace('\r', '').replace('\n', ' ').strip()

def cell_str(cell) -> str:
    """按单元格自身数字格式输出字符串"""
    return safe_cell_str(cell.value, getattr(cell, 'number_format', None))

def _resolve_part(base_part: str, target: str) -> str:
    """按 OPC 规范把关系 Target 解析为 zip 内绝对路径 (相对于 base_part 所在目录)"""
    if target.startswith('/'):
        return target.lstrip('/')
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))

class ExcelAssetExtractor:
    def __init__(self, xlsx_path):
        self.xlsx_path = xlsx_path
        self.ci_map = {}                # WPS cellimage ID -> media path
        self.sheet_drawing_anchors = {} # sheet_name -> {row: [(col, media_path)]}
        self._parse_underlying_media()

    def _parse_underlying_media(self):
        """解析 xlsx 容器底层的图片资产"""
        if not os.path.exists(self.xlsx_path):
            return

        try:
            with zipfile.ZipFile(self.xlsx_path, 'r') as z:
                namelist = set(z.namelist())

                # 1. 解析 WPS 单元格图片
                if 'xl/_rels/cellimages.xml.rels' in namelist and 'xl/cellimages.xml' in namelist:
                    rels = {
                        elem.attrib.get('Id'): elem.attrib.get('Target')
                        for elem in ET.fromstring(z.read('xl/_rels/cellimages.xml.rels'))
                    }
                    root_ci = ET.fromstring(z.read('xl/cellimages.xml'))
                    for elem in root_ci.iter():
                        if elem.tag.endswith('cellImage'):
                            pic_name, r_embed = None, None
                            for child in elem.iter():
                                if child.tag.endswith('cNvPr'):
                                    pic_name = child.attrib.get('name')
                                if child.tag.endswith('blip'):
                                    for k, v in child.attrib.items():
                                        if k.endswith('embed'):
                                            r_embed = v
                            if pic_name and r_embed:
                                t = rels.get(r_embed, '')
                                if t:
                                    self.ci_map[pic_name] = _resolve_part('xl/cellimages.xml', t)

                # 2. 解析 Sheet 与 Drawing 锚点
                if 'xl/_rels/workbook.xml.rels' in namelist and 'xl/workbook.xml' in namelist:
                    wb_rels = {
                        elem.attrib.get('Id'): elem.attrib.get('Target')
                        for elem in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
                    }
                    wb_root = ET.fromstring(z.read('xl/workbook.xml'))
                    sheet_name_to_xml = {}
                    for elem in wb_root.iter():
                        if elem.tag.endswith('sheet'):
                            sname = elem.attrib.get('name')
                            r_id = next((v for k, v in elem.attrib.items() if k.endswith('id')), None)
                            if sname and r_id and r_id in wb_rels:
                                sheet_name_to_xml[sname] = _resolve_part('xl/workbook.xml', wb_rels[r_id])

                    for sname, sxml in sheet_name_to_xml.items():
                        srels_name = os.path.dirname(sxml) + '/_rels/' + os.path.basename(sxml) + '.rels'
                        if srels_name in namelist:
                            srels = {
                                elem.attrib.get('Id'): elem.attrib.get('Target')
                                for elem in ET.fromstring(z.read(srels_name))
                            }
                            sroot = ET.fromstring(z.read(sxml))
                            drawing_xml = None
                            for elem in sroot.iter():
                                if elem.tag.endswith('drawing'):
                                    r_id = next((v for k, v in elem.attrib.items() if k.endswith('id')), None)
                                    dt = srels.get(r_id) if r_id else None
                                    if dt:
                                        drawing_xml = _resolve_part(sxml, dt)
                            if drawing_xml and drawing_xml in namelist:
                                drels_name = os.path.dirname(drawing_xml) + '/_rels/' + os.path.basename(drawing_xml) + '.rels'
                                drels = {
                                    elem.attrib.get('Id'): elem.attrib.get('Target')
                                    for elem in ET.fromstring(z.read(drels_name))
                                } if drels_name in namelist else {}
                                
                                droot = ET.fromstring(z.read(drawing_xml))
                                row_map = {}
                                for anchor in droot:
                                    from_r, from_c, blip_id = None, None, None
                                    for ch in anchor.iter():
                                        if ch.tag.endswith('from'):
                                            for sub in ch:
                                                if sub.tag.endswith('row') and sub.text:
                                                    from_r = int(sub.text) + 1
                                                if sub.tag.endswith('col') and sub.text:
                                                    from_c = int(sub.text) + 1
                                        if ch.tag.endswith('blip'):
                                            blip_id = next((v for k, v in ch.attrib.items() if k.endswith('embed')), None)
                                    if blip_id and from_r and from_c:
                                        target = drels.get(blip_id)
                                        if target:
                                            mpath = _resolve_part(drawing_xml, target)
                                            row_map.setdefault(from_r, []).append((from_c, mpath))
                                self.sheet_drawing_anchors[sname] = row_map
        except Exception as e:
            print(f"[ExcelAssetExtractor Warning] 底层媒体解包异常: {e}")

    def get_images_for_region(self, ws, sname, row_idx, col_min=1, col_max=30, col_tolerance=1):
        """
        获取特定 Sheet 在特定行和列区间内的所有图片路径（支持多图）
        col_tolerance: 浮动图允许越出子表边界的列数；左右并排子表时应传 0，避免图片串到相邻子表
        """
        imgs = []
        for c in range(col_min, col_max + 1):
            val = str(ws.cell(row_idx, c).value or '')
            if 'DISPIMG' in val:
                for m in re.finditer(r'\"(ID_[A-F0-9]+)\"', val):
                    img_id = m.group(1)
                    if img_id in self.ci_map and self.ci_map[img_id] not in imgs:
                        imgs.append(self.ci_map[img_id])
                        
        anchors = self.sheet_drawing_anchors.get(sname, {}).get(row_idx, [])
        for c, mp in anchors:
            if col_min - col_tolerance <= c <= col_max + col_tolerance:
                if mp not in imgs:
                    imgs.append(mp)
        return imgs

    def extract_image_bytes(self, media_path):
        """读取指定媒体路径的原图字节流"""
        if not os.path.exists(self.xlsx_path):
            return None
        try:
            with zipfile.ZipFile(self.xlsx_path, 'r') as z:
                if media_path in z.namelist():
                    return z.read(media_path)
        except Exception:
            pass
        return None

    def extract_sheet_sample_rows(self, ws, max_rows=10, max_cols=25):
        return extract_sheet_sample_rows(ws, max_rows, max_cols)


def extract_sheet_sample_rows(ws, max_rows=10, max_cols=25):
    """提取 Sheet 前几行样本用于 LLM 结构推导 (已做安全精度转换)"""
    samples = []
    for r in range(1, min(ws.max_row + 1, max_rows + 1)):
        row_data = []
        for c in range(1, min(ws.max_column + 1, max_cols + 1)):
            row_data.append(cell_str(ws.cell(r, c)))
        if any(row_data):
            samples.append({"row_idx": r, "cells": row_data})
    return samples
