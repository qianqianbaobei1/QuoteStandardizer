import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

# -*- coding: utf-8 -*-
"""
隐藏内容口径：Excel 里被隐藏的工作表 / 行 / 列一律不解析。
用真实原表（武汉第二电线电缆）做锚点断言：厂商把「系统结算价」整表隐藏、该表 58,413 行隐藏，
其中第 58885 行是 WDZBN-KYJYP2-450/750V 9×1 —— 加了口径后这行不能进交付。
"""
import os
import sys
import unittest

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
CLI_DIR = os.path.join(BASE_DIR, "cli")
if CLI_DIR not in sys.path:
    sys.path.insert(0, CLI_DIR)

import openpyxl

from config import SKIP_HIDDEN_CONTENT
from extractor import hidden_axes, hidden_content_report
from main import _collect_matrix_rows, _detect_matrix
from zip_standardizer_pipeline import collapse_to_single_sheet

REAL = os.path.join(BASE_DIR, "scratch", "武汉第二电线电缆有限公司", "2026电线出厂全系列价格表.xlsx")


class TestHiddenContent(unittest.TestCase):
    def test_default_rule_is_on(self):
        self.assertTrue(SKIP_HIDDEN_CONTENT, "隐藏内容默认必须不解析")

    @unittest.skipUnless(os.path.exists(REAL), "缺少真实原表样本")
    def test_real_workbook_hidden_axes(self):
        wb = openpyxl.load_workbook(REAL, data_only=True)
        self.assertEqual(wb['系统结算价'].sheet_state, 'hidden')
        hid_rows, hid_cols = hidden_axes(wb['系统结算价'])
        self.assertEqual(len(hid_rows), 58413)
        self.assertIn(58885, hid_rows)          # 9×1 那行本身是隐藏行
        self.assertEqual(wb['系统结算价'].cell(58885, 4).value, '9×1')
        self.assertEqual(hidden_axes(wb['电线1']), (set(), set()))

        # 只读流式审计脚本用 zip 直读的同一口径，两者必须一致
        rep = hidden_content_report(REAL)
        self.assertEqual(rep['系统结算价']['hidden'], True)
        self.assertEqual(rep['系统结算价']['hidden_rows'], hid_rows)
        self.assertEqual(rep['电线1']['hidden_rows'], set())
        self.assertEqual([s for s, i in rep.items() if i['hidden']], ['铝合金电缆', '系统结算价'])

    def test_matrix_skips_hidden_row_and_column(self):
        """矩阵表：隐藏行/列的价格格不得产出物料行"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['规格', 'BV', 'BVR', 'BVVB'])
        for i, spec in enumerate([1, 1.5, 2.5], start=2):
            ws.append([spec, 100 * i, 200 * i, 300 * i])   # 行 2/3/4
        ws.append([4, 500, 600, 700])                       # 行 5
        ws.row_dimensions[5].hidden = True                  # 隐藏档位
        ws.column_dimensions['C'].hidden = True             # 隐藏一整列型号

        hid_rows, hid_cols = hidden_axes(ws)
        self.assertEqual((hid_rows, hid_cols), ({5}, {3}))
        mx = _detect_matrix(ws, hid_rows, hid_cols)
        self.assertIsNotNone(mx)
        self.assertEqual(mx['value_cols'], [2, 4])           # 只剩 B 列与 D 列
        raw_rows, records = _collect_matrix_rows(ws, 'S', mx, {}, hid_rows, hid_cols)
        self.assertEqual(len(raw_rows), 6)                  # 3 个可见档位 × 2 个可见型号
        self.assertEqual(sorted(r['id'] for r in raw_rows),
                         ['S!R2C2', 'S!R2C4', 'S!R3C2', 'S!R3C4', 'S!R4C2', 'S!R4C4'])
        self.assertNotIn('4', [r['raw_cells_dict']['规格型号'] for r in raw_rows])

    def test_merge_drops_hidden_sheets_and_rows(self):
        """多表合并（zip 流水线）同样只合可见内容"""
        wb = openpyxl.Workbook()
        vis = wb.active
        vis.title = '可见表'
        vis.append(['名称', '单价'])
        vis.append(['A', 1])
        vis.append(['B', 2])
        vis.row_dimensions[3].hidden = True          # 隐藏行
        hid = wb.create_sheet('隐藏表')
        hid.append(['名称', '单价'])
        hid.append(['Z', 999])
        hid.sheet_state = 'hidden'

        merged = collapse_to_single_sheet(wb)
        values = [[c.value for c in row] for row in merged.iter_rows()]
        flat = [v for row in values for v in row if v not in (None, '')]
        self.assertIn('A', flat)
        self.assertNotIn('B', flat)      # 隐藏行被丢弃
        self.assertNotIn('Z', flat)      # 隐藏表被丢弃

    def test_multi_column_span_and_zero_dimension_hidden(self):
        """测试连续多列隐藏区间与行高/列宽拉成 0 的隐藏过滤"""
        wb = openpyxl.Workbook()
        ws = wb.active
        # 模拟 Excel 底层 XML 的连续列隐藏 <col min="5" max="8" hidden="1"/>
        cd = ws.column_dimensions['E']
        cd.min = 5
        cd.max = 8
        cd.hidden = True

        # 模拟行高拉为 0 和列宽拉为 0
        ws.row_dimensions[10].height = 0.0
        ws.column_dimensions['J'].width = 0.0

        hid_rows, hid_cols = hidden_axes(ws)
        self.assertIn(10, hid_rows)
        for c in (5, 6, 7, 8, 10):  # 5~8(E~H) 及 10(J)
            self.assertIn(c, hid_cols, f"列 {c} 必须在隐藏列集合中")

    def test_extract_sheet_sample_rows_masks_hidden_cols(self):
        """测试给 LLM 结构推导的样本行：隐藏列置空，隐藏行跳过"""
        from extractor import extract_sheet_sample_rows
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['序号', '品名', '内部底价', '外报价'])  # R1
        ws.append(['1', '镀锌钢管', '50.0', '80.0'])      # R2 (可见)
        ws.append(['2', '废弃行', '10.0', '20.0'])        # R3 (隐藏)

        ws.row_dimensions[3].hidden = True
        ws.column_dimensions['C'].hidden = True  # C 列（第 3 列内部底价）隐藏

        samples = extract_sheet_sample_rows(ws, max_rows=10, max_cols=10)
        self.assertEqual(len(samples), 2)  # R1 和 R2，R3 被跳过
        r1_cells = samples[0]["cells"]
        r2_cells = samples[1]["cells"]
        # 第 3 列 (C 列) 必须被置空为 ""，不能泄露 "内部底价" 或 "50.0"
        self.assertEqual(r1_cells[2], "")
        self.assertEqual(r2_cells[2], "")
        self.assertEqual(r2_cells[1], "镀锌钢管")
        self.assertEqual(r2_cells[3], "80.0")

    def test_build_headers_skips_hidden_cols(self):
        """测试表头读取跳过隐藏列，防止隐藏列抢占可见列表头名"""
        from main import _build_headers
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(['品名', '隐藏单位', '规格', '真实单位', '单价'])
        # 假设第 2 列是隐藏列
        hid_cols = {2}
        headers = _build_headers(ws, header_row=1, col_st=1, col_ed=5, hid_cols=hid_cols)
        self.assertNotIn(2, headers)
        self.assertEqual(headers[1], '品名')
        self.assertEqual(headers[3], '规格')
        self.assertEqual(headers[4], '真实单位')  # 真实单位未被抢占或编号
        self.assertEqual(headers[5], '单价')


if __name__ == '__main__':
    unittest.main()
