import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""多 Sheet 收口成单张「标准化报价」的单元测试。"""
import os
import sys
import unittest

import openpyxl
from openpyxl.comments import Comment

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
CLI_DIR = os.path.join(BASE_DIR, "cli")
if CLI_DIR not in sys.path:
    sys.path.insert(0, CLI_DIR)

from zip_standardizer_pipeline import collapse_to_single_sheet


class MergeSingleSheetTest(unittest.TestCase):
    def _wb(self):
        wb = openpyxl.Workbook()
        a = wb.active
        a.title = '184301'
        a.append(['名称规格', '市场价', '接口口径(mm)'])
        a.append(['UPVC地漏', 16.74, 'De110'])
        a['B2'].comment = Comment('推断来源: 图片', 'QS')

        b = wb.create_sheet('184701')
        b.append(['名称规格', '市场价', '规格(mm)'])
        b.append(['UPVC清扫口', 3.93, 'De160'])
        b.append(['UPVC清扫口', 1.8, 'De110'])
        return wb

    def test_merge(self):
        ws = collapse_to_single_sheet(self._wb())
        self.assertEqual(ws.parent.sheetnames, ['标准化报价'])
        self.assertEqual([c.value for c in ws[1]],
                         ['名称规格', '市场价', '接口口径(mm)', '规格(mm)'])
        self.assertEqual(ws.max_row - 1, 3, '行数必须等于各 Sheet 之和，不得丢行')
        self.assertEqual(ws['C2'].value, 'De110')
        self.assertEqual(ws['D3'].value, 'De160', '动态列必须按表头名落到自己的列')
        self.assertIsNone(ws['C3'].value, '别的 Sheet 没有的列留空')
        self.assertIsNotNone(ws['B2'].comment, '待复核批注必须跟着行迁移')

    def test_single_sheet_only_renamed(self):
        wb = openpyxl.Workbook()
        wb.active.title = '071101'
        wb.active.append(['名称规格'])
        ws = collapse_to_single_sheet(wb)
        self.assertEqual(ws.parent.sheetnames, ['标准化报价'])


if __name__ == '__main__':
    unittest.main()
