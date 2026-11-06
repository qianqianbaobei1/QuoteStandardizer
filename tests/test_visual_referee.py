import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

# -*- coding: utf-8 -*-
"""
原图三方裁决引擎的自检（不联网）：
最关键的是裁决方向不能被写反 —— 线上值来自 theirs，标准值来自 ours。
写反了报告会把"谁错"整个颠倒，而报告本身看起来毫无异常。
"""
import importlib.util
import json
import os
import tempfile
import unittest
from unittest.mock import patch

import openpyxl

ROOT = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))


def _load(mod_name, filename):
    spec = importlib.util.spec_from_file_location(mod_name, _os.path.join(ROOT, "cli", filename))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


HEADER = ['名称规格', '规格参数', '单位', '市场价', '厂家报价文件页码']


def _write_xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '标准化报价'
    ws.append(HEADER)
    for r in rows:
        ws.append(r)
    wb.save(path)


class TestVisualRefereeDirection(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.ours = _os.path.join(self.tmp.name, 'ours.xlsx')     # 我们的解析 = 标准/基准
        self.theirs = _os.path.join(self.tmp.name, 'theirs.xlsx')  # 线上系统
        _write_xlsx(self.ours, [['不锈钢栏杆', '品种:圆管式', 'm', 580, '第1页']])
        _write_xlsx(self.theirs, [['不锈钢栏杆', '品种:圆管式', 'm', 680, '第1页']])

    def tearDown(self):
        self.tmp.cleanup()

    def test_online_value_comes_from_theirs_and_verdict_maps_correctly(self):
        vr = _load('visual_referee_under_test', 'visual_referee.py')
        ev = vr._load_evaluate()
        collected = vr.collect_discrepancies(ev, self.ours, self.theirs, '', 'F1', max_per_file=50)
        price = [t for t in collected['diffs'] if t['field'] == '市场价']
        self.assertEqual(len(price), 1, collected['diffs'])
        task = price[0]
        self.assertEqual(str(task['standard']), '580', "标准值必须来自 ours")
        self.assertEqual(str(task['online']), '680', "线上值必须来自 theirs")

        # 裁判说"原图支持 580（我们的标准值）" → 线上是错的
        fake = {'verdict': 'STD_WIN', 'ground_truth_from_image': '580',
                'evidence_location': '第1页单价列', 'std_fix_reason': ''}
        with patch.object(vr, 'adjudicate', return_value=fake):
            verdicts = vr.adjudicate_batch([task])
        self.assertEqual(verdicts[0]['verdict'], 'STD_WIN')

        per_file = [dict(file_id='F1', rows_ours=1, rows_theirs=1, only_a=0, only_b=0,
                         diffs=[dict(task, **verdicts[0])])]
        out = os.path.join(self.tmp.name, 'out')
        xlsx_path, html_path = vr.write_reports(per_file, out, 'test')
        self.assertTrue(os.path.exists(xlsx_path) and os.path.exists(html_path))
        wb = openpyxl.load_workbook(xlsx_path)
        self.assertEqual(wb.sheetnames,
                         ['1_汇总', '2_逐行判定', '3_差异明细', '4_问题归因', '5_标准待修正'])
        detail = [r for r in wb['3_差异明细'].iter_rows(min_row=2, values_only=True)]
        self.assertEqual(detail[0][4], '680', "明细里线上值必须是我们没解析出的线上原值")
        self.assertEqual(detail[0][5], '580')
        with open(html_path, encoding='utf-8') as f:
            self.assertIn('4.2rem', f.read())

    def test_no_source_marks_unadjudicated_instead_of_blaming_online(self):
        """没有原图时必须如实记为未裁决 —— 绝不能默认把线上判错"""
        vr = _load('visual_referee_under_test2', 'visual_referee.py')
        task = {'source_file': None, 'page': 1, 'field': '市场价', 'online': '680',
                'standard': '580', 'name': '不锈钢栏杆', 'raw_spec': '', 'code': 'E03_PRICE'}
        with patch.object(vr, 'VISION_API_KEY', 'k'), patch.object(vr, 'VISION_BASE_URL', 'http://x'), \
                patch.object(vr, 'VISION_MODEL', 'm'), patch.object(vr, 'page_image_b64', return_value=''):
            out = vr.adjudicate(task)
        self.assertEqual(out['verdict'], 'UNADJUDICATED')

    def test_find_source_file_prefers_hint_then_pdf(self):
        vr = _load('visual_referee_under_test3', 'visual_referee.py')
        d = os.path.join(self.tmp.name, 'src', '甲')
        os.makedirs(d)
        for name in ('a报价单.pdf', 'b扫描件.png'):
            open(os.path.join(d, name), 'wb').close()
        self.assertTrue(vr.find_source_file(os.path.join(self.tmp.name, 'src')).endswith('.pdf'))
        self.assertTrue(vr.find_source_file(os.path.join(self.tmp.name, 'src'), 'b扫描件').endswith('.png'))
        self.assertIsNone(vr.find_source_file(os.path.join(self.tmp.name, '不存在')))

    def test_page_number_parsing(self):
        vr = _load('visual_referee_under_test4', 'visual_referee.py')
        self.assertEqual(vr.page_number_of({'厂家报价文件页码': '第3页'}), 3)
        self.assertEqual(vr.page_number_of({'厂家报价文件页码': 2}), 2)
        self.assertEqual(vr.page_number_of({}), 1)


if __name__ == '__main__':
    unittest.main(verbosity=2)
