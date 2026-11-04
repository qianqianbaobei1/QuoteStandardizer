# -*- coding: utf-8 -*-
"""不变式护栏：终审只提建议 (I1/I3/I6)、ZIP 流程落盘后对账、配图失败留痕 (I7)"""
import io
import os
import tempfile
import unittest
import zipfile
from unittest.mock import patch

import openpyxl

from models import CanonicalMaterialItem
from zip_standardizer_pipeline import ZipQuoteStandardizerPipeline

HEADERS = ['报价ID', '名称规格', '规格参数', '厂家原始规格参数', '单位', '市场价', '厂家报价文件页码', '图片']
ROWS = [
    ['WL001', '热镀锌钢管', '', '壁厚:3.25', 'm', 128.5, '1', ''],
    ['WL002', '闸阀中丝扣', '', 'DN50', '个', 66, '1', ''],
]


def _make_zip(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '钢管'
    ws.append(HEADERS)
    for r in ROWS:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('报价数据.xlsx', buf.getvalue())
    return buf.getvalue()


def _cell(ws, row, header):
    hdr = [c.value for c in ws[1]]
    return ws.cell(row, hdr.index(header) + 1)


class TestZipPipelineInvariants(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.zip_path = os.path.join(self.tmp.name, '测试供应商_1.zip')
        self.xlsx_bytes = _make_zip(self.zip_path)
        self.pipeline = ZipQuoteStandardizerPipeline(output_dir=os.path.join(self.tmp.name, 'out'))
        # 模型想改名、改价、换单位
        fake = {'need_healing': True, 'summary': 't', 'corrections': [
            {'seq': 2, 'reason': '[纠正错位]', 'fixes': {'名称规格': '闸阀', '市场价': 1, '单位': '台'}}]}
        p = patch.object(self.pipeline.auditor.client, 'generate_json', return_value=fake)
        p.start()
        self.addCleanup(p.stop)

    def test_auditor_only_suggests_and_output_reconciles(self):
        out = self.pipeline.process_zip(self.zip_path)
        self.assertEqual(self.pipeline.last_failures, [])

        ws = openpyxl.load_workbook(out)['标准化报价']
        self.assertEqual(_cell(ws, 3, '名称规格').value, '闸阀中丝扣', '品名以原表为准')
        self.assertEqual(_cell(ws, 3, '单位').value, '个', '单位不替换')
        self.assertEqual(_cell(ws, 3, '市场价').value, 66, '价格只取原表')
        note = _cell(ws, 3, '名称规格').comment
        self.assertIsNotNone(note)
        self.assertIn('[模型建议]', note.text)
        self.assertIn('闸阀', note.text)

    def test_reconcile_catches_price_change(self):
        out = self.pipeline.process_zip(self.zip_path)
        wb = openpyxl.load_workbook(out)
        _cell(wb['标准化报价'], 2, '市场价').value = 99
        wb.save(out)
        failures = self.pipeline._verify_against_source(self.pipeline._source_grid(self.xlsx_bytes), out)
        self.assertTrue(any('市场价' in f for f in failures), failures)


class TestIllustrationFailureIsRecorded(unittest.TestCase):
    def test_api_error_becomes_review_note(self):
        import material_illustration_service as mis

        def failing_fetch(materials, **kwargs):
            mis.default_illustration_service.last_errors = ['URLError: timed out']
            return {}

        item = CanonicalMaterialItem(quote_id='Q1', supplier='测试', brand='', category_code='', category_name='镀锌钢管')
        with patch('material_illustration_service.ENABLE_MATERIAL_ILLUSTRATION', True), \
                patch.object(mis.default_illustration_service, 'fetch_illustrations_batch', side_effect=failing_fetch), \
                patch.object(mis.default_illustration_service, 'last_errors', []):
            matched = mis.match_illustrations_for_canonical_items([item])
        self.assertEqual(matched, 0)
        self.assertTrue(any('配图接口调用失败' in n for n in item.note_list()), item.note_list())


if __name__ == '__main__':
    unittest.main()
