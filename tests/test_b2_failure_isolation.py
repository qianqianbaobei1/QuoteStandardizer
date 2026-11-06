import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
B2 验收单元测试: 模拟某页视觉调用抛异常，断言程序不崩溃、其余页有输出、对账有失败记录、退出码为 2。
"""

import os
import sys
import unittest
import tempfile
import shutil
import fitz
import openpyxl
from unittest.mock import patch

# 确保 QuoteStandardizer 根目录在 sys.path 中
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from main import process_pdf, main
from models import CanonicalMaterialItem
from config import DEFAULT_TEMPLATE_PATH


class TestB2FailureIsolation(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="b2_test_")
        self.output_dir = os.path.join(self.test_dir, "output")
        os.makedirs(self.output_dir, exist_ok=True)

        # 生成一个包含 2 页的测试 PDF (无文字层，强制进入视觉通道)
        self.pdf_path = os.path.join(self.test_dir, "test_2pages_scan.pdf")
        doc = fitz.open()
        p1 = doc.new_page(width=595, height=842)
        p1.draw_rect(fitz.Rect(50, 50, 500, 700), color=(0, 0, 0))
        p2 = doc.new_page(width=595, height=842)
        p2.draw_rect(fitz.Rect(50, 50, 500, 700), color=(0, 0, 0))
        doc.save(self.pdf_path)
        doc.close()

    def tearDown(self):
        if os.path.exists(self.test_dir):
            shutil.rmtree(self.test_dir)

    def test_page_failure_isolation_and_reconciliation(self):
        """测试单页视觉识别抛异常时的隔离机制与对账记录"""
        def mock_grids_from_vision(self_extractor, page, base):
            pno = page.number + 1
            if pno == 1:
                # 第 1 页成功返回
                return [[
                    ["品名", "规格型号", "单位", "单价"],
                    ["防水卷材", "1.5mm", "m²", "120.00"]
                ]]
            elif pno == 2:
                # 第 2 页模拟抛异常
                raise RuntimeError("视觉识别超时异常(测试模拟)")
            return []

        def mock_standardize_batch(self_agent, inp, supplier_hint="", brand_hint=""):
            return [
                {
                    "id": item["id"],
                    "standard_name": item.get("raw_name") or "防水卷材",
                    "structured_specs": {"厚度": "1.5mm"},
                    "uncertainty_note": ""
                }
                for item in inp
            ]

        with patch("pdf_extractor.PDFTableExtractor._grids_from_vision", side_effect=mock_grids_from_vision, autospec=True), \
             patch("llm_agent.DeepSeekStandardizer.standardize_batch", side_effect=mock_standardize_batch, autospec=True), \
             patch("main.verify_terms_supplier", return_value={"supplier": "测试建材"}), \
             patch("main.derive_supplier_and_brand_dynamic", return_value=("测试建材", "测试品牌", "CSJC", [])), \
             patch("main.EXPORT_INTEGRITY_CERTIFICATE", True):

            # 执行解析
            out_file, failures = process_pdf(self.pdf_path, DEFAULT_TEMPLATE_PATH, self.output_dir)

            # 1. 断言程序不崩溃，成功产出输出文件
            self.assertTrue(os.path.exists(out_file), "输出 Excel 应当生成成功")
            self.assertTrue(len(failures) > 0, "failures 列表中必须包含第 2 页失败信息")

            # 2. 检查 failures 中是否包含第 2 页异常说明
            has_p2_failure = any("第2页" in str(f) and "视觉识别超时异常" in str(f) for f in failures)
            self.assertTrue(has_p2_failure, f"未在 failures 中检测到第 2 页失败记录: {failures}")

            # 3. 读取 Excel 断言其余页有输出 (第 1 页防水卷材)
            wb = openpyxl.load_workbook(out_file, data_only=True)
            self.assertIn("标准化报价", wb.sheetnames)
            ws_main = wb["标准化报价"]
            self.assertGreaterEqual(ws_main.max_row, 2, "标准化报价 Sheet 必须包含第 1 页的物料输出")
            found_item = False
            for r in range(2, ws_main.max_row + 1):
                row_vals = [ws_main.cell(r, c).value for c in range(1, ws_main.max_column + 1)]
                if any(v == "防水卷材" or v == 120 or v == "120" for v in row_vals):
                    found_item = True
                    break
            self.assertTrue(found_item, "未在输出中找到第 1 页物料")

            # 4. 交付约定：主文件包含「标准化报价」；对账明细在独立体检报告文件里
            cert_path = str(out_file).replace('.xlsx', '_对账与防欺瞒体检报告.xlsx')
            self.assertTrue(os.path.exists(cert_path), f'缺少独立对账报告: {cert_path}')
            ws_recon = openpyxl.load_workbook(cert_path, data_only=True)['原表对账']
            found_recon_error = False
            for r in range(3, ws_recon.max_row + 1):
                status_val = str(ws_recon.cell(r, 2).value or "")
                ref_val = str(ws_recon.cell(r, 1).value or "")
                if "第2页" in ref_val and status_val.startswith("异常"):
                    found_recon_error = True
                    self.assertIn("视觉识别超时异常", status_val)
                    break
            self.assertTrue(found_recon_error, "未在原表对账 Sheet 中找到第 2 页的异常记录")

    def test_main_cli_exit_code_2_on_failed_pages(self):
        """测试包含失败页时 main() 最终退出码为 2"""
        def mock_grids_from_vision(self_extractor, page, base):
            if page.number == 1:
                raise RuntimeError("单页识别致命故障")
            return [[["品名", "单价"], ["测试品", "100.00"]]]

        def mock_standardize_batch(self_agent, inp, supplier_hint="", brand_hint=""):
            return [
                {
                    "id": item["id"],
                    "standard_name": item.get("raw_name") or "测试品",
                    "structured_specs": {},
                    "uncertainty_note": ""
                }
                for item in inp
            ]

        test_args = ["main.py", "--input", self.pdf_path, "--output_dir", self.output_dir]
        with patch.object(sys, "argv", test_args), \
             patch("pdf_extractor.PDFTableExtractor._grids_from_vision", side_effect=mock_grids_from_vision, autospec=True), \
             patch("llm_agent.DeepSeekStandardizer.standardize_batch", side_effect=mock_standardize_batch, autospec=True), \
             patch("main.verify_terms_supplier", return_value={"supplier": "测试供方"}), \
             patch("main.derive_supplier_and_brand_dynamic", return_value=("测试供方", "测试品牌", "TEST", [])):

            with self.assertRaises(SystemExit) as cm:
                main()
            self.assertEqual(cm.exception.code, 2, f"在存在失败页时，main() 退出码应当严格为 2，实际为 {cm.exception.code}")


if __name__ == "__main__":
    unittest.main()
