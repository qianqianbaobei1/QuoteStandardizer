import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

# -*- coding: utf-8 -*-
"""
并发改造的回归自检（不联网，全部 mock 外部调用）：
1. PDF 页级并发：多页同时发出、页序不变、单页失败只隔离该页
2. Gemini 批次并发：多批同时发出、全局页码映射不变
3. 分类召回缓存：同品名重复行命中缓存且互不污染
4. 批处理多供应商并发：多家同时解析、结果按原顺序
"""
import json
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

ROOT = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
SAMPLE_PDF = _os.path.join(ROOT, "deliveries", "1101_benchmark", "raw",
                           "广西未来新材料有限公司", "10.30_产品报价-石材_未来新材.pdf")
PAGE_COST = 0.35


class TestPdfPageConcurrency(unittest.TestCase):
    @unittest.skipUnless(_os.path.exists(SAMPLE_PDF), "样例 PDF 不在仓库里")
    def test_pages_run_concurrently_and_keep_order(self):
        from pdf_extractor import PDFTableExtractor
        n_pages = len(fitz.open(SAMPLE_PDF))
        self.assertGreater(n_pages, 1, "需要多页样例才测得出并发")

        def fake_vision(self, page, base):
            time.sleep(PAGE_COST)
            self._tl.title_sink.append(f"{base}抬头")
            if base == "第3页":
                raise RuntimeError("模拟单页失败")
            return [[["品名", "规格"], [f"{base}行", "20mm"]]]

        ext = PDFTableExtractor(SAMPLE_PDF)
        t0 = time.time()
        with patch.object(PDFTableExtractor, '_grids_from_vision', fake_vision):
            wb = ext.to_workbook()
        elapsed = time.time() - t0

        self.assertLess(elapsed, PAGE_COST * n_pages * 0.75, f"没有并发: {elapsed:.2f}s")
        self.assertEqual(wb.sheetnames, [f"第{i}页" for i in range(1, n_pages + 1)],
                         "失败页必须留在原页序位置")
        self.assertIn("识别失败", str(wb["第3页"].cell(1, 1).value))
        self.assertEqual(ext.page_titles, [f"第{i}页抬头" for i in range(1, n_pages + 1)
                                           if i != 3], "页标题必须按页序，不能被线程打乱")
        self.assertEqual(len(ext.failed_pages), 1)


class TestGeminiBatchConcurrency(unittest.TestCase):
    @unittest.skipUnless(_os.path.exists(SAMPLE_PDF), "样例 PDF 不在仓库里")
    def test_batches_run_concurrently_with_global_page_numbers(self):
        from gemini_pdf_extractor import GeminiPDFTableExtractor

        class FixedPageCount(GeminiPDFTableExtractor):
            def __init__(self, path, pages):
                super().__init__(path)
                self.total_pages = pages

        def fake_prepare(self, page_range=None):
            return f"{page_range[0]}:{page_range[1]}".encode()

        def fake_call(self, client, pdf_bytes):
            start, end = map(int, pdf_bytes.decode().split(":"))
            time.sleep(PAGE_COST)
            return json.dumps({"pages": [
                {"page_number": i, "page_title": "抬头",
                 "tables": [{"cells": [["品名"], [f"行{start + i}"]]}]}
                for i in range(1, end - start + 1)]})

        ext = FixedPageCount(SAMPLE_PDF, 10)
        t0 = time.time()
        with patch.object(GeminiPDFTableExtractor, '_prepare_enhanced_pdf_bytes', fake_prepare), \
                patch.object(GeminiPDFTableExtractor, '_call_gemini_raw', fake_call):
            wb = ext.to_workbook()
        elapsed = time.time() - t0

        self.assertLess(elapsed, PAGE_COST * 4 * 0.75, f"批次没有并发: {elapsed:.2f}s")
        self.assertEqual(wb.sheetnames, [f"第{i}页" for i in range(1, 11)])
        self.assertIn("行10", [c.value for c in wb["第10页"]["A"]], "最后一批的页码偏移错了")
        self.assertIn("行7", [c.value for c in wb["第7页"]["A"]])


class TestCategoryCandidatesCache(unittest.TestCase):
    def test_repeat_rows_hit_cache_without_bleeding(self):
        from category_attribute_service import retrieve_category_candidates

        args = ("石材", "中国黑花岗岩", "品种:中国黑;厚度(mm):20", {"石材"}, 5)
        t0 = time.time()
        first = retrieve_category_candidates(*args)
        cold = time.time() - t0
        self.assertTrue(first, "样例词条应召回到候选，否则本测试无意义")

        t0 = time.time()
        second = retrieve_category_candidates(*args)
        warm = time.time() - t0
        self.assertEqual(first, second)
        self.assertLess(warm, 0.002, f"没命中缓存: 冷 {cold:.3f}s / 热 {warm:.3f}s")

        second[0]["level3"] = "被调用方改写"
        third = retrieve_category_candidates(*args)
        self.assertNotEqual(third[0]["level3"], "被调用方改写", "缓存必须返回深拷贝")


class TestBatchVendorConcurrency(unittest.TestCase):
    def test_vendors_run_concurrently(self):
        import importlib.util
        # 不能`from cli import ...`：site-packages 里有个同名 `cli` 包会遮蔽本仓库目录
        spec = importlib.util.spec_from_file_location(
            "rbs_under_test", _os.path.join(ROOT, "cli", "run_batch_standardize.py"))
        rbs = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(rbs)

        def fake_run_single_vendor(vendor_dir, output_dir):
            time.sleep(PAGE_COST)
            return {"vendor": vendor_dir.name, "input_file": "x.pdf", "file_type": ".pdf",
                    "status": "PASS", "output_file": "x.xlsx", "item_count": 1,
                    "duration": PAGE_COST, "error_msg": None, "failures": None}

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "src"
            for name in ("A公司", "B公司", "C公司", "D公司"):
                (src / name).mkdir(parents=True)
            with patch.object(rbs, "run_single_vendor", fake_run_single_vendor):
                t0 = time.time()
                results = rbs.run_batch(src, Path(tmp) / "out", "自检批次", workers=4)
                elapsed = time.time() - t0

        self.assertLess(elapsed, PAGE_COST * 4 * 0.75, f"供应商没有并发: {elapsed:.2f}s")
        self.assertEqual([r["vendor"] for r in results], ["A公司", "B公司", "C公司", "D公司"],
                         "并发后结果顺序必须仍按目录字典序")


if __name__ == "__main__":
    unittest.main(verbosity=2)
