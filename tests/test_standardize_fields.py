# -*- coding: utf-8 -*-
"""standardize_workbook 拆出的逐字段取值函数：每个字段的取值规则与待复核说明"""
import unittest
from unittest.mock import patch

from main import _resolve_name_and_category, _resolve_price, _resolve_region, _resolve_specs, _resolve_unit, _row_brand

TERMS = {"province": "湖北省", "city": "武汉市"}


class TestFieldResolvers(unittest.TestCase):
    def test_price_always_from_source(self):
        self.assertEqual(_resolve_price(128.5, 128.5), (128.5, []))
        price, notes = _resolve_price(128.5, 120)
        self.assertEqual(price, 128.5)
        self.assertIn("已保留原表价", notes[0])
        price, notes = _resolve_price(0, "99")
        self.assertEqual(price, 0)
        self.assertIn("模型推测 99.0，未采用", notes[0])

    def test_unit_change_is_noted(self):
        self.assertEqual(_resolve_unit("个", "个"), ("个", []))
        self.assertEqual(_resolve_unit("", ""), ("", ["原表无计量单位"]))
        unit, notes = _resolve_unit("台", "个")
        self.assertEqual(unit, "台")
        self.assertIn("单位由原表“个”改为“台”", notes)

    def test_row_brand(self):
        self.assertEqual(_row_brand({"品牌": "联塑"}, "推断"), ("联塑", [], True))
        brand, notes, from_table = _row_brand({"品牌": "1340"}, "推断")
        self.assertEqual((brand, from_table), ("推断", False))
        self.assertIn("不像品牌", notes[0])
        self.assertEqual(_row_brand({"品牌系列": "A"}, "推断"), ("推断", [], False))
        self.assertEqual(_row_brand({"品牌": ""}, "推断"), ("推断", [], False))

    def test_region_row_overrides_terms(self):
        self.assertEqual(_resolve_region({}, TERMS)[:2], ("湖北省", "武汉市"))
        prov, _, notes = _resolve_region({"报价地区": "京津冀+山东"}, TERMS)
        self.assertEqual(prov, "北京市、天津市、河北省、山东省")
        self.assertTrue(notes)

    def test_specs_verbatim_kept(self):
        item = {"raw_cells_dict": {"规格型号": "厚度(mm):1.5;是否曝露使用:是"}, "raw_full_line": "厚度(mm):1.5;是否曝露使用:是"}
        _, specs, _, notes = _resolve_specs(item, {"structured_specs": {"厚度": "1.5mm"}}, "卷材")
        self.assertEqual(specs.get("厚度(mm)"), "1.5")
        self.assertNotIn("厚度", specs)

    def test_name_and_category_notes(self):
        with patch("main.fetch_category_attributes", return_value=None):
            name, code, _, notes = _resolve_name_and_category({"standard_name": ""}, {"raw_full_line": "x"})
        self.assertEqual((name, code), ("", ""))
        self.assertIn("未识别出物料名称", notes)


if __name__ == "__main__":
    unittest.main()
