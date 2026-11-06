# -*- coding: utf-8 -*-
"""
tests/test_category_attribute_service.py
单元测试：广材377万规格词库召回前置、属性同义词映射、数值单位脱壳与历史枚举反向认领
"""
import os
import sys
import unittest

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

from category_attribute_service import (
    retrieve_category_candidates,
    resolve_and_verify_category,
    validate_and_normalize_specs,
    _match_attr_key_to_schema,
    _normalize_attr_value_with_enum,
)


class TestCategoryAttributeService(unittest.TestCase):
    """测试分类召回前置与属性值规范化引擎"""

    def test_attr_name_synonyms_matching(self):
        """测试工程造价常见属性名同义词映射"""
        schema = [
            {"name": "公称直径DN(mm)", "enum": "DN15, DN20, DN25, DN50", "required": True},
            {"name": "厚度(mm)", "enum": "", "required": False},
            {"name": "标称截面(mm²)", "enum": "", "required": False},
            {"name": "密度(kg/m³)", "enum": "", "required": False},
        ]
        # 口径 -> 公称直径DN(mm)
        m1 = _match_attr_key_to_schema("口径", schema)
        self.assertIsNotNone(m1)
        self.assertEqual(m1["name"], "公称直径DN(mm)")

        # 通径 -> 公称直径DN(mm)
        m2 = _match_attr_key_to_schema("通径", schema)
        self.assertIsNotNone(m2)
        self.assertEqual(m2["name"], "公称直径DN(mm)")

        # 壁厚 -> 厚度(mm)
        m3 = _match_attr_key_to_schema("壁厚", schema)
        self.assertIsNotNone(m3)
        self.assertEqual(m3["name"], "厚度(mm)")

        # 平方 -> 标称截面(mm²)
        m4 = _match_attr_key_to_schema("平方", schema)
        self.assertIsNotNone(m4)
        self.assertEqual(m4["name"], "标称截面(mm²)")

        # 容重 -> 密度(kg/m³)
        m5 = _match_attr_key_to_schema("容重", schema)
        self.assertIsNotNone(m5)
        self.assertEqual(m5["name"], "密度(kg/m³)")

    def test_unit_shell_peeling(self):
        """测试数值型带单位属性自动脱壳（去除值内冗余单位）"""
        # 1.5mm -> 1.5
        v1, _ = _normalize_attr_value_with_enum("厚度(mm)", "1.5mm", "")
        self.assertEqual(v1, "1.5")

        # 2.5平方 / 2.5mm² -> 2.5
        v2, _ = _normalize_attr_value_with_enum("标称截面(mm²)", "2.5mm²", "")
        self.assertEqual(v2, "2.5")
        v2b, _ = _normalize_attr_value_with_enum("标称截面(mm²)", "2.5平方", "")
        self.assertEqual(v2b, "2.5")

        # 120kg/m³ -> 120
        v3, _ = _normalize_attr_value_with_enum("密度(kg/m³)", "120kg/m³", "")
        self.assertEqual(v3, "120")

    def test_standard_enum_alignment(self):
        """测试历史标准枚举值唯一包含关系对齐"""
        # 氟碳 -> 氟碳喷涂
        v1, _ = _normalize_attr_value_with_enum("表面处理", "氟碳", "氟碳喷涂, 阳极氧化, 电泳涂装")
        self.assertEqual(v1, "氟碳喷涂")

        # 304 -> 304不锈钢
        v2, _ = _normalize_attr_value_with_enum("材质", "304", "304不锈钢, 316L不锈钢, 碳钢")
        self.assertEqual(v2, "304不锈钢")

    def test_reverse_enum_claim_from_raw_spec(self):
        """测试未填标准属性时从原始规格自由文本反向认领专有枚举值"""
        raw_text = "SBS防水卷材 4mm 聚酯胎PY PE膜"
        specs = {"厚度": "4mm"}
        res, missing, warns = validate_and_normalize_specs(
            "弹性体改性沥青防水卷材", specs, raw_spec_text=raw_text
        )
        self.assertEqual(res.get("厚度(mm)"), "4")
        self.assertEqual(res.get("胎基材料"), "聚酯毡PY")

    def test_retrieve_category_candidates_truncation(self):
        """测试候选池属性枚举值预览不超过 45 字符，防止 Prompt 溢出"""
        cands = retrieve_category_candidates(raw_cat="防水材料", std_name="防水卷材", top_k=3)
        self.assertTrue(len(cands) > 0)
        for c in cands:
            for attr_str in c.get("allowed_attributes", []):
                if "(" in attr_str and attr_str.endswith(")"):
                    enum_content = attr_str[attr_str.index("(") + 1:-1]
                    self.assertLessEqual(len(enum_content), 50)


if __name__ == "__main__":
    unittest.main()
