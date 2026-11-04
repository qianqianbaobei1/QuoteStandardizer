#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""原表材料名称/原表分类/品牌推断：不许拿标准分类名顶替原件品名，不许把公司名截断当品牌"""
import os
import sys
import unittest

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import main
from config import _load_category_codes

# 用仓库自带的工料机类别库，不依赖本机 .env 的 CATEGORY_CODE_FILE
CATEGORY_CODE_MAP = _load_category_codes(os.path.join(BASE_DIR, '工料机类别.xlsx'))


class SourceNameAndCategoryTest(unittest.TestCase):
    def test_material_name_prefers_source_column(self):
        cells = {'地区': '广东省', '三级类别名称': '塑料门', '材料名称': 'CPL覆膜钢质门',
                 '品牌': '南粤', '单位': '樘'}
        self.assertEqual(main._pick_row_field(cells, main.SRC_NAME_KEYS, main.SRC_NAME_EXCLUDE),
                         'CPL覆膜钢质门', '材料名称列必须照抄，不能被类别名称/供应商名称抢走')

    def test_category_column_ignores_code_like_headers(self):
        cells = {'三级类别名称': '塑料门', '三级类别编码': '090501'}
        self.assertEqual(main._pick_row_field(cells, main.SRC_CAT_KEYS, ('编码',)), '塑料门')

    def test_source_category_resolves_in_tree(self):
        self.assertEqual(CATEGORY_CODE_MAP.get('塑料门'), '090501')
        self.assertEqual(CATEGORY_CODE_MAP.get('挡烟垂壁'), '091709')

    def test_brand_not_truncated_mid_word(self):
        _, brand, _, _ = main.derive_supplier_and_brand_dynamic('星牌优时吉建筑材料有限公司', {})
        self.assertEqual(brand, '星牌优时吉', '不得截成“星牌优时吉建”这种不存在的品牌')
        _, brand2, _, _ = main.derive_supplier_and_brand_dynamic('广东爱因新材料有限公司', {})
        self.assertEqual(brand2, '爱因', '“新材料”整词剥离，不能留下半截“爱因新”')

    def test_source_brand_wins_over_inference(self):
        cells = {'品牌': '南粤', '材料名称': 'CPL覆膜钢质门'}
        self.assertEqual(main._pick_row_field(cells, ('品牌',)), '南粤')

    def test_model_prefix_stripped_from_material_name(self):
        self.assertEqual(main._strip_model_prefix('NLED4916 吊装线型灯'), '吊装线型灯')
        self.assertEqual(main._strip_model_prefix('LED T806P 灯管'), '灯管')
        self.assertEqual(main._strip_model_prefix('LEDA45J球泡'), '球泡')
        # 字母开头但无型号数字/是材料缩写：一律不动
        self.assertEqual(main._strip_model_prefix('CPL覆膜钢质门'), 'CPL覆膜钢质门')
        self.assertEqual(main._strip_model_prefix('UPVC排水管'), 'UPVC排水管')
        self.assertEqual(main._strip_model_prefix('HDPE双壁波纹管'), 'HDPE双壁波纹管')


class CategoryBucketTest(unittest.TestCase):
    def test_other_bucket_leaves_trace(self):
        from category_attribute_service import _finish
        code, src, notes = _finish('000101', 'dict', [])
        self.assertEqual(code, '000101')
        self.assertTrue(any('其他类别' in n for n in notes), '掉进 000101 必须留痕，不能静默当正经分类')

    def test_other_bucket_not_traced_for_real_codes(self):
        from category_attribute_service import _finish
        self.assertEqual(_finish('221701', 'dict', [])[2], [])


if __name__ == '__main__':
    unittest.main()
