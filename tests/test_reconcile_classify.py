import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""交付件 vs 原表对账的分类判定测试（typo/append/fmt/fill/drop/diff + 致命判定）"""
import os
import sys
import unittest

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import reconcile_vs_source as R


class ClassifyTest(unittest.TestCase):
    def run_one(self, hdr, src_row, dst_row):
        src = (hdr, [src_row])
        dst = (hdr, [dst_row])
        rep = R.reconcile('测试', src, dst)
        return rep, rep['details']

    def test_kinds(self):
        hdr = ['名称规格', '规格参数', '市场价', '单位', '报价说明', '图片']
        src = ['广播劝放', '厚度:1.5;基价:18000', 10, 'm', '现款现结', '甲.png']
        dst = ['广播功放', '厚度: 1.5;\n基价: 18000', 10, 'm', '现款现结；追加说明', None]
        rep, details = self.run_one(hdr, src, dst)
        kinds = {d[2]: d[3] for d in details}
        self.assertEqual(kinds['名称规格'], 'typo', '错别字级差异应归为纠错，不是冲突')
        self.assertEqual(rep['per_field'][('规格参数', 0)]['fmt'], 1, '只差分隔符/空白不算改写')
        self.assertEqual(kinds['报价说明'], 'append', '原文保留在后面追加不算改写')
        self.assertEqual(kinds['图片'], 'drop', '我们清空了原图名 = 丢失')
        self.assertEqual(rep['stats']['same'], 2, '名称/规格/说明之外的价格和单位应完全一致')
        self.assertEqual(rep['per_field'][('图片', 0)]['drop'], 1)

    def test_diff_and_fatal(self):
        hdr = ['名称规格', '市场价']
        rep, details = self.run_one(hdr, ['班级信息发布终端', 10], ['光模块', 10])
        self.assertEqual(details[0][3], 'diff', '整体换成另一个物料名必须判为内容冲突')
        self.assertTrue(R.verdict(rep), '名称内容冲突必须报致命')

    def test_price_change_is_fatal(self):
        hdr = ['名称规格', '市场价']
        rep, details = self.run_one(hdr, ['甲', 10], ['甲', 12])
        self.assertEqual(details[0][3], 'diff')
        self.assertTrue(any('市场价' in f for f in R.verdict(rep)), '价格被改必须报致命')


if __name__ == '__main__':
    unittest.main()
