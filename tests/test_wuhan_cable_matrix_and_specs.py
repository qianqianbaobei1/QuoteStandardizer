# -*- coding: utf-8 -*-
"""
测试武汉第二电线电缆全系列与复杂矩阵表属性解构、多区子表头及分系列折扣：
1. 复合规格维度解构：N*S 芯数/截面，L*W*T 尺寸，纯数字线缆截面与额定电压
2. 矩阵表内局部子表头重绑定 (Multi-Zone Matrix)
3. 斜杠型号完整提取 (WDZC/B-BYJ, JZ-BV(7芯))
4. 分系列差异化工程折扣 (电缆 0.85, 电线 0.88)
"""
import os
import sys
import unittest

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import openpyxl
from main import extract_deterministic_kv, _collect_matrix_rows, _detect_matrix
from local_terms_engine import extract_terms_from_ocr_text, parse_discount_num


class TestWuhanCableAndMatrixSpecs(unittest.TestCase):

    def test_decompose_cable_and_board_dimensions(self):
        """测试规格维度确定性拆解"""
        # 1. 电线多芯: 2*2.5
        kv1 = extract_deterministic_kv("规格: 2*2.5; 型号: BVVB", {"品名": "护套线"})
        self.assertEqual(kv1.get("芯数"), "2")
        self.assertEqual(kv1.get("标称截面(mm²)"), "2.5")
        self.assertNotIn("规格型号", kv1)

        # 2. 板材尺寸: 600*600*15
        kv2 = extract_deterministic_kv("花岗岩板材 600*600*15mm 荔枝面", {"品名": "花岗岩"})
        self.assertEqual(kv2.get("长度(mm)"), "600")
        self.assertEqual(kv2.get("宽度(mm)"), "600")
        self.assertEqual(kv2.get("厚度(mm)"), "15")

        # 3. 单芯电线截面与电压: 1.5, 10kV
        kv3 = extract_deterministic_kv("规格: 1.5; 列: JKLYJ-10kV架空", {"品名": "架空绝缘电线"})
        self.assertEqual(kv3.get("标称截面(mm²)"), "1.5")
        self.assertEqual(kv3.get("芯数"), "1")
        self.assertEqual(kv3.get("额定电压"), "10kV")

        # 4. 7芯电线: JZ-BV(7芯)
        kv4 = extract_deterministic_kv("规格: 2.5; 列: JZ-BV(7芯)", {"品名": "电线"})
        self.assertEqual(kv4.get("标称截面(mm²)"), "2.5")
        self.assertEqual(kv4.get("芯数"), "7")

    def test_multi_zone_matrix_and_slash_models(self):
        """测试矩阵表中嵌入局部子表头 (如右侧 BLVVB/BVVB 护套线) 与带斜杠型号提取"""
        wb = openpyxl.Workbook()
        ws = wb.active
        # R1: 大标题
        ws.append(['电线电缆全系列价格表'])
        # R2~R3: 空行
        ws.append([])
        ws.append([])
        # R4: 顶层型号表头
        ws.append(['截面', 'BV', 'BVR', 'WDZC/B-BYJ', 'JKLYJ', ''])
        # R5: 二级表头
        ws.append(['规格', '铜芯软', '铜芯硬', '低烟无卤', '1KV架空', '10kV架空'])
        # R6~R10: 正常数据行
        ws.append(['1.5', 100, 200, 300, 400, 500])
        ws.append(['2.5', 150, 250, 350, 450, 550])
        ws.append(['4', 200, 300, 400, 500, 600])
        # R9: 右侧嵌入护套线局部子表头！
        # Col 1~4 仍是主表数据，Col 5 为型号，Col 6 为 BLVVB
        ws.append(['6', 300, 400, 500, '型号', 'BLVVB'])
        # R10: 局部二级表头：Col 5 规格，Col 6 铝芯护套
        ws.append(['10', 400, 500, 600, '规格', '铝芯护套'])
        # R11: 护套线子表开始数据
        ws.append(['16', 500, 600, 700, '2*1.5', 1044])

        mx = _detect_matrix(ws)
        self.assertIsNotNone(mx)
        raw_rows, records = _collect_matrix_rows(ws, 'Sheet1', mx, {})
        
        # 验证带斜杠的型号 WDZC/B-BYJ
        wdzc_items = [r for r in raw_rows if 'WDZC/B-BYJ' in r['raw_spec']]
        self.assertTrue(len(wdzc_items) > 0, "必须提取出带斜杠的完整型号 WDZC/B-BYJ")
        self.assertEqual(wdzc_items[0]['raw_cells_dict'].get('型号/系列'), 'WDZC/B-BYJ')

        # 验证右侧嵌入的 BLVVB 护套线，其规格正确认领了第 5 列的 2*1.5，价格为 1044
        blvvb_items = [r for r in raw_rows if 'BLVVB' in r['raw_spec'] or '铝芯护套' in r['raw_spec']]
        self.assertTrue(len(blvvb_items) > 0, "右侧嵌入的 BLVVB 护套线子表必须被正确识别")
        b_item = blvvb_items[0]
        self.assertEqual(b_item['raw_cells_dict']['规格型号'], '2*1.5')
        self.assertEqual(b_item['raw_price'], 1044.0)

    def test_series_based_discount_and_terms(self):
        """测试从报价说明中提取分系列差异化折扣与适用地区"""
        ocr_text = """
武汉第二电线电缆有限公司
报价说明与工程折扣单
适用范围：湖北省含运费
工程折扣优惠条款：
所有电缆系列，价格表下浮15%，起订量5万；
所有电线系列，价格表下浮12%，起订量1万；
付款方式：款到发货，质保1年。
"""
        terms = extract_terms_from_ocr_text(ocr_text)
        self.assertEqual(terms.get('province'), '湖北省')
        self.assertEqual(terms.get('city'), '武汉市')
        self.assertEqual(terms.get('is_freight_included'), '包含')

        # 检查分系列折扣
        s_map = terms.get('discount_series_map', {})
        self.assertEqual(s_map.get('电缆'), 0.85)
        self.assertEqual(s_map.get('电线'), 0.88)


if __name__ == '__main__':
    unittest.main()
