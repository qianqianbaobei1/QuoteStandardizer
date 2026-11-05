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
from main import extract_deterministic_kv, _collect_matrix_rows, _detect_matrix, _match_discount_rule
from local_terms_engine import extract_terms_from_ocr_text, parse_discount_num
from category_attribute_service import fetch_category_attributes


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

    def test_fallback_category_attributes(self):
        """测试国标大类通用属性兜底模板"""
        cable_attrs = fetch_category_attributes("电力电缆")
        cable_names = [a["name"] for a in cable_attrs]
        self.assertIn("标称截面(mm²)", cable_names)
        self.assertIn("芯数", cable_names)

        pipe_attrs = fetch_category_attributes("热镀锌钢管")
        pipe_names = [a["name"] for a in pipe_attrs]
        self.assertIn("公称直径DN(mm)", pipe_names)
        self.assertIn("壁厚(mm)", pipe_names)

        board_attrs = fetch_category_attributes("铝单板")
        board_names = [a["name"] for a in board_attrs]
        self.assertIn("厚度(mm)", board_names)

    def test_match_discount_rule_cable_vs_wire(self):
        """测试分系列折扣匹配器区分电缆(0.85)与电线(0.88)"""
        disc_map = {"所有电缆系列": 0.85, "所有电线系列": 0.88}

        # 1. 架空电缆
        r1, s1 = _match_discount_rule(
            item_context="JKLYJ 10kV架空 50",
            l2="电力电缆",
            l3="架空绝缘电缆",
            std_name="JKLYJ",
            raw_full_line="JKLYJ 10kV架空",
            disc_map=disc_map,
            default_ratio=0.9,
            default_series="普通折扣",
        )
        self.assertEqual(r1, 0.85)
        self.assertIn("电缆系列", s1)

        # 2. 布电线 BV
        r2, s2 = _match_discount_rule(
            item_context="BV 2.5",
            l2="电线电缆",
            l3="铜芯聚氯乙烯绝缘电线",
            std_name="BV",
            raw_full_line="BV 2.5",
            disc_map=disc_map,
            default_ratio=0.9,
            default_series="普通折扣",
        )
        self.assertEqual(r2, 0.88)
        self.assertIn("电线系列", s2)

        # 3. 护套线 BLVVB
        r3, s3 = _match_discount_rule(
            item_context="BLVVB 铝芯护套 2*1.5",
            l2="布电线",
            l3="铝芯聚氯乙烯护套线",
            std_name="BLVVB",
            raw_full_line="BLVVB 2*1.5",
            disc_map=disc_map,
            default_ratio=0.9,
            default_series="普通折扣",
        )
        self.assertEqual(r3, 0.88)
        self.assertIn("电线系列", s3)

    def test_bottom_clustered_tables_and_bv_7strand(self):
        """测试底部并排小表：单行表头不吞下一行价格、以及未写规格二字的独立小表 BV（7股）正确切分"""
        wb = openpyxl.Workbook()
        ws = wb.active
        # R1: 顶层表头
        ws.append(['截面', 'BV', 'BVR', 'WDZC', 'JKLYJ', 'JKLV', 'YJV'])
        # R2~R5: 主表数据行
        ws.append(['1.5', 100, 200, 300, 400, 500, 600])
        ws.append(['2.5', 150, 250, 350, 450, 550, 650])
        ws.append(['4', 200, 300, 400, 500, 600, 700])
        ws.append(['6', 300, 400, 500, 600, 700, 800])
        # R6: 空行
        ws.append([])
        # R7: 底部横幅大标题行
        ws.append(['音响线', '', '网线', '', '', 'BV（7股）'])
        # R8: 局部子表头行（单行表头）
        ws.append(['规格', '单价', '规格', '单价', '', 'BV（7股）', '价格'])
        # R9: 第一行数据行（价格 1814 绝不能被吞进列名！）
        ws.append(['2×0.5', 1814, '超五类', 2419, '', '2.5', 2563])
        # R10: 第二行数据
        ws.append(['2×0.75', 2620, '六类', 3150, '', '4', 4069])

        mx = _detect_matrix(ws)
        self.assertIsNotNone(mx)
        raw_rows, records = _collect_matrix_rows(ws, 'Sheet1', mx, {})

        # 1. 验证价格 1814 没有被拼进列标题
        audio_items = [r for r in raw_rows if '音响线' in r['raw_spec'] or '2×0.5' in r['raw_spec']]
        self.assertTrue(len(audio_items) > 0)
        self.assertNotIn('1814', audio_items[0]['raw_cells_dict'].get('列标题', ''), "单行表头绝不能把第一行数据的价格吞进列名")

        # 2. 验证未写规格二字的独立小表 BV（7股）被正确提取，且规格为 2.5，单价为 2563
        bv7_items = [r for r in raw_rows if 'BV（7股）' in r['raw_spec'] or '7股' in r['raw_full_line']]
        self.assertTrue(len(bv7_items) >= 2, "未写'规格'二字的 BV（7股）独立小表必须被正确切分识别")
        self.assertEqual(bv7_items[0]['raw_cells_dict']['规格型号'], '2.5')
        self.assertEqual(bv7_items[0]['raw_price'], 2563.0)


if __name__ == '__main__':
    unittest.main()

