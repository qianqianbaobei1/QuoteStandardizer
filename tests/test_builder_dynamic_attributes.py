# -*- coding: utf-8 -*-
"""
tests/test_builder_dynamic_attributes.py
验证动态多列拆分构建器：
根据广材三级品类标准Schema，将物料属性在“报价时间”后动态拆解为“一属性一列”，
红底白字高亮，单元格精准对位填入，图片列与时间列严密就位。
"""
import os
import shutil
import tempfile
import unittest
import openpyxl

from builder import StandardizedWorkbookBuilder
from models import CanonicalMaterialItem


class TestBuilderDynamicAttributes(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp(prefix="qs_test_dynamic_")

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_dynamic_attributes_expansion_for_round_steel(self):
        """测试圆钢物料：自动拆分 品种、直径(mm)、牌号、产品描述 为独立列"""
        item = CanonicalMaterialItem(
            quote_id="YG00010",
            supplier="测试钢铁贸易有限公司",
            brand="友发",
            category_name="热轧圆钢",
            category_code="010901",
            structured_specs={
                "品种": "热轧圆钢",
                "直径(mm)": "25",
                "牌号": "Q235B",
                "产品描述": "标准直条"
            },
            market_price=3850.0,
            unit="t",
            project_type="市政基建",
            project_volume="500吨",
            quote_date="2026-10-09"
        )

        out_path = os.path.join(self.tmp_dir, "round_steel_output.xlsx")
        img_dir = os.path.join(self.tmp_dir, "imgs")

        builder = StandardizedWorkbookBuilder(None, auto_expand_attributes=True)
        builder.build_from_canonical_items([item], out_path, img_dir)

        self.assertTrue(os.path.exists(out_path))
        wb = openpyxl.load_workbook(out_path)
        self.assertIn("标准化报价", wb.sheetnames)
        ws = wb["标准化报价"]

        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

        # 1. 验证 报价时间 锚点位置存在
        self.assertIn("报价时间", headers)
        time_col = headers.index("报价时间") + 1

        # 2. 验证动态属性独立成列且排在 报价时间 之后
        self.assertIn("品种", headers)
        self.assertIn("直径(mm)", headers)
        self.assertIn("牌号", headers)

        var_col = headers.index("品种") + 1
        diam_col = headers.index("直径(mm)") + 1
        grade_col = headers.index("牌号") + 1

        self.assertGreater(var_col, time_col)
        self.assertGreater(diam_col, time_col)
        self.assertGreater(grade_col, time_col)

        # 3. 验证图片文件夹紧随动态属性列之后
        img_folder_col = headers.index("图片文件夹") + 1
        self.assertGreater(img_folder_col, grade_col)

        # 4. 验证单元格数据精准填入对应列
        self.assertEqual(ws.cell(2, time_col).value, "2026-10-09")
        self.assertEqual(ws.cell(2, var_col).value, "热轧圆钢")
        self.assertEqual(ws.cell(2, diam_col).value, "25")
        self.assertEqual(ws.cell(2, grade_col).value, "Q235B")

        # 5. 验证表头视觉高亮（红框标杆设计：属性列橙红底白字高亮，报价时间浅绿底）
        time_cell = ws.cell(1, time_col)
        self.assertEqual(time_cell.fill.start_color.rgb, "00E2EFDA")

        attr_cell = ws.cell(1, diam_col)
        self.assertEqual(attr_cell.fill.start_color.rgb, "00C00000")
        self.assertEqual(attr_cell.font.color.rgb, "00FFFFFF")
        self.assertTrue(attr_cell.font.bold)

    def test_dynamic_attributes_expansion_for_stone(self):
        """测试石材物料：自动拆分 长度(mm)、宽度(mm)、厚度(mm)、表面处理、产地等独立列"""
        item = CanonicalMaterialItem(
            quote_id="SC00010",
            supplier="四川汇才石业有限公司",
            brand="汇才",
            category_name="雅士白大理石",
            category_code="060101",
            structured_specs={
                "品种": "雅士白大理石",
                "厚度(mm)": "18",
                "长度(mm)": "800",
                "宽度(mm)": "800",
                "表面处理": "抛光面",
                "产地": "进口"
            },
            market_price=360.0,
            unit="m²",
            quote_date="2026-10-15"
        )

        out_path = os.path.join(self.tmp_dir, "stone_output.xlsx")
        img_dir = os.path.join(self.tmp_dir, "imgs")

        builder = StandardizedWorkbookBuilder(None, auto_expand_attributes=True)
        builder.build_from_canonical_items([item], out_path, img_dir)

        wb = openpyxl.load_workbook(out_path)
        ws = wb["标准化报价"]
        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

        for expected_attr in ["品种", "厚度(mm)", "长度(mm)", "宽度(mm)", "表面处理", "产地"]:
            self.assertIn(expected_attr, headers)
            col_idx = headers.index(expected_attr) + 1
            val = ws.cell(2, col_idx).value
            self.assertEqual(val, item.structured_specs[expected_attr])

    def test_split_by_category_multiple_sheets(self):
        """测试多品类物料：自动拆分独立分类 Sheet 保存，每个 Sheet 专属展开各自的标准属性列"""
        item_steel = CanonicalMaterialItem(
            quote_id="YG00010",
            supplier="综合建材商",
            brand="友发",
            category_name="热轧圆钢",
            category_code="010901",
            structured_specs={"品种": "热轧圆钢", "直径(mm)": "25", "牌号": "Q235B"},
            market_price=3850.0,
            unit="t"
        )
        item_marble = CanonicalMaterialItem(
            quote_id="SC00010",
            supplier="综合建材商",
            brand="汇才",
            category_name="天然大理石板材",
            category_code="A080105",
            structured_specs={"品种": "金线米黄", "厚度(mm)": "20", "长度(mm)": "800", "宽度(mm)": "800"},
            market_price=180.0,
            unit="m²"
        )

        out_path = os.path.join(self.tmp_dir, "multi_cat_output.xlsx")
        img_dir = os.path.join(self.tmp_dir, "imgs")

        builder = StandardizedWorkbookBuilder(None, auto_expand_attributes=True)
        builder.build_from_canonical_items([item_steel, item_marble], out_path, img_dir)

        wb = openpyxl.load_workbook(out_path)
        # 1. 验证 Sheet 集合：既包含主表与原表，又包含各个独立的分类 Sheet
        self.assertIn("标准化报价", wb.sheetnames)
        self.assertIn("客户原始解析", wb.sheetnames)
        self.assertIn("010901", wb.sheetnames)
        self.assertIn("A080105", wb.sheetnames)

        # 2. 验证圆钢分类 Sheet (010901)：专属包含直径、牌号，不含大理石厚度长宽
        ws_steel = wb["010901"]
        steel_headers = [ws_steel.cell(1, c).value for c in range(1, ws_steel.max_column + 1)]
        self.assertIn("直径(mm)", steel_headers)
        self.assertIn("牌号", steel_headers)
        self.assertNotIn("厚度(mm)", steel_headers)
        self.assertEqual(ws_steel.max_row, 2)  # 1行表头 + 1行圆钢数据

        # 3. 验证大理石分类 Sheet (A080105)：专属包含厚度、长宽，不含圆钢直径与牌号
        ws_marble = wb["A080105"]
        marble_headers = [ws_marble.cell(1, c).value for c in range(1, ws_marble.max_column + 1)]
        self.assertIn("厚度(mm)", marble_headers)
        self.assertIn("长度(mm)", marble_headers)
        self.assertNotIn("直径(mm)", marble_headers)
        self.assertNotIn("牌号", marble_headers)
        self.assertEqual(ws_marble.max_row, 2)  # 1行表头 + 1行大理石数据


if __name__ == "__main__":
    unittest.main()
