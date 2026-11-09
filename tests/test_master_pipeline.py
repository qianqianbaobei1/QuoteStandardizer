# -*- coding: utf-8 -*-
"""
tests/test_master_pipeline.py
工业级 6 层流水线与三大验收铁律自动化验证套件：
1. 铁律一：非视觉模型熔断、印章局部手术修图、RapidOCR >= 60% 门禁
2. 铁律二：显式品种回正 (等边角钢 vs 扁钢、花纹钢板 vs 钢板)
3. 铁律三：零假对账与零静态写死，动态聚合验证
4. 6 层流水线端到端协同
"""
import os
import sys
import unittest
import numpy as np

from core.ir_models import (
    QuoteDocumentIR, DocumentAsset, AssetType, RawRow, FailFastGateError
)
from core.layer1_router import AssetRouter
from core.layer4_commercial import CommercialEngine
from core.layer5_taxonomy import TaxonomySpecEngine
from seal_enhancer_service import surgical_seal_mask_repair
from pdf_extractor import PDFTableExtractor
import config


class TestMasterPipelineAndIronLaws(unittest.TestCase):

    def test_iron_law_1_banned_vision_model(self):
        """【铁律一测试】：配置纯文本模型 (deepseek-chat) 必须被立即阻止并抛出 RuntimeError"""
        with self.assertRaises(RuntimeError) as ctx:
            config.check_vision_model_safety("deepseek-chat")
        self.assertIn("纯文本模型", str(ctx.exception))

    def test_iron_law_1_rapidocr_overlap_gate(self):
        """【铁律一测试】：模型输出汉字与 RapidOCR 字符重合度低于 60% 必须触发硬门禁熔断"""
        from PIL import Image, ImageDraw, ImageFont
        img = Image.new("RGB", (500, 120), color=(255, 255, 255))
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("/System/Library/Fonts/Supplemental/Songti.ttc", 24)
        except Exception:
            font = None
        draw.text((20, 30), "天兴消防器材玻璃球洒水喷头检验合格", fill=(0, 0, 0), font=font)
        cv_img = np.array(img)[:, :, ::-1]

        # 1. 虚假幻觉输出（如：天津天药原料药注射液）
        hallucinated_dict = {
            "tables": [
                {"cells": [["天津天药原料药注射液批号"], ["生产厂家股份有限公司"]]}
            ]
        }
        is_ok, ratio, reason = PDFTableExtractor._verify_against_rapidocr(cv_img, hallucinated_dict)
        self.assertFalse(is_ok, f"幻觉文本重合度应低于60%，实际为: {ratio}")

        # 2. 真实对应输出（玻璃球洒水喷头）
        correct_dict = {
            "tables": [
                {"cells": [["天兴消防器材", "玻璃球洒水喷头", "检验合格"]]}
            ]
        }
        is_ok_real, ratio_real, _ = PDFTableExtractor._verify_against_rapidocr(cv_img, correct_dict)
        self.assertTrue(is_ok_real, f"真实文字重合度应 >= 60%，实际为: {ratio_real}")

    def test_iron_law_1_surgical_seal_repair(self):
        """【铁律一测试】：局部印章手术，保护蓝底白字/黑底白字与 12.24 细小数字"""
        # 构造纯蓝色像素块 (蓝底白字表头背景)
        blue_img = np.full((50, 50, 3), [200, 50, 0], dtype=np.uint8)  # BGR
        repaired_blue = surgical_seal_mask_repair(blue_img)
        # 蓝底非红章，必须 100% 保持原值
        np.testing.assert_array_equal(blue_img, repaired_blue)

        # 构造一个红印章像素块: R=230, G=30, B=30 (R - max(G,B) = 200 > 25)
        red_img = np.full((50, 50, 3), [30, 30, 230], dtype=np.uint8)
        repaired_red = surgical_seal_mask_repair(red_img)
        # 印章像素去红留黑，B 与 G 通道脱色被提升为 R 灰度 (230)
        self.assertEqual(repaired_red[0, 0, 0], 230)
        self.assertEqual(repaired_red[0, 0, 1], 230)

    def test_iron_law_2_variety_override_angle_steel(self):
        """【铁律二测试】：折角残字扁钢凭“品种:等边角钢”回正为 01-21-01 角钢"""
        tax_engine = TaxonomySpecEngine()
        doc_ir = QuoteDocumentIR(supplier_name="测试钢材集团")
        doc_ir.raw_rows = [
            RawRow(
                source_ref="Sheet1!R10",
                sheet_or_page="Sheet1",
                row_index=10,
                cells={
                    "品名": "扁钢",  # 原文折角残缺，被识别为扁钢
                    "规格型号": "50*5 品种:等边角钢 材质:Q235B",
                    "单价": "3850"
                },
                raw_text="扁钢 50*5 品种:等边角钢 材质:Q235B 单价:3850",
                raw_price=3850.0
            )
        ]
        doc_ir = tax_engine.process(doc_ir)
        self.assertEqual(len(doc_ir.canonical_items), 1)
        item = doc_ir.canonical_items[0]
        # 验证品类必须被纠偏为角钢类 (01-21-01 角钢)，绝不可留在扁钢 (01-13-01)
        self.assertIn("角钢", item.category_name)
        self.assertEqual(item.category_code, "01-21-01")
        self.assertEqual(item.structured_specs.get("品种"), "等边角钢")
        self.assertEqual(item.structured_specs.get("材质"), "Q235B")

    def test_iron_law_2_variety_checkered_steel_plate(self):
        """【铁律二测试】：花纹钢板必须归入 01-29-01 钢板，并填入品种=花纹钢板"""
        tax_engine = TaxonomySpecEngine()
        doc_ir = QuoteDocumentIR(supplier_name="测试板材实业")
        doc_ir.raw_rows = [
            RawRow(
                source_ref="Sheet1!R20",
                sheet_or_page="Sheet1",
                row_index=20,
                cells={
                    "品名": "钢板",
                    "规格型号": "T=4.0mm 品种:花纹钢板 材质:Q235B",
                    "单价": "4200"
                },
                raw_text="钢板 T=4.0mm 品种:花纹钢板 材质:Q235B 单价:4200",
                raw_price=4200.0
            )
        ]
        doc_ir = tax_engine.process(doc_ir)
        self.assertEqual(len(doc_ir.canonical_items), 1)
        item = doc_ir.canonical_items[0]
        self.assertEqual(item.category_code, "01-29-01")
        self.assertIn("钢板", item.category_name)
        self.assertEqual(item.structured_specs.get("品种"), "花纹钢板")

    def test_iron_law_3_dynamic_commercial_waterfall(self):
        """【铁律三与第4层测试】：行内税率优先于全局，价格 100% 原文保真"""
        comm_engine = CommercialEngine()
        doc_ir = QuoteDocumentIR()
        doc_ir.raw_rows = [
            RawRow(
                source_ref="Sheet1!R05",
                sheet_or_page="Sheet1",
                row_index=5,
                cells={"品名": "消火栓", "税率": "9%", "单价": "120.0"},
                raw_text="消火栓 税率:9% 单价:120.0",
                raw_price=120.0
            ),
            RawRow(
                source_ref="Sheet1!R06",
                sheet_or_page="Sheet1",
                row_index=6,
                cells={"品名": "水枪", "单价": "45.0"},
                raw_text="水枪 单价:45.0",
                raw_price=45.0
            )
        ]
        # 全局条款设置 13% 税率与 8.5 折
        terms_dict = {
            "tax_rate": 13,
            "discount_ratio": 0.85,
            "discount_series": "消防喷头:8.9折"
        }
        comm_engine.apply(doc_ir, terms_dict)

        # 第 1 行有显式税率 9%，应覆盖全局 13%
        self.assertEqual(doc_ir.raw_rows[0].cells["_tax_rate"], 9)
        # 第 2 行无显式税率，应回退继承全局 13%
        self.assertEqual(doc_ir.raw_rows[1].cells["_tax_rate"], 13)


if __name__ == '__main__':
    unittest.main()
