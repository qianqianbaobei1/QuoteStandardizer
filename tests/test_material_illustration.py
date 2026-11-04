# -*- coding: utf-8 -*-
"""
单元测试：材料配图接口接入与图片链接匹配
- 验证从 工料机类别.xlsx Sheet3「材料配图接口」规范接入的配图服务
- 测试检索材料、任务轮询、批量匹配、本地缓存以及在 Excel 中的超链接匹配落盘
"""
import os
import unittest
import openpyxl
from unittest.mock import patch, MagicMock

from models import CanonicalMaterialItem
from builder import StandardizedWorkbookBuilder, STANDARD_HEADERS
from material_illustration_service import (
    MaterialIllustrationService,
    match_illustrations_for_canonical_items,
    match_illustrations_for_openpyxl_sheet,
    clean_material_name
)


class TestMaterialIllustrationService(unittest.TestCase):
    def setUp(self):
        self.output_dir = os.path.join(os.path.dirname(__file__), "scratch_test_output")
        os.makedirs(self.output_dir, exist_ok=True)
        # conftest 默认关闭配图接口；本测试类验证的就是配图链路，这里显式打开 (接口调用全部 mock)
        patcher = patch("material_illustration_service.ENABLE_MATERIAL_ILLUSTRATION", True)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_clean_material_name(self):
        self.assertEqual(clean_material_name("  镀锌钢管 \n "), "镀锌钢管")
        self.assertEqual(clean_material_name("UPVC管\t"), "UPVC管")
        self.assertEqual(clean_material_name(""), "")

    def test_fetch_illustrations_mocked_success(self):
        """测试配图服务创建任务与轮询结果的完整链路"""
        service = MaterialIllustrationService(enable=True)

        mock_create_resp = MagicMock()
        mock_create_resp.read.return_value = b'{"status": 200, "message": "success", "data": {"taskId": "mock-task-123"}}'
        mock_create_resp.__enter__.return_value = mock_create_resp

        mock_query_resp = MagicMock()
        mock_query_resp.read.return_value = b'''{
            "status": 200,
            "message": "success",
            "data": {
                "taskId": "mock-task-123",
                "status": "SUCCEEDED",
                "pictures": [
                    {
                        "uid": "item_1",
                        "pictureName": "\xe9\x95\x80\xe9\x94\x8c\xe9\x92\xa2\xe7\xae\xa1.jpg",
                        "pictureUrl": "https://example.com/dxgg.jpg",
                        "folderName": "m0000578"
                    }
                ]
            }
        }'''
        mock_query_resp.__enter__.return_value = mock_query_resp

        with patch("urllib.request.urlopen") as mock_urlopen, \
             patch.dict("material_illustration_service._ILLUSTRATION_CACHE", {}, clear=True), \
             patch("material_illustration_service.save_cache"):
            mock_urlopen.side_effect = [mock_create_resp, mock_query_resp]
            mats = [{"uid": "item_1", "materialName": "镀锌钢管测试料"}]
            res = service.fetch_illustrations_batch(mats, company_name="测试公司")

            self.assertIn("item_1", res)
            self.assertEqual(res["item_1"]["pictureUrl"], "https://example.com/dxgg.jpg")
            self.assertEqual(res["item_1"]["pictureName"], "镀锌钢管.jpg")
            self.assertEqual(res["item_1"]["folderName"], "m0000578")

    def test_canonical_items_matching_and_builder_excel(self):
        """测试 CanonicalMaterialItem 匹配到图片链接后，在构建器生成的 Excel 中正确写入超链接与元数据"""
        item1 = CanonicalMaterialItem(
            quote_id="TEST0001",
            supplier="测试友发钢管公司",
            brand="友发",
            category_name="镀锌钢管",
            category_code="184701",
            market_price=150.0,
            unit="m",
            raw_spec="DN100"
        )
        item2 = CanonicalMaterialItem(
            quote_id="TEST0002",
            supplier="测试友发钢管公司",
            brand="友发",
            category_name="未知无配图材料999",
            category_code="000101",
            market_price=50.0,
            unit="个"
        )

        hit = {
            "matched": True,
            "pictureUrl": "https://gcj-pub.oss-cn-beijing.aliyuncs.com/opcrm/pic_pictures/m0000578/1cd97a16dbbdc896.jpg",
            "pictureName": "镀锌钢管.jpg",
            "folderName": "m0000578"
        }

        # 实现按材料名去重后查询 (uid 为 uniq_N)，mock 按材料名命中，不依赖 uid 编号
        def fake_fetch(materials, **kwargs):
            return {m["uid"]: hit for m in materials if m["materialName"] == "镀锌钢管"}

        with patch("material_illustration_service.default_illustration_service.fetch_illustrations_batch", side_effect=fake_fetch):
            matched = match_illustrations_for_canonical_items([item1, item2], company_name="测试友发钢管公司")
            self.assertEqual(matched, 1)
            self.assertEqual(item1.picture_url, "https://gcj-pub.oss-cn-beijing.aliyuncs.com/opcrm/pic_pictures/m0000578/1cd97a16dbbdc896.jpg")
            self.assertEqual(item1.illustration_picture_name, "镀锌钢管.jpg")
            self.assertEqual(item1.illustration_folder, "m0000578")
            self.assertEqual(item2.picture_url, "")

        # 验证构建生成的 Excel 文件
        builder = StandardizedWorkbookBuilder(None)
        xlsx_path = os.path.join(self.output_dir, "test_illustration_delivery.xlsx")
        builder.build_from_canonical_items([item1, item2], xlsx_path, img_dir=os.path.join(self.output_dir, "imgs"))

        self.assertTrue(os.path.exists(xlsx_path))
        wb = openpyxl.load_workbook(xlsx_path)
        ws = wb["标准化报价"]

        headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        img_col = headers.index("图片") + 1
        img_file_col = headers.index("厂家报价图片文件名") + 1
        folder_col = headers.index("图片文件夹") + 1

        # 第 2 行 (item1 命中配图)
        cell_img1 = ws.cell(2, img_col)
        self.assertEqual(cell_img1.value, "https://gcj-pub.oss-cn-beijing.aliyuncs.com/opcrm/pic_pictures/m0000578/1cd97a16dbbdc896.jpg")
        self.assertIsNotNone(cell_img1.hyperlink)
        self.assertEqual(cell_img1.hyperlink.target, "https://gcj-pub.oss-cn-beijing.aliyuncs.com/opcrm/pic_pictures/m0000578/1cd97a16dbbdc896.jpg")
        self.assertIn(ws.cell(2, img_file_col).value, ("", None))
        self.assertEqual(ws.cell(2, folder_col).value, "m0000578")

        # 第 3 行 (item2 未命中配图)
        cell_img2 = ws.cell(3, img_col)
        self.assertIn(cell_img2.value, ("", None))

    def test_match_illustrations_for_openpyxl_sheet(self):
        """测试直接对 openpyxl Worksheet 执行配图链接回填"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(STANDARD_HEADERS)
        ws.append(["ID1", "测试商", "UPVC排水管", "184701", "联塑", "", "", "", "", 10, "广东省", "广州市", "m", 25.0])

        mock_hit = {
            "r_2": {
                "matched": True,
                "pictureUrl": "https://example.com/upvc.jpg",
                "pictureName": "UPVC排水管.jpg",
                "folderName": "m0002102"
            }
        }
        with patch("material_illustration_service.default_illustration_service.fetch_illustrations_batch", return_value=mock_hit):
            count = match_illustrations_for_openpyxl_sheet(ws, company_name="测试商")
            self.assertEqual(count, 1)

            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            img_c = headers.index("图片") + 1
            file_c = headers.index("厂家报价图片文件名") + 1
            folder_c = headers.index("图片文件夹") + 1

            self.assertEqual(ws.cell(2, img_c).value, "https://example.com/upvc.jpg")
            self.assertIsNotNone(ws.cell(2, img_c).hyperlink)
            self.assertIn(ws.cell(2, file_c).value, ("", None))
            self.assertEqual(ws.cell(2, folder_c).value, "m0002102")


if __name__ == '__main__':
    unittest.main()
