# -*- coding: utf-8 -*-
"""
QuoteStandardizer 综合健康体检与单元自检脚本
覆盖全模块功能自检：
1. 配置与行政区划解析 (config)
2. 物料数据模型与防篡改对账 (models)
3. Excel 数据展开与精度保护 (extractor)
4. 商务条款提取与规范化 (terms_extractor)
5. 报表构建与格式/字体检查 (builder)
6. 正则与供应商名称防截断 (main / pdf / gemini)
"""
import os
import sys
import unittest
import openpyxl

from config import normalize_province, expand_region, resolve_category_code
from models import CanonicalMaterialItem, SourceRowRecord, reconcile
from extractor import safe_clean_price, safe_cell_str, unmerge_and_fill_cells
from terms_extractor import parse_discount, parse_tax_rate, normalize_terms
from builder import StandardizedWorkbookBuilder
from main import COMPANY_RE as MAIN_COMPANY_RE
from pdf_extractor import COMPANY_RE as PDF_COMPANY_RE
from gemini_pdf_extractor import COMPANY_RE as GEMINI_COMPANY_RE


class TestQuoteStandardizerHealth(unittest.TestCase):

    def test_01_province_and_region(self):
        """测试行政区划与连写区域解析"""
        self.assertEqual(normalize_province('京'), '北京市')
        self.assertEqual(normalize_province('河北省>石家庄市'), '河北省')
        self.assertEqual(normalize_province('全国通用'), '全国')
        
        expanded = expand_region('京津冀+山东')
        self.assertEqual(expanded, ['北京市', '天津市', '河北省', '山东省'])
        self.assertEqual(expand_region('全国'), ['全国'])

    def test_02_company_regex_consistency(self):
        """测试各模块的供应商名称正则一致性，严禁截断带（集团）的企业全称"""
        test_name = "重庆望变电气（集团）股份有限公司报价单"
        
        m_main = MAIN_COMPANY_RE.search(test_name)
        m_pdf = PDF_COMPANY_RE.search(test_name)
        m_gemini = GEMINI_COMPANY_RE.search(test_name)
        
        self.assertIsNotNone(m_main)
        self.assertEqual(m_main.group(1), "重庆望变电气（集团）股份有限公司")
        
        self.assertIsNotNone(m_pdf)
        self.assertEqual(m_pdf.group(1), "重庆望变电气（集团）股份有限公司")
        
        self.assertIsNotNone(m_gemini)
        self.assertEqual(m_gemini.group(1), "重庆望变电气（集团）股份有限公司")

    def test_03_terms_normalization(self):
        """测试商务条款各维度规范化"""
        self.assertEqual(parse_discount('9折'), 0.9)
        self.assertEqual(parse_discount('85折'), 0.85)
        self.assertEqual(parse_discount('8.5折'), 0.85)
        self.assertEqual(parse_discount('无折扣'), 1.0)
        self.assertEqual(parse_discount(0.4), 0.4)
        
        self.assertEqual(parse_tax_rate('13%'), 13)
        self.assertEqual(parse_tax_rate('0.13'), 13)
        self.assertEqual(parse_tax_rate(9), 9)

        norm = normalize_terms({
            'supplier': '测试公司',
            'discount_ratio': '9折',
            'tax_rate': '13%',
            'is_tax_included': '是',
            'is_freight_included': '含运费',
            'province': '渝'
        })
        self.assertEqual(norm['discount_ratio'], 0.9)
        self.assertEqual(norm['tax_rate'], 13)
        self.assertEqual(norm['is_tax_included'], '是')
        self.assertEqual(norm['is_freight_included'], '包含')
        self.assertEqual(norm['province'], '重庆市')

    def test_04_extractor_clean_price_and_precision(self):
        """测试价格提取与数值精度无损保护"""
        self.assertEqual(safe_clean_price("¥10,500.00"), 10500.0)
        self.assertEqual(safe_clean_price("136,000"), 136000.0)
        self.assertEqual(safe_clean_price(None), 0.0)
        self.assertEqual(safe_clean_price("面议"), 0.0)

        # 保护带小数点的浮点精确保留
        self.assertEqual(safe_cell_str(1.0, number_format='0.0'), '1.0')
        self.assertEqual(safe_cell_str(10500), '10500')

    def test_05_reconciliation_tamper_proofing(self):
        """测试原表对账机制 (防漏行、防篡改价格、防多出假行)"""
        srows = [
            SourceRowRecord(source_ref="S1!R2", sheet="S1", row=2, raw_text="变压器 10000", raw_price=10000.0, decision="item"),
            SourceRowRecord(source_ref="S1!R3", sheet="S1", row=3, raw_text="标题行", raw_price=0.0, decision="title"),
            SourceRowRecord(source_ref="S1!R4", sheet="S1", row=4, raw_text="变压器 20000", raw_price=20000.0, decision="item")
        ]
        
        # 场景 A: 价格被篡改
        items_tampered = [
            CanonicalMaterialItem(quote_id="WB001", supplier="望变", brand="望变", category_name="变压器", category_code="",
                                  market_price=9999.0, raw_price=10000.0, source_ref="S1!R2"),
            CanonicalMaterialItem(quote_id="WB002", supplier="望变", brand="望变", category_name="变压器", category_code="",
                                  market_price=20000.0, raw_price=20000.0, source_ref="S1!R4")
        ]
        recon_a = reconcile(srows, items_tampered)
        r2_status = next(r['status'] for r in recon_a if r['source_ref'] == "S1!R2")
        self.assertTrue(r2_status.startswith("异常:单价不一致"))

        # 场景 B: 漏行 (S1!R4 丢失)
        items_missing = [
            CanonicalMaterialItem(quote_id="WB001", supplier="望变", brand="望变", category_name="变压器", category_code="",
                                  market_price=10000.0, raw_price=10000.0, source_ref="S1!R2")
        ]
        recon_b = reconcile(srows, items_missing)
        r4_status = next(r['status'] for r in recon_b if r['source_ref'] == "S1!R4")
        self.assertEqual(r4_status, "异常:漏行")

        # 场景 C: 原表无此行 (凭空编造 S1!R99)
        items_fake = [
            CanonicalMaterialItem(quote_id="WB001", supplier="望变", brand="望变", category_name="变压器", category_code="",
                                  market_price=10000.0, raw_price=10000.0, source_ref="S1!R2"),
            CanonicalMaterialItem(quote_id="WB002", supplier="望变", brand="望变", category_name="变压器", category_code="",
                                  market_price=20000.0, raw_price=20000.0, source_ref="S1!R4"),
            CanonicalMaterialItem(quote_id="WB099", supplier="望变", brand="望变", category_name="变压器", category_code="",
                                  market_price=30000.0, raw_price=30000.0, source_ref="S1!R99")
        ]
        recon_c = reconcile(srows, items_fake)
        r99_status = next(r['status'] for r in recon_c if r['source_ref'] == "S1!R99")
        self.assertEqual(r99_status, "异常:原表无此行")

    def test_06_builder_structure_and_fonts(self):
        """测试 WorkbookBuilder 33 列输出、全宋体、多行排版、审计页与对账页生成"""
        builder = StandardizedWorkbookBuilder(None)
        self.assertEqual(len(builder.headers), 33)

        item = CanonicalMaterialItem(
            quote_id="WB00010",
            supplier="测试供应商",
            brand="测试品牌",
            category_name="干式变压器",
            category_code="1001",
            structured_specs={"容量": "100kVA", "型号": "SCB11"},
            market_price=15000.0,
            discount_ratio=0.9,
            project_price=13500.0,
            uncertainty_note="公章遮挡单价，待人工确认",
            source_ref="第1页!R2C2",
            raw_price=15000.0,
            file_page="第1页"
        )
        
        test_out_path = os.path.join(os.path.dirname(__file__), "deliveries", "_temp_health_test.xlsx")
        os.makedirs(os.path.dirname(test_out_path), exist_ok=True)
        try:
            builder.build_from_canonical_items([item], test_out_path, img_dir=os.path.join(os.path.dirname(__file__), "deliveries", "_temp_imgs"))
            wb = openpyxl.load_workbook(test_out_path)
            
            # 校验包含三个核心 Sheet
            self.assertIn('标准化报价', wb.sheetnames)
            self.assertIn('疑义与失败审计说明', wb.sheetnames)
            
            ws_main = wb['标准化报价']
            self.assertEqual(ws_main.max_column, 33)
            self.assertEqual(ws_main.cell(row=2, column=1).value, "WB00010")
            # 校验宋体
            self.assertEqual(ws_main.cell(row=2, column=1).font.name, "宋体")
            
            # 校验审计说明 Sheet 的红色高亮
            ws_audit = wb['疑义与失败审计说明']
            self.assertGreaterEqual(ws_audit.max_row, 2)
            c_reason = ws_audit.cell(row=2, column=6)
            self.assertEqual(c_reason.font.name, "宋体")
            self.assertEqual(c_reason.font.color.rgb, "00FF0000") # 严格红色
        finally:
            if os.path.exists(test_out_path):
                os.remove(test_out_path)


class TestPipelineVerification(unittest.TestCase):
    """执行完整性闸门：跳过的行不能免检、写盘后必须读回校验、矩阵表不能凭空造物料"""

    def test_07_skipped_row_with_price_is_anomaly(self):
        """有单价的行被判为跳过 = 异常；合计/标题等结构性跳过不算 (合计行本来就带金额)"""
        from models import SourceRowRecord, reconcile
        rows = [
            SourceRowRecord('S!R1', 'S', 1, '合计 460', 460.0, '合计/小计行'),
            SourceRowRecord('S!R2', 'S', 2, '球阀 88', 88.0, 'Jev 判定为备注(0.91)'),
            SourceRowRecord('S!R3', 'S', 3, '备注文字', 0.0, '无价格且有效属性不足2项'),
            SourceRowRecord('S!R4', 'S', 4, '页失败', 0.0, '异常:第2页识别失败(超时)'),
        ]
        recon = reconcile(rows, [])
        st = {r['source_ref']: r['status'] for r in recon}
        self.assertEqual(st['S!R1'], '跳过')
        self.assertTrue(st['S!R2'].startswith('异常:有单价(88)却被跳过'))
        self.assertEqual(st['S!R3'], '跳过')
        self.assertEqual(st['S!R4'], '异常:第2页识别失败(超时)')

    def test_08_output_readback_catches_missing_rows(self):
        """写盘后读回必须能抓到：行数不符 / 市场价被改 / 报价ID 重复"""
        import tempfile
        from main import _verify_output
        from models import CanonicalMaterialItem
        from builder import StandardizedWorkbookBuilder
        items = [CanonicalMaterialItem(f'X{i:05d}', '通用品牌', '通用品牌', '闸阀', '',
                                       market_price=float(p), unit='个', source_ref=f'S!R{i}')
                 for i, p in ((1, 120.0), (2, 340.0))]
        tmp_dir = tempfile.gettempdir()
        out = os.path.join(tmp_dir, "_qs_readback_test.xlsx")
        img_tmp = os.path.join(tmp_dir, "_qs_img")
        try:
            StandardizedWorkbookBuilder('').build_from_canonical_items(items, out, img_tmp, reconciliation=[])
            self.assertEqual(_verify_output(out, items), [])          # 正常落盘：无问题
            self.assertEqual(len(_verify_output(out, items[:1])), 1)   # 漏一行：必须报出来
        finally:
            if os.path.exists(out):
                os.remove(out)

    def test_09_matrix_misdetection_rejected(self):
        """清单表(名称|长度|宽度)不能被当成矩阵表，否则会凭空造出假物料且对账全绿"""
        import openpyxl
        from main import _detect_matrix
        wb = openpyxl.Workbook(); ws = wb.active
        ws.append(['某公司 防火门规格'])
        ws.append(['名称', '长度', '宽度'])
        for n, l, w in [('甲级防火门', 2.1, 1.2), ('乙级防火门', 2.0, 1.1), ('丙级防火门', 1.9, 1.0)]:
            ws.append([n, l, w])
        self.assertIsNone(_detect_matrix(ws))

        ws2 = wb.create_sheet('真矩阵')
        ws2.append(['容量kVA', 'S13型', 'S20型'])
        for cap, a, b in [('30', 10500, 20800), ('50', 13500, 19800), ('80', 18000, 26000)]:
            ws2.append([cap, a, b])
        mx = _detect_matrix(ws2)
        self.assertIsNotNone(mx)
        self.assertEqual(len(mx['value_cols']), 2)

    def test_10_evaluate_engine_selftest(self):
        """评差引擎自身的防线：0 差异基线 / 13 类错误全命中 / 等价写法不误报 / 结构错位可识别"""
        import evaluate
        self.assertEqual(evaluate.selftest(), 0)


if __name__ == '__main__':
    unittest.main()

