import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

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
        """纯净单表交付：33 列、全宋体、多行排版，致命/严重行主表红字黄底留痕"""
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
        
        import tempfile
        test_out_path = os.path.join(tempfile.gettempdir(), "_qs_builder_health_test.xlsx")
        try:
            builder.build_from_canonical_items([item], test_out_path, img_dir=os.path.join(tempfile.gettempdir(), "_qs_builder_tmp_imgs"))
            wb = openpyxl.load_workbook(test_out_path)
            
            # 契约：只要【标准化报价】单张 Sheet（审计内容靠主表红字黄底+批注承载）
            self.assertEqual(wb.sheetnames, ['标准化报价'])

            ws_main = wb['标准化报价']
            self.assertEqual(ws_main.max_column, 33)
            self.assertEqual(ws_main.cell(row=2, column=1).value, "WB00010")
            # 校验宋体
            self.assertEqual(ws_main.cell(row=2, column=1).font.name, "宋体")

            # 审计分级落到主表：`公章遮挡单价` 属严重 → 整行红字黄底，市场价带批注
            name_cell = ws_main.cell(row=2, column=3)
            self.assertEqual(name_cell.font.color.rgb, "00C00000", "严重行应红字")
            self.assertEqual(name_cell.fill.start_color.rgb, "00FFF2CC", "严重行应黄底")
            price_col = next(i for i in range(1, 34) if ws_main.cell(row=1, column=i).value == '市场价')
            self.assertIsNotNone(ws_main.cell(row=2, column=price_col).comment, "市场价应带审计批注")
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

    def test_11_batch_pass_requires_integrity(self):
        """批量运行器不得把“行数>0”当 PASS：完整性缺陷必须让该家判 FAIL"""
        import shutil
        import tempfile
        from pathlib import Path
        from unittest.mock import patch
        import run_batch_standardize as rbs

        tmp = tempfile.mkdtemp(prefix="qs_batch_status_")
        vendor = Path(tmp) / "某供应商"
        out_dir = Path(tmp) / "out"
        vendor.mkdir(); out_dir.mkdir()
        wb = openpyxl.Workbook(); ws = wb.active
        ws.append(["名称", "单价"]); ws.append(["闸阀", 100])
        wb.save(vendor / "报价单.xlsx")
        out = out_dir / "某供应商-标准解析.xlsx"
        # 现在 PASS 要求交付件自身合格：单 Sheet + 有市场价且为正
        wb = openpyxl.Workbook()
        ws_out = wb.active
        ws_out.title = '标准化报价'
        ws_out.append(["报价ID", "市场价"]); ws_out.append(["A1", 100])
        wb.save(out)
        try:
            with patch.object(rbs, "process_excel", return_value=(str(out), [])):
                self.assertEqual(rbs.run_single_vendor(vendor, out_dir)["status"], "PASS")
            with patch.object(rbs, "process_excel", return_value=(str(out), ["物理行数不守恒: 50.0% < 100.0%"])):
                r = rbs.run_single_vendor(vendor, out_dir)
                self.assertEqual(r["status"], "FAIL")
                self.assertIn("完整性校验未通过", r["error_msg"])
        finally:
            shutil.rmtree(tmp, ignore_errors=True)


    # ---------------------------------------------------------------- 审计分级 / 动态列 / 金标不变量
    def test_12_audit_grading_and_pass_rate(self):
        """待复核说明要分级：推演类留痕不打扰人，致命/严重才进待核验；免审率要有分母"""
        from models import classify_note, summarize_audit
        self.assertEqual(classify_note('原表未读到有效单价')[0], '致命')
        self.assertEqual(classify_note('单价不一致: 原表 100 / 输出 99')[0], '致命')
        self.assertEqual(classify_note('品牌“望变”由供应商名称推断'), ('提示', '推断来源'))
        self.assertEqual(classify_note('类别编码未确定'), ('提示', '分类编码'))
        self.assertEqual(classify_note('单价(公章边缘遮挡数字，读数10500)'), ('严重', '识别遮挡'))
        self.assertEqual(classify_note('参数无原文出处: 厚度=3')[1], '参数出处')

        a = CanonicalMaterialItem('A00010', '供', '品', '闸阀', '', market_price=10.0)
        a.add_note('品牌“供”由供应商名称推断')
        b = CanonicalMaterialItem('A00020', '供', '品', '闸阀', '', market_price=20.0)
        b.add_note('单价(公章边缘遮挡数字，读数20)')
        c = CanonicalMaterialItem('A00030', '供', '品', '闸阀', '', market_price=0.0)
        c.add_note('原表未读到有效单价')
        headers = ['报价ID', '市场价', '品牌']
        s = summarize_audit([a, b, c], headers)
        self.assertEqual((s['total_rows'], s['review_rows']), (3, 2))
        self.assertEqual((s['fatal'], s['severe'], s['hint']), (1, 1, 1))
        self.assertEqual(s['row_pass_rate'], 33.3)
        self.assertEqual(a.worst_severity(), '提示')
        self.assertEqual(b.worst_severity(), '严重')
        self.assertEqual(c.worst_severity(), '致命')
        self.assertEqual(s['clusters'][0]['severity'], '致命')          # 致命排最前
        self.assertIn('识别遮挡', [cl['group'] for cl in s['clusters']])

    def test_13_dynamic_attribute_columns_follow_header_names(self):
        """模板动态属性列按表头名取值：AD/AE/AF 列序不同也能对得上 (112101 vs 112115)"""
        import shutil
        import tempfile
        from builder import StandardizedWorkbookBuilder

        item = CanonicalMaterialItem(
            'SBS00010', '北新禹王', '禹王', '弹性体改性沥青防水卷材', '112101',
            structured_specs={'品种': 'SBS I PY PE PE 3', '厚度(mm)': '3',
                              '胎基材料': '聚酯毡PY', '覆面材料': '聚乙烯膜PE', '是否曝露使用': '否'},
            market_price=28.0, unit='m²')
        expect = {'品种': 'SBS I PY PE PE 3', '厚度(mm)': '3',
                  '胎基材料': '聚酯毡PY', '覆面材料': '聚乙烯膜PE', '是否曝露使用': '否'}
        tmp = tempfile.mkdtemp(prefix='qs_dyn_')
        try:
            # 同一套物料、两种列序的模板：列位置变了，靠表头名写对
            for i, dyn in enumerate([['品种', '厚度(mm)', '胎基材料', '覆面材料', '是否曝露使用'],
                                     ['品种', '厚度(mm)', '是否曝露使用', '胎基材料', '覆面材料']]):
                tpl = os.path.join(tmp, f'tpl{i}.xlsx')
                wb = openpyxl.Workbook()
                wb.active.append(['报价ID', '名称规格', '市场价', '单位'] + dyn)
                wb.save(tpl)
                out = os.path.join(tmp, f'out{i}.xlsx')
                StandardizedWorkbookBuilder(tpl).build_from_canonical_items([item], out, os.path.join(tmp, f'img{i}'))
                ws = openpyxl.load_workbook(out)['标准化报价']
                cols = {str(c.value): c.column for c in ws[1] if c.value}
                for h, v in expect.items():
                    self.assertEqual(ws.cell(2, cols[h]).value, v, f'列序{i} 的 {h} 填错')
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    def test_14_golden_delivery_invariants(self):
        """已入库交付件的全量不变量：报价ID唯一、价>0、图片后缀合法、留痕项必进审计页"""
        import glob
        import re as _re
        files = glob.glob(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                       'deliveries', '*-标准解析.xlsx'))
        if not files:
            self.skipTest('deliveries 下无标准解析件')
        allowed_ext = ('.png', '.jpg', '.jpeg', '.emf', '.bmp', '.gif', '.webp', '.bin')
        for f in files:
            wb = openpyxl.load_workbook(f, data_only=True)
            ws = wb['标准化报价']
            col = {str(c.value): c.column for c in ws[1] if c.value}
            if '疑义与失败审计说明' not in wb.sheetnames and '市场价' not in col:
                continue
            # 单 Sheet 契约下没有审计页：留痕只落在批注与报价说明，仍要校验其余交付不变量
            has_audit = '疑义与失败审计说明' in wb.sheetnames
            zero_price = 0
            audit_ids, reason_col = self._audit_index(wb['疑义与失败审计说明']) if has_audit else ({}, 1)
            seen = []
            for r in range(2, ws.max_row + 1):
                qid = ws.cell(r, col['报价ID']).value
                if not qid:
                    continue
                seen.append(qid)
                price = ws.cell(r, col['市场价']).value
                self.assertIsInstance(price, (int, float), f'{os.path.basename(f)} 第{r}行 市场价非数值')
                if price <= 0:
                    # 原表本就无价的行允许留 0，但必须有留痕（批注或报价说明）
                    c = ws.cell(r, col['市场价']).comment
                    trace = (str(c.text) if c else '') + str(ws.cell(r, col['报价说明']).value or '')
                    self.assertIn('单价', trace, f'{os.path.basename(f)} 第{r}行 0 价却无留痕')
                    zero_price += 1
                img = str(ws.cell(r, col['厂家报价图片文件名']).value or '无')
                self.assertTrue(img == '无' or img.lower().endswith(allowed_ext), f'{qid} 图片文件名 {img}')
                notes = str(ws.cell(r, col['报价说明']).value or '')
                if has_audit:
                    if _re.search(r'待人工复核|留痕', notes):
                        self.assertIn(str(qid), audit_ids, f'{qid} 主表留痕但审计页漏项')
                    if not ws.cell(r, col['单位']).value:
                        reasons = ' '.join(audit_ids.get(str(qid), {}).get('reasons', []))
                        self.assertIn('计量单位', reasons, f'{qid} 无单位却没写原因')
                else:
                    # 无审计页时，无单位必须在主表留痕（批注或报价说明）
                    if not ws.cell(r, col['单位']).value:
                        note = ws.cell(r, col['市场价']).comment
                        text = (str(note.text) if note else '') + notes
                        self.assertIn('计量单位', text, f'{qid} 无单位且无任何留痕')
            self.assertEqual(len(seen), len(set(seen)), f'{os.path.basename(f)} 报价ID重复')
            self.assertLessEqual(zero_price, max(1, len(seen) * 0.5),
                                 f'{os.path.basename(f)} 过半行市场价不为正')
            self.assertGreaterEqual(reason_col, 1)

    @staticmethod
    def _audit_index(ws):
        """审计页列定位：兼容旧版(表头在第1行)与新版(首行免审率、表头在第2行)"""
        cols, header_row = {}, 0
        for hr in (1, 2, 3):
            vals = [str(ws.cell(hr, c).value or '') for c in range(1, (ws.max_column or 1) + 1)]
            if '报价ID' in vals:
                cols = {v: i + 1 for i, v in enumerate(vals) if v}
                header_row = hr
                break
        reason_col = next((c for h, c in cols.items() if '原因' in h), 0)
        index = {}
        for r in range(header_row + 1, (ws.max_row or 1) + 1):
            qid = ws.cell(r, cols.get('报价ID', 1)).value
            if not qid:
                continue
            reason = str(ws.cell(r, reason_col).value or '') if reason_col else ''
            index.setdefault(str(qid), {'reasons': []})['reasons'].append(reason)
        return index, reason_col
    def test_15_vision_payload_schema_rejected(self):
        """视觉模型响应的结构硬约束：不合格就整份作废，不能把半截表格当结果"""
        from pdf_extractor import validate_vision_payload as v
        self.assertEqual(v({'tables': [{'cells': [['a', 'b'], ['1', '2']]}], 'page_title': 'x'}), '')
        self.assertTrue(v({}))
        self.assertTrue(v({'tables': []}))
        self.assertTrue(v({'tables': [{'cells': []}]}))
        self.assertTrue(v({'tables': [{'cells': [['a', 'b'], ['1']]}]}))     # 列数不齐
        self.assertTrue(v({'tables': [{'cells': ['a', 'b']}]}))              # 行不是数组
        self.assertTrue(v({'tables': [{'cells': [['a']], 'uncertain': [[1]]}]}))  # uncertain 项过短
        self.assertTrue(v({'tables': [{'cells': [['a']]}], 'notes': 5}))

    def test_17_row_values_covers_standard_headers(self):
        """_row_values 的键必须覆盖 STANDARD_HEADERS 的所有非图片列，否则模板里那一列会静默为空"""
        from builder import STANDARD_HEADERS, StandardizedWorkbookBuilder
        from models import CanonicalMaterialItem
        item = CanonicalMaterialItem(quote_id='X1', supplier='s', brand='b',
                                     category_name='n', category_code='c')
        vals = StandardizedWorkbookBuilder._row_values(item, 10, '', '', '', '')
        missing = [h for h in STANDARD_HEADERS if h not in vals and h != '图片']
        self.assertEqual(missing, [], f'_row_values 漏了标准列: {missing}')

    def test_16_llm_auditor_cannot_touch_price_or_verbatim(self):
        """自愈审计铁律：模型不得改价格/原文留痕/标识列，且改写必须留痕（痕到批注，不污染业务字段）"""
        import openpyxl
        from unittest import mock
        from llm_auditor import LLMAuditor

        wb = openpyxl.Workbook()
        ws = wb.active
        headers = ['报价ID', '名称规格', '规格参数', '厂家原始规格参数', '单位', '市场价', '报价说明']
        ws.append(headers)
        ws.append(['WL00010', '旧品名', '厚度:1.5', '原文:厚度1.5;基价18000', '㎡', 128.5, ''])

        fake = {'choices': [{'message': {'content': __import__('json').dumps({
            'need_healing': True, 'summary': 'x', 'corrections': [{
                'row_id': 1,
                'fixes': {'名称规格': '成品名', '市场价': 99.0,
                          '厂家原始规格参数': '被模型改写', '单位': 'm²', '报价说明': '基价另计'}
            }]})}}]}
        with mock.patch('llm_auditor.requests.post') as post:
            post.return_value = mock.Mock(status_code=200, json=lambda: fake)
            LLMAuditor().verify_and_auto_heal(ws, headers)

        vals = {h: ws.cell(2, i + 1).value for i, h in enumerate(headers)}
        self.assertEqual(vals['市场价'], 128.5, '价格只能取原表，模型无权改写')
        self.assertEqual(vals['厂家原始规格参数'], '原文:厚度1.5;基价18000', '原文留痕不可改写')
        self.assertEqual(vals['名称规格'], '成品名', '允许字段应当被自愈')
        self.assertEqual(vals['报价说明'], '基价另计', '报价说明里的业务修正照常写入')
        self.assertIn('[模型自愈]', str(ws.cell(2, 2).comment.text), '自愈痕迹必须留痕到被改单元格的批注')

    def test_18_fixed_fields_reject_foreign_values(self):
        """固定字段值域门禁：品牌/地区列读到价格时必须拦下，不能当成品牌/省名落盘"""
        from main import _looks_like_brand, _region_fields, _field_value_problems
        # 品牌：价格/数量/规格串都不是品牌
        for v in ('1340', 1340, '¥2,450', '1,280.00', '4.8', '13%'):
            self.assertFalse(_looks_like_brand(v), f'{v!r} 不应被当成品牌')
        for v in ('山东佰腾', '3M', '奥斯特', 'TOTO'):
            self.assertTrue(_looks_like_brand(v), f'{v!r} 应被当成品牌')
        # 地区：认不出的原值不准回填到省列（特普丽原单“报价区域”列里就是单价）
        prov, city, note = _region_fields('11')
        self.assertEqual(prov, '')
        self.assertIn('不是省市名称', note)
        self.assertEqual(_region_fields('京津冀+山东')[0], '北京市、天津市、河北省、山东省')
        # 读回体检：整表里能定位到出问题的那一行
        header = ['品牌', '报价地区省']
        rows = [['奥斯特', '广东省'], ['1340', '11']]
        errs = _field_value_problems(header, rows)
        self.assertEqual(len(errs), 2)
        self.assertTrue(any('第3行' in e and '品牌' in e for e in errs))
        self.assertTrue(any('第3行' in e and '报价地区省' in e for e in errs))

    def test_19_seal_enhancement_and_occlusion_penetration(self):
        """
        公章全生命周期自动化质检门禁：
        1. 自动定位红色公章圆心与半径；
        2. 自动检测五角星朝向并计算旋转偏角，实现扶正；
        3. 极坐标展开 (Warp Polar) 将环形文字展开为水平长条；
        4. 光学多通道解耦脱印 (De-inking) 与反向非锐化掩模 (Unsharp Masking) 凸线边缘增强；
        5. 确保被遮挡区域文字高保真凸出穿透，彻底杜绝 E13 印章遮挡断笔/漏读。
        """
        import cv2
        import numpy as np
        from seal_enhancer_service import SealEnhancerService

        # 构造合成测试样张：含黑色文本与遮挡的红色倾斜印章
        h, w = 200, 200
        canvas = np.full((h, w, 3), 255, dtype=np.uint8)
        # 写入黑色表格字迹
        cv2.putText(canvas, "13%", (80, 100), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 0), 2)
        cv2.putText(canvas, "24.2", (80, 140), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 0), 2)

        # 盖上红色倾斜公章 (圆心 100, 100, 半径 60)
        cv2.circle(canvas, (100, 100), 60, (0, 0, 220), 3)
        # 内部画一个偏转的五角星/红色色块
        cv2.circle(canvas, (100, 100), 15, (0, 0, 220), -1)

        service = SealEnhancerService()
        res = service.process_image(canvas)

        self.assertTrue(res["has_seal"], "必须检出红色公章")
        self.assertIsNotNone(res["seal_bbox"], "必须输出公章 BBox 坐标")
        self.assertIsNotNone(res["penetrated_text_img"], "必须生成脱印穿透与凸线强化图像")
        self.assertIsNotNone(res["upright_seal_img"], "必须输出扶正公章图")
        self.assertIsNotNone(res["polar_strip_img"], "必须输出极坐标水平文本条")

        # 验证脱印穿透图：尺寸必须一致
        penetrated = res["penetrated_text_img"]
        self.assertEqual(penetrated.shape, (h, w), "穿透图像尺寸必须与原图一致")


if __name__ == '__main__':
    unittest.main()

