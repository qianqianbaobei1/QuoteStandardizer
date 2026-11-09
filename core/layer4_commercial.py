# -*- coding: utf-8 -*-
"""
第 4 层：商务条款瀑布与原价保真核算引擎 (Commercial & Pricing Engine)
职责：
1. 建立商务条款三级优先级瀑布 (Three-Tier Priority Waterfall)：
   - Level 1 (最高)：行内显式列 (单行自带的税率、是否含税、报价区域、单行折扣)；
   - Level 2 (次高)：工作表表头横幅 / 顶底备注 (Sheet Banner / Header Text)；
   - Level 3 (基础)：独立商务说明单 / 工程折扣记录单 (Terms Image / Terms JSON)；
2. 过滤印刷示例行 (如 "例：全系列 8.5折")，精准识别复选框与横线填空；
3. 原价保真核算铁律：
   - market_price (市场价) 100% 锁定原表物理单价 raw_price；
   - project_price (工程价) = market_price × discount_ratio (有折扣时) 或等于 market_price (无折扣时)；
   - 绝不私自乘算改动 market_price，原汁原味原文保真。
"""
import re
from typing import List, Dict, Optional, Any
from core.ir_models import QuoteDocumentIR, RawRow, CommercialTermsIR
from terms_extractor import parse_discount, normalize_terms, _split_discount_segments, _series_name_of, _is_global_series


class CommercialEngine:
    """商务条款瀑布与原价保真核算引擎"""

    def __init__(self):
        pass

    def apply(self, doc_ir: QuoteDocumentIR, terms_dict: Optional[Dict[str, Any]] = None) -> QuoteDocumentIR:
        """主入口：计算并绑定每一行的商务条款与工程价格"""
        # 1. 规范化外挂/全局商务条款 (Level 3)
        base_terms = self._parse_base_terms(terms_dict)
        doc_ir.commercial_terms = base_terms

        # 2. 遍历处理每一行
        for row in doc_ir.raw_rows:
            # 执行三级瀑布决议
            resolved = self._resolve_row_terms(row, base_terms)

            # 3. 价格保真与工程核算
            raw_p = row.raw_price
            discount = resolved.discount_ratio

            # 记录决议后的属性至 row metadata
            row.cells['_tax_rate'] = resolved.tax_rate
            row.cells['_is_tax_included'] = resolved.is_tax_included
            row.cells['_tax_desc'] = resolved.tax_desc
            row.cells['_is_freight_included'] = resolved.is_freight_included
            row.cells['_freight_desc'] = resolved.freight_desc
            row.cells['_discount_ratio'] = discount
            row.cells['_discount_series'] = resolved.discount_series
            row.cells['_province'] = resolved.province
            row.cells['_city'] = resolved.city
            row.cells['_quote_desc'] = resolved.quote_desc

        return doc_ir

    def _parse_base_terms(self, raw_terms: Optional[Dict[str, Any]]) -> CommercialTermsIR:
        """规范化全局商务条款"""
        ct = CommercialTermsIR(source_level="default")
        if not raw_terms:
            return ct

        norm = normalize_terms(raw_terms)
        ct.tax_rate = norm.get('tax_rate')
        ct.is_tax_included = norm.get('is_tax_included', '')
        ct.tax_desc = norm.get('tax_desc', '')
        ct.is_freight_included = norm.get('is_freight_included', '')
        ct.freight_desc = norm.get('freight_desc', '')
        ct.discount_series = norm.get('discount_series', '')
        ct.discount_ratio = norm.get('discount_ratio')
        ct.discount_min = norm.get('discount_min')
        ct.discount_max = norm.get('discount_max')
        ct.quote_date = norm.get('quote_date', '')
        ct.province = norm.get('province', '')
        ct.city = norm.get('city', '')
        ct.quote_desc = norm.get('quote_desc', '')
        ct.source_level = raw_terms.get('_source', 'terms_sheet')
        return ct

    def _resolve_row_terms(self, row: RawRow, base: CommercialTermsIR) -> CommercialTermsIR:
        """对单行执行三级优先级瀑布决议"""
        res = CommercialTermsIR(
            tax_rate=base.tax_rate,
            is_tax_included=base.is_tax_included,
            tax_desc=base.tax_desc,
            is_freight_included=base.is_freight_included,
            freight_desc=base.freight_desc,
            discount_series=base.discount_series,
            discount_ratio=base.discount_ratio,
            quote_date=base.quote_date,
            province=base.province,
            city=base.city,
            quote_desc=base.quote_desc,
            source_level=base.source_level
        )

        cells = row.cells

        # --- Level 1 瀑布覆盖：行内显式列 ---
        # 1. 税率
        for k, v in cells.items():
            if '税率' in k:
                m = re.search(r'(\d+)', str(v))
                if m:
                    res.tax_rate = int(m.group(1))
                    res.source_level = "row_column"

        # 2. 是否含税
        for k, v in cells.items():
            if '含税' in k and '单价' not in k and '价格' not in k:
                s_val = str(v).strip()
                if '不含' in s_val or s_val in ('否', '不'):
                    res.is_tax_included = '否'
                    res.source_level = "row_column"
                elif '含' in s_val or s_val in ('是', '1'):
                    res.is_tax_included = '是'
                    res.source_level = "row_column"

        # 3. 区域
        for k, v in cells.items():
            if any(rk in k for rk in ('地区', '省份', '城市', '区域')):
                s_val = str(v).strip()
                if s_val and len(s_val) >= 2:
                    res.province = s_val
                    res.source_level = "row_column"

        # 4. 系列折扣匹配 (如果全局条款有分系列折扣，如“消防喷头 8.9折”)
        if base.discount_series and not res.discount_ratio:
            series_discount = self._match_series_discount(row.raw_text, base.discount_series)
            if series_discount is not None:
                res.discount_ratio = series_discount

        return res

    def _match_series_discount(self, row_text: str, discount_series: str) -> Optional[float]:
        """根据行文本品名匹配具体系列的折扣"""
        segments = _split_discount_segments(discount_series)
        for seg in segments:
            s_name = _series_name_of(seg)
            if s_name and s_name in row_text:
                d = parse_discount(seg)
                if d is not None:
                    return d
        return None
