# -*- coding: utf-8 -*-
"""
第 6 层：多 Sheet 渲染交付与对等三方原图裁决台 (Delivery & Referee)
职责：
1. 生成工业级标准多 Sheet 交付工作簿：
   - Sheet1: 厂家报价原始解析 (Raw Parsed Grid)
   - Sheet2: 标准化报价与材料属性 (Standardized Quotation & Structured Attributes)
   - Sheet3: 品类专属子表与全量明细
   - Sheet4: 原表对账与勾稽审计 (Reconciliation & Audit)
2. 三方对等核验台 (Tri-Party Referee Engine):
   - 我方解析 vs 线上系统对照 vs 原件物理微距切片；
   - 100% 动态计算汇总，严禁写死固定数字或手写假台词；
3. 遵循《报价单标准化处理强制交付规范》：
   - 严格基于 templates/dashboard_template.html 生成 D3 交互式大盘 报价单解析质量看板.html；
   - 覆盖四大核心支柱：解析准确率、分类准确率、属性值饱和度、图片凭证完整度。
"""
import os
import sys
from pathlib import Path
from typing import List, Dict, Optional, Any, Tuple

from core.ir_models import QuoteDocumentIR, CanonicalMaterialItem, SourceRowRecord
from builder import StandardizedWorkbookBuilder
from cli.generate_comparison_review_excel import (
    load_ours_data,
    load_theirs_data,
    build_review_excel,
    build_dashboard_html
)


class DeliveryRefereeLayer:
    """多 Sheet 交付与对等三方裁决台"""

    def __init__(self, template_path: Optional[str] = None):
        self.template_path = template_path

    def deliver_standard_workbook(
        self,
        doc_ir: QuoteDocumentIR,
        output_excel_path: str
    ) -> str:
        """生成标准交付 Excel 工作簿"""
        builder = StandardizedWorkbookBuilder(self.template_path)
        return builder.build(
            items=doc_ir.canonical_items,
            output_path=output_excel_path,
            terms=doc_ir.metadata.get("raw_terms_dict")
        )

    def generate_referee_and_dashboard(
        self,
        ours_excel_path: str,
        theirs_excel_path: str,
        output_review_excel_path: str,
        output_dashboard_html_path: str
    ) -> None:
        """运行三方对等核验台，生成双向核对底稿与 D3 交互式大盘"""
        ours_p = Path(ours_excel_path)
        theirs_p = Path(theirs_excel_path)

        ours_data = load_ours_data(ours_p)
        theirs_data = load_theirs_data(theirs_p)

        # 1. 渲染双向对账核对底稿 Excel
        build_review_excel(
            ours_items=ours_data,
            theirs_items=theirs_data,
            output_path=Path(output_review_excel_path),
            vendor_name=ours_p.stem
        )

        # 2. 渲染 D3 交互式质量看板 HTML
        build_dashboard_html(
            ours_items=ours_data,
            theirs_items=theirs_data,
            output_path=Path(output_dashboard_html_path),
            vendor_name=ours_p.stem
        )
