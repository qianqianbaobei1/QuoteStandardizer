# -*- coding: utf-8 -*-
"""
QuoteStandardizer - 工业级六层流水线总编排器 (Master 6-Layer Pipeline)
将 6 个正交层有机串联，完全基于统一强类型 QuoteDocument IR 通信：
L1: 资产分流器 (AssetRouter)
L2: 三轨高保真物理感知层 (PerceptionLayer) [🔒硬门禁 1: RapidOCR >= 60% 重合]
L3: 二维表格拓扑解构引擎 (TableTopologyEngine)
L4: 商务条款瀑布与原价保真核算引擎 (CommercialEngine)
L5: 同构指纹聚类 + 三级漏斗分类与属性拆解引擎 (TaxonomySpecEngine) [🔒硬门禁 2: 属性字符溯源]
L6: 多 Sheet 渲染交付与对等三方原图裁决台 (DeliveryRefereeLayer)
"""
import os
import sys
from typing import Optional, Dict, Any

from core.ir_models import QuoteDocumentIR, FailFastGateError
from core.layer1_router import AssetRouter
from core.layer2_perception import PerceptionLayer
from core.layer3_topology import TableTopologyEngine
from core.layer4_commercial import CommercialEngine
from core.layer5_taxonomy import TaxonomySpecEngine
from core.layer6_delivery import DeliveryRefereeLayer


class MasterQuotePipeline:
    """工业级报价单全自动解析流水线"""

    def __init__(self, template_path: Optional[str] = None):
        self.router = None
        self.perception = PerceptionLayer()
        self.topology = TableTopologyEngine()
        self.commercial = CommercialEngine()
        self.taxonomy = TaxonomySpecEngine()
        self.delivery = DeliveryRefereeLayer(template_path=template_path)

    def run(
        self,
        input_path: str,
        output_excel_path: str,
        terms_dict: Optional[Dict[str, Any]] = None,
        theirs_excel_path: Optional[str] = None,
        output_review_excel_path: Optional[str] = None,
        output_dashboard_html_path: Optional[str] = None
    ) -> QuoteDocumentIR:
        """端到端高内聚执行 6 层流水线"""
        print(f"\n=======================================================")
        print(f"[*] 启动工程报价单工业级标准 6 层流水线")
        print(f"    输入路径: {input_path}")
        print(f"    输出交付: {output_excel_path}")
        print(f"=======================================================")

        # 第 1 层：资产分流
        print(f"\n[L1 资产分流器] 正在扫描并分类文件资产...")
        self.router = AssetRouter(input_path)
        doc_ir = self.router.route()
        doc_ir.metadata["raw_terms_dict"] = terms_dict
        print(f"    识别供应商: {doc_ir.supplier_name or '未探查到（以文件名/表头为准）'}")
        print(f"    资产统计: {len(doc_ir.assets)} 个文件")
        for a in doc_ir.assets:
            print(f"      - [{a.asset_type.value}] {os.path.basename(a.file_path)}: {a.summary}")

        # 第 2 层：物理感知
        print(f"\n[L2 物理感知层] 启动高保真物理提取...")
        doc_ir = self.perception.perceive(doc_ir)
        total_tables = len(doc_ir.tables)
        print(f"    提取表格: 共 {total_tables} 个物理表格网格")

        # 第 3 层：拓扑解构
        print(f"\n[L3 拓扑解构引擎] 展开表格并生成原子物料行...")
        doc_ir = self.topology.deconstruct(doc_ir)
        print(f"    解构产出: 共 {len(doc_ir.raw_rows)} 个扁平化物料行 (已剥离汇总与备注)")

        # 第 4 层：商务条款与核算
        print(f"\n[L4 商务条款瀑布与原价保真引擎] 计算条款与工程价格...")
        doc_ir = self.commercial.apply(doc_ir, terms_dict=terms_dict)
        print(f"    商务条款决议完成 (折扣系列: '{doc_ir.commercial_terms.discount_series or '无'}', 税率: {doc_ir.commercial_terms.tax_rate}%)")

        # 第 5 层：分类与属性拆解
        print(f"\n[L5 分类与属性引擎] 启动同构聚类与三级漏斗分类...")
        doc_ir = self.taxonomy.process(doc_ir)
        print(f"    标准化物料: 产出 {len(doc_ir.canonical_items)} 条标准化记录")

        # 第 6 层：交付渲染
        print(f"\n[L6 交付与裁决台] 生成标准交付工作簿...")
        os.makedirs(os.path.dirname(os.path.abspath(output_excel_path)), exist_ok=True)
        self.delivery.deliver_standard_workbook(doc_ir, output_excel_path)
        print(f"    [OK] 标准交付 Excel 已生成: {output_excel_path}")

        # 若提供了对照数据，自动执行三方裁决并输出 D3 质量看板
        if theirs_excel_path and os.path.exists(theirs_excel_path) and output_dashboard_html_path:
            print(f"\n[L6 交付与裁决台] 启动三方对等核验并生成 D3 交互式大盘...")
            stats = self.delivery.generate_referee_and_dashboard(
                ours_excel_path=output_excel_path,
                theirs_excel_path=theirs_excel_path,
                output_review_excel_path=output_review_excel_path or output_excel_path.replace(".xlsx", "_双向对账核对底稿.xlsx"),
                output_dashboard_html_path=output_dashboard_html_path
            )
            print(f"    [OK] 动态 D3 质量看板已落盘: {output_dashboard_html_path}")
            print(f"    [OK] 双向对账核对底稿已落盘: {output_review_excel_path}")

        print(f"\n=======================================================")
        print(f"🎉 工业级六层流水线顺利完成！")
        print(f"=======================================================\n")
        return doc_ir
