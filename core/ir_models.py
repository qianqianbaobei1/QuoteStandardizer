# -*- coding: utf-8 -*-
"""
QuoteStandardizer - 核心强类型中间表示 (Intermediate Representation, IR)
定义 6 层架构中跨层通信的标准数据结构，杜绝“临时假Excel中转”与“弱类型字典传递”。
"""
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Dict, Optional, Tuple, Any

# 重新复用已成熟且具备完备审计/多行规格排版能力的 CanonicalMaterialItem 与 SourceRowRecord
from models import CanonicalMaterialItem, SourceRowRecord, SEVERITY_ORDER, classify_note, note_field


class AssetType(str, Enum):
    """第 1 层资产路由类型"""
    QUOTE_MAIN = "quote_main"          # 报价主表 (Excel / PDF / 单据图片)
    TERMS_DISCOUNT = "terms_discount"  # 报价说明 / 工程折扣记录单
    ATTACHMENT_IGNORE = "attachment"    # 营业执照 / 检测报告 / 授权书 / 无关附件
    UNKNOWN = "unknown"                # 待进一步判定


class FailFastGateError(RuntimeError):
    """硬门禁熔断异常：当任一层数据未通过零信任校验时立即抛出，阻止脏数据流入下游或污染缓存"""
    def __init__(self, layer_name: str, gate_rule: str, detail: str):
        super().__init__(f"[{layer_name}] 硬门禁熔断: {gate_rule} -> {detail}")
        self.layer_name = layer_name
        self.gate_rule = gate_rule
        self.detail = detail


@dataclass
class DocumentAsset:
    """第 1 层：资产分流器产出单个资产描述"""
    file_path: str
    asset_type: AssetType
    confidence: float = 1.0
    page_count: int = 1
    ext: str = ""
    summary: str = ""
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class RawCell:
    """第 2 层：物理感知层提取的单元格原子物理信息"""
    row: int
    col: int
    value: Any
    formula: Optional[str] = None
    is_merged: bool = False
    bbox: Optional[Tuple[float, float, float, float]] = None  # (x0, y0, x1, y1)
    font_color: Optional[str] = None
    bg_color: Optional[str] = None


@dataclass
class RawRow:
    """第 3 层：二维表格拓扑解构引擎产出的扁平化物料行原子记录"""
    source_ref: str                   # 严格物理定位 (如 "Sheet1!R12#T1", "Page1!L08")
    sheet_or_page: str                # 来源 Sheet 或页码 (如 "第1页")
    row_index: int                    # 物理行号
    cells: Dict[str, Any]             # 列名 -> 单元格真实文本
    raw_text: str                     # 整行拼接原文
    raw_price: float = 0.0            # 原表提取的物理单价 (0幻觉物理真值)
    raw_qty: Optional[float] = None   # 原始数量 (若有)
    raw_amount: Optional[float] = None# 原始合价/金额 (若有)
    row_type: str = "item"            # "item", "title", "header", "summary", "note"
    section_title: str = ""           # 继承的上下文分组大标题 (如 "一、消火栓系统")
    is_matrix: bool = False           # 是否来自矩阵展开行
    matrix_col_label: str = ""        # 矩阵展开对应的列标题 (如 "壁厚 1.5mm")


@dataclass
class TableGridIR:
    """单张物理表格拓扑网格"""
    sheet_or_page: str
    headers: List[str]
    rows: List[RawRow]
    is_matrix: bool = False
    matrix_meta: Dict[str, Any] = field(default_factory=dict)
    title_groups: List[str] = field(default_factory=list)


@dataclass
class CommercialTermsIR:
    """第 4 层：商务条款与核算核准信息"""
    tax_rate: Optional[int] = None
    is_tax_included: str = ""
    tax_desc: str = ""
    freight_desc: str = ""
    is_freight_included: str = ""
    discount_series: str = ""
    discount_ratio: Optional[float] = None
    discount_min: Optional[float] = None
    discount_max: Optional[float] = None
    quote_date: str = ""
    province: str = ""
    city: str = ""
    quote_desc: str = ""
    source_level: str = "default"  # "row_column" > "banner" > "terms_sheet" > "default"
    raw_evidence: str = ""         # 提取该商务条款的原句证据


@dataclass
class QuoteDocumentIR:
    """统一内存中间表示 (跨层中枢流通载体)"""
    supplier_name: str = ""
    brand: str = ""
    assets: List[DocumentAsset] = field(default_factory=list)
    tables: List[TableGridIR] = field(default_factory=list)
    raw_rows: List[RawRow] = field(default_factory=list)
    commercial_terms: CommercialTermsIR = field(default_factory=CommercialTermsIR)
    canonical_items: List[CanonicalMaterialItem] = field(default_factory=list)
    reconciliation_records: List[SourceRowRecord] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)
