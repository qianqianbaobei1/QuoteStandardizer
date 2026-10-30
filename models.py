# -*- coding: utf-8 -*-
"""
统一物料记录 (CanonicalMaterialItem)、多行规格排版与原表对账
"""
from dataclasses import dataclass, field
from typing import List, Dict, Optional

@dataclass
class CanonicalMaterialItem:
    """统一物料记录。未知字段一律为空 ("" / None)，不预设任何业务默认值"""
    quote_id: str                      # 报价ID (品牌拼音前缀 + 序号)
    supplier: str                      # 供应商名称
    brand: str                         # 品牌
    category_name: str                 # 物料品类名称
    category_code: str                 # 分类编码 (来自标准编码库)

    structured_specs: Dict[str, str] = field(default_factory=dict)  # 结构化技术参数

    raw_spec: str = ""                 # 厂家原始规格参数 (原文留痕)
    file_page: str = ""                # 原始文件来源页 / Sheet
    unit: str = ""                     # 计量单位

    market_price: float = 0.0          # 原表单价
    discount_ratio: Optional[float] = None   # 工程折扣比例
    project_price: Optional[float] = None    # 工程价 = 原表单价 × 折扣 (折扣未知时为空)

    # 商务条款 (来自商务单识别或人工 JSON)
    discount_series: str = ""
    is_freight_included: str = ""
    freight_desc: str = ""
    is_tax_included: str = ""
    tax_desc: str = ""
    tax_rate: Optional[int] = None
    freight_settle: str = ""
    quote_desc: str = ""

    quote_type: str = ""               # 价格填报类型
    province: str = ""                 # 报价地区省
    city: str = ""                     # 报价地区市
    discount_min: Optional[float] = None
    discount_max: Optional[float] = None

    # 待复核说明：识别存疑、信息缺失、推断得出的字段
    uncertainty_note: str = ""

    # 溯源：对应原表位置 (如 "Sheet1!R12#T1" / "第3页!L15")，用于原表 ↔ 结果逐行对账
    source_ref: str = ""
    raw_price: float = 0.0             # 原表读到的单价 (未经任何模型处理)
    clean_source: str = ""             # 清洗来源: llm / fallback
    category_source: str = ""          # 编码来源: dict / llm / default

    def add_note(self, note: str):
        if note and note not in self.uncertainty_note:
            self.uncertainty_note = f"{self.uncertainty_note}；{note}" if self.uncertainty_note else note
    
    # 图片资产
    image_filenames: List[str] = field(default_factory=list)
    image_bytes_list: List[bytes] = field(default_factory=list)

    def format_multiline_specs(self) -> str:
        """规格参数多行展示：按参数原顺序每行一项 “参数名:值”；没有参数时为空"""
        return "\n".join(f"{k}:{v}" for k, v in self.structured_specs.items() if v not in (None, ""))


@dataclass
class SourceRowRecord:
    """原表中被扫描到的一行，以及解析器对它的判定 (用于对账)"""
    source_ref: str
    sheet: str
    row: int
    raw_text: str
    raw_price: float
    decision: str            # 'item' = 作为物料输出；其余为跳过原因
    has_formula_without_cache: bool = False


def reconcile(source_rows: List[SourceRowRecord], items: List[CanonicalMaterialItem]) -> List[dict]:
    """
    原表 ↔ 标准化结果逐行对账。返回对账明细 (每个原表行一条)，并把发现的问题写回 item.uncertainty_note。
    检查项：
      1. 判定为物料的原表行必须恰好对应 1 条输出 (漏行 / 重复)
      2. 输出单价必须与原表单价完全一致 (防模型篡改价格)
      3. 输出中不得出现原表不存在的行 (防模型编造)
    """
    by_ref: Dict[str, List[CanonicalMaterialItem]] = {}
    for it in items:
        by_ref.setdefault(it.source_ref, []).append(it)

    known_refs = {r.source_ref for r in source_rows}
    report = []
    for r in source_rows:
        outs = by_ref.get(r.source_ref, [])
        if r.decision != 'item':
            status = '跳过' if not outs else '异常:被跳过的行仍有输出'
            report.append({'source_ref': r.source_ref, 'status': status, 'detail': r.decision,
                           'raw_price': r.raw_price, 'out_price': None, 'quote_id': '', 'raw_text': r.raw_text})
            continue
        if not outs:
            report.append({'source_ref': r.source_ref, 'status': '异常:漏行', 'detail': '原表物料行未生成输出',
                           'raw_price': r.raw_price, 'out_price': None, 'quote_id': '', 'raw_text': r.raw_text})
            continue
        problems = []
        if len(outs) > 1:
            problems.append(f'一行生成了 {len(outs)} 条输出')
        it = outs[0]
        if abs(it.market_price - r.raw_price) > 1e-6:
            problems.append(f'单价不一致: 原表 {r.raw_price} / 输出 {it.market_price}')
        if r.raw_price <= 0:
            problems.append('原表未读到有效单价')
        if r.has_formula_without_cache:
            problems.append('含公式但无缓存值，请在 Excel 中打开并保存后重跑')
        for p in problems:
            for o in outs:
                o.add_note(p)
        status = '一致' if not problems else '异常:' + '；'.join(problems)
        if not problems and it.uncertainty_note:
            status = '一致(有待复核项)'
        report.append({'source_ref': r.source_ref, 'status': status, 'detail': it.uncertainty_note,
                       'raw_price': r.raw_price, 'out_price': it.market_price,
                       'quote_id': ', '.join(o.quote_id for o in outs), 'raw_text': r.raw_text})

    for ref, outs in by_ref.items():
        if ref not in known_refs:
            for o in outs:
                o.add_note('输出无法对应到原表行')
                report.append({'source_ref': ref or '(空)', 'status': '异常:原表无此行', 'detail': o.uncertainty_note,
                               'raw_price': None, 'out_price': o.market_price, 'quote_id': o.quote_id, 'raw_text': o.raw_spec})
    return report
