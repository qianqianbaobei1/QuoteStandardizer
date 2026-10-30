# -*- coding: utf-8 -*-
"""
统一物料记录 (CanonicalMaterialItem)、多行规格排版与原表对账
"""
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple

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
    """原表中被扫描到的一行，以及解析器对它的判定 (用于对账与数学勾稽)"""
    source_ref: str
    sheet: str
    row: int
    raw_text: str
    raw_price: float
    decision: str            # 'item' = 作为物料输出；其余为跳过原因
    has_formula_without_cache: bool = False
    raw_qty: Optional[float] = None       # 数量
    raw_amount: Optional[float] = None    # 合价 / 金额
    math_audit: str = ""                  # 数学勾稽结论


# 强物料特征词：即使某行含有“合计/小计”，若同时包含这些特征且有价格，坚决不能当汇总行跳过！
MATERIAL_FEATURE_KEYWORDS = (
    'dn', 'de', 'mm', 'cm', 'm²', 'm2', 'm³', 'm3', 'kg', 'kv', 'kva', 'w', 'kw',
    '型号', '规格', '系列', '米', '台', '套', '个', '支', '条', '件', '只', '根', '块', '卷'
)
STRUCTURAL_SKIP = {'分组标题行', '重复表头行', '合计/小计行'}


def reconcile(source_rows: List[SourceRowRecord], items: List[CanonicalMaterialItem]) -> Tuple[List[dict], dict]:
    """
    原表 ↔ 标准化结果逐行对账与四维防欺瞒体检。
    返回: (对账明细列表, 体检量化指标汇总)
    检查项：
      1. 物理行数守恒闭环：每个原表物理行必须有且仅有明确去向 (物料或显式合法跳过)
      2. 每一条输出单价必须与原表单价 100% 严格一致 (逐条校验，杜绝多条时漏检后项)
      3. 封堵 STRUCTURAL_SKIP 误杀漏洞：有单价且带物料特征的行坚决不豁免
      4. 输出中不得出现原表不存在的幽灵行
      5. 数学勾稽验证：单价 × 数量 == 合价
    """
    by_ref: Dict[str, List[CanonicalMaterialItem]] = {}
    for it in items:
        by_ref.setdefault(it.source_ref, []).append(it)

    known_refs = {r.source_ref for r in source_rows}
    report = []
    
    math_audited = 0
    math_passed = 0
    total_anomalies = 0

    for r in source_rows:
        outs = by_ref.get(r.source_ref, [])
        
        # 数学勾稽核算: 单价 * 数量 == 合价
        if r.raw_price > 0 and r.raw_qty and r.raw_qty > 0 and r.raw_amount and r.raw_amount > 0:
            math_audited += 1
            expected = round(r.raw_price * r.raw_qty, 2)
            diff = abs(expected - r.raw_amount)
            # 允许 5 分钱以内的微小乘积精度容差
            if diff <= max(0.05, r.raw_amount * 0.002):
                math_passed += 1
                r.math_audit = f"勾稽自洽: {r.raw_price:g}×{r.raw_qty:g}≈{r.raw_amount:g}"
            else:
                r.math_audit = f"勾稽异常: 单价({r.raw_price:g})×数量({r.raw_qty:g})={expected:g}≠合价({r.raw_amount:g})"

        if r.decision != 'item':
            # 跳过行的防漏检审计
            is_suspicious_kill = False
            if r.raw_price > 0:
                raw_lower = r.raw_text.lower()
                # 检查是否包含规格特征词
                has_spec_feat = any(k in raw_lower for k in MATERIAL_FEATURE_KEYWORDS)
                if has_spec_feat or r.decision not in STRUCTURAL_SKIP:
                    is_suspicious_kill = True

            if outs:
                status = '异常:被跳过的行仍有输出'
                total_anomalies += 1
            elif r.decision.startswith('异常'):
                status = r.decision
                total_anomalies += 1
            elif is_suspicious_kill:
                status = f'异常:有单价({r.raw_price:g})却被跳过'
                total_anomalies += 1
            else:
                status = '跳过'
            report.append({
                'source_ref': r.source_ref, 'status': status, 'detail': r.decision,
                'raw_price': r.raw_price, 'out_price': None, 'quote_id': '',
                'raw_text': r.raw_text, 'math_audit': r.math_audit
            })
            continue

        if not outs:
            total_anomalies += 1
            report.append({
                'source_ref': r.source_ref, 'status': '异常:漏行', 'detail': '原表物料行未生成任何输出',
                'raw_price': r.raw_price, 'out_price': None, 'quote_id': '',
                'raw_text': r.raw_text, 'math_audit': r.math_audit
            })
            continue

        problems = []
        if len(outs) > 1:
            problems.append(f'一行生成了 {len(outs)} 条输出')

        # 逐条全面核验单价 (杜绝仅比对 outs[0] 的漏洞)
        for idx, o_it in enumerate(outs):
            if abs(o_it.market_price - r.raw_price) > 1e-6:
                diff_label = f'单价不一致: 原表 {r.raw_price} / 输出#{idx+1} {o_it.market_price}' if len(outs) > 1 else f'单价不一致: 原表 {r.raw_price} / 输出 {o_it.market_price}'
                problems.append(diff_label)
                total_anomalies += 1

        if r.raw_price <= 0:
            problems.append('原表未读到有效单价')
        if r.has_formula_without_cache:
            problems.append('含公式但无缓存值，请在 Excel 中打开并保存后重跑')
        if r.math_audit and r.math_audit.startswith("勾稽异常"):
            problems.append(r.math_audit)
            total_anomalies += 1

        for p in problems:
            for o in outs:
                o.add_note(p)

        status = '一致' if not problems else '异常:' + '；'.join(problems)
        if not problems and any(o.uncertainty_note for o in outs):
            status = '一致(有待复核项)'
        report.append({
            'source_ref': r.source_ref, 'status': status,
            'detail': '；'.join(o.uncertainty_note for o in outs if o.uncertainty_note),
            'raw_price': r.raw_price, 'out_price': outs[0].market_price,
            'quote_id': ', '.join(o.quote_id for o in outs),
            'raw_text': r.raw_text, 'math_audit': r.math_audit
        })

    # 反查原表无此行的幽灵物料
    for ref, outs in by_ref.items():
        if ref not in known_refs:
            total_anomalies += len(outs)
            for o in outs:
                o.add_note('输出无法对应到原表行')
                report.append({
                    'source_ref': ref or '(空)', 'status': '异常:原表无此行',
                    'detail': o.uncertainty_note or '原表无此行(幽灵行)', 'raw_price': None,
                    'out_price': o.market_price, 'quote_id': o.quote_id,
                    'raw_text': o.raw_spec, 'math_audit': ''
                })

    # 计算四维防欺瞒量化体检指标
    total_src = len(source_rows)
    skipped_cnt = sum(1 for rep in report if rep['status'] == '跳过')
    matched_cnt = sum(1 for rep in report if rep['status'].startswith('一致'))
    
    # 行数守恒率 = (合法物料行 + 合法跳过行) / 原表总物理扫描行
    row_conservation_rate = round(((matched_cnt + skipped_cnt) / total_src * 100), 2) if total_src > 0 else 0.0
    math_pass_rate = round((math_passed / math_audited * 100), 2) if math_audited > 0 else 100.0

    integrity_summary = {
        "total_source_rows": total_src,
        "matched_items": matched_cnt,
        "skipped_rows": skipped_cnt,
        "anomalies_count": total_anomalies,
        "row_conservation_rate": row_conservation_rate,
        "math_audited_count": math_audited,
        "math_passed_count": math_passed,
        "math_pass_rate": math_pass_rate,
        "is_zero_defect": total_anomalies == 0 and row_conservation_rate == 100.0 and math_pass_rate == 100.0
    }

    return ReconciliationReport(report, integrity_summary)


class ReconciliationReport(list):
    """兼具 list 列表原生兼容性与四维防欺瞒量化指标 summary 的增强对账结果对象"""
    def __init__(self, report: list, summary: dict):
        super().__init__(report)
        self.summary = summary
        self.integrity_summary = summary


