# -*- coding: utf-8 -*-
"""
统一物料记录 (CanonicalMaterialItem)、多行规格排版与原表对账
"""
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple


# ---------------------------------------------------------------- 待复核说明分级
# 每条说明归到一个严重级别与一个“同类问题组”：审计页先看致命/严重，提示级只留痕不打扰。
# 级别决定的是“人要不要看”，不是“这条错得有多离谱”：推断来源、编码未配置这类留痕项
# 每天都有一堆，混在致命项里会让审计页失去分辨力 (165 行全红的审计页等于没有审计页)。
SEVERITY_ORDER = {'致命': 0, '严重': 1, '提示': 2}

# (关键词, 严重级别, 同类问题组) —— 顺序即优先级，先匹配先算
NOTE_RULES = (
    # 致命：这项结果不可信 / 原表对不上，不核对不许交付
    (('漏行', '原表无此行', '一行生成了', '输出无法对应', '却被跳过', '单价不一致',
      '未识别出物料名称', '原表未读到有效单价', '识别失败', '0 条输出'), '致命', '行数与价格'),
    # 严重：读到了但可能读错，必须对着原件核这一个格子
    (('两遍识别不一致', '第二遍识别中找不到', '视觉模型标注存疑', '列标题识别存疑',
      '识别存疑', '遮挡', '看不清'), '严重', '识别遮挡'),
    (('模型返回价格', '未经大模型清洗', '模型推测', '未采用'), '严重', '清洗与比对'),
    (('参数无原文出处',), '严重', '参数出处'),
    (('低于上一档',), '严重', '价格走向'),
    (('单位由原表', '原表无计量单位'), '严重', '计量单位'),
    (('商务条款缺失', '未提供工程折扣', '未提供或未识别出商务条款',
      '无法识别为省级行政区', '不予采信'), '严重', '商务条款'),
    (('含公式但无缓存值',), '严重', '结构'),
    # 提示：只是留痕，不需要人工核对
    (('Jev',), '提示', '结构判定'),
    (('由文件名推断', '由供应商名称推断', '由页面标题推断'), '提示', '推断来源'),
    (('未经标准编码库', '类别编码未确定', '未配置标准编码库'), '提示', '分类编码'),
    (('个省份',), '提示', '报价地区'),
)
DEFAULT_RULE = ('提示', '其他')

# 说明里写的是“单价”，输出表头叫“市场价”
FIELD_SYNONYMS = {'单价': '市场价'}


def classify_note(note: str):
    """待复核说明 → (严重级别, 同类问题组)"""
    s = str(note or '')
    for keys, severity, group in NOTE_RULES:
        if any(k in s for k in keys):
            return severity, group
    return DEFAULT_RULE


def note_field(note: str, headers=None) -> str:
    """说明指的是哪个单元格：取说明里最早出现的表头名（“品牌…由供应商名称推断”算品牌，不算供应商）；
    都对不上就按同义词，再都对不上就是整行问题"""
    s = str(note or '')
    best = None
    for h in headers or []:
        h = str(h or '')
        pos = s.find(h) if len(h) >= 2 else -1
        if pos >= 0:
            key = (pos, -len(h))       # 位置越靠前越算；同位置取更长的表头
            if best is None or key < best[0]:
                best = (key, h)
    if best:
        return best[1]
    for alias, real in FIELD_SYNONYMS.items():
        if alias in s:
            return real
    return '整行'

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
    quote_image_filename: str = ""     # 厂家报价原件截图文件名 (对应页面/Sheet的原件截图)
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
    project_type: str = ""             # 项目类型
    project_volume: str = ""           # 工程用量
    quote_date: str = ""               # 报价时间
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

    # 逐条待复核说明 (审计页一行一条，按条分级)；uncertainty_note 是它的合并串，供报价说明/高亮用
    notes: List[str] = field(default_factory=list)

    def add_note(self, note: str):
        # 按“逐条说明”去重：不能用子串包含（“单价不一致”会被“单价不一致: 原表 100 / 输出 90”吞掉）
        if note and note not in self.notes:
            self.uncertainty_note = f"{self.uncertainty_note}；{note}" if self.uncertainty_note else note
            self.notes.append(note)

    def note_list(self) -> List[str]:
        """逐条说明。老数据只有合并串时按“；”拆开，宁可多拆也不漏项"""
        if self.notes:
            return list(self.notes)
        return [n for n in str(self.uncertainty_note or '').split('；') if n.strip()]

    def worst_severity(self) -> str:
        """最坏级别：'' (无待复核项) / 提示 / 严重 / 致命"""
        levels = [classify_note(n)[0] for n in self.note_list()]
        return min(levels, key=lambda s: SEVERITY_ORDER[s]) if levels else ''
    
    # 图片资产
    image_filenames: List[str] = field(default_factory=list)
    image_bytes_list: List[bytes] = field(default_factory=list)

    # 材料配图接口资产 (来自工料机类别.xlsx Sheet3 材料配图接口)
    picture_url: str = ""                               # 配图在线链接
    picture_urls: List[str] = field(default_factory=list) # 全部配图在线链接
    illustration_picture_name: str = ""                 # 配图规范文件名
    illustration_folder: str = ""                       # 配图所属目录

    def format_multiline_specs(self) -> str:
        """规格参数多行展示：按参数原顺序每行一项 “参数名:值”；过滤非物理参数内部编码；没有参数时为空"""
        forbidden = ('表头', '表标题', '基价', '单价', '总价', '金额', '指导价', '调价',
                     '产品编码', '物料编码', '商品编码', '内部编码', '条形码', '编码', '条码', '代码',
                     '税号', '账号', '帐号', '开户行', '一级类别', '二级类别', '三级类别', '材料名称', '属性项属性值',
                     '物料名称', '类别名称', '分类编码', '原始描述', '列标题', '行标题')
        lines = []
        for k, v in self.structured_specs.items():
            if v in (None, ""):
                continue
            k_str = str(k).strip()
            if any(fb in k_str for fb in forbidden):
                continue
            lines.append(f"{k}:{v}")
        return "\n".join(lines)


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


# 同类问题组 → 一次处理建议 (同一组的问题可以用同一个动作批量解决)
GROUP_ADVICE = {
    '行数与价格': '先核对或重跑，这条不许带着交付',
    '识别遮挡': '对照原件该单元格/公章区域截图核对',
    '清洗与比对': '对照原表该行原文核对，模型结论不作数',
    '参数出处': '该参数原文里找不到，确认后再保留',
    '价格走向': '档位越高价格反而越低，核对原表该列',
    '计量单位': '按原件补录单位',
    '商务条款': '补齐报价说明/折扣单，或补录 terms.json 后重跑',
    '结构': '在 Excel 中打开原表保存后重跑 (公式缺缓存值)',
    '结构判定': '留痕，无需处理',
    '推断来源': '留痕，无需处理 (已在输出中标明出处)',
    '分类编码': '留痕；配置 CATEGORY_CODE_FILE 后可消除',
    '报价地区': '留痕，无需处理',
}


def summarize_audit(items, headers=None) -> Dict:
    """
    审计分级汇总：给审计页首行与“同类问题组汇总”用。
    - 行级免审率：没有致命/严重的行 / 总行数 (人真正要动的比例)
    - 字段级免审率：1 - 存疑单元格数 / (总行 × 输出列数)
    """
    total = len(items)
    counts = {'致命': 0, '严重': 0, '提示': 0}
    review_rows = set()
    suspect_cells = set()
    clusters: Dict[tuple, Dict] = {}

    for it in items:
        for note in it.note_list():
            severity, group = classify_note(note)
            counts[severity] += 1
            field_name = note_field(note, headers)
            if severity != '提示':
                review_rows.add(it.quote_id)
            suspect_cells.add((it.quote_id, field_name))
            c = clusters.setdefault((severity, group), {'severity': severity, 'group': group,
                                                       'count': 0, 'quote_ids': [], 'fields': set()})
            c['count'] += 1
            if it.quote_id not in c['quote_ids']:
                c['quote_ids'].append(it.quote_id)
            c['fields'].add(field_name)

    field_total = total * len(headers or [])
    cluster_list = sorted(clusters.values(), key=lambda c: (SEVERITY_ORDER[c['severity']], -c['count']))
    for c in cluster_list:
        c['advice'] = GROUP_ADVICE.get(c['group'], '对照原表核对')
        c['fields'] = sorted(c['fields'])

    return {
        'total_rows': total,
        'fatal': counts['致命'],
        'severe': counts['严重'],
        'hint': counts['提示'],
        'review_rows': len(review_rows),
        'row_pass_rate': round(100 * (total - len(review_rows)) / total, 1) if total else 100.0,
        'field_pass_rate': round(100 * (1 - len(suspect_cells) / field_total), 1) if field_total else 100.0,
        'clusters': cluster_list,
    }


def reconcile(source_rows: List[SourceRowRecord], items: List[CanonicalMaterialItem]) -> 'ReconciliationReport':
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
            # 已明确标注“矩阵单元格非数值”的格是说明文字（如“DN250以上是PN10价格”），
            # 数字只是文字的一部分，属于自证跳过，不能再算“有单价却被跳过”
            explained = str(r.decision).startswith('矩阵单元格非数值')
            if r.raw_price > 0 and not explained:
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


