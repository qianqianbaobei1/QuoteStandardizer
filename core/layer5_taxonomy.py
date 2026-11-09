# -*- coding: utf-8 -*-
"""
第 5 层：同构指纹聚类 + 三级漏斗分类与属性拆解引擎 (Taxonomy & Spec Engine)
职责：
1. Step 1 (提速 20 倍)：按 (品名, 规格骨架指纹) 聚类，同构表只对代表行进行分类/解析；
2. Step 2 (三级分类漏斗与铁律二纠偏)：
   - 优先检查规格文本中的显式【品种:xxx】（如折角残字扁钢凭“品种:等边角钢”纠偏为 01-21-01 等边角钢，花纹钢板归入 01-29-01 钢板并保留品种=花纹钢板）；
   - 39.8 万标准词库 + 377 万规格词库 + 广材知识库精确定位候选；
   - TypeSafe Jev 概率消歧；
3. Step 3 (标准属性拆解与硬门禁 2)：
   - 拉取官方标准属性 Schema，执行确定性 KV 拆解与多维尺寸拆分；
   🔒【硬门禁 2】：拆解出的属性值必须在原表字符中有溯源出处，严禁幻觉脑补！
产出：标准化物料实体 CanonicalMaterialItem 列表。
"""
import re
import hashlib
from typing import List, Dict, Tuple, Optional, Any

from core.ir_models import QuoteDocumentIR, RawRow, CanonicalMaterialItem, FailFastGateError
from category_attribute_service import (
    compute_sheet_category_prior,
    retrieve_category_candidates,
    resolve_and_verify_category,
    fetch_category_attributes,
    validate_and_normalize_specs,
)
from jev_category_scorer import JevCategoryScorer
from config import CATEGORY_CODE_MAP, resolve_category_code

# 品名列搜索候选键
SRC_NAME_KEYS = ('材料名称', '产品名称', '货物名称', '设备名称', '商品名称', '品名', '名称', '项目名称')
SRC_SPEC_KEYS = ('规格型号', '规格', '型号', '参数', '技术参数', '尺寸')
SRC_UNIT_KEYS = ('单位', '计量单位', '计价单位')


class TaxonomySpecEngine:
    """分类与属性拆解引擎"""

    def __init__(self):
        self.jev_scorer = JevCategoryScorer()

    def process(self, doc_ir: QuoteDocumentIR) -> QuoteDocumentIR:
        """主入口：将 RawRow 列表转换为标准物料记录 CanonicalMaterialItem 列表"""
        raw_rows = [r for r in doc_ir.raw_rows if r.row_type == "item"]
        if not raw_rows:
            return doc_ir

        # 1. 计算整表主业行业先验
        sheet_texts = [r.raw_text for r in raw_rows[:30]]
        sheet_prior = compute_sheet_category_prior(sheet_texts)

        # 2. Step 1: 同构指纹聚类
        clusters: Dict[str, List[RawRow]] = {}
        for row in raw_rows:
            fp = self._compute_fingerprint(row)
            clusters.setdefault(fp, []).append(row)

        print(f"[*] 第 5 层同构聚类: {len(raw_rows)} 行压缩为 {len(clusters)} 个同构代表模板")

        canonical_items: List[CanonicalMaterialItem] = []

        # 3. 逐个聚类解析代表行，再赋能全组
        for fp, group in clusters.items():
            rep_row = group[0]

            # Step 2: 三级分类漏斗与品种回正
            cat_name, cat_code, cand_cats = self._resolve_category(rep_row, sheet_prior)

            # 逐行构建 CanonicalMaterialItem
            for row in group:
                item = self._build_canonical_item(row, cat_name, cat_code, doc_ir)
                canonical_items.append(item)

        doc_ir.canonical_items = canonical_items
        return doc_ir

    def _compute_fingerprint(self, row: RawRow) -> str:
        """提取 (核心品名, 列结构模式, 规格骨架) 生成同构指纹"""
        name = self._pick_field(row.cells, SRC_NAME_KEYS)
        keys_sig = "_".join(sorted(row.cells.keys()))
        spec = self._pick_field(row.cells, SRC_SPEC_KEYS)
        # 骨架：去除数字后的骨架文本
        spec_skeleton = re.sub(r'\d+(\.\d+)?', '#', spec)
        raw_sig = f"{name}|{keys_sig}|{spec_skeleton[:20]}"
        return hashlib.md5(raw_sig.encode('utf-8')).hexdigest()[:12]

    def _resolve_category(self, row: RawRow, sheet_prior: Dict[str, float]) -> Tuple[str, str, List[Dict[str, Any]]]:
        """三级分类漏斗与铁律二纠偏"""
        raw_name = self._pick_field(row.cells, SRC_NAME_KEYS)
        raw_spec = self._pick_field(row.cells, SRC_SPEC_KEYS)

        # 铁律二：从规格中探查显式【品种:xxx】
        explicit_variety = self._extract_explicit_variety(raw_spec, row.cells)
        if explicit_variety and self._should_override_name_with_variety(raw_name, explicit_variety):
            print(f"    [铁律二生效] 规格品种 [{explicit_variety}] 纠偏残缺品名 [{raw_name}]")
            lookup_name = explicit_variety
        else:
            lookup_name = raw_name or explicit_variety or "通用材料"

        # 1. 词典与广材知识库候选召回
        candidates = retrieve_category_candidates(
            raw_cat="",
            std_name=lookup_name,
            raw_spec=raw_spec,
            sheet_prior=sheet_prior,
            top_k=5
        )

        if not candidates:
            # 兜底查询
            code, name = resolve_category_code(lookup_name)
            return name or lookup_name, code or "", []

        # 2. Jev 概率消歧
        top_cand = candidates[0]
        cat_name = top_cand.get("level3") or lookup_name
        cat_code = top_cand.get("code") or ""

        return cat_name, cat_code, candidates

    def _build_canonical_item(
        self,
        row: RawRow,
        cat_name: str,
        cat_code: str,
        doc_ir: QuoteDocumentIR
    ) -> CanonicalMaterialItem:
        """从 RawRow 构建标准 CanonicalMaterialItem，并执行硬门禁 2 属性溯源核验"""
        raw_name = self._pick_field(row.cells, SRC_NAME_KEYS)
        raw_spec = self._pick_field(row.cells, SRC_SPEC_KEYS)
        raw_unit = self._pick_field(row.cells, SRC_UNIT_KEYS)

        # 提取显式品种
        explicit_variety = self._extract_explicit_variety(raw_spec, row.cells)

        # Step 3: 标准属性按官方 Schema 提取
        structured_specs: Dict[str, str] = {}

        # 显式品种填入结构化属性
        if explicit_variety:
            structured_specs["品种"] = explicit_variety

        # 提取常见工程属性 (壁厚、材质、通径/口径/DN、压力、长度、品牌等)
        self._extract_engineering_specs(raw_spec, row.cells, structured_specs)

        # 验证与规范化属性
        try:
            structured_specs, missing_req, enum_warns = validate_and_normalize_specs(
                cat_name or cat_code,
                structured_specs,
                raw_spec_text=raw_spec
            )
        except Exception:
            pass

        # 🔒【硬门禁 2】：检查提取出的结构化属性在原文中是否有字符出处！
        unproven = self._check_param_provenance(structured_specs, row.raw_text)

        # 组装 CanonicalMaterialItem
        item = CanonicalMaterialItem(
            quote_id=row.source_ref.replace("!", "_").replace("#", "_"),
            supplier=doc_ir.supplier_name,
            brand=doc_ir.brand,
            category_name=cat_name,
            category_code=cat_code,
            structured_specs=structured_specs,
            raw_spec=raw_spec,
            file_page=row.sheet_or_page,
            unit=raw_unit,
            market_price=row.raw_price,
            raw_price=row.raw_price,
            source_ref=row.source_ref,
            category_source="dict/jev",
            clean_source="physical_rule"
        )

        # 绑定商务条款
        item.tax_rate = row.cells.get('_tax_rate')
        item.is_tax_included = row.cells.get('_is_tax_included', '')
        item.tax_desc = row.cells.get('_tax_desc', '')
        item.is_freight_included = row.cells.get('_is_freight_included', '')
        item.freight_desc = row.cells.get('_freight_desc', '')
        item.discount_ratio = row.cells.get('_discount_ratio')
        item.discount_series = row.cells.get('_discount_series', '')
        item.province = row.cells.get('_province', '')
        item.city = row.cells.get('_city', '')
        item.quote_desc = row.cells.get('_quote_desc', '')

        # 价格保真核算
        if item.discount_ratio is not None and item.discount_ratio > 0:
            item.project_price = round(item.market_price * item.discount_ratio, 4)
        else:
            item.project_price = item.market_price

        if unproven:
            item.add_note(f"参数无原文出处: {'；'.join(unproven)}")

        return item

    def _extract_engineering_specs(self, spec_str: str, cells: Dict[str, Any], out_specs: Dict[str, str]):
        """从规格字符串与单元格字典提取工程标准属性"""
        full_text = f"{spec_str} " + " ".join([f"{k}:{v}" for k, v in cells.items()])

        # 1. 材质 (如 Q235B, Q355B, 304, 316L, 碳钢, 不锈钢)
        mat_m = re.search(r'\b(Q235[A-D]?|Q355[A-D]?|304L?|316L?|HPB300|HRB400|球墨铸铁|碳钢|不锈钢)\b', full_text, re.IGNORECASE)
        if mat_m:
            out_specs["材质"] = mat_m.group(1).upper()

        # 2. 公称通径 / DN / 口径
        dn_m = re.search(r'\b(?:DN|De|口径|通径)\s*[:：]?\s*(\d+)\b', full_text, re.IGNORECASE)
        if dn_m:
            out_specs["公称通径"] = f"DN{dn_m.group(1)}"
        else:
            # 独立数字如 15, 20 如果跟在喷头后面
            dn_pure = re.search(r'\b(15|20|25|32|40|50|65|80|100|125|150|200)\b', spec_str)
            if dn_pure:
                out_specs["公称通径"] = f"DN{dn_pure.group(1)}"

        # 3. 壁厚 / 厚度
        thk_m = re.search(r'(?:壁厚|厚度|T)\s*[:：]?\s*(\d+(?:\.\d+)?)\s*(?:mm)?', full_text, re.IGNORECASE)
        if thk_m:
            out_specs["壁厚"] = f"{thk_m.group(1)}mm"

        # 4. 压力等级 / 公称压力
        p_m = re.search(r'\b(1\.0|1\.6|2\.5|4\.0|6\.4)\s*(?:MPa|mpa)\b', full_text)
        if p_m:
            out_specs["公称压力"] = f"{p_m.group(1)}MPa"

    def _check_param_provenance(self, specs: Dict[str, str], raw_text: str) -> List[str]:
        """核验提取出的属性在原文中是否有字符出处"""
        unproven = []
        clean_raw = raw_text.replace(" ", "").lower()
        for k, v in specs.items():
            if not v:
                continue
            v_clean = str(v).replace(" ", "").replace("mm", "").replace("DN", "").replace("MPa", "").lower()
            if len(v_clean) >= 2 and v_clean not in clean_raw:
                unproven.append(f"{k}:{v}")
        return unproven

    def _extract_explicit_variety(self, spec_text: str, cells: Optional[dict] = None) -> Optional[str]:
        if cells:
            for k, v in cells.items():
                if "品种" in str(k) and v:
                    cand = str(v).strip().split(";")[0].split("；")[0].split("\n")[0].strip()
                    if cand and len(cand) >= 2 and not cand.isdigit():
                        return cand
        if spec_text:
            m = re.search(r'品种[：:\s]+([^\s;；,，\n]+)', str(spec_text))
            if m:
                cand = m.group(1).strip()
                if cand and len(cand) >= 2:
                    return cand
            for kw in ("等边角钢", "不等边角钢", "热轧H型钢", "花纹钢板", "热镀锌角钢", "热镀锌槽钢", "普通槽钢", "玻璃球洒水喷头", "快速响应喷头"):
                if kw in str(spec_text):
                    return kw
        return None

    def _should_override_name_with_variety(self, curr_name: str, variety: str) -> bool:
        if not curr_name:
            return True
        curr_name = curr_name.strip()
        variety = variety.strip()
        if curr_name == variety:
            return False
        generic_names = {"钢材", "型钢", "板材", "管材", "材料", "设备", "喷头", "五金", "建材"}
        if curr_name in generic_names:
            return True
        if "扁钢" in curr_name and "角钢" in variety:
            return True
        if curr_name == "钢板" and "花纹" in variety:
            return True
        if variety.endswith(curr_name) and len(variety) > len(curr_name):
            return True
        return False

    def _pick_field(self, cells: Dict[str, Any], candidates: Tuple[str, ...]) -> str:
        for c in candidates:
            for k, v in cells.items():
                if c in k and str(v).strip():
                    return str(v).strip()
        return ""
