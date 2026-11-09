# -*- coding: utf-8 -*-
"""
standardizer_engine.py - 广材标准分类与动态属性拆分标准化门面 SDK
"""
import sys
import os

# 确保包内模块平滑寻址
_CUR_DIR = os.path.dirname(os.path.abspath(__file__))
if _CUR_DIR not in sys.path:
    sys.path.insert(0, _CUR_DIR)
_CORE_DIR = os.path.join(_CUR_DIR, "core")
if _CORE_DIR not in sys.path:
    sys.path.insert(0, _CORE_DIR)

from typing import Dict, List, Optional, Any
from core.category_attribute_service import (
    retrieve_category_candidates,
    validate_and_normalize_specs,
    get_category_standard_attr_names,
    compute_sheet_category_prior,
    _load_catalog
)
from core.jev_category_scorer import JevCategoryScorer, catalog_by_t2
from core.guangcai_kb_service import get_guangcai_kb_service
from core.name_norm import normalize_material_name


class StandardizerEngine:
    """
    生产级物料分类与属性标准化引擎
    """
    def __init__(self, use_jev: bool = True):
        self.use_jev = use_jev
        self.scorer = JevCategoryScorer() if use_jev else None
        self.kb_service = get_guangcai_kb_service()
        self._t2_map = catalog_by_t2()

    def standardize_item(
        self,
        raw_name: str,
        raw_spec: str = "",
        raw_cat: str = "",
        context: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        单条物料标准化：返回标准三级分类、标准三级编码、官方标准属性清单、拆解后的属性键值对
        """
        ctx = context or {}
        sheet_prior = ctx.get("sheet_prior")

        # 1. 候选集生成 (融合大库品名、规格词库、企业实盘分布、工料机倒排索引)
        cands = retrieve_category_candidates(
            raw_cat=raw_cat,
            std_name=raw_name,
            raw_spec=raw_spec,
            sheet_prior=sheet_prior,
            top_k=6
        )

        selected_code = ""
        category_name = ""
        level1, level2 = "", ""
        confidence = 0.0
        reason = "未命中"
        tier = "fallback"

        # 2. 若启用 JEV 模型且有候选，进行语义打分消歧
        if self.use_jev and self.scorer and self.scorer.enabled and cands:
            item_payload = [{
                "key": "single_0",
                "std_name": raw_name,
                "raw_cat": raw_cat,
                "raw_spec": raw_spec,
                "candidates": cands
            }]
            sheet_ctx = {
                "supplier": ctx.get("supplier", ""),
                "brand": ctx.get("brand", ""),
                "dominant_t2": list(sheet_prior) if sheet_prior else [],
                "sample_names": [raw_name]
            }
            res = self.scorer.score_batch_items(item_payload, sheet_ctx, catalog_t2_map=self._t2_map)
            if "single_0" in res:
                r = res["single_0"]
                selected_code = r.get("code", "")
                level1 = r.get("level1", "")
                level2 = r.get("level2", "")
                category_name = r.get("level3", "")
                confidence = float(r.get("confidence", 0.0))
                tier = r.get("tier", "tier1")
                reason = f"JEV模型仲裁({tier}, 置信度:{confidence})"

        # 若 JEV 未命中或未启用，回退到规则第一候选
        if not selected_code and cands:
            best = cands[0]
            selected_code = best["code"]
            level1 = best["level1"]
            level2 = best["level2"]
            category_name = best["level3"]
            confidence = round(float(best.get("score", 0)) / 100.0, 3)
            reason = best.get("reason", "规则候选库第一名")
            tier = "rule_best"

        # 3. 调取官方标准属性体系并执行物理脱壳
        std_attrs = get_category_standard_attr_names(selected_code) if selected_code else []
        norm_specs, raw_unmapped, warnings = validate_and_normalize_specs(
            level3_name=category_name,
            specs={},
            raw_spec_text=raw_spec
        )

        return {
            "query_name": raw_name,
            "raw_spec": raw_spec,
            "raw_category": raw_cat,
            "category_code": selected_code,
            "level1": level1,
            "level2": level2,
            "level3": category_name,
            "confidence": confidence,
            "tier": tier,
            "reason": reason,
            "standard_attributes": std_attrs,
            "extracted_attributes": norm_specs,
            "unmapped_attributes": raw_unmapped,
            "warnings": warnings
        }

    def standardize_batch(
        self,
        items: List[Dict[str, Any]],
        context: Optional[Dict[str, Any]] = None
    ) -> List[Dict[str, Any]]:
        """
        批量物料标准化 (带整表主业先验推断)
        items: [{"id": 1, "name": "...", "spec": "...", "raw_cat": "..."}]
        """
        ctx = context or {}
        raw_rows = [{"raw_cat": it.get("raw_cat", ""), "std_name": it.get("name", "")} for it in items]
        sheet_prior = compute_sheet_category_prior(raw_rows)
        ctx["sheet_prior"] = sheet_prior

        prepared = []
        cands_map = {}
        for idx, it in enumerate(items):
            key = str(it.get("id", idx))
            cands = retrieve_category_candidates(
                raw_cat=it.get("raw_cat", ""),
                std_name=it.get("name", ""),
                raw_spec=it.get("spec", ""),
                sheet_prior=sheet_prior,
                top_k=6
            )
            cands_map[key] = cands
            prepared.append({
                "key": key,
                "std_name": it.get("name", ""),
                "raw_cat": it.get("raw_cat", ""),
                "raw_spec": it.get("spec", ""),
                "candidates": cands
            })

        jev_res = {}
        if self.use_jev and self.scorer and self.scorer.enabled:
            sheet_ctx = {
                "supplier": ctx.get("supplier", ""),
                "brand": ctx.get("brand", ""),
                "dominant_t2": list(sheet_prior) if sheet_prior else [],
                "sample_names": [it.get("name", "") for it in items[:15]]
            }
            jev_res = self.scorer.score_batch_items(prepared, sheet_ctx, catalog_t2_map=self._t2_map)

        results = []
        for idx, it in enumerate(items):
            key = str(it.get("id", idx))
            cands = cands_map.get(key, [])
            r = jev_res.get(key)
            if r and r.get("code"):
                selected_code = r.get("code", "")
                level1 = r.get("level1", "")
                level2 = r.get("level2", "")
                category_name = r.get("level3", "")
                confidence = float(r.get("confidence", 0.0))
                tier = r.get("tier", "tier1")
                reason = f"JEV模型仲裁({tier})"
            elif cands:
                best = cands[0]
                selected_code = best["code"]
                level1 = best["level1"]
                level2 = best["level2"]
                category_name = best["level3"]
                confidence = round(float(best.get("score", 0)) / 100.0, 3)
                reason = best.get("reason", "规则候选第一名")
                tier = "rule_best"
            else:
                selected_code, level1, level2, category_name = "", "", "", ""
                confidence = 0.0
                reason = "未命中工料机库"
                tier = "none"

            std_attrs = get_category_standard_attr_names(selected_code) if selected_code else []
            norm_specs, raw_unmapped, warnings = validate_and_normalize_specs(
                level3_name=category_name,
                specs={},
                raw_spec_text=it.get("spec", "")
            )

            results.append({
                "id": it.get("id", idx),
                "query_name": it.get("name", ""),
                "raw_spec": it.get("spec", ""),
                "category_code": selected_code,
                "level1": level1,
                "level2": level2,
                "level3": category_name,
                "confidence": confidence,
                "tier": tier,
                "reason": reason,
                "standard_attributes": std_attrs,
                "extracted_attributes": norm_specs,
                "warnings": warnings
            })

        return results
