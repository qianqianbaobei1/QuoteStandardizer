# -*- coding: utf-8 -*-
"""
guangcai_kb_service.py
广材标准分类与属性智能回填服务（基于 11GB 全量大库编译的 guangcai_knowledge_base.db）

核心能力：
1. 【品名查分类】通过图片/PDF提取的材料名称，毫秒级反查广材权威三级品类与置信度；
2. 【提取标准Schema】提取该分类官方标准属性列表（包含标准Key、单位、高频枚举值）；
3. 【属性智能对标回填】将图片/PDF解析出的原始散乱参数，按官方属性规范认领、脱壳、归一化回填。
"""
import os
import re
import json
import sqlite3
from typing import Dict, List, Optional, Any, Tuple

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "guangcai_knowledge_base.db")

class GuangcaiKnowledgeBaseService:
    def __init__(self, db_path: str = DB_PATH):
        self.db_path = db_path
        self._conn = None
        if os.path.exists(self.db_path):
            self._conn = sqlite3.connect(self.db_path, check_same_thread=False)

    def is_ready(self) -> bool:
        return self._conn is not None

    def match_category_and_schema(self, material_name: str) -> Optional[Dict[str, Any]]:
        """
        通过物料名称匹配广材标准三级分类及官方属性体系规范
        """
        if not self.is_ready() or not material_name:
            return None

        cur = self._conn.cursor()
        clean_name = material_name.strip()

        # 1. 精确匹配
        cur.execute(
            "SELECT level1, level2, level3, sample_count, confidence FROM product_category_index WHERE product_name = ?",
            (clean_name,)
        )
        row = cur.fetchone()

        # 2. 前缀/模糊包含匹配
        if not row and len(clean_name) >= 3:
            cur.execute(
                "SELECT level1, level2, level3, sample_count, confidence FROM product_category_index "
                "WHERE product_name LIKE ? OR ? LIKE ('%' || product_name || '%') "
                "ORDER BY sample_count DESC LIMIT 1",
                (f"%{clean_name}%", clean_name)
            )
            row = cur.fetchone()

        if not row:
            return None

        l1, l2, l3, count, conf = row

        # 3. 提取该三级分类的官方标准属性体系
        cur.execute(
            "SELECT total_specs, attributes_json, sample_enums_json FROM category_attribute_schema WHERE level3 = ?",
            (l3,)
        )
        s_row = cur.fetchone()
        attrs = json.loads(s_row[1]) if s_row and s_row[1] else []
        enums = json.loads(s_row[2]) if s_row and s_row[2] else {}

        # 4. 获取 3 个代表性参考规格
        cur.execute(
            "SELECT spec_json FROM category_specs WHERE level3 = ? LIMIT 3",
            (l3,)
        )
        sample_specs = [json.loads(r[0]) for r in cur.fetchall() if r[0]]

        return {
            "query_name": clean_name,
            "level1": l1,
            "level2": l2,
            "level3": l3,
            "confidence": conf,
            "sample_count": count,
            "standard_attributes": attrs,
            "sample_enums": enums,
            "sample_specs": sample_specs
        }

    def backfill_attributes(self, raw_spec_text: str, standard_attrs: List[Dict[str, Any]]) -> Dict[str, str]:
        """
        根据官方属性Schema要求，将图片/PDF解析出来的散乱文本按要求回填
        自动处理：
        1. 单位脱壳（如 '20mm' 填入 '厚度(mm)' 时自动变为 '20'）
        2. 常见别名映射（如 '口径' 填入 '公称外径De(mm)'）
        3. 规格复合模式拆解（如 '800*800' 拆为 长度800、宽度800）
        """
        filled = {}
        if not raw_spec_text or not standard_attrs:
            return filled

        text = raw_spec_text.strip()

        # 预拆解常见长宽厚几何尺寸 (如 800×800 或 800*800*20)
        dim_match = re.search(r'(\d+(?:\.\d+)?)\s*[×\*xX]\s*(\d+(?:\.\d+)?)(?:\s*[×\*xX]\s*(\d+(?:\.\d+)?))?', text)
        dim_w, dim_h, dim_t = None, None, None
        if dim_match:
            dim_w = dim_match.group(1)
            dim_h = dim_match.group(2)
            dim_t = dim_match.group(3)

        # 预拆解管材常见外径*壁厚 (例如 De50*2.0 或 20*2.0mm)
        pipe_m = re.search(r'(?:De|dn)?\s*(\d+(?:\.\d+)?)\s*[×\*xX]\s*(\d+(?:\.\d+)?)\s*(?:mm)?', text, re.I)
        pipe_de, pipe_t = None, None
        if pipe_m:
            pipe_de = pipe_m.group(1)
            pipe_t = pipe_m.group(2)

        # 逐个对标标准属性
        for attr in standard_attrs:
            key_name = attr.get("name", "")
            unit = attr.get("unit", "")
            core_key = re.sub(r'[(（].*?[)）]', '', key_name).strip()

            val = None

            # A. 显式键值匹配 (例如 '厚度:20' 或 '厚度(mm)=20')
            kv_pattern = rf'(?:{re.escape(key_name)}|{re.escape(core_key)})\s*[:：=]\s*([^;,\n\r]+)'
            kv_m = re.search(kv_pattern, text, re.I)
            if kv_m:
                val = kv_m.group(1).strip()

            # B. 管材外径与壁厚
            if not val:
                if ('外径' in core_key or 'De' in key_name) and pipe_de:
                    val = pipe_de
                elif ('壁厚' in core_key) and pipe_t:
                    val = pipe_t

            # C. 几何尺寸回填 (长、宽、厚)
            if not val:
                if '长' in core_key and dim_w:
                    val = dim_w
                elif '宽' in core_key and dim_h:
                    val = dim_h
                elif ('厚' in core_key or '壁厚' in core_key) and dim_t:
                    val = dim_t

            # D. 常用公称外径 / 壁厚 / 压力正则
            if not val:
                if '外径' in core_key or 'De' in key_name:
                    de_m = re.search(r'[Dd][Ee]\s*(\d+(?:\.\d+)?)', text)
                    if de_m:
                        val = de_m.group(1)
                elif '壁厚' in core_key:
                    t_m = re.search(r'(?:壁厚|厚度|厚)\s*[:：]?\s*(\d+(?:\.\d+)?)\s*(?:mm)?', text)
                    if t_m:
                        val = t_m.group(1)
                elif '压力' in core_key or 'PN' in key_name:
                    pn_m = re.search(r'(?:PN|压力)?\s*(\d+(?:\.\d+)?)\s*MPa', text, re.I)
                    if pn_m:
                        val = pn_m.group(1)
                elif '管系' in core_key:
                    s_m = re.search(r'(S\d+(?:\.\d+)?)', text, re.I)
                    if s_m:
                        val = s_m.group(1)

            # E. 枚举值直接或子串命中
            if not val:
                for cand_val in attr.get("common_values", []):
                    if not cand_val or len(cand_val) < 2:
                        continue
                    if cand_val in text:
                        val = cand_val
                        break
                    # 去掉尾部辅助词模糊匹配 (如 电动式闭门器 -> 电动闭门器)
                    cand_simp = cand_val.replace('式', '')
                    if cand_simp in text and len(cand_simp) >= 2:
                        val = cand_val
                        break

            # 统一单位脱壳与清洗：如果字段名自带单位，去掉值里的冗余单位
            if val:
                val = val.strip().strip(';，,')
                if unit:
                    val = re.sub(rf'\s*{re.escape(unit)}$', '', val, flags=re.I).strip()
                filled[key_name] = val

        return filled
