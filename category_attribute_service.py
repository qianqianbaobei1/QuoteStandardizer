# -*- coding: utf-8 -*-
"""
材料类别特定参数及要求服务 (Material Category Attribute Service)
- 对应 Sheet1: 876条三级分类编码标准库
- 对应 Sheet2: 获取材料类别特定参数及要求接口 (https://ugw.m2.com.cn/agents-toolkit/api/material-category/attrs)
- 负责：
  1. 编码匹配与映射 (resolve_standard_category)
  2. 类别属性与规范约束缓存拉取 (fetch_category_attributes)
  3. 参数规范化、合法性校验与回填 (validate_and_normalize_specs)
"""
import os
import re
import json
import urllib.request
import urllib.error
from typing import Dict, Any, Optional, Tuple, List

CACHE_FILE = os.path.join(os.path.dirname(__file__), "category_attrs_cache.json")
# 类别属性接口地址：不配置则不调用，直接用本地规则 (CATEGORY_CODE_FILE)
API_URL = os.getenv("CATEGORY_ATTRS_API_URL", "")

# 内存缓存
_ATTR_CACHE: Dict[str, Any] = {}

def load_cache():
    global _ATTR_CACHE
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                _ATTR_CACHE = json.load(f)
        except Exception:
            _ATTR_CACHE = {}

def save_cache():
    try:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(_ATTR_CACHE, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

load_cache()

def fetch_category_attributes(category_code: str, timeout: int = 5) -> Optional[Dict[str, Any]]:
    """
    根据三级类别编码获取属性组和填写要求
    """
    if not category_code:
        return None
    if not API_URL:
        return None  # 未配置接口地址，不拼接残缺 URL
    category_code = str(category_code).strip()
    if category_code in _ATTR_CACHE:
        return _ATTR_CACHE[category_code]

    url = f"{API_URL}?categoryCode={category_code}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (QuoteStandardizer)"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            if resp.status == 200:
                payload = json.loads(resp.read().decode("utf-8"))
                if payload.get("status") == 200 and "data" in payload:
                    data = payload["data"]
                    _ATTR_CACHE[category_code] = data
                    save_cache()
                    return data
    except Exception as e:
        # 网络异常时静默或记录，不阻断主流程
        pass
    return None

def validate_and_normalize_specs(category_code: str, specs: Dict[str, Any]) -> Tuple[Dict[str, Any], List[str]]:
    """
    对照 Sheet2 接口返回的规则，对物料的具体名称/参数描述进行校验与规范化。
    规则要求：
    1. attrValueEnterType = 2 (限定录入):
       - 值必须从 values[].value 中选择。如果当前值存在于候选中或经同义转换可匹配，则纠正为规范值；
       - 若完全无法匹配，记录待复核 notes。
    2. attrValueEnterType = 1 (不限定录入):
       - 可不强制从 values 取值，原样或规范化保留。
    3. isRequired = True (必填):
       - 如果缺少该属性且已有上下文/名称可推断，补充录入；否则记录缺失警告。
    """
    if not category_code:
        return dict(specs), []

    cat_data = fetch_category_attributes(category_code)
    if not cat_data:
        return dict(specs), []

    groups = cat_data.get("attributeGroups", [])
    norm_specs = dict(specs)
    notes = []

    # 建立属性名快速查找表 (规范名称 -> group)
    attr_map = {}
    for g in groups:
        attr = g.get("attribute", {})
        desc = attr.get("description", "").strip()
        if desc:
            attr_map[desc] = g
            # 兼容去掉括号的名称
            clean_desc = re.sub(r'[（(][^）)]*[）)]', '', desc).strip()
            if clean_desc and clean_desc not in attr_map:
                attr_map[clean_desc] = g

    # 1. 检查和规范现有参数
    for k, v in list(norm_specs.items()):
        k_clean = re.sub(r'[（(][^）)]*[）)]', '', str(k)).strip()
        g = attr_map.get(str(k).strip()) or attr_map.get(k_clean)
        if not g:
            continue

        attr = g.get("attribute", {})
        attr_desc = attr.get("description")
        enter_type = attr.get("attrValueEnterType") # 1: 不限定, 2: 限定
        valid_values = [item.get("value") for item in g.get("values", []) if item.get("value")]

        curr_val = str(v).strip()
        if enter_type == 2 and valid_values:
            # 严格限定录入：必须在 values 中
            if curr_val in valid_values:
                norm_specs[k] = curr_val
            else:
                # 尝试模糊/同义匹配
                matched = None
                for candidate in valid_values:
                    if curr_val.lower() == candidate.lower():
                        matched = candidate
                        break
                    # 如 "50" 匹配 "De50" 或 "DN50"
                    if candidate.endswith(curr_val) and any(candidate.startswith(p) for p in ['DN', 'De', 'φ', 'Φ']):
                        matched = candidate
                        break
                if matched:
                    norm_specs[k] = matched
                else:
                    notes.append(f"参数[{attr_desc}]值“{curr_val}”不在限定可选值范围内")

    # 2. 检查必填项是否缺失
    for g in groups:
        attr = g.get("attribute", {})
        desc = attr.get("description", "").strip()
        is_required = attr.get("isRequired", False)
        if is_required:
            k_clean = re.sub(r'[（(][^）)]*[）)]', '', desc).strip()
            found = False
            for k, val in norm_specs.items():
                if val not in (None, "") and (k == desc or re.sub(r'[（(][^）)]*[）)]', '', str(k)).strip() == k_clean):
                    found = True
                    break
            if not found:
                notes.append(f"缺少必填参数: {desc}")

    return norm_specs, notes


def jev_score_category(material_name: str, raw_specs: str, candidate_codes: List[str]) -> Optional[Dict[str, Any]]:
    """
    使用 TypeSafe Jev 对候选分类进行选择题裁决与概率打分。
    返回: {'category_code': code, 'confidence': float, 'probabilities': dict} 或 None
    """
    from llm_agent import JevJudge
    judge = JevJudge()
    if not judge.available or not candidate_codes:
        return None

    criteria = {}
    for code in candidate_codes[:6]: # Jev 选项不宜过多，精选 Top 候选
        cat_data = fetch_category_attributes(code)
        desc = ""
        sample_types = []
        if cat_data:
            desc = cat_data.get("type", {}).get("description", "")
            # 取前几个限定品种作为辨别特征
            for g in cat_data.get("attributeGroups", []):
                if g.get("attribute", {}).get("description") == "品种":
                    sample_types = [v.get("value") for v in g.get("values", [])[:6] if v.get("value")]
                    break
        type_hint = f" (典型品种: {', '.join(sample_types)})" if sample_types else ""
        criteria[code] = f"【{code}】{desc}{type_hint}"

    criteria["NONE"] = "【NONE】以上分类均不符合"

    state = f"待判定物料名称: {material_name}\n原始技术参数与描述: {raw_specs or '(空)'}"
    questions = {
        "cat_choice": {
            "type": "choice",
            "question": f"请为待判定物料“{material_name}”选择最准确的工程材料三级分类编码：",
            "criteria": criteria
        }
    }

    ans = judge.ask(state, questions, timeout=12)
    if not ans or "cat_choice" not in ans:
        return None

    choice_res = ans["cat_choice"]
    chosen_code = choice_res.get("choice")
    confidence = choice_res.get("confidence", 0.0)
    probabilities = choice_res.get("probabilities", {})

    if chosen_code == "NONE":
        return {"category_code": "", "confidence": confidence, "probabilities": probabilities}

    return {
        "category_code": chosen_code,
        "confidence": confidence,
        "probabilities": probabilities
    }

def resolve_and_verify_category(material_name: str, raw_specs: str = "", llm_code: Optional[str] = None) -> Tuple[str, str, List[str]]:
    """
    结合标准库匹配、Jev 概率裁决打分与接口限定值校验的综合判定：
    返回: (category_code, source, notes)
    source: 'dict' (词库直中/接口校验通过), 'jev' (Jev 高置信度裁决), 'llm' (模型推荐), '' (未确定)
    """
    notes = []
    name = (material_name or "").strip()

    # 1. 尝试直接词库/后缀最长匹配
    matched_code = ""
    candidate_codes = []

    from config import CATEGORY_CODE_MAP
    if name in CATEGORY_CODE_MAP:
        candidate_codes.append(CATEGORY_CODE_MAP[name])
    hits = [k for k in CATEGORY_CODE_MAP if name.endswith(k) or k in name]
    if hits:
        # 按匹配长度倒序加入候选
        for k in sorted(hits, key=len, reverse=True)[:5]:
            code = CATEGORY_CODE_MAP[k]
            if code not in candidate_codes:
                candidate_codes.append(code)

    if llm_code and str(llm_code).strip() not in candidate_codes:
        candidate_codes.append(str(llm_code).strip())

    # 2. 如果候选多于 1 个或者匹配不确切，调用 Jev 裁判进行置信度概率打分
    if len(candidate_codes) > 1 or (candidate_codes and not (name in CATEGORY_CODE_MAP)):
        jev_res = jev_score_category(name, raw_specs, candidate_codes)
        if jev_res and jev_res.get("category_code"):
            conf = jev_res.get("confidence", 0.0)
            chosen = jev_res.get("category_code")
            if conf >= 0.8:
                return chosen, "jev", notes
            else:
                notes.append(f"Jev 分类裁决置信度偏低 ({conf:.2f})，归属为 {chosen}，需人工复核")
                return chosen, "jev", notes

    if candidate_codes:
        return candidate_codes[0], "dict", notes

    if llm_code and re.fullmatch(r'[A-Za-z0-9]+', str(llm_code).strip()):
        return str(llm_code).strip(), "llm", [f"类别编码 {llm_code} 未经标准编码库核验"]

    return "", "", ["类别编码未确定"]
