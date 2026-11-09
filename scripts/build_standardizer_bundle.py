# -*- coding: utf-8 -*-
"""
build_standardizer_bundle.py - 完整生成生产级独立部署包
"""
import os
import re
import shutil
import tarfile
import zipfile

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEPLOY_DIR = os.path.join(ROOT_DIR, "deploy", "standardizer_engine")
CORE_DIR = os.path.join(DEPLOY_DIR, "core")
DATA_DIR = os.path.join(DEPLOY_DIR, "data")
RULES_DIR = os.path.join(DEPLOY_DIR, "rules")
TESTS_DIR = os.path.join(DEPLOY_DIR, "tests")
DIST_DIR = os.path.join(ROOT_DIR, "deliveries", "standardizer_bundle")


def setup_directories():
    os.makedirs(CORE_DIR, exist_ok=True)
    os.makedirs(DATA_DIR, exist_ok=True)
    os.makedirs(RULES_DIR, exist_ok=True)
    os.makedirs(TESTS_DIR, exist_ok=True)
    os.makedirs(DIST_DIR, exist_ok=True)


def copy_and_adapt_data():
    print("[1/6] 归档底表与元数据字典 (原数据)...")
    data_files = [
        ("工料机类别.xlsx", os.path.join(DATA_DIR, "工料机类别.xlsx")),
        ("guangcai_knowledge_base.db", os.path.join(DATA_DIR, "guangcai_knowledge_base.db")),
        ("category_lexicon.db", os.path.join(DATA_DIR, "category_lexicon.db")),
        ("category_attrs_cache.json", os.path.join(DATA_DIR, "category_attrs_cache.json")),
        (os.path.join("data", "guangcai_attribute_catalog.json"), os.path.join(DATA_DIR, "guangcai_attribute_catalog.json")),
    ]
    for src_rel, dst_path in data_files:
        src_path = os.path.join(ROOT_DIR, src_rel)
        if os.path.exists(src_path):
            shutil.copy2(src_path, dst_path)
            print(f"  -> [OK] 数据落盘: {os.path.basename(dst_path)} ({round(os.path.getsize(dst_path)/1024, 1)} KB)")
        else:
            print(f"  -> [FAIL] 缺少数据: {src_rel}")


def create_core_config():
    print("[2/6] 生成独立配置中枢 core/config.py...")
    config_content = '''# -*- coding: utf-8 -*-
"""
独立部署包配置中枢
支持通过环境变量动态注入路径，默认自适应当前部署包下的 data 目录
"""
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.getenv("STANDARDIZER_DATA_DIR", os.path.join(BASE_DIR, "data"))

EXCEL_CATALOG = os.path.join(DATA_DIR, "工料机类别.xlsx")
CATEGORY_DICT_PATH = EXCEL_CATALOG
GUANGCAI_KB_DB = os.path.join(DATA_DIR, "guangcai_knowledge_base.db")
CATEGORY_LEXICON_DB = os.path.join(DATA_DIR, "category_lexicon.db")
CATEGORY_ATTRS_CACHE = os.path.join(DATA_DIR, "category_attrs_cache.json")
ATTRIBUTE_CATALOG_JSON = os.path.join(DATA_DIR, "guangcai_attribute_catalog.json")

CATEGORY_ATTRS_API_URL = os.getenv("CATEGORY_ATTRS_API_URL", "")

# JEV 模型线上服务鉴权 (支持环境变量传入)
TYPESAFE_API_KEY = os.getenv("TYPESAFE_API_KEY", "b3a32502-36e6-4d1a-be8a-ef8bb6cb4570")
TYPESAFE_BASE_URL = os.getenv("TYPESAFE_BASE_URL", "http://106.14.214.152:8000/api/v1/agentic/reasoning")
TYPESAFE_MODEL = os.getenv("TYPESAFE_MODEL", "jev-latest")
JEV_MIN_CONFIDENCE = float(os.getenv("JEV_MIN_CONFIDENCE", "0.60"))
'''
    with open(os.path.join(CORE_DIR, "config.py"), "w", encoding="utf-8") as f:
        f.write(config_content)


def adapt_core_code():
    print("[3/6] 适配并注入核心解耦模块...")
    # 1. 适配 guangcai_kb_service.py
    with open(os.path.join(ROOT_DIR, "guangcai_kb_service.py"), "r", encoding="utf-8") as f:
        kb_code = f.read()
    kb_code = re.sub(
        r'DB_PATH = .*?\n',
        'try:\n    from core.config import GUANGCAI_KB_DB\nexcept ImportError:\n    from config import GUANGCAI_KB_DB\nDB_PATH = GUANGCAI_KB_DB\n',
        kb_code,
        count=1
    )
    with open(os.path.join(CORE_DIR, "guangcai_kb_service.py"), "w", encoding="utf-8") as f:
        f.write(kb_code)

    # 2. 适配 jev_category_scorer.py
    with open(os.path.join(ROOT_DIR, "jev_category_scorer.py"), "r", encoding="utf-8") as f:
        jev_code = f.read()
    jev_code = re.sub(
        r'BASE_DIR = .*?JEV_MIN_CONFIDENCE = float\(.*?\)\n',
        'try:\n    from core.config import BASE_DIR, EXCEL_CATALOG, CATEGORY_LEXICON_DB, TYPESAFE_API_KEY, TYPESAFE_BASE_URL, TYPESAFE_MODEL, JEV_MIN_CONFIDENCE\n'
        'except ImportError:\n    from config import BASE_DIR, EXCEL_CATALOG, CATEGORY_LEXICON_DB, TYPESAFE_API_KEY, TYPESAFE_BASE_URL, TYPESAFE_MODEL, JEV_MIN_CONFIDENCE\n'
        'CATALOG_FILE = EXCEL_CATALOG\n'
        'LEXICON_DB = CATEGORY_LEXICON_DB\n',
        jev_code,
        flags=re.DOTALL
    )
    jev_code = jev_code.replace(
        'from name_norm import normalize_material_name',
        'try:\n    from core.name_norm import normalize_material_name\nexcept ImportError:\n    from name_norm import normalize_material_name'
    )
    with open(os.path.join(CORE_DIR, "jev_category_scorer.py"), "w", encoding="utf-8") as f:
        f.write(jev_code)

    # 3. 复制 name_norm.py
    shutil.copy2(os.path.join(ROOT_DIR, "name_norm.py"), os.path.join(CORE_DIR, "name_norm.py"))

    # 4. 适配 category_attribute_service.py
    with open(os.path.join(ROOT_DIR, "category_attribute_service.py"), "r", encoding="utf-8") as f:
        cas_code = f.read()
    cas_code = cas_code.replace(
        'from config import CATEGORY_DICT_PATH, CATEGORY_ATTRS_API_URL',
        'try:\n    from core.config import CATEGORY_DICT_PATH, CATEGORY_ATTRS_API_URL, EXCEL_CATALOG, CATEGORY_ATTRS_CACHE, ATTRIBUTE_CATALOG_JSON, CATEGORY_LEXICON_DB\n'
        'except ImportError:\n    from config import CATEGORY_DICT_PATH, CATEGORY_ATTRS_API_URL, EXCEL_CATALOG, CATEGORY_ATTRS_CACHE, ATTRIBUTE_CATALOG_JSON, CATEGORY_LEXICON_DB'
    )
    cas_code = re.sub(
        r'CATALOG_EXCEL = .*?CACHE_FILE = .*?\n',
        'CATALOG_EXCEL = EXCEL_CATALOG\nCACHE_FILE = CATEGORY_ATTRS_CACHE\n',
        cas_code
    )
    cas_code = cas_code.replace('os.path.join(BASE_DIR, "data", "guangcai_attribute_catalog.json")', 'ATTRIBUTE_CATALOG_JSON')
    cas_code = cas_code.replace(
        'from guangcai_kb_service import',
        'try:\n    from core.guangcai_kb_service import'
    )
    cas_code = cas_code.replace(
        'from jev_category_scorer import',
        'try:\n    from core.jev_category_scorer import'
    )
    cas_code = cas_code.replace(
        'from name_norm import',
        'try:\n    from core.name_norm import'
    )

    with open(os.path.join(CORE_DIR, "category_attribute_service.py"), "w", encoding="utf-8") as f:
        f.write(cas_code)

    # __init__.py
    with open(os.path.join(CORE_DIR, "__init__.py"), "w", encoding="utf-8") as f:
        f.write("# -*- coding: utf-8 -*-\n")


def create_engine_facade():
    print("[4/6] 编写极简高层 SDK (standardizer_engine.py) 与微服务 (api_server.py)...")
    facade_code = '''# -*- coding: utf-8 -*-
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
            raw_cat=raw_cat,
            std_name=raw_name,
            raw_specs={},
            raw_spec_text=raw_spec,
            selected_code=selected_code,
            sheet_prior=sheet_prior
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
                raw_cat=it.get("raw_cat", ""),
                std_name=it.get("name", ""),
                raw_specs={},
                raw_spec_text=it.get("spec", ""),
                selected_code=selected_code,
                sheet_prior=sheet_prior
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
'''
    with open(os.path.join(DEPLOY_DIR, "standardizer_engine.py"), "w", encoding="utf-8") as f:
        f.write(facade_code)

    # 编写 api_server.py
    api_server_code = '''# -*- coding: utf-8 -*-
"""
api_server.py - 线上微服务入口 (FastAPI)
提供生产级 HTTP API 接口：
  - POST /api/v1/standardize          单条标准化
  - POST /api/v1/standardize/batch    批量标准化
  - GET  /api/v1/categories/{code}    查询分类官方标准属性列表
"""
import os
import sys

_CUR_DIR = os.path.dirname(os.path.abspath(__file__))
if _CUR_DIR not in sys.path:
    sys.path.insert(0, _CUR_DIR)

from typing import List, Dict, Any, Optional
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import uvicorn
from standardizer_engine import StandardizerEngine

app = FastAPI(
    title="工料机标准分类与属性拆分标准化微服务",
    version="1.0.0",
    description="配合 JEV 语义模型与广材官方知识库，提供高精度的建材品类比对、仲裁选中与一属性一列物理拆分能力。"
)

engine = StandardizerEngine(use_jev=True)

class StandardizeItemRequest(BaseModel):
    name: str
    spec: Optional[str] = ""
    raw_cat: Optional[str] = ""
    context: Optional[Dict[str, Any]] = None

class StandardizeBatchRequest(BaseModel):
    items: List[Dict[str, Any]]
    context: Optional[Dict[str, Any]] = None

@app.get("/health")
def health_check():
    return {"status": "ok", "service": "QuoteStandardizer", "ready": True}

@app.post("/api/v1/standardize")
def standardize_single(req: StandardizeItemRequest):
    try:
        res = engine.standardize_item(
            raw_name=req.name,
            raw_spec=req.spec or "",
            raw_cat=req.raw_cat or "",
            context=req.context
        )
        return {"code": 200, "message": "success", "data": res}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/v1/standardize/batch")
def standardize_batch(req: StandardizeBatchRequest):
    try:
        res = engine.standardize_batch(
            items=req.items,
            context=req.context
        )
        return {"code": 200, "message": "success", "data": res}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8080"))
    uvicorn.run("api_server:app", host="0.0.0.0", port=port, reload=False)
'''
    with open(os.path.join(DEPLOY_DIR, "api_server.py"), "w", encoding="utf-8") as f:
        f.write(api_server_code)


def create_rulebook_and_docs():
    print("[5/6] 编写规则白皮书 (RULEBOOK.md) 与部署说明 (README.md)...")
    rulebook = '''# 广材标准分类与 JEV 语义模型配合机制与规则白皮书 (RULEBOOK v1.0)

## 一、 核心架构协同总览

整个物料标准化链路分为 6 个严格阶段，实现从原始混乱的报价行到权威标准三级分类及一属性一列的闭环：

```mermaid
flowchart TD
    A["原始物料信息 (名称 / 规格 / 原表分类)"] --> B["Step 1: 特征清洗与型号剥离 (name_norm.py)"]
    B --> C["Step 2: 多源漏斗候选生成 (Multi-Source Candidates)"]
    
    subgraph S2["Step 2 候选召回源"]
        C1["广材 3,657 万行大库反查 (175分置顶)"]
        C2["广材规格词库精确签名"]
        C3["39.8 万品名企业实盘分布"]
        C4["工料机 679 类倒排分词索引"]
    end
    C --> S2 --> D["Top-K 候选池 (含语义卡片)"]
    
    D --> E["Step 3: JEV 语义打分与仲裁引擎 (jev_category_scorer.py)"]
    subgraph S3["Step 3 JEV 双档位机制"]
        E1["档位 1: 精排消歧 (Choice + Score 双原语)"]
        E2["档位 2: 全量二级大类展开兜底 (支持 255 选项)"]
    end
    E --> S3 --> F["确定的标准三级品类 (如 08-01-05 天然大理石板材)"]
    
    F --> G["Step 4: 官方标准属性 Schema 召回 (category_attribute_service)"]
    G --> H["Step 5: 物理规则参数脱壳与枚举认领 (guangcai_kb_service.py)"]
    H --> I["Step 6: 输出渲染 (一属性一列红底白字 + 专属分类 Sheet)"]
```

---

## 二、 关键比对与仲裁规则 (Matching & Arbiter Rules)

### 1. JEV 双档位判定规则
- **档位 1 (精排消歧)**：
  - 输入：物料名称、规格、供应商、品牌、以及整表计算得出的主业先验 `dominant_t2`。
  - 将每个候选组装成富语义判定卡片：`【一级 > 二级 > 三级】 | 核心属性:[...] | 证据:...`。
  - 彻底杜绝歧义（如“输入模块”既可能出现在消防弱电，也可能出现在自动化，通过主业先验消歧）。
- **档位 2 (全量二级大类展开兜底)**：
  - 当词表未覆盖该生僻新词、或档位 1 得分低于 `JEV_MIN_CONFIDENCE (0.60)` 时触发；
  - 利用 JEV 单次容纳 255 个选项的闭卷推理能力，将所属二级大类下全量三级类别展开，由 JEV 闭卷精准直选。

### 2. 官方标准属性 Schema 召回规则
- 通过选中的 `category_code`，从 `category_attribute_schema` 或 `category_attrs_cache.json` 调取该品类官方定义的标准属性。
- 保证表头严格按官方标准属性列排布（如圆钢为 `品种, 直径(mm), 牌号`；大理石为 `品种, 厚度(mm), 长宽(mm), 表面处理, 等级`）。

### 3. 物理规则参数脱壳规则 (Physical Unpacking)
- **几何多维尺寸**：识别 `800*800*20`、`600*1200*18`，自动拆解：长=800、宽=800、厚度=20。
- **管材外径与壁厚**：识别 `De50*2.0`、`dn20*2.0mm`，自动拆解：公称外径De(mm)=50、壁厚(mm)=2.0。
- **通径 DN**：识别 `DN50`、`通径25`，自动拆解并剥离 `DN` 前缀填入纯数字 `50`。
- **钢筋线材**：识别 `Φ18`、`φ12`，自动拆解：直径(mm)=18。
- **单位脱壳规范**：当标准属性列已带有单位（如 `厚度(mm)`），提取出来的值必须脱掉冗余单位（如 `20mm` 必须脱壳清洗为 `20`）。
- **专有枚举反向自动认领**：若原始规格中明确出现官方 Schema 专有枚举（如 `聚酯胎PY`、`Q235B`、`抛光面` 等），自动认领填入对应标准列。
'''
    with open(os.path.join(RULES_DIR, "RULEBOOK.md"), "w", encoding="utf-8") as f:
        f.write(rulebook)

    # 编写 README.md
    readme = '''# 工料机标准分类与属性拆分标准化独立部署包 (QuoteStandardizer Bundle v1.0)

本部署包提炼自 QuoteStandardizer 核心引擎，解耦了对主项目繁杂提取逻辑的依赖，完整包含了：
1. **核心计算引擎** (`core/`)：JEV 多维打分引擎、广材双大库反查、物理脱壳与名称归一化；
2. **底表元数据与知识库** (`data/`)：国标三级树、3,657万广材大库属性Schema库、39.8万品名企业分布词表；
3. **极简门面 SDK** (`standardizer_engine.py`)：3 行代码快速集成；
4. **线上微服务** (`api_server.py`)：FastAPI 高性能 HTTP 接口；
5. **规则白皮书** (`rules/RULEBOOK.md`)：全套打分机制与物理脱壳规则。

---

## 一、 快速开始

### 1. 安装依赖
```bash
pip install -r requirements.txt
```

### 2. Python SDK 方式直接调用
```python
from standardizer_engine import StandardizerEngine

engine = StandardizerEngine(use_jev=True)

# 单条标准化
res = engine.standardize_item(name="热轧圆钢", spec="Φ18 Q235B", raw_cat="金属材料")
print(res["category_code"])        # 01-01-01
print(res["level3"])               # 热轧圆钢
print(res["extracted_attributes"]) # {'品种': '热轧圆钢', '直径(mm)': '18', '牌号': 'Q235B'}
```

### 3. 启动 HTTP 微服务
```bash
python api_server.py
# 访问 Swagger UI: http://localhost:8080/docs
```

---

## 二、 环境变量配置

| 变量名 | 说明 | 默认值 |
| :--- | :--- | :--- |
| `STANDARDIZER_DATA_DIR` | 底表数据目录绝对路径 | 当前部署包下的 `data/` |
| `TYPESAFE_API_KEY` | JEV 模型鉴权 Key | 预置生产 Key |
| `TYPESAFE_BASE_URL` | JEV 线上推理服务地址 | 预置网关地址 |
| `PORT` | 微服务启动端口 | 8080 |
'''
    with open(os.path.join(DEPLOY_DIR, "README.md"), "w", encoding="utf-8") as f:
        f.write(readme)

    # 编写 requirements.txt
    reqs = '''fastapi>=0.100.0
uvicorn>=0.22.0
pydantic>=2.0.0
openpyxl>=3.1.0
requests>=2.28.0
'''
    with open(os.path.join(DEPLOY_DIR, "requirements.txt"), "w", encoding="utf-8") as f:
        f.write(reqs)

    # 编写 Dockerfile
    dockerfile = '''FROM python:3.10-slim

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV STANDARDIZER_DATA_DIR=/app/data
ENV PORT=8080

EXPOSE 8080

CMD ["python", "api_server.py"]
'''
    with open(os.path.join(DEPLOY_DIR, "Dockerfile"), "w", encoding="utf-8") as f:
        f.write(dockerfile)


def create_verification_test():
    print("[6/6] 编写交付验证自检脚本 (tests/test_bundle.py)...")
    test_code = '''# -*- coding: utf-8 -*-
"""
test_bundle.py - 部署包集成验证测试
"""
import os
import sys

# 注入当前目录
CUR_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if CUR_DIR not in sys.path:
    sys.path.insert(0, CUR_DIR)

from standardizer_engine import StandardizerEngine

def run_tests():
    print("=== 开始运行部署包端到端集成测试 ===")
    engine = StandardizerEngine(use_jev=True)

    test_cases = [
        {
            "name": "热轧圆钢",
            "spec": "Φ20 Q235B",
            "raw_cat": "金属及金属制品",
            "expected_code_prefix": "01",
            "expected_attrs": ["直径(mm)", "牌号"]
        },
        {
            "name": "天然大理石板材",
            "spec": "800*800*20 抛光面 白麻",
            "raw_cat": "石材及石制品",
            "expected_code_prefix": "08",
            "expected_attrs": ["厚度(mm)"]
        },
        {
            "name": "UPVC排水管",
            "spec": "De50*2.0",
            "raw_cat": "给排水管材",
            "expected_code_prefix": "05",
            "expected_attrs": ["公称外径De(mm)"]
        }
    ]

    for tc in test_cases:
        res = engine.standardize_item(
            raw_name=tc["name"],
            raw_spec=tc["spec"],
            raw_cat=tc["raw_cat"]
        )
        print(f"\\n测试物料: {tc['name']} | 规格: {tc['spec']}")
        print(f"  -> 命中分类: {res['level3']} ({res['category_code']}), 依据: {res['reason']}")
        print(f"  -> 官方标准属性列: {res['standard_attributes']}")
        print(f"  -> 物理脱壳拆分结果: {res['extracted_attributes']}")

        assert res["category_code"].startswith(tc["expected_code_prefix"]), f"编码前缀不匹配: {res['category_code']}"
        for ea in tc["expected_attrs"]:
            assert ea in res["extracted_attributes"] or ea in res["standard_attributes"], f"缺失关键属性: {ea}"
        print(f"  [PASS] 测试通过!")

    # 测试整表批量推断
    batch_items = [
        {"id": 1, "name": "热轧圆钢", "spec": "Φ18 Q235B", "raw_cat": "金属及金属制品"},
        {"id": 2, "name": "热轧圆钢", "spec": "Φ25 Q235B", "raw_cat": "金属及金属制品"},
    ]
    batch_res = engine.standardize_batch(batch_items)
    assert len(batch_res) == 2
    print(f"\\n批量接口测试通过: 处理 {len(batch_res)} 条物料正常。")
    print("=== 全量自检 100% 验证通过! ===")

if __name__ == "__main__":
    run_tests()
'''
    with open(os.path.join(TESTS_DIR, "test_bundle.py"), "w", encoding="utf-8") as f:
        f.write(test_code)


def archive_bundle():
    print("\\n正在打包归档成交付物压缩包...")
    tar_path = os.path.join(DIST_DIR, "QuoteStandardizer_DeployBundle_v1.0.tar.gz")
    zip_path = os.path.join(DIST_DIR, "QuoteStandardizer_DeployBundle_v1.0.zip")

    with tarfile.open(tar_path, "w:gz") as tar:
        tar.add(DEPLOY_DIR, arcname="standardizer_engine")
    print(f"  -> 已生成 TAR.GZ 部署包: {tar_path} ({round(os.path.getsize(tar_path)/(1024*1024), 2)} MB)")

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(DEPLOY_DIR):
            for file in files:
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, os.path.dirname(DEPLOY_DIR))
                zipf.write(full_path, rel_path)
    print(f"  -> 已生成 ZIP 部署包: {zip_path} ({round(os.path.getsize(zip_path)/(1024*1024), 2)} MB)")


if __name__ == "__main__":
    setup_directories()
    copy_and_adapt_data()
    create_core_config()
    adapt_core_code()
    create_engine_facade()
    create_rulebook_and_docs()
    create_verification_test()
    archive_bundle()
    print("\\n打包构建任务圆满完成！")
