# -*- coding: utf-8 -*-
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
