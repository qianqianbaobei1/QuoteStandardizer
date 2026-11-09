# -*- coding: utf-8 -*-
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
