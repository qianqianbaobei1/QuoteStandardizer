# -*- coding: utf-8 -*-
"""
测试环境隔离：在任何项目模块 import 之前把外部服务相关配置清空/关闭，
保证测试不读本机 .env 里的密钥、不访问网络，在 CI 与本机结果一致。
(python-dotenv 默认不覆盖已存在的环境变量，所以这里的值优先于 .env)
"""
import os
import sys

_HERMETIC_ENV = {
    "DEEPSEEK_API_KEY": "",
    "VISION_API_KEY": "",
    "GEMINI_API_KEY": "",
    "TYPESAFE_API_KEY": "",
    "CATEGORY_ATTRS_API_URL": "",
    "CATEGORY_CODE_FILE": "",
    "TEMPLATE_PATH": "",
    "ENABLE_MATERIAL_ILLUSTRATION": "0",
    "EXTRACTOR_ENGINE": "traditional",
    "STRUCTURE_ENGINE": "llm",
}
for _k, _v in _HERMETIC_ENV.items():
    os.environ[_k] = _v

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
