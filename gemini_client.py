# -*- coding: utf-8 -*-
"""
gemini_client.py - 企业级 Google Gemini 统一调用客户端
支持:
1. 首选 gemini-3.8-flash，遇 503/拥堵/超时 毫秒级无缝降级至 gemini-3.6-flash / gemini-3.5-flash
2. 极速 REST API + SDK 双轨容灾
3. JSON 格式强制约束与鲁棒解析
"""
import json
import os
import re
import time
from typing import Any, Dict, List, Optional

import requests

from config import (
    GEMINI_API_KEY,
    GEMINI_MODEL,
    GEMINI_FALLBACK_MODEL,
    GEMINI_THINKING_BUDGET,
    LLM_MAX_RETRIES,
    LLM_TEMPERATURE
)

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None


class GeminiClient:
    """统一 Gemini 调用中枢，保障高可用与极速响应"""

    FALLBACK_MODELS = []

    _quota_exhausted = False

    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        self.api_key = api_key or GEMINI_API_KEY
        self.primary_model = model or GEMINI_MODEL or "gemini-3.8-flash"
        self.models_priority = [self.primary_model]
        for m in self.FALLBACK_MODELS:
            if m not in self.models_priority:
                self.models_priority.append(m)

    def generate_json(self, prompt: str, system_prompt: Optional[str] = None,
                      timeout: int = 40) -> Optional[Dict[str, Any]]:
        """调用 Gemini 并返回解析后的 JSON 字典对象"""
        text = self.generate_text(prompt, system_prompt=system_prompt, is_json=True, timeout=timeout)
        if not text:
            return None
        return self._extract_json(text)

    def generate_text(self, prompt: str, system_prompt: Optional[str] = None,
                      is_json: bool = False, timeout: int = 40) -> Optional[str]:
        """
        统一调用文本生成：
        优先使用 REST API (无冗余 SDK 包装，连接快)，若发生异常自动切换降级模型
        """
        if not self.api_key or GeminiClient._quota_exhausted:
            return None

        # 遍历模型梯队重试
        for cur_model in self.models_priority:
            if GeminiClient._quota_exhausted:
                return None
            for attempt in range(1, LLM_MAX_RETRIES + 1):
                # 尝试 REST API
                result, should_fallback = self._call_rest_api(cur_model, prompt, system_prompt, is_json, timeout)
                if result:
                    return result
                if should_fallback or GeminiClient._quota_exhausted:
                    break  # 遇 503 等拥堵立即切换降级模型，不硬等
                time.sleep(1)

        # 若 REST 均未成功，尝试 SDK 兜底 (若已配额超限则不重复尝试)
        # 只兜底首选模型：REST 已经把全梯队试过一轮，SDK 再逐个重试只是把卡死时间乘以 5
        if genai and not GeminiClient._quota_exhausted:
            for cur_model in self.models_priority[:1]:
                try:
                    # 不给 http timeout，SDK 会无限等下去（主流程曾因此挂死 30 分钟）
                    client = genai.Client(api_key=self.api_key,
                                          http_options={"timeout": max(1, int(timeout)) * 1000})
                    config_args = {"temperature": LLM_TEMPERATURE}
                    if is_json and hasattr(types, "GenerateContentConfig"):
                        config_args["response_mime_type"] = "application/json"
                    if system_prompt:
                        config_args["system_instruction"] = system_prompt
                    if hasattr(types, "ThinkingConfig"):
                        config_args["thinking_config"] = types.ThinkingConfig(thinking_budget=GEMINI_THINKING_BUDGET)

                    cfg = types.GenerateContentConfig(**config_args) if hasattr(types, "GenerateContentConfig") else None
                    resp = client.models.generate_content(
                        model=cur_model,
                        contents=prompt,
                        config=cfg
                    )
                    if resp and resp.text:
                        return resp.text
                except Exception as sdk_err:
                    print(f"[GeminiClient SDK Fallback] 模型 {cur_model} SDK 调用异常: {sdk_err}")

        return None

    def _call_rest_api(self, model: str, prompt: str, system_prompt: Optional[str],
                       is_json: bool, timeout: int) -> tuple[Optional[str], bool]:
        """返回 (result_text, should_fallback)"""
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={self.api_key}"
        
        contents = []
        if prompt:
            contents.append({"parts": [{"text": prompt}]})

        payload: Dict[str, Any] = {
            "contents": contents,
            "generationConfig": {
                "temperature": LLM_TEMPERATURE
            }
        }
        if is_json:
            payload["generationConfig"]["responseMimeType"] = "application/json"
        if system_prompt:
            payload["systemInstruction"] = {
                "parts": [{"text": system_prompt}]
            }

        try:
            r = requests.post(url, json=payload, timeout=timeout)
            if r.status_code == 200:
                res_data = r.json()
                candidates = res_data.get("candidates", [])
                if candidates:
                    parts = candidates[0].get("content", {}).get("parts", [])
                    texts = [p.get("text", "") for p in parts if "text" in p]
                    return "".join(texts), False
            elif r.status_code in (429, 503):
                if r.status_code == 429:
                    GeminiClient._quota_exhausted = True
                    print(f"[GeminiClient REST] 检测到 429 配额用尽，启动熔断保护，秒级直通 DeepSeek/规则降级通道...")
                    return None, True
                print(f"[GeminiClient REST] 模型 {model} 遭遇并发高峰 503，无缝切换备选模型...")
                return None, True
            else:
                print(f"[GeminiClient REST] 模型 {model} 响应 [{r.status_code}]: {r.text[:150]}")
                return None, False
        except requests.exceptions.Timeout:
            GeminiClient._quota_exhausted = True
            print(f"[GeminiClient REST] 模型 {model} 请求超时 ({timeout}s)，立即熔断保护并秒级直通备用通道...")
            return None, True
        except Exception as e:
            GeminiClient._quota_exhausted = True
            print(f"[GeminiClient REST Error] 模型 {model} 请求异常: {e}，立即熔断...")
            return None, True

    @staticmethod
    def _extract_json(text: str) -> Optional[Dict[str, Any]]:
        clean_text = text.strip()
        m_block = re.search(r'```(?:json)?\s*(\{.*\}|\[.*\])\s*```', clean_text, re.DOTALL)
        if m_block:
            clean_text = m_block.group(1).strip()
        elif clean_text.startswith("```"):
            clean_text = re.sub(r"^```(?:json)?\s*", "", clean_text)
            clean_text = re.sub(r"\s*```$", "", clean_text).strip()
        else:
            m_brace = re.search(r'(\{.*\})', clean_text, re.DOTALL)
            if m_brace:
                clean_text = m_brace.group(1).strip()

        try:
            return json.loads(clean_text)
        except Exception as e:
            print(f"[GeminiClient] JSON 解析异常: {e} | 内容: {clean_text[:200]}")
            return None
