# -*- coding: utf-8 -*-
"""
Gemini 原生全模态 PDF 转录提取器 (B 轨新流水线)
严格遵循 Google 官方文档规范 (google-genai / Gemini 3.8 Flash)：
1. 小文件支持 Base64 Inline 传输 (interactions.create / models.generate_content)
2. 大文件支持 Files API 上传 (files.upload + file.uri)
3. 原生视觉全模态一眼看懂多达 1000 页的整篇文档、表格、公式、图表
4. 深度防御: 内置 SDK 调用 + REST API 双重兜底，输出标准兼容的 openpyxl.Workbook
"""
import base64
import json
import os
import re
import time
from typing import Dict, List, Tuple

import fitz  # PyMuPDF
import openpyxl
import requests

from config import GEMINI_API_KEY, GEMINI_MODEL, LLM_MAX_RETRIES

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

COMPANY_RE = re.compile(r'([\u4e00-\u9fa5（）\(\)]{4,30}?(?:股份有限公司|有限责任公司|有限公司|股份公司))')

GEMINI_PDF_TRANSCRIPTION_PROMPT = """你是一个国家级建设工程造价与采招数据专家。请直接审视这份原始 PDF 报价单文档，把其中包含的全部页面的所有表格原样逐格转录为 JSON 结构。

【转录铁律与业务规范】：
1. 完整还原每一页的表格网格，绝不遗漏任何一页、任何一行、任何一列。
2. 对于每一张表格，按真实网格输出二维数组 cells：
   - 每一行的列数必须完全相等，对齐最细的列划分；
   - 合并单元格：把同一个值重复填写到它覆盖的每一个格子（横向、纵向都要填充，严禁留空）；
   - 原文空白格填 ""；
   - 数字与符号必须 100% 逐位照抄（严禁四舍五入、严禁擅自补零或篡改，如“10500”、“136,000”、“m²”一律保留）；
3. 【印章遮挡与模糊存疑显式标注】：
   - 凡是被公章（红色印章）、折痕、污损、噪点遮挡或模糊无法 100% 确认的格子：仍填写你最可能的读数，并必须在 uncertain 中记录 [行号, 列号, 存疑原因]（行号与列号均从 0 开始计）；
4. 表格外的大标题（如公司名称、报价单表名）放入 page_title，页脚备注或说明放入 notes。

【输出格式要求】：
严格只返回一个 JSON 对象，禁止输出任何 Markdown 说明或冗余文本：
{
  "pages": [
    {
      "page_number": 1,
      "page_title": "重庆望变电气（集团）股份有限公司 报价单",
      "tables": [
        {
          "cells": [
            ["容量kVA", "S13型（NX3）", "S20型（NX2）", "S22型（NX1）"],
            ["30", "10500", "", "20800"],
            ["50", "13500", "19800", "33000"]
          ],
          "uncertain": [
            [1, 1, "公章文字遮挡数字，疑似10500"]
          ]
        }
      ],
      "notes": "1000公里以内含运费"
    }
  ]
}
"""


class GeminiPDFTableExtractor:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.page_titles: List[str] = []
        self.cell_issue: Dict[Tuple[str, int, int], str] = {}
        self.text_source: Dict[str, str] = {}
        self.failed_pages: List[Tuple[str, str]] = []
        self.total_pages = 0

        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF 文件不存在: {pdf_path}")

        try:
            with fitz.open(pdf_path) as doc:
                self.total_pages = len(doc)
        except Exception as e:
            print(f"[GeminiPDFTableExtractor] 打开 PDF 失败: {e}")

    def to_workbook(self) -> openpyxl.Workbook:
        """
        调用 Gemini 3.8 Flash 原生全模态读取 PDF，并将结果转换为 openpyxl.Workbook
        """
        if not GEMINI_API_KEY:
            raise ValueError("未配置 GEMINI_API_KEY，无法使用 Gemini 原生全模态线路")

        print(f"[Gemini B-Rail] 启动 Google GenAI 原生全模态直传流水线: {self.pdf_path}")
        print(f"[*] 使用模型: {GEMINI_MODEL} | 文档总页数: {self.total_pages}")

        file_size = os.path.getsize(self.pdf_path)
        raw_json_str = None
        start_t = time.time()

        # 优先使用 google-genai 官方 SDK
        if genai:
            client = genai.Client(api_key=GEMINI_API_KEY)
            uploaded_file = None

            try:
                # 方式 1: < 20MB 文件使用 Base64 Inline 方式传输 (官方文档推荐示例)
                if file_size < 20 * 1024 * 1024:
                    with open(self.pdf_path, "rb") as f:
                        pdf_bytes = f.read()
                    b64_data = base64.b64encode(pdf_bytes).decode("utf-8")

                    for attempt in range(1, LLM_MAX_RETRIES + 2):
                        try:
                            # 优先尝试官方文档的 interactions.create 接口
                            if hasattr(client, "interactions"):
                                interaction = client.interactions.create(
                                    model=GEMINI_MODEL,
                                    input=[
                                        {
                                            "type": "document",
                                            "data": b64_data,
                                            "mime_type": "application/pdf"
                                        },
                                        {"type": "text", "text": GEMINI_PDF_TRANSCRIPTION_PROMPT}
                                    ]
                                )
                                raw_json_str = getattr(interaction, "output_text", None) or str(interaction)
                            else:
                                # 回退到 models.generate_content
                                pdf_part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")
                                response = client.models.generate_content(
                                    model=GEMINI_MODEL,
                                    contents=[pdf_part, GEMINI_PDF_TRANSCRIPTION_PROMPT],
                                    config=types.GenerateContentConfig(
                                        response_mime_type="application/json",
                                        temperature=0.1
                                    )
                                )
                                raw_json_str = response.text
                            if raw_json_str:
                                break
                        except Exception as e:
                            print(f"[Gemini B-Rail] SDK 调用异常 (重试 {attempt}/{LLM_MAX_RETRIES}): {e}")
                            time.sleep(2)
                else:
                    # 方式 2: >= 20MB 文件使用 Files API 上传
                    print(f"[*] 文件大小 {file_size / (1024*1024):.1f}MB，使用 Google Files API 上传...")
                    uploaded_file = client.files.upload(file=self.pdf_path)
                    
                    for attempt in range(1, LLM_MAX_RETRIES + 2):
                        try:
                            if hasattr(client, "interactions"):
                                interaction = client.interactions.create(
                                    model=GEMINI_MODEL,
                                    input=[
                                        {
                                            "type": "document",
                                            "uri": uploaded_file.uri,
                                            "mime_type": uploaded_file.mime_type or "application/pdf"
                                        },
                                        {"type": "text", "text": GEMINI_PDF_TRANSCRIPTION_PROMPT}
                                    ]
                                )
                                raw_json_str = getattr(interaction, "output_text", None) or str(interaction)
                            else:
                                response = client.models.generate_content(
                                    model=GEMINI_MODEL,
                                    contents=[uploaded_file, GEMINI_PDF_TRANSCRIPTION_PROMPT],
                                    config=types.GenerateContentConfig(
                                        response_mime_type="application/json",
                                        temperature=0.1
                                    )
                                )
                                raw_json_str = response.text
                            if raw_json_str:
                                break
                        except Exception as e:
                            print(f"[Gemini B-Rail] Files API 调用异常 (重试 {attempt}/{LLM_MAX_RETRIES}): {e}")
                            time.sleep(2)
            finally:
                if uploaded_file:
                    try:
                        client.files.delete(name=uploaded_file.name)
                    except Exception:
                        pass

        # 兜底: 若 SDK 失败或未安装，使用官方 REST API (https://generativelanguage.googleapis.com/v1beta/interactions)
        if not raw_json_str:
            print("[Gemini B-Rail] 启动官方 REST API 备选调用通道...")
            with open(self.pdf_path, "rb") as f:
                b64_pdf = base64.b64encode(f.read()).decode("utf-8")

            rest_url = f"https://generativelanguage.googleapis.com/v1beta/interactions?key={GEMINI_API_KEY}"
            payload = {
                "model": GEMINI_MODEL,
                "input": [
                    {
                        "type": "document",
                        "data": b64_pdf,
                        "mime_type": "application/pdf"
                    },
                    {"type": "text", "text": GEMINI_PDF_TRANSCRIPTION_PROMPT}
                ]
            }
            resp = requests.post(rest_url, json=payload, headers={"Content-Type": "application/json"}, timeout=120)
            if resp.status_code == 200:
                res_data = resp.json()
                raw_json_str = res_data.get("output_text") or (
                    res_data.get("steps", [{}])[-1].get("content", [{}])[0].get("text", "")
                )
            else:
                raise RuntimeError(f"REST API 调用失败 [{resp.status_code}]: {resp.text}")

        if not raw_json_str:
            raise RuntimeError("Gemini 返回内容为空")

        elapsed = time.time() - start_t
        print(f"[Gemini B-Rail] 原生全模态转录完成，总耗时: {elapsed:.2f}s")

        # 解析与格式化
        clean_json_str = raw_json_str.strip()
        m_block = re.search(r'```(?:json)?\s*(\{.*\}|\[.*\])\s*```', clean_json_str, re.DOTALL)
        if m_block:
            clean_json_str = m_block.group(1).strip()
        elif clean_json_str.startswith("```"):
            clean_json_str = re.sub(r"^```(?:json)?\s*", "", clean_json_str)
            clean_json_str = re.sub(r"\s*```$", "", clean_json_str).strip()
        else:
            m_brace = re.search(r'(\{.*\})', clean_json_str, re.DOTALL)
            if m_brace:
                clean_json_str = m_brace.group(1).strip()

        try:
            parsed = json.loads(clean_json_str)
        except Exception as err:
            print(f"[Gemini B-Rail] JSON 解析失败: {err}，原始内容预览: {clean_json_str[:300]}")
            raise err
        return self._build_workbook_from_gemini_json(parsed)

    def _build_workbook_from_gemini_json(self, data: dict) -> openpyxl.Workbook:
        """
        把 Gemini 转录的 JSON 数据转换为标准的 openpyxl.Workbook
        """
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        pages = data.get("pages", [])
        if not pages and isinstance(data, list):
            pages = data

        for idx, page_info in enumerate(pages, 1):
            sheet_name = f"第{idx}页"
            ws = wb.create_sheet(title=sheet_name)
            self.text_source[sheet_name] = f"gemini_{GEMINI_MODEL}"

            page_title = page_info.get("page_title", "")
            if page_title:
                self.page_titles.append(page_title)

            tables = page_info.get("tables", [])
            current_r = 1

            # 若页面有大标题，写在首行作为页面标识存根
            if page_title and (not tables or not any(page_title in str(r) for r in tables[0].get("cells", []))):
                ws.cell(row=current_r, column=1, value=page_title)
                current_r += 1

            for tbl_idx, tbl in enumerate(tables):
                cells = tbl.get("cells", [])
                uncertain = tbl.get("uncertain", [])

                if not cells:
                    continue

                table_start_row = current_r
                for r_idx, row_vals in enumerate(cells):
                    for c_idx, val in enumerate(row_vals):
                        cell_val = str(val).strip() if val is not None else ""
                        ws.cell(row=table_start_row + r_idx, column=c_idx + 1, value=cell_val)

                # 记录存疑单元格
                for u in uncertain:
                    if len(u) >= 3:
                        u_r, u_c, reason = u[0], u[1], u[2]
                        actual_r = table_start_row + int(u_r)
                        actual_c = int(u_c) + 1
                        self.cell_issue[(sheet_name, actual_r, actual_c)] = f"({reason})"

                current_r = table_start_row + len(cells) + 1

            notes = page_info.get("notes", "")
            if notes:
                ws.cell(row=current_r, column=1, value=f"【说明】{notes}")

        return wb

    def guess_supplier(self) -> str:
        """从识别出的页面大标题中推断供应商全称"""
        for title in self.page_titles:
            m = COMPANY_RE.search(title)
            if m:
                return m.group(1)
        return ""

    def save_page_images(self, out_dir: str) -> Dict[int, str]:
        """
        利用 PyMuPDF 导出每一页的高清原件截图存根，供构建器做佐证留痕
        """
        os.makedirs(out_dir, exist_ok=True)
        saved = {}
        try:
            with fitz.open(self.pdf_path) as doc:
                for idx, page in enumerate(doc, 1):
                    pix = page.get_pixmap(dpi=200)
                    fn = f"quote_img{idx}.png"
                    dst = os.path.join(out_dir, fn)
                    pix.save(dst)
                    saved[idx] = fn
        except Exception as e:
            print(f"[GeminiPDFTableExtractor] 保存原件截图失败: {e}")
        return saved
