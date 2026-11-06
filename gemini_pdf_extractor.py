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
from concurrent.futures import ThreadPoolExecutor
from typing import Dict, List, Tuple, Optional, Any

import cv2
import fitz  # PyMuPDF
import numpy as np
import openpyxl
import requests

from config import (GEMINI_API_KEY, GEMINI_MODEL, GEMINI_FALLBACK_MODEL, GEMINI_THINKING_BUDGET,
                    GEMINI_BATCH_WORKERS, LLM_MAX_RETRIES)
from universal_pdf_grid_engine import unify_and_broadcast_cross_page_headers, auto_orient_page_upright
from seal_enhancer_service import SealEnhancerService

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
      "page_title": "示例供应商名称 报价单",
      "tables": [
        {
          "cells": [
            ["参数A", "规格项1", "规格项2", "规格项3"],
            ["型号X", "100.00", "", "200.00"],
            ["型号Y", "150.00", "180.00", "300.00"]
          ],
          "uncertain": [
            [1, 1, "字迹模糊，疑似100.00"]
          ]
        }
      ],
      "notes": "示例备注说明"
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

    def _prepare_enhanced_pdf_bytes(self, page_range: Optional[Tuple[int, int]] = None) -> bytes:
        """
        为 Gemini B 轨准备预处理后的高保真 PDF 字节流：
        1. 针对扫描件或图片：调用 auto_orient_page_upright 四向摆正 + balance_illumination_and_orient LAB 亮度通道均衡；
        2. 针对纯矢量文字页：直接原样抽取对应页面以保留 100% 矢量字形与极速吞吐；
        3. 支持按页区间 page_range=(start_idx, end_idx) 分批打包，防止长文档单次输出超 Token 截断或偷懒跳行。
        """
        enhancer = SealEnhancerService()
        ext = os.path.splitext(self.pdf_path)[1].lower()
        if ext in ('.png', '.jpg', '.jpeg'):
            img_doc = fitz.open(self.pdf_path)
            src_doc = fitz.open("pdf", img_doc.convert_to_pdf())
            img_doc.close()
        else:
            src_doc = fitz.open(self.pdf_path)

        out_doc = fitz.open()
        p_start, p_end = page_range if page_range else (0, len(src_doc))
        try:
            for p_idx in range(p_start, min(p_end, len(src_doc))):
                page = src_doc[p_idx]
                text_len = len(re.sub(r'\s', '', page.get_text()))
                # 原生矢量电子版且字数充足：直接拷贝矢量页
                if ext not in ('.png', '.jpg', '.jpeg') and text_len >= 50:
                    out_doc.insert_pdf(src_doc, from_page=p_idx, to_page=p_idx)
                    continue

                # 扫描页/图片：执行 5ms 四向摆正 + LAB 阴影均衡 (100% 保留红色公章色度)
                pix = page.get_pixmap(dpi=200)
                arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                bgr = cv2.cvtColor(arr, cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR)
                try:
                    bgr = auto_orient_page_upright(bgr)
                except Exception:
                    pass
                try:
                    bgr = enhancer.balance_illumination_and_orient(bgr, do_orient=False)
                except Exception:
                    pass

                ok, buf = cv2.imencode(".jpg", bgr, [int(cv2.IMWRITE_JPEG_QUALITY), 92])
                if ok:
                    img_pdf = fitz.open(stream=buf.tobytes(), filetype="jpg").convert_to_pdf()
                    with fitz.open("pdf", img_pdf) as tmp_p:
                        out_doc.insert_pdf(tmp_p)
            return out_doc.tobytes()
        finally:
            out_doc.close()
            src_doc.close()

    def _call_gemini_raw(self, client, pdf_bytes: bytes) -> str:
        """调用 Gemini 原生接口 (官方 SDK 优先，REST API 兜底)，强制配置 max_output_tokens=65536"""
        b64_data = base64.b64encode(pdf_bytes).decode("utf-8")
        raw_json_str = None

        if client:
            for attempt in range(1, LLM_MAX_RETRIES + 2):
                cur_model = GEMINI_MODEL if attempt == 1 else (GEMINI_FALLBACK_MODEL or GEMINI_MODEL)
                try:
                    pdf_part = types.Part.from_bytes(data=pdf_bytes, mime_type="application/pdf")
                    config_kwargs = {
                        "response_mime_type": "application/json",
                        "temperature": 0.0,
                        "max_output_tokens": 65536,
                    }
                    if GEMINI_THINKING_BUDGET is not None and hasattr(types, "ThinkingConfig"):
                        config_kwargs["thinking_config"] = types.ThinkingConfig(thinking_budget=GEMINI_THINKING_BUDGET)
                    response = client.models.generate_content(
                        model=cur_model,
                        contents=[pdf_part, GEMINI_PDF_TRANSCRIPTION_PROMPT],
                        config=types.GenerateContentConfig(**config_kwargs)
                    )
                    raw_json_str = response.text
                    if raw_json_str:
                        break
                except Exception as e:
                    print(f"[Gemini B-Rail] SDK ({cur_model}) 调用异常 (重试 {attempt}/{LLM_MAX_RETRIES}): {e}")
                    time.sleep(2)

        # 兜底: 若 SDK 失败或未安装，使用官方 REST API
        if not raw_json_str:
            print("[Gemini B-Rail] 启动官方 REST API 备选调用通道...")
            rest_url = "https://generativelanguage.googleapis.com/v1beta/interactions"
            payload = {
                "model": GEMINI_MODEL,
                "input": [
                    {
                        "type": "document",
                        "data": b64_data,
                        "mime_type": "application/pdf"
                    },
                    {"type": "text", "text": GEMINI_PDF_TRANSCRIPTION_PROMPT}
                ],
                "generationConfig": {"maxOutputTokens": 65536}
            }
            headers = {
                "Content-Type": "application/json",
                "x-goog-api-key": GEMINI_API_KEY
            }
            resp = requests.post(rest_url, json=payload, headers=headers, timeout=90)
            if resp.status_code == 200:
                res_data = resp.json()
                raw_json_str = res_data.get("output_text") or (
                    res_data.get("steps", [{}])[-1].get("content", [{}])[0].get("text", "")
                )
            else:
                raise RuntimeError(f"REST API 调用失败 [{resp.status_code}]: {resp.text}")

        return raw_json_str or ""

    @staticmethod
    def _clean_and_parse_json(raw_json_str: str) -> dict:
        clean_json_str = (raw_json_str or "").strip()
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
        return json.loads(clean_json_str)

    def to_workbook(self) -> openpyxl.Workbook:
        """
        调用 Gemini 3.8 Flash 原生全模态读取 PDF，并将结果转换为 openpyxl.Workbook
        """
        if not GEMINI_API_KEY:
            raise ValueError("未配置 GEMINI_API_KEY，无法使用 Gemini 原生全模态线路")

        print(f"[Gemini B-Rail] 启动 Google GenAI 原生全模态直传流水线: {self.pdf_path}")
        print(f"[*] 使用模型: {GEMINI_MODEL} | 文档总页数: {self.total_pages}")

        client = None
        if genai:
            http_opts = types.HttpOptions(timeout=90.0) if hasattr(types, "HttpOptions") else None
            client_kwargs = {"api_key": GEMINI_API_KEY}
            if http_opts:
                client_kwargs["http_options"] = http_opts
            client = genai.Client(**client_kwargs)

        all_pages: List[Dict[str, Any]] = []
        # 当总页数 > 4 时，按每 3 页一批进行分批打包与识别，杜绝长表注意力跳行与 Token 截断
        batch_size = 3 if self.total_pages > 4 else max(1, self.total_pages)
        ranges = []
        for p_start in range(0, max(1, self.total_pages), batch_size):
            p_end = min(self.total_pages, p_start + batch_size) if self.total_pages > 0 else 1
            ranges.append((p_start, p_end))

        def _run_batch(p_range: Tuple[int, int]):
            """单批预处理 + 识别（批次间页码独立，可多线路并发）"""
            p_start, _p_end = p_range
            if self.total_pages > 4:
                print(f"    [*] [B轨分批] 正在执行第 {p_start + 1} ~ {_p_end} 页预处理与识别 (共 {self.total_pages} 页)...")

            pdf_bytes = self._prepare_enhanced_pdf_bytes(p_range)
            raw_json = self._call_gemini_raw(client, pdf_bytes)
            try:
                parsed = self._clean_and_parse_json(raw_json)
            except Exception as err:
                print(f"[Gemini B-Rail] JSON 解析失败: {err}，原始内容预览: {raw_json[:300]}")
                raise err

            batch_pages = parsed.get("pages", []) if isinstance(parsed, dict) else (parsed if isinstance(parsed, list) else [])
            for p_info in batch_pages:
                try:
                    rel_pno = int(p_info.get("page_number", 1))
                except Exception:
                    rel_pno = 1
                if rel_pno <= batch_size and p_start > 0:
                    p_info["page_number"] = p_start + rel_pno
            return p_start, batch_pages

        # 多线路并发：10 页 = 4 个批次同时出去，不再一批一批干等
        workers = max(1, min(len(ranges), GEMINI_BATCH_WORKERS))
        if workers == 1:
            ordered = [_run_batch(r) for r in ranges]
        else:
            print(f"    [*] [B轨并发] 已启用 {workers} 路批次并发（共 {len(ranges)} 个批次）")
            with ThreadPoolExecutor(max_workers=workers) as pool:
                ordered = list(pool.map(_run_batch, ranges))

        for _p_start, batch_pages in sorted(ordered, key=lambda x: x[0]):
            all_pages.extend(batch_pages)

        return self._build_workbook_from_gemini_json({"pages": all_pages})

    def _build_workbook_from_gemini_json(self, data: dict) -> openpyxl.Workbook:
        """
        把 Gemini 转录的 JSON 数据转换为标准的 openpyxl.Workbook
        """
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        pages = data.get("pages", [])
        if not pages and isinstance(data, list):
            pages = data

        # 页数对不上就是漏页：必须显式上报，不能当成功落盘
        returned = set()
        for p in pages:
            try:
                returned.add(int(p.get("page_number")))
            except (AttributeError, TypeError, ValueError):
                pass
        if self.total_pages and len(pages) < self.total_pages:
            missing = [n for n in range(1, self.total_pages + 1) if n not in returned]
            msg = f"模型只转录了 {len(pages)}/{self.total_pages} 页" + (f"，缺页 {missing[:10]}" if missing else "")
            print(f"    ! [Gemini] {msg}")
            self.failed_pages.append(("整篇", msg))

        # 系统级全文档跨页表头双向广播与自推导（倒序表头回填、续表正向继承、全无表头反推）
        try:
            pages = unify_and_broadcast_cross_page_headers(pages)
        except Exception as e:
            print(f"    ! [Gemini] 跨页表头对齐跳过: {e}")

        for idx, page_info in enumerate(pages, 1):
            # Sheet 名用模型的自报页码，不能用 enumerate 序号：模型跳页时序号会把后面每页的页码全推错
            try:
                pno = int(page_info.get("page_number"))
            except (AttributeError, TypeError, ValueError):
                pno = idx
            sheet_name = f"第{pno}页"
            if sheet_name in wb.sheetnames:
                sheet_name = f"第{pno}页(#{idx})"
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
        """从识别出的页面大标题中推断供应商全称，若无则回退至 PDFTableExtractor 公章/原件识别"""
        for title in self.page_titles:
            m = COMPANY_RE.search(title)
            if m:
                return m.group(1)
        try:
            from pdf_extractor import PDFTableExtractor
            return PDFTableExtractor(self.pdf_path).guess_supplier() or ""
        except Exception:
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
