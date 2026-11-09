# -*- coding: utf-8 -*-
"""
PDF 报价单 → 表格网格还原器
把 PDF 每一页还原成一个二维表格 (openpyxl Worksheet)，之后与 Excel 走完全相同的解析/清洗/对账流程：
1. 有文字层的电子版：PyMuPDF find_tables 直接取表格；没有表格时按文本行切分
2. 扫描件：视觉大模型 (OpenAI 兼容接口) 直读页面图片，返回完整二维表格
   - 第一遍读增强图 (去红章、压暗褪色笔画)，第二遍读原图，两遍不一致的格子标记待复核
   - 模型自报看不清 / 被公章遮挡的格子同样标记
   - 视觉模型不可用时直接报错，不产出不可信的识别结果
3. 每个存疑单元格记录原因 (cell_issue)，供下游写入待复核说明
"""
import base64
import json
import os
import re
import threading
import time
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from typing import Optional, List, Dict, Any, Tuple

import cv2
import fitz
import numpy as np
import openpyxl
import requests

from config import (VISION_API_KEY, VISION_BASE_URL, VISION_MODEL, VISION_DOUBLE_CHECK,
                    VISION_GAMMA, VISION_MAX_ROWS_PER_TILE, VISION_TILE_OVERLAP_ROWS,
                    VISION_MAX_WORKERS, LLM_MAX_RETRIES)
from seal_enhancer_service import SealEnhancerService, surgical_seal_mask_repair
from universal_pdf_grid_engine import (
    auto_orient_page_upright,
    detect_physical_table_row_bands,
    slice_dense_page_with_row_rulers,
    is_valid_table_header_row,
    broadcast_headers_to_page_grids,
)

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

_SEAL_ENHANCER = SealEnhancerService()

COMPANY_RE = re.compile(r'([\u4e00-\u9fa5（）\(\)]{4,30}?(?:股份有限公司|有限责任公司|有限公司|股份公司))')
SCAN_DPI = 300
VISION_DPI = 200
VISION_MAX_EDGE = 2400

VISION_TABLE_PROMPT = """你是报价单扫描件的表格转录员。请把图片中的每一张表格逐格、原样转录为 JSON，不做任何理解、计算或改写。

【转录规则】
1. 按表格的真实网格输出二维数组 cells：每一行的列数必须相同，等于该表最细的列划分数；
2. 合并单元格：把同一个值重复填写到它覆盖的每一个格子 (横向、纵向都要填)；
3. 空白格填 ""；
4. 数字必须逐位照抄 (不得四舍五入、不得补零、不得删除千分位以外的字符)；单元格里有换行的，用一个空格连接；
   文字逐字照抄，原文的用字、异体字、符号 (如“曝露”“m²”“；”) 一律保留，不得替换为你认为更规范的写法；
5. 被公章遮挡、模糊、无法确定的格子：仍填写你最可能的读数，并在 uncertain 中列出 [行号, 列号, 原因] (行列号从 0 开始)；
6. 表格外的标题文字 (如公司名、表名) 放到 page_title，页脚备注放到 notes；
7. 一页有多张表格时，按从上到下、从左到右顺序放入 tables。

【返回格式】只返回 JSON 对象：
{"page_title": "...", "tables": [{"cells": [["表头1", "表头2"], ["值", "值"]], "uncertain": [[1, 1, "公章遮挡"]]}], "notes": "..."}
"""


def validate_vision_payload(obj):
    """
    视觉模型响应的结构硬约束：不合规整份作废，退回重试，不产出半可信的表格。
    返回 '' 表示通过，否则返回作废原因。
    """
    if not isinstance(obj, dict):
        return '不是 JSON 对象'
    tables = obj.get('tables')
    if not isinstance(tables, list) or not tables:
        return '缺少 tables 或为空'
    for t in tables:
        cells = t.get('cells') if isinstance(t, dict) else None
        if not isinstance(cells, list) or not cells:
            return 'tables[].cells 缺失或为空'
        if any(not isinstance(r, list) for r in cells):
            return 'cells 内含非数组行'
        widths = {len(r) for r in cells}
        if len(widths) != 1 or 0 in widths:
            return f'cells 列数不齐 {sorted(widths)}'
        unc = t.get('uncertain', [])
        if not isinstance(unc, list) or any(not isinstance(u, list) or len(u) < 2 for u in unc):
            return 'uncertain 格式错误'
    for k in ('page_title', 'notes'):
        if obj.get(k) is not None and not isinstance(obj.get(k), str):
            return f'{k} 不是字符串'
    return ''


class PDFTableExtractor:
    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        ext = os.path.splitext(pdf_path)[1].lower()
        if ext in ('.png', '.jpg', '.jpeg'):
            # 单张图片转换为单页内存 PDF 文档
            img_doc = fitz.open(pdf_path)
            pdf_bytes = img_doc.convert_to_pdf()
            self.doc = fitz.open("pdf", pdf_bytes)
        else:
            self.doc = fitz.open(pdf_path)
        self.cell_issue = {}      # (sheet, row, col) -> 存疑原因
        self.page_titles = []     # 每页表格上方的文字 (用于识别供应商)
        self.text_source = {}     # sheet -> 'text' / 'vision'

    # ------------------------------------------------------------------ 入口
    def to_workbook(self, max_pages=None):
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        total_p = len(self.doc) if max_pages is None else min(max_pages, len(self.doc))
        use_vision = bool(VISION_API_KEY and VISION_BASE_URL and VISION_MODEL)
        print(f"[PDFTableExtractor] 正在还原 {total_p} 页表格 (扫描件识别: {VISION_MODEL if use_vision else '未配置视觉模型'})...")
        
        def _extract_one(p_idx):
            """单页还原（线程内各开一份 Document，页面之间互不共享状态）"""
            page = self._thread_page(p_idx)
            base = f"第{p_idx + 1}页"
            self._tl.title_sink = []
            if len(re.sub(r'\s', '', page.get_text())) >= 30:
                return {"base": base, "grids": self._grids_from_text_layer(page), "src": 'text'}
            if not use_vision:
                return {"base": base, "error": "no_vision"}
            titles = self._tl.title_sink
            try:
                grids = self._grids_from_vision(page, base)
                return {"base": base, "grids": grids, "src": "vision", "titles": titles}
            except Exception as page_err:
                return {"base": base, "error": str(page_err)}

        # 多线路并发：各页彼此独立，串行等 API 纯属浪费（10 页从 10×T 降到 ~T）
        workers = max(1, min(total_p, VISION_MAX_WORKERS))
        if workers == 1:
            results = [_extract_one(i) for i in range(total_p)]
        else:
            print(f"    [*] 已启用 {workers} 路线程并发页面识别")
            with ThreadPoolExecutor(max_workers=workers) as pool:
                results = list(pool.map(_extract_one, range(total_p)))

        pages_records = []
        for rec in results:                     # pool.map 保序，严格按页序汇总
            base = rec["base"]
            if rec.get("error") == "no_vision":
                raise RuntimeError(f"{base} 是扫描件，需要可用的视觉模型 (DEEPSEEK_* 或 VISION_*)")
            if rec.get("error"):
                print(f"    ! [隔离] {base} 识别失败: {rec['error']}")
                if not hasattr(self, 'failed_pages'):
                    self.failed_pages = []
                self.failed_pages.append((base, rec["error"]))
                # 失败页占位保留在页序里，稍后按原位置建异常 Sheet（不执行中断，确保后续页面继续处理并如实对账）
                pages_records.append({"base": base, "error": rec["error"]})
                continue
            self.page_titles.extend(rec.get("titles") or [])
            pages_records.append({
                "base": base,
                "grids": rec["grids"],
                "src": rec["src"],
            })

        # 全文档跨页表头双向广播与自推导（支持第10页表头倒序回填第1~9页、续表正向继承、全无表头列角色推导）
        pages_records = broadcast_headers_to_page_grids(pages_records)

        # 写入 Workbook
        for rec in pages_records:
            base = rec["base"]
            if rec.get("error"):
                ws_err = wb.create_sheet(base)
                ws_err.cell(1, 1, value=f"异常:{base}识别失败({rec['error']})")
                self.text_source[base] = 'error'
                continue
            grids = rec["grids"]
            src = rec["src"]
            for t_idx, grid in enumerate(grids, 1):
                # 一页多表时拆成 “第N页-表K”，存疑单元格记录的 key 同步改为最终 Sheet 名
                sname = base if len(grids) == 1 else f"{base}-表{t_idx}"
                for key in [k for k in self.cell_issue if k[0] == f"{base}#{t_idx}"]:
                    self.cell_issue[(sname,) + key[1:]] = self.cell_issue.pop(key)
                ws = wb.create_sheet(sname)
                self.text_source[sname] = src
                for r, row in enumerate(grid, 1):
                    for c, v in enumerate(row, 1):
                        if v not in (None, ''):
                            ws.cell(r, c, value=str(v))
                label = {'text': '文字层', 'vision': '视觉模型'}[src]
                print(f"    -> {sname}: {label} 还原 {len(grid)} 行 x {max((len(r) for r in grid), default=0)} 列")
        return wb

    # ------------------------------------------------------------------ 视觉模型与硬门禁
    @staticmethod
    def _verify_against_rapidocr(img_bgr: Optional[np.ndarray], vision_dict: dict) -> Tuple[bool, float, str]:
        """
        【硬门禁 1】：模型输出汉字与本地 RapidOCR 重合度必须 >=60%，否则熔断判废
        """
        if not vision_dict or not isinstance(vision_dict, dict):
            return False, 0.0, "模型输出为空或非合规JSON"

        vision_chars = set()
        for t in vision_dict.get("tables", []):
            for row in t.get("cells", []):
                for cell in row:
                    for c in str(cell):
                        if '\u4e00' <= c <= '\u9fa5':
                            vision_chars.add(c)

        if len(vision_chars) < 5:
            return True, 1.0, "模型输出汉字少于5个，放行"

        if img_bgr is None or img_bgr.size == 0:
            return True, 1.0, "无原始像素数据，放行"

        try:
            from rapidocr_onnxruntime import RapidOCR
            ocr = RapidOCR()
            res, _ = ocr(img_bgr)
            ocr_chars = set()
            if res:
                for line in res:
                    if len(line) > 1:
                        for c in str(line[1]):
                            if '\u4e00' <= c <= '\u9fa5':
                                ocr_chars.add(c)

            if len(ocr_chars) < 5:
                return True, 1.0, "本地OCR未检测到足够汉字，放行"

            overlap_chars = vision_chars & ocr_chars
            overlap_ratio = len(overlap_chars) / len(vision_chars)

            if overlap_ratio < 0.60:
                return False, overlap_ratio, (
                    f"模型输出汉字与本地RapidOCR重合度仅为 {overlap_ratio:.1%} "
                    f"(<60% 熔断阈值，模型汉字数 {len(vision_chars)}，OCR汉字数 {len(ocr_chars)})"
                )

            return True, overlap_ratio, f"重合度 {overlap_ratio:.1%} (>=60% 门禁通过)"
        except Exception as e:
            print(f"    [!] RapidOCR门禁执行异常: {e}，安全放行")
            return True, 1.0, "门禁异常放行"

    def _page_png_b64(self, page, enhanced=True):
        """
        渲染页面给视觉模型。保留高清晰度原生彩色通道，自动执行姿态矫正与尺寸自适应；
        enhanced=True 时采用【顶层铁律 2：红章局部掩膜修复】，仅脱去红章覆盖区域印泥，
        其余 95% 表格区域（黑底白字表头、蓝底白字表头、细小数字等）100% 保持原图原貌！
        返回 (b64_str, bgr_img)。
        """
        dpi = SCAN_DPI if enhanced else VISION_DPI
        pix = page.get_pixmap(dpi=dpi)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        bgr = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR)
        try:
            bgr = auto_orient_page_upright(bgr)
        except Exception:
            pass
        if enhanced:
            try:
                bgr = surgical_seal_mask_repair(bgr)
            except Exception:
                pass
        out = bgr
        scale = VISION_MAX_EDGE / max(out.shape[:2])
        if scale < 1:
            out = cv2.resize(out, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
        b64 = base64.b64encode(cv2.imencode('.jpg', out, [int(cv2.IMWRITE_JPEG_QUALITY), 92])[1].tobytes()).decode()
        return b64, out

    def _execute_multimodal_vision(self, b64: str, prompt_text: str, kind: str,
                                   img_bgr: Optional[np.ndarray] = None) -> Optional[dict]:
        """统一多模态视觉中枢：支持 Gemini 原生 / OpenAI 兼容接口，严密挂接 RapidOCR 60% 熔断门禁与 CAS 缓存"""
        import cas_cache
        cache_key = b64 + '|' + prompt_text
        _cas = cas_cache.get(kind, VISION_MODEL or '', cache_key)
        if _cas:
            try:
                cached = json.loads(_cas)
                if not validate_vision_payload(cached):
                    return cached
            except Exception:
                pass

        parsed = None
        is_gemini = (VISION_MODEL.startswith("gemini-") or "generativelanguage" in (VISION_BASE_URL or ""))

        if is_gemini:
            # 1. 优先使用 Google GenAI 官方 SDK
            if genai and VISION_API_KEY:
                for attempt in range(min(LLM_MAX_RETRIES + 1, 2)):
                    try:
                        client = genai.Client(api_key=VISION_API_KEY)
                        img_bytes = base64.b64decode(b64)
                        img_part = types.Part.from_bytes(data=img_bytes, mime_type="image/jpeg")
                        cfg = types.GenerateContentConfig(
                            response_mime_type="application/json",
                            temperature=0.0,
                            max_output_tokens=16384
                        )
                        full_prompt = f"{VISION_TABLE_PROMPT}\n\n{prompt_text}"
                        resp = client.models.generate_content(
                            model=VISION_MODEL,
                            contents=[img_part, full_prompt],
                            config=cfg
                        )
                        if resp and resp.text:
                            parsed = json.loads(_strip_fence(resp.text))
                            break
                    except Exception as g_err:
                        if attempt == 0:
                            time.sleep(1)

            # 2. 若 SDK 失败或未安装，使用官方 REST API 备选
            if not parsed and VISION_API_KEY:
                try:
                    rest_url = f"https://generativelanguage.googleapis.com/v1beta/models/{VISION_MODEL}:generateContent?key={VISION_API_KEY}"
                    full_prompt = f"{VISION_TABLE_PROMPT}\n\n{prompt_text}"
                    payload = {
                        "contents": [{
                            "parts": [
                                {"inlineData": {"mimeType": "image/jpeg", "data": b64}},
                                {"text": full_prompt}
                            ]
                        }],
                        "generationConfig": {
                            "responseMimeType": "application/json",
                            "temperature": 0.0,
                            "maxOutputTokens": 16384
                        }
                    }
                    resp = requests.post(rest_url, json=payload, timeout=60)
                    if resp.status_code == 200:
                        res_data = resp.json()
                        candidates = res_data.get("candidates", [])
                        if candidates:
                            parts = candidates[0].get("content", {}).get("parts", [])
                            text_str = "".join(p.get("text", "") for p in parts if "text" in p)
                            parsed = json.loads(_strip_fence(text_str))
                except Exception as r_err:
                    print(f"[Gemini REST Vision Warning] {r_err}")
        else:
            # OpenAI 兼容多模态接口 (如 Qwen-VL 等)
            url = f"{VISION_BASE_URL.rstrip('/')}/chat/completions"
            headers = {"Authorization": f"Bearer {VISION_API_KEY}", "Content-Type": "application/json"}
            payload = {
                "model": VISION_MODEL,
                "temperature": 0,
                "messages": [
                    {"role": "system", "content": VISION_TABLE_PROMPT},
                    {"role": "user", "content": [
                        {"type": "text", "text": prompt_text},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{b64}"}},
                    ]},
                ],
                "response_format": {"type": "json_object"},
            }
            for attempt in range(min(LLM_MAX_RETRIES + 1, 2)):
                try:
                    resp = requests.post(url, headers=headers, json=payload, timeout=45)
                    if resp.status_code == 400 and "response_format" in payload:
                        payload.pop("response_format")
                        continue
                    if resp.status_code == 200:
                        choice = resp.json()["choices"][0]
                        if choice.get("finish_reason") == "length":
                            return {"_truncated": True}
                        parsed = json.loads(_strip_fence(choice["message"]["content"]))
                        break
                except Exception:
                    pass
                if attempt < 1:
                    time.sleep(1)

        if not parsed:
            return None

        if parsed.get("_truncated"):
            return parsed

        bad = validate_vision_payload(parsed)
        if bad:
            print(f"[Vision Warning] 响应结构不合规（{bad}），整份作废")
            return None

        # 🔒【硬门禁 1】：模型输出汉字与本地 RapidOCR 重合度必须 >=60%，否则熔断判废！
        if img_bgr is not None:
            ok, overlap, reason = self._verify_against_rapidocr(img_bgr, parsed)
            if not ok:
                print(f"    [!] 触发【硬门禁 1】熔断判废: {reason}！拒绝采纳并废除脏数据！")
                return None

        # 仅在通过硬门禁后写入 CAS 内容寻址缓存
        try:
            cas_cache.put(kind, VISION_MODEL or '', cache_key, json.dumps(parsed, ensure_ascii=False))
        except Exception:
            pass

        return parsed

    def _vision_call(self, b64_or_tuple):
        """调用统一多模态接口转录整页表格"""
        if isinstance(b64_or_tuple, tuple):
            b64, img_bgr = b64_or_tuple
        else:
            b64, img_bgr = b64_or_tuple, None
        return self._execute_multimodal_vision(b64, "请转录这一页中的全部表格，返回 JSON。",
                                              kind='vision_table', img_bgr=img_bgr)

    @staticmethod
    def _normalize_tables(result):
        tables = []
        for t in (result or {}).get("tables") or []:
            cells = t.get("cells") if isinstance(t, dict) else t
            if not isinstance(cells, list) or not cells:
                continue
            width = max(len(r) for r in cells if isinstance(r, list)) if any(isinstance(r, list) for r in cells) else 0
            grid = [[str(v).strip() if v is not None else '' for v in (r if isinstance(r, list) else [r])]
                    for r in cells]
            grid = [row + [''] * (width - len(row)) for row in grid]
            unc = [u for u in (t.get("uncertain") or []) if isinstance(t, dict) and isinstance(u, list) and len(u) >= 2]
            tables.append((grid, unc))
        return tables


    def _slice_page_tiles(self, page, enhanced=True, num_tiles=2):
        dpi = SCAN_DPI if enhanced else VISION_DPI
        pix = page.get_pixmap(dpi=dpi)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        bgr = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR)
        try:
            bgr = auto_orient_page_upright(bgr)
        except Exception:
            pass
        if enhanced:
            try:
                bgr = surgical_seal_mask_repair(bgr)
            except Exception:
                pass
        out = bgr
        H, W = out.shape[:2]
        tile_h = int(H / num_tiles)
        overlap_h = int(tile_h * 0.1) # 10% 物理重叠
        b64_list = []
        for i in range(num_tiles):
            y0 = max(0, i * tile_h - (overlap_h if i > 0 else 0))
            y1 = min(H, (i + 1) * tile_h + (overlap_h if i < num_tiles - 1 else 0))
            crop = out[y0:y1, :]
            scale = VISION_MAX_EDGE / max(crop.shape[:2])
            if scale < 1:
                crop = cv2.resize(crop, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
            b64_str = base64.b64encode(cv2.imencode('.jpg', crop, [int(cv2.IMWRITE_JPEG_QUALITY), 92])[1].tobytes()).decode()
            b64_list.append((b64_str, crop))
        return b64_list

    def _grids_from_tiles(self, page, base, num_tiles=2):
        print(f"    [*] 页面数据量大或发生截断，正在执行水平 {num_tiles} 分块切片识别...")
        tile_items = self._slice_page_tiles(page, enhanced=True, num_tiles=num_tiles)
        merged_cells = []
        all_unc = []

        # 并发请求切片识别，消灭串行排队超时
        def _fetch_tile(idx_item):
            idx, (b64, crop) = idx_item
            p_text = f"请转录切片图片中的全部表格，返回 JSON。这是第 {idx+1}/{num_tiles} 个垂直分块。"
            return idx, self._vision_call_prompt(b64, p_text, img_bgr=crop)

        with ThreadPoolExecutor(max_workers=min(num_tiles, 3)) as pool:
            tile_results = list(pool.map(_fetch_tile, enumerate(tile_items)))

        tile_results.sort(key=lambda x: x[0])
        target_cols = None

        def _is_hdr(row):
            t = ' '.join(str(c) for c in row)
            return any(k in t for k in ('材料名称', '品名', '规格', '单价', '税率')) and ('序号' in t or '单位' in t)

        def _row_keys(row):
            seq = None
            for c in row[:2]:
                s = str(c).strip()
                if re.fullmatch(r'\d+', s):
                    seq = s
                    break
            spec = ''
            for c in row:
                s = str(c).strip()
                if any(k in s for k in ('×', '*', 'mm', '#', 'DN', 'Q235', 'Q355')):
                    spec = s
                    break
            price = ''
            for c in row:
                s = str(c).strip().replace(',', '')
                if re.fullmatch(r'\d+(\.\d+)?', s) and float(s) > 10:
                    price = s
                    break
            return seq, spec, price

        def _align_row_to_hdr(row, hdr):
            if len(row) == len(hdr):
                return row
            slots = {}
            r = [str(c).strip() for c in row if str(c).strip() != '']
            if r and re.fullmatch(r'\d+', r[-1]) and int(r[-1]) in (0, 3, 6, 9, 13):
                slots['税率'] = r.pop()
            if r and r[-1] in ('是', '否', '含', '不含'):
                slots['是否含税'] = r.pop()
            if r and re.fullmatch(r'\d+(\.\d+)?', r[-1].replace(',', '')) and float(r[-1].replace(',', '')) > 50:
                slots['单价'] = r.pop()
            if r and r[-1] in ('t', 'm', 'kg', '套', '个', '吨', '米', '批'):
                slots['单位'] = r.pop()
            if r and re.fullmatch(r'\d+', r[0]):
                slots['序号'] = r.pop(0)
            if r:
                slots['材料名称'] = r.pop(0)
            if r and len(r) >= 2:
                slots['品牌'] = r.pop(0)
                slots['属性项属性值'] = ' '.join(r)
            elif r:
                slots['属性项属性值'] = r.pop(0)

            aligned = []
            for h in hdr:
                h_clean = str(h or '').strip()
                val = ''
                if any(k in h_clean for k in ('序号', '编号', 'No')):
                    val = slots.get('序号', '')
                elif any(k in h_clean for k in ('材料名称', '品名', '名称')):
                    val = slots.get('材料名称', '')
                elif any(k in h_clean for k in ('品牌', '厂家', '生产厂')):
                    val = slots.get('品牌', '')
                elif any(k in h_clean for k in ('属性', '规格', '参数')):
                    val = slots.get('属性项属性值', '')
                elif any(k in h_clean for k in ('单价', '价格', '含税价', '不含税价')):
                    val = slots.get('单价', '')
                elif any(k in h_clean for k in ('单位', '计量单位')):
                    val = slots.get('单位', 't')
                elif any(k in h_clean for k in ('是否含税', '含税')):
                    val = slots.get('是否含税', '是')
                elif any(k in h_clean for k in ('税率', '增值税率')):
                    val = slots.get('税率', '13')
                aligned.append(val)
            return aligned

        for idx, res in tile_results:
            tables = self._normalize_tables(res)
            if not tables or not tables[0][0]:
                print(f"    ! [切片告警] {base} 第 {idx+1}/{num_tiles} 切片识别失败，插入异常标记行以防静默丢失")
                cols = target_cols or 2
                err_row = [f"异常:{base}第{idx+1}分块识别失败"] + [''] * (cols - 1)
                merged_cells.append(err_row)
                continue
            grid, unc = tables[0]
            if idx == 0:
                target_cols = len(grid[0])
                merged_cells.extend(grid)
            else:
                # 过滤后续切片顶部的重复表头
                clean_grid = []
                for r in grid:
                    if not clean_grid and _is_hdr(r):
                        continue
                    clean_grid.append(r)

                # 建立前密切片尾部最近 12 行指纹索引
                tail_keys = []
                for t_i, tr in enumerate(merged_cells[-12:]):
                    t_abs = len(merged_cells) - 12 + t_i if len(merged_cells) >= 12 else t_i
                    s, sp, p = _row_keys(tr)
                    tail_keys.append((t_abs, s, sp, p, tr))

                hdr_row = merged_cells[0] if (merged_cells and _is_hdr(merged_cells[0])) else None

                start_r = 0
                last_matched_t_abs = None
                for r_i, row in enumerate(clean_grid):
                    s2, sp2, p2 = _row_keys(row)
                    r_text = ''.join(str(c) for c in row)
                    matched = None
                    for t_abs, s1, sp1, p1, r1 in tail_keys:
                        if s1 and s2:
                            if s1 == s2:
                                matched = (t_abs, r1)
                                break
                        else:
                            # 缺序号时，需规格与单价双重对齐才允许判定为同一行
                            if sp1 and sp2 and sp1 == sp2 and p1 and p2 and p1 == p2:
                                matched = (t_abs, r1)
                                break
                    if matched:
                        t_abs, r1 = matched
                        if last_matched_t_abs is None or t_abs > last_matched_t_abs:
                            last_matched_t_abs = t_abs
                        if len(r_text) > len(''.join(str(c) for c in r1)):
                            merged_cells[t_abs] = _align_row_to_hdr(row, hdr_row) if hdr_row else row
                        start_r = r_i + 1
                    else:
                        # 检查切片交界处的碎片残留行
                        if '品种' not in r_text and '规格' not in r_text and len(r_text) < 40 and r_i == 0:
                            start_r = r_i + 1
                            continue
                        start_r = r_i
                        break

                # 关键修复：截断前密切片尾部由于重叠切割产生的残影碎片行
                if last_matched_t_abs is not None:
                    merged_cells = merged_cells[:last_matched_t_abs + 1]

                hdr_row = merged_cells[0] if (merged_cells and _is_hdr(merged_cells[0])) else None
                for row in clean_grid[start_r:]:
                    if not any(str(c).strip() for c in row):
                        continue
                    if hdr_row and len(row) != len(hdr_row):
                        row = _align_row_to_hdr(row, hdr_row)
                    elif target_cols and len(row) != target_cols:
                        row = row + [''] * (target_cols - len(row)) if len(row) < target_cols else row[:target_cols]
                    merged_cells.append(row)
            all_unc.extend(unc)

        if not merged_cells:
            return None
        return [(merged_cells, all_unc)]

    def _vision_call_prompt(self, b64_or_tuple, prompt_text, img_bgr=None):
        """调用统一多模态接口转录局部切片表格，受 RapidOCR 60% 重合度硬门禁保护"""
        if isinstance(b64_or_tuple, tuple):
            b64, img_bgr = b64_or_tuple
        else:
            b64 = b64_or_tuple
        return self._execute_multimodal_vision(b64, prompt_text, kind='vision_table_tile', img_bgr=img_bgr)

    def _thread_page(self, p_idx):
        """每线程独占一份 PyMuPDF Document —— Document 不是线程安全的，跨线程共享会读坏页面"""
        tl = getattr(self, '_tl', None)
        if tl is None:
            self._tl = tl = threading.local()
        doc = getattr(tl, 'doc', None)
        if doc is None:
            doc = tl.doc = fitz.open(self.pdf_path)
        return doc[p_idx]

    def _grids_from_vision(self, page, base):
        # 步骤 1：OpenCV 物理网格先验探测
        pix = page.get_pixmap(dpi=VISION_DPI)
        bgr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        bgr = cv2.cvtColor(bgr, cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR)
        bgr = auto_orient_page_upright(bgr)
        phys_bands = detect_physical_table_row_bands(bgr)
        phys_rows = len(phys_bands)

        tables = None
        first = None

        # 当单页物理表格行数 > 20 行（如东方雨虹第 4 页 34 行，未来新材 24 行），
        # 大模型极易发生同质行注意力滑移（Attention Skipping 漏行跳行），主动执行切片识别！
        if phys_rows > 20:
            num_tiles = 2 if phys_rows <= 32 else 3
            print(f"    [*] {base}: 物理网格探测发现 {phys_rows} 行密集表格，主动执行 {num_tiles} 分块切片防跳行识别...")
            tables = self._grids_from_tiles(page, base, num_tiles=num_tiles)
        else:
            # 第一遍：增强图 (以此为准)；第二遍：原图 (预处理不同，错法不同，用于交叉校验)
            first = self._vision_call(self._page_png_b64(page, enhanced=True))
            if first and first.get("_truncated"):
                tables = self._grids_from_tiles(page, base, num_tiles=3)
            else:
                tables = self._normalize_tables(first)
                if tables and len(tables[0][0]) > VISION_MAX_ROWS_PER_TILE:
                    n_rows = len(tables[0][0])
                    empty_price_count = sum(1 for r in tables[0][0][1:] if len(r) > 3 and not str(r[3]).strip())
                    if empty_price_count > 5 or n_rows >= 60:
                        print(f"    [*] {base}: 页面行数 {n_rows} 且有 {empty_price_count} 行单价缺失，主动执行 3 分块切片高精识别...")
                        t_sliced = self._grids_from_tiles(page, base, num_tiles=3)
                        if t_sliced:
                            tables = t_sliced
                    else:
                        self.cell_issue[(f"{base}#1", 1, 1)] = (
                            f"本页识别 {n_rows} 行，超过单块阈值 {VISION_MAX_ROWS_PER_TILE} 行")

        if not tables:
            # 切块重试一次 (num_tiles=3)
            tables = self._grids_from_tiles(page, base, num_tiles=3)
        if not tables:
            raise RuntimeError(f"{base}: 视觉模型未返回表格，请检查 VISION_* 配置或模型是否支持图片输入")
        if first and first.get("page_title"):
            # 并发时各线程写各自的 title_sink，主线程再按页序收拢，避免页标题顺序被线程打乱
            sink = getattr(getattr(self, '_tl', None), 'title_sink', None)
            (sink if sink is not None else self.page_titles).append(str(first["page_title"]))

        second_tables = None
        if VISION_DOUBLE_CHECK:
            second = self._vision_call(self._page_png_b64(page, enhanced=False))
            second_tables = self._normalize_tables(second) if second else None
            if not second_tables:
                print(f"    ! {base}: 第二遍识别失败，本页未做双读校验")

        grids = []
        for t_idx, (grid, unc) in enumerate(tables, 1):
            key = f"{base}#{t_idx}"
            for u in unc:
                try:
                    r, c = int(u[0]), int(u[1])
                except (TypeError, ValueError):
                    continue
                reason = str(u[2]) if len(u) > 2 else "看不清"
                self.cell_issue[(key, r + 1, c + 1)] = f"视觉模型标注存疑({reason})"
            if second_tables is not None:
                g2 = second_tables[t_idx - 1][0] if t_idx <= len(second_tables) else []
                diff, unmatched = self._cross_check(grid, g2, key)
                msg = f"不一致 {diff} 格" + (f"，{unmatched} 行无法对齐(已标记)" if unmatched else "")
                print(f"    -> {base} 表{t_idx}: 双读校验完成，{msg}")
            grids.append(grid)
        return grids

    def _cross_check(self, grid, g2, key):
        """
        逐格比对两遍识别结果。两遍行数不同 (多/少识别一行表头等) 时按首列行标题对齐；
        列数不同或行标题对不上的行整行标记。返回 (不一致格数, 无法对齐行数)
        """
        def norm(v):
            # NFKC 归一化全角/半角标点符号，去除空白和千分位逗号
            s = unicodedata.normalize('NFKC', str(v or ''))
            return s.replace(' ', '').replace(',', '').replace('　', '')

        width_ok = bool(g2) and len(g2[0]) == len(grid[0])
        if width_ok and len(g2) == len(grid):
            pairs = list(zip(range(len(grid)), g2))
        else:
            by_label = {}
            for row in g2 if width_ok else []:
                by_label.setdefault(norm(row[0]), row)
            pairs = [(r, by_label.get(norm(grid[r][0]))) for r in range(len(grid))]

        diff = unmatched = 0
        for r, row2 in pairs:
            if row2 is None:
                if any(grid[r][1:]):
                    unmatched += 1
                    for c, v in enumerate(grid[r]):
                        if v:
                            self.cell_issue.setdefault((key, r + 1, c + 1), "第二遍识别中找不到对应行")
                continue
            for c in range(len(grid[r])):
                a, b = grid[r][c], row2[c] if c < len(row2) else ''
                if norm(a) != norm(b):
                    diff += 1
                    self.cell_issue[(key, r + 1, c + 1)] = f"两遍识别不一致: “{a}” / “{b}”"
        return diff, unmatched

    def save_page_images(self, img_dir):
        """保存每页原件截图 quote_img{页码}.png，返回 {页码: 文件名}"""
        os.makedirs(img_dir, exist_ok=True)
        out = {}
        for i, page in enumerate(self.doc, 1):
            fn = f"quote_img{i}.png"
            page.get_pixmap(dpi=150).save(os.path.join(img_dir, fn))
            out[i] = fn
        return out

    def _ocr_seal_image(self, img_bgr) -> Optional[str]:
        """识别扶正/展开后的公章图像中的企业名称"""
        if img_bgr is None:
            return None
        # 1. 优先尝试本地 RapidOCR (复用全局单例，杜绝重复加载模型)
        try:
            from universal_pdf_grid_engine import _OCR_ENGINE
            ocr = _OCR_ENGINE
            if ocr is None:
                from rapidocr_onnxruntime import RapidOCR
                ocr = RapidOCR()
            result, _ = ocr(img_bgr)
            if result:
                full_text = "".join(line[1] for line in result)
                if COMPANY_RE.search(full_text):
                    return full_text
        except Exception:
            pass

        # 2. 调用多模态视觉模型提取企业名称 (使用通用自由文本 Prompt，不走表格校验)
        try:
            import base64
            b64 = base64.b64encode(cv2.imencode('.png', img_bgr)[1].tobytes()).decode()
            if not VISION_API_KEY:
                return None
            url = f"{VISION_BASE_URL.rstrip('/')}/chat/completions"
            headers = {"Authorization": f"Bearer {VISION_API_KEY}", "Content-Type": "application/json"}
            payload = {
                "model": VISION_MODEL,
                "temperature": 0,
                "messages": [
                    {"role": "system", "content": "你是一位专业的文档与印章文字识别专家。请仔细识别图片中红色印章上的文字，只输出印章上的单位/企业全称（必须包含有限公司、股份有限公司、厂或经营部等），不要输出任何其他说明。"},
                    {"role": "user", "content": [
                        {"type": "text", "text": "请提取此印章中的企业全称："},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                    ]},
                ]
            }
            resp = requests.post(url, headers=headers, json=payload, timeout=20)
            if resp.status_code == 200:
                txt = resp.json()["choices"][0]["message"]["content"]
                return txt.strip()
        except Exception:
            pass
        return None

    def guess_supplier(self):
        """从页面表格上方的标题文字，或旋转扶正后的红色公章中识别公司全称"""
        for t in self.page_titles:
            m = COMPANY_RE.search(re.sub(r'\s', '', t))
            # 只认出后半截 (如 “集团）股份有限公司”) 时不可用
            if m and not re.match(r'^[（\(]?(集团|股份|有限|公司)', m.group(1)):
                return m.group(1).replace('(', '（').replace(')', '）')

        # 兜底：尝试通过 SealEnhancerService 从页面中检测并扶正公章提取供应商 (优先末页与首页)
        try:
            with fitz.open(self.pdf_path) as doc:
                pages_to_check = [doc[-1], doc[0]] if len(doc) > 1 else ([doc[0]] if len(doc) > 0 else [])
                for page in pages_to_check:
                    pix = page.get_pixmap(dpi=150)
                    bgr = cv2.cvtColor(np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n),
                                       cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR)
                    seal_res = _SEAL_ENHANCER.process_document_image(bgr)
                    if seal_res.get("has_seal"):
                        # 优先读取旋转扶正后的正立印章，其次读取极坐标展开条
                        for key in ("upright_seal_img", "polar_strip_img"):
                            target_img = seal_res.get(key)
                            if target_img is not None:
                                txt = self._ocr_seal_image(target_img)
                                if txt:
                                    m = COMPANY_RE.search(re.sub(r'\s', '', txt))
                                    if m and not re.match(r'^[（\(]?(集团|股份|有限|公司)', m.group(1)):
                                        return m.group(1).replace('(', '（').replace(')', '）')
        except Exception:
            pass
        return None

    # ------------------------------------------------------------------ 文字层
    def _lines_from_text(self, page):
        """没有表格线时按文本行切分 (整页文字)"""
        self.page_titles.append(page.get_text()[:200])
        return [[[p for p in re.split(r'\s{2,}|\t', line.strip()) if p]
                 for line in page.get_text().split('\n') if line.strip()]]

    def _grid_from_borderless_text(self, page):
        """当没有矢量边框线 (find_tables 为空) 时，通过单词物理坐标聚类与表头列锚点还原规整无框表格"""
        words = page.get_text('words')
        if not words:
            return None
        # 按 y 坐标归并视觉文本行
        lines = []
        curr_line = []
        curr_y = None
        for w in sorted(words, key=lambda x: (round(x[1], 1), x[0])):
            if curr_y is None or abs(w[1] - curr_y) <= 3.5:
                curr_line.append(w)
                curr_y = w[1] if curr_y is None else curr_y
            else:
                lines.append(curr_line)
                curr_line = [w]
                curr_y = w[1]
        if curr_line:
            lines.append(curr_line)

        header_idx = None
        header_keywords = ['类别', '名称', '品名', '规格', '属性', '单位', '单价', '报价', '品牌', '金额', '备注', '序号', '型号']
        for idx, l in enumerate(lines[:15]):
            line_text = ''.join(w[4] for w in l)
            matched = sum(1 for kw in header_keywords if kw in line_text)
            if matched >= 2:
                header_idx = idx
                break

        if header_idx is None:
            return None

        header_line = lines[header_idx]
        col_starts = [w[0] for w in header_line]
        col_names = [w[4] for w in header_line]

        col_intervals = []
        for i in range(len(header_line)):
            left = 0.0 if i == 0 else col_starts[i] - 2.0
            right = 10000.0 if i == len(header_line) - 1 else col_starts[i + 1] - 2.0
            col_intervals.append((left, right, col_names[i]))

        data_lines = lines[header_idx + 1:]
        grid_rows = [[c[2] for c in col_intervals]]
        current_row = [''] * len(col_intervals)

        # 查找单位/单价/品牌所在列索引作为行锚点
        anchor_cols = [c_idx for c_idx, (_, _, name) in enumerate(col_intervals) if any(k in name for k in ['单位', '单价', '报价', '品牌', '金额'])]

        for dl in data_lines:
            dl_cols = {}
            for w in dl:
                wx = (w[0] + w[2]) / 2.0
                for c_idx, (c_left, c_right, _) in enumerate(col_intervals):
                    if c_left <= wx < c_right:
                        dl_cols.setdefault(c_idx, []).append(w[4])
                        break
            has_anchor = any(c in dl_cols for c in anchor_cols) or len(dl_cols) >= max(2, len(col_intervals) // 2)
            if not has_anchor:
                for c_idx, wl in dl_cols.items():
                    current_row[c_idx] += ''.join(wl)
            else:
                for c_idx, wl in dl_cols.items():
                    sep = ' ' if c_idx not in anchor_cols else ''
                    current_row[c_idx] = (current_row[c_idx] + sep + ' '.join(wl)).strip()
                grid_rows.append(current_row)
                current_row = [''] * len(col_intervals)

        if len(grid_rows) >= 2:
            self.page_titles.append(page.get_text()[:200])
            return grid_rows
        return None

    def _grids_from_text_layer(self, page):
        """一页可能有多张表，必须全部返回；原实现只取面积最大的一张，其余整张静默丢弃"""
        try:
            tables = page.find_tables().tables
        except Exception:
            tables = []
        if not tables:
            borderless_grid = self._grid_from_borderless_text(page)
            if borderless_grid and len(borderless_grid) > 1 and len(borderless_grid[0]) >= 3:
                return [borderless_grid]
            return self._lines_from_text(page)
        cand = []
        for tb in sorted(tables, key=lambda t: (round(t.bbox[1] / 5), t.bbox[0])):
            rows = tb.extract()
            if rows:
                cand.append((tb, rows, (tb.bbox[2] - tb.bbox[0]) * (tb.bbox[3] - tb.bbox[1])))
        max_area = max((a for _, _, a in cand), default=0)
        grids = []
        for tb, rows, area in cand:
            # 面积不到最大表 20% 的当碎片丢弃 (页眉装饰框)，其余全部解析并逐张打印
            if max_area and area < max_area * 0.2:
                print(f"    ! 忽略碎片表格 (面积占比 {area / max_area:.0%})")
                continue
            self.page_titles.append(page.get_text(clip=fitz.Rect(0, 0, page.rect.width, tb.bbox[1])))
            # find_tables 对合并单元格返回 None：向下/向右填充，与 Excel 合并单元格展开语义一致
            grid = [[(c or '').replace('\n', ' ').strip() for c in r] for r in rows]
            grids.append(_fill_merged_none(rows, grid))
        if len(grids) > 1:
            print(f"    -> 本页识别到 {len(grids)} 张表格，全部解析")
        return grids or self._lines_from_text(page)

    @staticmethod
    def _remove_red(img):
        """把红色公章像素置白 (黑色文字保留)"""
        b, g, r = [c.astype(np.int16) for c in cv2.split(img)]
        out = img.copy()
        out[(r > 120) & (r - g > 40) & (r - b > 40)] = 255
        return out


def _strip_fence(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r'^```(?:json)?', '', text).rstrip('`').strip()
    return text


def _fill_merged_none(raw_rows, grid):
    """
    PyMuPDF 表格中 None 表示被合并覆盖的格子。
    在报价单/数据表中，跨行合并（如分类、品牌、单位等向下合并）占绝大多数；
    跨列合并在同一行中也可能出现，但绝不能把左侧的值（如价格、规格）横向填入右侧不同的列（如品牌、单位列）。
    因此：优先同列向上寻找合并起始值（纵向穿透）；仅当同列上方无有效值且同一行左侧为有效合并时才考虑横向。
    """
    num_rows = len(raw_rows)
    num_cols = max(len(r) for r in raw_rows) if raw_rows else 0
    for c in range(num_cols):
        for r in range(num_rows):
            if r < len(raw_rows) and c < len(raw_rows[r]):
                if raw_rows[r][c] is None:
                    # 优先沿同列向上继承（纵向合并单元格）
                    if r > 0 and grid[r - 1][c]:
                        grid[r][c] = grid[r - 1][c]
                    elif c > 0 and grid[r][c - 1] and raw_rows[r][c - 1] is not None:
                        grid[r][c] = grid[r][c - 1]
    return grid
