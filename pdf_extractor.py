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
import time
import unicodedata

import cv2
import fitz
import numpy as np
import openpyxl
import requests

from config import (VISION_API_KEY, VISION_BASE_URL, VISION_MODEL, VISION_DOUBLE_CHECK,
                    VISION_GAMMA, VISION_MAX_ROWS_PER_TILE, VISION_TILE_OVERLAP_ROWS, LLM_MAX_RETRIES)

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
        for p_idx in range(total_p):
            page = self.doc[p_idx]
            base = f"第{p_idx + 1}页"
            if len(re.sub(r'\s', '', page.get_text())) >= 30:
                grids = self._grids_from_text_layer(page)
                src = 'text'
            elif use_vision:
                try:
                    grids = self._grids_from_vision(page, base)
                    src = 'vision'
                except Exception as page_err:
                    print(f"    ! [隔离] {base} 识别失败: {page_err}")
                    if not hasattr(self, 'failed_pages'):
                        self.failed_pages = []
                    self.failed_pages.append((base, str(page_err)))
                    # 失败页面留痕隔离 (不执行中断，确保后续页面被继续处理并如实对账)
                    ws_err = wb.create_sheet(base)
                    ws_err.cell(1, 1, value=f"异常:{base}识别失败({page_err})")
                    self.text_source[base] = 'error'
                    continue
            else:
                raise RuntimeError(f"{base} 是扫描件，需要可用的视觉模型 (DEEPSEEK_* 或 VISION_*)")
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

    # ------------------------------------------------------------------ 视觉模型
    def _page_png_b64(self, page, enhanced=True):
        """
        渲染页面给视觉模型。enhanced=True 时做识别增强：
        去除红色公章 → 灰度 → gamma 压暗褪色笔画 (防止褪色笔画被误读)
        """
        dpi = SCAN_DPI if enhanced else VISION_DPI
        pix = page.get_pixmap(dpi=dpi)
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        if enhanced:
            bgr = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(self._remove_red(bgr), cv2.COLOR_BGR2GRAY)
            out = (255 * (gray / 255.0) ** VISION_GAMMA).astype(np.uint8)
        else:
            out = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR)
        scale = VISION_MAX_EDGE / max(out.shape[:2])
        if scale < 1:
            out = cv2.resize(out, None, fx=scale, fy=scale, interpolation=cv2.INTER_AREA)
        return base64.b64encode(cv2.imencode('.png', out)[1].tobytes()).decode()

    def _vision_call(self, b64):
        """调用 OpenAI 兼容的多模态接口，返回解析后的 dict；失败返回 None"""
        url = f"{VISION_BASE_URL.rstrip('/')}/chat/completions"
        headers = {"Authorization": f"Bearer {VISION_API_KEY}", "Content-Type": "application/json"}
        payload = {
            "model": VISION_MODEL,
            "temperature": 0,
            "messages": [
                {"role": "system", "content": VISION_TABLE_PROMPT},
                {"role": "user", "content": [
                    {"type": "text", "text": "请转录这一页中的全部表格，返回 JSON。"},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                ]},
            ],
            "response_format": {"type": "json_object"},
        }
        last = ""
        for attempt in range(min(LLM_MAX_RETRIES + 1, 2)):
            try:
                resp = requests.post(url, headers=headers, json=payload, timeout=12)
                if resp.status_code == 400 and "response_format" in payload:
                    # 部分模型不支持 response_format，去掉后重试
                    payload.pop("response_format")
                    continue
                if resp.status_code == 200:
                    data = resp.json()
                    choice = data["choices"][0]
                    finish_reason = choice.get("finish_reason")
                    content = choice["message"]["content"]
                    if finish_reason == "length":
                        print("[Vision Warning] 视觉模型输出触发 length 截断！")
                        return {"_truncated": True}
                    try:
                        parsed = json.loads(_strip_fence(content))
                        bad = validate_vision_payload(parsed)
                        if bad:
                            print(f"[Vision Warning] 响应结构不合规（{bad}），整份作废")
                            last = f"结构不合规: {bad}"
                            continue
                        return parsed
                    except Exception as json_err:
                        print(f"[Vision Warning] JSON解析失败(疑似截断): {json_err}")
                        return {"_truncated": True}
                last = f"HTTP {resp.status_code}: {resp.text[:200]}"
                if 400 <= resp.status_code < 500 and resp.status_code != 429:
                    break
            except Exception as e:
                last = str(e)
                # 遇到证书或连接断开直接终止，防止无效重试阻塞
                if "SSLError" in str(e) or "Max retries exceeded" in str(e):
                    break
            if attempt < 1:
                time.sleep(1)
        print(f"[Vision Warning] 视觉模型调用失败: {last}")
        return None

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
        if enhanced:
            bgr = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR)
            gray = cv2.cvtColor(self._remove_red(bgr), cv2.COLOR_BGR2GRAY)
            out = (255 * (gray / 255.0) ** VISION_GAMMA).astype(np.uint8)
        else:
            out = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR)
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
            b64_list.append(base64.b64encode(cv2.imencode('.png', crop)[1].tobytes()).decode())
        return b64_list

    def _grids_from_tiles(self, page, base, num_tiles=2):
        print(f"    [*] 页面数据量大或发生截断，正在执行水平 {num_tiles} 分块切片识别...")
        b64_tiles = self._slice_page_tiles(page, enhanced=True, num_tiles=num_tiles)
        merged_cells = []
        all_unc = []
        known_headers = None
        target_cols = None

        for idx, b64 in enumerate(b64_tiles):
            prompt = f"请转录切片图片中的全部表格，返回 JSON。这是第 {idx+1}/{num_tiles} 个垂直分块。"
            if known_headers:
                prompt += f" 表格列数必须为 {target_cols} 列，首块表头为: {known_headers}。"
            res = self._vision_call_prompt(b64, prompt)
            tables = self._normalize_tables(res)
            if not tables or not tables[0][0]:
                print(f"    ! [切片告警] {base} 第 {idx+1}/{num_tiles} 切片识别失败，插入异常标记行以防静默丢失")
                cols = target_cols or 2
                err_row = [f"异常:{base}第{idx+1}分块识别失败"] + [''] * (cols - 1)
                merged_cells.append(err_row)
                continue
            grid, unc = tables[0]
            if idx == 0:
                known_headers = grid[0]
                target_cols = len(grid[0])
                merged_cells.extend(grid)
            else:
                # 消除重叠区重复行 (按整行文本或去除空格匹配)
                start_r = 0
                for r_i, row in enumerate(grid):
                    row_str = ''.join(str(c).strip() for c in row)
                    # 检查是否已存在于已合并的最后5行
                    is_dup = False
                    for existing in merged_cells[-6:]:
                        ex_str = ''.join(str(c).strip() for c in existing)
                        if row_str and row_str == ex_str:
                            is_dup = True
                            break
                    if not is_dup:
                        start_r = r_i
                        break
                for row in grid[start_r:]:
                    # 列数对齐检查
                    if target_cols and len(row) != target_cols:
                        row = row + [''] * (target_cols - len(row)) if len(row) < target_cols else row[:target_cols]
                    merged_cells.append(row)
            all_unc.extend(unc)

        if not merged_cells:
            return None
        return [(merged_cells, all_unc)]

    def _vision_call_prompt(self, b64, prompt_text):
        url = f"{VISION_BASE_URL.rstrip('/')}/chat/completions"
        headers = {"Authorization": f"Bearer {VISION_API_KEY}", "Content-Type": "application/json"}
        payload = {
            "model": VISION_MODEL,
            "temperature": 0,
            "messages": [
                {"role": "system", "content": VISION_TABLE_PROMPT},
                {"role": "user", "content": [
                    {"type": "text", "text": prompt_text},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
                ]},
            ],
            "response_format": {"type": "json_object"},
        }
        for attempt in range(LLM_MAX_RETRIES + 1):
            try:
                resp = requests.post(url, headers=headers, json=payload, timeout=180)
                if resp.status_code == 200:
                    choice = resp.json()["choices"][0]
                    content = choice["message"]["content"]
                    parsed = json.loads(_strip_fence(content))
                    bad = validate_vision_payload(parsed)
                    if bad:
                        raise ValueError(f"响应结构不合规: {bad}")
                    return parsed
            except Exception:
                pass
            if attempt < LLM_MAX_RETRIES:
                time.sleep(2 ** attempt)
        return None

    def _grids_from_vision(self, page, base):
        # 第一遍：增强图 (以此为准)；第二遍：原图 (预处理不同，错法不同，用于交叉校验)
        first = self._vision_call(self._page_png_b64(page, enhanced=True))
        if first and first.get("_truncated"):
            tables = self._grids_from_tiles(page, base, num_tiles=3)
        else:
            tables = self._normalize_tables(first)
            # 若单页识别出的行数大于阈值，也支持容错
            if tables and len(tables[0][0]) > VISION_MAX_ROWS_PER_TILE and len(tables[0][0]) >= 60:
                print(f"    [*] 页面行数 ({len(tables[0][0])}) 超过阈值，切块验证重合...")

        if not tables:
            # 切块重试一次 (num_tiles=3)
            tables = self._grids_from_tiles(page, base, num_tiles=3)
        if not tables:
            raise RuntimeError(f"{base}: 视觉模型未返回表格，请检查 VISION_* 配置或模型是否支持图片输入")
        if first.get("page_title"):
            self.page_titles.append(str(first["page_title"]))

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

    def guess_supplier(self):
        """从页面表格上方的标题文字中识别公司全称"""
        for t in self.page_titles:
            m = COMPANY_RE.search(re.sub(r'\s', '', t))
            # 只认出后半截 (如 “集团）股份有限公司”) 时不可用
            if m and not re.match(r'^[（\(]?(集团|股份|有限|公司)', m.group(1)):
                return m.group(1).replace('(', '（').replace(')', '）')
        return None

    # ------------------------------------------------------------------ 文字层
    def _lines_from_text(self, page):
        """没有表格线时按文本行切分 (整页文字)"""
        self.page_titles.append(page.get_text()[:200])
        return [[[p for p in re.split(r'\s{2,}|\t', line.strip()) if p]
                 for line in page.get_text().split('\n') if line.strip()]]

    def _grids_from_text_layer(self, page):
        """一页可能有多张表，必须全部返回；原实现只取面积最大的一张，其余整张静默丢弃"""
        try:
            tables = page.find_tables().tables
        except Exception:
            tables = []
        if not tables:
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
    """PyMuPDF 表格中 None 表示被合并覆盖的格子：用左侧/上方的值填充"""
    for r, row in enumerate(raw_rows):
        for c, v in enumerate(row):
            if v is None:
                if c > 0 and raw_rows[r][c - 1] is not None or (c > 0 and grid[r][c - 1]):
                    grid[r][c] = grid[r][c - 1]
                elif r > 0:
                    grid[r][c] = grid[r - 1][c]
    return grid
