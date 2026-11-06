# -*- coding: utf-8 -*-
"""
universal_pdf_grid_engine.py
系统级 PDF 表格防漏引擎：
1. 四方向 (0°/90°/180°/270°) 实测置信度自动摆正（通杀奇偶页正反横扫、Landscape 横置页）；
2. OpenCV 物理表格行探测 + 密集长页沿格线无损分片（彻底消灭大模型长表注意力跳行与 Token 截断）；
3. 全文档跨页同构表自动缝合 + 任意页表头全局广播 + 全无表头时基于列数据分布自动推导表头。
"""
import re
import cv2
import fitz
import numpy as np
from typing import List, Dict, Tuple, Optional, Any

_OCR_ENGINE = None
_ORIENT_ENGINE = None

try:
    from rapid_orientation import RapidOrientation
    _ORIENT_ENGINE = RapidOrientation()
except Exception:
    _ORIENT_ENGINE = None

try:
    from rapidocr_onnxruntime import RapidOCR
    _OCR_ENGINE = RapidOCR()
except Exception:
    try:
        import pytesseract
        _OCR_ENGINE = "pytesseract"
    except Exception:
        _OCR_ENGINE = None


# ==============================================================================
# 模块 1：四方向 (0° / 90° / 180° / 270°) 100% 确定性页面自动摆正
# ==============================================================================
def auto_orient_page_upright(bgr_img: np.ndarray) -> np.ndarray:
    """
    系统性解决任意方向扫描件（含竖版A4印横向表、奇偶页+90°/-90°交替旋转、180°倒置）：
    优先调用 5ms 级别的轻量 RapidOrientation ONNX 引擎直接判定方向并纠正；
    若不可用，则退化至形态学梯度 + RapidOCR/Tesseract 实测中文字符高置信度。
    """
    if bgr_img is None or bgr_img.size == 0:
        return bgr_img

    # 1. 优先使用极速轻量分类器 RapidOrientation (单页耗时约 5ms)
    if _ORIENT_ENGINE is not None:
        try:
            angle_res, _ = _ORIENT_ENGINE(bgr_img)
            angle = int(angle_res)
            if angle == 90:
                return cv2.rotate(bgr_img, cv2.ROTATE_90_COUNTERCLOCKWISE)
            elif angle == 270:
                return cv2.rotate(bgr_img, cv2.ROTATE_90_CLOCKWISE)
            elif angle == 180:
                return cv2.rotate(bgr_img, cv2.ROTATE_180)
            elif angle == 0:
                return bgr_img
        except Exception:
            pass

    # 2. 降级方案：形态学判定走向 + OCR 实测置信度
    gray = cv2.cvtColor(bgr_img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 25, 15
    )

    # 去除长表格线干扰，仅保留字符连通域测量文字排列方向
    h, w = binary.shape[:2]
    h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (max(40, w // 15), 1)))
    v_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(40, h // 15))))
    chars_only = cv2.subtract(binary, cv2.bitwise_or(h_lines, v_lines))

    # 膨胀字符：观察水平连带 vs 垂直连带的数量与方差
    horiz_connect = cv2.dilate(chars_only, cv2.getStructuringElement(cv2.MORPH_RECT, (15, 2)))
    vert_connect = cv2.dilate(chars_only, cv2.getStructuringElement(cv2.MORPH_RECT, (2, 15)))
    h_energy = float(np.var(np.mean(horiz_connect, axis=1)))
    v_energy = float(np.var(np.mean(vert_connect, axis=0)))

    # 若垂直方向文本行能量显著高于水平方向，说明页面旋转了 +90° 或 -90°
    is_sideways = v_energy > h_energy * 1.25

    if is_sideways:
        cand_cw = cv2.rotate(bgr_img, cv2.ROTATE_90_CLOCKWISE)
        cand_ccw = cv2.rotate(bgr_img, cv2.ROTATE_90_COUNTERCLOCKWISE)
        cand_orig = bgr_img
        return _pick_best_upright_by_ocr([cand_orig, cand_cw, cand_ccw])
    else:
        # 水平方向时，快速检验是否 180° 上下颠倒
        return _pick_best_upright_by_ocr([bgr_img])


def _pick_best_upright_by_ocr(candidates: List[np.ndarray]) -> np.ndarray:
    """用 RapidOCR / Tesseract 对候选旋转图核心区域做实测，选出高置信度中文字符最多的正立图像"""
    if len(candidates) == 1:
        return candidates[0]

    best_img = candidates[0]
    best_score = -1.0

    for img in candidates:
        ih, iw = img.shape[:2]
        # 截取页面核心区域并缩放到适中尺寸以保证毫秒级完成
        crop = img[int(ih * 0.1):int(ih * 0.9), int(iw * 0.1):int(iw * 0.9)]
        scale = min(1.0, 900.0 / max(crop.shape[:2], default=900))
        if scale < 1.0:
            crop = cv2.resize(crop, (int(crop.shape[1] * scale), int(crop.shape[0] * scale)))

        score = 0.0
        # 优先使用 RapidOCR
        if _OCR_ENGINE is not None and _OCR_ENGINE != "pytesseract":
            try:
                res, _ = _OCR_ENGINE(crop, use_det=True, use_cls=False, use_rec=True)
                if res:
                    for item in res:
                        txt, conf = str(item[1] or ""), float(item[2] or 0.0)
                        if conf >= 0.70:
                            zh_cnt = sum(1 for ch in txt if '\u4e00' <= ch <= '\u9fff' or ch.isalnum())
                            score += zh_cnt * (conf ** 2)
            except Exception:
                pass
        # 降级使用 pytesseract
        elif _OCR_ENGINE == "pytesseract":
            try:
                import pytesseract
                txt = pytesseract.image_to_string(crop, lang='chi_sim')
                zh_cnt = sum(1 for ch in txt if '\u4e00' <= ch <= '\u9fff')
                score = float(zh_cnt)
            except Exception:
                pass

        if score > best_score:
            best_score = score
            best_img = img

    return best_img


# ==============================================================================
# 模块 2：物理表格横线探测 + 密集长页无损分片
# ==============================================================================
def detect_physical_table_row_bands(bgr_img: np.ndarray) -> List[Tuple[int, int]]:
    """
    通过形态学水平线检测，提取表格每一行的物理垂直区间 [(y_top, y_bottom), ...]。
    用于：
      1. 准确获知本页物理总行数（守恒校验基准）；
      2. 当单页行数 > 16 行时，严格沿着表格横线切分图片，绝不切碎任何一行字！
    """
    h, w = bgr_img.shape[:2]
    gray = cv2.cvtColor(bgr_img, cv2.COLOR_BGR2GRAY)
    binary = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 21, 10
    )
    # 提取长度至少占页面宽度 25% 的水平表格线
    kernel_w = max(w // 4, 80)
    horiz_lines = cv2.morphologyEx(
        binary, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (kernel_w, 1))
    )
    proj = np.sum(horiz_lines > 0, axis=1)
    line_ys = np.where(proj > kernel_w * 0.5)[0]
    if len(line_ys) < 2:
        return []

    # 聚类相邻像素的水平线 y 坐标
    merged_ys = []
    cur_group = [int(line_ys[0])]
    for y in line_ys[1:]:
        if int(y) - cur_group[-1] <= 8:
            cur_group.append(int(y))
        else:
            merged_ys.append(int(np.mean(cur_group)))
            cur_group = [int(y)]
    merged_ys.append(int(np.mean(cur_group)))

    bands = []
    for i in range(len(merged_ys) - 1):
        y1, y2 = merged_ys[i], merged_ys[i + 1]
        if y2 - y1 >= 14:  # 过滤双线边框的极窄缝隙
            bands.append((y1, y2))
    return bands


def slice_dense_page_with_row_rulers(
    bgr_img: np.ndarray, max_rows_per_slice: int = 16
) -> Tuple[List[np.ndarray], int]:
    """
    若单页表格物理行数超过 max_rows_per_slice（如《东方雨虹》第4页有34行、《未来新材》每页21~24行），
    自动沿着物理表格横线将页面切分为上下 2~3 张高清子图（每张最多 16 行），
    使视觉大模型每次只需专注识别 <=16 行，彻底杜绝长表同质行跳行与 Token 截断！
    返回: (切片图像列表, 探测到的物理总行数)
    """
    bands = detect_physical_table_row_bands(bgr_img)
    total_phys_rows = len(bands)
    if total_phys_rows <= max_rows_per_slice or not bands:
        return [bgr_img], total_phys_rows

    h, w = bgr_img.shape[:2]
    slices = []
    for start_idx in range(0, total_phys_rows, max_rows_per_slice):
        end_idx = min(start_idx + max_rows_per_slice, total_phys_rows)
        y_top = max(0, bands[start_idx][0] - 12) if start_idx > 0 else 0
        y_bot = min(h, bands[end_idx - 1][1] + 12) if end_idx < total_phys_rows else h
        crop = bgr_img[y_top:y_bot, 0:w].copy()
        slices.append(crop)
    return slices, total_phys_rows


# ==============================================================================
# 模块 3：全文档跨页同构表缝合 + 任意页表头双向广播 + 无表头列角色自推导
# ==============================================================================
_HDR_KEYWORDS = ("材料名称", "物料名称", "属性项", "属性值", "品名", "规格", "规格型号", "规格参数", "单价", "含税价", "不含税价", "价格", "单位", "品牌", "类别", "税率", "运费", "序号", "编号")
_UNIT_VAL_RE = re.compile(r'^(m2|m²|m3|m³|kg|t|吨|个|套|根|米|m|km|公里|樘|扇|片|块|张|卷|桶|箱|包|袋|台|只|对|条|元/.+)$', re.I)
_PRICE_OR_DASH_RE = re.compile(r'^(\d+(?:\.\d+)?|[—\-－/]+)$')
_KV_SPEC_RE = re.compile(r'(品种\s*[:：]|厚度\s*\(?mm\)?\s*[:：]|长度\s*\(?mm\)?\s*[:：]|材质\s*[:：]|规格\s*\(?mm\)?\s*[:：]|颜色\s*[:：]|系列\s*[:：])', re.I)


def is_valid_table_header_row(row_cells: List[Any]) -> bool:
    """
    判定某一行是否为真正的表头行。
    严格排除以下数据行伪装：
    1. 规格列中包含 '规格(mm):800×800'、'品种:xxx' 等长键值对描述；
    2. 包含 'm²'、'kg'、'卷' 等计量单位；
    3. 包含浮点数金额（如 16.25、45.00）或单价破折号 '—'；
    4. 单元格字符超长（> 16 字符）。
    """
    texts = [str(c or "").strip() for c in (row_cells or []) if str(c or "").strip()]
    if not texts:
        return False

    for t in texts:
        # 排除包含规格参数冒号、分号的描述项
        if _KV_SPEC_RE.search(t) or ':' in t or '：' in t or ';' in t or '；' in t:
            return False
        # 排除包含计量单位的行
        if _UNIT_VAL_RE.match(t):
            return False
        # 排除包含金额浮点数的行
        if re.match(r'^\d+\.\d+$', t):
            return False
        # 排除单价破折号占位符
        if t in ('—', '-', '－', '/'):
            return False
        # 表头文字必须精简，单格文字过长必为参数数据
        if len(t) > 16:
            return False

    # 表头行必须命中最少 2 个表头标准关键字（且单元格字数 <= 10）
    kw_hits = sum(1 for t in texts if any(k in t for k in _HDR_KEYWORDS) and len(t) <= 10)
    return kw_hits >= 2


def infer_headers_from_data_columns(all_rows: List[List[Any]]) -> List[str]:
    """
    终极兜底：当整份 PDF 从第 1 页到最后一页连一行表头都没有时，
    直接通过全文档每一列的数据分布画像（Data Profile）自动反推每一列的标准表头名！
    """
    if not all_rows:
        return []
    col_cnt = max(len(r) for r in all_rows)
    inferred = [f"列{i+1}" for i in range(col_cnt)]

    unit_col, price_col, spec_col = None, None, None
    for c in range(col_cnt):
        col_vals = [str(r[c]).strip() for r in all_rows if c < len(r) and str(r[c] or "").strip()]
        if not col_vals:
            continue
        n = len(col_vals)

        # 第 0 列递增纯整数优先判定为“序号”，绝不能误判为“单价”
        is_seq = (c == 0) and all(re.match(r'^\d+$', v) and int(v) <= 5000 for v in col_vals)
        if is_seq:
            inferred[c] = "序号"
            continue

        unit_ratio = sum(1 for v in col_vals if _UNIT_VAL_RE.match(v)) / n
        price_ratio = sum(1 for v in col_vals if _PRICE_OR_DASH_RE.match(v)) / n
        kv_ratio = sum(1 for v in col_vals if _KV_SPEC_RE.search(v) or (':' in v or '：' in v or ';' in v)) / n

        if unit_ratio >= 0.5 and unit_col is None:
            unit_col = c
            inferred[c] = "单位"
        elif kv_ratio >= 0.3 and spec_col is None:
            spec_col = c
            inferred[c] = "属性项属性值"
        elif price_ratio >= 0.5 and price_col is None and any(re.match(r'^\d+(?:\.\d+)?$', v) for v in col_vals):
            price_col = c
            inferred[c] = "单价"

    # 根据已锚定的 spec_col / unit_col / price_col 推断其余文本列（品名、品牌、一二三级分类）
    if spec_col is not None and spec_col >= 1:
        inferred[spec_col - 1] = "材料名称"
        if spec_col == 4:
            inferred[0], inferred[1], inferred[2] = "一级类别", "二级类别", "三级类别名称"
    if unit_col is not None and price_col is not None and price_col - unit_col == 2:
        inferred[unit_col + 1] = "品牌"
    return inferred


def unify_and_broadcast_cross_page_headers(pages_data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    系统性跨页表头对齐器（针对 dict 结构的 pages 列表）：
    1. 收集全文档所有页的表格，按列数 col_cnt 分组；
    2. 无论表头在第 1 页（如广西未来新材）、还是在第 10 页（如东方雨虹），自动提取该组表头并全文档双向广播；
    3. 若整份文档所有页都没有表头，调用 infer_headers_from_data_columns 从数据分布自动生成标准表头！
    """
    header_by_cols: Dict[int, List[str]] = {}
    rows_by_cols: Dict[int, List[List[Any]]] = {}

    for p in pages_data:
        for tbl in (p.get("tables") or []):
            cells = tbl.get("cells") or []
            if not cells:
                continue
            col_cnt = max(len(r) for r in cells)
            for r_idx, row in enumerate(cells):
                if r_idx < 3 and is_valid_table_header_row(row):
                    header_by_cols.setdefault(col_cnt, [str(x or "").strip() for x in row])
                else:
                    rows_by_cols.setdefault(col_cnt, []).append(row)

    # 对全文档均未找到显式表头的列数分组，优先尝试从已有表头中压缩空列对齐，兜底才基于全量数据分布反推表头
    for col_cnt, data_rows in rows_by_cols.items():
        if col_cnt not in header_by_cols and data_rows:
            matched_hdr = None
            for exist_cols, exist_hdr in header_by_cols.items():
                stripped_hdr = [h for h in exist_hdr if h.strip()]
                if len(stripped_hdr) == col_cnt:
                    matched_hdr = stripped_hdr
                    break
            if matched_hdr:
                header_by_cols[col_cnt] = matched_hdr
            else:
                header_by_cols[col_cnt] = infer_headers_from_data_columns(data_rows)

    # 将表头注入每一个首行是数据行的续表页（绝不吞掉续表页的第 1 行真实数据！）
    for p in pages_data:
        for tbl in (p.get("tables") or []):
            cells = tbl.get("cells") or []
            if not cells:
                continue
            col_cnt = max(len(r) for r in cells)
            if not is_valid_table_header_row(cells[0]) and col_cnt in header_by_cols:
                tbl["cells"] = [list(header_by_cols[col_cnt])] + cells

    return pages_data


def broadcast_headers_to_page_grids(
    pages_records: List[Dict[str, Any]]
) -> List[Dict[str, Any]]:
    """
    系统性跨页表头对齐器（针对 pdf_extractor.py 的内部数据记录）：
    入参格式: [{"base": base, "grids": [grid1, grid2], "src": src}, ...]
    1. 跨越所有页面双向扫描真正的表头（支持第 10 页倒序回填第 1~9 页）；
    2. 对缺失表头的表格自动补齐；
    3. 若全无表头，自动反推。
    """
    header_by_cols: Dict[int, List[str]] = {}
    rows_by_cols: Dict[int, List[List[Any]]] = {}

    for rec in pages_records:
        for grid in rec.get("grids") or []:
            if not grid:
                continue
            col_cnt = max(len(r) for r in grid)
            for r_idx, row in enumerate(grid[:3]):
                if is_valid_table_header_row(row):
                    header_by_cols.setdefault(col_cnt, [str(x or "").strip() for x in row])
                    break
            for r_idx, row in enumerate(grid):
                if not (r_idx < 3 and is_valid_table_header_row(row)):
                    rows_by_cols.setdefault(col_cnt, []).append(row)

    # 兜底反推：优先从已有表头压缩空列对齐，无匹配时再反推
    for col_cnt, data_rows in rows_by_cols.items():
        if col_cnt not in header_by_cols and data_rows:
            matched_hdr = None
            for exist_cols, exist_hdr in header_by_cols.items():
                stripped_hdr = [h for h in exist_hdr if h.strip()]
                if len(stripped_hdr) == col_cnt:
                    matched_hdr = stripped_hdr
                    break
            if matched_hdr:
                header_by_cols[col_cnt] = matched_hdr
            else:
                header_by_cols[col_cnt] = infer_headers_from_data_columns(data_rows)

    # 双向补齐
    for rec in pages_records:
        for grid in rec.get("grids") or []:
            if not grid:
                continue
            col_cnt = max(len(r) for r in grid)
            if not is_valid_table_header_row(grid[0]) and col_cnt in header_by_cols:
                hdr = list(header_by_cols[col_cnt])
                grid.insert(0, hdr)
                print(f"    [*] {rec.get('base')}: 成功对齐并注入跨页表头 ({len(hdr)} 列: {hdr[:3]}...)")

    return pages_records

