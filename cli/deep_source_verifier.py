# -*- coding: utf-8 -*-
"""
cli/deep_source_verifier.py —— 原文件四通道深潜核验引擎 (Peer-to-Peer Deep Source Referee)

核心原则：
1. 对等零偏见原则：不预设文件A或文件B谁是标准答案，双方都只是待审候选人；
2. 四大新技术通道深入核验：
   • 通道 1 (物理层直取)：针对 Excel 物理单元格与原生 PDF 矢量文字流，100% 确定真值；
   • 通道 2 (多光谱微距取证)：四向正立摆正 + RapidOCR 行坐标锚定 + 2.5倍放大 + 原色/脱红章双光谱对比拼图 + VLM 盲读复核；
   • 通道 3 (知识图谱与算术定理)：针对派生字段（工程价折扣公式、377万规格词库分类编码、属性出处溯源）；
   • 通道 4 (诚实兜底转人工)：置信度 < 0.80 或原图损毁模糊时，绝不强行乱判，强制输出 NEED_HUMAN_VERIFY 并附微距切片。
"""
import base64
import json
import os
import re
import cv2
import fitz
import numpy as np
import openpyxl
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, Any, Tuple, List, Optional

from universal_pdf_grid_engine import auto_orient_page_upright, _OCR_ENGINE
from seal_enhancer_service import SealEnhancerService
from category_attribute_service import (
    resolve_and_verify_category,
    validate_and_normalize_specs,
    fetch_category_attributes,
    _CODE_MAP,
)
from config import VISION_API_KEY, VISION_BASE_URL, VISION_MODEL

_SEAL_ENHANCER = SealEnhancerService()

# 手册第二章 C 类（外部商务配置）：原图未体现时一律 PENDING，不得判任何一方错
C_CLASS_FIELDS = {
    '增值税税率', '工程价折扣比例', '折扣系列', '折扣下限', '折扣上限',
    '价格填报类型', '报价地区省', '报价地区市'
}

REFEREE_BLIND_PROMPT = """【报价单质量 · 双光谱微距取证审判庭】
你正在审查一张争议物料行的【3倍微距双光谱取证证据卡】（上方为LAB去阴影原色图，下方为R通道脱红章黑字增强图）。

审理原则：
1. 唯一依据是你在该微距截图中看到的真实文字与数字；
2. 标准值和线上值都可能错（可能是OCR识别错、错行串行、无中生有算错）；
3. 若截图模糊受损、字迹难辨、或无法确定是哪一行，置信度请给低分（< 0.80）。

请先独立读出图片中该字段的真实值（blind_truth），再对比候选值A（标准）与候选值B（线上）：
- "ONLINE_WIN"：图片明确支持线上值 B（标准值A有错）
- "STD_WIN"：图片明确支持标准值 A（线上值B有错）
- "BOTH_WRONG"：图片明确显示两个候选值都与真实数值不符
- "NEED_HUMAN_VERIFY"：图片严重模糊、反光、无法清晰辨认数字，或置信度 < 0.80

只输出合法 JSON，不要任何其他文字：
{
  "verdict": "ONLINE_WIN|STD_WIN|BOTH_WRONG|NEED_HUMAN_VERIFY",
  "confidence": 0.95,
  "blind_truth": "从图片读出的确切文字或数字",
  "evidence_location": "在微距图中的具体位置（如：左侧品名列/中下方单价列）",
  "reason": "判断依据说明"
}"""


def _canon_num_str(val: Any) -> Optional[str]:
    """数值字符串归一化（用于精确比对 10500 vs 10,500.00 vs 10500.0）"""
    if val is None or val == "":
        return None
    s = str(val).strip().replace(",", "").replace("¥", "").replace("￥", "").replace(" ", "")
    m = re.search(r'-?\d+(?:\.\d+)?', s)
    if not m:
        return None
    try:
        d = Decimal(m.group(0))
        # 归一化为无冗余尾随零字符串
        return f"{d.normalize():f}"
    except Exception:
        return m.group(0)


class DeepSourceVerifier:
    """利用原文件多技术栈（物理DOM + 3倍双光谱微距视觉 + 377万词库 + 算术定理）核验 A/B 分歧"""

    def __init__(self, evidence_out_dir: str = ""):
        self.evidence_out_dir = evidence_out_dir
        if evidence_out_dir:
            os.makedirs(evidence_out_dir, exist_ok=True)
        self._page_bgr_cache: Dict[Tuple[str, int], np.ndarray] = {}
        self._page_ocr_cache: Dict[Tuple[str, int], List[Any]] = {}
        self._excel_cache: Dict[str, openpyxl.Workbook] = {}
        self._vector_pdf_text_cache: Dict[Tuple[str, int], List[str]] = {}

    def load_upright_page_bgr(self, src_path: str, page_no: int = 1) -> Optional[np.ndarray]:
        """加载原文件指定页，并自动完成 0/90/180/270 四向正立摆正"""
        key = (src_path, page_no)
        if key in self._page_bgr_cache:
            return self._page_bgr_cache[key]
        if not src_path or not os.path.exists(src_path):
            return None
        try:
            ext = os.path.splitext(src_path)[1].lower()
            if ext == ".pdf":
                with fitz.open(src_path) as doc:
                    idx = max(0, min(int(page_no or 1) - 1, len(doc) - 1))
                    pix = doc[idx].get_pixmap(dpi=220)
                    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
                    bgr = cv2.cvtColor(arr, cv2.COLOR_RGBA2BGR if pix.n == 4 else cv2.COLOR_RGB2BGR)
            else:
                img_bytes = np.fromfile(src_path, dtype=np.uint8)
                bgr = cv2.imdecode(img_bytes, cv2.IMREAD_COLOR)
            if bgr is not None:
                bgr = auto_orient_page_upright(bgr)
                self._page_bgr_cache[key] = bgr
            return bgr
        except Exception:
            return None

    # -------------------------------------------------------------------------
    # 通道 1：原文件物理层直取 (Excel / 矢量 PDF)
    # -------------------------------------------------------------------------
    def verify_channel1_physical(
        self, src_path: str, page_no: int, field: str, val_a: Any, val_b: Any, row_context: Dict[str, Any]
    ) -> Optional[Dict[str, Any]]:
        """直接读取底层物理单元格或矢量文本流，只要有一方与底层物理字节100%吻合即直接定案"""
        if not src_path or not os.path.exists(src_path):
            return None

        ext = os.path.splitext(src_path)[1].lower()
        num_a = _canon_num_str(val_a)
        num_b = _canon_num_str(val_b)
        str_a = str(val_a or "").strip()
        str_b = str(val_b or "").strip()

        anchors = [
            str(row_context.get("name") or row_context.get("名称规格") or "").strip(),
            str(row_context.get("raw_spec") or row_context.get("规格参数") or "").strip(),
        ]
        clean_anchors = [a for a in anchors if len(a) >= 2]

        # 1.1 若原文件是 Excel 格式
        if ext in (".xlsx", ".xlsm", ".xltx"):
            try:
                if src_path not in self._excel_cache:
                    self._excel_cache[src_path] = openpyxl.load_workbook(src_path, data_only=True)
                wb = self._excel_cache[src_path]
                for ws in wb.worksheets:
                    # 遍历行
                    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
                        row_vals = [str(c).strip() for c in row if c is not None]
                        row_line = " ".join(row_vals)
                        # 检查该行是否为目标物料行
                        if clean_anchors and any(a in row_line for a in clean_anchors):
                            # 在该行单元格中精准查找
                            hit_a = False
                            hit_b = False
                            hit_cell_val = ""
                            for c in row:
                                if c is None:
                                    continue
                                c_str = str(c).strip()
                                c_num = _canon_num_str(c)
                                if num_a and c_num == num_a:
                                    hit_a = True
                                    hit_cell_val = c_str
                                if num_b and c_num == num_b:
                                    hit_b = True
                                    hit_cell_val = c_str
                                if str_a and (c_str == str_a or str_a in c_str):
                                    hit_a = True
                                    hit_cell_val = c_str
                                if str_b and (c_str == str_b or str_b in c_str):
                                    hit_b = True
                                    hit_cell_val = c_str

                            if hit_a and not hit_b:
                                return {
                                    "verdict": "STD_WIN",
                                    "confidence": 1.0,
                                    "ground_truth_from_image": hit_cell_val or str_a,
                                    "truth": hit_cell_val or str_a,
                                    "evidence_location": f"原Excel工作表[{ws.title}]第{row_idx}行物理单元格",
                                    "std_fix_reason": "",
                                    "channel": "CH1_EXCEL_PHYSICAL",
                                }
                            elif hit_b and not hit_a:
                                return {
                                    "verdict": "ONLINE_WIN",
                                    "confidence": 1.0,
                                    "ground_truth_from_image": hit_cell_val or str_b,
                                    "truth": hit_cell_val or str_b,
                                    "evidence_location": f"原Excel工作表[{ws.title}]第{row_idx}行物理单元格",
                                    "std_fix_reason": f"原Excel物理格为{hit_cell_val or str_b}，标准值误填为{str_a}",
                                    "channel": "CH1_EXCEL_PHYSICAL",
                                }
            except Exception:
                pass

        # 1.2 若原文件是含文字图层的原生矢量 PDF
        if ext == ".pdf":
            try:
                key = (src_path, page_no)
                if key not in self._vector_pdf_text_cache:
                    with fitz.open(src_path) as doc:
                        idx = max(0, min(int(page_no or 1) - 1, len(doc) - 1))
                        txt = doc[idx].get_text("text") or ""
                        lines = [l.strip() for l in txt.split("\n") if l.strip()]
                        self._vector_pdf_text_cache[key] = lines
                pdf_lines = self._vector_pdf_text_cache.get(key, [])
                if pdf_lines:
                    # 匹配包含锚点词的文本行
                    for l_idx, pline in enumerate(pdf_lines):
                        if clean_anchors and any(a in pline for a in clean_anchors):
                            # 观察周边 3 行文字
                            block_slice = " ".join(pdf_lines[max(0, l_idx - 1) : min(len(pdf_lines), l_idx + 3)])
                            hit_a = (str_a and str_a in block_slice) or (num_a and num_a in block_slice)
                            hit_b = (str_b and str_b in block_slice) or (num_b and num_b in block_slice)
                            if hit_a and not hit_b:
                                return {
                                    "verdict": "STD_WIN",
                                    "confidence": 1.0,
                                    "ground_truth_from_image": str_a,
                                    "truth": str_a,
                                    "evidence_location": f"原PDF第{page_no}页矢量文字流",
                                    "std_fix_reason": "",
                                    "channel": "CH1_PDF_VECTOR",
                                }
                            elif hit_b and not hit_a:
                                return {
                                    "verdict": "ONLINE_WIN",
                                    "confidence": 1.0,
                                    "ground_truth_from_image": str_b,
                                    "truth": str_b,
                                    "evidence_location": f"原PDF第{page_no}页矢量文字流",
                                    "std_fix_reason": f"原PDF矢量字节流明确包含{str_b}，标准答案有误",
                                    "channel": "CH1_PDF_VECTOR",
                                }
            except Exception:
                pass

        return None

    # -------------------------------------------------------------------------
    # 通道 3：知识图谱 + 377万词库 + 算术定理联审 (B 类派生字段)
    # -------------------------------------------------------------------------
    def verify_channel3_derived(
        self, field: str, val_a: Any, val_b: Any, row_context: Dict[str, Any]
    ) -> Optional[Dict[str, Any]]:
        """针对 B 类规则派生字段（工程价 / 类别编码 / 规格参数）进行公式定理与知识库硬核验算"""
        val_a_str = str(val_a or "").strip()
        val_b_str = str(val_b or "").strip()

        # 3.1 工程价算术定理核验：工程价 == 市场价 * 折扣比例
        if field == "工程价":
            try:
                mkt_raw = row_context.get("market_price") or row_context.get("市场价") or row_context.get("true_market_price")
                disc_raw = row_context.get("discount_rate") or row_context.get("工程价折扣比例")
                if mkt_raw and disc_raw:
                    mkt = Decimal(str(_canon_num_str(mkt_raw) or mkt_raw))
                    disc = Decimal(str(_canon_num_str(disc_raw) or disc_raw))
                    if disc > Decimal("1") and disc <= Decimal("10"):
                        disc = disc / Decimal("10")
                    elif disc > Decimal("10"):
                        disc = disc / Decimal("100")
                    if mkt > Decimal("0") and Decimal("0") < disc <= Decimal("1"):
                        expected = (mkt * disc).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                        da_num = Decimal(str(_canon_num_str(val_a_str) or "-9999"))
                        db_num = Decimal(str(_canon_num_str(val_b_str) or "-9999"))
                        da_ok = abs(da_num - expected) <= Decimal("0.02")
                        db_ok = abs(db_num - expected) <= Decimal("0.02")
                        if da_ok and not db_ok:
                            return {
                                "verdict": "STD_WIN",
                                "confidence": 1.0,
                                "ground_truth_from_image": str(expected),
                                "truth": str(expected),
                                "evidence_location": "算术公式计算(市场价×折扣比例)",
                                "std_fix_reason": "",
                                "channel": "CH3_MATH_THEOREM",
                            }
                        elif db_ok and not da_ok:
                            return {
                                "verdict": "ONLINE_WIN",
                                "confidence": 1.0,
                                "ground_truth_from_image": str(expected),
                                "truth": str(expected),
                                "evidence_location": "算术公式计算(市场价×折扣比例)",
                                "std_fix_reason": f"标准工程价填错({val_a_str})，按公式({mkt}×{disc})应为{expected}",
                                "channel": "CH3_MATH_THEOREM",
                            }
                        elif not da_ok and not db_ok:
                            return {
                                "verdict": "BOTH_WRONG",
                                "confidence": 1.0,
                                "ground_truth_from_image": str(expected),
                                "truth": str(expected),
                                "evidence_location": "算术公式计算(市场价×折扣比例)",
                                "std_fix_reason": f"双方工程价均算错({val_a_str} vs {val_b_str})，按公式({mkt}×{disc})应为{expected}",
                                "channel": "CH3_MATH_THEOREM",
                            }
            except Exception:
                pass

        # 3.2 类别编码核验：调用 39.8 万品名词库 + 377 万规格词库
        if field == "类别编码":
            name = str(row_context.get("name") or row_context.get("名称规格") or "")
            raw_spec = str(row_context.get("raw_spec") or row_context.get("规格参数") or "")
            _, _, best_l3, best_code, is_ok, reason = resolve_and_verify_category("", name, raw_spec)
            if is_ok and best_code:
                ca = val_a_str.replace("-", "")
                cb = val_b_str.replace("-", "")
                target_c = best_code.replace("-", "")
                if ca == target_c and cb != target_c:
                    return {
                        "verdict": "STD_WIN",
                        "confidence": 0.95,
                        "ground_truth_from_image": best_code,
                        "truth": best_code,
                        "evidence_location": f"广材377万词库与工料机字典精准判定({best_l3})",
                        "std_fix_reason": "",
                        "channel": "CH3_LEXICON_KNOWLEDGE",
                    }
                elif cb == target_c and ca != target_c:
                    return {
                        "verdict": "ONLINE_WIN",
                        "confidence": 0.95,
                        "ground_truth_from_image": best_code,
                        "truth": best_code,
                        "evidence_location": f"广材377万词库与工料机字典精准判定({best_l3})",
                        "std_fix_reason": f"广材词表实盘确定为 {best_code}({best_l3})，标准答案分类有误",
                        "channel": "CH3_LEXICON_KNOWLEDGE",
                    }
                elif ca != target_c and cb != target_c and is_ok:
                    return {
                        "verdict": "BOTH_WRONG",
                        "confidence": 0.95,
                        "ground_truth_from_image": best_code,
                        "truth": best_code,
                        "evidence_location": f"广材377万词库与工料机字典精准判定({best_l3})",
                        "std_fix_reason": f"双方编码均与词库实盘不符({val_a_str} vs {val_b_str})，正解为 {best_code}({best_l3})",
                        "channel": "CH3_LEXICON_KNOWLEDGE",
                    }

        # 3.3 规格参数出处溯源与幻觉排查
        if field.startswith("规格参数") or field in ("参数", "规格"):
            raw_spec = str(row_context.get("raw_spec") or row_context.get("规格参数") or "")
            if raw_spec:
                # 检查属性值是否有字符出处（幻觉排查 L3-HALLU）
                has_a = bool(val_a_str and (val_a_str in raw_spec or _canon_num_str(val_a_str) in raw_spec))
                has_b = bool(val_b_str and (val_b_str in raw_spec or _canon_num_str(val_b_str) in raw_spec))
                if has_a and not has_b:
                    return {
                        "verdict": "STD_WIN",
                        "confidence": 0.90,
                        "ground_truth_from_image": val_a_str,
                        "truth": val_a_str,
                        "evidence_location": "原表原始规格文本出处溯源",
                        "std_fix_reason": "",
                        "channel": "CH3_PROVENANCE_CHECK",
                    }
                elif has_b and not has_a:
                    return {
                        "verdict": "ONLINE_WIN",
                        "confidence": 0.90,
                        "ground_truth_from_image": val_b_str,
                        "truth": val_b_str,
                        "evidence_location": "原表原始规格文本出处溯源",
                        "std_fix_reason": f"线上值有原文依据({val_b_str})，标准值无原文出处({val_a_str})",
                        "channel": "CH3_PROVENANCE_CHECK",
                    }

        return None

    # -------------------------------------------------------------------------
    # 通道 2：多光谱微距切片生成 (Micro-Crop Forensic Panel)
    # -------------------------------------------------------------------------
    def build_forensic_micro_crop(
        self,
        src_path: str,
        page_no: int,
        anchor_texts: List[str],
        file_id: str = "",
        field: str = "",
    ) -> Tuple[Optional[np.ndarray], str, str]:
        """
        新技术取证：
        1. 用 RapidOCR 在正立摆正后的原图上定位争议行所在的纵坐标带 [y1, y2]；
        2. 局部截取该窄条并进行 2.0x 放大；
        3. 拼接生成【上：LAB去阴影原色图】+【下：R通道脱红章黑字增强图】双光谱对比图；
        4. 保存至 evidence 目录以供证据溯源与人工核验。
        """
        bgr = self.load_upright_page_bgr(src_path, page_no)
        if bgr is None:
            return None, "无法加载原图", ""

        h, w = bgr.shape[:2]
        key = (src_path, page_no)
        if key not in self._page_ocr_cache and _OCR_ENGINE is not None:
            try:
                res, _ = _OCR_ENGINE(bgr)
                self._page_ocr_cache[key] = res or []
            except Exception:
                self._page_ocr_cache[key] = []

        ocr_boxes = self._page_ocr_cache.get(key, [])
        matched_ys = []
        clean_anchors = [re.sub(r'\s+', '', str(t))[:12] for t in anchor_texts if str(t or "").strip()]

        for item in ocr_boxes:
            box, txt = item[0], re.sub(r'\s+', '', str(item[1] or ""))
            if any(a and (a in txt or txt in a) for a in clean_anchors if len(a) >= 2):
                ys = [pt[1] for pt in box]
                matched_ys.extend(ys)

        if matched_ys:
            y_center = int(np.median(matched_ys))
            pad = max(90, int(h * 0.07))
            y1, y2 = max(0, y_center - pad), min(h, y_center + pad)
            loc_desc = f"原图第{page_no}页纵向坐标区[{y1}:{y2}]"
        else:
            y1, y2 = 0, h
            loc_desc = f"原图第{page_no}页全幅"

        crop_bgr = bgr[y1:y2, 0:w].copy()
        if (y2 - y1) < h * 0.35:
            crop_bgr = cv2.resize(crop_bgr, (0, 0), fx=2.0, fy=2.0, interpolation=cv2.INTER_CUBIC)

        lab_bgr = _SEAL_ENHANCER.balance_illumination_and_orient(crop_bgr, do_orient=False)
        deseal = _SEAL_ENHANCER.penetrate_and_enhance_text(crop_bgr)
        if len(deseal.shape) == 2:
            deseal_bgr = cv2.cvtColor(deseal, cv2.COLOR_GRAY2BGR)
        else:
            deseal_bgr = deseal

        forensic_panel = np.vstack([lab_bgr, deseal_bgr])

        # 保存切片
        saved_relpath = ""
        if self.evidence_out_dir:
            try:
                safe_fid = re.sub(r'[\\/:*?"<>|]', '_', str(file_id or 'evidence'))
                safe_fld = re.sub(r'[\\/:*?"<>|]', '_', str(field or 'diff'))
                fname = f"{safe_fid}_p{page_no}_{safe_fld}.png"
                fpath = os.path.join(self.evidence_out_dir, fname)
                cv2.imencode(".png", forensic_panel)[1].tofile(fpath)
                saved_relpath = os.path.join("evidence", fname)
            except Exception:
                pass

        return forensic_panel, loc_desc, saved_relpath

    # -------------------------------------------------------------------------
    # 通道 2 + 4：微距视觉双模盲读与置信度门控
    # -------------------------------------------------------------------------
    def verify_channel2_visual(
        self,
        panel_bgr: np.ndarray,
        task: Dict[str, Any],
        loc_desc: str,
        saved_crop_relpath: str = "",
    ) -> Dict[str, Any]:
        """将微距证据卡发送给视觉大模型进行盲读与复核，若置信度 < 0.80 自动转入人工核验"""
        undecided = {
            "verdict": "UNADJUDICATED",
            "confidence": 0.0,
            "ground_truth_from_image": "",
            "evidence_location": loc_desc,
            "std_fix_reason": "",
            "channel": "CH2_MICRO_VISION",
            "evidence_crop_path": saved_crop_relpath,
        }
        if not VISION_API_KEY or not VISION_BASE_URL or not VISION_MODEL:
            return dict(undecided, std_fix_reason="未配置视觉模型(VISION_*)")

        success, enc = cv2.imencode(".png", panel_bgr)
        if not success:
            return dict(undecided, std_fix_reason="微距图编码失败")
        b64_img = base64.b64encode(enc.tobytes()).decode("utf-8")

        prompt = (
            f"【待裁决争议物料】\n"
            f"- 原图位置：{loc_desc}\n"
            f"- 品名规格：品名={task.get('name', '')} | 原始规格={task.get('raw_spec', '')}\n"
            f"- 争议字段：【{task['field']}】\n"
            f"- 候选值 A (标准值)：\"{task['standard']}\"\n"
            f"- 候选值 B (线上值)：\"{task['online']}\"\n\n"
            f"请看图片中的微距特写（上方原色，下方脱红章增强黑字），独立读出真值并给出裁决与置信度。"
        )
        payload = {
            "model": VISION_MODEL,
            "temperature": 0,
            "max_tokens": 450,
            "messages": [
                {"role": "system", "content": REFEREE_BLIND_PROMPT},
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": prompt},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64_img}"}},
                    ],
                },
            ],
        }

        try:
            import requests
            resp = requests.post(
                f"{VISION_BASE_URL.rstrip('/')}/chat/completions",
                headers={"Authorization": f"Bearer {VISION_API_KEY}"},
                json=payload,
                timeout=90,
            )
            if resp.status_code != 200:
                return dict(undecided, std_fix_reason=f"裁判接口 HTTP {resp.status_code}")
            content = resp.json()["choices"][0]["message"]["content"]
            m = re.search(r"\{.*\}", content, re.S)
            parsed = json.loads(m.group(0)) if m else {}
            verdict = str(parsed.get("verdict") or "").upper()
            conf = float(parsed.get("confidence") or 0.0)
            blind_truth = str(parsed.get("blind_truth") or "").strip()
            reason = str(parsed.get("reason") or "").strip()

            # 置信度门控（< 0.80 诚实转入人工核验专席，绝不乱猜）
            if conf < 0.80 or verdict == "NEED_HUMAN_VERIFY":
                return {
                    "verdict": "NEED_HUMAN_VERIFY",
                    "confidence": conf,
                    "ground_truth_from_image": blind_truth,
                    "evidence_location": loc_desc,
                    "std_fix_reason": f"视觉置信度不足({conf:.2f})，转人工核验: {reason}",
                    "channel": "CH4_HUMAN_REVIEW",
                    "evidence_crop_path": saved_crop_relpath,
                }

            if verdict not in ("ONLINE_WIN", "STD_WIN", "BOTH_WRONG"):
                verdict = "NEED_HUMAN_VERIFY"

            return {
                "verdict": verdict,
                "confidence": conf,
                "ground_truth_from_image": blind_truth,
                "evidence_location": loc_desc,
                "std_fix_reason": f"原图微距盲读为'{blind_truth}'(置信度{conf:.2f}): {reason}" if verdict == "ONLINE_WIN" else reason,
                "channel": "CH2_MICRO_VISION",
                "evidence_crop_path": saved_crop_relpath,
            }
        except Exception as e:
            return dict(undecided, std_fix_reason=f"视觉裁决调用异常: {e}")

    # -------------------------------------------------------------------------
    # 统一调度总入口：四大通道链式核验
    # -------------------------------------------------------------------------
    def adjudicate_task(self, task: Dict[str, Any], cache: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """单处分歧的四通道对等核验"""
        # 手册公理 2：C 类外部配置字段原图未体现时直接 PENDING，不冤枉任何一方
        if task.get("pre_verdict"):
            return {
                "verdict": task["pre_verdict"],
                "confidence": 1.0,
                "ground_truth_from_image": "",
                "evidence_location": "C类外部配置字段，原图未体现",
                "std_fix_reason": "",
                "channel": "CH3_PRE_RULE",
                "evidence_crop_path": "",
            }

        sig = json.dumps(
            [task.get("source_file"), task.get("page"), task.get("field"), str(task.get("online")), str(task.get("standard"))],
            ensure_ascii=False,
        )
        if cache is not None and sig in cache:
            return cache[sig]

        src_file = task.get("source_file")
        page_no = int(task.get("page") or 1)
        field = task.get("field") or ""
        val_a = task.get("standard")
        val_b = task.get("online")
        row_ctx = task.get("row_context") or {}

        # 关卡 1：原文件物理层直取 (Excel 单元格 / 矢量 PDF)
        if src_file:
            res_ch1 = self.verify_channel1_physical(src_file, page_no, field, val_a, val_b, row_ctx)
            if res_ch1:
                res_ch1["evidence_crop_path"] = ""
                if cache is not None:
                    cache[sig] = res_ch1
                return res_ch1

        # 关卡 3：知识图谱 + 377万词库 + 算术定理联审 (B 类派生字段)
        res_ch3 = self.verify_channel3_derived(field, val_a, val_b, row_ctx)
        if res_ch3:
            res_ch3["evidence_crop_path"] = ""
            if cache is not None:
                cache[sig] = res_ch3
            return res_ch3

        # 关卡 2 + 关卡 4：微距视觉双模盲读与置信度门控
        if src_file and os.path.exists(src_file):
            anchor_texts = [
                str(task.get("name") or row_ctx.get("name") or "").strip(),
                str(task.get("raw_spec") or row_ctx.get("raw_spec") or "").strip(),
                str(val_a or "").strip(),
                str(val_b or "").strip(),
            ]
            panel_bgr, loc_desc, saved_crop = self.build_forensic_micro_crop(
                src_file, page_no, anchor_texts, file_id=task.get("file_id", ""), field=field
            )
            if panel_bgr is not None:
                res_ch2 = self.verify_channel2_visual(panel_bgr, task, loc_desc, saved_crop)
                if cache is not None:
                    cache[sig] = res_ch2
                return res_ch2

        # 兜底：缺失原件，如实记录为未裁决
        undecided = {
            "verdict": "UNADJUDICATED",
            "confidence": 0.0,
            "ground_truth_from_image": "",
            "evidence_location": "无原件文件",
            "std_fix_reason": "缺少原始报价单图片/PDF/Excel，无法核验",
            "channel": "UNADJUDICATED",
            "evidence_crop_path": "",
        }
        if cache is not None:
            cache[sig] = undecided
        return undecided
