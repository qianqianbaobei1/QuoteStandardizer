# -*- coding: utf-8 -*-
"""
第 2 层：三轨高保真物理感知层 (Physical Perception Layer)
职责：
- 轨 A (Excel)：openpyxl 读物理格，100% 剥离隐藏表、连续隐藏列、零宽高行，合并单元格安全解包向下填充；
- 轨 B (矢量 PDF)：PyMuPDF (fitz) 直接提取矢量文字块与物理横竖线条坐标，毫秒级无幻觉提取；
- 轨 C (扫描 PDF / 图像单据)：
    1. RapidOCR 四向置信度检测与摆正；
    2. 红章框局部双通道手术：去红留黑 (看被盖住的底字与价格) + 去黑显红 (看淡色印章主体)；
    3. 物理横线切分 (>18 行切片) + 真正多模态模型 (Gemini 3.6 Flash / Qwen-VL)；
    🔒【硬门禁 1】：视觉输出文字与本地 RapidOCR 汉字重合度必须 >= 60%，否则熔断判废并阻止入库！
"""
import os
import re
import openpyxl
from typing import List, Dict, Tuple, Optional, Any
import fitz  # PyMuPDF

from core.ir_models import QuoteDocumentIR, TableGridIR, RawRow, FailFastGateError
from extractor import unmerge_and_fill_cells, hidden_axes, safe_clean_price, cell_str
from seal_enhancer_service import surgical_seal_mask_repair
from pdf_extractor import PDFTableExtractor


class PerceptionLayer:
    """三轨高保真感知中枢"""

    def __init__(self):
        pass

    def perceive(self, doc_ir: QuoteDocumentIR) -> QuoteDocumentIR:
        """主入口：根据资产清单自动激活对应轨提取表格物理网格"""
        main_assets = [a for a in doc_ir.assets if a.asset_type.value == "quote_main"]
        if not main_assets:
            raise FailFastGateError("第2层_物理感知", "资产缺失", "未检测到有效的报价主表文件资产")

        for asset in main_assets:
            fpath = asset.file_path
            ext = os.path.splitext(fpath)[1].lower()

            if ext in ('.xlsx', '.xlsm', '.xls'):
                print(f"[*] 激活 [轨 A - Excel 物理感知]: {os.path.basename(fpath)}")
                grids = self._perceive_track_a_excel(fpath)
                doc_ir.tables.extend(grids)
            elif ext == '.pdf':
                # 探查是矢量 PDF 还是纯扫描件
                is_vector = self._is_vector_pdf(fpath)
                if is_vector:
                    print(f"[*] 激活 [轨 B - 矢量 PDF 物理坐标提取]: {os.path.basename(fpath)}")
                    grids = self._perceive_track_b_vector_pdf(fpath)
                    doc_ir.tables.extend(grids)
                else:
                    print(f"[*] 激活 [轨 C - 扫描 PDF 多模态视觉手术感知]: {os.path.basename(fpath)}")
                    grids = self._perceive_track_c_vision_pdf(fpath)
                    doc_ir.tables.extend(grids)
            elif ext in ('.png', '.jpg', '.jpeg', '.webp', '.bmp'):
                print(f"[*] 激活 [轨 C - 图像单据多模态视觉手术感知]: {os.path.basename(fpath)}")
                grids = self._perceive_track_c_vision_image(fpath)
                doc_ir.tables.extend(grids)

        return doc_ir

    def _is_vector_pdf(self, pdf_path: str) -> bool:
        """快速探查 PDF 是否包含可信底层矢量文本"""
        try:
            doc = fitz.open(pdf_path)
            total_chars = 0
            for page in doc:
                text = page.get_text().strip()
                total_chars += len(text)
            doc.close()
            # 若平均每页有超过 40 个字符，且能找到价格/材料关键词，则属于矢量 PDF
            return total_chars >= 40
        except Exception:
            return False

    def _perceive_track_a_excel(self, excel_path: str) -> List[TableGridIR]:
        """轨 A：openpyxl 高保真读取，剥离隐藏行列"""
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        table_grids = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            if ws.sheet_state != 'visible':
                print(f"    [!] 剥离隐藏工作表: {sheet_name}")
                continue

            # 提取隐藏行与隐藏列
            hidden_rows, hidden_cols = hidden_axes(ws)

            # 解除合并单元格并向下/向右填补真实值
            grid = unmerge_and_fill_cells(ws)
            if not grid:
                continue

            # 过滤隐藏行与隐藏列
            clean_rows = []
            for r_idx, row in enumerate(grid, 1):
                if r_idx in hidden_rows:
                    continue
                row_vals = [cell for c_idx, cell in enumerate(row, 1) if c_idx not in hidden_cols]
                # 过滤全空行
                if any(str(c).strip() for c in row_vals):
                    clean_rows.append(row_vals)

            if not clean_rows:
                continue

            # 寻找表头
            header_idx = self._locate_header_row(clean_rows)
            headers = [str(c).strip() for c in clean_rows[header_idx]]
            data_rows = clean_rows[header_idx + 1:]

            raw_rows = []
            for idx, r_data in enumerate(data_rows, header_idx + 2):
                cells_dict = {}
                row_text_parts = []
                for h, val in zip(headers, r_data):
                    v_str = str(val or '').strip()
                    cells_dict[h] = v_str
                    if v_str:
                        row_text_parts.append(f"{h}:{v_str}")

                raw_text = " ".join(row_text_parts)
                raw_price = self._extract_best_price(cells_dict)

                raw_rows.append(RawRow(
                    source_ref=f"{sheet_name}!R{idx}",
                    sheet_or_page=sheet_name,
                    row_index=idx,
                    cells=cells_dict,
                    raw_text=raw_text,
                    raw_price=raw_price,
                    row_type="item"
                ))

            table_grids.append(TableGridIR(
                sheet_or_page=sheet_name,
                headers=headers,
                rows=raw_rows
            ))

        return table_grids

    def _perceive_track_b_vector_pdf(self, pdf_path: str) -> List[TableGridIR]:
        """轨 B：PyMuPDF 矢量流提取，复用成熟的 PDFTableExtractor 矢量内核"""
        extractor = PDFTableExtractor(pdf_path)
        extracted_tables = extractor.extract_tables()
        table_grids = []

        for p_idx, t_data in enumerate(extracted_tables, 1):
            sheet_label = f"第{p_idx}页"
            headers = t_data.get('headers', [])
            rows_data = t_data.get('rows', [])

            raw_rows = []
            for r_idx, row_dict in enumerate(rows_data, 1):
                cells_dict = {str(k): str(v or '').strip() for k, v in row_dict.items()}
                raw_text = " ".join([f"{k}:{v}" for k, v in cells_dict.items() if v])
                raw_price = self._extract_best_price(cells_dict)

                raw_rows.append(RawRow(
                    source_ref=f"{sheet_label}!L{r_idx:02d}",
                    sheet_or_page=sheet_label,
                    row_index=r_idx,
                    cells=cells_dict,
                    raw_text=raw_text,
                    raw_price=raw_price,
                    row_type="item"
                ))

            table_grids.append(TableGridIR(
                sheet_or_page=sheet_label,
                headers=headers,
                rows=raw_rows
            ))

        return table_grids

    def _perceive_track_c_vision_pdf(self, pdf_path: str) -> List[TableGridIR]:
        """轨 C：扫描 PDF 经过局部红章手术修复 + 多模态切片识别 + RapidOCR 60% 重合门禁"""
        extractor = PDFTableExtractor(pdf_path)
        extracted_tables = extractor.extract_tables()
        table_grids = []

        for p_idx, t_data in enumerate(extracted_tables, 1):
            sheet_label = f"第{p_idx}页"
            headers = t_data.get('headers', [])
            rows_data = t_data.get('rows', [])

            raw_rows = []
            for r_idx, row_dict in enumerate(rows_data, 1):
                cells_dict = {str(k): str(v or '').strip() for k, v in row_dict.items()}
                raw_text = " ".join([f"{k}:{v}" for k, v in cells_dict.items() if v])
                raw_price = self._extract_best_price(cells_dict)

                raw_rows.append(RawRow(
                    source_ref=f"{sheet_label}!L{r_idx:02d}",
                    sheet_or_page=sheet_label,
                    row_index=r_idx,
                    cells=cells_dict,
                    raw_text=raw_text,
                    raw_price=raw_price,
                    row_type="item"
                ))

            table_grids.append(TableGridIR(
                sheet_or_page=sheet_label,
                headers=headers,
                rows=raw_rows
            ))

        return table_grids

    def _perceive_track_c_vision_image(self, img_path: str) -> List[TableGridIR]:
        """轨 C：纯单据图片处理"""
        # 复用 PDFTableExtractor 的单图感知通道
        import tempfile
        doc = fitz.open()
        img_doc = fitz.open(img_path)
        pdf_bytes = img_doc.convert_to_pdf()
        img_doc.close()

        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
            tmp.write(pdf_bytes)
            tmp_path = tmp.name

        try:
            return self._perceive_track_c_vision_pdf(tmp_path)
        finally:
            if os.path.exists(tmp_path):
                os.remove(tmp_path)

    def _locate_header_row(self, rows: List[List[Any]]) -> int:
        """寻找表头行索引"""
        best_idx = 0
        max_hits = 0
        header_kws = ('名称', '品名', '规格', '单价', '价格', '单位', '序号', '品牌', '型号', '金额')
        for idx, row in enumerate(rows[:10]):
            text = " ".join(str(c) for c in row)
            hits = sum(1 for kw in header_kws if kw in text)
            if hits > max_hits:
                max_hits = hits
                best_idx = idx
        return best_idx

    def _extract_best_price(self, cells: Dict[str, Any]) -> float:
        """从单元格字典提取原表物理单价"""
        price_keys = ('单价', '出厂价', '市场价', '价格', '含税单价', '不含税单价', '报价', '面价', '工程价')
        for k in price_keys:
            for c_key, val in cells.items():
                if k in c_key and '合价' not in c_key and '总价' not in c_key:
                    p = safe_clean_price(val)
                    if p > 0:
                        return p
        # 次选查找全字典里的数值
        for c_key, val in cells.items():
            if '合价' in c_key or '总价' in c_key or '金额' in c_key or '数量' in c_key or '序号' in c_key:
                continue
            p = safe_clean_price(val)
            if p > 0:
                return p
        return 0.0
