# -*- coding: utf-8 -*-
"""
第 3 层：二维表格拓扑解构引擎 (2D Table Topology Engine)
职责：
1. 全文档跨页表头双向广播（跨页缺失表头时，自动继承首页或最近锚定表头）；
2. 矩阵表格展开 (行标题 × 列标题 → 独立物料行，保留列属性维度)；
3. 表内分组小标题状态机（如 "一、管网系统"、"分区A"）与“同上/空白”单元格向下继承；
4. 过滤统计汇总行（小计/合计/总计）与纯注释说明行；
产出：扁平化、带完整物理溯源坐标的原子物料行列表 List[RawRow]。
"""
import re
from typing import List, Dict, Optional, Tuple, Any
from core.ir_models import QuoteDocumentIR, TableGridIR, RawRow

SUMMARY_KEYWORDS = {'合计', '小计', '总计', '总价合计', '本页小计'}
MATRIX_HEADER_INDICATORS = ('壁厚', '厚度', '外径', '口径', '规格', '高度', '截面', '容量', '档位')


class TableTopologyEngine:
    """二维表格拓扑解构引擎"""

    def __init__(self):
        pass

    def deconstruct(self, doc_ir: QuoteDocumentIR) -> QuoteDocumentIR:
        """主入口：解构所有表格网格，统一扁平化为原始物料行列表"""
        if not doc_ir.tables:
            return doc_ir

        # 1. 跨页表头广播
        canonical_headers = self._broadcast_headers(doc_ir.tables)

        flattened_rows: List[RawRow] = []

        # 2. 逐表解构
        for table in doc_ir.tables:
            if not table.headers and canonical_headers:
                table.headers = list(canonical_headers)

            # 判断是否为矩阵式报价表
            is_matrix = self._detect_matrix_table(table)
            table.is_matrix = is_matrix

            if is_matrix:
                matrix_rows = self._unroll_matrix_table(table)
                flattened_rows.extend(matrix_rows)
            else:
                list_rows = self._process_list_table(table)
                flattened_rows.extend(list_rows)

        doc_ir.raw_rows = flattened_rows
        return doc_ir

    def _broadcast_headers(self, tables: List[TableGridIR]) -> List[str]:
        """从所有表格中挑选最完整、列数最多的标准表头，双向广播给缺失表头的表格"""
        best_headers: List[str] = []
        max_score = 0

        header_kws = ('名称', '品名', '规格', '单价', '价格', '单位', '型号')
        for t in tables:
            h_text = " ".join(t.headers)
            hits = sum(1 for kw in header_kws if kw in h_text)
            if hits > max_score and len(t.headers) >= 3:
                max_score = hits
                best_headers = t.headers

        return best_headers

    def _detect_matrix_table(self, table: TableGridIR) -> bool:
        """根据表头是否包含大量数值或特定物理维度判断矩阵表"""
        if len(table.headers) < 4:
            return False
        # 统计表头中包含数值或厚度/壁厚关键词的比例
        numeric_headers = 0
        for h in table.headers[2:]:
            if re.search(r'\d+(\.\d+)?(mm|寸|寸半|分|#|A|V|W)?', h) or any(ind in h for ind in MATRIX_HEADER_INDICATORS):
                numeric_headers += 1
        return numeric_headers >= 3 and (numeric_headers / max(1, len(table.headers) - 2) >= 0.5)

    def _process_list_table(self, table: TableGridIR) -> List[RawRow]:
        """清单式表格：状态机维护当前分组标题，向下传播“同上”与“空白”"""
        res_rows: List[RawRow] = []
        current_section = ""
        last_material_name = ""

        for row in table.rows:
            raw_text = row.raw_text.strip()
            # 1. 跳过汇总行
            if any(raw_text.startswith(kw) or f":{kw}" in raw_text for kw in SUMMARY_KEYWORDS):
                row.row_type = "summary"
                continue

            # 2. 检查是否为分组标题行 (如 "一、管网喷头系统"、"镀锌钢管系列")
            if self._is_section_title(row):
                current_section = self._clean_title_text(raw_text)
                row.row_type = "title"
                continue

            # 3. 关联当前分组
            row.section_title = current_section

            # 4. 同上 / 空白向下继承物料品名
            name_key = self._find_name_key(row.cells)
            if name_key:
                val = str(row.cells.get(name_key, '')).strip()
                if val in ('同上', '"', '〃', '“', '”', '同前') or (not val and row.raw_price > 0 and last_material_name):
                    row.cells[name_key] = last_material_name
                elif val:
                    last_material_name = val

            # 5. 过滤完全没有价格且没有有效规格的纯备注行
            if row.raw_price <= 0 and not self._has_valid_specs(row.cells):
                row.row_type = "note"
                continue

            row.row_type = "item"
            res_rows.append(row)

        return res_rows

    def _unroll_matrix_table(self, table: TableGridIR) -> List[RawRow]:
        """矩阵式表格展开：每一列数值展开为一个独立物料行"""
        res_rows: List[RawRow] = []
        # 前两列通常为行维度 (如：名称、规格)
        dim_headers = table.headers[:2]
        val_headers = table.headers[2:]

        for r_idx, row in enumerate(table.rows, 1):
            base_name = str(row.cells.get(dim_headers[0], '')).strip() if dim_headers else ''
            base_spec = str(row.cells.get(dim_headers[1], '')).strip() if len(dim_headers) > 1 else ''

            if not base_name and not base_spec:
                continue

            for col_idx, col_h in enumerate(val_headers, 3):
                col_val = str(row.cells.get(col_h, '')).strip()
                if not col_val or col_val in ('-', '/', '无', '暂无'):
                    continue

                # 提取单价
                from extractor import safe_clean_price
                price = safe_clean_price(col_val)
                if price <= 0:
                    continue

                # 构造展开后的新行
                unrolled_cells = {
                    dim_headers[0]: base_name,
                    "列维度": col_h,
                    "单价": str(price)
                }
                if base_spec:
                    unrolled_cells[dim_headers[1]] = base_spec

                combined_spec = f"{base_spec} {col_h}".strip()
                unrolled_cells["规格型号"] = combined_spec

                res_rows.append(RawRow(
                    source_ref=f"{table.sheet_or_page}!R{r_idx}C{col_idx}",
                    sheet_or_page=table.sheet_or_page,
                    row_index=r_idx,
                    cells=unrolled_cells,
                    raw_text=f"{base_name} {combined_spec} 单价:{price}",
                    raw_price=price,
                    row_type="item",
                    is_matrix=True,
                    matrix_col_label=col_h
                ))

        return res_rows

    def _is_section_title(self, row: RawRow) -> bool:
        """判定是否为小标题/分组行"""
        # 如果有价格，绝不能作为小标题跳过！
        if row.raw_price > 0:
            return False
        text = row.raw_text
        if re.search(r'^[一二三四五六七八九十0-9]+[、.．\s]+[\u4e00-\u9fa5]{2,15}', text):
            return True
        if any(kw in text for kw in ('系统', '分区', '系列', '分部', '工程', '配件类', '管件类')) and len(text) < 30:
            return True
        return False

    def _clean_title_text(self, text: str) -> str:
        # 去掉键值前缀如 "材料名称:"
        cleaned = re.sub(r'^[^\s:]+[:：]', '', text).strip()
        return cleaned

    def _find_name_key(self, cells: Dict[str, Any]) -> Optional[str]:
        for k in ('名称', '品名', '材料名称', '物料名称', '产品名称'):
            for c_key in cells:
                if k in c_key:
                    return c_key
        return None

    def _has_valid_specs(self, cells: Dict[str, Any]) -> bool:
        for k, v in cells.items():
            if any(sk in k for sk in ('规格', '型号', '材质', '参数', 'DN', '壁厚')) and str(v).strip():
                return True
        return False
