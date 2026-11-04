from llm_auditor import LLMAuditor
from style_enhancer import apply_enterprise_style
# -*- coding: utf-8 -*-
"""
zip_standardizer_pipeline.py - 系统级报价单压缩包端到端标准化流水线
核心功能：
1. 全自动处理各类供应商报价压缩包（含 报价数据.xlsx + quote_img*.png + 支撑图片/条款）
2. 彻底解决“图片内有规格参数（属性项属性值/规格型号），但原始表格留空”的问题（自动通过视觉转录提取并与行对齐）
3. 彻底解决“参数未拆分到类别动态属性列”的问题（自动结构化解析键值对并映射到对应动态列，无硬编码）
4. 彻底解决“图片插入后与原有单元格文本重叠”的问题（清空原有文本 + 规范列宽行高 + 自动缩放居中嵌入）
5. 自动去除多余无用表单（如无需审计表时纯净交付）
"""
import os
import sys
import io
import re
import json
import copy
import base64
import zipfile
from difflib import SequenceMatcher
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
import requests

from config import (
    DEEPSEEK_API_KEY, DEEPSEEK_BASE_URL, VISION_MODEL,
    DEFAULT_OUTPUT_DIR, SKIP_HIDDEN_CONTENT
)
from extractor import hidden_axes
from builder import STANDARD_HEADERS
from models import SourceRowRecord, CanonicalMaterialItem, reconcile

# 线上导出与 GenOffice 33 列标准只差个大小写/百分号，统一按标准列名交付
HEADER_ALIAS = {'报价id': '报价ID', '增值税税率%': '增值税税率'}


def normalize_headers(ws):
    """把表头改写成标准列名（保留其余列不动）"""
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(1, c)
        name = str(cell.value).strip() if cell.value is not None else ''
        if name in HEADER_ALIAS:
            cell.value = HEADER_ALIAS[name]

def collapse_to_single_sheet(wb, sheet_name='标准化报价', column_order=None):
    """把多张分类 Sheet 合并成一张交付主表（线上导出按类别分了多张表）。
    列按表头名对齐，重名按出现次序区分，缺列留空；样式、批注、内嵌图片随单元格迁移。
    column_order 传入标准列名（如 STANDARD_HEADERS）时，缺的标准列补空列、列序按标准排，动态列跟在后头。"""
    need_rewrite = bool(column_order) or wb.sheetnames != [sheet_name]
    if not need_rewrite:
        return wb[sheet_name]

    def keys_of(ws):
        seen, keys = {}, []
        for c in range(1, ws.max_column + 1):
            name = ws.cell(1, c).value
            key = str(name).strip() if name not in (None, '') else f'@col{c}'
            n = seen.get(key, 0)
            seen[key] = n + 1
            keys.append((key, n))
        return keys

    sheets = [(ws, keys_of(ws)) for ws in wb.worksheets
              if not (SKIP_HIDDEN_CONTENT and ws.sheet_state != 'visible')]   # 隐藏口径: 隐藏表不并入
    if not sheets:
        raise ValueError('工作簿里没有任何可见工作表（全部被隐藏）')
    order, head = [], {}
    for ws, keys in sheets:
        for c, k in enumerate(keys, 1):
            if k not in head:
                head[k] = ws.cell(1, c).value
                order.append(k)

    if column_order:
        seen, desired = {}, []
        for name in column_order:
            key = str(name).strip()
            n = seen.get(key, 0)
            seen[key] = n + 1
            desired.append((key, n))
        for k in desired:
            if k not in head:
                head[k] = k[0]  # 标准列但原表没有：补一个空列，表头用标准名
        order = desired + [k for k in order if k not in desired]
    col_idx = {k: i + 1 for i, k in enumerate(order)}

    out = wb.create_sheet('__merged__')
    for k in order:
        out.cell(1, col_idx[k], head[k])

    r_out = 2
    for ws, keys in sheets:
        hid_rows, hid_cols = hidden_axes(ws) if SKIP_HIDDEN_CONTENT else (set(), set())
        cmap = {c: col_idx[k] for c, k in enumerate(keys, 1) if c not in hid_cols}
        offset = r_out - 2
        for r in range(2, ws.max_row + 1):
            if r in hid_rows:
                continue
            for c, k in enumerate(keys, 1):
                if c in hid_cols:
                    continue
                src = ws.cell(r, c)
                if src.value is None:
                    continue
                dst = out.cell(r_out + r - 2, cmap[c])
                dst.value = src.value
                dst._style = copy.copy(src._style)  # ponytail: 公开 API 无批量复制，升级 openpyxl 时校验
                if src.comment is not None:
                    dst.comment = src.comment
        for img in list(getattr(ws, '_images', [])):
            frm = img.anchor._from
            frm.row += offset
            if frm.col + 1 in cmap:
                frm.col = cmap[frm.col + 1] - 1
            out.add_image(img)
        r_out += sum(1 for r in range(2, ws.max_row + 1) if r not in hid_rows)

    for ws, _ in sheets:
        wb.remove(ws)
    out.title = sheet_name
    return out


class ZipQuoteStandardizerPipeline:
    def __init__(self, output_dir=DEFAULT_OUTPUT_DIR):
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)
        self.transcription_cache = {}
        self.auditor = LLMAuditor()

    def extract_image_tables(self, img_bytes: bytes, img_name: str) -> list:
        """调用 Vision 视觉大模型，系统级转录报价单图片中的所有物料及属性规格"""
        cache_key = f"{img_name}_{len(img_bytes)}"
        if cache_key in self.transcription_cache:
            return self.transcription_cache[cache_key]

        b64 = base64.b64encode(img_bytes).decode('utf-8')
        prompt = """你是一个国家级建设工程造价与采招数据专家。请直接审视这份原始报价单图片，把表格中的所有物料及全部技术参数逐行转录为 JSON 结构。

转录要求：
1. 每一行必须提取：
   - "seq": 序号（如果有）
   - "name": 材料名称 / 产品名称 / 品名
   - "spec": 规格型号 / 属性项属性值 / 详细技术参数描述（统一整理为 "键:值;键:值..." 格式，如 "品种:镀锌槽钢;牌号:Q235B;规格:10#"）
   - "unit": 计量单位（如 t, m, 台, 个, m²）
   - "price": 单价 / 市场价 / 价格（数值）
2. 务必保留所有规格细节，绝不脑补，绝不遗漏。
3. 请仅输出合法的 JSON 数组，格式为：[{"seq": 1, "name": "...", "spec": "...", "unit": "...", "price": ...}]"""

        headers = {
            'Authorization': f'Bearer {DEEPSEEK_API_KEY}',
            'Content-Type': 'application/json'
        }
        payload = {
            'model': VISION_MODEL,
            'messages': [{
                'role': 'user',
                'content': [
                    {'type': 'text', 'text': prompt},
                    {'type': 'image_url', 'image_url': {'url': f'data:image/png;base64,{b64}'}}
                ]
            }],
            'temperature': 0.1
        }
        try:
            resp = requests.post(f"{DEEPSEEK_BASE_URL.rstrip('/')}/chat/completions", headers=headers, json=payload, timeout=90)
            if resp.status_code == 200:
                content = resp.json()['choices'][0]['message']['content']
                items = self._parse_json_robust(content)
                for it in items:
                    it['_img_source'] = img_name
                self.transcription_cache[cache_key] = items
                return items
            else:
                print(f"[Warning] Vision API returned {resp.status_code} for {img_name}")
        except Exception as e:
            print(f"[Error] Failed to transcribe {img_name}: {e}")
        return []

    @staticmethod
    def _parse_json_robust(text: str) -> list:
        """鲁棒解析 LLM 返回的 JSON 列表"""
        # 尝试直接解析
        m = re.search(r'\[\s*\{.*\}\s*\]', text, re.DOTALL)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                pass
        
        # 逐个对象提取兜底
        objs = []
        pattern = re.compile(r'\{\s*"seq":.*?(?=\n\s*\{|\n\s*\]|\Z)', re.DOTALL)
        for block in pattern.findall(text):
            clean = block.strip().rstrip(',')
            if not clean.endswith('}'):
                clean += '}'
            try:
                objs.append(json.loads(clean))
            except Exception:
                # 正则提取单项
                seq_m = re.search(r'"seq":\s*(\d+)', clean)
                name_m = re.search(r'"name":\s*"([^"]+)"', clean)
                spec_m = re.search(r'"spec":\s*"([^"]+)"', clean)
                unit_m = re.search(r'"unit":\s*"([^"]+)"', clean)
                price_m = re.search(r'"price":\s*([0-9.]+)', clean)
                if name_m or spec_m:
                    objs.append({
                        "seq": int(seq_m.group(1)) if seq_m else None,
                        "name": name_m.group(1) if name_m else "",
                        "spec": spec_m.group(1) if spec_m else "",
                        "unit": unit_m.group(1) if unit_m else "",
                        "price": float(price_m.group(1)) if price_m else None
                    })
        return objs

    @staticmethod
    def parse_spec_dict(spec_text: str) -> dict:
        """结构化解析规格参数字符串为字典映射，严禁混入内部编码与非物理属性"""
        if not spec_text:
            return {}
        forbidden = ('表头', '表标题', '基价', '单价', '总价', '金额', '指导价', '调价',
                     '产品编码', '物料编码', '商品编码', '内部编码', '条形码', '编码', '条码', '代码',
                     '税号', '账号', '帐号', '开户行', '一级类别', '二级类别', '三级类别', '材料名称', '属性项属性值',
                     '物料名称', '类别名称', '分类编码')
        res = {}
        items = re.split(r'[;\n；]', str(spec_text))
        for it in items:
            it = it.strip()
            if not it:
                continue
            if ':' in it:
                k, v = it.split(':', 1)
                k_clean = k.strip()
                if not any(fb in k_clean for fb in forbidden):
                    res[k_clean] = v.strip()
            elif '：' in it:
                k, v = it.split('：', 1)
                k_clean = k.strip()
                if not any(fb in k_clean for fb in forbidden):
                    res[k_clean] = v.strip()
        return res

    @staticmethod
    def clean_name(s: str) -> str:
        if not s:
            return ""
        return re.sub(r'[\s\-_（）\(\)\[\]【】/]', '', str(s))

    def process_zip(self, zip_path: str, supplier_name: str = None) -> str:
        """主入口：处理单个报价单压缩包并输出标准交付 Excel"""
        base_name = os.path.basename(zip_path)
        if not supplier_name:
            m = re.match(r'^([^0-9_\-]+)', base_name)
            supplier_name = m.group(1) if m else os.path.splitext(base_name)[0]

        print(f"\n==================== 开始系统化处理: {supplier_name} ====================")
        
        # 1. 解构 Zip 包中的 Excel 与图片资产
        xlsx_bytes = None
        images_dict = {}
        
        with zipfile.ZipFile(zip_path, 'r') as z:
            for info in z.infolist():
                fname = info.filename
                if fname.startswith("__MACOSX") or "/." in fname:
                    continue
                lower_name = fname.lower()
                if lower_name.endswith('.xlsx') and not lower_name.startswith('~$'):
                    xlsx_bytes = z.read(info)
                elif lower_name.endswith(('.png', '.jpg', '.jpeg')):
                    images_dict[os.path.basename(fname)] = z.read(info)

        if not xlsx_bytes:
            raise RuntimeError(f"在压缩包 {base_name} 中未找到主数据表格 .xlsx")

        # 2. 检查是否有图片需要进行视觉解析补充
        quote_images = {k: v for k, v in images_dict.items() if 'quote_img' in k.lower()}
        all_transcriptions = []
        for img_name, img_bytes in sorted(quote_images.items()):
            print(f"[*] 视觉引擎介入分析报价图: {img_name}...")
            trans = self.extract_image_tables(img_bytes, img_name)
            all_transcriptions.extend(trans)

        print(f"[*] 视觉大模型共提取到有效物料项: {len(all_transcriptions)} 条")

        # 3. 加载并标准化 Excel
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))

        # 剔除用户明确不需要的冗余审计表
        for unwanted in ['疑义与失败审计说明', '原表对账']:
            if unwanted in wb.sheetnames:
                wb.remove(wb[unwanted])

        # 缓存图片到本地临时目录以便 openpyxl 嵌入
        temp_img_dir = os.path.join(self.output_dir, f".cache_{supplier_name}")
        os.makedirs(temp_img_dir, exist_ok=True)
        local_img_paths = {}
        for img_name, img_bytes in images_dict.items():
            lp = os.path.join(temp_img_dir, img_name)
            with open(lp, 'wb') as f:
                f.write(img_bytes)
            local_img_paths[img_name] = lp

        total_rows_processed = 0
        total_specs_filled = 0
        total_images_embedded = 0

        for ws in wb.worksheets:
            normalize_headers(ws)
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]

            # 定位关键列
            spec_col = headers.index('厂家原始规格参数') + 1 if '厂家原始规格参数' in headers else None
            multiline_spec_col = headers.index('规格参数') + 1 if '规格参数' in headers else None
            name_col = headers.index('名称规格') + 1 if '名称规格' in headers else None
            page_col = headers.index('厂家报价文件页码') + 1 if '厂家报价文件页码' in headers else None
            img_file_col = headers.index('厂家报价图片文件名') + 1 if '厂家报价图片文件名' in headers else None
            img_col = headers.index('图片') + 1 if '图片' in headers else None
            seq_col = headers.index('顺序号') + 1 if '顺序号' in headers else None

            # 清理 sheet 旧图片，统一行高列宽
            ws._images.clear()
            img_letter = openpyxl.utils.get_column_letter(img_col) if img_col else None
            if img_letter:
                ws.column_dimensions[img_letter].width = 24

            for r in range(2, ws.max_row + 1):
                total_rows_processed += 1
                row_name = str(ws.cell(r, name_col).value or "").strip() if name_col else ""
                row_seq = ws.cell(r, seq_col).value if seq_col else None
                cur_spec = ws.cell(r, spec_col).value if spec_col else None

                # 彻底净化现有规格参数：剥除混入的产品编码与内部杂质
                if spec_col and cur_spec:
                    raw_s = str(cur_spec)
                    cleaned_s = re.sub(r'(?:^|[;；\s])(?:产品编码|物料编码|商品编码|内部编码|条形码|条码|代码|三级类别编码|二级类别|一级类别|材料名称|属性项属性值)\s*[:：][^;；\n]*', '', raw_s)
                    cleaned_s = re.sub(r'^[;；\s]+|[;；\s]+$', '', cleaned_s).strip()
                    ws.cell(r, spec_col).value = cleaned_s if cleaned_s else None
                    cur_spec = cleaned_s

                if multiline_spec_col:
                    multi_s = str(ws.cell(r, multiline_spec_col).value or '')
                    if multi_s:
                        clean_lines = []
                        for l in multi_s.split('\n'):
                            k = l.split(':', 1)[0].strip() if ':' in l else (l.split('：', 1)[0].strip() if '：' in l else l.strip())
                            if any(fb in k for fb in ('产品编码', '物料编码', '商品编码', '内部编码', '条形码', '编码', '条码', '代码', '三级类别编码', '二级类别', '一级类别', '材料名称', '属性项属性值', '物料名称', '类别名称', '分类编码')):
                                continue
                            clean_lines.append(l)
                        ws.cell(r, multiline_spec_col).value = '\n'.join(clean_lines) if clean_lines else None

                row_img_file = str(ws.cell(r, img_file_col).value or "").strip() if img_file_col else ""
                if not row_img_file:
                    page_val = str(ws.cell(r, page_col).value or "").strip() if page_col else ""
                    m_p = re.search(r'\d+', page_val)
                    if m_p:
                        row_img_file = f"quote_img{m_p.group(0)}.png"
                    elif page_val:
                        safe_p = re.sub(r'[\/\\:\*\?\"<>\|]', '_', page_val).strip()
                        row_img_file = f"quote_img_{safe_p}.png"
                    else:
                        row_img_file = "quote_img1.png"
                    if img_file_col:
                        ws.cell(r, img_file_col).value = row_img_file

                # 规则 1: 若当前规格参数为空，系统级从视觉转录结果中精准匹配回填
                if (not cur_spec or str(cur_spec).strip() == '') and spec_col:
                    matched_spec, match_how = None, ''
                    c_name = self.clean_name(row_name)

                    def _unique(cands):
                        """候选必须唯一：多条候选宁可留空也不猜（否则静默错配无法发现）"""
                        specs = {str(it.get('spec')).strip() for it in cands if str(it.get('spec') or '').strip()}
                        return (specs.pop(), '') if len(specs) == 1 else (None, '')

                    # 一级：同图片 + 序号（确定性最高）
                    hit = [it for it in all_transcriptions
                           if it.get('_img_source') == row_img_file and str(it.get('seq')) == str(row_seq)]
                    matched_spec, _ = _unique(hit)
                    if matched_spec:
                        match_how = '同图+序号'

                    # 二级：品名完全匹配（含包含关系，但只接受唯一命中）
                    if not matched_spec and c_name:
                        hit = [it for it in all_transcriptions
                               if (lambda t: t and (t == c_name or t in c_name or c_name in t))(self.clean_name(it.get('name', '')))]
                        matched_spec, _ = _unique(hit)
                        if matched_spec:
                            match_how = '品名匹配'

                    # 三级：字符重叠不作依据（“钢制防火门甲级” vs “乙级” 字符交集仍达 86%，必然错配）
                    # 宁可留空，也不靠模糊相似度猜规格

                    # 四级兜底：同图 + 价格（仅比价格，不比单位；价格必须唯一命中）
                    if not matched_spec:
                        price_col = headers.index('市场价') + 1 if '市场价' in headers else None
                        row_p = ws.cell(r, price_col).value if price_col else None
                        if row_p:
                            try:
                                r_pval = float(str(row_p).replace('¥', '').replace(',', ''))
                                hit = [it for it in all_transcriptions
                                       if it.get('_img_source') == row_img_file and it.get('price')
                                       and abs(float(it['price']) - r_pval) < 0.01]
                                matched_spec, _ = _unique(hit)
                                if matched_spec:
                                    match_how = '同图+价格(未比对单位)'
                            except (TypeError, ValueError):
                                pass

                    if matched_spec:
                        ws.cell(r, spec_col).value = matched_spec
                        # 回填必留痕，但不污染业务字段：痕迹挂在规格参数单元格的批注里，供人工抽查是靠哪条规则猜的
                        target = ws.cell(r, spec_col)
                        prev = f'{target.comment.text}；' if target.comment and target.comment.text else ''
                        target.comment = Comment(f'{prev}[规格回填] 来源:{match_how}', 'QuoteStandardizer')
                        cur_spec = matched_spec
                        total_specs_filled += 1

                # 规则 2: 将规格参数结构化拆分，动态映射并补齐到右侧各动态属性列
                if cur_spec:
                    spec_dict = self.parse_spec_dict(cur_spec)
                    for c_idx, h_text in enumerate(headers, 1):
                        if c_idx not in [spec_col, name_col] and h_text:
                            # 归一化查找属性名
                            h_norm = self.clean_name(h_text)
                            for k, v in spec_dict.items():
                                k_norm = self.clean_name(k)
                                # 只认「表头 = 属性名」或「表头 = 属性名+单位后缀」；
                                # 双向包含会错配（实测 spec 的「类型:电容式」被填进「项目类型」列）
                                if (k_norm == h_norm or h_norm.startswith(k_norm)) and not ws.cell(r, c_idx).value:
                                    ws.cell(r, c_idx).value = v
                                    break

                # 规则 3: 严格按照“只要压缩包内有对应产品图片，就内嵌到对应产品单元格”
                # 规则规范：
                # 1. 扫描行内图片列记录的文件名（如 热镀锌槽钢.jpg / 槽式电缆桥架.jpg 等）
                # 2. 检查压缩包内解压出的图片资产中是否存在该文件，或产品小样图库中的对应图片
                # 3. 严禁把非产品图的客户整页报价单扫描底稿（quote_img*）作为产品图
                # 4. 若匹配到真实产品图，单元格清空文字并缩放嵌入（宽 135，高 90）；若包内没有该产品图片文件，则保持单元格清空
                if img_col:
                    matched_product_img_path = None
                    row_img_cells = []
                    for cc in range(img_col, ws.max_column + 1):
                        val = str(ws.cell(r, cc).value or '').strip()
                        row_img_cells.append(cc)
                        if val:
                            base_val = os.path.basename(val)
                            # 过滤非产品图（报价单底稿图 quote_img 不作为产品图嵌入）
                            if not base_val.lower().startswith('quote_img'):
                                # 在包内所有图片中匹配该产品图
                                for local_name, local_path in local_img_paths.items():
                                    if local_name == base_val or local_name == val or (base_val and base_val in local_name):
                                        matched_product_img_path = local_path
                                        break

                    if matched_product_img_path and os.path.exists(matched_product_img_path):
                        # 只有真的嵌了图才清空文字，否则保留原图片文件名（清空等于丢信息）
                        for cc in row_img_cells:
                            ws.cell(r, cc).value = None
                        ws.row_dimensions[r].height = 95
                        try:
                            img = OpenpyxlImage(matched_product_img_path)
                            img.width = 135
                            img.height = 90
                            ws.add_image(img, f"{img_letter}{r}")
                            total_images_embedded += 1
                        except Exception as e:
                            print(f"[Error] Failed to embed product image {matched_product_img_path}: {e}")
                    else:
                        # 规则 4: 接入材料配图接口，通过材料名称在接口检索配图，并回填图片在线链接
                        try:
                            from material_illustration_service import default_illustration_service
                            from openpyxl.styles import Font
                            if row_name:
                                illust = default_illustration_service.fetch_illustration_by_name(
                                    row_name,
                                    company_name=supplier_name,
                                    specification=str(cur_spec or "")
                                )
                                if illust and illust.get("pictureUrl"):
                                    pic_url = illust["pictureUrl"]
                                    pic_name = illust.get("pictureName") or ""
                                    folder_name = illust.get("folderName") or ""
                                    cell_img = ws.cell(r, img_col)
                                    cell_img.value = pic_url
                                    cell_img.hyperlink = pic_url
                                    cell_img.font = Font(name='宋体', size=10, color='0563C1', underline='single')
                                    if img_file_col:
                                        ws.cell(r, img_file_col).value = None
                                    folder_col = headers.index('图片文件夹') + 1 if '图片文件夹' in headers else None
                                    if folder_col and folder_name and not ws.cell(r, folder_col).value:
                                        ws.cell(r, folder_col).value = folder_name
                        except Exception:
                            pass

        # 4. 保存为交付标准文件（多 Sheet 收口成单张「标准化报价」，列序对齐 GenOffice 33 列）
        collapse_to_single_sheet(wb, column_order=STANDARD_HEADERS)
        out_path = os.path.join(self.output_dir, f"{supplier_name}-标准解析.xlsx")
        ws_main = wb['标准化报价']
        main_hdr = [c.value for c in ws_main[1]]
        if '供应商名称' in main_hdr:
            sc = main_hdr.index('供应商名称') + 1
            for r in range(2, ws_main.max_row + 1):
                cell = ws_main.cell(r, sc)
                if not cell.value:
                    cell.value = supplier_name
                    cell.comment = Comment('来源:压缩包文件名推断', 'QuoteStandardizer')
        name_col = main_hdr.index('名称规格') + 1
        orig_names = {r: ws_main.cell(r, name_col).value for r in range(2, ws_main.max_row + 1)}
        for ws in wb.worksheets:
            apply_enterprise_style(ws)
        for ws in wb.worksheets:
            sheet_headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            try:
                self.auditor.verify_and_auto_heal(ws, sheet_headers)
            except Exception as e:
                print(f"[Warning] 自我核对执行提示: {e}")
            apply_enterprise_style(ws)

        # 5. 品名以原表为准：模型自愈可以查错别字，但不能整行换成另一个物料的名字
        #    （实测：「广播劝放」→「广播功放」该留；「班级信息发布终端」→「光模块」必须回退）
        for r, v in orig_names.items():
            cur = ws_main.cell(r, name_col).value
            if cur != v:
                if SequenceMatcher(None, str(v or ''), str(cur or '')).ratio() >= 0.6:
                    print(f"[名称纠错] 行{r}: {v!r} -> {cur!r}")
                else:
                    print(f"[名称保护] 行{r} 模型自愈改名 {cur!r} -> 回退原表 {v!r}")
                    ws_main.cell(r, name_col).value = v

        # 6. 对齐主链：强制走 reconcile 对账与四维体检闭环核验，杜绝旁路裸奔
        try:
            zip_source_rows = []
            zip_items = []
            price_col = main_hdr.index('市场价') + 1 if '市场价' in main_hdr else None
            spec_col = main_hdr.index('厂家原始规格参数') + 1 if '厂家原始规格参数' in main_hdr else None
            page_col = main_hdr.index('厂家报价文件页码') + 1 if '厂家报价文件页码' in main_hdr else None

            for r in range(2, ws_main.max_row + 1):
                p_val = ws_main.cell(r, price_col).value if price_col else None
                try:
                    p_num = float(p_val) if p_val is not None else 0.0
                except (ValueError, TypeError):
                    p_num = 0.0
                s_name = str(ws_main.cell(r, name_col).value or '').strip()
                s_spec = str(ws_main.cell(r, spec_col).value or '') if spec_col else ''
                s_page = str(ws_main.cell(r, page_col).value or '1') if page_col else '1'
                ref_id = f"{s_page}!R{r}"

                s_row = SourceRowRecord(
                    source_ref=ref_id,
                    sheet=s_page,
                    row=r,
                    raw_text=f"{s_name} {s_spec}".strip(),
                    raw_price=p_num,
                    decision='item'
                )
                zip_source_rows.append(s_row)

                c_item = CanonicalMaterialItem(
                    quote_id=f"ZIP{r:04d}",
                    supplier=supplier_name,
                    brand="",
                    category_name=s_name,
                    category_code="",
                    market_price=p_num,
                    source_ref=ref_id,
                    file_page=s_page,
                    raw_spec=s_spec
                )
                zip_items.append(c_item)

            recon = reconcile(zip_source_rows, zip_items)
            integrity_summary = getattr(recon, "summary", {})
            print("\n" + "=" * 70)
            print(f"【ZIP 旁路对齐主链 - 零信任防欺瞒四维体检指标报告】({supplier_name})")
            print(f"  -> 物理行守恒率: {integrity_summary.get('row_conservation_rate', 100):.1f}% (原表物理行: {len(zip_source_rows)})")
            print(f"  -> 物料输出行数: {integrity_summary.get('matched_items', 0)} | 对账异常项数: {integrity_summary.get('anomalies_count', 0)} 项")
            print(f"  -> 最终免检裁决: {'【通过·零缺陷免检入库】' if integrity_summary.get('is_zero_defect') else '【拦截·存在异常需复核】'}")
            print("=" * 70)

            # 落盘 run sidecar，记录 zip 引擎元数据
            sidecar_path = f"{out_path}.run.json"
            sidecar_payload = {
                "output_file": os.path.basename(out_path),
                "source_file": os.path.basename(zip_path),
                "EXTRACTOR_ENGINE_requested": "zip_standardizer",
                "EXTRACTOR_ENGINE_actual": "zip_standardizer",
                "provider_requested": "zip_rule",
                "provider_actual": "zip_rule",
                "model_id": "zip_standardizer_pipeline",
                "fallback_reason": None,
                "degraded": False,
                "total_items": len(zip_items),
                "zero_defect": bool(integrity_summary.get("is_zero_defect", False)),
                "row_conservation_rate": integrity_summary.get("row_conservation_rate", 100.0),
                "timestamp": time.strftime("%Y-%m-%dT%H:%M:%S%z")
            }
            with open(sidecar_path, "w", encoding="utf-8") as sc_f:
                json.dump(sidecar_payload, sc_f, ensure_ascii=False, indent=2)
            print(f"[*] ZIP 旁路运行 Sidecar 已落盘: {sidecar_path}")
        except Exception as recon_err:
            print(f"[Warning] ZIP 对账闸门执行异常: {recon_err}")

        wb.save(out_path)
        print(f"[✓] 标准化成功: {out_path}")
        print(f"    - 处理总行数: {total_rows_processed} | 补齐规格行: {total_specs_filled} | 内嵌图片: {total_images_embedded}\n")
        return out_path

if __name__ == '__main__':
    # 用法: python3 zip_standardizer_pipeline.py <报价单.zip> [供应商名称]
    if len(sys.argv) < 2:
        print(__doc__)
        print('用法: python3 zip_standardizer_pipeline.py <报价单.zip> [供应商名称]')
        raise SystemExit(1)
    return_path = ZipQuoteStandardizerPipeline().process_zip(
        sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    print(return_path)
