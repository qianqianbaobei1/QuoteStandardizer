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
import io
import re
import json
import base64
import zipfile
import openpyxl
from openpyxl.drawing.image import Image as OpenpyxlImage
from openpyxl.utils import get_column_letter
import requests

from config import (
    DEEPSEEK_API_KEY, DEEPSEEK_BASE_URL, VISION_MODEL,
    DEFAULT_OUTPUT_DIR
)

class ZipQuoteStandardizerPipeline:
    def __init__(self, output_dir=DEFAULT_OUTPUT_DIR):
        self.output_dir = output_dir
        os.makedirs(self.output_dir, exist_ok=True)
        self.transcription_cache = {}

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
        """结构化解析规格参数字符串为字典映射"""
        if not spec_text:
            return {}
        res = {}
        items = re.split(r'[;\n；]', str(spec_text))
        for it in items:
            it = it.strip()
            if not it:
                continue
            if ':' in it:
                k, v = it.split(':', 1)
                res[k.strip()] = v.strip()
            elif '：' in it:
                k, v = it.split('：', 1)
                res[k.strip()] = v.strip()
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
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            
            # 定位关键列
            spec_col = headers.index('厂家原始规格参数') + 1 if '厂家原始规格参数' in headers else None
            name_col = headers.index('名称规格') + 1 if '名称规格' in headers else None
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
                row_img_file = str(ws.cell(r, img_file_col).value or "quote_img1.png").strip() if img_file_col else "quote_img1.png"

                # 规则 1: 若当前规格参数为空，系统级从视觉转录结果中精准匹配回填
                if (not cur_spec or str(cur_spec).strip() == '') and spec_col:
                    matched_spec = None
                    c_name = self.clean_name(row_name)
                    # 优先按同图片+序号匹配
                    for it in all_transcriptions:
                        if it.get('_img_source') == row_img_file and str(it.get('seq')) == str(row_seq):
                            matched_spec = it.get('spec')
                            break
                    # 次优先按品名完全匹配
                    if not matched_spec:
                        for it in all_transcriptions:
                            t_name = self.clean_name(it.get('name', ''))
                            if t_name and (t_name == c_name or t_name in c_name or c_name in t_name):
                                matched_spec = it.get('spec')
                                break
                    # 再次优先跨图名称部分匹配
                    if not matched_spec:
                        for it in all_transcriptions:
                            t_name = self.clean_name(it.get('name', ''))
                            common = set(c_name) & set(t_name)
                            if len(common) >= 3 or (len(c_name) <= 3 and len(common) >= 2):
                                matched_spec = it.get('spec')
                                break
                    # 最后兜底：基于同图+价格+单位精准兜底对齐（解决 OCR 品名误差）
                    if not matched_spec:
                        price_col = headers.index('市场价') + 1 if '市场价' in headers else None
                        row_p = ws.cell(r, price_col).value if price_col else None
                        if row_p:
                            try:
                                r_pval = float(str(row_p).replace('¥', '').replace(',', ''))
                                for it in all_transcriptions:
                                    if it.get('_img_source') == row_img_file and it.get('price'):
                                        if abs(float(it['price']) - r_pval) < 0.01:
                                            matched_spec = it.get('spec')
                                            break
                            except Exception:
                                pass

                    if matched_spec:
                        ws.cell(r, spec_col).value = matched_spec
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
                                if (k_norm == h_norm or k_norm in h_norm or h_norm in k_norm) and not ws.cell(r, c_idx).value:
                                    ws.cell(r, c_idx).value = v
                                    break

                # 规则 3: 严格区分“产品图”与“客户报价单图”
                # 【重要规范】：
                # - “厂家报价图片文件名”列：记录客户报价单/扫描件图文件名（如 quote_img1.png），用于溯源对账
                # - “图片”列：只允许放置物料对应的“真实产品效果图/样品图/实物图”；
                #   严禁把客户报价单/带公章扫描件（quote_img*）作为产品图塞入！若原包未提供真实产品图库，则“图片”列保持纯净留空。
                if img_col:
                    # 检查是否有真实产品图
                    product_img_found = None
                    for cc in range(img_col, ws.max_column + 1):
                        val = str(ws.cell(r, cc).value or '').strip()
                        # 如果单元格内记录了具体的产品图片名称（非 quote_img 且非空）
                        if val and not val.lower().startswith('quote_img') and val in local_img_paths:
                            product_img_found = local_img_paths[val]
                        ws.cell(r, cc).value = None  # 清空文字，保持单元格纯净

                    if product_img_found:
                        ws.row_dimensions[r].height = 95
                        try:
                            img = OpenpyxlImage(product_img_found)
                            img.width = 135
                            img.height = 90
                            ws.add_image(img, f"{img_letter}{r}")
                            total_images_embedded += 1
                        except Exception as e:
                            print(f"[Error] Failed to embed product image {product_img_found}: {e}")

        # 4. 保存为交付标准文件
        out_path = os.path.join(self.output_dir, f"{supplier_name}-标准解析.xlsx")
        for ws in wb.worksheets:
            apply_enterprise_style(ws)
        wb.save(out_path)
        print(f"[✓] 标准化成功: {out_path}")
        print(f"    - 处理总行数: {total_rows_processed} | 补齐规格行: {total_specs_filled} | 内嵌图片: {total_images_embedded}\n")
        return out_path

if __name__ == '__main__':
    pipeline = ZipQuoteStandardizerPipeline()
    test_zips = [
        ('/Users/qianqianawodebaobei/.gemini/antigravity/brain/ff156e75-4bef-45d5-828b-c11d0e4d6b2e/.user_uploaded/media_1793673628042_60ff8e76.zip', '青岛沃丰钢铁有限公司'),
        ('/Users/qianqianawodebaobei/.gemini/antigravity/brain/ff156e75-4bef-45d5-828b-c11d0e4d6b2e/.user_uploaded/media_1793673628730_4ec3dd31.zip', '天津柴进电缆桥架制造有限公司'),
        ('/Users/qianqianawodebaobei/.gemini/antigravity/brain/ff156e75-4bef-45d5-828b-c11d0e4d6b2e/.user_uploaded/media_1793673628039_d9369ba2.zip', '深圳市凝光时代科技有限公司')
    ]
    for zp, sname in test_zips:
        pipeline.process_zip(zp, sname)
