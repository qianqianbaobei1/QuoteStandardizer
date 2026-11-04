# -*- coding: utf-8 -*-
"""
商务条款 (报价说明 / 折扣记录单) 解析
1. 图片交给视觉模型 (OpenAI 兼容接口) 直接识别
2. 识别结果统一经 normalize_terms 规范 ("4折"→0.4、"13%"→13)
3. 没有提供、或识别不出的条款一律留空 (None / "")，绝不填写默认值
"""
import base64
import io
import json
import os
import re

import requests
from PIL import Image

from config import VISION_API_KEY, VISION_BASE_URL, VISION_MODEL, GEMINI_API_KEY, GEMINI_MODEL, GEMINI_THINKING_BUDGET, normalize_province

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

SYSTEM_PROMPT_TERMS_VISION = """你是工程采购商务条款的转录员。请从提供的《报价说明》《工程折扣记录单》等商务附件图片中提取条款。

【提取字段】
1. "supplier": 公章或签章上的公司全称；
2. "brand": 供应品牌名称；
3. "province": 报价/供货适用区域，照抄原文 (如“全国”“京津冀+山东”)；
4. "is_tax_included": 是否含税 ("是" 或 "否")；
5. "tax_rate": 增值税税率数字；
6. "tax_desc": 税费说明；
7. "is_freight_included": 是否含运费 ("包含" 或 "不含")；
8. "freight_desc": 运费约定；
9. "discount_ratio": 工程折扣率浮点数 (9折填 0.9，明确写“无折扣”填 1.0)；
10. "discount_series": 折扣适用系列与折扣描述；
11. "quote_desc": 付款条件与报价说明。

【规则】
- 只返回一个 JSON 对象；图片中没有、或看不清的字段一律填 null，严禁猜测或补默认值；
- 以“例:”开头的行是表格填写示例，不是本供应商的条款，必须忽略；
- 勾选框：☑/√/■ 表示选中，□ 表示未选中；只有选中的选项才能写入；
- freight_desc、tax_desc、quote_desc、discount_series 必须照抄原文中供应商实际填写/勾选的内容，不得改写、概括或补充；
- 表单页眉/页脚里的平台或模板 Logo 不是供应商；供应商只看公章或签章文字，看不清填 null。
"""

CN_NUM = {'一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
TERMS_IMAGE_KEYS = ['报价说明', '折扣说明', '折扣记录', '折扣单', '商务条款', '说明']
TEXT_FIELDS = ['discount_series', 'freight_desc', 'tax_desc', 'freight_settle', 'quote_desc', 'quote_type']


def parse_discount(val):
    """'4折'→0.4、'85折'→0.85、'8.5折'→0.85、'八五折'→0.85、'40%'→0.4、0.4→0.4、'无折扣'→1.0；无法解析返回 None"""
    if val is None or val == '':
        return None
    if isinstance(val, (int, float)):
        x = float(val)
    else:
        s = str(val).strip().replace(' ', '')
        if any(k in s for k in ('无折扣', '不打折', '原价')):
            return 1.0
        cn = re.search(r'([一二两三四五六七八九])([一二三四五六七八九])?折', s)
        if cn and not re.search(r'\d', s):
            x = CN_NUM[cn.group(1)] + (CN_NUM[cn.group(2)] / 10 if cn.group(2) else 0)
            return round(x / 10, 4)
        m = re.search(r'(\d+(?:\.\d+)?)', s)
        if not m:
            return None
        x = float(m.group(1))
        if '%' in s:
            return round(x / 100, 4) if 0 < x <= 100 else None
        if '折' in s:
            if x >= 10:  # 85折 → 8.5折
                x = x / 10
            return round(x / 10, 4) if 0 < x <= 10 else None
    if 0 < x <= 1:
        return round(x, 4)
    if 1 < x <= 10:
        return round(x / 10, 4)
    if 10 < x <= 100:
        return round(x / 100, 4)
    return None


def parse_tax_rate(val):
    """'13%'→13、0.13→13、'9'→9；无法解析返回 None"""
    if val is None or val == '':
        return None
    m = re.search(r'(\d+(?:\.\d+)?)', str(val))
    if not m:
        return None
    x = float(m.group(1))
    if 0 < x < 1:
        x *= 100
    return int(round(x)) if 0 <= x <= 100 else None


def _yes_no(val, yes, no):
    """规范是/否类字段；无法判断返回 ''"""
    s = str(val or '').strip()
    if not s:
        return ''
    if s in (yes, no):
        return s
    if any(k in s for k in ['不含', '否', '未含', 'false', 'False', '另计']):
        return no
    if any(k in s for k in ['含', '是', '包', 'true', 'True']):
        return yes
    return ''


def normalize_terms(terms: dict) -> dict:
    """把识别/人工给出的商务条款规范为下游可用的类型；缺失项保持为空"""
    given = {k: v for k, v in (terms or {}).items() if v not in (None, '')}
    out = dict(given)

    ratio = parse_discount(given.get('discount_ratio'))
    if ratio is None:
        ratio = parse_discount(given.get('discount_series'))
    if ratio is None and given.get('discount_ratio') is not None:
        print(f"[Terms Warning] 无法解析折扣 {given.get('discount_ratio')!r}，工程折扣留空")
    out['discount_ratio'] = ratio
    out['discount_min'] = parse_discount(given.get('discount_min')) or ratio
    out['discount_max'] = parse_discount(given.get('discount_max')) or ratio
    if out['discount_min'] is not None and out['discount_max'] is not None and out['discount_min'] > out['discount_max']:
        out['discount_min'], out['discount_max'] = out['discount_max'], out['discount_min']

    out['tax_rate'] = parse_tax_rate(given.get('tax_rate'))
    out['is_tax_included'] = _yes_no(given.get('is_tax_included'), '是', '否')
    out['is_freight_included'] = _yes_no(given.get('is_freight_included'), '包含', '不含')
    out['province'] = normalize_province(given.get('province'))
    out['city'] = str(given.get('city') or '').strip() or ('全国' if out['province'] == '全国' else '')
    for k in TEXT_FIELDS:
        out[k] = str(given.get(k) or '')
    return out


def missing_terms(terms: dict) -> list:
    """列出规范化后仍为空的关键条款，供待复核说明使用"""
    labels = {'discount_ratio': '工程折扣', 'tax_rate': '增值税税率', 'is_tax_included': '是否含税',
              'is_freight_included': '是否含运费', 'province': '报价地区'}
    return [name for k, name in labels.items() if terms.get(k) in (None, '')]


def find_terms_image(folder: str, exclude_file: str = None):
    """在报价单同目录查找商务说明/折扣单图片 (文件名含关键词即可，如 “报价说明20260918.png”)，自动排除输入文件本身"""
    try:
        files = [f for f in os.listdir(folder) if os.path.splitext(f)[1].lower() in ('.png', '.jpg', '.jpeg')]
    except OSError:
        return None
    exclude_base = os.path.basename(exclude_file) if exclude_file else None
    for key in TERMS_IMAGE_KEYS:
        for f in sorted(files):
            if exclude_base and f == exclude_base:
                continue
            if key in f:
                return os.path.join(folder, f)
    return None


class TermsExtractor:
    def __init__(self, image_path):
        self.image_path = image_path

    def extract_terms(self) -> dict:
        """识别商务单图片，返回原始条款字典 (由调用方统一 normalize_terms)；失败返回 {'_source': 'none'}"""
        if not os.path.exists(self.image_path):
            return {'_source': 'none'}
        if not (GEMINI_API_KEY or (VISION_API_KEY and VISION_BASE_URL and VISION_MODEL)):
            print("[TermsExtractor Warning] 未配置视觉模型，商务条款无法识别。")
            return {'_source': 'none'}
        parsed = self._extract_by_vision()
        if parsed:
            parsed['_source'] = 'vision'
            print(f"[*] 商务单识别结果: 供应商={parsed.get('supplier')} | 折扣={parsed.get('discount_series')}")
            return parsed
        print("[TermsExtractor Warning] 商务单识别失败，商务条款留空。")
        return {'_source': 'none'}

    def _chat(self, messages, timeout=90):
        payload = {"model": VISION_MODEL, "messages": messages, "temperature": 0,
                   "response_format": {"type": "json_object"}}
        headers = {"Authorization": f"Bearer {VISION_API_KEY}", "Content-Type": "application/json"}
        try:
            url = f"{VISION_BASE_URL.rstrip('/')}/chat/completions"
            resp = requests.post(url, headers=headers, json=payload, timeout=timeout)
            if resp.status_code == 400:
                # 部分多模态模型不支持 response_format，去掉后重试
                payload.pop("response_format")
                resp = requests.post(url, headers=headers, json=payload, timeout=timeout)
            if resp.status_code != 200:
                print(f"[TermsExtractor Warning] HTTP {resp.status_code}: {resp.text[:200]}")
                return None
            parsed = json.loads(self._clean_json_str(resp.json()["choices"][0]["message"]["content"]))
            return parsed if isinstance(parsed, dict) else None
        except Exception as e:
            print(f"[TermsExtractor Warning] 解析异常: {e}")
            return None

    def _extract_by_vision(self):
        # 1. 优先使用 Gemini 原生全模态视觉
        if genai and GEMINI_API_KEY:
            try:
                # 增加客户端超时保护 (15秒超时保护，满足 Google API 最小 10s 要求)
                http_opts = types.HttpOptions(timeout=15.0) if hasattr(types, "HttpOptions") else None
                client_kwargs = {"api_key": GEMINI_API_KEY}
                if http_opts: client_kwargs["http_options"] = http_opts
                client = genai.Client(**client_kwargs)
                with open(self.image_path, "rb") as f:
                    img_bytes = f.read()
                ext = os.path.splitext(self.image_path)[1].lower()
                mime = "image/png" if ext == ".png" else "image/jpeg"
                part = types.Part.from_bytes(data=img_bytes, mime_type=mime)
                resp = client.models.generate_content(
                    model=GEMINI_MODEL,
                    contents=[part, SYSTEM_PROMPT_TERMS_VISION + "\n请严格以 JSON 格式输出识别到的条款："],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        temperature=0.0,
                        thinking_config=types.ThinkingConfig(thinking_budget=GEMINI_THINKING_BUDGET)
                    )
                )
                if resp.text:
                    parsed = json.loads(self._clean_json_str(resp.text))
                    if isinstance(parsed, dict):
                        return parsed
            except Exception as e:
                print(f"[TermsExtractor] Gemini 原生全模态识别异常: {e}，回退至兼容通道...")

        # 2. 传统兼容通道
        # 商务单图片分辨率常很低：放大到约 1600px 宽再送模型，提升小字与勾选框识别
        img = Image.open(self.image_path).convert('RGB')
        if img.width < 1600:
            scale = 1600 / img.width
            img = img.resize((int(img.width * scale), int(img.height * scale)), Image.LANCZOS)
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        b64 = base64.b64encode(buf.getvalue()).decode("utf-8")
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT_TERMS_VISION},
            {"role": "user", "content": [
                {"type": "text", "text": "请提取这张商务单据中的全部商务条款，以 JSON 返回："},
                {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}}
            ]}
        ]
        return self._chat(messages)

    @staticmethod
    def _clean_json_str(text: str) -> str:
        text = text.strip()
        if text.startswith("```json"): text = text[7:]
        elif text.startswith("```"): text = text[3:]
        if text.endswith("```"): text = text[:-3]
        return text.strip()
