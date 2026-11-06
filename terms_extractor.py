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
9. "discount_ratio": 工程折扣率浮点数 (9折填 0.9，明确写“无折扣”填 1.0)；若折扣只对某个具体系列/品类生效 (如“消防喷头 8.9折”)，此字段填 null，折扣只写进 discount_series；
10. "discount_series": 折扣适用系列与折扣描述，必须保留原文，包括“无折扣，见工程再议”这类无数字的声明，以及起订量/起订金额 (如“起订量:1000m²”“起订金额:10000元”)；
11. "quote_desc": 付款条件与报价说明。

【规则】
- 只返回一个 JSON 对象；图片中没有、或看不清的字段一律填 null，严禁猜测或补默认值；
- 以“例:”“例：”“例如”“示例”开头的行（如“例：全系列 8.5 折 起订量 1000 件”）是表格印刷示例，不是本供应商的条款，必须 100% 忽略，只提取表格里实际填写的折扣行；
- 表格里写了“无折扣���见工程再议”这类不含数字的折扣行，绝不能丢弃：discount_ratio 填 1.0，原文照抄进 discount_series；
- 勾选框：☑/√/■ 表示选中，☒/□ 表示未选中；勾选标记可能被印章或手写签名遮盖，必须结合“横线填空处是否真的填了内容（如 13、含重庆市区运费、可供货全国、100公里）”确认哪一项生效；只有选中的选项才能写入；
- freight_desc、tax_desc、quote_desc、discount_series 必须照抄原文中供应商实际填写/勾选的内容，不得改写、概括或补充；
- 图片底部若附有系统自动扶正、去除黑字手签干扰并拉直放大后的公章特写，请结合原图公章与特写准确读出 supplier（含“有限公司/股份公司/集团”的完整企业名称）；
- 表单页眉/页脚里的平台或模板 Logo 不是供应商；供应商只看公章或签章文字，看不清填 null。
"""

CN_NUM = {'一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
TERMS_IMAGE_KEYS = ['报价说明', '折扣说明', '折扣记录', '折扣单', '商务条款', '说明']
TEXT_FIELDS = ['discount_series', 'freight_desc', 'tax_desc', 'freight_settle', 'quote_desc', 'quote_type']

# “例：”“例如”开头的表格印刷示例；以及不含折扣语义的分句（如“起订量1000m²”）
_EXAMPLE_PREFIX_RE = re.compile(r'^\s*(?:例\s*[:：]|例如|示例|举例)')
_DISCOUNT_TOKEN_RE = re.compile(r'折|%|％|无折扣|不打折|原价')
_GLOBAL_SERIES_TOKENS = ('全', '全部', '所有', '全系列', '全品类', '全品', '通用', '通用系列', '所有产品', '所有系列', '全系列产品', '各类产品')


def _split_discount_segments(text: str) -> list:
    """把折扣描述拆成分句并剔除“例：”示例行"""
    parts = re.split(r'[；;\n]+', str(text or ''))
    return [p.strip() for p in parts if p.strip() and not _EXAMPLE_PREFIX_RE.match(p.strip())]


def _series_name_of(segment: str) -> str:
    """取分句里的系列/品类名；“消防喷头:8.9折”→消防喷头，“所有电缆系列下浮15%”→电缆系列，无前缀返回 ''"""
    s = segment.strip()
    m = re.match(r'^(?:所有)?\s*([\u4e00-\u9fa5A-Za-z0-9]{1,10}?)\s*(?:系列|品类)?\s*[:：]', s)
    if m:
        return m.group(1)
    m = re.match(r'^(?:所有)?\s*([\u4e00-\u9fa5A-Za-z0-9]{1,10}?(?:系列|品类))', s)
    return m.group(1) if m else ''


def _is_global_series(name: str) -> bool:
    """未限定系列，或写的就是全系列 → 属于全局折扣"""
    return not name or name in _GLOBAL_SERIES_TOKENS or '全系列' in name


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
            if any(w in s for w in ('下浮', '优惠', '降', '减')):
                return round(1.0 - x / 100, 4) if 0 < x < 100 else None
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
    """把识别/人工给出的商务条款规范为下游可用的类型；缺失项保持为空

    折扣口径 (三处真实样本踩过的坑)：
    1. “例：全系列 8.5 折”属表格印刷示例，必须剔除，不能当真折扣；
    2. “无折扣，见工程再议”无数值但必须保留原文，不能因为解析不出数字就丢掉；
    3. 只对“消防喷头 8.9折”这类具体系列生效的折扣只进 discount_series_map，
       绝不写入全局 discount_ratio，避免消火栓/灭火器被误乘折扣。
    """
    given = {k: v for k, v in (terms or {}).items() if v not in (None, '')}
    out = dict(given)

    # ---- 工程折扣：分句解析，区分全局折扣与分系列折扣 ----
    raw_series_text = str(given.get('discount_series') or '')
    segments = _split_discount_segments(raw_series_text)
    global_rates, series_map = [], {}
    for seg in segments:
        if not _DISCOUNT_TOKEN_RE.search(seg):
            continue  # 纯文字补充（如“具体实际项目可再做商议”），保留原文但不参与折扣解析
        rate = parse_discount(seg)
        if rate is None:
            continue
        name = _series_name_of(seg)
        if _is_global_series(name):
            global_rates.append(rate)
        else:
            series_map[name.replace('所有', '').replace('系列', '').strip() or name] = rate

    explicit = parse_discount(given.get('discount_ratio'))
    if explicit is None and given.get('discount_ratio') is not None:
        print(f"[Terms Warning] 无法解析折扣 {given.get('discount_ratio')!r}，工程折扣留空")
    if global_rates:
        ratio = global_rates[0]
    elif series_map or (raw_series_text.strip() and not segments):
        ratio = None  # 只有分系列折扣 / 只提识别出示例行 → 不外溢成全局折扣
    elif explicit is not None:
        ratio = explicit
    else:
        ratio = parse_discount(raw_series_text)
    out['discount_ratio'] = ratio
    if series_map:
        out['discount_series_map'] = series_map
    else:
        out.pop('discount_series_map', None)

    all_rates = global_rates + list(series_map.values())
    given_min, given_max = parse_discount(given.get('discount_min')), parse_discount(given.get('discount_max'))
    lo = given_min if given_min is not None else (min(all_rates) if all_rates else ratio)
    hi = given_max if given_max is not None else (max(all_rates) if all_rates else ratio)
    if lo is not None and hi is not None and lo > hi:
        lo, hi = hi, lo
    out['discount_min'], out['discount_max'] = lo, hi

    out['tax_rate'] = parse_tax_rate(given.get('tax_rate'))
    out['is_tax_included'] = _yes_no(given.get('is_tax_included'), '是', '否')
    out['is_freight_included'] = _yes_no(given.get('is_freight_included'), '包含', '不含')
    out['province'] = normalize_province(given.get('province'))
    out['city'] = str(given.get('city') or '').strip() or ('全国' if out['province'] == '全国' else '')
    for k in TEXT_FIELDS:
        out[k] = str(given.get(k) or '')
    # 折扣原文按分句重排（剔除示例行），起订量/“见工程再议”等文字在此保留
    out['discount_series'] = '；'.join(segments)

    # “可供货全国”这类运费备注里写明的供货范围，回填到报价地区
    if not out['province']:
        freight_blob = f"{out.get('freight_desc', '')} {out.get('quote_desc', '')}"
        if any(w in freight_blob for w in ('可供货全国', '全国供货', '发全国', '全国可发')):
            out['province'] = '全国'
            out['city'] = out['city'] or '全国'
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

    def _prepare_vision_image(self):
        """准备送模型的图像：必要放大到 1600px 宽，并在底部拼接【扶正原色章 + 去黑显影章 + 极坐标拉直条】"""
        with Image.open(self.image_path) as im:
            img = im.convert('RGB')
            img.load()
        if img.width < 1600:
            scale = 1600 / img.width
            img = img.resize((int(img.width * scale), int(img.height * scale)), Image.LANCZOS)
        return self._append_seal_banner(img)

    @staticmethod
    def _append_seal_banner(img):
        """把公章特写条拼在主图下方；无公章或 OpenCV 不可用时静默原图返回"""
        try:
            import cv2
            import numpy as np
            from seal_enhancer_service import SealEnhancerService

            bgr = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
            banner = SealEnhancerService().build_seal_enhancement_banner(bgr)
            if banner is None or banner.size == 0:
                return img
            target_w = img.width
            target_h = max(120, int(banner.shape[0] * target_w / max(banner.shape[1], 1)))
            banner = cv2.resize(banner, (target_w, target_h), interpolation=cv2.INTER_CUBIC)
            banner_rgb = cv2.cvtColor(banner, cv2.COLOR_BGR2RGB)
            canvas = Image.new('RGB', (img.width, img.height + target_h + 8), 'white')
            canvas.paste(img, (0, 0))
            canvas.paste(Image.fromarray(banner_rgb), (0, img.height + 8))
            return canvas
        except Exception as e:
            print(f"[TermsExtractor] 公章显影拼图跳过: {e}")
            return img

    def _extract_by_vision(self):
        img = self._prepare_vision_image()
        buf = io.BytesIO()
        img.save(buf, format='PNG')
        img_bytes = buf.getvalue()

        # 1. 优先使用 Gemini 原生全模态视觉
        if genai and GEMINI_API_KEY:
            try:
                # 增加客户端超时保护 (15秒超时保护，满足 Google API 最小 10s 要求)
                http_opts = types.HttpOptions(timeout=15.0) if hasattr(types, "HttpOptions") else None
                client_kwargs = {"api_key": GEMINI_API_KEY}
                if http_opts: client_kwargs["http_options"] = http_opts
                client = genai.Client(**client_kwargs)
                part = types.Part.from_bytes(data=img_bytes, mime_type="image/png")
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
        b64 = base64.b64encode(img_bytes).decode("utf-8")
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
