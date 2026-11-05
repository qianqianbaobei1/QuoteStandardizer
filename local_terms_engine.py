# -*- coding: utf-8 -*-
"""
local_terms_engine.py - 基于 macOS 系统原生 Vision OCR 的离线超高精度商务条款与说明提取引擎 (精炼纯净版)
"""

import os
import re
import subprocess
from typing import Dict, Any, Optional, List

import shutil
import platform

CN_NUM = {'一': 1, '二': 2, '两': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}

def run_macos_vision_ocr(image_path: str) -> str:
    """调用 macOS 原生 Vision Framework 进行 OCR 识别（非 macOS 或未安装 swift 时自动静默降级）"""
    if platform.system() != 'Darwin' or not shutil.which('swift'):
        return ""
    if not os.path.exists(image_path):
        return ""
    swift_code = '''
    import Cocoa
    import Vision

    let imagePath = CommandLine.arguments[1]
    let imageURL = URL(fileURLWithPath: imagePath)

    guard let image = NSImage(contentsOf: imageURL),
          let cgImage = image.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
        exit(1)
    }

    let request = VNRecognizeTextRequest { request, error in
        guard let observations = request.results as? [VNRecognizedTextObservation] else { return }
        for observation in observations {
            if let topCandidate = observation.topCandidates(1).first {
                print(topCandidate.string)
            }
        }
    }
    request.recognitionLanguages = ["zh-Hans", "en-US"]
    request.recognitionLevel = .accurate
    let handler = VNImageRequestHandler(cgImage: cgImage, options: [:])
    try? handler.perform([request])
    '''
    try:
        res = subprocess.run(
            ['swift', '-e', swift_code, os.path.abspath(image_path)],
            capture_output=True,
            text=True,
            timeout=25
        )
        if res.returncode == 0:
            return res.stdout
    except Exception as e:
        print(f"[LocalOCR Error] {e}")
    return ""

def parse_discount_num(val_str: str) -> Optional[float]:
    if not val_str:
        return None
    s = str(val_str).strip().replace(' ', '')
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
        if x >= 10:
            x = x / 10
        return round(x / 10, 4) if 0 < x <= 10 else None
    if 0 < x <= 1:
        return round(x, 4)
    if 1 < x <= 10:
        return round(x / 10, 4)
    if 10 < x <= 100:
        return round(x / 100, 4)
    return None

def extract_terms_from_ocr_text(ocr_text: str) -> Dict[str, Any]:
    """从 OCR 提取的全文中，精炼解析商务字段，彻底过滤模版示例文本"""
    raw_lines = [line.strip() for line in ocr_text.split('\n') if line.strip()]
    
    # 彻底过滤模版示例干扰行与模版提示词
    filtered_lines = []
    for l in raw_lines:
        if l.startswith('例:') or l.startswith('例：') or l.startswith('例 '):
            continue
        if any(k in l for k in [
            '面向建筑单位批量采购时给予',
            '请在折扣说明中记录下来',
            '请务必包含您报价单中的全部商品',
            '（请根据该报价实际情况二选一）',
            '（请根据该报价的实际情况二选一）',
            '（如：含多少公里',
            '（如100元/吨/公里）',
            '（如 100元/吨/公里）',
            '如100元/吨/公里',
            '产品类别/系列',
            '其他说明（非必填）'
        ]):
            continue
        filtered_lines.append(l)

    full_text = " ".join(filtered_lines)
    
    terms = {
        'is_tax_included': '',
        'tax_rate': None,
        'tax_desc': '',
        'is_freight_included': '',
        'freight_desc': '',
        'quote_desc': '',
        'discount_ratio': None,
        'discount_series': '',
        'freight_settle': '',
        'quote_type': ''
    }
    
    # 1. 税率与是否含税（未提取到坚决留空，绝不塞入 13% 假数据）
    m_tax = re.search(r'(?:增值税税率|税率|加税)[：:\s]*(\d{1,2})%', full_text)
    if m_tax:
        terms['tax_rate'] = int(m_tax.group(1))
        terms['tax_desc'] = f"增值税税率{m_tax.group(1)}%"

    if any(re.search(r'[☑√■\✓]\s*不含税', l) for l in raw_lines):
        terms['is_tax_included'] = '否'
    elif any(re.search(r'[☑√■\✓]\s*含税', l) for l in raw_lines):
        terms['is_tax_included'] = '是'
    elif '不含税' in full_text and '含税' not in full_text:
        terms['is_tax_included'] = '否'
    elif any(w in full_text for w in ('含税价', '报价含税', '均含税', '价格已含税')):
        terms['is_tax_included'] = '是'

    # 2. 运费
    # 检测勾选
    is_freight_no = any(re.search(r'[☑√■\✓]\s*不含运费', l) for l in raw_lines) or ('不含运费' in full_text and '运费另算' in full_text)
    is_freight_yes = any(re.search(r'[☑√■\✓]\s*含运费', l) for l in raw_lines) or ('包运' in full_text)
    
    if is_freight_no:
        terms['is_freight_included'] = '不含'
    elif is_freight_yes:
        terms['is_freight_included'] = '包含'

    # 通用运费描述提取：匹配包含运费、运输、送货、包邮、自提等通用物流语义的说明行（严禁绑定特定厂家原话）
    freight_desc_parts = []
    FREIGHT_KEYWORDS = ('运费', '运输', '配送', '送货', '包运', '包邮', '自提', '货到', '车板交货')
    for l in filtered_lines:
        if any(k in l for k in FREIGHT_KEYWORDS):
            clean_l = re.sub(r'^[☑√■\✓口\s]+', '', l)
            clean_l = re.sub(r'\(.*?\)|（.*?）', '', clean_l).strip('，, ；;')
            if clean_l and clean_l not in freight_desc_parts:
                freight_desc_parts.append(clean_l)
    if freight_desc_parts:
        terms['freight_desc'] = "；".join(freight_desc_parts)

    # 3. 付款方式与报价说明（通用提取，支持账期、付款方式、定尺、结算等商务条款）
    quote_parts = []
    m_term = re.search(r'接受账期\s*(\d+)\s*个月[，,]?\s*价格上浮\s*([\d\.]+)%', full_text)
    if m_term:
        quote_parts.append(f"接受账期{m_term.group(1)}个月，价格上浮{m_term.group(2)}%")

    COMMERCIAL_KEYWORDS = ('账期', '付款', '现款', '结算', '现结', '款到', '加价', '质保', '订货', '有效期')
    for l in filtered_lines:
        if any(k in l for k in COMMERCIAL_KEYWORDS):
            clean_l = re.sub(r'^[☑√■\✓口\s]+', '', l).strip('，, ；;')
            if clean_l and not any(clean_l in p for p in quote_parts) and clean_l not in terms['freight_desc']:
                quote_parts.append(clean_l)

    if quote_parts:
        terms['quote_desc'] = "；".join(quote_parts)

    # 4. 工程折扣提取（严格过滤例句）
    # 查找表格中的产品类别与折扣行
    real_disc_items = []
    in_disc_table = False
    for l in raw_lines:
        if l.startswith('例:') or l.startswith('例：') or l.startswith('例 '):
            continue
        if '产品类别' in l or '工程折扣' in l:
            in_disc_table = True
            continue
        if in_disc_table:
            if '其他说明' in l or '公司盖章' in l or '签字' in l:
                in_disc_table = False
                break
            # 真实表格行
            if any(k in l for k in ['折', '%']) or l in ('1', '1.0'):
                real_disc_items.append(l)

    # 解析 real_disc_items
    valid_ratios = []
    valid_series = []
    for it in real_disc_items:
        r = parse_discount_num(it)
        if r is not None:
            valid_ratios.append(r)
            valid_series.append(it)

    if valid_ratios:
        terms['discount_ratio'] = valid_ratios[0]
        terms['discount_series'] = "；".join(valid_series)
    else:
        # 如果表格没写但有文字折扣（且不是示例）
        non_example_text = re.sub(r'例[：:].*?(\n|$)', '', ocr_text)
        m_disc = re.findall(r'(\b\d+(?:\.\d+)?折|\b\d+%)', non_example_text)
        discs = [parse_discount_num(d) for d in m_disc if parse_discount_num(d) is not None]
        if discs:
            terms['discount_ratio'] = discs[0]
            terms['discount_series'] = m_disc[0]

    return terms

def extract_terms_from_image_local(image_path: str) -> Dict[str, Any]:
    ocr_text = run_macos_vision_ocr(image_path)
    if not ocr_text:
        return {}
    return extract_terms_from_ocr_text(ocr_text)
