# -*- coding: utf-8 -*-
"""
name_norm.py - 物料名称归一化（唯一入口）

词表建库、在线查表、交付件对账必须共用这一套归一化，否则两边 key 对不上：
  1. 去 [572305] / 【123】 这类原表自带的类别编码前缀
  2. 去型号前缀（NLED4916 吊装线型灯 → 吊装线型灯）
  3. 去括号里的规格补充（热塑性聚烯烃(TPO)防水卷材 → 热塑性聚烯烃防水卷材）
  4. 全角转半角、去空白、统一小写
注意：字母开头但没有数字的（CPL覆膜钢质门、UPVC排水管、HDPE双壁波纹管）不动，它们是品名的一部分。
"""
import re
import unicodedata

_BRACKET_CODE_RE = re.compile(r'^[\[【][A-Za-z0-9\-_]+[\]】]\s*')
# 必须以大写字母开头再带数字（NLED4916 / DN100）；纯数字开头的是规格（2mm橡胶卷材），不是型号
_MODEL_PREFIX_RE = re.compile(r'^(?:[A-Z]{1,6}[\s\-])?[A-Z]{1,6}\d[A-Z0-9\-]*[\s　]*')
_PAREN_RE = re.compile(r'[（(][^）)]*[）)]')
_SPACE_RE = re.compile(r'\s+')


def strip_model_prefix(name: str) -> str:
    """剥掉品名前的 [572305] 类别编码与型号编码；剥完必须还剩汉字，否则原样返回"""
    orig = _SPACE_RE.sub(' ', str(name or '')).strip()
    s = orig
    s_bracket = _BRACKET_CODE_RE.sub('', s).strip()
    if s_bracket and re.search(r'[\u4e00-\u9fa5]', s_bracket):
        s = s_bracket
    for _ in range(2):
        s2 = _MODEL_PREFIX_RE.sub('', s, count=1).strip()
        if s2 == s:
            break
        s = s2
    return s if re.search(r'[\u4e00-\u9fa5]', s) else orig


def normalize_material_name(name: str) -> str:
    """物料名称 → 查表 key"""
    s = unicodedata.normalize('NFKC', str(name or '')).strip()
    s = strip_model_prefix(s)
    s = _PAREN_RE.sub('', s)
    return _SPACE_RE.sub('', s).lower()


if __name__ == '__main__':
    cases = {
        'NLED4916 吊装线型灯': '吊装线型灯',
        'LED T806P灯管': '灯管',
        'LEDA45J球泡': '球泡',
        '[572305]光纤收发器': '光纤收发器',
        '热塑性聚烯烃(TPO)防水卷材': '热塑性聚烯烃防水卷材',
        'CPL覆膜钢质门': 'cpl覆膜钢质门',
        'UPVC排水管': 'upvc排水管',
        '2mm橡胶卷材': '2mm橡胶卷材',          # 数字开头是规格，不许当型号剥掉
        '[FT-01]球阀': '球阀',                  # 方括号编码前缀要剥
        '110KV/220KV变电站水喷雾系统': '110kv/220kv变电站水喷雾系统',
        'YJV电力电缆': 'yjv电力电缆',          # YJV 是型号但不是"字母+数字"，保留
    }
    bad = {k: (normalize_material_name(k), v) for k, v in cases.items() if normalize_material_name(k) != v}
    assert not bad, bad
    print(f'[OK] name_norm 自检 {len(cases)} 例通过')
