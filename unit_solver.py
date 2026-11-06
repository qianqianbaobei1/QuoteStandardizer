# -*- coding: utf-8 -*-
"""
unit_solver.py - 计量单位量纲/包装换算（只算"折合"，绝不改原件的 市场价/单位）

痛点：同一物料不同厂家用不同口径报价，入库后无法横向比价——
  BV 2.5mm²  2511 元/km  ／ 251.1 元/百米 ／ 238.5 元/卷(95米) ／ 8.6 元/根(3.6米)
33 列标准里「市场价 + 单位」是**原件口径**，改了就与原件对不上、对账必炸。
所以本模块只产出 (基准数量, 基准单位, 折合单价) 供留痕与比价，原始两个字段一律不动。

支持：
  纯量纲   km/千米/公里 → m(×1000)；百米/100m → m(×100)；t/吨 → kg(×1000)；kg → kg
  复合定尺  根/3.6米、卷/95米、桶/20kg、箱/12支  → 折合到 m / kg / 支
  不可换算  樘、套、个、台、把……  → 返回 None（不做任何猜测）
"""
import re
from typing import Optional, Tuple

# 同一量纲内的换算基数（都折到基准单位）
BASE_UNIT = {'m': 'm', 'km': 'm', '米': 'm', '千米': 'm', '公里': 'm', 'm²': 'm²', '平方米': 'm²',
             'kg': 'kg', 'g': 'kg', '克': 'kg', '千克': 'kg', '公斤': 'kg', 't': 'kg', '吨': 'kg'}

# 1 个该单位 = 多少基准单位
TO_BASE = {'m': 1.0, '米': 1.0, 'km': 1000.0, '千米': 1000.0, '公里': 1000.0,
           'm²': 1.0, '平方米': 1.0,
           'kg': 1.0, '千克': 1.0, '公斤': 1.0, 'g': 0.001, '克': 0.001, 't': 1000.0, '吨': 1000.0}

_NUM = r'(\d+(?:\.\d+)?)'
# 复合单位：「根/3.6米」「卷/95米」「桶/20kg」「箱/12支」
_COMPOUND_RE = re.compile(r'^([^/／]+)[/／]\s*' + _NUM + r'\s*([A-Za-z²]+|[\u4e00-\u9fa5]+)?$')
# 「百米」「100m」「千米」
_PREFIX_QTY_RE = re.compile(r'^(?:' + _NUM + r')?\s*(百米|千米|公里|100m|km)$', re.I)


def _norm_unit(u: str) -> str:
    s = str(u or '').strip()
    s = s.replace('㎡', 'm²').replace('m2', 'm²').replace('M2', 'm²')
    s = re.sub(r'\s+', '', s)
    return s


def solve(unit: str, price: Optional[float]) -> Optional[Tuple[float, str, float, str]]:
    """
    返回 (每计价单位的基准数量, 基准单位, 折合单价, 说明) ；无法确定时返回 None。
    例：('根/3.6米', 8.6) → (3.6, 'm', 2.389, '1根=3.6米')
    """
    u = _norm_unit(unit)
    if not u:
        return None
    try:
        p = float(price) if price not in (None, '') else None
    except (TypeError, ValueError):
        p = None

    # 1) 复合定尺/包装：X/Y单位 或 X/Y数字单位
    m = _COMPOUND_RE.match(u)
    if m:
        head = m.group(1).strip()
        qty = float(m.group(2))
        tail = (m.group(3) or '').strip()
        if qty <= 0:
            return None
        base = BASE_UNIT.get(tail)
        if base and base in ('m', 'm²', 'kg'):
            per = qty * TO_BASE.get(tail, 1.0)
            note = f'1{head}={qty}{tail}'
            return per, base, (round(p / per, 3) if p else None), note
        if not tail:                      # 「根/3.6」这种缺单位的，猜测空间太大，不做
            return None

    # 2) 纯量纲：km / 百米 / 吨 …
    pm = _PREFIX_QTY_RE.match(u)
    if pm:
        qty = float(pm.group(1)) if pm.group(1) else (100.0 if pm.group(2) == '百米' else None)
        if qty is None:
            qty = 1000.0 if pm.group(2).lower() in ('km', '千米', '公里') else 1.0
        base = BASE_UNIT.get(u) or ('m' if pm.group(2) in ('百米', 'km', '千米', '公里', '100m') else None)
        if base == 'm' and qty > 1:
            return qty, 'm', (round(p / qty, 3) if p else None), f'1{u}={qty:g}m'
        return None

    if u in TO_BASE:
        qty = TO_BASE[u]
        base = BASE_UNIT[u]
        if qty == 1.0:
            return None                   # 已经是基准单位，无需折算
        return qty, base, (round(p / qty, 3) if p else None), f'1{u}={qty:g}{base}'

    return None


def note_of(unit: str, price: Optional[float]) -> str:
    """给「报价说明」用的折合留痕文字；不需要折算时返回空串"""
    r = solve(unit, price)
    if not r:
        return ''
    per, base, base_price, why = r
    if base_price is None:
        return ''
    return f"【折算】原件 {price:g} 元/{unit}（{why}）≈ {base_price:g} 元/{base}"


if __name__ == '__main__':
    cases = [
        ('根/3.6米', 8.6, (3.6, 'm', 2.389)),
        ('km', 2511.0, (1000.0, 'm', 2.511)),
        ('百米', 251.1, (100.0, 'm', 2.511)),
        ('卷/95米', 238.5, (95.0, 'm', 2.511)),
        ('t', 3200.0, (1000.0, 'kg', 3.2)),
        ('桶/20kg', 160.0, (20.0, 'kg', 8.0)),
        ('樘', 1475.0, None),
        ('套', 300.0, None),
        ('m', 2.5, None),          # 已是基准单位，不折算
        ('根/3.6', 8.6, None),     # 缺单位，不猜
    ]
    for unit, price, want in cases:
        got = solve(unit, price)
        key = None if got is None else (got[0], got[1], got[2])
        assert key == want, (unit, price, key, want)
    print(f'[OK] unit_solver 自检 {len(cases)} 例通过')
