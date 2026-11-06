# -*- coding: utf-8 -*-
"""
price_guard.py - 跨行价格离群护栏（只提示，绝不改价）

能抓什么：同一张表里「同类别编码 + 同单位」的一批行，出现 10x/100x 的小数点漏读、
          km→m、吨→kg 这类量纲错配（单行表无法判断，直接跳过，不许硬报）。
抓不了什么：整表只有 1~2 行、或全表同错（同一批同一错法）——那要靠外部基准，本模块不猜。

口径：
  · 分组 = (类别编码, 单位)，且组内 ≥ MIN_GROUP 行才判；
  · 用组内 P10~P90 做自适应区间（电缆 1.5→240mm² 这种合法 100 倍跨度也能容忍）；
  · 超出 [P10/5, P90*5] 才算离群，报「严重」并给出组内中位数，不修改任何数值。
"""
import statistics
from typing import Dict, List, Optional, Tuple

MIN_GROUP = 5          # 少于 5 行不判（样本太少的"离群"没有统计意义）
OUTLIER_FACTOR = 5.0   # 超出 P10/5 或 P90*5
NEIGHBOR_JUMP = 10.0   # 相邻档位金额跳变倍数（配合“同时超过中位数 10 倍”）


def _quantile(sorted_vals: List[float], q: float) -> float:
    if not sorted_vals:
        return 0.0
    if len(sorted_vals) == 1:
        return sorted_vals[0]
    pos = (len(sorted_vals) - 1) * q
    lo, hi = int(pos), min(int(pos) + 1, len(sorted_vals) - 1)
    return sorted_vals[lo] + (sorted_vals[hi] - sorted_vals[lo]) * (pos - lo)


def find_outliers(rows: List[Dict]) -> List[Dict]:
    """
    rows: [{'key','price','unit','code'}]
    返回: [{'key','price','median','low','high','group_size','ratio'}]
    """
    groups: Dict[Tuple[str, str], List[Tuple[str, float]]] = {}
    for r in rows:
        try:
            price = float(r.get('price'))
        except (TypeError, ValueError):
            continue
        if price <= 0:
            continue
        key = (str(r.get('code') or ''), str(r.get('unit') or '').strip())
        groups.setdefault(key, []).append((r.get('key'), price))

    out = []
    for (_code, _unit), members in groups.items():
        if len(members) < MIN_GROUP:
            continue
        prices = sorted(p for _k, p in members)
        p10, p90 = _quantile(prices, 0.10), _quantile(prices, 0.90)
        med = statistics.median(prices)
        low, high = p10 / OUTLIER_FACTOR, p90 * OUTLIER_FACTOR
        for k, p in members:
            if p < low or p > high:
                ratio = (p / med) if med else 0.0
                out.append({'key': k, 'price': p, 'median': med, 'low': round(low, 4),
                            'high': round(high, 4), 'group_size': len(members),
                            'ratio': round(ratio, 3)})
    # 追加规则：相邻档位跳变（单调递增的电缆/管材价格序列里，单个点突然大 ≥10 倍）
    # 单靠 P90 会被这个离群点自己污染；用"与最近的低价档位比值"更稳
    flagged = {h['key'] for h in out}
    for (_code, _unit), members in groups.items():
        if len(members) < MIN_GROUP:
            continue
        ordered = sorted(members, key=lambda kv: kv[1])
        med = statistics.median([p for _k, p in ordered])
        for idx in range(1, len(ordered)):
            k, p = ordered[idx]
            prev_p = ordered[idx - 1][1]
            if prev_p <= 0 or p / prev_p < NEIGHBOR_JUMP:
                continue
            if med and p < med * 10:
                continue
            if k in flagged:
                continue
            out.append({'key': k, 'price': p, 'median': med, 'low': round(prev_p, 4),
                        'high': round(prev_p * OUTLIER_FACTOR, 4), 'group_size': len(members),
                        'ratio': round(p / prev_p, 3)})
            flagged.add(k)
    return out


def note_of(hit: Dict) -> str:
    return (f"价格离群：{hit['price']:g} 与本组中位数 {hit['median']:g} 相差 "
            f"{hit['ratio']:.2f} 倍（同类别同单位 {hit['group_size']} 行，合理区间 "
            f"{hit['low']:g}~{hit['high']:g}），疑似小数点/单位量纲错读，需核对原件")


if __name__ == '__main__':
    # 1) 正常宽跨度（电缆 1.5→240mm²）不应误报
    cable = [{'key': i, 'price': p, 'unit': 'm', 'code': '25-11-01'}
             for i, p in enumerate([2.5, 3.1, 4.2, 6.8, 11.0, 18.5, 35.0, 62.0, 120.0, 240.0])]
    assert find_outliers(cable) == [], '合法宽跨度被误报'
    # 2) 混进一条 km 误读（1000 倍）必须抓到
    bad = cable + [{'key': 'x', 'price': 2511.0, 'unit': 'm', 'code': '25-11-01'}]
    hits = find_outliers(bad)
    assert len(hits) == 1 and hits[0]['key'] == 'x', hits
    # 3) 组内不足 5 行不判
    small = [{'key': i, 'price': p, 'unit': 'm', 'code': 'x'} for i, p in enumerate([1, 2, 3, 999])]
    assert find_outliers(small) == [], '样本不足还硬报'
    # 4) 单位不同不混组
    mixed = cable + [{'key': 'y', 'price': 2511.0, 'unit': 'km', 'code': '25-11-01'}]
    assert find_outliers(mixed) == [], '不同单位被混组'
    # 5) 单调递增序列里混进 100 倍错价（P90 会被自己污染，靠相邻档位规则抓）
    mono = [{'key': i, 'price': p, 'unit': 'km', 'code': '25-11-01'}
            for i, p in enumerate([2511, 3200, 4800, 8000, 12000, 18500, 251100])]
    hits = find_outliers(mono)
    assert any(h['key'] == 6 for h in hits), hits
    print('[OK] price_guard 自检 5 例通过')
