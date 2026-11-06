# -*- coding: utf-8 -*-
"""真实《报价说明-工程折扣记录单》样本踩过的折扣口径回归测试"""
from local_terms_engine import parse_discount_num
from terms_extractor import normalize_terms


def test_specific_series_discount_does_not_pollute_global():
    """天兴消防：表里只写“消防喷头 8.9折”，不能把 0.89 套到消火栓/灭火器上"""
    t = normalize_terms({'discount_series': '消防喷头:8.9折(起订量:10000只)', 'discount_ratio': 0.89})
    assert t['discount_ratio'] is None
    assert t['discount_series_map'] == {'消防喷头': 0.89}
    assert '10000只' in t['discount_series']


def test_example_rows_are_ignored():
    """表单印刷示例“例：全系列 8.5 折”必须剔除，真折扣照常保留"""
    t = normalize_terms({'discount_series': '例：全系列 8.5 折 起订量 1000 件；消防喷头:8.9折', 'discount_ratio': 0.85})
    assert '8.5' not in t['discount_series']
    assert t['discount_series_map'] == {'消防喷头': 0.89}
    assert t['discount_ratio'] is None


def test_no_discount_text_is_kept():
    """宁波/集实筑镀：“无折扣，见工程再议”不能因为解析不出数字就丢掉，无折扣时折扣率保持 None，严禁臆测 1.0"""
    t = normalize_terms({'discount_series': '全系列:无折扣见工程再议(起订量:1000m²)'})
    assert t['discount_ratio'] is None
    assert '见工程再议' in t['discount_series']
    assert '1000m²' in t['discount_series']


def test_global_series_discount_stays_global():
    """重庆齐汇总：全系列 95 折 → 全局 0.95，原文“具体实际项目可再做商议”保留"""
    t = normalize_terms({'discount_series': '全系列95折(具体实际项目可再做商议)'})
    assert t['discount_ratio'] == 0.95
    assert '再做商议' in t['discount_series']


def test_multi_series_map_and_range():
    t = normalize_terms({'discount_series': '所有电缆系列下浮15%；所有电线系列下浮12%'})
    assert t['discount_ratio'] is None
    assert t['discount_series_map'] == {'电缆': 0.85, '电线': 0.88}
    assert (t['discount_min'], t['discount_max']) == (0.85, 0.88)


def test_province_from_freight_note():
    t = normalize_terms({'freight_desc': '不含运费：可供货全国'})
    assert t['province'] == '全国'


def test_local_engine_reads_zhe_discount():
    assert parse_discount_num('95折') == 0.95
    assert parse_discount_num('8.9折') == 0.89


if __name__ == '__main__':
    for _name, _fn in list(globals().items()):
        if _name.startswith('test_') and callable(_fn):
            _fn()
    print('ok')
