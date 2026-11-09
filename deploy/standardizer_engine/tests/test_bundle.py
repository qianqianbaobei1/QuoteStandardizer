# -*- coding: utf-8 -*-
"""
test_bundle.py - 部署包集成验证测试
"""
import os
import sys

# 注入当前目录
CUR_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if CUR_DIR not in sys.path:
    sys.path.insert(0, CUR_DIR)

from standardizer_engine import StandardizerEngine

def run_tests():
    print("=== 开始运行部署包端到端集成测试 ===")
    engine = StandardizerEngine(use_jev=True)

    test_cases = [
        {
            "name": "热轧圆钢",
            "spec": "Φ20 Q235B",
            "raw_cat": "金属及金属制品",
            "expected_code_prefix": "01",
            "expected_attrs": ["直径(mm)", "牌号"]
        },
        {
            "name": "天然大理石板材",
            "spec": "800*800*20 抛光面 白麻",
            "raw_cat": "石材及石制品",
            "expected_code_prefix": "08",
            "expected_attrs": ["厚度(mm)"]
        },
        {
            "name": "UPVC排水管",
            "spec": "De50*2.0",
            "raw_cat": "给排水管材",
            "expected_code_prefix": "14",
            "expected_attrs": ["公称外径De(mm)"]
        }
    ]

    for tc in test_cases:
        res = engine.standardize_item(
            raw_name=tc["name"],
            raw_spec=tc["spec"],
            raw_cat=tc["raw_cat"]
        )
        print(f"\n测试物料: {tc['name']} | 规格: {tc['spec']}")
        print(f"  -> 命中分类: {res['level3']} ({res['category_code']}), 依据: {res['reason']}")
        print(f"  -> 官方标准属性列: {res['standard_attributes']}")
        print(f"  -> 物理脱壳拆分结果: {res['extracted_attributes']}")

        assert res["category_code"].startswith(tc["expected_code_prefix"]), f"编码前缀不匹配: {res['category_code']}"
        for ea in tc["expected_attrs"]:
            assert ea in res["extracted_attributes"] or ea in res["standard_attributes"], f"缺失关键属性: {ea}"
        print(f"  [PASS] 测试通过!")

    # 测试整表批量推断
    batch_items = [
        {"id": 1, "name": "热轧圆钢", "spec": "Φ18 Q235B", "raw_cat": "金属及金属制品"},
        {"id": 2, "name": "热轧圆钢", "spec": "Φ25 Q235B", "raw_cat": "金属及金属制品"},
    ]
    batch_res = engine.standardize_batch(batch_items)
    assert len(batch_res) == 2
    print(f"\n批量接口测试通过: 处理 {len(batch_res)} 条物料正常。")
    print("=== 全量自检 100% 验证通过! ===")

if __name__ == "__main__":
    run_tests()
