# -*- coding: utf-8 -*-
"""
batch_four_layer_audit.py
基于《报价单解析质量对比方案（四层校验）》（@轩轩 2026-11-05）
对 50 份供应商报价单进行全网四层校验统一扫描与总报告生成。
"""
import os
import sys
import json
import openpyxl
from four_layer_validator import run_four_layer_audit_for_vendor, export_five_sheet_excel

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DELIVERIES_DIR = os.path.join(BASE_DIR, 'deliveries', 'batch_50_eval')
THEIRS_DIR = os.path.join(DELIVERIES_DIR, 'theirs')

def run_all_50_vendors():
    theirs_vendors = sorted(os.listdir(THEIRS_DIR))
    print(f"开始批量运行 50 份供应商四层校验比对...")
    
    all_row_verdicts = []
    all_diff_details = []
    all_standard_fixes = []
    
    total_origin_rows = 0
    total_std_rows = 0
    total_our_rows = 0
    total_pass_rows = 0
    total_l1_pass = 0
    total_l2_pass = 0
    total_l3_pass = 0
    total_l4_pass = 0
    
    vendor_summaries = []

    for vname in theirs_vendors:
        theirs_xlsx = os.path.join(THEIRS_DIR, vname, '报价数据.xlsx')
        if not os.path.exists(theirs_xlsx):
            continue
        try:
            stats = run_four_layer_audit_for_vendor(vname)
            vendor_summaries.append(stats)
            
            total_origin_rows += stats['origin_rows']
            total_std_rows += stats['std_rows']
            total_our_rows += stats['our_rows']
            total_pass_rows += stats['pass_cnt']
            total_l1_pass += stats['l1_pass']
            total_l2_pass += stats['l2_pass']
            total_l3_pass += stats['l3_pass']
            total_l4_pass += stats['l4_pass']
            
            all_row_verdicts.extend(stats['row_verdicts'])
            all_diff_details.extend(stats['diff_details'])
            all_standard_fixes.extend(stats['standard_fixes'])
        except Exception as e:
            print(f"Error processing {vname}: {e}")

    # 合并生成全局统计
    from collections import defaultdict
    root_cause_counts = defaultdict(lambda: {'affected_rows': 0, 'root_causes': set(), 'level': ''})
    for d in all_diff_details:
        code = d['issue_code']
        lvl = d['level']
        root_cause_counts[code]['affected_rows'] += 1
        root_cause_counts[code]['level'] = lvl
        if code in ['L1-TERMS', 'L1-DATE', 'L1-SUPPLIER', 'L4-SRC-PAGE', 'L4-SRC-FILE']:
            root_cause_counts[code]['root_causes'].add(f"{d['row_id']}_PAGE")
        else:
            root_cause_counts[code]['root_causes'].add(d['row_id'])

    global_stats = {
        'vendor': '全网 50 批供应商汇总',
        'origin_rows': total_origin_rows,
        'std_rows': total_std_rows,
        'our_rows': total_our_rows,
        'matched_rows': total_pass_rows,
        'overall_pass_rate': total_pass_rows / total_origin_rows if total_origin_rows else 0,
        'pass_cnt': total_pass_rows,
        'l1_pass': total_l1_pass,
        'l2_pass': total_l2_pass,
        'l3_pass': total_l3_pass,
        'l4_pass': total_l4_pass,
        'diff_details': all_diff_details,
        'row_verdicts': all_row_verdicts,
        'root_cause_counts': root_cause_counts,
        'standard_fixes': all_standard_fixes
    }

    out_file = os.path.join(DELIVERIES_DIR, "全网50批报价单_四层校验解析质量对比总览报告.xlsx")
    export_five_sheet_excel(global_stats, out_file)
    print(f"\n==================== 全网 50 批四层校验对账完成 ====================")
    print(f"总报告已成功生成落盘: {out_file}")

if __name__ == '__main__':
    run_all_50_vendors()
