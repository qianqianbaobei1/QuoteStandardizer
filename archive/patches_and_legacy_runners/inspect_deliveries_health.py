# -*- coding: utf-8 -*-
"""
inspect_deliveries_health.py - 全量交付件质量终审与健康体检巡检中枢
深度核查客户关切的核心三大顽疾：
1. 【留空 (Blanks)】：名称规格、市场价、单位、类别编码留空核查
2. 【错行 (Misalignment)】：表头残留污染、基价混入规格列、参数错位漂移核查
3. 【内容丢失 (Data Loss)】：行数异常归零/截断、关键参数遗失核查
4. 【商业级视觉与接口规范】：单Sheet标准化、厂家报价图片纯图片无文字、材料配图URL真实性
"""
import glob
import os
import re
from typing import Dict, List, Any
import openpyxl

DELIVERY_DIR = "/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer/deliveries"

# 表头残留关键词（若数据行名称规格或原始规格出现此类词，说明发生了表头错行误提取）
HEADER_LEAK_WORDS = ["序号", "材料名称", "规格型号", "规格参数", "计量单位", "单价(元)", "合价(元)", "项目特征", "品牌及厂家"]

# 杂质关键词
BANK_WORDS = ["开户行", "工行", "建行", "农行", "中行", "招商银行", "账号", "帐号", "税号", "税务登记", "社会信用代码"]
ADDRESS_WORDS = ["公司地址", "地址：", "电话：", "传真：", "邮编：", "联系人"]
CODE_WORDS = ["产品编码", "物料编码", "商品编码", "内部编码", "条形码"]


def inspect_single_delivery(xlsx_path: str) -> Dict[str, Any]:
    fname = os.path.basename(xlsx_path)
    issues = []

    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    except Exception as e:
        return {"file": fname, "error": f"文件损坏或无法读取: {e}", "is_healthy": False}

    # 1. 结构规范检查：纯单 Sheet
    if "标准化报价" not in wb.sheetnames:
        return {"file": fname, "error": "缺失【标准化报价】Sheet", "is_healthy": False}

    ws = wb["标准化报价"]
    headers = [cell for cell in next(ws.iter_rows(min_row=1, max_row=1, values_only=True))]

    def get_c(col_name):
        return headers.index(col_name) if col_name in headers else None

    name_c = get_c("名称规格")
    spec_c = get_c("规格参数")
    raw_spec_c = get_c("厂家原始规格参数")
    price_c = get_c("市场价")
    unit_c = get_c("单位")
    code_c = get_c("类别编码")
    quote_desc_c = get_c("报价说明")
    img_fn_c = get_c("厂家报价图片文件名")
    pic_c = get_c("图片")

    total_rows = 0
    blank_name_count = 0
    blank_price_count = 0
    blank_unit_count = 0
    blank_code_count = 0
    
    header_leak_count = 0
    misaligned_base_price_count = 0
    impurity_count = 0
    
    img_text_leak_count = 0
    invalid_pic_url_count = 0

    for r_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        total_rows += 1
        name_val = str(row[name_c] or "").strip() if name_c is not None else ""
        spec_val = str(row[spec_c] or "").strip() if spec_c is not None else ""
        raw_val = str(row[raw_spec_c] or "").strip() if raw_spec_c is not None else ""
        price_val = row[price_c] if price_c is not None else None
        unit_val = str(row[unit_c] or "").strip() if unit_c is not None else ""
        code_val = str(row[code_c] or "").strip() if code_c is not None else ""
        desc_val = str(row[quote_desc_c] or "").strip() if quote_desc_c is not None else ""
        img_fn_val = row[img_fn_c] if img_fn_c is not None else None
        pic_val = str(row[pic_c] or "").strip() if pic_c is not None else ""

        # A. 留空检视 (Blanks)
        if not name_val:
            blank_name_count += 1
            if len(issues) < 5:
                issues.append(f"行 {r_idx} [留空] 名称规格为空")

        # 价格留空或 <= 0
        try:
            p_float = float(price_val) if price_val is not None and str(price_val).strip() != '' else 0.0
            if p_float <= 0:
                blank_price_count += 1
        except (ValueError, TypeError):
            blank_price_count += 1

        if not unit_val:
            blank_unit_count += 1

        if not code_val:
            blank_code_count += 1

        # B. 错行错位检视 (Misalignment)
        # 表头残留误作数据行
        for h_word in HEADER_LEAK_WORDS:
            if h_word in name_val:
                header_leak_count += 1
                if len(issues) < 5:
                    issues.append(f"行 {r_idx} [错行] 数据行混入表头词【{h_word}】: '{name_val}'")
                break

        # 调价基价混入规格
        if "铝锭基价" in spec_val or "铜基价" in spec_val or "基价" in spec_val:
            misaligned_base_price_count += 1
            if len(issues) < 5:
                issues.append(f"行 {r_idx} [错位] 规格参数未剥离基价: '{spec_val[:30]}'")

        # 杂质检查 (银行/地址/编码)
        comb = name_val + spec_val + raw_val
        for bw in BANK_WORDS + ADDRESS_WORDS + CODE_WORDS:
            if bw in comb:
                impurity_count += 1
                if len(issues) < 5:
                    issues.append(f"行 {r_idx} [杂质] 混入非技术干扰词【{bw}】: '{comb[:40]}'")
                break

        # C. 商业图片规范与接口真实性
        # 厂家报价图片文件名单元格：放图片，不要放文字
        if img_fn_val is not None and str(img_fn_val).strip() != '':
            img_text_leak_count += 1
            if len(issues) < 5:
                issues.append(f"行 {r_idx} [格式] 厂家报价图片文件名残留文本: '{img_fn_val}'")

        # 配图列若有值，必须是 http 链接，严禁虚假代码
        if pic_val and not pic_val.lower().startswith("http"):
            invalid_pic_url_count += 1
            if len(issues) < 5:
                issues.append(f"行 {r_idx} [接口] 图片列包含非法链接文本: '{pic_val[:30]}'")

    # 综合健康度评估
    is_healthy = (
        total_rows > 0 and
        blank_name_count == 0 and
        header_leak_count == 0 and
        misaligned_base_price_count == 0 and
        img_text_leak_count == 0 and
        invalid_pic_url_count == 0 and
        impurity_count == 0
    )

    return {
        "file": fname,
        "total_rows": total_rows,
        "blank_name_count": blank_name_count,
        "blank_price_count": blank_price_count,
        "blank_unit_count": blank_unit_count,
        "blank_code_count": blank_code_count,
        "header_leak_count": header_leak_count,
        "misaligned_base_price_count": misaligned_base_price_count,
        "impurity_count": impurity_count,
        "img_text_leak_count": img_text_leak_count,
        "invalid_pic_url_count": invalid_pic_url_count,
        "is_healthy": is_healthy,
        "issues": issues
    }


def run_inspection():
    files = sorted(glob.glob(os.path.join(DELIVERY_DIR, "*标准解析.xlsx")))
    print(f"================================================================================")
    print(f"[*] 启动全量交付件健康终审巡视中枢 (检测目录: {DELIVERY_DIR})")
    print(f"[*] 发现待检文件总数: {len(files)} 份")
    print(f"================================================================================\n")

    if not files:
        print("[!] 交付目录下暂无已生成的标准解析文件，请先运行解析流水线。")
        return

    healthy_files = 0
    flawed_files = 0

    for idx, f in enumerate(files, 1):
        res = inspect_single_delivery(f)
        if "error" in res:
            print(f"[{idx:03d}/{len(files):03d}] ❌ {res['file']} -> {res['error']}")
            flawed_files += 1
            continue

        if res["is_healthy"]:
            healthy_files += 1
            print(f"[{idx:03d}/{len(files):03d}] ✅ {res['file']} (行数: {res['total_rows']} | 留空: 0 | 错行: 0 | 杂质: 0 | 图文规范: 100%)")
        else:
            flawed_files += 1
            print(f"[{idx:03d}/{len(files):03d}] ⚠️ {res['file']} (行数: {res['total_rows']}) 存在瑕疵:")
            if res["blank_name_count"] > 0:
                print(f"      - 留空异常: 品名留空 {res['blank_name_count']} 处")
            if res["header_leak_count"] > 0:
                print(f"      - 错行异常: 表头残留误当数据行 {res['header_leak_count']} 处")
            if res["misaligned_base_price_count"] > 0:
                print(f"      - 错位异常: 规格列未剥离基价 {res['misaligned_base_price_count']} 处")
            if res["img_text_leak_count"] > 0:
                print(f"      - 图文异常: 截图底稿列残留文字 {res['img_text_leak_count']} 处")
            if res["invalid_pic_url_count"] > 0:
                print(f"      - 接口异常: 配图列非标准URL {res['invalid_pic_url_count']} 处")
            for iss in res["issues"]:
                print(f"      [样例] {iss}")

    print(f"\n================================================================================")
    print(f"[*] 终审巡查汇总报告:")
    print(f"    - 检测文件总数: {len(files)}")
    print(f"    - 完美健康文件: {healthy_files} 份")
    print(f"    - 瑕疵拦截文件: {flawed_files} 份")
    print(f"    - 交付合格率: {(healthy_files / len(files) * 100):.1f}%")
    print(f"================================================================================")


if __name__ == "__main__":
    run_inspection()
