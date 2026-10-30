# -*- coding: utf-8 -*-
"""
run_batch_standardize.py - 批量自动化执行 QuoteStandardizer
为桌面程雅洁与张艳鸽两批供应商原始文件进行端到端解析并输出至桌面交付目录
"""
import os
import sys
import re
import time
import zipfile
import traceback
from pathlib import Path
import pandas as pd

from main import process_excel, process_pdf
from terms_extractor import find_terms_image, TermsExtractor
from config import DEFAULT_TEMPLATE_PATH

def extract_zip_if_needed(vendor_dir: Path):
    """如果目录内存在未解压的 zip 且没有解压出的核心文件，自动安全解压并自动纠正 GBK 编码"""
    for zpath in vendor_dir.glob("*.zip"):
        extract_dir = zpath.parent / zpath.stem
        if not extract_dir.exists() or len(list(extract_dir.iterdir())) == 0:
            extract_dir.mkdir(parents=True, exist_ok=True)
            try:
                with zipfile.ZipFile(zpath, 'r') as z:
                    for info in z.infolist():
                        # 尝试纠正编码: zipfile 默认把非 utf-8 解码为 cp437，需转回 gbk
                        try:
                            fname = info.filename.encode('cp437').decode('gbk')
                        except Exception:
                            fname = info.filename
                        if fname.startswith("__MACOSX") or "/." in fname:
                            continue
                        target_path = extract_dir / fname
                        if info.is_dir():
                            target_path.mkdir(parents=True, exist_ok=True)
                        else:
                            target_path.parent.mkdir(parents=True, exist_ok=True)
                            with z.open(info) as src, open(target_path, 'wb') as dst:
                                dst.write(src.read())
            except Exception as e:
                print(f"[Warning] 解压 {zpath.name} 失败: {e}")

def find_primary_quote_file(vendor_dir: Path) -> Path:
    """动态寻找供应商目录下的真实原始报价单文件（强化安全隔离版）。
    安全铁律：
    1. 严禁识别任何由线上系统导出的 zip 压缩包及其解压文件夹（特征: _\\d{14}）。
    2. 严禁识别名为 '报价数据.xlsx'、'*标准*.xlsx' 的系统产出表格作为原单。
    3. 排除商务说明、折扣记录等纯条款图片。
    4. 优先挑选供应商真实原厂原单: 原生原厂Excel >> 原厂PDF >> 扫描图片。
    """
    candidates = []
    
    # 遍历所有文件，严格执行防污染黑名单过滤
    for p in vendor_dir.rglob("*"):
        if not p.is_file() or p.name.startswith(".") or "__MACOSX" in str(p):
            continue
        # 1. 绝对拉黑线上系统导出特征
        if re.search(r"_\d{14}", str(p)) or "报价数据.xlsx" in p.name or "标准解析" in p.name:
            continue
        # 2. 排除纯说明和纯折扣图片（若为 PDF/Excel 即使带“说明”也可能是主报价函）
        if p.suffix.lower() in [".png", ".jpg", ".jpeg"] and any(k in p.name for k in ["报价说明", "折扣说明", "折扣记录", "商务说明"]):
            continue
            
        ext = p.suffix.lower()
        if ext in [".xlsx", ".xlsm", ".xls", ".pdf", ".png", ".jpg", ".jpeg"]:
            score = 0
            # 格式权重：真原厂 Excel 优先，PDF 次之，图片
            if ext in [".xlsx", ".xlsm", ".xls"]:
                score += 100
            elif ext == ".pdf":
                score += 80
            else:
                score += 40
                
            # 关键词加分
            name_lower = p.name.lower()
            if any(k in name_lower for k in ["报价", "价格", "清单", "明细", "产品"]):
                score += 20
            if "盖章" in name_lower or "扫描" in name_lower:
                score -= 5
                
            candidates.append((score, p.stat().st_size, p))
            
    if not candidates:
        return None
    # 按得分高、文件大小大排序
    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return candidates[0][2]

def run_single_vendor(vendor_dir: Path, output_dir: Path) -> dict:
    vname = vendor_dir.name
    res = {
        "vendor": vname,
        "input_file": None,
        "file_type": None,
        "status": "FAIL",
        "output_file": None,
        "item_count": 0,
        "duration": 0.0,
        "error_msg": None
    }
    
    start_t = time.time()
    try:
        primary_file = find_primary_quote_file(vendor_dir)
        if not primary_file:
            res["error_msg"] = "未在目录中找到有效的报价单文件"
            return res
            
        res["input_file"] = primary_file.name
        ext = primary_file.suffix.lower()
        res["file_type"] = ext
        
        # 查找商务条款图片
        terms_path = find_terms_image(str(vendor_dir), exclude_file=str(primary_file))
        terms_dict = None
        if terms_path and os.path.exists(terms_path):
            try:
                terms_dict = TermsExtractor(terms_path).extract_terms()
            except Exception as e:
                print(f"[{vname}] 商务条款提取跳过/异常: {e}")
                
        output_dir.mkdir(parents=True, exist_ok=True)
        
        if ext in [".xlsx", ".xlsm"]:
            out_file, failures = process_excel(str(primary_file), DEFAULT_TEMPLATE_PATH, str(output_dir), terms=terms_dict)
        elif ext in [".pdf", ".png", ".jpg", ".jpeg"]:
            out_file, failures = process_pdf(str(primary_file), DEFAULT_TEMPLATE_PATH, str(output_dir), terms=terms_dict)
        else:
            res["error_msg"] = f"不支持的文件类型: {ext}"
            return res
            
        res["output_file"] = out_file
        res["duration"] = round(time.time() - start_t, 2)
        
        # 统计解析出的有效物料条数
        if out_file and os.path.exists(out_file):
            df = pd.read_excel(out_file)
            res["item_count"] = len(df)
            res["status"] = "PASS" if len(df) > 0 else "FAIL"
            if len(df) == 0:
                res["error_msg"] = "输出物料条数为 0"
        else:
            res["error_msg"] = "未生成交付 Excel 文件"
            
    except Exception as e:
        res["duration"] = round(time.time() - start_t, 2)
        res["error_msg"] = f"解析异常: {e}\n{traceback.format_exc()}"
        print(f"[{vname}] 解析发生异常: {e}")
        
    return res

if __name__ == "__main__":
    desktop = Path("/Users/qianqianawodebaobei/Desktop")
    tasks = [
        ("程雅洁", desktop / "程雅洁-10月30日", desktop / "本地自跑解析结果_1030/程雅洁"),
        ("张艳鸽", desktop / "张艳鸽-源文件1030", desktop / "本地自跑解析结果_1030/张艳鸽")
    ]
    
    for person, src_root, out_root in tasks:
        print(f"\n==================== 开始批量自跑: {person} ====================")
        vdirs = sorted([d for d in src_root.iterdir() if d.is_dir() and not d.name.startswith(".")])
        for idx, vd in enumerate(vdirs, 1):
            print(f"\n[{idx}/{len(vdirs)}] 处理: {vd.name}...")
            r = run_single_vendor(vd, out_root)
            print(f"    -> 结果: {r['status']} | 物料: {r['item_count']} 行 | 耗时: {r['duration']}s | 文件: {r['output_file']}")
