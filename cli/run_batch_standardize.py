import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

# -*- coding: utf-8 -*-
"""
run_batch_standardize.py - 批量解析一个批次下的所有供应商原文件

用法: python3 run_batch_standardize.py <源目录> <输出目录> [--label 批次名]
"""
import os
import sys
import re
import time
import zipfile
import traceback
from pathlib import Path

from main import process_excel, process_pdf, _field_value_problems, COMPANY_RE
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

def _count_delivery_rows(path: str) -> int:
    """交付件数据行数（用已依赖的 openpyxl，不再为一个 len() 拉进 pandas）"""
    import openpyxl
    try:
        ws = openpyxl.load_workbook(path, read_only=True, data_only=True)['标准化报价']
    except Exception:
        return 0
    return max(0, (ws.max_row or 1) - 1)


def verify_delivery_file(path: str) -> list:
    """交付件独立体检：文件可读回、主表单 Sheet、有数据行、市场价为正数值。
    给不经过 main._verify_output 的链路（ZIP 流水线、批处理脚本）使用。"""
    import openpyxl
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        return [f"交付文件无法读回: {e}"]
    if '标准化报价' not in wb.sheetnames:
        return [f"缺少「标准化报价」Sheet: {wb.sheetnames}"]
    errs = []
    extra = [s for s in wb.sheetnames if s not in ('客户原始解析', '标准化报价')]
    if extra:
        errs.append(f"主交付含多余 Sheet: {extra}")
    ws = wb['标准化报价']
    col = {str(c.value): c.column for c in ws[1] if c.value}
    if '市场价' not in col:
        return errs + ["缺少「市场价」列"]
    rows = 0
    zero = 0
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, col['市场价']).value
        if v in (None, ''):
            continue
        rows += 1
        try:
            if float(v) <= 0:
                zero += 1
        except (TypeError, ValueError):
            errs.append(f"第{r}行 市场价非数值: {v}")
        if len(errs) > 10:
            break
    if rows == 0:
        errs.append("没有任何数据行")
    # 个别 0 价行由对账/批注负责（原表本就无价）；过半才是真坏了
    elif zero > rows * 0.5:
        errs.append(f"{zero}/{rows} 行市场价不为正")
    # 固定字段值域：品牌列不许是价格，地区列不许是价格/规格串
    all_rows = [r for r in ws.iter_rows(min_row=2, values_only=True)
                if any(v not in (None, '') for v in r)]
    errs.extend(_field_value_problems([c.value for c in ws[1]], all_rows))
    return errs


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
        "error_msg": None,
        "failures": None
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

        # 仅当目录名确实符合企业全称规范 (如“xx有限公司/股份有限公司/厂/经营部”) 时，才允许作为供应商候选；
        # 若目录名仅为批次名或普通编号（如“2026年11月询价批次”、“图1”），严禁覆盖公章和表内识别出的合法企业名称！
        dir_supplier = re.sub(r'^\d+[\._、\s]*', '', vendor_dir.name).strip()
        terms_dict = dict(terms_dict or {})
        detected = terms_dict.get('supplier')
        m_comp = COMPANY_RE.search(dir_supplier)
        if m_comp and terms_dict.get('_source') != 'manual':
            valid_dir_company = m_comp.group(1).strip()
            if not detected:
                terms_dict['supplier'] = valid_dir_company
                terms_dict['_source'] = 'dir'
            elif detected != valid_dir_company:
                terms_dict['_supplier_note'] = f"公章/表内识别为“{detected}”，源目录名为“{valid_dir_company}”"
                print(f"[{vname}] 公章识别为“{detected}”，与目录名“{valid_dir_company}”对照留存旁证")

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
        res["failures"] = failures

        # PASS 只认 process_* 的完整性校验结果 (对账异常/漏行/不守恒)；行数>0 不算通过
        if not out_file or not os.path.exists(out_file):
            res["error_msg"] = "未生成交付 Excel 文件"
            return res
        res["item_count"] = _count_delivery_rows(out_file)
        if failures:
            res["error_msg"] = f"完整性校验未通过 {len(failures)} 项: {failures[0]}"
        elif res["item_count"] == 0:
            res["error_msg"] = "输出物料条数为 0"
        else:
            res["status"] = "PASS"
            
    except Exception as e:
        res["duration"] = round(time.time() - start_t, 2)
        res["error_msg"] = f"解析异常: {e}\n{traceback.format_exc()}"
        print(f"[{vname}] 解析发生异常: {e}")
        
    return res

if __name__ == "__main__":
    import argparse
    ap = argparse.ArgumentParser(description="批量解析一个批次下所有供应商原文件")
    ap.add_argument("src", help="批次源目录 (每家一个子目录)")
    ap.add_argument("out", help="输出目录")
    ap.add_argument("--label", default="", help="批次名，仅用于日志")
    args = ap.parse_args()

    src_root, out_root = Path(args.src), Path(args.out)
    label = args.label or src_root.name
    print(f"\n==================== 开始批量自跑: {label} ====================")
    vdirs = sorted([d for d in src_root.iterdir() if d.is_dir() and not d.name.startswith(".")])
    results = []
    for idx, vd in enumerate(vdirs, 1):
        print(f"\n[{idx}/{len(vdirs)}] 处理: {vd.name}...")
        r = run_single_vendor(vd, out_root)
        results.append(r)
        print(f"    -> 结果: {r['status']} | 物料: {r['item_count']} 行 | 耗时: {r['duration']}s | 文件: {r['output_file']}")
        if r["error_msg"]:
            print(f"    ! {r['error_msg']}")

    passed = sum(1 for r in results if r["status"] == "PASS")
    print(f"\n批次 {label}: {passed}/{len(results)} PASS")
    for r in results:
        if r["status"] != "PASS":
            print(f"  ✗ {r['vendor']}: {r['error_msg']}")
    sys.exit(0 if passed == len(results) else 2)
