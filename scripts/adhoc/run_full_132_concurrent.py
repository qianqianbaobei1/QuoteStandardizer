# -*- coding: utf-8 -*-
"""
run_full_132_concurrent.py - 全量 132 家厂商多线程并发高精度解析引擎
【核心规范与约束】：
1. 覆盖全量 132 家原始文件（四位经办人完整档案库）
2. 使用标准材料类别库与 CATEGORY_ATTRS_API_URL 指向的实时属性接口
3. 标准化品名与类别编码（由 Jev 概率最高裁决）
4. 严格剥离非物理技术参数（基价、单价、表标题等转入说明）
5. 纯数值价格，纯单 Sheet (标准化报价)
6. 多进程/线程并发流水线提速
"""
from _adhoc_paths import SOURCE_ROOT
import os
import re
import sys
import time
import shutil
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import openpyxl

# 加载环境
from config import CATEGORY_CODE_MAP, CATEGORY_ATTRS_API_URL
from run_batch_standardize import find_primary_quote_file, run_single_vendor, verify_delivery_file
from zip_standardizer_pipeline import ZipQuoteStandardizerPipeline
from style_enhancer import apply_enterprise_style
from llm_auditor import LLMAuditor

output_dir = Path('deliveries')
output_dir.mkdir(parents=True, exist_ok=True)

# 1. 扫描原始供应商目录：命令行第一个参数 > 环境变量 QS_SOURCE_ROOT > 默认全量经办人目录
DEFAULT_SOURCE_ROOT = SOURCE_ROOT
cli_arg = sys.argv[1] if (len(sys.argv) > 1 and not sys.argv[1].startswith('-')) else ""
source_path_str = cli_arg or os.getenv('QS_SOURCE_ROOT', DEFAULT_SOURCE_ROOT)
root_agents = Path(source_path_str)

if not str(root_agents).strip() or str(root_agents) in ('.', './') or not root_agents.is_dir():
    print(f'未找到有效源文件目录: {root_agents}')
    print('用法: python3 run_full_132_concurrent.py [<02_四位经办人分类完整档案 目录>] [--force]')
    sys.exit(1)
all_tasks = []

for agent_dir in sorted(root_agents.iterdir()):
    if not agent_dir.is_dir() or agent_dir.name.startswith('.'): continue
    for sub in agent_dir.iterdir():
        if '原始' in sub.name and sub.is_dir():
            for v in sorted(sub.iterdir()):
                if v.is_dir() and not v.name.startswith('.'):
                    all_tasks.append((agent_dir.name, v.name, v))

def _clean_name(vendor_name: str) -> str:
    return re.sub(r'^\d+[\._、\s]*', '', vendor_name).strip()


FORCE_RERUN = os.getenv('FORCE_RERUN', '0').strip().lower() in ('1', 'true') or '--force' in sys.argv

def already_done(vendor_name: str):
    """断点续跑：若开启 FORCE_RERUN 则全部重跑；否则交付件存在且体检合格才跳过"""
    if FORCE_RERUN:
        return False
    xlsx = output_dir / f"{_clean_name(vendor_name)}-标准解析.xlsx"
    if not xlsx.exists():
        return False
    return not verify_delivery_file(str(xlsx))


print(f"==================== 全量 132 家建材报价单并发标准化解析启动 ====================")
print(f"  -> 待解析目标总数: {len(all_tasks)} 家")
print(f"  -> 标准三级分类库容量: {len(CATEGORY_CODE_MAP)} 条")
print(f"  -> 属性规范接口: {'已配置' if CATEGORY_ATTRS_API_URL else '未配置(仅用本地编码库)'}")

pipeline_zip = ZipQuoteStandardizerPipeline(str(output_dir))
auditor = LLMAuditor()

def process_single_task(args):
    agent_name, vendor_name, vendor_path = args
    start_t = time.time()
    # 供应商规范清洗名称
    clean_vname = _clean_name(vendor_name)
    if already_done(vendor_name):
        print(f"[SKIP] [{agent_name}] {clean_vname} 已交付且体检合格，跳过")
        return (clean_vname, agent_name, "SKIP", "已完成", 0, 0)
    
    primary_file = find_primary_quote_file(vendor_path)
    if not primary_file:
        return (clean_vname, agent_name, "FAIL", "未找到主报价文件", 0, 0)
        
    ext = primary_file.suffix.lower()
    out_xlsx = None
    try:
        if ext == '.zip':
            out_xlsx = pipeline_zip.process_zip(str(primary_file), clean_vname)
        elif ext in ['.xlsx', '.xls', '.pdf', '.png', '.jpg', '.jpeg']:
            res_dict = run_single_vendor(vendor_path, output_dir)
            out_xlsx = res_dict.get('output_file')
            # 完整性校验未过就是 FAIL：有文件不等于可信
            if res_dict.get('failures'):
                return (clean_vname, agent_name, "FAIL",
                        f"完整性校验未通过 {len(res_dict['failures'])} 项: {res_dict['failures'][0]}", 0, 0)
            
        if out_xlsx and os.path.exists(out_xlsx):
            # 标准统一重命名为 <供应商名称>-标准解析.xlsx
            target_final_xlsx = str(output_dir / f"{clean_vname}-标准解析.xlsx")
            if os.path.abspath(out_xlsx) != os.path.abspath(target_final_xlsx):
                try:
                    shutil.move(out_xlsx, target_final_xlsx)
                    out_xlsx = target_final_xlsx
                except Exception:
                    pass
                    
            wb = openpyxl.load_workbook(out_xlsx)
            # 严格保证纯单 Sheet：只留「标准化报价」；没有主表时退回首个 Sheet 并改名
            if '标准化报价' in wb.sheetnames:
                ws = wb['标准化报价']
                for s in list(wb.sheetnames):
                    if s != '标准化报价':
                        wb.remove(wb[s])
            else:
                ws = wb.worksheets[0]
                ws.title = '标准化报价'
                
            headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
            
            # 交付前模型业务逻辑终审与就地自愈
            try:
                auditor.verify_and_auto_heal(ws, headers)
            except Exception as e:
                print(f"    [{clean_vname}] 自检跳过: {e}")
                
            # 美学样式刷新
            apply_enterprise_style(ws)
            wb.save(out_xlsx)
            
            rows_cnt = ws.max_row - 1
            dur = round(time.time() - start_t, 2)
            print(f"[✓ PASS] [{agent_name}] {clean_vname} 解析成功: {rows_cnt} 行 ({dur}s)")
            return (clean_vname, agent_name, "PASS", out_xlsx, rows_cnt, dur)
        else:
            return (clean_vname, agent_name, "FAIL", "未能生成有效标准表", 0, 0)
    except Exception as e:
        return (clean_vname, agent_name, "FAIL", str(e), 0, 0)

def main():
    # 使用 4 线程并发执行，平衡 API 吞吐与系统负载
    max_workers_env = os.getenv("MAX_WORKERS", "4")
    MAX_WORKERS = int(max_workers_env) if max_workers_env.isdigit() else 4
    completed_results = []

    print(f"\n[*] 启动 {MAX_WORKERS} 线程池执行...")
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(process_single_task, task): task for task in all_tasks}
        for fut in as_completed(futures):
            res = fut.result()
            completed_results.append(res)
            vname, aname, st, info, cnt, d = res
            print(f"进度: {len(completed_results)}/{len(all_tasks)} | {vname}: {st} ({cnt}行)")

    print("\n==================== 全量 132 家解析汇总报告 ====================")
    success_cnt = sum(1 for r in completed_results if r[2] == "PASS")
    skip_cnt = sum(1 for r in completed_results if r[2] == "SKIP")
    fail_cnt = sum(1 for r in completed_results if r[2] not in ("PASS", "SKIP"))
    total_items = sum(r[4] for r in completed_results)
    print(f"总计: {len(completed_results)} 家 | 成功: {success_cnt} 家 | 跳过(已完成): {skip_cnt} 家 | "
          f"失败/空单: {fail_cnt} 家 | 本次解析标准物料: {total_items} 条")
    report_path = output_dir / "_批量解析结果.txt"
    with open(report_path, "w", encoding="utf-8") as fh:
        fh.write(f"总计 {len(completed_results)} 家 | 成功 {success_cnt} | 跳过 {skip_cnt} | 失败 {fail_cnt}\n\n")
        for vname, aname, st, info, cnt, d in sorted(completed_results, key=lambda x: (x[2], x[0])):
            fh.write(f"{st:5s} {cnt:5d}行 [{aname}] {vname}\n        {info}\n")
    print(f"[报告] 明细已写入 {report_path}")
    if fail_cnt:
        print("\n失败清单:")
        for vname, aname, st, info, cnt, d in completed_results:
            if st not in ('PASS', 'SKIP'):
                print(f"  ✗ [{aname}] {vname}: {info}")
    # 任一家失败即退出码 2（与 run_batch_standardize 口径一致）
    sys.exit(0 if fail_cnt == 0 else 2)


if __name__ == '__main__':
    main()
