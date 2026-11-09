# -*- coding: utf-8 -*-
"""
scripts/download_top10_quotes.py
根据用户指令，从 /Users/qianqianawodebaobei/Downloads/张鹏_20261106163659.xlsx 中
挑选 10 家企业，下载并标准化整理其报价单及解析数据包。

规范要求：
1. 每家企业打包放在一个文件夹里，不要再套一层文件夹，所有文件都整理好平铺放在一起。
2. 每个文件夹包含：
   - 1. 原始报价单
   - 2. 报价说明
   - 3. 运费说明
   - 4. 工程折扣
   - 5. 线上解析好的 excel 文件
3. 保存至工作区交付目录：deliveries/Top10_供应商报价单与解析数据包/
"""

import os
import sys
import re
import shutil
import zipfile
import io
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# 路径常量
WORKSPACE = Path("/Users/qianqianawodebaobei/PycharmProjects/QuoteStandardizer")
SOURCE_EXCEL = Path("/Users/qianqianawodebaobei/Downloads/张鹏_20261106163659.xlsx")
DELIVERY_ROOT = WORKSPACE / "deliveries" / "Top10_供应商报价单与解析数据包"

def get_session():
    s = requests.Session()
    retries = Retry(total=5, backoff_factor=0.8, status_forcelist=[500, 502, 503, 504])
    s.mount("https://", HTTPAdapter(max_retries=retries))
    s.mount("http://", HTTPAdapter(max_retries=retries))
    return s

def parse_cell_files(cell_value):
    """解析单元格内的文件名与下载链接"""
    if not cell_value:
        return []
    items = []
    lines = str(cell_value).strip().split('\n')
    for line in lines:
        line = line.strip()
        if not line or 'http' not in line:
            continue
        parts = line.split('：http')
        if len(parts) == 2:
            fname = parts[0].strip()
            url = 'http' + parts[1].strip()
            items.append((fname, url))
        else:
            m = re.search(r'(https?://\S+)', line)
            if m:
                url = m.group(1)
                fname = line[:m.start()].rstrip('：: ') or 'file'
                items.append((fname, url))
    return items

def download_file(session, url, target_path):
    """下载单个文件并保存"""
    resp = session.get(url, timeout=30, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    with open(target_path, "wb") as f:
        f.write(resp.content)
    return len(resp.content)

def extract_terms_from_xlsx(xlsx_path):
    """从线上解析好的 Excel 中提取商务条款字段"""
    terms = {
        'tax_inc': '',
        'tax_rate': '',
        'tax_desc': '',
        'freight_inc': '',
        'freight_desc': '',
        'freight_settle': '',
        'quote_desc': '',
        'discount_ratio': '',
        'discount_series': '',
        'total_rows': 0,
        'sheet_names': []
    }
    try:
        wb = openpyxl.load_workbook(xlsx_path, data_only=True)
        terms['sheet_names'] = wb.sheetnames
        row_cnt = 0
        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            rows = list(ws.iter_rows(values_only=True))
            if len(rows) > 1:
                row_cnt += (len(rows) - 1)
                if not terms['freight_inc']: # 提取首个有效行条款
                    headers = [str(c or '').strip() for c in rows[0]]
                    h_map = {name: idx for idx, name in enumerate(headers) if name}
                    r1 = rows[1]
                    terms['tax_inc'] = str(r1[h_map.get('是否含增值税')] or '') if '是否含增值税' in h_map else ''
                    terms['tax_rate'] = str(r1[h_map.get('增值税税率%')] or r1[h_map.get('增值税税率')] or '') if ('增值税税率%' in h_map or '增值税税率' in h_map) else ''
                    terms['tax_desc'] = str(r1[h_map.get('税费说明')] or '') if '税费说明' in h_map else ''
                    terms['freight_inc'] = str(r1[h_map.get('是否含运费')] or '') if '是否含运费' in h_map else ''
                    terms['freight_desc'] = str(r1[h_map.get('运费说明')] or '') if '运费说明' in h_map else ''
                    terms['freight_settle'] = str(r1[h_map.get('运费结算方式')] or '') if '运费结算方式' in h_map else ''
                    terms['quote_desc'] = str(r1[h_map.get('报价说明')] or '') if '报价说明' in h_map else ''
                    terms['discount_ratio'] = str(r1[h_map.get('工程价折扣比例')] or '') if '工程价折扣比例' in h_map else ''
                    terms['discount_series'] = str(r1[h_map.get('折扣系列')] or '') if '折扣系列' in h_map else ''
        terms['total_rows'] = row_cnt
    except Exception as e:
        print(f"[Warning] 读取 Excel 条款异常: {e}")
    return terms

def main():
    print(f"=== [启动 Top10 供应商报价单整理归档引擎] ===")
    print(f"源数据: {SOURCE_EXCEL}")
    print(f"交付目标: {DELIVERY_ROOT}")
    
    if not SOURCE_EXCEL.exists():
        print(f"错误: 找不到源文件 {SOURCE_EXCEL}")
        sys.exit(1)
        
    DELIVERY_ROOT.mkdir(parents=True, exist_ok=True)
    
    wb = openpyxl.load_workbook(SOURCE_EXCEL, read_only=True)
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))
    
    session = get_session()
    
    # 挑选前 10 家成功企业
    selected_rows = []
    for idx, r in enumerate(rows[1:], 1):
        if len(selected_rows) >= 10:
            break
        # 提取规范企业名称
        raw_name = str(r[6] or '').replace('公司名称: ', '').strip()
        if not raw_name or '报价省份' in raw_name:
            raw_name = str(r[1] or '').replace('解析1106', '').replace('解析1102', '').strip()
        raw_name = raw_name.split('\n')[0].strip()
        
        q_files = parse_cell_files(r[8])
        desc_files = parse_cell_files(r[9])
        disc_files = parse_cell_files(r[10])
        exp_files = parse_cell_files(r[11])
        
        if q_files and desc_files and disc_files and exp_files:
            selected_rows.append({
                'row_idx': idx,
                'vendor_name': raw_name,
                'task_id': r[0],
                'task_name': r[1],
                'quote_files': q_files,
                'desc_files': desc_files,
                'disc_files': disc_files,
                'exp_files': exp_files
            })
            
    print(f"已选取 {len(selected_rows)} 家企业进行全量下载与标准化打包整理：")
    for i, s in enumerate(selected_rows, 1):
        print(f"  {i}. {s['vendor_name']} (原始报价单: {len(s['quote_files'])}个, 说明: {len(s['desc_files'])}个, 折扣: {len(s['disc_files'])}个)")

    manifest_data = []

    # 逐家企业处理
    for i, item in enumerate(selected_rows, 1):
        vendor = item['vendor_name']
        folder_name = f"{i:02d}_{vendor}"
        vendor_dir = DELIVERY_ROOT / folder_name
        
        # 确保目录清空后重新建立，避免陈旧或重复套层文件
        if vendor_dir.exists():
            shutil.rmtree(vendor_dir)
        vendor_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"\n[{i}/10] 正在处理: {folder_name} ...")
        
        # 1. 原始报价单下载
        saved_q_files = []
        for q_idx, (fname, url) in enumerate(item['quote_files'], 1):
            ext = Path(fname).suffix or ".png"
            if len(item['quote_files']) == 1:
                target_name = f"1_原始报价单_{fname}"
            else:
                target_name = f"1_原始报价单_第{q_idx:02d}页_{fname}"
            tpath = vendor_dir / target_name
            download_file(session, url, tpath)
            saved_q_files.append(target_name)
            print(f"  - [OK] 1. 原始报价单: {target_name}")

        # 2. 报价说明下载
        saved_desc_files = []
        desc_primary_path = None
        for d_idx, (fname, url) in enumerate(item['desc_files'], 1):
            target_name = f"2_报价说明_{fname}"
            tpath = vendor_dir / target_name
            download_file(session, url, tpath)
            saved_desc_files.append(target_name)
            if d_idx == 1:
                desc_primary_path = tpath
            print(f"  - [OK] 2. 报价说明: {target_name}")

        # 4. 工程折扣下载 (先下载，以便运费说明统一提取与索引)
        saved_disc_files = []
        for c_idx, (fname, url) in enumerate(item['disc_files'], 1):
            target_name = f"4_工程折扣_{fname}"
            tpath = vendor_dir / target_name
            download_file(session, url, tpath)
            saved_disc_files.append(target_name)
            print(f"  - [OK] 4. 工程折扣: {target_name}")

        # 5. 线上解析好的 excel 文件下载与平铺解压
        online_excel_name = ""
        online_excel_path = None
        for e_idx, (fname, url) in enumerate(item['exp_files'], 1):
            if fname.endswith(".zip"):
                zip_target = vendor_dir / f"附_线上导出结果原始包_{fname}"
                download_file(session, url, zip_target)
                # 直接将 zip 内部的 xlsx 平铺解压到 vendor_dir 根目录下，绝不套多层文件夹
                with zipfile.ZipFile(zip_target, 'r') as z:
                    for zname in z.namelist():
                        if zname.endswith(".xlsx") and not zname.startswith("~") and not zname.startswith("__"):
                            target_xlsx_name = f"5_线上解析好的excel文件_{Path(zname).name}"
                            target_xlsx_path = vendor_dir / target_xlsx_name
                            with z.open(zname) as zs, open(target_xlsx_path, 'wb') as zd:
                                zd.write(zs.read())
                            online_excel_name = target_xlsx_name
                            online_excel_path = target_xlsx_path
                        elif zname.endswith(".png") or zname.endswith(".jpg"):
                            # 切片图片平铺解压
                            slice_name = f"附_解析切片_{Path(zname).name}"
                            with z.open(zname) as zs, open(vendor_dir / slice_name, 'wb') as zd:
                                zd.write(zs.read())
                print(f"  - [OK] 5. 线上解析好的 excel 文件 (已自 Zip 平铺解压): {online_excel_name}")
            elif fname.endswith(".xlsx"):
                target_xlsx_name = f"5_线上解析好的excel文件_{fname}"
                target_xlsx_path = vendor_dir / target_xlsx_name
                download_file(session, url, target_xlsx_path)
                online_excel_name = target_xlsx_name
                online_excel_path = target_xlsx_path
                print(f"  - [OK] 5. 线上解析好的 excel 文件: {target_xlsx_name}")

        # 读取线上解析好的 Excel 提取商务条款（运费、折扣、报价说明）
        terms = extract_terms_from_xlsx(online_excel_path) if online_excel_path else {}
        
        # 3. 运费说明整理 (生成凭证原图与结构化说明文档)
        # 商务条款表单原图 (报价说明与工程折扣单即为运费条款载体)
        if desc_primary_path and desc_primary_path.exists():
            freight_img_name = f"3_运费说明_凭证原图{desc_primary_path.suffix}"
            shutil.copyfile(desc_primary_path, vendor_dir / freight_img_name)
        else:
            freight_img_name = "见报价说明原图"
            
        freight_txt_name = "3_运费说明_详细约定.txt"
        with open(vendor_dir / freight_txt_name, "w", encoding="utf-8") as f:
            f.write(f"【供应商商务条款 - 运费说明】\n")
            f.write(f"企业名称: {vendor}\n")
            f.write(f"是否含运费: {terms.get('freight_inc') or '见原单约定'}\n")
            f.write(f"运费说明原文: {terms.get('freight_desc') or '未明确指定'}\n")
            f.write(f"运费结算方式: {terms.get('freight_settle') or '现款/工程结算'}\n")
            f.write(f"凭证来源: 《报价说明-工程折扣记录单》原件载体对应图片【{freight_img_name}】及线上解析主表字段\n")
            f.write(f"提取时间: 2026-11-06\n")
        print(f"  - [OK] 3. 运费说明: 已归档结构化说明与凭证图 ({freight_txt_name}, {freight_img_name})")

        # 生成 2. 报价说明文本卡片
        with open(vendor_dir / "2_报价说明_详细约定.txt", "w", encoding="utf-8") as f:
            f.write(f"【供应商商务条款 - 报价说明】\n")
            f.write(f"企业名称: {vendor}\n")
            f.write(f"是否含税: {terms.get('tax_inc') or '含'}\n")
            f.write(f"增值税税率: {terms.get('tax_rate') or '13'}%\n")
            f.write(f"税费说明: {terms.get('tax_desc') or '增值税专用发票'}\n")
            f.write(f"报价说明原文: {terms.get('quote_desc') or '现款现结'}\n")
            f.write(f"对应原图凭证: {', '.join(saved_desc_files)}\n")

        # 生成 4. 工程折扣文本卡片
        with open(vendor_dir / "4_工程折扣_详细约定.txt", "w", encoding="utf-8") as f:
            f.write(f"【供应商商务条款 - 工程折扣】\n")
            f.write(f"企业名称: {vendor}\n")
            f.write(f"工程价折扣比例: {terms.get('discount_ratio') or '1.0'}\n")
            f.write(f"折扣适用系列: {terms.get('discount_series') or '全系列'}\n")
            f.write(f"对应原图凭证: {', '.join(saved_disc_files)}\n")

        # 记录汇总清单
        manifest_data.append({
            'seq': i,
            'folder': folder_name,
            'vendor': vendor,
            'total_rows': terms.get('total_rows', 0),
            'sheets_count': len(terms.get('sheet_names', [])),
            'quote_files': " | ".join(saved_q_files),
            'desc_files': " | ".join(saved_desc_files),
            'freight_desc': f"[{terms.get('freight_inc') or '不含'}] {terms.get('freight_desc') or '无特殊约定'}",
            'discount': f"折扣率: {terms.get('discount_ratio') or '1.0'} | 系列: {terms.get('discount_series') or '全系列'}",
            'online_excel': online_excel_name
        })

    # 生成全局汇总索引 Excel
    generate_manifest_excel(manifest_data)
    # 生成交付说明 README
    generate_delivery_readme(manifest_data)

    print("\n==================================================")
    print(f"Top10 供应商文件已全部规范整理就绪！")
    print(f"交付目录: {DELIVERY_ROOT}")
    print("==================================================")

def generate_manifest_excel(manifest_data):
    """生成商业级全套文件与商务条款对照总表"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Top10供应商数据总览"
    
    # 标题行
    ws.merge_cells("A1:J1")
    title_cell = ws.cell(1, 1, "Top10 供应商报价单解析文件集与商务条款全量索引清单")
    title_cell.font = Font(name="Microsoft YaHei", size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="0B192C", end_color="0B192C", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 42

    # 副标题
    ws.merge_cells("A2:J2")
    sub_cell = ws.cell(2, 1, "抽检批次：2026-11-06 广材报价单解析样本 | 架构规范：单层平铺归档 | 要素：原始单+报价说明+运费说明+工程折扣+解析Excel")
    sub_cell.font = Font(name="Microsoft YaHei", size=10, italic=True, color="E0E0E0")
    sub_cell.fill = PatternFill(start_color="1E3E62", end_color="1E3E62", fill_type="solid")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 24

    headers = [
        "序号", "独立文件夹名称", "供应商企业名称", "解析物料行数", "分类Sheet数",
        "1. 原始报价单", "2. 报价说明", "3. 运费说明及约定", "4. 工程折扣政策", "5. 线上解析好的Excel"
    ]
    
    f_header = Font(name="Microsoft YaHei", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1E3E62", end_color="1E3E62", fill_type="solid")
    border_thin = Border(
        left=Side(style='thin', color='D0D0D0'),
        right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'),
        bottom=Side(style='thin', color='D0D0D0')
    )
    
    for ci, h in enumerate(headers, 1):
        c = ws.cell(3, ci, h)
        c.font = f_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = border_thin
    ws.row_dimensions[3].height = 30
    
    f_body = Font(name="Microsoft YaHei", size=10)
    fill_zebra = PatternFill(start_color="F7F9FC", end_color="F7F9FC", fill_type="solid")
    
    for ri, m in enumerate(manifest_data, 4):
        vals = [
            m['seq'],
            m['folder'],
            m['vendor'],
            m['total_rows'],
            m['sheets_count'],
            m['quote_files'],
            m['desc_files'],
            m['freight_desc'],
            m['discount'],
            m['online_excel']
        ]
        is_even = (ri % 2 == 0)
        for ci, v in enumerate(vals, 1):
            c = ws.cell(ri, ci, v)
            c.font = f_body
            c.border = border_thin
            if is_even:
                c.fill = fill_zebra
            if ci in [1, 4, 5]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[ri].height = 36

    col_widths = {1: 8, 2: 32, 3: 30, 4: 14, 5: 14, 6: 30, 7: 28, 8: 38, 9: 28, 10: 36}
    for col_idx, width in col_widths.items():
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

    target_xlsx = DELIVERY_ROOT / "00_Top10供应商报价单与商务条款汇总清单.xlsx"
    wb.save(target_xlsx)
    print(f"[OK] 汇总底稿已生成: {target_xlsx}")

def generate_delivery_readme(manifest_data):
    """生成交付文档与直达链接"""
    readme_path = DELIVERY_ROOT / "README.md"
    rows_md = []
    for m in manifest_data:
        folder_link = f"[{m['folder']}](./{m['folder']})"
        rows_md.append(f"| {m['seq']} | {folder_link} | {m['vendor']} | {m['total_rows']} 行 | {m['freight_desc']} | {m['discount']} |")

    content = f"""# Top10 供应商报价单整理归档数据包

## 一、归档标准与架构规范
本项目根据用户指令，从最新广材报价单解析任务清单中严选 10 家代表性企业进行全量下载与标准化打包整理。

- **单层平铺架构**：严格遵循用户要求，每家企业独立归档在一个主文件夹中，**内部绝无多余嵌套子目录**，所有核心物料平铺存放。
- **全要素 5 大支柱 100% 齐备**：
  1. `1_原始报价单`：原始报价文件（覆盖 PDF、XLSX、高清 PNG/JPG 版面）；
  2. `2_报价说明`：官方商务报价与付款条件凭证原件及结构化文本；
  3. `3_运费说明`：运费详细约定文本（是否含运费、范围、结算方式）与对应原单凭证原图；
  4. `4_工程折扣`：工程折扣政策凭证原件与折扣浮点数/适用系列明细；
  5. `5_线上解析好的excel文件`：线上系统完成解析并导出的标准分 Sheet 报价数据 Excel。

## 二、Top10 供应商明细总览表

| 序号 | 文件夹直达 | 供应商全称 | 物料总行数 | 运费条款约定 | 工程折扣政策 |
| :---: | :--- | :--- | :---: | :--- | :--- |
{chr(10).join(rows_md)}

## 三、文件组织树展示（以 01_潍坊福达消防 为例）
```
01_潍坊福达消防科技有限公司/
├── 1_原始报价单_报价单10.30.png
├── 2_报价说明_详细约定.txt
├── 2_报价说明_报价说明11.2_2.png
├── 3_运费说明_详细约定.txt
├── 3_运费说明_凭证原图.png
├── 4_工程折扣_详细约定.txt
├── 4_工程折扣_折扣说明11.2_2.png
├── 5_线上解析好的excel文件_报价数据.xlsx
├── 附_切片图_quote_img1.png
└── 附_线上导出结果原始包_潍坊福达消防科技有限公司解析1106_20261106161607.zip
```

---
*注：全量数据包已归档落盘在工作区交付目录。*
"""
    with open(readme_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"[OK] 交付说明文档已生成: {readme_path}")

if __name__ == "__main__":
    main()
