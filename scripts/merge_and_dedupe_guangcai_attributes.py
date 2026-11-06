# -*- coding: utf-8 -*-
"""
scripts/merge_and_dedupe_guangcai_attributes.py
广材两份历史大库（762MB + 11GB，共 3657 万行）全量整合、按属性名称分类与属性值去重保留脚本

业务规则：
1. 全量流式整合两份数据源：
   - 文件1: result_广材后台数据库_db_op_gldjc_20261105171750.csv (762 MB / 237 万行)
   - 文件2: result_广材后台数据库_db_op_gldjc_20261105172104.csv (11.04 GB / 3420 万行)
2. 按属性名称（如 '公称外径De(mm)', '公称壁厚(mm)', '品种', '材质', '颜色' 等）分类归集；
3. 对每个属性名称下的属性值进行严格去重（加属性值去重）；
4. 同时在各三级品类维度下进行去重（同分类同规格只保留唯一样本）；
5. 结果全量持久化保留至：
   - SQLite 核心索引库: guangcai_knowledge_base.db
   - JSON 便捷查询字典: data/guangcai_attribute_catalog.json
   - Excel 交互字典: deliveries/广材标准属性分类与取值字典.xlsx
"""
import os
import sys
import csv
import json
import time
import re
import hashlib
import sqlite3
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FILE_1 = "/Users/qianqianawodebaobei/Downloads/result_广材后台数据库_db_op_gldjc_20261105171750.csv"
FILE_2 = "/Users/qianqianawodebaobei/Downloads/result_广材后台数据库_db_op_gldjc_20261105172104.csv"
OUTPUT_DB = "guangcai_knowledge_base.db"
OUTPUT_JSON = "data/guangcai_attribute_catalog.json"
OUTPUT_XLSX = "deliveries/广材标准属性分类与取值字典.xlsx"

REPORT_INTERVAL = 1000000

def run_integration():
    sources = [f for f in [FILE_1, FILE_2] if os.path.exists(f)]
    if not sources:
        print("[Error] 未找到广材数据源文件！")
        sys.exit(1)

    print("=== [广材双大库全量整合与属性分类去重] 启动 ===")
    for idx, s in enumerate(sources, 1):
        print(f"数据源 {idx}: {s} ({os.path.getsize(s) / (1024**3):.2f} GB)")
    print(f"目标数据库: {OUTPUT_DB}")

    os.makedirs("data", exist_ok=True)
    os.makedirs("deliveries", exist_ok=True)

    # 1. 初始化 SQLite 结构
    conn = sqlite3.connect(OUTPUT_DB)
    cur = conn.cursor()
    cur.execute("PRAGMA synchronous = OFF;")
    cur.execute("PRAGMA journal_mode = WAL;")
    cur.execute("PRAGMA cache_size = -128000;")
    cur.execute("PRAGMA temp_store = MEMORY;")

    # 清理并重建表
    cur.execute("DROP TABLE IF EXISTS global_attribute_values;")
    cur.execute("DROP TABLE IF EXISTS category_attr_values;")
    cur.execute("DROP TABLE IF EXISTS category_attribute_schema;")
    cur.execute("DROP TABLE IF EXISTS category_specs;")
    cur.execute("DROP TABLE IF EXISTS product_category_index;")

    # 表1: 全局按属性名称分类的去重属性值库
    cur.execute("""
        CREATE TABLE global_attribute_values (
            attr_name TEXT,
            attr_value TEXT,
            frequency INTEGER,
            unit TEXT,
            PRIMARY KEY (attr_name, attr_value)
        )
    """)

    # 表2: 品类 + 属性名称维度去重属性值库
    cur.execute("""
        CREATE TABLE category_attr_values (
            level3 TEXT,
            attr_name TEXT,
            attr_value TEXT,
            frequency INTEGER,
            PRIMARY KEY (level3, attr_name, attr_value)
        )
    """)

    # 表3: 三级品类标准属性体系表
    cur.execute("""
        CREATE TABLE category_attribute_schema (
            level3 TEXT PRIMARY KEY,
            level1 TEXT,
            level2 TEXT,
            total_specs INTEGER,
            attributes_json TEXT,
            sample_enums_json TEXT
        )
    """)

    # 表4: 同分类同规格去重规格表
    cur.execute("""
        CREATE TABLE category_specs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            level1 TEXT,
            level2 TEXT,
            level3 TEXT,
            product_name TEXT,
            unit TEXT,
            spec_json TEXT,
            spec_hash TEXT
        )
    """)

    # 表5: 产品名称到分类索引表
    cur.execute("""
        CREATE TABLE product_category_index (
            product_name TEXT PRIMARY KEY,
            level1 TEXT,
            level2 TEXT,
            level3 TEXT,
            sample_count INTEGER,
            confidence REAL
        )
    """)
    conn.commit()

    # 内存归集容器
    # 1. attr_name -> attr_val -> count (按属性名称分类，加属性值去重)
    global_attr_values = defaultdict(Counter)
    # 2. cat3 -> attr_name -> attr_val -> count (按品类+属性名称分类)
    cat_attr_values = defaultdict(lambda: defaultdict(Counter))
    # 3. cat3 -> (cat1, cat2)
    cat_hierarchy = {}
    # 4. 同分类同规格指纹池
    seen_spec_hashes = set()
    # 5. 产品名称投票池
    product_category_votes = defaultdict(Counter)

    specs_batch = []
    total_processed = 0
    unique_specs_count = 0
    t_start = time.time()

    for src_idx, src_file in enumerate(sources, 1):
        print(f"\n[*] 正在读取源文件 {src_idx}/{len(sources)}: {os.path.basename(src_file)} ...")
        file_rows = 0
        with open(src_file, "r", encoding="utf-8-sig", errors="replace", buffering=16 * 1024 * 1024) as f:
            reader = csv.reader(f)
            header = next(reader)

            for row in reader:
                total_processed += 1
                file_rows += 1
                if len(row) < 12:
                    continue

                p_name = row[5].strip()
                unit = row[7].strip()
                spec_str = row[8].strip()
                cat1 = row[9].strip()
                cat2 = row[10].strip()
                cat3 = row[11].strip()

                if not cat3 or not spec_str:
                    continue

                if p_name:
                    product_category_votes[p_name][(cat1, cat2, cat3)] += 1

                # 规范化 spec_json
                parsed_d = None
                if spec_str.startswith("{") and spec_str.endswith("}"):
                    try:
                        parsed_d = json.loads(spec_str)
                        norm_json = json.dumps(parsed_d, sort_keys=True, ensure_ascii=False)
                    except Exception:
                        norm_json = spec_str
                else:
                    norm_json = spec_str

                # 同分类同规格去重
                h = hashlib.md5((cat3 + "\n" + norm_json).encode("utf-8")).digest()
                if h not in seen_spec_hashes:
                    seen_spec_hashes.add(h)
                    unique_specs_count += 1
                    hex_hash = h.hex()
                    specs_batch.append((cat1, cat2, cat3, p_name, unit, norm_json, hex_hash))

                    if cat3 not in cat_hierarchy:
                        cat_hierarchy[cat3] = (cat1, cat2)

                # 提取属性值并按属性名称归并与去重统计
                if parsed_d and isinstance(parsed_d, dict):
                    for k, v in parsed_d.items():
                        k_clean = str(k).strip()
                        v_clean = str(v).strip()
                        if k_clean and v_clean:
                            # 长度适中属性值收录（过滤超长文章描述）
                            if len(v_clean) <= 100:
                                global_attr_values[k_clean][v_clean] += 1
                                cat_attr_values[cat3][k_clean][v_clean] += 1

                # 批量写入 category_specs
                if len(specs_batch) >= 50000:
                    cur.executemany(
                        "INSERT INTO category_specs (level1, level2, level3, product_name, unit, spec_json, spec_hash) "
                        "VALUES (?, ?, ?, ?, ?, ?, ?)",
                        specs_batch
                    )
                    conn.commit()
                    specs_batch = []

                if total_processed % REPORT_INTERVAL == 0:
                    elapsed = time.time() - t_start
                    rate = total_processed / elapsed if elapsed > 0 else 0
                    print(
                        f"  [进度: {total_processed // 1000000}M 行] 速率: {rate:,.0f} 行/s "
                        f"| 属性名称数: {len(global_attr_values)} | 唯一规格数: {unique_specs_count:,}"
                    )

        print(f"    -> 完成此文件读取: {file_rows:,} 行")

    if specs_batch:
        cur.executemany(
            "INSERT INTO category_specs (level1, level2, level3, product_name, unit, spec_json, spec_hash) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            specs_batch
        )
        conn.commit()
        specs_batch = []

    t_read = time.time()
    print(f"\n[*] 全部文件扫描完成！耗时: {(t_read - t_start) / 60:.2f} 分钟，共处理 {total_processed:,} 行")

    # ---------------- 阶段二：持久化全局属性名称分类去重库 ----------------
    print("[*] 正在写入全局属性名称去重表 (global_attribute_values)...")
    global_rows = []
    total_deduped_values = 0
    catalog_export = {}

    for attr_name, vals_counter in global_attr_values.items():
        # 提取单位
        u_match = re.search(r'[(（]([A-Za-z0-9/%²³\-_]+)[)）]', attr_name)
        unit = u_match.group(1) if u_match else ""

        catalog_export[attr_name] = {
            "unit": unit,
            "unique_values_count": len(vals_counter),
            "values": [val for val, _ in vals_counter.most_common()]
        }
        total_deduped_values += len(vals_counter)

        for val, freq in vals_counter.items():
            global_rows.append((attr_name, val, freq, unit))
            if len(global_rows) >= 50000:
                cur.executemany(
                    "INSERT OR REPLACE INTO global_attribute_values (attr_name, attr_value, frequency, unit) VALUES (?, ?, ?, ?)",
                    global_rows
                )
                conn.commit()
                global_rows = []

    if global_rows:
        cur.executemany(
            "INSERT OR REPLACE INTO global_attribute_values (attr_name, attr_value, frequency, unit) VALUES (?, ?, ?, ?)",
            global_rows
        )
        conn.commit()
    print(f"    -> 成功归集标准属性名称: {len(global_attr_values)} 类，去重后属性值总量: {total_deduped_values:,} 项")

    # ---------------- 阶段三：持久化品类属性去重库与模型表 ----------------
    print("[*] 正在写入品类属性模型与品类去重表...")
    cat_attr_rows = []
    schema_rows = []

    for cat3, (cat1, cat2) in cat_hierarchy.items():
        attr_dict = cat_attr_values.get(cat3, {})
        attr_schema_list = []
        sample_enums_dict = {}

        for attr_name, vals_cnt in sorted(attr_dict.items(), key=lambda x: sum(x[1].values()), reverse=True):
            u_match = re.search(r'[(（]([A-Za-z0-9/%²³\-_]+)[)）]', attr_name)
            unit = u_match.group(1) if u_match else ""
            top_vals = [v for v, _ in vals_cnt.most_common(20)]
            sample_enums_dict[attr_name] = top_vals

            attr_schema_list.append({
                "name": attr_name,
                "freq": sum(vals_cnt.values()),
                "unique_values_count": len(vals_cnt),
                "unit": unit,
                "common_values": top_vals[:8]
            })

            for val, freq in vals_cnt.items():
                cat_attr_rows.append((cat3, attr_name, val, freq))
                if len(cat_attr_rows) >= 50000:
                    cur.executemany(
                        "INSERT OR REPLACE INTO category_attr_values (level3, attr_name, attr_value, frequency) VALUES (?, ?, ?, ?)",
                        cat_attr_rows
                    )
                    conn.commit()
                    cat_attr_rows = []

        schema_rows.append((
            cat3, cat1, cat2,
            len(attr_dict),
            json.dumps(attr_schema_list, ensure_ascii=False),
            json.dumps(sample_enums_dict, ensure_ascii=False)
        ))

    if cat_attr_rows:
        cur.executemany(
            "INSERT OR REPLACE INTO category_attr_values (level3, attr_name, attr_value, frequency) VALUES (?, ?, ?, ?)",
            cat_attr_rows
        )
        conn.commit()

    if schema_rows:
        cur.executemany(
            "INSERT OR REPLACE INTO category_attribute_schema (level3, level1, level2, total_specs, attributes_json, sample_enums_json) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            schema_rows
        )
        conn.commit()
    print(f"    -> 成功构建 {len(schema_rows)} 个三级品类的标准属性体系")

    # ---------------- 阶段四：产品名称索引表 ----------------
    print("[*] 正在写入产品名称索引表...")
    prod_rows = []
    for p_name, cat_counter in product_category_votes.items():
        if not p_name:
            continue
        tot = sum(cat_counter.values())
        (b_cat1, b_cat2, b_cat3), top_votes = cat_counter.most_common(1)[0]
        conf = round(top_votes / tot, 4)
        prod_rows.append((p_name, b_cat1, b_cat2, b_cat3, tot, conf))

        if len(prod_rows) >= 50000:
            cur.executemany(
                "INSERT OR REPLACE INTO product_category_index (product_name, level1, level2, level3, sample_count, confidence) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                prod_rows
            )
            conn.commit()
            prod_rows = []

    if prod_rows:
        cur.executemany(
            "INSERT OR REPLACE INTO product_category_index (product_name, level1, level2, level3, sample_count, confidence) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            prod_rows
        )
        conn.commit()
    print(f"    -> 沉淀产品词条: {len(product_category_votes):,} 条")

    # ---------------- 阶段五：构建高性能索引 ----------------
    print("[*] 正在构建 B-Tree 索引...")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_global_attr_name ON global_attribute_values(attr_name);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_cat_attr_name ON category_attr_values(level3, attr_name);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_specs_level3 ON category_specs(level3);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_specs_name ON category_specs(product_name);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_specs_hash ON category_specs(spec_hash);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_prod_cat3 ON product_category_index(level3);")
    conn.commit()

    cur.execute("PRAGMA optimize;")
    conn.commit()
    conn.close()

    # ---------------- 阶段六：导出 JSON 与 Excel 交付件 ----------------
    print(f"[*] 导出轻量 JSON 字典: {OUTPUT_JSON} ...")
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(catalog_export, f, ensure_ascii=False, indent=2)

    print(f"[*] 导出 Excel 属性字典: {OUTPUT_XLSX} ...")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    NAVY_FILL = PatternFill(start_color="0B192C", end_color="0B192C", fill_type="solid")
    ICE_FILL = PatternFill(start_color="1E3E62", end_color="1E3E62", fill_type="solid")
    FONT_TH = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    FONT_BODY = Font(name="微软雅黑", size=10, color="141413")
    THIN = Border(
        left=Side(style='thin', color='D0D7DE'),
        right=Side(style='thin', color='D0D7DE'),
        top=Side(style='thin', color='D0D7DE'),
        bottom=Side(style='thin', color='D0D7DE')
    )

    # Sheet 1: 按属性名称分类汇总
    ws1 = wb.create_sheet("按属性名称分类汇总")
    ws1.views.sheetView[0].showGridLines = True
    ws1.cell(1, 1, "广材官方属性名称分类与取值字典 (去重汇总)").font = Font(name="微软雅黑", size=14, bold=True, color="0B192C")
    
    headers1 = ["序号", "属性名称", "标准单位", "去重属性值总数", "涉及频次", "高频代表取值 (Top 5 样例)"]
    for c, h in enumerate(headers1, 1):
        cell = ws1.cell(3, c, h)
        cell.fill = NAVY_FILL
        cell.font = FONT_TH
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN

    sorted_attrs = sorted(global_attr_values.keys(), key=lambda k: len(global_attr_values[k]), reverse=True)
    for idx, a_name in enumerate(sorted_attrs, 1):
        r = 3 + idx
        vals = global_attr_values[a_name]
        top5_str = "、".join([v for v, _ in vals.most_common(5)])
        u_match = re.search(r'[(（]([A-Za-z0-9/%²³\-_]+)[)）]', a_name)
        unit = u_match.group(1) if u_match else "—"

        row_vals = [idx, a_name, unit, len(vals), sum(vals.values()), top5_str]
        for c, v in enumerate(row_vals, 1):
            cell = ws1.cell(r, c, v)
            cell.font = FONT_BODY
            cell.border = THIN
            cell.alignment = Alignment(horizontal="center" if c in [1, 3, 4, 5] else "left", vertical="center")

    # Sheet 2: 属性值明细库 (Top 20000 高频去重属性值)
    ws2 = wb.create_sheet("属性值明细去重库")
    ws2.views.sheetView[0].showGridLines = True
    ws2.cell(1, 1, "广材标准属性值去重明细表 (按属性名称归类)").font = Font(name="微软雅黑", size=14, bold=True, color="0B192C")

    headers2 = ["序号", "归属属性名称", "去重属性值", "全局出现频次", "标准单位"]
    for c, h in enumerate(headers2, 1):
        cell = ws2.cell(3, c, h)
        cell.fill = ICE_FILL
        cell.font = FONT_TH
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN

    detail_row = 4
    item_seq = 1
    for a_name in sorted_attrs:
        u_match = re.search(r'[(（]([A-Za-z0-9/%²³\-_]+)[)）]', a_name)
        unit = u_match.group(1) if u_match else "—"
        for val, freq in global_attr_values[a_name].most_common(200):  # 每个属性保留前200高频
            for c, v in enumerate([item_seq, a_name, val, freq, unit], 1):
                cell = ws2.cell(detail_row, c, v)
                cell.font = FONT_BODY
                cell.border = THIN
                cell.alignment = Alignment(horizontal="center" if c in [1, 4, 5] else "left", vertical="center")
            detail_row += 1
            item_seq += 1
            if detail_row > 100000:
                break

    # 列宽自适应
    for ws in [ws1, ws2]:
        for col in ws.columns:
            cl = get_column_letter(col[0].column)
            mlen = max(sum(2 if ord(ch) > 127 else 1 for ch in str(cell.value or '')) for cell in col)
            ws.column_dimensions[cl].width = max(12, min(mlen + 3, 50))

    wb.save(OUTPUT_XLSX)
    print(f"    -> 成功生成 Excel 字典: {OUTPUT_XLSX}")

    db_mb = os.path.getsize(OUTPUT_DB) / (1024 * 1024)
    t_tot = time.time() - t_start
    print(f"\n🎉 [全量整合与属性去重顺利完成] 总耗时: {t_tot / 60:.2f} 分钟！")
    print(f"    - 输入大库总行数: {total_processed:,} 行")
    print(f"    - 提炼标准属性名称分类: {len(global_attr_values)} 种")
    print(f"    - 提炼去重属性值总量: {total_deduped_values:,} 项")
    print(f"    - 核心知识库 DB: {OUTPUT_DB} ({db_mb:.2f} MB)")
    print(f"    - Excel 字典交付物: {OUTPUT_XLSX}")

if __name__ == "__main__":
    run_integration()

