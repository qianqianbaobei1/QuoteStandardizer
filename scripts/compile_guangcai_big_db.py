# -*- coding: utf-8 -*-
"""
scripts/compile_guangcai_big_db.py
广材后台数据库（11GB / 34,198,511行）高性能流式预编译与去重脚本

业务目标：
1. "同一个分类下的同一个规格下的只保留一个就够用了"（极高压缩比去重，消除99%+冗余）；
2. 提取并建立：
   - 表 category_specs: (level1, level2, level3, product_name, unit, spec_json, spec_hash)
   - 表 product_category_index: 产品名称 -> (level1, level2, level3, 频次, 置信度) 用于图片/PDF名称快速反查标准分类
   - 表 category_attribute_schema: 三级品类 -> 官方标准属性Schema（属性名、计量单位、是否必填、枚举值池）
   - 表 attribute_enums: (level3, attr_key, attr_value) 用于属性值归一化与校验
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

CSV_PATH = "/Users/qianqianawodebaobei/Downloads/result_广材后台数据库_db_op_gldjc_20261105172104.csv"
OUTPUT_DB_PATH = "guangcai_knowledge_base.db"
TOTAL_ESTIMATED_ROWS = 34198511
REPORT_INTERVAL = 1000000  # 每 100 万行打印一次进度
BATCH_INSERT_SIZE = 50000  # SQLite 批量提交批次

def run_etl():
    if not os.path.exists(CSV_PATH):
        print(f"[Error] 找不到输入文件: {CSV_PATH}")
        sys.exit(1)

    print(f"=== [广材全量大库流式编译] 启动 ===")
    print(f"输入文件: {CSV_PATH} ({os.path.getsize(CSV_PATH) / (1024**3):.2f} GB)")
    print(f"输出数据库: {OUTPUT_DB_PATH}")
    print(f"去重规则: 相同 (三级分类 + 规范化规格JSON) 仅保留唯一样本")

    # 如果旧 DB 存在则先移除重新构建
    if os.path.exists(OUTPUT_DB_PATH):
        try:
            os.remove(OUTPUT_DB_PATH)
            print("已清理旧的数据库文件。")
        except Exception as e:
            print(f"清理旧数据库失败: {e}")

    conn = sqlite3.connect(OUTPUT_DB_PATH)
    cur = conn.cursor()

    # SQLite 高性能写入参数调优
    cur.execute("PRAGMA synchronous = OFF;")
    cur.execute("PRAGMA journal_mode = WAL;")
    cur.execute("PRAGMA cache_size = -128000;")  # 128MB 缓存
    cur.execute("PRAGMA temp_store = MEMORY;")

    # 1. 规格唯一样本表 (同一个分类下的同一个规格只保留一个)
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

    # 2. 产品名称 -> 标准分类反查索引表
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

    # 3. 三级品类标准属性体系表
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

    # 4. 属性高频合法枚举值表
    cur.execute("""
        CREATE TABLE attribute_enums (
            level3 TEXT,
            attr_key TEXT,
            attr_value TEXT,
            freq INTEGER,
            PRIMARY KEY (level3, attr_key, attr_value)
        )
    """)
    conn.commit()

    # 内存哈希池：存放 (level3_hash + spec_hash)，使用 16 字节 bytes 极致节省内存
    seen_spec_hashes = set()
    # 统计产品名称到分类的投票分布: p_name -> { (cat1, cat2, cat3): count }
    product_category_votes = defaultdict(Counter)
    # 统计分类下的属性键及枚举值: cat3 -> { (cat1, cat2): True }, cat3 -> { attr_key: Counter({val: count}) }
    cat_hierarchy = {}
    cat_attr_keys = defaultdict(Counter)
    cat_attr_enums = defaultdict(lambda: defaultdict(Counter))

    specs_batch = []
    total_processed = 0
    unique_specs_count = 0
    t_start = time.time()
    t_last = t_start

    print("[*] 开始流式扫描 3420 万行数据...")

    with open(CSV_PATH, "r", encoding="utf-8-sig", errors="replace", buffering=16 * 1024 * 1024) as f:
        reader = csv.reader(f)
        header = next(reader)  # 跳过表头

        for row in reader:
            total_processed += 1
            if len(row) < 12:
                continue

            # 字段抽取
            p_name = row[5].strip()
            unit = row[7].strip()
            spec_str = row[8].strip()
            cat1 = row[9].strip()
            cat2 = row[10].strip()
            cat3 = row[11].strip()

            if not cat3 or not spec_str:
                continue

            # 记录产品名称到三级品类的投票
            if p_name:
                product_category_votes[p_name][(cat1, cat2, cat3)] += 1

            # 规范化属性 JSON（按 Key 字母排序，消除无序字典带来的伪重复）
            parsed_d = None
            if spec_str.startswith("{") and spec_str.endswith("}"):
                try:
                    parsed_d = json.loads(spec_str)
                    norm_json = json.dumps(parsed_d, sort_keys=True, ensure_ascii=False)
                except Exception:
                    norm_json = spec_str
            else:
                norm_json = spec_str

            # 计算分类+规格的唯一指纹 (MD5 16字节)
            h = hashlib.md5((cat3 + "\n" + norm_json).encode("utf-8")).digest()
            if h not in seen_spec_hashes:
                seen_spec_hashes.add(h)
                unique_specs_count += 1
                hex_hash = h.hex()
                specs_batch.append((cat1, cat2, cat3, p_name, unit, norm_json, hex_hash))

                # 记录分类层次
                if cat3 not in cat_hierarchy:
                    cat_hierarchy[cat3] = (cat1, cat2)

                # 收集该分类下的属性与枚举分布
                if parsed_d and isinstance(parsed_d, dict):
                    for k, v in parsed_d.items():
                        k_clean = str(k).strip()
                        v_clean = str(v).strip()
                        if k_clean:
                            cat_attr_keys[cat3][k_clean] += 1
                            if v_clean and len(v_clean) <= 60:
                                cat_attr_enums[cat3][k_clean][v_clean] += 1

            # 定期批量写入 SQLite
            if len(specs_batch) >= BATCH_INSERT_SIZE:
                cur.executemany(
                    "INSERT INTO category_specs (level1, level2, level3, product_name, unit, spec_json, spec_hash) "
                    "VALUES (?, ?, ?, ?, ?, ?, ?)",
                    specs_batch,
                )
                conn.commit()
                specs_batch = []

            # 进度汇报
            if total_processed % REPORT_INTERVAL == 0:
                now = time.time()
                elapsed = now - t_start
                rate = total_processed / elapsed if elapsed > 0 else 0
                pct = (total_processed / TOTAL_ESTIMATED_ROWS) * 100
                rem_sec = (TOTAL_ESTIMATED_ROWS - total_processed) / rate if rate > 0 else 0
                rem_m = rem_sec / 60
                print(
                    f"  [{total_processed // 1000000}M / {TOTAL_ESTIMATED_ROWS // 1000000}M] {pct:.1f}% "
                    f"| 速率: {rate:,.0f} 行/s | 已保留唯一规格: {unique_specs_count:,} | 剩余预计: {rem_m:.1f} 分钟"
                )

        # 写入最后一批规格
        if specs_batch:
            cur.executemany(
                "INSERT INTO category_specs (level1, level2, level3, product_name, unit, spec_json, spec_hash) "
                "VALUES (?, ?, ?, ?, ?, ?, ?)",
                specs_batch,
            )
            conn.commit()
            specs_batch = []

    t_pass1 = time.time()
    print(f"[*] 全表 3,420 万行扫描去重完成！耗时: {(t_pass1 - t_start) / 60:.2f} 分钟")
    print(f"    - 原文件行数: {total_processed:,}")
    print(f"    - 提炼唯一规格数: {unique_specs_count:,} (压缩率: {unique_specs_count / max(1, total_processed):.2%})")

    # ---------------- 阶段二：生成产品名称反查索引表 ----------------
    print("[*] 正在构建产品名称反查索引表 (product_category_index)...")
    prod_rows = []
    for p_name, cat_counter in product_category_votes.items():
        if not p_name:
            continue
        total_votes = sum(cat_counter.values())
        (best_cat1, best_cat2, best_cat3), top_votes = cat_counter.most_common(1)[0]
        confidence = round(top_votes / total_votes, 4)
        prod_rows.append((p_name, best_cat1, best_cat2, best_cat3, total_votes, confidence))

        if len(prod_rows) >= 50000:
            cur.executemany(
                "INSERT OR REPLACE INTO product_category_index (product_name, level1, level2, level3, sample_count, confidence) "
                "VALUES (?, ?, ?, ?, ?, ?)",
                prod_rows,
            )
            conn.commit()
            prod_rows = []

    if prod_rows:
        cur.executemany(
            "INSERT OR REPLACE INTO product_category_index (product_name, level1, level2, level3, sample_count, confidence) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            prod_rows,
        )
        conn.commit()
        prod_rows = []
    print(f"    -> 成功沉淀唯一标准产品词条: {len(product_category_votes):,} 条")

    # ---------------- 阶段三：生成三级品类标准属性体系表与枚举表 ----------------
    print("[*] 正在构建三级品类标准属性模型 (category_attribute_schema & attribute_enums)...")
    schema_rows = []
    enums_rows = []

    for cat3, (cat1, cat2) in cat_hierarchy.items():
        keys_counter = cat_attr_keys.get(cat3, Counter())
        tot_specs = sum(1 for h in seen_spec_hashes)  # approx
        # 按出现频次倒序排列标准属性
        attr_schema_list = []
        sample_enums_dict = {}

        for k, cnt in keys_counter.most_common(50):
            # 提取单位提示：如果属性名含括号单位如 (mm)，标记 unit
            unit_match = re.search(r'[(（]([A-Za-z0-9/%²³\-_]+)[)）]', k)
            unit_hint = unit_match.group(1) if unit_match else ""
            
            top_enums = [val for val, _ in cat_attr_enums[cat3][k].most_common(15)]
            sample_enums_dict[k] = top_enums

            attr_schema_list.append({
                "name": k,
                "freq": cnt,
                "unit": unit_hint,
                "common_values": top_enums[:5]
            })

            # 枚举表写入
            for val, v_cnt in cat_attr_enums[cat3][k].most_common(30):
                enums_rows.append((cat3, k, val, v_cnt))

        schema_rows.append((
            cat3, cat1, cat2,
            len(keys_counter),
            json.dumps(attr_schema_list, ensure_ascii=False),
            json.dumps(sample_enums_dict, ensure_ascii=False)
        ))

        if len(enums_rows) >= 50000:
            cur.executemany(
                "INSERT OR REPLACE INTO attribute_enums (level3, attr_key, attr_value, freq) VALUES (?, ?, ?, ?)",
                enums_rows
            )
            conn.commit()
            enums_rows = []

    if schema_rows:
        cur.executemany(
            "INSERT OR REPLACE INTO category_attribute_schema (level3, level1, level2, total_specs, attributes_json, sample_enums_json) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            schema_rows
        )
        conn.commit()

    if enums_rows:
        cur.executemany(
            "INSERT OR REPLACE INTO attribute_enums (level3, attr_key, attr_value, freq) VALUES (?, ?, ?, ?)",
            enums_rows
        )
        conn.commit()
    print(f"    -> 成功构建标准分类属性模型: {len(schema_rows):,} 个三级品类")

    # ---------------- 阶段四：创建高性能索引 ----------------
    print("[*] 正在创建高性能索引 (B-Tree Indexes)...")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_specs_level3 ON category_specs(level3);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_specs_product_name ON category_specs(product_name);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_specs_hash ON category_specs(spec_hash);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_prod_idx_cat3 ON product_category_index(level3);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_attr_enums_key ON attribute_enums(level3, attr_key);")
    conn.commit()

    # 数据库瘦身与空间紧缩
    print("[*] 执行 VACUUM 与空间紧缩...")
    cur.execute("PRAGMA optimize;")
    conn.commit()
    conn.close()

    db_size_mb = os.path.getsize(OUTPUT_DB_PATH) / (1024 * 1024)
    t_total = time.time() - t_start
    print(f"\n🎉 [广材知识库全量编译完成] 总耗时: {t_total / 60:.2f} 分钟！")
    print(f"    - 原文件: {CSV_PATH} (11 GB)")
    print(f"    - 目标库: {OUTPUT_DB_PATH} ({db_size_mb:.2f} MB)")
    print(f"    - 存储总规格数: {unique_specs_count:,}")
    print(f"    - 产品词典容量: {len(product_category_votes):,}")
    print(f"    - 三级品类模型: {len(schema_rows):,}")

if __name__ == "__main__":
    run_etl()

