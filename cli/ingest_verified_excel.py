# -*- coding: utf-8 -*-
"""
ingest_verified_excel.py - 把人工复核过的交付件回流成词表数据（飞轮闭环）

用法：
    python3 cli/ingest_verified_excel.py deliveries/已审核目录          # 目录或单个文件
    python3 cli/ingest_verified_excel.py --dry-run  deliveries/已审核   # 只看会写什么

只做三件事：
  1. 读交付件「名称规格 / 类别编码 / 供应商名称」三列；
  2. 按 名称归一化（name_norm）折叠，把 (品名 → 类别) 追加进 category_lexicon.db 的 lex 表：
     · 已有同 (品名, 类别) → companies += 1（同一供应商只加一次，避免自己刷自己）
     · 新增组合 → 插入一条 companies=1
  3. 每条都写进 lex_feedback 表留痕（来源文件、供应商、时间），可追溯、可回滚

不碰原始快照（lex_meta.source_file 指向的那份 CSV），词表只增不改。
"""
import argparse
import glob
import os
import sqlite3
import sys
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

import openpyxl

from name_norm import normalize_material_name
from category_attribute_service import _load_catalog, _CODE_MAP

DEFAULT_DB = os.path.join(BASE_DIR, 'category_lexicon.db')


def _rows_of(path: str):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            hdr = [str(c.value or '') for c in ws[1]]
            if not any(h in hdr for h in ('名称规格', '材料名称', '品名')):
                continue
            def col(*names):
                for n in names:
                    if n in hdr:
                        return hdr.index(n)
                return -1
            c_name = col('名称规格', '材料名称', '品名')
            c_code = col('类别编码', '分类编码')
            c_sup = col('供应商名称')
            if c_name < 0 or c_code < 0:
                continue
            for r in ws.iter_rows(min_row=2, values_only=True):
                if not r or len(r) <= max(c_name, c_code):
                    continue
                yield (str(r[c_name] or '').strip(), str(r[c_code] or '').strip(),
                       str(r[c_sup] or '').strip() if c_sup >= 0 else '')
    finally:
        wb.close()


def ingest(target: str, db_path: str = DEFAULT_DB, dry_run: bool = False) -> dict:
    if not os.path.exists(db_path):
        raise SystemExit(f'词表库不存在: {db_path}（先跑 cli/build_category_lexicon.py）')
    _load_catalog()
    files = [target] if os.path.isfile(target) else sorted(glob.glob(os.path.join(target, '*.xlsx')))
    files = [f for f in files if not os.path.basename(f).startswith('~$')]

    con = sqlite3.connect(db_path)
    con.execute("""CREATE TABLE IF NOT EXISTS lex_feedback(
        nk TEXT, t1 TEXT, t2 TEXT, t3 TEXT, code TEXT, supplier TEXT,
        source_file TEXT, created_at TEXT)""")
    stats = {'files': 0, 'rows': 0, 'new': 0, 'bumped': 0, 'skipped': 0}
    for f in files:
        stats['files'] += 1
        for name, code, supplier in _rows_of(f):
            stats['rows'] += 1
            nk = normalize_material_name(name)
            norm_code = code.replace('-', '')
            hit = next((v for k, v in _CODE_MAP.items() if k.replace('-', '') == norm_code), None)
            if not nk or not hit:
                stats['skipped'] += 1
                continue
            t1, t2, t3, real_code = hit
            if dry_run:
                stats['new'] += 1
                continue
            existed = con.execute("SELECT companies FROM lex WHERE nk=? AND t3=?", (nk, t3)).fetchone()
            same_supplier = con.execute(
                "SELECT 1 FROM lex_feedback WHERE nk=? AND code=? AND supplier=?",
                (nk, real_code, supplier)).fetchone()
            if existed:
                if not same_supplier and supplier:
                    con.execute("UPDATE lex SET companies=companies+1 WHERE nk=? AND t3=?", (nk, t3))
                    stats['bumped'] += 1
            else:
                con.execute("INSERT INTO lex(nk,t1,t2,t3,rows,companies,sample) VALUES(?,?,?,?,?,?,?)",
                            (nk, t1, t2, t3, 1, 1, name))
                stats['new'] += 1
            con.execute("""INSERT INTO lex_feedback(nk,t1,t2,t3,code,supplier,source_file,created_at)
                           VALUES(?,?,?,?,?,?,?,?)""",
                        (nk, t1, t2, t3, real_code, supplier, os.path.basename(f),
                         datetime.now().isoformat(timespec='seconds')))
    if not dry_run:
        con.commit()
    con.close()
    return stats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('target', help='已审核交付件目录或单个 xlsx')
    ap.add_argument('--db', default=DEFAULT_DB)
    ap.add_argument('--dry-run', action='store_true')
    a = ap.parse_args()
    st = ingest(a.target, a.db, a.dry_run)
    tag = '（dry-run，未写库）' if a.dry_run else ''
    print(f"扫描 {st['files']} 个文件 / {st['rows']} 行 → 新增 {st['new']} 条、"
          f"企业支持度+1 {st['bumped']} 条、跳过(无品名或无合法编码) {st['skipped']} 条 {tag}")


if __name__ == '__main__':
    main()
