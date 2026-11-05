# -*- coding: utf-8 -*-
"""
build_category_lexicon.py - 把广材后台导出的「产品名称 → 工料机三级品类」快照编译成查表词库

去重只做"事实折叠"，不丢信息：
  1. 无分类的行：丢弃（176 行）
  2. 同一 (归一化品名, 三级类别) 的重复出现：折叠成一条，但保留 行数 / **去重企业数** 作为支持度
     （支持度按企业算，不按行算——电缆大户一家能刷 4 万行）
  3. 同一品名对多个类别：**不折叠**，原样存成分布（如"电线"对 9 个类别），由在线判定用主业先验去挑

产物：
  category_lexicon.db   SQLite：lex / lex_best / lex_meta
用法：
  python3 cli/build_category_lexicon.py --csv <导出文件> [--db category_lexicon.db]
"""
import argparse
import hashlib
import os
import sqlite3
import sys

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, BASE_DIR)

import pandas as pd

from name_norm import normalize_material_name

REQUIRED_COLS = ['企业名称', '产品ID', '产品名称', '一级工料机品类', '二级工料机品类', '三级工料机品类']


def build(csv_path: str, db_path: str) -> dict:
    raw = pd.read_csv(csv_path, encoding='utf-8-sig', dtype=str)
    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        raise SystemExit(f"CSV 缺少列: {missing}")

    n_raw = len(raw)
    df = raw.dropna(subset=['三级工料机品类']).copy()
    n_blank = n_raw - len(df)
    n_dup_product = int(df.duplicated(subset=['产品ID', '三级工料机品类']).sum())
    df = df.drop_duplicates(subset=['产品ID', '三级工料机品类'])

    df['nk'] = df['产品名称'].map(normalize_material_name)
    df = df[df['nk'] != '']
    for c in ('一级工料机品类', '二级工料机品类', '三级工料机品类'):
        df[c] = df[c].str.strip()

    facts = (df.groupby(['nk', '一级工料机品类', '二级工料机品类', '三级工料机品类'])
               .agg(rows=('产品ID', 'size'), companies=('企业名称', 'nunique'),
                    sample=('产品名称', 'first'))
               .reset_index())

    if os.path.exists(db_path):
        os.remove(db_path)
    con = sqlite3.connect(db_path)
    con.execute("""CREATE TABLE lex(
        nk TEXT NOT NULL, t1 TEXT, t2 TEXT, t3 TEXT,
        rows INTEGER, companies INTEGER, sample TEXT,
        PRIMARY KEY(nk, t3))""")
    con.executemany(
        "INSERT INTO lex(nk,t1,t2,t3,rows,companies,sample) VALUES(?,?,?,?,?,?,?)",
        [(r['nk'], r['一级工料机品类'], r['二级工料机品类'], r['三级工料机品类'],
          int(r['rows']), int(r['companies']), r['sample']) for _, r in facts.iterrows()])
    # 每个品名的最优类别 + 歧义度（在线判定先看这个视图）
    con.execute("""CREATE VIEW lex_best AS
        SELECT nk, COUNT(*) AS n_codes,
               (SELECT t3 FROM lex b WHERE b.nk=a.nk ORDER BY companies DESC, rows DESC LIMIT 1) AS t3_top,
               (SELECT t1 FROM lex b WHERE b.nk=a.nk ORDER BY companies DESC, rows DESC LIMIT 1) AS t1_top,
               (SELECT t2 FROM lex b WHERE b.nk=a.nk ORDER BY companies DESC, rows DESC LIMIT 1) AS t2_top,
               MAX(companies) AS companies_top, MAX(rows) AS rows_top
        FROM lex a GROUP BY nk""")
    sha = hashlib.sha256(open(csv_path, 'rb').read()).hexdigest()
    con.execute("CREATE TABLE lex_meta(k TEXT PRIMARY KEY, v TEXT)")
    con.executemany("INSERT INTO lex_meta(k,v) VALUES(?,?)", [
        ('source_file', os.path.basename(csv_path)), ('sha256', sha),
        ('raw_rows', str(n_raw)), ('rows_used', str(len(df))),
        ('facts', str(len(facts))), ('names', str(df['nk'].nunique())),
    ])
    con.commit()
    stats = {
        'raw': n_raw, 'blank_dropped': n_blank, 'dup_product_dropped': n_dup_product,
        'rows_used': len(df), 'facts': len(facts), 'names': int(df['nk'].nunique()),
        'ambiguous_names': int((facts.groupby('nk').size() > 1).sum()),
        'multi_company_names': int((facts.groupby('nk')['companies'].max() >= 2).sum()),
    }
    con.close()
    return stats


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--csv', required=True)
    ap.add_argument('--db', default=os.path.join(BASE_DIR, 'category_lexicon.db'))
    args = ap.parse_args()
    st = build(args.csv, args.db)
    print(f"原始 {st['raw']} 行 → 丢弃无分类 {st['blank_dropped']} / 重复产品 {st['dup_product_dropped']} "
          f"→ 用 {st['rows_used']} 行")
    print(f"折叠成 {st['facts']} 条(品名,类别)事实 | {st['names']} 个品名 "
          f"(压缩 {st['rows_used']/max(st['facts'],1):.1f}:1)")
    print(f"其中歧义品名 {st['ambiguous_names']} 个；有 ≥2 家企业印证的 {st['multi_company_names']} 个")
    print(f"已写入 {args.db}")


if __name__ == '__main__':
    main()
