# -*- coding: utf-8 -*-
"""
cas_cache.py - 内容寻址缓存（CAS）给所有昂贵的外部调用做本地缓存

作用：同一份输入的第 2 次以后调用直接本地返回 —— 断点续跑、改下游渲染规则后秒级回归、
      以及同构行去重后的重复请求都不再花钱。
键 = SHA256(kind + model + prompt/输入内容)，值 = 原始响应文本（JSON 字符串）。

默认开启（CAS_CACHE=0 可关）。DB: llm_cas_cache.db（已在 .gitignore 里）
"""
import hashlib
import os
import sqlite3
import threading

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.getenv('CAS_CACHE_DB', os.path.join(BASE_DIR, 'llm_cas_cache.db'))
ENABLED = os.getenv('CAS_CACHE', '1') not in ('0', 'false', 'False')

_LOCK = threading.Lock()
_INITIALIZED = False
_STATS = {'hit': 0, 'miss': 0, 'put': 0}


def _ensure():
    global _INITIALIZED
    if _INITIALIZED:
        return
    with _LOCK:
        if _INITIALIZED:
            return
        con = sqlite3.connect(DB_PATH)
        con.execute("""CREATE TABLE IF NOT EXISTS cas(
            k TEXT PRIMARY KEY, kind TEXT, model TEXT, payload TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
        con.commit()
        con.close()
        _INITIALIZED = True


def make_key(kind: str, model: str, payload: str) -> str:
    h = hashlib.sha256()
    h.update(kind.encode())
    h.update(b'\x00')
    h.update((model or '').encode())
    h.update(b'\x00')
    h.update((payload or '').encode())
    return h.hexdigest()


def get(kind: str, model: str, payload: str):
    if not ENABLED:
        return None
    _ensure()
    k = make_key(kind, model, payload)
    try:
        con = sqlite3.connect(DB_PATH)
        row = con.execute("SELECT payload FROM cas WHERE k=?", (k,)).fetchone()
        con.close()
    except Exception:
        return None
    if row:
        _STATS['hit'] += 1
        return row[0]
    _STATS['miss'] += 1
    return None


def put(kind: str, model: str, payload: str, value: str):
    if not ENABLED or value is None:
        return
    _ensure()
    k = make_key(kind, model, payload)
    try:
        con = sqlite3.connect(DB_PATH)
        con.execute("INSERT OR REPLACE INTO cas(k,kind,model,payload) VALUES(?,?,?,?)",
                    (k, kind, model or '', value))
        con.commit()
        con.close()
        _STATS['put'] += 1
    except Exception:
        pass


def stats() -> dict:
    return dict(_STATS)
