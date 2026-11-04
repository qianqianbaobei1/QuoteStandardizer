# -*- coding: utf-8 -*-
"""
一次性/批次专用脚本的公共路径。路径一律来自环境变量（或 .env），不写个人目录。

  QS_DELIVERIES_DIR      交付目录，默认 config.DEFAULT_OUTPUT_DIR (OUTPUT_DIR 或 ./deliveries)
  QS_BACKUP_DIR          历史交付备份目录，默认 <项目根>/deliveries_backup_20261103_1828
  QS_SOURCE_ROOT         「02_四位经办人分类完整档案」源目录
  QS_ZHOUWEI_1103_DIR    「周伟E化11.03」解包目录，默认 <QS_BACKUP_DIR>/周伟E化11.03_unpacked/周伟E化11.03
  QS_WUHAN_SOURCE_XLSX   武汉第二电线电缆《2026电线出厂全系列价格表.xlsx》原件
  QS_ZIP_DIR             线上导出压缩包所在目录，默认 ~/Desktop
"""
import os
import sys

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if REPO_ROOT not in sys.path:
    sys.path.insert(0, REPO_ROOT)

from config import DEFAULT_OUTPUT_DIR  # noqa: E402  (config 负责加载 .env)

DELIVERIES_DIR = os.getenv("QS_DELIVERIES_DIR") or DEFAULT_OUTPUT_DIR
BACKUP_DIR = os.getenv("QS_BACKUP_DIR") or os.path.join(REPO_ROOT, "deliveries_backup_20261103_1828")
SOURCE_ROOT = os.getenv("QS_SOURCE_ROOT", "")
ZHOUWEI_1103_DIR = os.getenv("QS_ZHOUWEI_1103_DIR") or os.path.join(BACKUP_DIR, "周伟E化11.03_unpacked", "周伟E化11.03")
ZIP_DIR = os.getenv("QS_ZIP_DIR") or os.path.expanduser("~/Desktop")


def require_env(name: str) -> str:
    """读取必填路径；未配置时给出用法并退出，而不是带着空路径继续跑"""
    v = os.getenv(name, "").strip()
    if not v:
        print(f"缺少环境变量 {name}（可写在项目根目录 .env 中），见 scripts/adhoc/_adhoc_paths.py")
        raise SystemExit(1)
    return v


def wuhan_source_xlsx() -> str:
    return require_env("QS_WUHAN_SOURCE_XLSX")
