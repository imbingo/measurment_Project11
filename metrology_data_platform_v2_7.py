#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
量测数据采集配置平台 V2.7 - 稳定化增强版

当前推荐入口：
    python .\metrology_data_platform_v2_7.py

核心能力：
- 管理生产编号、量测项、指标和模板；
- 采集 CSV、Excel、Image OCR 数据源；
- 支持局域网部署、定时采集、结果查询、导出和审计日志；
- 采集结果以“作废”替代物理删除，保留可追溯记录；
- Image OCR 结果保存原图路径、OCR 原文、ROI 和配置快照；
- 采集任务生命周期、角色权限、模板套用快照和配置完整性检查。

正式部署前请设置管理员密码环境变量 MDCP_ADMIN_PASSWORD。
UNC 路径必须由运行本程序的电脑或服务器账号具备读取权限。
"""

import csv
import glob
import hashlib
import html
import io
import json
import math
import os
import re
import secrets
import shutil
import sqlite3
import threading
import time
import traceback
import uuid
import zipfile
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeoutError
from datetime import datetime, timedelta, timezone
from http import cookies
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlencode, urlparse

APP_VERSION = "V2.7"
APP_TITLE = "量测数据采集配置平台 V2.7 - 稳定化增强版"
SCHEMA_VERSION = "V2.7.1"
DB_FILE = os.environ.get("MDCP_DB_FILE", "metrology_config_v1.db")
BACKUP_DIR = Path(os.environ.get("MDCP_BACKUP_DIR", "backup"))
HOST = os.environ.get("MDCP_HOST", "127.0.0.1")
PORT = int(os.environ.get("MDCP_PORT", "8023"))
SESSIONS = {}
SCHEDULER_STOP = threading.Event()
# Items whose real collection is currently running (possibly blocked on a slow/dead
# network path). Used to avoid piling up duplicate reads and exhausting the read pool.
INFLIGHT_LOCK = threading.Lock()
INFLIGHT_ITEMS = set()
APP_TZ = timezone(timedelta(hours=8))  # Asia/Shanghai / Asia/Singapore, UTC+8
READ_TIMEOUT_SECONDS = int(os.environ.get("MDCP_READ_TIMEOUT_SECONDS", "20"))
READ_RETRY_COUNT = int(os.environ.get("MDCP_READ_RETRY_COUNT", "3"))
READ_RETRY_INTERVAL_SECONDS = float(os.environ.get("MDCP_READ_RETRY_INTERVAL_SECONDS", "1.0"))
FILE_STABLE_WAIT_SECONDS = float(os.environ.get("MDCP_FILE_STABLE_WAIT_SECONDS", "0.4"))
READ_EXECUTOR = ThreadPoolExecutor(max_workers=int(os.environ.get("MDCP_READ_WORKERS", "4")))
DISPLAY_IP = os.environ.get("MDCP_DISPLAY_IP", "10.21.210.75")
TEMPLATE_CACHE_DIR = Path(os.environ.get("MDCP_TEMPLATE_CACHE_DIR", "template_upload_cache"))
DEMO_DATA_DIR = Path(os.environ.get("MDCP_DEMO_DATA_DIR", str(Path(__file__).resolve().parent / "demo_data")))
DEMO_CSV_FILENAME = "customer_acceptance_results.csv"
# Default UNC path shown in the new-item form. Kept as a module constant because an
# f-string expression cannot contain a backslash on Python < 3.12 (PEP 701).
DEFAULT_DATA_SOURCE_PATH_EXAMPLE = r"\\192.168.1.100\share\result.xlsx"
DEFAULT_IMAGE_SOURCE_PATH_EXAMPLE = r"\\192.168.1.100\share\image_inbox"
IMAGE_SUFFIXES = (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")
DEFAULT_IMAGE_PARSE_CONFIG_JSON = json.dumps({
    "file_pattern": "*",
    "collect_mode": "all_stable",
    "max_files_per_collect": 50,
    "route_by_image_production_code": True,
    "production_code_from_filename_regex": r"(?P<production_code>.+)_[^_.]+\.[^.]+$",
    "process_from_filename_regex": r"_(?P<process_step>[^_.]+)\.",
    "ocr": {
        "lang": "eng",
        "psm": 6,
        "scale": 2.0,
        "threshold": True
    },
    "metrics": {
        "Rx": {"roi": [0.05, 0.10, 0.30, 0.12], "regex": r"Rx\s*[:=]?\s*([-+]?\d+(?:\.\d+)?)"},
        "Ry": {"roi": [0.05, 0.24, 0.30, 0.12], "regex": r"Ry\s*[:=]?\s*([-+]?\d+(?:\.\d+)?)"},
        "Z": {"roi": [0.05, 0.38, 0.30, 0.12], "regex": r"Z\s*[:=]?\s*([-+]?\d+(?:\.\d+)?)"}
    }
}, ensure_ascii=False, indent=2)

DEMO_CSV_TEXT = """生产编号,工序,批次,设备,CD,Overlay_X,Overlay_Y,Rz
DEMO_WAFER_A01,PHOTO_CD,LOT-A-0726,CDSEM-01,49.8,0.4,-0.3,9.6
DEMO_WAFER_A01,ETCH_PROFILE,LOT-A-0726,PROFILER-02,51.2,0.7,-0.4,10.2
DEMO_WAFER_A01,FINAL_AOI,LOT-A-0726,AOI-03,52.6,2.1,-0.8,12.8
DEMO_WAFER_B02,PHOTO_CD,LOT-B-0726,CDSEM-01,48.9,-0.9,1.0,8.4
DEMO_WAFER_B02,ETCH_PROFILE,LOT-B-0726,PROFILER-02,54.4,1.8,-2.2,13.2
DEMO_WAFER_B02,FINAL_AOI,LOT-B-0726,AOI-03,56.3,3.4,-2.9,14.8
"""

DEMO_METRICS = [
    {"metric_name": "CD", "source_column": "CD", "unit": "nm", "target": 50.0, "lsl": 45.0, "usl": 55.0, "lcl": 48.5, "ucl": 51.5, "sort_order": 1},
    {"metric_name": "Overlay_X", "source_column": "Overlay_X", "unit": "nm", "target": 0.0, "lsl": -3.0, "usl": 3.0, "lcl": -1.5, "ucl": 1.5, "sort_order": 2},
    {"metric_name": "Overlay_Y", "source_column": "Overlay_Y", "unit": "nm", "target": 0.0, "lsl": -3.0, "usl": 3.0, "lcl": -1.5, "ucl": 1.5, "sort_order": 3},
    {"metric_name": "Rz", "source_column": "Rz", "unit": "nm", "target": 10.0, "lsl": 6.0, "usl": 14.0, "lcl": 8.0, "ucl": 12.0, "sort_order": 4},
]


# ==========================================================
# Utility
# ==========================================================

def now_str():
    return datetime.now(APP_TZ).strftime("%Y-%m-%d %H:%M:%S")


def h(value):
    return html.escape("" if value is None else str(value))


def hash_password(password: str) -> str:
    return hashlib.sha256(("mdcp_v1_salt_" + password).encode("utf-8")).hexdigest()


def row_hash(row: dict) -> str:
    raw = json.dumps(row, ensure_ascii=False, sort_keys=True)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


def parse_cookie(header):
    jar = cookies.SimpleCookie()
    if header:
        jar.load(header)
    return {k: morsel.value for k, morsel in jar.items()}


def redirect(location):
    return 302, {"Location": location}, b""


def safe_int(value, default=0):
    try:
        return int(value)
    except Exception:
        return default


def safe_float(value):
    if value is None:
        return None
    text = str(value).strip()
    if text == "":
        return None
    try:
        return float(text)
    except Exception:
        return None


VALID_ROLES = ("viewer", "engineer", "admin")


def user_role(user):
    # 旧 session 里没有 role 时按 admin 处理，避免升级瞬间把当前管理员锁在外面。
    role = (user or {}).get("role", "admin")
    return role if role in VALID_ROLES else "viewer"


def has_role(user, *roles):
    return bool(user) and user_role(user) in roles


def can_manage_config(user):
    return has_role(user, "admin", "engineer")


def can_collect(user):
    return has_role(user, "admin", "engineer")


def can_void_results(user):
    return has_role(user, "admin")


def can_manage_users(user):
    return has_role(user, "admin")


def can_manage_system(user):
    return has_role(user, "admin")


def require_permission(user, allowed, message="当前账号没有权限执行该操作"):
    if not allowed:
        raise PermissionError(message)


# ==========================================================
# Database
# ==========================================================

def get_conn():
    # timeout + WAL: reduces "database is locked" errors when scheduler and UI write concurrently.
    conn = sqlite3.connect(DB_FILE, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=30000")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def ensure_column(cur, table_name, column_name, column_sql):
    existing = {r["name"] for r in cur.execute(f"PRAGMA table_info({table_name})").fetchall()}
    if column_name not in existing:
        cur.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_sql}")
        return True
    return False


def sqlite_tables(conn):
    return {r["name"] for r in conn.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}


def schema_version_from_conn(conn):
    tables = sqlite_tables(conn)
    if "schema_meta" not in tables:
        return ""
    row = conn.execute("SELECT value FROM schema_meta WHERE key='schema_version'").fetchone()
    return row["value"] if row else ""


def database_needs_migration(db_file):
    db_path = Path(db_file)
    if not db_path.exists() or db_path.stat().st_size == 0:
        return False, ""
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    try:
        tables = sqlite_tables(conn)
        if not tables:
            return False, ""
        current_version = schema_version_from_conn(conn)
        required_tables = {
            "admin_user", "production_config", "measurement_item_config", "metric_config",
            "measurement_result", "collect_log", "collect_job", "audit_log",
            "template_config", "template_metric_config", "template_apply_log", "schema_meta"
        }
        if not required_tables.issubset(tables):
            return True, current_version or "unknown"
        required_columns = {
            "production_config": {"production_code", "production_name", "product_model", "process_version", "description", "status", "created_at", "updated_at"},
            "admin_user": {"role", "enabled", "updated_at"},
            "measurement_item_config": {"production_id", "item_name", "process_step", "execution_time_text", "equipment_name", "data_source_type", "data_source_path", "excel_sheet_name", "image_parse_config_json", "header_row_index", "csv_encoding", "delimiter", "production_code_column", "process_step_column", "scan_frequency_seconds", "enabled", "last_collect_time", "last_collect_status", "created_at", "updated_at"},
            "metric_config": {"item_id", "metric_name", "source_column", "unit", "data_type", "target", "lsl", "usl", "lcl", "ucl", "enabled", "sort_order", "created_at", "updated_at"},
            "measurement_result": {"production_id", "production_code", "item_id", "measurement_item_name", "process_step", "execution_time_text", "equipment_name", "metric_name", "metric_value_text", "metric_value_number", "unit", "target", "lsl", "usl", "lcl", "ucl", "result_status", "source_path", "source_row_hash", "source_metric_hash", "collect_time", "raw_row_json", "is_voided", "voided_by", "voided_at", "void_reason", "source_file_path", "source_file_mtime", "ocr_raw_text", "ocr_roi_json", "ocr_confidence", "ocr_config_json"},
            "collect_log": {"production_id", "production_code", "item_id", "measurement_item_name", "data_source_path", "status", "message", "matched_rows", "inserted_count", "skipped_count", "created_at"},
            "audit_log": {"username", "action", "object_type", "object_id", "detail", "ip_address", "created_at"},
            "template_config": {"template_name", "template_version", "data_source_type", "header_row_index", "delimiter", "encoding", "excel_sheet_name", "production_code_column", "process_step_column", "sample_fields_json", "description", "created_at", "updated_at"},
            "template_metric_config": {"template_id", "metric_name", "source_column", "data_type", "unit", "target", "lsl", "usl", "lcl", "ucl", "sort_order", "created_at", "updated_at"},
            "template_apply_log": {"template_id", "production_id", "production_code", "item_id", "applied_by", "applied_at", "template_name", "template_version", "snapshot_json", "detail"},
        }
        for table, columns in required_columns.items():
            existing = {r["name"] for r in conn.execute(f"PRAGMA table_info({table})").fetchall()} if table in tables else set()
            if not columns.issubset(existing):
                return True, current_version or "unknown"
        if current_version != SCHEMA_VERSION:
            return True, current_version or "unknown"
        return False, current_version
    finally:
        conn.close()


def backup_database_before_migration(db_file, from_version, to_version):
    db_path = Path(db_file)
    if not db_path.exists() or db_path.stat().st_size == 0:
        return ""
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now(APP_TZ).strftime("%Y%m%d_%H%M%S")
    safe_from = re.sub(r"[^A-Za-z0-9_.-]+", "_", from_version or "unknown")
    safe_to = re.sub(r"[^A-Za-z0-9_.-]+", "_", to_version or "unknown")
    backup_path = BACKUP_DIR / f"{db_path.stem}_{safe_from}_to_{safe_to}_{timestamp}.db"
    try:
        src = sqlite3.connect(str(db_path))
        dst = sqlite3.connect(str(backup_path))
        try:
            src.backup(dst)
        finally:
            dst.close()
            src.close()
    except Exception:
        shutil.copy2(db_path, backup_path)
    return str(backup_path)


def init_db():
    needs_migration, from_schema_version = database_needs_migration(DB_FILE)
    backup_path = ""
    if needs_migration:
        backup_path = backup_database_before_migration(DB_FILE, from_schema_version, SCHEMA_VERSION)
        print(f"[db_migration] backup created before schema migration: {backup_path}")

    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS schema_meta (
        key TEXT PRIMARY KEY,
        value TEXT,
        updated_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS admin_user (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT DEFAULT 'admin',
        enabled INTEGER DEFAULT 1,
        updated_at TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS production_config (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        production_code TEXT UNIQUE NOT NULL,
        production_name TEXT,
        product_model TEXT,
        process_version TEXT,
        description TEXT,
        status TEXT DEFAULT 'enabled',
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS measurement_item_config (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        production_id INTEGER NOT NULL,
        item_name TEXT NOT NULL,
        process_step TEXT,
        execution_time_text TEXT,
        equipment_name TEXT,
        data_source_type TEXT DEFAULT 'auto',
        data_source_path TEXT,
        excel_sheet_name TEXT,
        image_parse_config_json TEXT,
        header_row_index INTEGER DEFAULT 1,
        csv_encoding TEXT DEFAULT 'auto',
        delimiter TEXT DEFAULT ',',
        production_code_column TEXT DEFAULT '生产编号',
        process_step_column TEXT,
        scan_frequency_seconds INTEGER DEFAULT 60,
        enabled INTEGER DEFAULT 1,
        last_collect_time TEXT,
        last_collect_status TEXT,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT,
        FOREIGN KEY(production_id) REFERENCES production_config(id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS metric_config (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_id INTEGER NOT NULL,
        metric_name TEXT NOT NULL,
        source_column TEXT NOT NULL,
        unit TEXT,
        data_type TEXT DEFAULT 'number',
        target REAL,
        lsl REAL,
        usl REAL,
        lcl REAL,
        ucl REAL,
        enabled INTEGER DEFAULT 1,
        sort_order INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT,
        FOREIGN KEY(item_id) REFERENCES measurement_item_config(id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS measurement_result (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        production_id INTEGER,
        production_code TEXT,
        item_id INTEGER,
        measurement_item_name TEXT,
        process_step TEXT,
        execution_time_text TEXT,
        equipment_name TEXT,
        metric_name TEXT,
        metric_value_text TEXT,
        metric_value_number REAL,
        unit TEXT,
        target REAL,
        lsl REAL,
        usl REAL,
        lcl REAL,
        ucl REAL,
        result_status TEXT,
        source_path TEXT,
        source_row_hash TEXT,
        source_metric_hash TEXT UNIQUE,
        collect_time TEXT DEFAULT CURRENT_TIMESTAMP,
        raw_row_json TEXT,
        is_voided INTEGER DEFAULT 0,
        voided_by TEXT,
        voided_at TEXT,
        void_reason TEXT,
        source_file_path TEXT,
        source_file_mtime TEXT,
        ocr_raw_text TEXT,
        ocr_roi_json TEXT,
        ocr_confidence REAL,
        ocr_config_json TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS collect_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        production_id INTEGER,
        production_code TEXT,
        item_id INTEGER,
        measurement_item_name TEXT,
        data_source_path TEXT,
        status TEXT,
        message TEXT,
        matched_rows INTEGER DEFAULT 0,
        inserted_count INTEGER DEFAULT 0,
        skipped_count INTEGER DEFAULT 0,
        created_at TEXT DEFAULT CURRENT_TIMESTAMP
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS collect_job (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        job_uuid TEXT UNIQUE,
        item_id INTEGER,
        production_id INTEGER,
        production_code TEXT,
        measurement_item_name TEXT,
        data_source_type TEXT,
        data_source_path TEXT,
        trigger_type TEXT,
        status TEXT,
        scheduled_at TEXT,
        started_at TEXT,
        finished_at TEXT,
        duration_ms INTEGER,
        matched_rows INTEGER DEFAULT 0,
        collect_rows INTEGER DEFAULT 0,
        inserted_count INTEGER DEFAULT 0,
        skipped_count INTEGER DEFAULT 0,
        failed_count INTEGER DEFAULT 0,
        retry_count INTEGER DEFAULT 0,
        error_message TEXT,
        traceback_text TEXT,
        created_at TEXT,
        updated_at TEXT
    )
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_collect_job_item_status ON collect_job(item_id, status)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_collect_job_created_at ON collect_job(created_at)")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS audit_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT,
        action TEXT,
        object_type TEXT,
        object_id TEXT,
        detail TEXT,
        ip_address TEXT,
        created_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS template_config (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_name TEXT NOT NULL,
        template_version TEXT DEFAULT 'v1.0',
        data_source_type TEXT DEFAULT 'csv',
        header_row_index INTEGER DEFAULT 1,
        delimiter TEXT DEFAULT ',',
        encoding TEXT DEFAULT 'auto',
        excel_sheet_name TEXT,
        production_code_column TEXT NOT NULL,
        process_step_column TEXT,
        sample_fields_json TEXT,
        description TEXT,
        created_at TEXT,
        updated_at TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS template_metric_config (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_id INTEGER NOT NULL,
        metric_name TEXT NOT NULL,
        source_column TEXT NOT NULL,
        data_type TEXT DEFAULT 'number',
        unit TEXT,
        target REAL,
        lsl REAL,
        usl REAL,
        lcl REAL,
        ucl REAL,
        sort_order INTEGER DEFAULT 0,
        created_at TEXT,
        updated_at TEXT,
        FOREIGN KEY(template_id) REFERENCES template_config(id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS template_apply_log (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        template_id INTEGER,
        production_id INTEGER,
        production_code TEXT,
        item_id INTEGER,
        applied_by TEXT,
        applied_at TEXT,
        template_name TEXT,
        template_version TEXT,
        snapshot_json TEXT,
        detail TEXT
    )
    """)

    ensure_column(cur, "production_config", "production_name", "production_name TEXT")
    ensure_column(cur, "production_config", "product_model", "product_model TEXT")
    ensure_column(cur, "production_config", "process_version", "process_version TEXT")
    ensure_column(cur, "production_config", "description", "description TEXT")
    ensure_column(cur, "production_config", "status", "status TEXT DEFAULT 'enabled'")
    ensure_column(cur, "production_config", "created_at", "created_at TEXT")
    ensure_column(cur, "production_config", "updated_at", "updated_at TEXT")
    ensure_column(cur, "admin_user", "role", "role TEXT DEFAULT 'admin'")
    ensure_column(cur, "admin_user", "enabled", "enabled INTEGER DEFAULT 1")
    ensure_column(cur, "admin_user", "updated_at", "updated_at TEXT")
    ensure_column(cur, "measurement_item_config", "production_id", "production_id INTEGER")
    ensure_column(cur, "measurement_item_config", "item_name", "item_name TEXT")
    ensure_column(cur, "measurement_item_config", "process_step", "process_step TEXT")
    ensure_column(cur, "measurement_item_config", "execution_time_text", "execution_time_text TEXT")
    ensure_column(cur, "measurement_item_config", "equipment_name", "equipment_name TEXT")
    ensure_column(cur, "measurement_item_config", "data_source_type", "data_source_type TEXT DEFAULT 'auto'")
    ensure_column(cur, "measurement_item_config", "data_source_path", "data_source_path TEXT")
    ensure_column(cur, "measurement_item_config", "excel_sheet_name", "excel_sheet_name TEXT")
    ensure_column(cur, "measurement_item_config", "image_parse_config_json", "image_parse_config_json TEXT")
    ensure_column(cur, "measurement_item_config", "header_row_index", "header_row_index INTEGER DEFAULT 1")
    ensure_column(cur, "measurement_item_config", "csv_encoding", "csv_encoding TEXT DEFAULT 'auto'")
    ensure_column(cur, "measurement_item_config", "delimiter", "delimiter TEXT DEFAULT ','")
    ensure_column(cur, "measurement_item_config", "production_code_column", "production_code_column TEXT DEFAULT '生产编号'")
    ensure_column(cur, "measurement_item_config", "process_step_column", "process_step_column TEXT")
    ensure_column(cur, "measurement_item_config", "scan_frequency_seconds", "scan_frequency_seconds INTEGER DEFAULT 60")
    ensure_column(cur, "measurement_item_config", "enabled", "enabled INTEGER DEFAULT 1")
    ensure_column(cur, "measurement_item_config", "last_collect_time", "last_collect_time TEXT")
    ensure_column(cur, "measurement_item_config", "last_collect_status", "last_collect_status TEXT")
    ensure_column(cur, "measurement_item_config", "created_at", "created_at TEXT")
    ensure_column(cur, "measurement_item_config", "updated_at", "updated_at TEXT")
    ensure_column(cur, "metric_config", "item_id", "item_id INTEGER")
    ensure_column(cur, "metric_config", "metric_name", "metric_name TEXT")
    ensure_column(cur, "metric_config", "source_column", "source_column TEXT")
    ensure_column(cur, "metric_config", "unit", "unit TEXT")
    ensure_column(cur, "metric_config", "data_type", "data_type TEXT DEFAULT 'number'")
    ensure_column(cur, "metric_config", "target", "target REAL")
    ensure_column(cur, "metric_config", "lsl", "lsl REAL")
    ensure_column(cur, "metric_config", "usl", "usl REAL")
    ensure_column(cur, "metric_config", "lcl", "lcl REAL")
    ensure_column(cur, "metric_config", "ucl", "ucl REAL")
    ensure_column(cur, "metric_config", "enabled", "enabled INTEGER DEFAULT 1")
    ensure_column(cur, "metric_config", "sort_order", "sort_order INTEGER DEFAULT 0")
    ensure_column(cur, "metric_config", "created_at", "created_at TEXT")
    ensure_column(cur, "metric_config", "updated_at", "updated_at TEXT")
    ensure_column(cur, "measurement_result", "production_id", "production_id INTEGER")
    ensure_column(cur, "measurement_result", "production_code", "production_code TEXT")
    ensure_column(cur, "measurement_result", "item_id", "item_id INTEGER")
    ensure_column(cur, "measurement_result", "measurement_item_name", "measurement_item_name TEXT")
    ensure_column(cur, "measurement_result", "process_step", "process_step TEXT")
    ensure_column(cur, "measurement_result", "execution_time_text", "execution_time_text TEXT")
    ensure_column(cur, "measurement_result", "equipment_name", "equipment_name TEXT")
    ensure_column(cur, "measurement_result", "metric_name", "metric_name TEXT")
    ensure_column(cur, "measurement_result", "metric_value_text", "metric_value_text TEXT")
    ensure_column(cur, "measurement_result", "metric_value_number", "metric_value_number REAL")
    ensure_column(cur, "measurement_result", "unit", "unit TEXT")
    ensure_column(cur, "measurement_result", "target", "target REAL")
    ensure_column(cur, "measurement_result", "lsl", "lsl REAL")
    ensure_column(cur, "measurement_result", "usl", "usl REAL")
    ensure_column(cur, "measurement_result", "lcl", "lcl REAL")
    ensure_column(cur, "measurement_result", "ucl", "ucl REAL")
    ensure_column(cur, "measurement_result", "result_status", "result_status TEXT")
    ensure_column(cur, "measurement_result", "source_path", "source_path TEXT")
    ensure_column(cur, "measurement_result", "source_row_hash", "source_row_hash TEXT")
    ensure_column(cur, "measurement_result", "source_metric_hash", "source_metric_hash TEXT")
    ensure_column(cur, "measurement_result", "collect_time", "collect_time TEXT")
    ensure_column(cur, "measurement_result", "raw_row_json", "raw_row_json TEXT")
    ensure_column(cur, "measurement_result", "is_voided", "is_voided INTEGER DEFAULT 0")
    ensure_column(cur, "measurement_result", "voided_by", "voided_by TEXT")
    ensure_column(cur, "measurement_result", "voided_at", "voided_at TEXT")
    ensure_column(cur, "measurement_result", "void_reason", "void_reason TEXT")
    ensure_column(cur, "measurement_result", "source_file_path", "source_file_path TEXT")
    ensure_column(cur, "measurement_result", "source_file_mtime", "source_file_mtime TEXT")
    ensure_column(cur, "measurement_result", "ocr_raw_text", "ocr_raw_text TEXT")
    ensure_column(cur, "measurement_result", "ocr_roi_json", "ocr_roi_json TEXT")
    ensure_column(cur, "measurement_result", "ocr_confidence", "ocr_confidence REAL")
    ensure_column(cur, "measurement_result", "ocr_config_json", "ocr_config_json TEXT")
    ensure_column(cur, "collect_log", "production_id", "production_id INTEGER")
    ensure_column(cur, "collect_log", "production_code", "production_code TEXT")
    ensure_column(cur, "collect_log", "item_id", "item_id INTEGER")
    ensure_column(cur, "collect_log", "measurement_item_name", "measurement_item_name TEXT")
    ensure_column(cur, "collect_log", "data_source_path", "data_source_path TEXT")
    ensure_column(cur, "collect_log", "status", "status TEXT")
    ensure_column(cur, "collect_log", "message", "message TEXT")
    ensure_column(cur, "collect_log", "matched_rows", "matched_rows INTEGER DEFAULT 0")
    ensure_column(cur, "collect_log", "inserted_count", "inserted_count INTEGER DEFAULT 0")
    ensure_column(cur, "collect_log", "skipped_count", "skipped_count INTEGER DEFAULT 0")
    ensure_column(cur, "collect_log", "created_at", "created_at TEXT")
    ensure_column(cur, "audit_log", "username", "username TEXT")
    ensure_column(cur, "audit_log", "action", "action TEXT")
    ensure_column(cur, "audit_log", "object_type", "object_type TEXT")
    ensure_column(cur, "audit_log", "object_id", "object_id TEXT")
    ensure_column(cur, "audit_log", "detail", "detail TEXT")
    ensure_column(cur, "audit_log", "ip_address", "ip_address TEXT")
    ensure_column(cur, "audit_log", "created_at", "created_at TEXT")
    ensure_column(cur, "template_config", "template_name", "template_name TEXT")
    ensure_column(cur, "template_config", "template_version", "template_version TEXT DEFAULT 'v1.0'")
    ensure_column(cur, "template_config", "data_source_type", "data_source_type TEXT DEFAULT 'csv'")
    ensure_column(cur, "template_config", "header_row_index", "header_row_index INTEGER DEFAULT 1")
    ensure_column(cur, "template_config", "delimiter", "delimiter TEXT DEFAULT ','")
    ensure_column(cur, "template_config", "encoding", "encoding TEXT DEFAULT 'auto'")
    ensure_column(cur, "template_config", "excel_sheet_name", "excel_sheet_name TEXT")
    ensure_column(cur, "template_config", "production_code_column", "production_code_column TEXT")
    ensure_column(cur, "template_config", "process_step_column", "process_step_column TEXT")
    ensure_column(cur, "template_config", "sample_fields_json", "sample_fields_json TEXT")
    ensure_column(cur, "template_config", "description", "description TEXT")
    ensure_column(cur, "template_config", "created_at", "created_at TEXT")
    ensure_column(cur, "template_config", "updated_at", "updated_at TEXT")
    ensure_column(cur, "template_metric_config", "template_id", "template_id INTEGER")
    ensure_column(cur, "template_metric_config", "metric_name", "metric_name TEXT")
    ensure_column(cur, "template_metric_config", "source_column", "source_column TEXT")
    ensure_column(cur, "template_metric_config", "data_type", "data_type TEXT DEFAULT 'number'")
    ensure_column(cur, "template_metric_config", "unit", "unit TEXT")
    ensure_column(cur, "template_metric_config", "target", "target REAL")
    ensure_column(cur, "template_metric_config", "lsl", "lsl REAL")
    ensure_column(cur, "template_metric_config", "usl", "usl REAL")
    ensure_column(cur, "template_metric_config", "lcl", "lcl REAL")
    ensure_column(cur, "template_metric_config", "ucl", "ucl REAL")
    ensure_column(cur, "template_metric_config", "sort_order", "sort_order INTEGER DEFAULT 0")
    ensure_column(cur, "template_metric_config", "created_at", "created_at TEXT")
    ensure_column(cur, "template_metric_config", "updated_at", "updated_at TEXT")
    ensure_column(cur, "template_apply_log", "template_id", "template_id INTEGER")
    ensure_column(cur, "template_apply_log", "production_id", "production_id INTEGER")
    ensure_column(cur, "template_apply_log", "production_code", "production_code TEXT")
    ensure_column(cur, "template_apply_log", "item_id", "item_id INTEGER")
    ensure_column(cur, "template_apply_log", "applied_by", "applied_by TEXT")
    ensure_column(cur, "template_apply_log", "applied_at", "applied_at TEXT")
    ensure_column(cur, "template_apply_log", "template_name", "template_name TEXT")
    ensure_column(cur, "template_apply_log", "template_version", "template_version TEXT")
    ensure_column(cur, "template_apply_log", "snapshot_json", "snapshot_json TEXT")
    ensure_column(cur, "template_apply_log", "detail", "detail TEXT")
    cur.execute("UPDATE admin_user SET role='admin' WHERE COALESCE(role,'')=''")
    cur.execute("UPDATE admin_user SET enabled=1 WHERE enabled IS NULL")

    def set_schema_meta(key, value):
        row = cur.execute("SELECT key FROM schema_meta WHERE key=?", (key,)).fetchone()
        if row:
            cur.execute("UPDATE schema_meta SET value=?, updated_at=? WHERE key=?", (value, now_str(), key))
        else:
            cur.execute("INSERT INTO schema_meta (key, value, updated_at) VALUES (?, ?, ?)", (key, value, now_str()))

    set_schema_meta("schema_version", SCHEMA_VERSION)
    set_schema_meta("app_version", APP_VERSION)
    if backup_path:
        set_schema_meta("last_migration_backup", backup_path)
        set_schema_meta("last_migration_from", from_schema_version or "unknown")
        set_schema_meta("last_migration_at", now_str())

    admin_username = os.environ.get("MDCP_ADMIN_USERNAME", "admin")
    admin_password = os.environ.get("MDCP_ADMIN_PASSWORD", "admin123")

    cur.execute("SELECT COUNT(*) AS c FROM admin_user")
    if cur.fetchone()["c"] == 0:
        cur.execute(
            "INSERT INTO admin_user (username, password_hash, role, enabled, updated_at) VALUES (?, ?, 'admin', 1, ?)",
            (admin_username, hash_password(admin_password), now_str())
        )
    elif os.environ.get("MDCP_ADMIN_PASSWORD"):
        # If env password is explicitly supplied, update/create that admin account.
        cur.execute("SELECT id FROM admin_user WHERE username=?", (admin_username,))
        if cur.fetchone():
            cur.execute("UPDATE admin_user SET password_hash=?, role='admin', enabled=1, updated_at=? WHERE username=?", (hash_password(admin_password), now_str(), admin_username))
        else:
            cur.execute("INSERT INTO admin_user (username, password_hash, role, enabled, updated_at) VALUES (?, ?, 'admin', 1, ?)", (admin_username, hash_password(admin_password), now_str()))

    conn.commit()
    conn.close()


# ==========================================================
# Audit log
# ==========================================================

def write_audit(username, action, object_type="", object_id="", detail="", ip_address=""):
    try:
        conn = get_conn()
        conn.execute("""
        INSERT INTO audit_log (username, action, object_type, object_id, detail, ip_address, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (username or "system", action, object_type, str(object_id or ""), detail or "", ip_address or "", now_str()))
        conn.commit()
        conn.close()
    except Exception as ex:
        print("[audit_log_failed]", ex)


# ==========================================================
# Demo data
# ==========================================================

def demo_csv_path():
    return DEMO_DATA_DIR / DEMO_CSV_FILENAME


def ensure_demo_csv_file():
    DEMO_DATA_DIR.mkdir(parents=True, exist_ok=True)
    path = demo_csv_path()
    if not path.exists() or path.read_text(encoding="utf-8-sig") != DEMO_CSV_TEXT:
        path.write_text(DEMO_CSV_TEXT, encoding="utf-8-sig")
    return str(path)


def upsert_demo_production(cur, production_code, production_name, description):
    row = cur.execute("SELECT id FROM production_config WHERE production_code=?", (production_code,)).fetchone()
    vals = (
        production_name,
        "Demo-Metrology",
        "V1-demo",
        description,
        "enabled",
        now_str(),
        production_code,
    )
    if row:
        cur.execute("""
            UPDATE production_config
            SET production_name=?, product_model=?, process_version=?, description=?, status=?, updated_at=?
            WHERE production_code=?
        """, vals)
        return row["id"]
    cur.execute("""
        INSERT INTO production_config (
            production_code, production_name, product_model, process_version, description, status, created_at, updated_at
        ) VALUES (?, ?, 'Demo-Metrology', 'V1-demo', ?, 'enabled', ?, ?)
    """, (production_code, production_name, description, now_str(), now_str()))
    return cur.lastrowid


def upsert_demo_template(cur):
    template_name = "5分钟客户演示 CSV 模板"
    sample_fields = ["生产编号", "工序", "批次", "设备", "CD", "Overlay_X", "Overlay_Y", "Rz"]
    row = cur.execute("SELECT id FROM template_config WHERE template_name=?", (template_name,)).fetchone()
    if row:
        template_id = row["id"]
        cur.execute("""
            UPDATE template_config
            SET template_version='v1.0', data_source_type='csv', header_row_index=1, delimiter=',',
                encoding='auto', excel_sheet_name='', production_code_column='生产编号',
                process_step_column='工序', sample_fields_json=?, description=?, updated_at=?
            WHERE id=?
        """, (
            json.dumps(sample_fields, ensure_ascii=False),
            "客户演示模板：从一份 CSV 采集两组生产编号、三道工序和四个指标。",
            now_str(),
            template_id,
        ))
    else:
        cur.execute("""
            INSERT INTO template_config (
                template_name, template_version, data_source_type, header_row_index, delimiter, encoding,
                excel_sheet_name, production_code_column, process_step_column, sample_fields_json,
                description, created_at, updated_at
            ) VALUES (?, 'v1.0', 'csv', 1, ',', 'auto', '', '生产编号', '工序', ?, ?, ?, ?)
        """, (
            template_name,
            json.dumps(sample_fields, ensure_ascii=False),
            "客户演示模板：从一份 CSV 采集两组生产编号、三道工序和四个指标。",
            now_str(),
            now_str(),
        ))
        template_id = cur.lastrowid

    for metric in DEMO_METRICS:
        row = cur.execute(
            "SELECT id FROM template_metric_config WHERE template_id=? AND metric_name=?",
            (template_id, metric["metric_name"]),
        ).fetchone()
        vals = (
            metric["source_column"], "number", metric["unit"], metric["target"],
            metric["lsl"], metric["usl"], metric["lcl"], metric["ucl"],
            metric["sort_order"], now_str(), template_id, metric["metric_name"],
        )
        if row:
            cur.execute("""
                UPDATE template_metric_config
                SET source_column=?, data_type=?, unit=?, target=?, lsl=?, usl=?, lcl=?, ucl=?,
                    sort_order=?, updated_at=?
                WHERE template_id=? AND metric_name=?
            """, vals)
        else:
            cur.execute("""
                INSERT INTO template_metric_config (
                    template_id, metric_name, source_column, data_type, unit, target, lsl, usl, lcl, ucl,
                    sort_order, created_at, updated_at
                ) VALUES (?, ?, ?, 'number', ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                template_id, metric["metric_name"], metric["source_column"], metric["unit"],
                metric["target"], metric["lsl"], metric["usl"], metric["lcl"], metric["ucl"],
                metric["sort_order"], now_str(), now_str(),
            ))
    return template_id


def upsert_demo_item(cur, production_id, csv_path):
    item_name = "Demo CSV 三工序量测"
    row = cur.execute(
        "SELECT id FROM measurement_item_config WHERE production_id=? AND item_name=?",
        (production_id, item_name),
    ).fetchone()
    vals = (
        "", "工序", "工序完成后自动采集", "Demo-Metrology-Line",
        "csv", csv_path, "", "", 1, "auto", ",", "生产编号", 60, 1, now_str(),
    )
    if row:
        item_id = row["id"]
        cur.execute("""
            UPDATE measurement_item_config
            SET process_step=?, process_step_column=?, execution_time_text=?, equipment_name=?,
                data_source_type=?, data_source_path=?, excel_sheet_name=?, image_parse_config_json=?,
                header_row_index=?, csv_encoding=?, delimiter=?, production_code_column=?,
                scan_frequency_seconds=?, enabled=?, updated_at=?
            WHERE id=?
        """, vals + (item_id,))
    else:
        cur.execute("""
            INSERT INTO measurement_item_config (
                production_id, item_name, process_step, process_step_column, execution_time_text, equipment_name,
                data_source_type, data_source_path, excel_sheet_name, image_parse_config_json,
                header_row_index, csv_encoding, delimiter, production_code_column,
                scan_frequency_seconds, enabled, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (production_id, item_name) + vals[:-1] + (now_str(), now_str()))
        item_id = cur.lastrowid

    for metric in DEMO_METRICS:
        row = cur.execute(
            "SELECT id FROM metric_config WHERE item_id=? AND metric_name=?",
            (item_id, metric["metric_name"]),
        ).fetchone()
        vals = (
            metric["source_column"], "number", metric["unit"], metric["target"],
            metric["lsl"], metric["usl"], metric["lcl"], metric["ucl"], 1,
            metric["sort_order"], now_str(), item_id, metric["metric_name"],
        )
        if row:
            cur.execute("""
                UPDATE metric_config
                SET source_column=?, data_type=?, unit=?, target=?, lsl=?, usl=?, lcl=?, ucl=?,
                    enabled=?, sort_order=?, updated_at=?
                WHERE item_id=? AND metric_name=?
            """, vals)
        else:
            cur.execute("""
                INSERT INTO metric_config (
                    item_id, metric_name, source_column, data_type, unit, target, lsl, usl, lcl, ucl,
                    enabled, sort_order, created_at, updated_at
                ) VALUES (?, ?, ?, 'number', ?, ?, ?, ?, ?, ?, 1, ?, ?, ?)
            """, (
                item_id, metric["metric_name"], metric["source_column"], metric["unit"],
                metric["target"], metric["lsl"], metric["usl"], metric["lcl"], metric["ucl"],
                metric["sort_order"], now_str(), now_str(),
            ))
    return item_id


def seed_demo_data(user, ip_address=""):
    require_permission(user, can_manage_config(user))
    csv_path = ensure_demo_csv_file()
    conn = get_conn()
    cur = conn.cursor()
    template_id = upsert_demo_template(cur)
    production_specs = [
        ("DEMO_WAFER_A01", "Demo 晶圆 A01", "演示：整体稳定，仅末道 AOI 有轻微 MS2 风险。"),
        ("DEMO_WAFER_B02", "Demo 晶圆 B02", "演示：末道 AOI 出现超 MS2，需要客户追溯来源文件。"),
    ]
    item_ids = []
    for code, name, description in production_specs:
        production_id = upsert_demo_production(cur, code, name, description)
        item_ids.append(upsert_demo_item(cur, production_id, csv_path))
    cur.execute("""
        UPDATE measurement_result
        SET is_voided=0, voided_by='', voided_at='', void_reason=''
        WHERE production_code IN ('DEMO_WAFER_A01', 'DEMO_WAFER_B02')
    """)
    conn.commit()
    conn.close()

    collect_results = []
    for item_id in item_ids:
        collect_results.append(COLLECTOR.collect_item(item_id, trigger_type="demo_seed", dry_run=False))

    conn = get_conn()
    summary = conn.execute("""
        SELECT COUNT(*) AS total,
               SUM(CASE WHEN result_status IN ('MS3_PASS','PASS') THEN 1 ELSE 0 END) AS ms3_count,
               SUM(CASE WHEN result_status IN ('MS2_PASS','OOC') THEN 1 ELSE 0 END) AS ms2_count,
               SUM(CASE WHEN result_status IN ('MISS_MS2','OOS') THEN 1 ELSE 0 END) AS miss_count
        FROM measurement_result
        WHERE production_code IN ('DEMO_WAFER_A01', 'DEMO_WAFER_B02') AND COALESCE(is_voided,0)=0
    """).fetchone()
    conn.close()
    inserted = sum(safe_int(r.get("inserted", 0), 0) for r in collect_results)
    skipped = sum(safe_int(r.get("skipped", 0), 0) for r in collect_results)
    write_audit(
        user.get("username") if user else "system",
        "SEED_DEMO_DATA",
        "demo",
        ",".join(str(x) for x in item_ids),
        f"加载客户演示 Demo 数据，新增 {inserted} 条，跳过重复 {skipped} 条，模板ID={template_id}",
        ip_address,
    )
    return {
        "csv_path": csv_path,
        "template_id": template_id,
        "item_ids": item_ids,
        "collect_results": collect_results,
        "inserted": inserted,
        "skipped": skipped,
        "total": summary["total"] or 0,
        "ms3_count": summary["ms3_count"] or 0,
        "ms2_count": summary["ms2_count"] or 0,
        "miss_count": summary["miss_count"] or 0,
    }


def log_request_exception(user, method, path, ex, ip_address=""):
    request_id = secrets.token_hex(4)
    detail = traceback.format_exc()
    print(f"[request_error:{request_id}] {method} {path} user={(user or {}).get('username', 'anonymous')} ip={ip_address}\n{detail}")
    write_audit(
        (user or {}).get("username", "unknown"),
        "REQUEST_ERROR",
        "route",
        path,
        f"请求处理失败，错误编号 {request_id}：{ex}",
        ip_address,
    )
    return request_id


def page_request_error(user, request_id):
    return base_layout("处理失败", f"""
    <h1>处理失败</h1>
    <div class="card">
      <p class="error">当前操作没有完成。请记录错误编号并联系管理员查看服务端日志。</p>
      <p><b>错误编号：</b>{h(request_id)}</p>
      <p class="note">页面不会向客户展示内部路径和 traceback；详细信息已写入服务端日志和审计日志。</p>
      <a class="btn secondary" href="/">返回 Dashboard</a>
    </div>
    """, user)

# ==========================================================
# CSV collection service
# ==========================================================

def _read_file_bytes_stably(path: str):
    """Read bytes with retry and size/mtime stability check.

    This avoids parsing a half-written CSV when the equipment is writing the file.
    It does not request write access and it never modifies the source file.
    """
    last_error = None
    for attempt in range(1, READ_RETRY_COUNT + 1):
        try:
            if not path:
                raise FileNotFoundError("数据源路径为空")
            if not os.path.exists(path):
                raise FileNotFoundError(f"路径不存在或无权限访问：{path}")

            st1 = os.stat(path)
            time.sleep(FILE_STABLE_WAIT_SECONDS)
            st2 = os.stat(path)
            if (st1.st_size != st2.st_size) or (int(st1.st_mtime_ns) != int(st2.st_mtime_ns)):
                raise RuntimeError("文件仍在写入或变化中，等待下次重试")

            # Open read-only. On Windows, if writer uses exclusive lock, this may raise PermissionError.
            with open(path, "rb") as f:
                data = f.read()

            st3 = os.stat(path)
            if (st2.st_size != st3.st_size) or (int(st2.st_mtime_ns) != int(st3.st_mtime_ns)):
                raise RuntimeError("读取期间文件发生变化，等待下次重试")
            return data, st3
        except (PermissionError, OSError, RuntimeError, FileNotFoundError) as ex:
            last_error = ex
            if attempt < READ_RETRY_COUNT:
                time.sleep(READ_RETRY_INTERVAL_SECONDS * attempt)
                continue
            raise last_error


def read_csv_rows(path: str, encoding: str, delimiter: str):
    """Read CSV as DictReader with retry, stability checks and encoding fallback.

    Read robustness:
    - retries PermissionError/OSError when equipment is writing or locking the file;
    - checks size/mtime stability before and after read;
    - decodes from an in-memory byte snapshot, so parsing does not hold the source file;
    - supports encoding auto-detection fallback.
    """
    if delimiter == "\\t":
        delimiter = "\t"

    data, stat_info = _read_file_bytes_stably(path)

    encodings = []
    configured = (encoding or "auto").strip().lower()
    if configured and configured != "auto":
        encodings.append(configured)
    for fallback in ["utf-8-sig", "utf-8", "gb18030", "gbk", "cp936", "big5"]:
        if fallback not in encodings:
            encodings.append(fallback)

    last_error = None
    for enc in encodings:
        try:
            text = data.decode(enc, errors="strict")
            reader = csv.DictReader(io.StringIO(text), delimiter=delimiter or ",")
            rows = [dict(r) for r in reader]
            fieldnames = reader.fieldnames or []
            return fieldnames, rows, enc
        except UnicodeDecodeError as ex:
            last_error = ex
            continue

    # Last-resort tolerant read. This protects production collection from one abnormal character.
    text = data.decode("gb18030", errors="replace")
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter or ",")
    rows = [dict(r) for r in reader]
    fieldnames = reader.fieldnames or []
    return fieldnames, rows, "gb18030(errors=replace)"


def excel_column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in (cell_ref or "") if ch.isalpha()).upper()
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return max(0, idx - 1)


def normalize_xlsx_target(target: str) -> str:
    target = (target or "").replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return "xl/" + target.lstrip("/")


def cell_ref(col_index: int, row_index: int) -> str:
    col = ""
    n = col_index + 1
    while n:
        n, rem = divmod(n - 1, 26)
        col = chr(ord("A") + rem) + col
    return f"{col}{row_index}"


def xlsx_list_sheets_from_bytes(data: bytes):
    ns_main = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    ns_rel = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}
    office_rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rels = {rel.attrib.get("Id"): rel.attrib.get("Target", "") for rel in rels_root.findall("r:Relationship", ns_rel)}
        sheets = []
        for sheet in workbook.findall(".//x:sheet", ns_main):
            name = sheet.attrib.get("name", "")
            rid = sheet.attrib.get(office_rel, "")
            target = normalize_xlsx_target(rels.get(rid, ""))
            sheets.append((name, target))
        return sheets


def _xlsx_rows_matrix_from_bytes(data: bytes, sheet_name: str = ""):
    ns_main = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        shared_strings = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("x:si", ns_main):
                shared_strings.append("".join(t.text or "" for t in si.findall(".//x:t", ns_main)))

        sheets = xlsx_list_sheets_from_bytes(data)
        if not sheets:
            raise ValueError("Excel 文件中没有可读取的 Sheet。")
        selected = sheets[0]
        if sheet_name:
            selected = next((s for s in sheets if s[0] == sheet_name), None)
            if selected is None:
                available = "，".join(name for name, _target in sheets)
                raise ValueError(f"找不到 Sheet：{sheet_name}。可用 Sheet：{available}")

        worksheet = ET.fromstring(zf.read(selected[1]))
        parsed_rows = []
        for row in worksheet.findall(".//x:sheetData/x:row", ns_main):
            values = {}
            for cell in row.findall("x:c", ns_main):
                col_idx = excel_column_index(cell.attrib.get("r", ""))
                cell_type = cell.attrib.get("t", "")
                if cell_type == "inlineStr":
                    value = "".join(t.text or "" for t in cell.findall(".//x:t", ns_main))
                else:
                    v = cell.find("x:v", ns_main)
                    raw = "" if v is None or v.text is None else v.text
                    if cell_type == "s":
                        try:
                            value = shared_strings[int(raw)]
                        except Exception:
                            value = raw
                    elif cell_type == "b":
                        value = "TRUE" if raw == "1" else "FALSE"
                    else:
                        value = raw
                values[col_idx] = value
            if values:
                max_col = max(values.keys())
                parsed_rows.append([values.get(i, "") for i in range(max_col + 1)])
    return parsed_rows, selected[0]


def _rows_matrix_to_dicts(parsed_rows, selected_name: str, header_row_index: int = 1, preview_limit=None):
    if not parsed_rows:
        return [], [], f"xlsx:{selected_name}"
    header_idx = max(0, header_row_index - 1)
    if header_idx >= len(parsed_rows):
        raise ValueError(f"表头所在行 {header_row_index} 超出 Excel 数据范围。")
    fieldnames = [str(v).strip() for v in parsed_rows[header_idx]]
    rows = []
    for raw_row in parsed_rows[header_idx + 1:]:
        if not any(str(v).strip() for v in raw_row):
            continue
        row_dict = {}
        for idx, name in enumerate(fieldnames):
            if name:
                row_dict[name] = raw_row[idx] if idx < len(raw_row) else ""
        rows.append(row_dict)
        if preview_limit is not None and len(rows) >= preview_limit:
            break
    return fieldnames, rows, f"xlsx:{selected_name}"


def _parse_xlsx_rows_from_bytes(data: bytes, sheet_name: str = "", header_row_index: int = 1):
    parsed_rows, selected_name = _xlsx_rows_matrix_from_bytes(data, sheet_name)
    return _rows_matrix_to_dicts(parsed_rows, selected_name, header_row_index, preview_limit=10)


def read_xlsx_rows(path: str, sheet_name: str = "", header_row_index: int = 1):
    if Path(path or "").suffix.lower() == ".xls":
        raise ValueError("当前版本支持 .xlsx/.xlsm；旧版 .xls 请先另存为 .xlsx。")
    data, _stat_info = _read_file_bytes_stably(path)
    parsed_rows, selected_name = _xlsx_rows_matrix_from_bytes(data, sheet_name)
    return _rows_matrix_to_dicts(parsed_rows, selected_name, header_row_index, preview_limit=None)


def is_image_path(path: str) -> bool:
    return Path(path or "").suffix.lower() in IMAGE_SUFFIXES


def has_glob_pattern(path: str) -> bool:
    return any(ch in (path or "") for ch in "*?[")


def truthy_config(value):
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return value != 0
    return str(value or "").strip().lower() in ("1", "true", "yes", "y", "on")


def validate_regex(pattern: str, label: str):
    if not str(pattern or "").strip():
        return
    try:
        re.compile(pattern)
    except re.error as ex:
        raise ValueError(f"Invalid Image OCR regex for {label}: {ex}") from ex


def normalize_image_roi(roi, metric_name=""):
    if isinstance(roi, dict):
        values = [roi.get("x"), roi.get("y"), roi.get("w", roi.get("width")), roi.get("h", roi.get("height"))]
    elif isinstance(roi, (list, tuple)) and len(roi) == 4:
        values = list(roi)
    else:
        raise ValueError(f"Image OCR metric {metric_name} missing valid roi [x,y,w,h].")
    try:
        x, y, w, hgt = [float(v) for v in values]
    except Exception as ex:
        raise ValueError(f"Image OCR metric {metric_name} roi must contain numbers.") from ex
    if x < 0 or y < 0 or w <= 0 or hgt <= 0 or x + w > 1.000001 or y + hgt > 1.000001:
        raise ValueError(f"Image OCR metric {metric_name} roi must be normalized inside 0..1.")
    return [x, y, w, hgt]


def get_image_metadata_field_config(config: dict, field_name: str):
    fields = config.get("fields")
    aliases = {
        "production_code": ("production_code", "生产编号", "prod_code"),
        "process_step": ("process_step", "工序", "工序名", "step"),
    }.get(field_name, (field_name,))
    if isinstance(fields, dict):
        for alias in aliases:
            field_cfg = fields.get(alias)
            if isinstance(field_cfg, dict):
                return field_cfg
    legacy_cfg = config.get(f"{field_name}_ocr")
    if isinstance(legacy_cfg, dict):
        return legacy_cfg
    return None


def image_config_routes_by_production(config: dict):
    if truthy_config(config.get("route_by_image_production_code")):
        return True
    if str(config.get("production_code_from_filename_regex") or config.get("filename_production_code_regex") or "").strip():
        return True
    return get_image_metadata_field_config(config, "production_code") is not None


def image_config_process_column(config: dict, configured_process_column: str):
    configured_process_column = (configured_process_column or "").strip()
    if configured_process_column:
        return configured_process_column
    has_filename_process = str(config.get("process_from_filename_regex") or config.get("filename_process_regex") or "").strip()
    if has_filename_process or get_image_metadata_field_config(config, "process_step"):
        return "process_step"
    return ""


def parse_image_parse_config(config_json: str, required_metric_columns=None):
    if not str(config_json or "").strip():
        raise ValueError("Image OCR config JSON is required.")
    try:
        config = json.loads(config_json)
    except json.JSONDecodeError as ex:
        raise ValueError(f"Invalid Image OCR config JSON: {ex}") from ex
    if not isinstance(config, dict):
        raise ValueError("Image OCR config JSON must be an object.")
    metric_configs = config.get("metrics")
    if not isinstance(metric_configs, dict) or not metric_configs:
        raise ValueError("Image OCR config JSON must include a non-empty metrics object.")
    production_code_pattern = config.get("production_code_from_filename_regex") or config.get("filename_production_code_regex") or ""
    validate_regex(production_code_pattern, "production_code_from_filename_regex")
    validate_regex(config.get("process_from_filename_regex") or config.get("filename_process_regex") or "", "process_from_filename_regex")
    if truthy_config(config.get("route_by_image_production_code")) and not str(production_code_pattern).strip() and not get_image_metadata_field_config(config, "production_code"):
        raise ValueError("Image OCR route_by_image_production_code requires production_code_from_filename_regex or fields.production_code OCR config.")
    for metadata_name in ("production_code", "process_step"):
        metadata_cfg = get_image_metadata_field_config(config, metadata_name)
        if not metadata_cfg:
            continue
        if "roi" not in metadata_cfg:
            raise ValueError(f"Image OCR metadata field {metadata_name} missing roi.")
        normalize_image_roi(metadata_cfg.get("roi"), metadata_name)
        if not str(metadata_cfg.get("regex") or "").strip():
            raise ValueError(f"Image OCR metadata field {metadata_name} missing regex.")
        validate_regex(metadata_cfg.get("regex"), metadata_name)
    required = [str(c).strip() for c in (required_metric_columns or metric_configs.keys()) if str(c).strip()]
    for metric_name in required:
        metric_cfg = metric_configs.get(metric_name)
        if not isinstance(metric_cfg, dict):
            raise ValueError(f"Image OCR config missing metric config for {metric_name}.")
        if "roi" not in metric_cfg:
            raise ValueError(f"Image OCR metric {metric_name} missing roi.")
        normalize_image_roi(metric_cfg.get("roi"), metric_name)
        if not str(metric_cfg.get("regex") or "").strip():
            raise ValueError(f"Image OCR metric {metric_name} missing regex.")
        validate_regex(metric_cfg.get("regex"), metric_name)
    return config


def image_collects_multiple_files(config: dict):
    mode = str(config.get("collect_mode") or config.get("file_selection") or "latest").strip().lower()
    return mode in ("all", "all_stable", "all_new", "batch", "folder")


def image_max_files_per_collect(config: dict):
    default = 50 if image_collects_multiple_files(config) else 1
    value = safe_int(config.get("max_files_per_collect", default), default)
    return max(1, min(value, 500))


def find_stable_image_files(path: str, config: dict):
    if not path:
        raise FileNotFoundError("Image source path is empty.")
    source = Path(path)
    if source.is_dir():
        pattern = str(config.get("file_pattern") or "*")
        candidates = list(source.glob(pattern))
    elif has_glob_pattern(path):
        candidates = [Path(p) for p in glob.glob(path, recursive=True)]
    else:
        candidates = [source]
    candidates = [p for p in candidates if p.is_file() and p.suffix.lower() in IMAGE_SUFFIXES]
    if not candidates:
        raise FileNotFoundError(f"No supported image files found for path: {path}")
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    max_files = image_max_files_per_collect(config)
    last_error = None
    selected = []
    for candidate in candidates:
        try:
            data, stat_info = _read_file_bytes_stably(str(candidate))
            selected.append((str(candidate), data, stat_info))
            if len(selected) >= max_files:
                break
        except Exception as ex:
            last_error = ex
            continue
    if selected:
        return selected
    raise RuntimeError(f"No stable image file could be read from {path}: {last_error}")


def find_stable_image_file(path: str, config: dict):
    return find_stable_image_files(path, {**config, "collect_mode": "latest", "max_files_per_collect": 1})[0]


def preprocess_image_roi_for_ocr(image_bytes: bytes, roi, ocr_config: dict):
    try:
        from PIL import Image
        import cv2
        import numpy as np
    except ImportError as ex:
        raise RuntimeError("Image OCR requires Pillow, opencv-python-headless, numpy and pytesseract. Install requirements_ocr.txt.") from ex

    image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
    width, height = image.size
    x, y, w, hgt = normalize_image_roi(roi)
    left = max(0, min(width - 1, int(round(x * width))))
    top = max(0, min(height - 1, int(round(y * height))))
    right = max(left + 1, min(width, int(round((x + w) * width))))
    bottom = max(top + 1, min(height, int(round((y + hgt) * height))))
    cropped = image.crop((left, top, right, bottom)).convert("L")
    arr = np.array(cropped)
    scale = float(ocr_config.get("scale", 2.0) or 1.0)
    if scale != 1.0:
        arr = cv2.resize(arr, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    if ocr_config.get("threshold", True):
        arr = cv2.threshold(arr, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    return Image.fromarray(arr), {"left": left, "top": top, "right": right, "bottom": bottom}


def extract_named_regex_from_text(text: str, pattern: str, preferred_group: str = "", label: str = ""):
    match = re.search(pattern, text or "", re.IGNORECASE | re.MULTILINE)
    if not match:
        raise ValueError(f"Image OCR text for {label} did not match regex.")
    if preferred_group and preferred_group in match.groupdict():
        return match.group(preferred_group).strip()
    if "value" in match.groupdict():
        return match.group("value").strip()
    if match.groups():
        return match.group(1).strip()
    return match.group(0).strip()


def extract_regex_value(text: str, pattern: str, metric_name=""):
    return extract_named_regex_from_text(text, pattern, "value", metric_name)


def run_image_ocr_with_metadata(image_bytes: bytes, image_path: str, config: dict, required_metric_columns):
    try:
        import pytesseract
    except ImportError as ex:
        raise RuntimeError("Image OCR requires pytesseract. Install requirements_ocr.txt.") from ex
    tesseract_cmd = os.environ.get("MDCP_TESSERACT_CMD", "").strip()
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    metric_configs = config.get("metrics") or {}
    ocr_config = config.get("ocr") if isinstance(config.get("ocr"), dict) else {}
    lang = str(ocr_config.get("lang") or "eng")
    psm = safe_int(ocr_config.get("psm", 6), 6)
    extra_config = str(ocr_config.get("config") or "").strip()
    tesseract_config = f"--psm {psm}" + (f" {extra_config}" if extra_config else "")
    values = {}
    metadata_values = {}
    debug = {
        "image_path": image_path,
        "fields": {},
        "metrics": {}
    }
    metadata_fields = {
        "production_code": get_image_metadata_field_config(config, "production_code"),
        "process_step": get_image_metadata_field_config(config, "process_step"),
    }
    for field_name, field_cfg in metadata_fields.items():
        if not field_cfg:
            continue
        processed_image, pixel_roi = preprocess_image_roi_for_ocr(image_bytes, field_cfg["roi"], ocr_config)
        raw_text = pytesseract.image_to_string(processed_image, lang=lang, config=tesseract_config)
        value = extract_named_regex_from_text(raw_text, field_cfg["regex"], field_name, field_name)
        metadata_values[field_name] = value
        debug["fields"][field_name] = {
            "roi": field_cfg["roi"],
            "pixel_roi": pixel_roi,
            "regex": field_cfg["regex"],
            "ocr_text": raw_text,
            "value": value
        }
    for metric_name in required_metric_columns:
        metric_cfg = metric_configs.get(metric_name)
        processed_image, pixel_roi = preprocess_image_roi_for_ocr(image_bytes, metric_cfg["roi"], ocr_config)
        raw_text = pytesseract.image_to_string(processed_image, lang=lang, config=tesseract_config)
        value = extract_regex_value(raw_text, metric_cfg["regex"], metric_name)
        values[metric_name] = value
        debug["metrics"][metric_name] = {
            "roi": metric_cfg["roi"],
            "pixel_roi": pixel_roi,
            "regex": metric_cfg["regex"],
            "ocr_text": raw_text,
            "value": value
        }
    return values, debug, metadata_values


def run_image_ocr(image_bytes: bytes, image_path: str, config: dict, required_metric_columns):
    values, debug, _metadata_values = run_image_ocr_with_metadata(image_bytes, image_path, config, required_metric_columns)
    return values, debug


def extract_from_filename(image_path: str, pattern: str, preferred_group: str = ""):
    if not pattern:
        return ""
    target = Path(image_path).name
    match = re.search(pattern, target)
    if not match:
        return ""
    if preferred_group and preferred_group in match.groupdict():
        return match.group(preferred_group).strip()
    if "value" in match.groupdict():
        return match.group("value").strip()
    if match.groups():
        return match.group(1).strip()
    return match.group(0).strip()


def extract_production_code_from_filename(image_path: str, config: dict):
    pattern = (config.get("production_code_from_filename_regex") or config.get("filename_production_code_regex") or "").strip()
    return extract_from_filename(image_path, pattern, "production_code")


def extract_process_from_filename(image_path: str, config: dict):
    pattern = (config.get("process_from_filename_regex") or config.get("filename_process_regex") or "").strip()
    return extract_from_filename(image_path, pattern, "process_step")


def read_image_rows(path: str, image_config_json: str, production_code: str, code_column: str,
                    fixed_process_step: str, process_column: str, required_metric_columns):
    required_metric_columns = [str(c).strip() for c in (required_metric_columns or []) if str(c).strip()]
    config = parse_image_parse_config(image_config_json, required_metric_columns)
    code_field = code_column or "production_code"
    process_column = image_config_process_column(config, process_column)
    rows = []
    labels = []
    parse_errors = []
    for image_path, image_bytes, stat_info in find_stable_image_files(path, config):
        try:
            values, debug, metadata_values = run_image_ocr_with_metadata(image_bytes, image_path, config, required_metric_columns)
            parsed_production_code = (
                extract_production_code_from_filename(image_path, config)
                or metadata_values.get("production_code")
                or production_code
            )
            row = {
                code_field: parsed_production_code,
                "_source_path": image_path,
                "_source_mtime": datetime.fromtimestamp(stat_info.st_mtime, APP_TZ).strftime("%Y-%m-%d %H:%M:%S"),
                "_ocr": debug
            }
            if process_column:
                row[process_column] = (
                    extract_process_from_filename(image_path, config)
                    or metadata_values.get("process_step")
                    or fixed_process_step
                    or ""
                )
            for metric_name in required_metric_columns:
                row[metric_name] = values.get(metric_name, "")
            rows.append(row)
        except Exception as ex:
            parse_errors.append({"image_path": image_path, "error": str(ex)})
        labels.append(image_path)
    if not rows and parse_errors:
        raise RuntimeError("No image rows could be parsed. " + "; ".join(f"{e['image_path']}: {e['error']}" for e in parse_errors[:5]))
    if parse_errors and rows:
        rows[0]["_image_parse_errors"] = parse_errors[:20]
    fieldnames = [code_field]
    if process_column:
        fieldnames.append(process_column)
    fieldnames.extend(required_metric_columns)
    fieldnames.extend(["_source_path", "_source_mtime", "_ocr"])
    if parse_errors:
        fieldnames.append("_image_parse_errors")
    return fieldnames, rows, "image:" + ",".join(labels[:5]) + ("..." if len(labels) > 5 else "")


def read_source_rows(data_source_type: str, path: str, encoding: str, delimiter: str, sheet_name: str = "", header_row_index: int = 1,
                     image_config_json: str = "", production_code: str = "", code_column: str = "",
                     fixed_process_step: str = "", process_column: str = "", required_metric_columns=None):
    source_type = (data_source_type or "auto").strip().lower()
    suffix = Path(path or "").suffix.lower()
    if source_type == "auto":
        if suffix in (".xlsx", ".xlsm", ".xls"):
            source_type = "excel"
        elif suffix in IMAGE_SUFFIXES or (path and (Path(path).is_dir() or has_glob_pattern(path))):
            source_type = "image"
        else:
            source_type = "csv"
    if source_type == "excel":
        return read_xlsx_rows(path, sheet_name, header_row_index)
    if source_type == "image":
        return read_image_rows(
            path, image_config_json, production_code, code_column,
            fixed_process_step, process_column, required_metric_columns or []
        )
    return read_csv_rows(path, encoding, delimiter)


def save_template_cache_file(file_bytes: bytes, filename: str):
    TEMPLATE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    safe_suffix = Path(filename or "template.xlsx").suffix.lower() or ".xlsx"
    token = secrets.token_urlsafe(16)
    path = TEMPLATE_CACHE_DIR / f"{token}{safe_suffix}"
    path.write_bytes(file_bytes)
    return token, str(path)


def get_template_cache_path(token: str):
    if not token:
        return None
    for p in TEMPLATE_CACHE_DIR.glob(token + ".*"):
        return str(p)
    return None

def judge_status(value_number, ms2_lower, ms2_upper, ms3_lower, ms3_upper):
    if value_number is None:
        return "TEXT"
    has_ms2 = ms2_lower is not None or ms2_upper is not None
    has_ms3 = ms3_lower is not None or ms3_upper is not None
    if ms2_lower is not None and value_number < ms2_lower:
        return "MISS_MS2"
    if ms2_upper is not None and value_number > ms2_upper:
        return "MISS_MS2"
    if ms3_lower is not None and value_number < ms3_lower:
        return "MS2_PASS"
    if ms3_upper is not None and value_number > ms3_upper:
        return "MS2_PASS"
    if has_ms3:
        return "MS3_PASS"
    if has_ms2:
        return "MS2_PASS"
    return "PASS"


def metric_ocr_debug(row: dict, source_col: str):
    debug = row.get("_ocr") if isinstance(row, dict) else None
    if not isinstance(debug, dict):
        return {}
    metrics = debug.get("metrics")
    if not isinstance(metrics, dict):
        return {}
    metric_debug = metrics.get(source_col)
    return metric_debug if isinstance(metric_debug, dict) else {}


def ocr_roi_snapshot(metric_debug: dict):
    if not metric_debug:
        return ""
    return json.dumps({
        "roi": metric_debug.get("roi"),
        "pixel_roi": metric_debug.get("pixel_roi"),
        "regex": metric_debug.get("regex")
    }, ensure_ascii=False)


def result_trace_fields(row: dict, source_col: str, actual_source_path: str, image_config_json: str):
    metric_debug = metric_ocr_debug(row, source_col)
    source_file_path = row.get("_source_path") or actual_source_path
    source_file_mtime = row.get("_source_mtime") or ""
    if not metric_debug:
        return source_file_path, source_file_mtime, "", "", None, ""
    # The current OCR path uses pytesseract.image_to_string, which does not expose
    # per-ROI confidence. Keep NULL here until image_to_data is introduced.
    return (
        source_file_path,
        source_file_mtime,
        metric_debug.get("ocr_text") or "",
        ocr_roi_snapshot(metric_debug),
        None,
        image_config_json or ""
    )


def image_ocr_metric_details(rows, metrics, code_column: str, process_column: str,
                             fallback_production_code: str, fallback_process_step: str):
    details = []
    for row in rows or []:
        if not isinstance(row, dict) or not row.get("_ocr"):
            continue
        production_code = str(row.get(code_column, "")).strip() or fallback_production_code or ""
        process_step = str(row.get(process_column, "")).strip() if process_column else ""
        process_step = process_step or fallback_process_step or ""
        for metric in metrics:
            source_col = metric["source_column"]
            metric_debug = metric_ocr_debug(row, source_col)
            details.append({
                "image_path": row.get("_source_path") or "",
                "process_step": process_step,
                "production_code": production_code,
                "metric_name": metric["metric_name"],
                "source_column": source_col,
                "ocr_raw_text": metric_debug.get("ocr_text") if metric_debug else "",
                "extracted_value": row.get(source_col),
                "roi": metric_debug.get("roi") if metric_debug else "",
                "pixel_roi": metric_debug.get("pixel_roi") if metric_debug else ""
            })
    return details


def collect_item_core(item_id: int, dry_run=False):
    """Collect one measurement item. If dry_run=True, only test and return detail."""
    conn = get_conn()
    cur = conn.cursor()
    item = cur.execute("""
        SELECT mi.*, p.production_code, p.production_name
        FROM measurement_item_config mi
        JOIN production_config p ON p.id = mi.production_id
        WHERE mi.id=?
    """, (item_id,)).fetchone()

    if not item:
        conn.close()
        return {"ok": False, "status": "NOT_FOUND", "message": "量测项不存在"}

    metrics = cur.execute("""
        SELECT * FROM metric_config
        WHERE item_id=? AND enabled=1
        ORDER BY sort_order, id
    """, (item_id,)).fetchall()

    if not metrics:
        status = "NO_METRICS_CONFIGURED"
        msg = "该量测项下没有启用的指标配置。请先进入“指标”，添加 Dx1、Dy1、Dx2、Dy2、Rz 等指标后再采集。"
        if not dry_run:
            write_collect_log(cur, item, status, msg, 0, 0, 0)
            update_item_status(cur, item_id, status)
            conn.commit()
        conn.close()
        return {
            "ok": False,
            "status": status,
            "message": msg,
            "matched_rows": 0,
            "inserted": 0,
            "skipped": 0,
            "metric_preview": {}
        }

    production_code = item["production_code"]
    data_source_path = item["data_source_path"]
    code_column = item["production_code_column"] or "生产编号"
    process_column = (item["process_step_column"] if "process_step_column" in item.keys() else "") or ""
    process_column = process_column.strip()
    fixed_process_step = (item["process_step"] or "").strip()
    data_source_type = (item["data_source_type"] if "data_source_type" in item.keys() else "auto") or "auto"
    image_source_hint = (
        data_source_type.strip().lower() == "image"
        or (
            data_source_type.strip().lower() == "auto"
            and (
                Path(data_source_path or "").suffix.lower() in IMAGE_SUFFIXES
                or (data_source_path and (Path(data_source_path).is_dir() or has_glob_pattern(data_source_path)))
            )
        )
    )
    image_config_for_collect = {}
    image_dynamic_production = False
    if image_source_hint:
        try:
            image_config_for_collect = json.loads(item["image_parse_config_json"] or "{}")
            if isinstance(image_config_for_collect, dict):
                process_column = image_config_process_column(image_config_for_collect, process_column)
                image_dynamic_production = image_config_routes_by_production(image_config_for_collect)
        except Exception:
            image_config_for_collect = {}

    # 禁止“无工序”采集。
    # 必须配置固定工序，或配置工序字段并从数据源中读取工序。
    if not process_column and not fixed_process_step:
        status = "PROCESS_STEP_REQUIRED"
        msg = "该量测项未配置固定工序，也未配置工序字段名。为避免生成无工序采集结果，系统已拒绝采集。请在量测项配置中填写“固定量测工序”或“工序字段名”。"
        if not dry_run:
            write_collect_log(cur, item, status, msg, 0, 0, 0)
            update_item_status(cur, item_id, status)
            conn.commit()
        conn.close()
        return {
            "ok": False,
            "status": status,
            "message": msg,
            "matched_rows": 0,
            "collect_rows": 0,
            "inserted": 0,
            "skipped": 0,
            "metric_preview": {}
        }

    inserted = 0
    skipped = 0
    try:
        fieldnames, rows, used_encoding = read_source_rows(
            item["data_source_type"] if "data_source_type" in item.keys() else "auto",
            data_source_path,
            item["csv_encoding"] or "auto",
            item["delimiter"] or ",",
            item["excel_sheet_name"] if "excel_sheet_name" in item.keys() else "",
            item["header_row_index"] if "header_row_index" in item.keys() else 1,
            item["image_parse_config_json"] if "image_parse_config_json" in item.keys() else "",
            production_code,
            code_column,
            fixed_process_step,
            process_column,
            [m["source_column"] for m in metrics]
        )

        if code_column not in fieldnames:
            status = "MISSING_CODE_COLUMN"
            msg = f"数据源中找不到生产编号字段：{code_column}。当前字段：{', '.join(fieldnames)}"
            if not dry_run:
                write_collect_log(cur, item, status, msg, 0, 0, 0)
                update_item_status(cur, item_id, status)
                conn.commit()
            conn.close()
            return {
                "ok": False,
                "status": status,
                "message": msg,
                "fieldnames": fieldnames,
                "used_encoding": used_encoding
            }

        if process_column and process_column not in fieldnames:
            status = "MISSING_PROCESS_COLUMN"
            msg = f"数据源中找不到工序字段：{process_column}。当前字段：{', '.join(fieldnames)}"
            if not dry_run:
                write_collect_log(cur, item, status, msg, 0, 0, 0)
                update_item_status(cur, item_id, status)
                conn.commit()
            conn.close()
            return {
                "ok": False,
                "status": status,
                "message": msg,
                "fieldnames": fieldnames,
                "used_encoding": used_encoding
            }

        missing_metric_columns = [m["source_column"] for m in metrics if m["source_column"] not in fieldnames]
        if missing_metric_columns:
            status = "MISSING_METRIC_COLUMN"
            msg = f"数据源中找不到指标字段：{', '.join(missing_metric_columns)}。当前字段：{', '.join(fieldnames)}"
            if not dry_run:
                write_collect_log(cur, item, status, msg, 0, 0, 0)
                update_item_status(cur, item_id, status)
                conn.commit()
            conn.close()
            return {
                "ok": False,
                "status": status,
                "message": msg,
                "fieldnames": fieldnames,
                "missing_metric_columns": missing_metric_columns,
                "used_encoding": used_encoding
            }

        unknown_production_rows = 0
        if image_dynamic_production:
            production_rows = cur.execute("SELECT id, production_code FROM production_config").fetchall()
            production_lookup = {str(r["production_code"]).strip(): r["id"] for r in production_rows}
            matched = []
            for row in rows:
                row_code = str(row.get(code_column, "")).strip()
                if row_code and row_code in production_lookup:
                    row["_matched_production_code"] = row_code
                    row["_matched_production_id"] = production_lookup[row_code]
                    matched.append(row)
                else:
                    unknown_production_rows += 1
        else:
            matched = [r for r in rows if str(r.get(code_column, "")).strip() == str(production_code).strip()]
        if not matched:
            status = "NO_MATCHED_PRODUCTION_CODE"
            if image_dynamic_production:
                seen_codes = sorted({str(r.get(code_column, "")).strip() for r in rows if str(r.get(code_column, "")).strip()})
                msg = f"图片中未识别到已配置的生产编号。识别到的编号：{', '.join(seen_codes) if seen_codes else '无'}。"
            else:
                msg = f"未找到生产编号 {production_code} 对应的数据行。"
            if not dry_run:
                write_collect_log(cur, item, status, msg, 0, 0, 0)
                update_item_status(cur, item_id, status)
                conn.commit()
            conn.close()
            return {
                "ok": False,
                "status": status,
                "message": msg,
                "fieldnames": fieldnames,
                "used_encoding": used_encoding,
                "matched_rows": 0,
                "unknown_production_rows": unknown_production_rows
            }

        blank_process_rows = 0
        # 配置工序字段后，同一生产编号可对应多行不同工序，逐行采集。
        if process_column:
            target_rows = [r for r in matched if str(r.get(process_column, "")).strip()]
            blank_process_rows = len(matched) - len(target_rows)
            if not target_rows:
                status = "NO_VALID_PROCESS_ROWS"
                msg = f"生产编号 {production_code} 匹配到 {len(matched)} 行，但工序字段 {process_column} 均为空，已拒绝采集。"
                if not dry_run:
                    write_collect_log(cur, item, status, msg, len(matched), 0, 0)
                    update_item_status(cur, item_id, status)
                    conn.commit()
                conn.close()
                return {
                    "ok": False,
                    "status": status,
                    "message": msg,
                    "fieldnames": fieldnames,
                    "used_encoding": used_encoding,
                    "matched_rows": len(matched),
                    "collect_rows": 0
                }
        else:
            if len(matched) > 1 and not (item["process_step"] or "").strip():
                status = "MULTIPLE_ROWS_REQUIRE_PROCESS_COLUMN"
                msg = f"生产编号 {production_code} 匹配到 {len(matched)} 行，但量测项未配置工序字段。请在量测项中填写“工序字段名”，否则系统无法判断每行属于哪道工序。"
                if not dry_run:
                    write_collect_log(cur, item, status, msg, len(matched), 0, 0)
                    update_item_status(cur, item_id, status)
                    conn.commit()
                conn.close()
                return {
                    "ok": False,
                    "status": status,
                    "message": msg,
                    "fieldnames": fieldnames,
                    "used_encoding": used_encoding,
                    "matched_rows": len(matched),
                    "collect_rows": 0
                }
            target_rows = [matched[-1]]
        preview_rows = []
        for row in target_rows[:20]:
            row_process_step = str(row.get(process_column, "")).strip() if process_column else ""
            preview_rows.append({
                "production_code": str(row.get(code_column, "")).strip() or production_code,
                "process_step": row_process_step or item["process_step"] or "",
                "metrics": {m["metric_name"]: row.get(m["source_column"]) for m in metrics},
                "row": row
            })
        preview = preview_rows[0]["metrics"] if preview_rows else {}
        image_ocr_details = image_ocr_metric_details(target_rows, metrics, code_column, process_column, production_code, fixed_process_step)
        if dry_run:
            conn.close()
            return {
                "ok": True,
                "status": "TEST_SUCCESS",
                "message": "测试读取成功。",
                "fieldnames": fieldnames,
                "used_encoding": used_encoding,
                "matched_rows": len(matched),
                "selected_row": target_rows[-1],
                "selected_rows": target_rows[:20],
                "dynamic_production_routing": image_dynamic_production,
                "unknown_production_rows": unknown_production_rows,
                "process_step_column": process_column,
                "collect_rows": len(target_rows),
                "blank_process_rows_skipped": blank_process_rows,
                "metric_preview": preview,
                "row_previews": preview_rows,
                "image_ocr": [r.get("_ocr") for r in target_rows if isinstance(r, dict) and r.get("_ocr")],
                "image_ocr_details": image_ocr_details
            }

        for target_row in target_rows:
            base_hash = row_hash(target_row)
            actual_source_path = target_row.get("_source_path") or data_source_path
            row_process_step = str(target_row.get(process_column, "")).strip() if process_column else ""
            effective_process_step = row_process_step or item["process_step"] or ""
            effective_production_code = target_row.get("_matched_production_code") if image_dynamic_production else production_code
            effective_production_code = str(effective_production_code or production_code).strip()
            effective_production_id = target_row.get("_matched_production_id") if image_dynamic_production else item["production_id"]
            effective_production_id = effective_production_id or item["production_id"]
            for m in metrics:
                source_col = m["source_column"]
                value_text = "" if target_row.get(source_col) is None else str(target_row.get(source_col)).strip()
                value_number = safe_float(value_text) if m["data_type"] == "number" else None
                status = judge_status(value_number, m["lsl"], m["usl"], m["lcl"], m["ucl"])
                source_file_path, source_file_mtime, ocr_raw_text, ocr_roi_json, ocr_confidence, ocr_config_json = result_trace_fields(
                    target_row,
                    source_col,
                    actual_source_path,
                    item["image_parse_config_json"] if "image_parse_config_json" in item.keys() else ""
                )
                metric_hash = hashlib.sha256(
                    f"{item_id}|{m['id']}|{base_hash}|{m['metric_name']}|{effective_process_step}|{value_text}".encode("utf-8")
                ).hexdigest()

                try:
                    cur.execute("""
                    INSERT INTO measurement_result (
                        production_id, production_code, item_id, measurement_item_name,
                        process_step, execution_time_text, equipment_name,
                        metric_name, metric_value_text, metric_value_number, unit,
                        target, lsl, usl, lcl, ucl, result_status,
                        source_path, source_row_hash, source_metric_hash, collect_time, raw_row_json,
                        source_file_path, source_file_mtime, ocr_raw_text, ocr_roi_json, ocr_confidence, ocr_config_json
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        effective_production_id, effective_production_code, item_id, item["item_name"],
                        effective_process_step, item["execution_time_text"], item["equipment_name"],
                        m["metric_name"], value_text, value_number, m["unit"],
                        m["target"], m["lsl"], m["usl"], m["lcl"], m["ucl"], status,
                        actual_source_path, base_hash, metric_hash, now_str(),
                        json.dumps(target_row, ensure_ascii=False),
                        source_file_path, source_file_mtime, ocr_raw_text, ocr_roi_json, ocr_confidence, ocr_config_json
                    ))
                    inserted += 1
                except sqlite3.IntegrityError:
                    skipped += 1

        log_status = "SUCCESS"
        extra = f"，跳过空工序 {blank_process_rows} 行" if blank_process_rows else ""
        if unknown_production_rows:
            extra += f"，跳过未配置生产编号图片 {unknown_production_rows} 张"
        msg = f"采集成功：匹配 {len(matched)} 行，采集 {len(target_rows)} 行，新增 {inserted} 条，跳过重复 {skipped} 条{extra}。"
        write_collect_log(cur, item, log_status, msg, len(matched), inserted, skipped)
        update_item_status(cur, item_id, log_status)
        conn.commit()
        conn.close()
        return {
            "ok": True,
            "status": log_status,
            "message": msg,
            "matched_rows": len(matched),
            "collect_rows": len(target_rows),
            "inserted": inserted,
            "skipped": skipped,
            "selected_row": target_rows[-1],
            "selected_rows": target_rows[:20],
            "dynamic_production_routing": image_dynamic_production,
            "unknown_production_rows": unknown_production_rows,
            "process_step_column": process_column,
            "blank_process_rows_skipped": blank_process_rows,
            "metric_preview": preview,
            "row_previews": preview_rows,
            "image_ocr": [r.get("_ocr") for r in target_rows if isinstance(r, dict) and r.get("_ocr")],
            "image_ocr_details": image_ocr_details
        }

    except Exception as ex:
        status = "READ_ERROR"
        msg = f"读取失败：{ex}"
        if not dry_run:
            try:
                write_collect_log(cur, item, status, msg, 0, inserted, skipped)
                update_item_status(cur, item_id, status)
                conn.commit()
            except Exception:
                pass
        conn.close()
        return {
            "ok": False,
            "status": status,
            "message": msg,
            "traceback": traceback.format_exc()
        }


def _write_timeout_log(item_id: int, dry_run: bool, message: str):
    if dry_run:
        return
    try:
        conn = get_conn()
        cur = conn.cursor()
        item = cur.execute("""
            SELECT mi.*, p.production_code, p.production_name
            FROM measurement_item_config mi
            JOIN production_config p ON p.id = mi.production_id
            WHERE mi.id=?
        """, (item_id,)).fetchone()
        if item:
            write_collect_log(cur, item, "READ_TIMEOUT", message, 0, 0, 0)
            update_item_status(cur, item_id, "READ_TIMEOUT")
            conn.commit()
        conn.close()
    except Exception as ex:
        print("[timeout_log_failed]", ex)


def _clear_inflight(item_id: int):
    with INFLIGHT_LOCK:
        INFLIGHT_ITEMS.discard(item_id)


def collect_item_with_timeout_core(item_id: int, dry_run=False):
    """Protect UI requests from hanging on slow/broken UNC paths.

    Note: Python cannot forcibly kill a blocked OS-level UNC read inside a thread.
    This returns control to the UI after READ_TIMEOUT_SECONDS; the underlying read may
    finish later. For strict industrial isolation, move collectors into separate worker
    processes/services.

    Because a hung read keeps occupying a pool worker until the OS call returns, we guard
    real (non-dry-run) collections with an in-flight set: while a previous read for the
    same item is still blocked, we skip submitting another one. This keeps a single dead
    path from accumulating duplicate jobs and exhausting READ_EXECUTOR, which would
    otherwise stall collection for every item.
    """
    if not dry_run:
        with INFLIGHT_LOCK:
            if item_id in INFLIGHT_ITEMS:
                return {
                    "ok": False,
                    "status": "READ_IN_PROGRESS",
                    "message": "上一次采集仍在进行中（数据源路径可能无响应），本次已跳过，避免读取任务堆积、占满采集线程池。",
                    "matched_rows": 0,
                    "inserted": 0,
                    "skipped": 0
                }
            INFLIGHT_ITEMS.add(item_id)

    try:
        future = READ_EXECUTOR.submit(collect_item_core, item_id, dry_run)
    except Exception:
        if not dry_run:
            _clear_inflight(item_id)
        raise

    if not dry_run:
        # Clear the in-flight flag only when the underlying read truly finishes, even if
        # we already stopped waiting for it below after the timeout.
        future.add_done_callback(lambda _f: _clear_inflight(item_id))

    try:
        return future.result(timeout=READ_TIMEOUT_SECONDS)
    except FutureTimeoutError:
        msg = f"读取超时：超过 {READ_TIMEOUT_SECONDS} 秒未返回。可能是共享路径网络异常、权限问题或设备正在独占写入。系统已放弃本次前台等待，后台后续周期会重试。"
        _write_timeout_log(item_id, dry_run, msg)
        return {
            "ok": False,
            "status": "READ_TIMEOUT",
            "message": msg,
            "matched_rows": 0,
            "inserted": 0,
            "skipped": 0
        }


class CollectorService:
    """Single entry point for manual, scheduler and test collection jobs."""

    def create_job(self, item_id: int, trigger_type="manual", retry_count=0):
        conn = get_conn()
        cur = conn.cursor()
        item = cur.execute("""
            SELECT mi.*, p.production_code, p.production_name
            FROM measurement_item_config mi
            JOIN production_config p ON p.id=mi.production_id
            WHERE mi.id=?
        """, (item_id,)).fetchone()
        if not item:
            conn.close()
            raise ValueError("量测项不存在，无法创建采集任务")
        job_uuid = str(uuid.uuid4())
        ts = now_str()
        cur.execute("""
            INSERT INTO collect_job (
                job_uuid, item_id, production_id, production_code, measurement_item_name,
                data_source_type, data_source_path, trigger_type, status, scheduled_at,
                retry_count, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?, ?, ?)
        """, (
            job_uuid, item["id"], item["production_id"], item["production_code"], item["item_name"],
            item["data_source_type"] if "data_source_type" in item.keys() else "auto",
            item["data_source_path"], trigger_type, ts, retry_count, ts, ts
        ))
        job_id = cur.lastrowid
        conn.commit()
        conn.close()
        return job_id

    def update_job_status(self, job_id: int, status: str, **fields):
        allowed = {
            "started_at", "finished_at", "duration_ms", "matched_rows", "collect_rows",
            "inserted_count", "skipped_count", "failed_count", "retry_count",
            "error_message", "traceback_text"
        }
        updates = {"status": status, "updated_at": now_str()}
        for key, value in fields.items():
            if key in allowed:
                updates[key] = value
        assignments = ", ".join(f"{key}=?" for key in updates)
        conn = get_conn()
        conn.execute(f"UPDATE collect_job SET {assignments} WHERE id=?", list(updates.values()) + [job_id])
        conn.commit()
        conn.close()

    def run_job(self, job_id: int, dry_run=False):
        conn = get_conn()
        job = conn.execute("SELECT * FROM collect_job WHERE id=?", (job_id,)).fetchone()
        if not job:
            conn.close()
            raise ValueError("采集任务不存在")
        running = conn.execute("""
            SELECT id FROM collect_job
            WHERE item_id=? AND status='running' AND id<>?
            ORDER BY started_at DESC, id DESC LIMIT 1
        """, (job["item_id"], job_id)).fetchone()
        started_at = now_str()
        start_tick = time.monotonic()
        if running:
            finished_at = now_str()
            message = f"量测项已有运行中的采集任务 #{running['id']}，本次任务跳过。"
            conn.execute("""
                UPDATE collect_job
                SET status='skipped', started_at=?, finished_at=?, duration_ms=0,
                    error_message=?, updated_at=?
                WHERE id=?
            """, (started_at, finished_at, message, finished_at, job_id))
            conn.commit()
            conn.close()
            return {
                "ok": False,
                "status": "READ_IN_PROGRESS",
                "job_status": "skipped",
                "job_id": job_id,
                "job_uuid": job["job_uuid"],
                "message": message,
                "matched_rows": 0,
                "collect_rows": 0,
                "inserted": 0,
                "skipped": 0
            }
        conn.execute("""
            UPDATE collect_job SET status='running', started_at=?, updated_at=? WHERE id=?
        """, (started_at, started_at, job_id))
        conn.commit()
        conn.close()

        try:
            result = collect_item_with_timeout_core(job["item_id"], dry_run=dry_run)
        except Exception as ex:
            result = {
                "ok": False,
                "status": "READ_ERROR",
                "message": str(ex),
                "traceback": traceback.format_exc()
            }

        finished_at = now_str()
        duration_ms = int((time.monotonic() - start_tick) * 1000)
        result_status = result.get("status") or ("SUCCESS" if result.get("ok") else "READ_ERROR")
        if result.get("ok"):
            job_status = "success"
        elif result_status == "READ_TIMEOUT":
            job_status = "timeout"
        elif result_status == "READ_IN_PROGRESS":
            job_status = "skipped"
        else:
            job_status = "failed"
        failed_count = 0 if job_status in ("success", "skipped") else 1
        self.update_job_status(
            job_id,
            job_status,
            finished_at=finished_at,
            duration_ms=duration_ms,
            matched_rows=safe_int(result.get("matched_rows", 0), 0),
            collect_rows=safe_int(result.get("collect_rows", result.get("matched_rows", 0)), 0),
            inserted_count=safe_int(result.get("inserted", result.get("inserted_count", 0)), 0),
            skipped_count=safe_int(result.get("skipped", result.get("skipped_count", 0)), 0),
            failed_count=failed_count,
            error_message="" if result.get("ok") else result.get("message", ""),
            traceback_text=result.get("traceback", "")
        )
        result["job_id"] = job_id
        result["job_uuid"] = job["job_uuid"]
        result["job_status"] = job_status
        result["duration_ms"] = duration_ms
        return result

    def collect_item(self, item_id: int, trigger_type="manual", dry_run=False):
        job_id = self.create_job(item_id, trigger_type=trigger_type)
        return self.run_job(job_id, dry_run=dry_run)

    def collect_item_test(self, item_id: int):
        return self.collect_item(item_id, trigger_type="test", dry_run=True)


COLLECTOR = CollectorService()


def collect_item(item_id: int, dry_run=False):
    return COLLECTOR.collect_item(item_id, trigger_type="test" if dry_run else "manual", dry_run=dry_run)


def collect_item_with_timeout(item_id: int, dry_run=False, trigger_type=None):
    trigger = trigger_type or ("test" if dry_run else "manual")
    return COLLECTOR.collect_item(item_id, trigger_type=trigger, dry_run=dry_run)


def write_collect_log(cur, item, status, message, matched_rows, inserted_count, skipped_count):
    cur.execute("""
    INSERT INTO collect_log (
        production_id, production_code, item_id, measurement_item_name,
        data_source_path, status, message, matched_rows, inserted_count, skipped_count, created_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        item["production_id"], item["production_code"], item["id"], item["item_name"],
        item["data_source_path"], status, message, matched_rows, inserted_count, skipped_count, now_str()
    ))


def update_item_status(cur, item_id, status):
    cur.execute("""
    UPDATE measurement_item_config
    SET last_collect_time=?, last_collect_status=?, updated_at=?
    WHERE id=?
    """, (now_str(), status, now_str(), item_id))


def scheduler_loop():
    """Simple polling scheduler for the single-file deployment."""
    last_run = {}
    while not SCHEDULER_STOP.is_set():
        try:
            conn = get_conn()
            rows = conn.execute("""
                SELECT id, scan_frequency_seconds, last_collect_time
                FROM measurement_item_config
                WHERE enabled=1
            """).fetchall()
            conn.close()

            now_ts = time.time()
            for r in rows:
                freq = max(10, int(r["scan_frequency_seconds"] or 60))
                item_id = r["id"]
                prev = last_run.get(item_id, 0)
                if now_ts - prev >= freq:
                    last_run[item_id] = now_ts
                    collect_item_with_timeout(item_id, dry_run=False, trigger_type="scheduler")
        except Exception as ex:
            print("[scheduler]", ex)
        SCHEDULER_STOP.wait(5)


# ==========================================================
# Auth
# ==========================================================

def current_user(handler):
    c = parse_cookie(handler.headers.get("Cookie"))
    sid = c.get("sid")
    if not sid or sid not in SESSIONS:
        return None
    return SESSIONS[sid]


# ==========================================================
# HTML layout
# ==========================================================

def base_layout(title, body, user=None):
    css = """
    :root{--bg:#f5f7fb;--card:#fff;--text:#172033;--muted:#667085;--line:#e5e7eb;--primary:#2563eb;--primary2:#1d4ed8;--danger:#dc2626;--ok:#16a34a;--warn:#ea580c;}
    *{box-sizing:border-box} body{margin:0;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','Microsoft YaHei',Arial,sans-serif;background:var(--bg);color:var(--text)}
    a{color:var(--primary);text-decoration:none}.topbar{height:56px;background:#111827;color:#fff;display:flex;align-items:center;justify-content:space-between;padding:0 22px}.topbar a{color:#bfdbfe}.brand{font-weight:800}.layout{display:flex;min-height:calc(100vh - 56px)}
    .sidebar{width:230px;background:white;border-right:1px solid var(--line);padding:16px 10px}.sidebar a{display:block;padding:11px 14px;border-radius:10px;color:#374151;margin-bottom:6px}.sidebar a:hover{background:#eff6ff;color:var(--primary2)}
    .content{flex:1;padding:24px;overflow:auto}.card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:18px;margin-bottom:18px;box-shadow:0 2px 10px rgba(15,23,42,.04)}
    h1{margin:0 0 18px;font-size:24px} h2{margin:0 0 12px;font-size:18px}.grid{display:grid;grid-template-columns:repeat(4,minmax(160px,1fr));gap:16px}.metric .label{color:var(--muted);font-size:13px}.metric .value{font-size:28px;font-weight:800;margin-top:8px}
    table{width:100%;border-collapse:collapse;font-size:14px} th,td{border-bottom:1px solid var(--line);padding:10px 8px;text-align:left;white-space:nowrap} th{background:#f9fafb;color:#344054}.table-wrap{overflow:auto}
    input,select,textarea{border:1px solid #d0d5dd;border-radius:10px;padding:9px 10px;min-width:180px;background:#fff} textarea{font-family:ui-monospace,Consolas,monospace}.form-row{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-bottom:12px}.form-grid{display:grid;grid-template-columns:180px 1fr;gap:12px;align-items:center;max-width:950px}
    button,.btn{background:var(--primary);border:none;color:#fff;padding:9px 14px;border-radius:10px;font-weight:700;cursor:pointer;display:inline-block}.btn.secondary{background:#475467}.btn.danger,button.danger{background:var(--danger)}button:hover,.btn:hover{background:var(--primary2)}button.danger:hover,.btn.danger:hover{background:#991b1b}.inline-form{display:inline}.inline-form button{margin:0}
    .badge{display:inline-block;padding:3px 8px;border-radius:999px;font-size:12px;font-weight:700}.enabled,.SUCCESS,.PASS,.MS3_PASS,.TEST_SUCCESS,.success{background:#dcfce7;color:#166534}.running,.pending{background:#dbeafe;color:#1d4ed8}.skipped{background:#f1f5f9;color:#475569}.disabled,.failed,.timeout,.READ_ERROR,.READ_TIMEOUT,.OOS,.MISS_MS2,.MISSING_CODE_COLUMN,.MISSING_PROCESS_COLUMN,.MISSING_METRIC_COLUMN,.NO_VALID_PROCESS_ROWS,.MULTIPLE_ROWS_REQUIRE_PROCESS_COLUMN,.PROCESS_STEP_REQUIRED{background:#fee2e2;color:#991b1b}.NO_MATCHED_PRODUCTION_CODE,.OOC,.MS2_PASS{background:#ffedd5;color:#9a3412}.TEXT{background:#e0e7ff;color:#3730a3}
    .note{font-size:13px;color:var(--muted);line-height:1.7}.error{color:var(--danger);font-size:14px}.success{color:var(--ok);font-size:14px}.login-wrap{min-height:100vh;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,#0f172a,#1d4ed8)}.login-card{width:390px;background:white;border-radius:18px;padding:28px;box-shadow:0 18px 60px rgba(0,0,0,.25)}.login-card input,.login-card button{width:100%;margin:8px 0}.login-card p{color:var(--muted)}
    pre{background:#0b1020;color:#e5e7eb;padding:14px;border-radius:12px;overflow:auto}.actions{display:flex;gap:8px;flex-wrap:wrap}.small{font-size:12px;color:var(--muted)}
    .dash-grid{display:grid;grid-template-columns:1.15fr .85fr;gap:18px}.chart{width:100%;min-height:220px}.stack{height:12px;background:#eef2f7;border-radius:999px;overflow:hidden;display:flex}.seg-pass{background:#16a34a}.seg-ooc{background:#ea580c}.seg-oos{background:#dc2626}.bar-cell{min-width:180px}.bar-bg{height:10px;background:#eef2f7;border-radius:999px;overflow:hidden}.bar-fill{height:10px;background:#2563eb;border-radius:999px}.risk-fill{background:#dc2626}.warn-fill{background:#ea580c}.muted-fill{background:#64748b}
    .quick-start{border-color:#bfdbfe;background:linear-gradient(180deg,#eff6ff,#fff)}.step-list{display:grid;grid-template-columns:repeat(4,minmax(170px,1fr));gap:12px;margin-top:12px}.step{border:1px solid var(--line);border-radius:12px;background:#fff;padding:12px}.step b{display:block;margin-bottom:6px}.step .num{display:inline-block;width:24px;height:24px;line-height:24px;text-align:center;border-radius:999px;background:#dbeafe;color:#1d4ed8;font-weight:800;margin-right:6px}.demo-actions{display:flex;gap:10px;flex-wrap:wrap;margin-top:14px}.status-strip{display:flex;gap:10px;flex-wrap:wrap}.status-strip span{display:inline-block;border:1px solid var(--line);border-radius:12px;background:#fff;padding:9px 12px;font-weight:700}
    details.process-risk{border:1px solid var(--line);border-radius:12px;margin-bottom:10px;background:#fff}details.process-risk summary{cursor:pointer;padding:12px 14px;font-weight:700;display:flex;gap:14px;align-items:center;justify-content:space-between}.risk-meta{font-weight:500;color:var(--muted);font-size:13px}.risk-inner{padding:0 14px 14px}.link-button{background:none;border:none;color:var(--primary);padding:0;font-weight:700;cursor:pointer}.modal-backdrop{position:fixed;inset:0;background:rgba(15,23,42,.55);display:none;align-items:center;justify-content:center;padding:24px;z-index:20}.modal-card{background:#fff;border-radius:16px;max-width:920px;width:100%;padding:18px;box-shadow:0 24px 80px rgba(0,0,0,.32)}.modal-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}
    @media(max-width:900px){.grid{grid-template-columns:repeat(2,1fr)}.dash-grid{grid-template-columns:1fr}.step-list{grid-template-columns:1fr}.layout{flex-direction:column}.sidebar{width:100%;display:flex;overflow:auto}.sidebar a{white-space:nowrap}.form-grid{grid-template-columns:1fr}}
    """
    if user:
        users_link = '<a href="/users">用户管理</a>' if can_manage_users(user) else ""
        audit_link = '<a href="/audit_logs">审计日志</a>' if can_manage_system(user) else ""
        import_link = '<a href="/import_config">导入配置</a>' if can_manage_config(user) else ""
        shell = f"""
        <div class="topbar"><div class="brand">{h(APP_TITLE)}</div><div>版本：{APP_VERSION} ｜ 用户：{h(user.get('username'))} ｜ 角色：{h(user_role(user))} ｜ <a href="/logout">退出</a></div></div>
        <div class="layout">
          <aside class="sidebar">
            <a href="/">首页 Dashboard</a>
            <a href="/productions">生产编号管理</a>
            <a href="/templates">模板库</a>
            <a href="/results">采集结果</a>
            <a href="/collect_jobs">采集任务</a>
            <a href="/logs">采集日志</a>
            <a href="/config_checks">配置检查</a>
            {audit_link}
            {import_link}
            {users_link}
            <a href="/about">说明</a>
          </aside>
          <main class="content">{body}</main>
        </div>
        """
    else:
        shell = body
    return f"""<!doctype html><html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>{h(title)}</title><style>{css}</style></head><body>{shell}</body></html>"""


def display_status(text):
    return {
        "MS3_PASS": "MS3达成",
        "MS2_PASS": "MS2达成",
        "MISS_MS2": "未达MS2",
        "PASS": "PASS",
        "OOC": "OOC",
        "OOS": "OOS",
        "TEXT": "TEXT",
        "VOIDED": "已作废",
    }.get("" if text is None else str(text), "" if text is None else str(text))


def badge(text):
    return f'<span class="badge {h(text)}">{h(display_status(text))}</span>'


def page_login(error=""):
    return base_layout("登录", f"""
    <div class="login-wrap">
      <form class="login-card" method="post" action="/login">
        <h1>{h(APP_TITLE)}</h1>
        <p>管理员登录</p>
        {f'<div class="error">{h(error)}</div>' if error else ''}
        <input name="username" placeholder="管理员账号" required autocomplete="username">
        <input name="password" type="password" placeholder="管理员密码" required autocomplete="current-password">
        <button type="submit">登录</button>
        <div class="note" style="margin-top:12px">默认账号：admin<br>默认密码：admin123<br>正式使用建议通过环境变量 MDCP_ADMIN_USERNAME / MDCP_ADMIN_PASSWORD 设置管理员账号密码。</div>
      </form>
    </div>
    """)


def page_demo_ready(user, result):
    collect_rows = "".join(
        f"<tr><td>{h(r.get('job_id'))}</td><td>{badge(r.get('job_status'))}</td><td>{h(r.get('matched_rows'))}</td><td>{h(r.get('inserted'))}</td><td>{h(r.get('skipped'))}</td><td>{h(r.get('message'))}</td></tr>"
        for r in result.get("collect_results", [])
    ) or "<tr><td colspan='6'>暂无采集任务结果</td></tr>"
    return base_layout("Demo 数据已加载", f"""
    <h1>Demo 数据已加载</h1>
    <div class="card quick-start">
      <h2>客户演示数据准备完成</h2>
      <div class="status-strip">
        <span>结果 {result.get('total', 0)} 条</span>
        <span>MS3 {result.get('ms3_count', 0)}</span>
        <span>仅 MS2 {result.get('ms2_count', 0)}</span>
        <span>未达 MS2 {result.get('miss_count', 0)}</span>
        <span>本次新增 {result.get('inserted', 0)}</span>
      </div>
      <p class="note">Demo CSV：{h(result.get('csv_path'))}。重复加载会跳过已入库的重复结果，便于多次演示。</p>
      <div class="demo-actions">
        <a class="btn" href="/">查看 Dashboard</a>
        <a class="btn secondary" href="/results?production_code=DEMO_WAFER">查看采集结果</a>
        <a class="btn secondary" href="/collect_jobs">查看采集任务</a>
        <a class="btn secondary" href="/templates">查看模板库</a>
      </div>
    </div>
    <div class="card">
      <h2>本次采集任务</h2>
      <div class="table-wrap"><table><tr><th>任务ID</th><th>状态</th><th>匹配行</th><th>新增</th><th>跳过</th><th>说明</th></tr>{collect_rows}</table></div>
    </div>
    """, user)


# ==========================================================
# Pages
# ==========================================================

def percent(part, total):
    return 0.0 if not total else round(part * 100.0 / total, 1)


def table_bar(value, max_value, css_class="bar-fill"):
    width = 0 if not max_value else max(2, min(100, int(value * 100 / max_value)))
    return f'<div class="bar-bg"><div class="{css_class}" style="width:{width}%"></div></div>'


def parse_date_value(text, default_value):
    try:
        return datetime.strptime(text or "", "%Y-%m-%d").date()
    except Exception:
        return default_value


def dashboard_date_range(query):
    today = datetime.now(APP_TZ).date()
    start_date = parse_date_value((query or {}).get("start_date", [""])[0], today)
    end_date = parse_date_value((query or {}).get("end_date", [""])[0], today)
    if start_date > end_date:
        start_date, end_date = end_date, start_date
    return start_date, end_date, f"{start_date} 00:00:00", f"{end_date} 23:59:59"


def status_stack(ms3_count, ms2_only_count, miss_count, total):
    if total <= 0:
        return '<div class="stack"></div>'
    pass_w = percent(ms3_count, total)
    ooc_w = percent(ms2_only_count, total)
    oos_w = percent(miss_count, total)
    return f"""
    <div class="stack" title="MS3达成 {pass_w}% / 仅MS2达成 {ooc_w}% / 未达MS2 {round(oos_w, 1)}%">
      <div class="seg-pass" style="width:{pass_w}%"></div>
      <div class="seg-ooc" style="width:{ooc_w}%"></div>
      <div class="seg-oos" style="width:{oos_w}%"></div>
    </div>
    """


def svg_status_pie(ms3_count, ms2_only_count, miss_count, total):
    """状态分布饼图：MS3达成 / 仅达MS2 / 未达MS2。

    使用实心扇区 path，不再使用圆环/堆叠条，便于一眼看比例。
    """
    cx, cy, r = 130, 120, 82
    slices = [
        ("MS3达成", ms3_count or 0, "#16a34a"),
        ("仅达MS2", ms2_only_count or 0, "#ea580c"),
        ("未达MS2", miss_count or 0, "#dc2626"),
    ]
    if total <= 0:
        return """
        <svg class="chart" viewBox="0 0 520 260" role="img" aria-label="状态分布饼图">
          <circle cx="130" cy="120" r="82" fill="#eef2f7"/>
          <text x="130" y="116" text-anchor="middle" font-size="24" font-weight="800" fill="#172033">0</text>
          <text x="130" y="140" text-anchor="middle" font-size="12" fill="#667085">结果数</text>
          <rect x="280" y="68" width="12" height="12" rx="2" fill="#16a34a"/><text x="300" y="78" font-size="13" fill="#344054">MS3达成：0（0.0%）</text>
          <rect x="280" y="100" width="12" height="12" rx="2" fill="#ea580c"/><text x="300" y="110" font-size="13" fill="#344054">仅达MS2：0（0.0%）</text>
          <rect x="280" y="132" width="12" height="12" rx="2" fill="#dc2626"/><text x="300" y="142" font-size="13" fill="#344054">未达MS2：0（0.0%）</text>
        </svg>
        """

    def polar_to_xy(angle_rad):
        return cx + r * math.cos(angle_rad), cy + r * math.sin(angle_rad)

    paths = []
    start_angle = -math.pi / 2
    for label, count, color in slices:
        if count <= 0:
            continue
        angle = 2 * math.pi * count / total
        end_angle = start_angle + angle
        x1, y1 = polar_to_xy(start_angle)
        x2, y2 = polar_to_xy(end_angle)
        large_arc = 1 if angle > math.pi else 0

        if count == total:
            paths.append(f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="{color}"><title>{h(label)}：{count}（{percent(count,total)}%）</title></circle>')
        else:
            d = f"M {cx} {cy} L {x1:.2f} {y1:.2f} A {r} {r} 0 {large_arc} 1 {x2:.2f} {y2:.2f} Z"
            paths.append(f'<path d="{d}" fill="{color}" stroke="#ffffff" stroke-width="2"><title>{h(label)}：{count}（{percent(count,total)}%）</title></path>')
        start_angle = end_angle

    legend = []
    y = 78
    for label, count, color in slices:
        legend.append(f'<rect x="280" y="{y - 10}" width="12" height="12" rx="2" fill="{color}"/>')
        legend.append(f'<text x="300" y="{y}" font-size="13" fill="#344054">{h(label)}：{count}（{percent(count, total)}%）</text>')
        y += 32

    return f"""
    <svg class="chart" viewBox="0 0 520 260" role="img" aria-label="状态分布饼图">
      {''.join(paths)}
      <circle cx="{cx}" cy="{cy}" r="34" fill="rgba(255,255,255,0.92)"/>
      <text x="{cx}" y="{cy - 4}" text-anchor="middle" font-size="22" font-weight="800" fill="#172033">{total}</text>
      <text x="{cx}" y="{cy + 18}" text-anchor="middle" font-size="12" fill="#667085">结果数</text>
      {''.join(legend)}
    </svg>
    """


def svg_bar_chart(day_rows):
    width, height = 760, 240
    pad_l, pad_r, pad_t, pad_b = 46, 18, 20, 42
    chart_w = width - pad_l - pad_r
    chart_h = height - pad_t - pad_b
    max_total = max([r["total"] for r in day_rows] + [1])
    bar_group = chart_w / max(1, len(day_rows))
    parts = [
        f'<svg class="chart" viewBox="0 0 {width} {height}" role="img" aria-label="最近7天采集状态趋势">',
        f'<line x1="{pad_l}" y1="{pad_t + chart_h}" x2="{width - pad_r}" y2="{pad_t + chart_h}" stroke="#d0d5dd"/>'
    ]
    for idx, r in enumerate(day_rows):
        x = pad_l + idx * bar_group + bar_group * 0.18
        bar_w = bar_group * 0.64
        scale = chart_h / max_total
        oos_h = r["miss"] * scale
        ooc_h = r["ms2"] * scale
        pass_h = max(0, r["ms3"] * scale)
        y = pad_t + chart_h
        parts.append(f'<rect x="{x:.1f}" y="{y - pass_h:.1f}" width="{bar_w:.1f}" height="{pass_h:.1f}" fill="#16a34a" rx="3"/>')
        y -= pass_h
        parts.append(f'<rect x="{x:.1f}" y="{y - ooc_h:.1f}" width="{bar_w:.1f}" height="{ooc_h:.1f}" fill="#ea580c" rx="3"/>')
        y -= ooc_h
        parts.append(f'<rect x="{x:.1f}" y="{y - oos_h:.1f}" width="{bar_w:.1f}" height="{oos_h:.1f}" fill="#dc2626" rx="3"/>')
        label = r["day"][5:]
        parts.append(f'<text x="{x + bar_w / 2:.1f}" y="{height - 16}" text-anchor="middle" font-size="12" fill="#667085">{h(label)}</text>')
        parts.append(f'<text x="{x + bar_w / 2:.1f}" y="{max(14, y - oos_h - 6):.1f}" text-anchor="middle" font-size="12" fill="#344054">{r["total"]}</text>')
    parts.append('<text x="48" y="16" font-size="12" fill="#667085">MS3达成/仅MS2达成/未达MS2 堆叠数量</text>')
    parts.append('</svg>')
    return "".join(parts)


def sigma_expr():
    return "ROUND(CASE WHEN COUNT(metric_value_number) > 1 THEN sqrt(AVG(metric_value_number * metric_value_number) - AVG(metric_value_number) * AVG(metric_value_number)) ELSE 0 END, 4)"


def svg_line_chart(points, title, min_value=None, max_value=None, sigma_value=None):
    width, height = 820, 300
    pad_l, pad_r, pad_t, pad_b = 52, 24, 36, 48
    values = [p["value"] for p in points if p["value"] is not None]
    if not values:
        return f"<p class='note'>{h(title)} 暂无可绘制的数值数据。</p>"
    lo, hi = min(values), max(values)
    if lo == hi:
        lo -= 1
        hi += 1
    chart_w = width - pad_l - pad_r
    chart_h = height - pad_t - pad_b
    def xy(idx, val):
        x = pad_l + (idx * chart_w / max(1, len(values) - 1))
        y = pad_t + chart_h - ((val - lo) * chart_h / (hi - lo))
        return x, y
    coords = [xy(i, v) for i, v in enumerate(values)]
    poly = " ".join(f"{x:.1f},{y:.1f}" for x, y in coords)
    parts = [
        f'<svg class="chart" viewBox="0 0 {width} {height}" role="img" aria-label="{h(title)}">',
        f'<text x="{pad_l}" y="18" font-size="13" font-weight="700" fill="#172033">{h(title)}</text>',
        f'<text x="{pad_l}" y="34" font-size="12" fill="#667085">max={h(max_value)} min={h(min_value)} sigma={h(sigma_value)}</text>',
        f'<line x1="{pad_l}" y1="{pad_t + chart_h}" x2="{width - pad_r}" y2="{pad_t + chart_h}" stroke="#d0d5dd"/>',
        f'<line x1="{pad_l}" y1="{pad_t}" x2="{pad_l}" y2="{pad_t + chart_h}" stroke="#d0d5dd"/>',
        f'<polyline fill="none" stroke="#2563eb" stroke-width="3" points="{poly}"/>'
    ]
    for x, y in coords:
        parts.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3" fill="#2563eb"/>')
    parts.append(f'<text x="{pad_l}" y="{height - 16}" font-size="12" fill="#667085">{h(points[0]["time"][:10])}</text>')
    parts.append(f'<text x="{width - pad_r}" y="{height - 16}" text-anchor="end" font-size="12" fill="#667085">{h(points[-1]["time"][:10])}</text>')
    parts.append('</svg>')
    return "".join(parts)


def config_check_issues(limit=200):
    issues = []
    conn = get_conn()

    def add(issue, production_code="", item_name="", item_id=None, template_name="", template_id=None, suggestion=""):
        issues.append({
            "issue": issue,
            "production_code": production_code or "",
            "item_name": item_name or template_name or "",
            "id": item_id or "",
            "template_id": template_id or "",
            "suggestion": suggestion,
        })

    for r in conn.execute("""
        SELECT mi.id, p.production_code, mi.item_name
        FROM measurement_item_config mi JOIN production_config p ON p.id=mi.production_id
        WHERE COALESCE(mi.process_step,'')='' AND COALESCE(mi.process_step_column,'')=''
        ORDER BY mi.id DESC
    """).fetchall():
        add("无工序配置", r["production_code"], r["item_name"], r["id"], suggestion="填写固定量测工序，或填写工序字段名。")

    for r in conn.execute("""
        SELECT mi.id, p.production_code, mi.item_name
        FROM measurement_item_config mi JOIN production_config p ON p.id=mi.production_id
        WHERE NOT EXISTS (SELECT 1 FROM metric_config mc WHERE mc.item_id=mi.id AND mc.enabled=1)
        ORDER BY mi.id DESC
    """).fetchall():
        add("无启用指标", r["production_code"], r["item_name"], r["id"], suggestion="进入指标配置，至少启用一个指标。")

    for r in conn.execute("""
        SELECT mi.id, p.production_code, mi.item_name
        FROM measurement_item_config mi JOIN production_config p ON p.id=mi.production_id
        WHERE COALESCE(mi.data_source_path,'')=''
        ORDER BY mi.id DESC
    """).fetchall():
        add("数据源路径为空", r["production_code"], r["item_name"], r["id"], suggestion="填写 CSV/Excel/图片共享路径。")

    excel_rows = conn.execute("""
        SELECT mi.id, p.production_code, mi.item_name, mi.data_source_type, mi.data_source_path, mi.excel_sheet_name
        FROM measurement_item_config mi JOIN production_config p ON p.id=mi.production_id
        WHERE LOWER(COALESCE(mi.data_source_type,'auto')) IN ('excel','auto')
        ORDER BY mi.id DESC
    """).fetchall()
    for r in excel_rows:
        suffix = Path(r["data_source_path"] or "").suffix.lower()
        if suffix in (".xlsx", ".xlsm", ".xls") and not (r["excel_sheet_name"] or "").strip():
            add("Excel Sheet 未填写", r["production_code"], r["item_name"], r["id"], suggestion="多 Sheet 文件建议明确填写 Excel Sheet 名称。")

    image_rows = conn.execute("""
        SELECT mi.id, p.production_code, mi.item_name, mi.image_parse_config_json
        FROM measurement_item_config mi JOIN production_config p ON p.id=mi.production_id
        WHERE LOWER(COALESCE(mi.data_source_type,''))='image'
        ORDER BY mi.id DESC
    """).fetchall()
    for r in image_rows:
        config_text = r["image_parse_config_json"] or ""
        if not config_text.strip():
            add("Image OCR 配置问题", r["production_code"], r["item_name"], r["id"], suggestion="填写 Image OCR config JSON。")
            continue
        try:
            json.loads(config_text)
        except Exception as ex:
            add("Image OCR 配置问题", r["production_code"], r["item_name"], r["id"], suggestion=f"修正 JSON 格式：{ex}")

    for r in conn.execute("""
        SELECT t.id, t.template_name
        FROM template_config t
        WHERE NOT EXISTS (SELECT 1 FROM template_metric_config tm WHERE tm.template_id=t.id)
        ORDER BY t.id DESC
    """).fetchall():
        add("模板没有指标", template_name=r["template_name"], template_id=r["id"], suggestion="编辑模板，至少保留一个指标。")

    for r in conn.execute("""
        SELECT id, template_name
        FROM template_config
        WHERE COALESCE(production_code_column,'')=''
        ORDER BY id DESC
    """).fetchall():
        add("模板缺少生产编号字段", template_name=r["template_name"], template_id=r["id"], suggestion="编辑模板，选择生产编号字段。")

    conn.close()
    summary = {}
    for issue in issues:
        summary[issue["issue"]] = summary.get(issue["issue"], 0) + 1
    return issues[:limit], summary, len(issues)


def page_dashboard(user, query=None):
    query = query or {}
    start_date, end_date, start_dt, end_dt = dashboard_date_range(query)
    selected_production = query.get("production_code", [""])[0].strip()
    conn = get_conn()
    p_count = conn.execute("SELECT COUNT(*) AS c FROM production_config").fetchone()["c"]
    item_count = conn.execute("SELECT COUNT(*) AS c FROM measurement_item_config WHERE enabled=1").fetchone()["c"]
    productions = conn.execute("SELECT production_code FROM production_config ORDER BY production_code").fetchall()
    production_options = '<option value="">当前全部生产编号</option>' + "".join(
        f'<option value="{h(r["production_code"])}" {"selected" if selected_production == r["production_code"] else ""}>{h(r["production_code"])}</option>'
        for r in productions
    )
    result_filter = "collect_time BETWEEN ? AND ? AND COALESCE(is_voided,0)=0"
    result_params = [start_dt, end_dt]
    if selected_production:
        result_filter += " AND production_code=?"
        result_params.append(selected_production)
    else:
        result_filter += " AND production_code IN (SELECT production_code FROM production_config)"
    result_filter += """
        AND EXISTS (
            SELECT 1
            FROM measurement_item_config mi
            WHERE mi.id=measurement_result.item_id
        )
    """
    status_stats = conn.execute("""
        SELECT COUNT(*) AS total,
               SUM(CASE WHEN result_status IN ('MS3_PASS','PASS') THEN 1 ELSE 0 END) AS ms3_count,
               SUM(CASE WHEN result_status IN ('MS2_PASS','OOC') THEN 1 ELSE 0 END) AS ms2_count,
               SUM(CASE WHEN result_status IN ('MISS_MS2','OOS') THEN 1 ELSE 0 END) AS miss_count
        FROM measurement_result
        WHERE """ + result_filter, result_params).fetchone()
    last_collect = conn.execute("SELECT MAX(collect_time) AS t FROM measurement_result WHERE " + result_filter, result_params).fetchone()["t"] or "暂无"
    trend_rows = conn.execute("""
        SELECT substr(collect_time,1,10) AS day,
               COUNT(*) AS total,
               SUM(CASE WHEN result_status IN ('MS3_PASS','PASS') THEN 1 ELSE 0 END) AS ms3,
               SUM(CASE WHEN result_status IN ('MS2_PASS','OOC') THEN 1 ELSE 0 END) AS ms2,
               SUM(CASE WHEN result_status IN ('MISS_MS2','OOS') THEN 1 ELSE 0 END) AS miss
        FROM measurement_result
        WHERE """ + result_filter + """
        GROUP BY substr(collect_time,1,10)
    """, result_params).fetchall()
    trend_map = {r["day"]: {"total": r["total"], "ms3": r["ms3"] or 0, "ms2": r["ms2"] or 0, "miss": r["miss"] or 0} for r in trend_rows}
    day_rows = []
    days = min((end_date - start_date).days, 90)
    for offset in range(days + 1):
        day = (start_date + timedelta(days=offset)).strftime("%Y-%m-%d")
        day_rows.append({"day": day, **trend_map.get(day, {"total": 0, "ms3": 0, "ms2": 0, "miss": 0})})
    process_rows = conn.execute("""
        SELECT COALESCE(NULLIF(process_step,''),'未填工序') AS process_name,
               COUNT(*) AS total,
               SUM(CASE WHEN result_status IN ('MS3_PASS','PASS') THEN 1 ELSE 0 END) AS ms3_count,
               SUM(CASE WHEN result_status IN ('MS2_PASS','OOC') THEN 1 ELSE 0 END) AS ms2_count,
               SUM(CASE WHEN result_status IN ('MISS_MS2','OOS') THEN 1 ELSE 0 END) AS miss_count
        FROM measurement_result
        WHERE """ + result_filter + """
        GROUP BY COALESCE(NULLIF(process_step,''),'未填工序')
        ORDER BY miss_count DESC, ms2_count DESC, total DESC
        LIMIT 12
    """, result_params).fetchall()
    process_metric_map = {}
    trend_charts = {}
    for p_idx, process in enumerate(process_rows):
        process_name = process["process_name"]
        metric_rows = conn.execute("""
            SELECT metric_name,
                   COUNT(*) AS total,
                   ROUND(AVG(metric_value_number), 4) AS avg_value,
                   ROUND(MIN(metric_value_number), 4) AS min_value,
                   ROUND(MAX(metric_value_number), 4) AS max_value,
                   AVG(metric_value_number * metric_value_number) AS avg_square,
                   SUM(CASE WHEN result_status IN ('MS3_PASS','PASS') THEN 1 ELSE 0 END) AS ms3_count,
                   SUM(CASE WHEN result_status IN ('MS2_PASS','OOC') THEN 1 ELSE 0 END) AS ms2_count,
                   SUM(CASE WHEN result_status IN ('MISS_MS2','OOS') THEN 1 ELSE 0 END) AS miss_count
            FROM measurement_result
            WHERE """ + result_filter + """
              AND COALESCE(NULLIF(process_step,''),'未填工序')=?
              AND metric_name IS NOT NULL AND metric_name <> ''
            GROUP BY metric_name
            ORDER BY miss_count DESC, ms2_count DESC, total DESC
        """, result_params + [process_name]).fetchall()
        enriched = []
        for m_idx, metric in enumerate(metric_rows):
            sigma = ""
            if metric["avg_value"] is not None and metric["avg_square"] is not None:
                sigma = round(max(0, metric["avg_square"] - float(metric["avg_value"]) * float(metric["avg_value"])) ** 0.5, 4)
            chart_rows = conn.execute("""
                SELECT collect_time AS time, metric_value_number AS value
                FROM measurement_result
                WHERE """ + result_filter + """
                  AND COALESCE(NULLIF(process_step,''),'未填工序')=?
                  AND metric_name=?
                  AND metric_value_number IS NOT NULL
                ORDER BY collect_time ASC
                LIMIT 300
            """, result_params + [process_name, metric["metric_name"]]).fetchall()
            key = f"p{p_idx}_m{m_idx}"
            points = [{"time": r["time"], "value": r["value"]} for r in chart_rows]
            trend_charts[key] = svg_line_chart(points, f"{process_name} / {metric['metric_name']}", metric["min_value"], metric["max_value"], sigma)
            enriched.append((metric, sigma, key))
        process_metric_map[process_name] = enriched
    abnormal_rows = conn.execute("""
        SELECT * FROM measurement_result
        WHERE """ + result_filter + """
          AND result_status IN ('MISS_MS2','OOS')
        ORDER BY collect_time DESC
        LIMIT 10
    """, result_params).fetchall()

    recent_24 = (datetime.now(APP_TZ) - timedelta(hours=24)).strftime("%Y-%m-%d %H:%M:%S")
    job_stats = conn.execute("""
        SELECT
            SUM(CASE WHEN status='success' THEN 1 ELSE 0 END) AS success_count,
            SUM(CASE WHEN status='failed' THEN 1 ELSE 0 END) AS failed_count,
            SUM(CASE WHEN status='timeout' THEN 1 ELSE 0 END) AS timeout_count,
            SUM(CASE WHEN status='running' THEN 1 ELSE 0 END) AS running_count
        FROM collect_job
        WHERE COALESCE(created_at, scheduled_at, started_at)>=?
    """, (recent_24,)).fetchone()
    recent_failed_jobs = conn.execute("""
        SELECT id, production_code, measurement_item_name, data_source_path, trigger_type,
               status, started_at, finished_at, duration_ms, error_message
        FROM collect_job
        WHERE status IN ('failed','timeout')
        ORDER BY COALESCE(finished_at, started_at, created_at) DESC, id DESC
        LIMIT 5
    """).fetchall()
    failure_statuses = ("READ_ERROR", "PROCESS_STEP_REQUIRED", "NO_VALID_PROCESS_ROWS", "NO_METRICS_CONFIGURED", "OCR_ERROR", "IMAGE_ROUTE_SKIPPED", "IMAGE_NO_MATCH")
    failure_placeholders = ",".join("?" for _ in failure_statuses)
    failure_records = conn.execute(f"""
        SELECT data_source_path, status, message, created_at
        FROM collect_log
        WHERE created_at>=? AND status IN ({failure_placeholders})
        ORDER BY created_at DESC
    """, [recent_24] + list(failure_statuses)).fetchall()
    failure_map = {}
    for row in failure_records:
        key = row["data_source_path"] or "(空路径)"
        if key not in failure_map:
            failure_map[key] = {"data_source_path": key, "failure_count": 0, "last_failed_at": row["created_at"], "last_message": row["message"]}
        failure_map[key]["failure_count"] += 1
    failure_top = sorted(failure_map.values(), key=lambda r: (-r["failure_count"], r["last_failed_at"] or ""))[:5]

    no_data_rows = conn.execute("""
        SELECT p.production_code, mi.item_name, mi.equipment_name, mi.data_source_path, mi.last_collect_status, mi.last_collect_time
        FROM measurement_item_config mi
        JOIN production_config p ON p.id=mi.production_id
        WHERE mi.enabled=1
          AND NOT EXISTS (
            SELECT 1 FROM measurement_result mr
            WHERE mr.item_id=mi.id
              AND COALESCE(mr.is_voided,0)=0
              AND mr.collect_time>=?
          )
        ORDER BY COALESCE(mi.last_collect_time,'') ASC, mi.id DESC
        LIMIT 10
    """, (recent_24,)).fetchall()

    ocr_raw_empty_count = conn.execute("""
        SELECT COUNT(*) AS c FROM measurement_result
        WHERE COALESCE(is_voided,0)=0
          AND COALESCE(ocr_config_json,'')<>''
          AND COALESCE(ocr_raw_text,'')=''
    """).fetchone()["c"]
    ocr_conf_empty_count = conn.execute("""
        SELECT COUNT(*) AS c FROM measurement_result
        WHERE COALESCE(is_voided,0)=0
          AND COALESCE(ocr_config_json,'')<>''
          AND ocr_confidence IS NULL
    """).fetchone()["c"]
    ocr_risk_rows = conn.execute("""
        SELECT collect_time, production_code, process_step, metric_name, metric_value_text, result_status,
               source_file_path, ocr_raw_text, ocr_confidence
        FROM measurement_result
        WHERE COALESCE(is_voided,0)=0
          AND COALESCE(ocr_config_json,'')<>''
          AND (ocr_confidence IS NULL OR COALESCE(ocr_raw_text,'')='' OR result_status IN ('TEXT','OCR_ERROR','IMAGE_ROUTE_SKIPPED','IMAGE_NO_MATCH'))
        ORDER BY collect_time DESC
        LIMIT 8
    """).fetchall()

    no_process_count = conn.execute("""
        SELECT COUNT(*) AS c FROM measurement_item_config
        WHERE COALESCE(process_step,'')='' AND COALESCE(process_step_column,'')=''
    """).fetchone()["c"]
    no_metric_count = conn.execute("""
        SELECT COUNT(*) AS c FROM measurement_item_config mi
        WHERE NOT EXISTS (
            SELECT 1 FROM metric_config mc
            WHERE mc.item_id=mi.id AND mc.enabled=1
        )
    """).fetchone()["c"]
    empty_path_count = conn.execute("""
        SELECT COUNT(*) AS c FROM measurement_item_config
        WHERE COALESCE(data_source_path,'')=''
    """).fetchone()["c"]
    image_items = conn.execute("""
        SELECT mi.id, mi.production_id, p.production_code, mi.item_name, mi.image_parse_config_json
        FROM measurement_item_config mi
        JOIN production_config p ON p.id=mi.production_id
        WHERE mi.data_source_type='image'
    """).fetchall()
    image_config_issues = []
    for item in image_items:
        config_text = item["image_parse_config_json"] or ""
        if not config_text.strip():
            image_config_issues.append((item, "Image OCR 配置为空"))
            continue
        try:
            json.loads(config_text)
        except Exception as ex:
            image_config_issues.append((item, f"Image OCR JSON 解析失败：{ex}"))
    config_issue_rows = []
    for row in conn.execute("""
        SELECT mi.id, mi.production_id, p.production_code, mi.item_name, '无工序配置' AS issue
        FROM measurement_item_config mi JOIN production_config p ON p.id=mi.production_id
        WHERE COALESCE(mi.process_step,'')='' AND COALESCE(mi.process_step_column,'')=''
        LIMIT 10
    """).fetchall():
        config_issue_rows.append(row)
    for row in conn.execute("""
        SELECT mi.id, mi.production_id, p.production_code, mi.item_name, '无启用指标' AS issue
        FROM measurement_item_config mi JOIN production_config p ON p.id=mi.production_id
        WHERE NOT EXISTS (SELECT 1 FROM metric_config mc WHERE mc.item_id=mi.id AND mc.enabled=1)
        LIMIT 10
    """).fetchall():
        config_issue_rows.append(row)
    for row in conn.execute("""
        SELECT mi.id, mi.production_id, p.production_code, mi.item_name, '数据源路径为空' AS issue
        FROM measurement_item_config mi JOIN production_config p ON p.id=mi.production_id
        WHERE COALESCE(mi.data_source_path,'')=''
        LIMIT 10
    """).fetchall():
        config_issue_rows.append(row)
    for item, issue in image_config_issues[:10]:
        config_issue_rows.append({
            "id": item["id"],
            "production_id": item["production_id"],
            "production_code": item["production_code"],
            "item_name": item["item_name"],
            "issue": issue
        })
    demo_result_count = conn.execute("""
        SELECT COUNT(*) AS c FROM measurement_result
        WHERE production_code IN ('DEMO_WAFER_A01', 'DEMO_WAFER_B02') AND COALESCE(is_voided,0)=0
    """).fetchone()["c"]
    conn.close()
    config_issues, config_summary, config_issue_total = config_check_issues(limit=200)
    no_process_count = config_summary.get("无工序配置", 0)
    no_metric_count = config_summary.get("无启用指标", 0)
    empty_path_count = config_summary.get("数据源路径为空", 0)
    image_config_issue_count = config_summary.get("Image OCR 配置问题", 0)
    config_issue_rows = config_issues[:10]

    total = status_stats["total"] or 0
    ms3 = status_stats["ms3_count"] or 0
    ms2 = status_stats["ms2_count"] or 0
    miss = status_stats["miss_count"] or 0
    blocks = []
    for process in process_rows:
        process_name = process["process_name"]
        p_total = process["total"] or 0
        p_ms3 = process["ms3_count"] or 0
        p_ms2 = process["ms2_count"] or 0
        p_miss = process["miss_count"] or 0
        metric_html = ""
        for metric, sigma, key in process_metric_map.get(process_name, []):
            metric_html += f"""
            <tr><td><button class="link-button" type="button" onclick="showTrend('{key}')">{h(metric['metric_name'])}</button></td><td>{metric['total']}</td><td>{h(metric['avg_value'])}</td><td>{h(metric['min_value'])}</td><td>{h(metric['max_value'])}</td><td>{h(sigma)}</td><td>{metric['ms3_count'] or 0}</td><td>{(metric['ms2_count'] or 0) + (metric['ms3_count'] or 0)}</td><td>{metric['ms2_count'] or 0}</td><td>{metric['miss_count'] or 0}</td><td>{percent(metric['ms3_count'] or 0, metric['total'])}%</td><td>{percent((metric['ms2_count'] or 0) + (metric['ms3_count'] or 0), metric['total'])}%</td></tr>
            """
        metric_html = metric_html or "<tr><td colspan='12'>暂无指标明细</td></tr>"
        blocks.append(f"""
        <details class="process-risk">
          <summary><span>{h(process_name)}</span><span class="risk-meta">结果 {p_total} ｜ MS3 {p_ms3} ｜ MS2达成 {p_ms2 + p_ms3} ｜ 仅MS2 {p_ms2} ｜ 未达 {p_miss} ｜ MS3达成率 {percent(p_ms3, p_total)}%</span></summary>
          <div class="risk-inner"><div class="table-wrap"><table><tr><th>指标</th><th>结果数</th><th>均值</th><th>最小</th><th>最大</th><th>Sigma</th><th>MS3达成</th><th>MS2达成</th><th>仅MS2</th><th>未达</th><th>MS3达成率</th><th>MS2达成率</th></tr>{metric_html}</table></div></div>
        </details>
        """)
    process_html = "".join(blocks) or "<p class='note'>暂无工序统计</p>"
    abnormal_html = "".join(f"<tr><td>{h(r['collect_time'])}</td><td>{h(r['production_code'])}</td><td>{h(r['process_step'])}</td><td>{h(r['metric_name'])}</td><td>{h(r['metric_value_text'])}</td><td>{badge(r['result_status'])}</td></tr>" for r in abnormal_rows) or "<tr><td colspan='6'>暂无异常记录</td></tr>"
    failure_html = "".join(
        f"<tr><td>{h(r['data_source_path'])}</td><td>{r['failure_count']}</td><td>{h(r['last_failed_at'])}</td><td>{h(r['last_message'])}</td></tr>"
        for r in failure_top
    ) or "<tr><td colspan='4'>最近24小时暂无采集失败数据源</td></tr>"
    no_data_html = "".join(
        f"<tr><td>{h(r['production_code'])}</td><td>{h(r['item_name'])}</td><td>{h(r['equipment_name'])}</td><td>{h(r['data_source_path'])}</td><td>{badge(r['last_collect_status'] or 'NA')}</td><td>{h(r['last_collect_time'] or '')}</td></tr>"
        for r in no_data_rows
    ) or "<tr><td colspan='6'>最近24小时所有启用量测项均有入库结果</td></tr>"
    failed_job_html = "".join(
        f"<tr><td>#{r['id']}</td><td>{h(r['production_code'])}</td><td>{h(r['measurement_item_name'])}</td><td>{h(r['trigger_type'])}</td><td>{badge(r['status'])}</td><td>{h(r['finished_at'] or r['started_at'])}</td><td>{h(r['duration_ms'])}</td><td>{h(r['error_message'])}</td></tr>"
        for r in recent_failed_jobs
    ) or "<tr><td colspan='8'>暂无失败任务</td></tr>"
    ocr_risk_html = "".join(
        f"<tr><td>{h(r['collect_time'])}</td><td>{h(r['production_code'])}</td><td>{h(r['process_step'])}</td><td>{h(r['metric_name'])}</td><td>{h(r['metric_value_text'])}</td><td>{badge(r['result_status'])}</td><td>{h(r['source_file_path'])}</td><td>{h((r['ocr_raw_text'] or '')[:80])}</td><td>{h(r['ocr_confidence'])}</td></tr>"
        for r in ocr_risk_rows
    ) or "<tr><td colspan='9'>暂无 OCR 追溯风险记录</td></tr>"
    config_issue_parts = []
    for r in config_issue_rows[:10]:
        if r.get("id"):
            action = f'<a class="btn secondary" href="/item_edit?id={h(r["id"])}">编辑量测项</a>'
        elif r.get("template_id"):
            action = f'<a class="btn secondary" href="/template_detail?id={h(r["template_id"])}">查看模板</a>'
        else:
            action = ""
        config_issue_parts.append(
            f"<tr><td>{h(r['issue'])}</td><td>{h(r['production_code'])}</td><td>{h(r['item_name'])}</td><td>{h(r.get('suggestion',''))}</td><td>{action}</td></tr>"
        )
    config_issue_html = "".join(config_issue_parts) or "<tr><td colspan='5'>暂无配置完整性问题</td></tr>"
    chart_json = json.dumps(trend_charts, ensure_ascii=False)
    if total == 0 and can_manage_config(user):
        quick_start_html = """
        <div class="card quick-start">
          <h2>5分钟客户演示路径</h2>
          <p class="note">当前库还没有采集结果。可以先加载一组演示 CSV 数据，系统会自动创建生产编号、模板、量测项、指标，并执行一次真实采集。</p>
          <div class="step-list">
            <div class="step"><b><span class="num">1</span>加载 Demo</b><span class="note">生成两组生产编号、三道工序、四个指标。</span></div>
            <div class="step"><b><span class="num">2</span>看 Dashboard</b><span class="note">马上看到 MS3/MS2/未达分布和趋势。</span></div>
            <div class="step"><b><span class="num">3</span>追溯结果</b><span class="note">查看来源 CSV、采集任务和日志。</span></div>
            <div class="step"><b><span class="num">4</span>复用模板</b><span class="note">展示新增料号时如何复用配置。</span></div>
          </div>
          <form method="post" action="/demo_seed" class="demo-actions">
            <button type="submit">一键加载 Demo 数据</button>
            <a class="btn secondary" href="/production_new">手动新增生产编号</a>
            <a class="btn secondary" href="/template_upload">上传模板</a>
          </form>
        </div>
        """
    elif demo_result_count:
        quick_start_html = f"""
        <div class="card quick-start">
          <h2>Demo 数据已就绪</h2>
          <div class="status-strip">
            <span>Demo 结果 {demo_result_count} 条</span>
            <span>MS3 {ms3}</span>
            <span>仅 MS2 {ms2}</span>
            <span>未达 MS2 {miss}</span>
          </div>
          <div class="demo-actions">
            <a class="btn" href="/results?production_code=DEMO_WAFER">查看采集结果</a>
            <a class="btn secondary" href="/collect_jobs">查看采集任务</a>
            <a class="btn secondary" href="/templates">查看模板库</a>
            <a class="btn secondary" href="/config_checks">配置检查</a>
          </div>
        </div>
        """
    else:
        quick_start_html = """
        <div class="card quick-start">
          <h2>开始配置</h2>
          <p class="note">建议按“生产编号 → 模板/量测项 → 指标 → 测试读取 → 采集结果”的顺序完成首条数据闭环。</p>
          <div class="demo-actions"><a class="btn" href="/productions">进入生产编号管理</a><a class="btn secondary" href="/templates">查看模板库</a></div>
        </div>
        """
    return base_layout("首页", f"""
    <h1>Dashboard</h1>
    {quick_start_html}
    <div class="card"><form class="form-row" method="get" action="/"><label>生产编号</label><select name="production_code">{production_options}</select><label>开始日期</label><input type="date" name="start_date" value="{start_date}"><label>结束日期</label><input type="date" name="end_date" value="{end_date}"><button type="submit">更新看板</button><a class="btn secondary" href="/">今天</a></form><p class="note">当前统计区间：{start_date} 00:00:00 至 {end_date} 23:59:59。默认只统计当前仍在“生产编号管理”中的生产编号，已删除配置的历史结果不会进入 Dashboard。</p></div>
    <div class="grid">
      <div class="card metric"><div class="label">区间结果数</div><div class="value">{total}</div></div>
      <div class="card metric"><div class="label">MS3达成率</div><div class="value">{percent(ms3, total)}%</div></div>
      <div class="card metric"><div class="label">MS2达成率</div><div class="value">{percent(ms2 + ms3, total)}%</div></div>
      <div class="card metric"><div class="label">未达MS2率</div><div class="value">{percent(miss, total)}%</div></div>
      <div class="card metric"><div class="label">最近采集时间</div><div class="value" style="font-size:18px">{h(last_collect)}</div></div>
      <div class="card metric"><div class="label">生产编号数量</div><div class="value">{p_count}</div></div>
      <div class="card metric"><div class="label">启用量测项</div><div class="value">{item_count}</div></div>
      <div class="card metric"><div class="label">区间未达MS2数</div><div class="value">{miss}</div></div>
      <div class="card metric"><div class="label">24h成功任务</div><div class="value">{job_stats['success_count'] or 0}</div></div>
      <div class="card metric"><div class="label">24h失败/超时任务</div><div class="value">{(job_stats['failed_count'] or 0) + (job_stats['timeout_count'] or 0)}</div></div>
      <div class="card metric"><div class="label">运行中任务</div><div class="value">{job_stats['running_count'] or 0}</div></div>
      <div class="card metric"><div class="label">配置问题总数</div><div class="value">{config_issue_total}</div></div>
    </div>
    <div class="dash-grid"><div class="card"><h2>区间采集趋势</h2>{svg_bar_chart(day_rows)}</div><div class="card"><h2>状态分布</h2>{svg_status_pie(ms3, ms2, miss, total)}<p class="note">未达成率：{percent(miss, total)}% ｜ MS2达成率：{percent(ms2 + ms3, total)}% ｜ MS3达成率：{percent(ms3, total)}%</p></div></div>
    <div class="dash-grid">
      <div class="card"><h2>最近采集失败 Top 数据源</h2><div class="table-wrap"><table><tr><th>数据源路径</th><th>失败次数</th><th>最近失败时间</th><th>最近错误信息</th></tr>{failure_html}</table></div></div>
      <div class="card"><h2>最近24小时无数据的启用量测项</h2><div class="table-wrap"><table><tr><th>生产编号</th><th>量测项</th><th>设备</th><th>数据源</th><th>上次状态</th><th>上次采集时间</th></tr>{no_data_html}</table></div></div>
    </div>
    <div class="card"><h2>最近失败采集任务 Top 5</h2><div class="table-wrap"><table><tr><th>任务ID</th><th>生产编号</th><th>量测项</th><th>触发</th><th>状态</th><th>时间</th><th>耗时ms</th><th>失败原因</th></tr>{failed_job_html}</table></div></div>
    <div class="card"><h2>OCR 追溯风险提示</h2><p class="note">OCR 原文为空：{ocr_raw_empty_count} ｜ OCR confidence 为空：{ocr_conf_empty_count}。当前 OCR 使用离线 Tesseract 文本识别，暂未写入逐 ROI 置信度。</p><div class="table-wrap"><table><tr><th>采集时间</th><th>生产编号</th><th>工序</th><th>指标</th><th>值</th><th>状态</th><th>图片路径</th><th>OCR原文</th><th>置信度</th></tr>{ocr_risk_html}</table></div></div>
    <div class="card"><h2>配置检查</h2><p class="note">问题总数：{config_issue_total} ｜ 无工序配置：{no_process_count} ｜ 无启用指标：{no_metric_count} ｜ 数据源路径为空：{empty_path_count} ｜ Image OCR 配置问题：{image_config_issue_count}</p><div class="table-wrap"><table><tr><th>问题</th><th>生产编号</th><th>对象</th><th>建议修复</th><th>处理</th></tr>{config_issue_html}</table></div></div>
    <div class="card"><h2>工序/指标风险排行</h2>{process_html}</div>
    <div class="card"><h2>最近未达MS2记录</h2><div class="table-wrap"><table><tr><th>采集时间</th><th>生产编号</th><th>工序</th><th>指标</th><th>值</th><th>状态</th></tr>{abnormal_html}</table></div></div>
    <div class="modal-backdrop" id="trendModal"><div class="modal-card"><div class="modal-head"><h2>指标趋势</h2><button class="secondary" type="button" onclick="closeTrend()">关闭</button></div><div id="trendBody"></div></div></div>
    <script>
      const trendCharts = {chart_json};
      function showTrend(key) {{ document.getElementById('trendBody').innerHTML = trendCharts[key] || '<p class="note">暂无趋势数据</p>'; document.getElementById('trendModal').style.display = 'flex'; }}
      function closeTrend() {{ document.getElementById('trendModal').style.display = 'none'; }}
      document.getElementById('trendModal').addEventListener('click', function(e) {{ if (e.target === this) closeTrend(); }});
    </script>
    """, user)


def page_productions(user):
    conn = get_conn()
    rows = conn.execute("""
        SELECT p.*, COUNT(mi.id) AS item_count
        FROM production_config p
        LEFT JOIN measurement_item_config mi ON mi.production_id=p.id
        GROUP BY p.id
        ORDER BY p.id DESC
    """).fetchall()
    conn.close()
    manage = can_manage_config(user)
    rows_html = "".join(f"""
    <tr>
      <td>{h(r['production_code'])}</td><td>{h(r['production_name'])}</td><td>{h(r['product_model'])}</td><td>{h(r['process_version'])}</td>
      <td>{r['item_count']}</td><td>{badge(r['status'])}</td>
      <td class="actions">
        <a class="btn" href="/items?production_id={r['id']}">量测项</a>
        {f'<a class="btn secondary" href="/production_edit?id={r["id"]}">编辑</a>' if manage else ''}
        <a class="btn secondary" href="/export_config?production_id={r['id']}">导出配置</a>
        {f'''<form class="inline-form" method="post" action="/production_delete" onsubmit="return confirm('确认删除该生产编号及其量测项/指标配置？历史采集结果会保留。')">
          <input type="hidden" name="production_id" value="{r['id']}">
          <button class="danger" type="submit">删除</button>
        </form>''' if manage else ''}
      </td>
    </tr>
    """ for r in rows) or "<tr><td colspan='7'>暂无生产编号</td></tr>"
    manage_buttons = f'<a class="btn" href="/production_new">新增生产编号</a> <a class="btn secondary" href="/templates">模板库</a> <a class="btn secondary" href="/export_all_config">导出全部配置</a> <a class="btn secondary" href="/import_config">导入配置 JSON</a>' if manage else '<a class="btn secondary" href="/templates">模板库</a> <a class="btn secondary" href="/export_all_config">导出全部配置</a>'
    return base_layout("生产编号管理", f"""
    <h1>生产编号管理</h1>
    <div class="card">{manage_buttons}</div>
    <div class="card"><div class="table-wrap"><table><tr><th>生产编号</th><th>生产名称</th><th>产品型号</th><th>工艺版本</th><th>量测项数量</th><th>状态</th><th>操作</th></tr>{rows_html}</table></div></div>
    """, user)


def page_production_form(user, production_id=None, error=""):
    row = None
    if production_id:
        conn = get_conn()
        row = conn.execute("SELECT * FROM production_config WHERE id=?", (production_id,)).fetchone()
        conn.close()
    title = "编辑生产编号" if row else "新增生产编号"
    return base_layout(title, f"""
    <h1>{title}</h1>
    <div class="card">
      {f'<div class="error">{h(error)}</div>' if error else ''}
      <form method="post" action="/production_save">
        <input type="hidden" name="id" value="{h(row['id'] if row else '')}">
        <div class="form-grid">
          <label>生产编号 *</label><input name="production_code" value="{h(row['production_code'] if row else '')}" required placeholder="例如 PROD_A_V1">
          <label>生产名称</label><input name="production_name" value="{h(row['production_name'] if row else '')}" placeholder="例如 产品A V1">
          <label>产品型号</label><input name="product_model" value="{h(row['product_model'] if row else '')}">
          <label>工艺版本</label><input name="process_version" value="{h(row['process_version'] if row else '')}">
          <label>状态</label><select name="status"><option value="enabled" {'selected' if (row and row['status']=='enabled') or not row else ''}>enabled</option><option value="disabled" {'selected' if row and row['status']=='disabled' else ''}>disabled</option></select>
          <label>描述</label><textarea name="description" rows="4">{h(row['description'] if row else '')}</textarea>
        </div>
        <br><button type="submit">保存</button> <a class="btn secondary" href="/productions">返回</a>
      </form>
    </div>
    """, user)


def page_items(user, production_id):
    conn = get_conn()
    prod = conn.execute("SELECT * FROM production_config WHERE id=?", (production_id,)).fetchone()
    if not prod:
        conn.close()
        return base_layout("未找到", "<h1>生产编号不存在</h1>", user)
    rows = conn.execute("SELECT * FROM measurement_item_config WHERE production_id=? ORDER BY id DESC", (production_id,)).fetchall()
    conn.close()
    manage = can_manage_config(user)
    collect_allowed = can_collect(user)
    rows_html = "".join(f"""
    <tr>
      <td>{h(r['item_name'])}</td><td>{h(r['process_step'])}</td><td>{h(r['process_step_column'] if 'process_step_column' in r.keys() else '')}</td><td>{h(r['execution_time_text'])}</td><td>{h(r['equipment_name'])}</td>
      <td>{h(r['data_source_path'])}</td><td>{h(r['scan_frequency_seconds'])} s</td><td>{badge('enabled' if r['enabled'] else 'disabled')}</td><td>{badge(r['last_collect_status'] or 'NA')}</td>
      <td class="actions">
        <a class="btn" href="/metrics?item_id={r['id']}">指标</a>
        {f'<a class="btn secondary" href="/item_edit?id={r["id"]}">编辑</a>' if manage else ''}
        {f'<a class="btn secondary" href="/test_collect?item_id={r["id"]}">测试读取</a>' if collect_allowed else ''}
        {f'<a class="btn secondary" href="/collect_now?item_id={r["id"]}">立即采集</a>' if collect_allowed else ''}
        {f'''<form class="inline-form" method="post" action="/item_delete" onsubmit="return confirm('确认删除该量测项及其指标配置？历史采集结果会保留。')">
          <input type="hidden" name="item_id" value="{r['id']}">
          <button class="danger" type="submit">删除</button>
        </form>''' if manage else ''}
      </td>
    </tr>
    """ for r in rows) or "<tr><td colspan='10'>暂无量测项</td></tr>"
    return base_layout("量测项配置", f"""
    <h1>量测项配置：{h(prod['production_code'])}</h1>
    <div class="card">
      {f'<a class="btn" href="/item_new?production_id={production_id}">新增量测项</a>' if manage else ''}
      {f'<a class="btn secondary" href="/template_apply?production_id={production_id}">从模板新增量测项</a>' if manage else ''}
      <a class="btn secondary" href="/productions">返回生产编号</a>
      <a class="btn secondary" href="/export_config?production_id={production_id}">导出该生产编号配置</a>
    </div>
    <div class="card"><div class="table-wrap"><table><tr><th>量测项</th><th>固定工序</th><th>工序字段</th><th>执行时间</th><th>设备</th><th>数据源路径</th><th>频率</th><th>启用</th><th>最近状态</th><th>操作</th></tr>{rows_html}</table></div></div>
    """, user)


def page_item_form(user, item_id=None, production_id=None, error=""):
    conn = get_conn()
    item = None
    if item_id:
        item = conn.execute("SELECT * FROM measurement_item_config WHERE id=?", (item_id,)).fetchone()
        production_id = item["production_id"] if item else production_id
    prod = conn.execute("SELECT * FROM production_config WHERE id=?", (production_id,)).fetchone() if production_id else None
    conn.close()
    if not prod:
        return base_layout("错误", "<h1>请先选择生产编号</h1>", user)
    title = "编辑量测项" if item else "新增量测项"
    image_config_value = item["image_parse_config_json"] if item and "image_parse_config_json" in item.keys() and item["image_parse_config_json"] else DEFAULT_IMAGE_PARSE_CONFIG_JSON
    path_placeholder = DEFAULT_DATA_SOURCE_PATH_EXAMPLE
    if item and "data_source_type" in item.keys() and item["data_source_type"] == "image":
        path_placeholder = DEFAULT_IMAGE_SOURCE_PATH_EXAMPLE
    return base_layout(title, f"""
    <h1>{title}：{h(prod['production_code'])}</h1>
    <div class="card">
      {f'<div class="error">{h(error)}</div>' if error else ''}
      <form method="post" action="/item_save">
        <input type="hidden" name="id" value="{h(item['id'] if item else '')}">
        <input type="hidden" name="production_id" value="{h(production_id)}">
        <div class="form-grid">
          <label>量测项名称 *</label><input name="item_name" value="{h(item['item_name'] if item else '')}" required placeholder="例如 光刻后CD量测">
          <label>固定量测工序</label><input name="process_step" value="{h(item['process_step'] if item else '')}" placeholder="无工序字段时使用，例如 PHOTO_CD_MEAS">
          <label>工序字段名</label><input name="process_step_column" value="{h(item['process_step_column'] if item and 'process_step_column' in item.keys() else '')}" placeholder="例如 贴装工序；留空则只采集生产编号最后一行">
          <label>量测执行时间</label><input name="execution_time_text" value="{h(item['execution_time_text'] if item else '')}" placeholder="例如 光刻后 / 每日10:00 / 工序完成后">
          <label>量测设备</label><input name="equipment_name" value="{h(item['equipment_name'] if item else '')}" placeholder="例如 CDSEM01">
          <label>数据源类型</label><select name="data_source_type"><option value="auto" {'selected' if (item and 'data_source_type' in item.keys() and item['data_source_type']=='auto') or not item else ''}>自动判断</option><option value="csv" {'selected' if item and 'data_source_type' in item.keys() and item['data_source_type']=='csv' else ''}>CSV</option><option value="excel" {'selected' if item and 'data_source_type' in item.keys() and item['data_source_type']=='excel' else ''}>Excel xlsx/xlsm</option><option value="image" {'selected' if item and 'data_source_type' in item.keys() and item['data_source_type']=='image' else ''}>Image OCR</option></select>
          <label>数据源路径 *</label><input name="data_source_path" value="{h(item['data_source_path'] if item else path_placeholder)}" required style="min-width:520px">
          <label>Excel Sheet 名称</label><input name="excel_sheet_name" value="{h(item['excel_sheet_name'] if item and 'excel_sheet_name' in item.keys() else '')}" placeholder="Excel 多 Sheet 时填写，例如 Sheet1">
          <label>表头所在行</label><input name="header_row_index" value="{h(item['header_row_index'] if item and 'header_row_index' in item.keys() else 1)}" type="number" min="1">
          <label>Image OCR config JSON</label><textarea name="image_parse_config_json" rows="12" style="grid-column:1 / -1; font-family:Consolas,monospace">{h(image_config_value)}</textarea>
          <label>CSV编码</label><select name="csv_encoding">
  <option value="auto" {'selected' if (item and item['csv_encoding']=='auto') or not item else ''}>auto 自动识别</option>
  <option value="utf-8-sig" {'selected' if item and item['csv_encoding']=='utf-8-sig' else ''}>utf-8-sig</option>
  <option value="utf-8" {'selected' if item and item['csv_encoding']=='utf-8' else ''}>utf-8</option>
  <option value="gb18030" {'selected' if item and item['csv_encoding']=='gb18030' else ''}>gb18030</option>
  <option value="gbk" {'selected' if item and item['csv_encoding']=='gbk' else ''}>gbk</option>
</select>
          <label>分隔符</label><input name="delimiter" value="{h(item['delimiter'] if item else ',')}" placeholder=", 或 \t">
          <label>生产编号字段名</label><input name="production_code_column" value="{h(item['production_code_column'] if item else '生产编号')}" placeholder="生产编号">
          <label>抓取频率/秒</label><input name="scan_frequency_seconds" value="{h(item['scan_frequency_seconds'] if item else 60)}" type="number" min="10">
          <label>是否启用</label><select name="enabled"><option value="1" {'selected' if (item and item['enabled']) or not item else ''}>启用</option><option value="0" {'selected' if item and not item['enabled'] else ''}>停用</option></select>
        </div>
        <br><button type="submit">保存</button> <a class="btn secondary" href="/items?production_id={production_id}">返回</a>
      </form>
      <p class="note">说明：CSV/Excel 如果一个生产编号在同一个 Sheet 里有多道工序，请填写“工序字段名”，例如 贴装工序。图片 OCR 正式自动采集建议把“数据源路径”填成设备共享图片目录，config 使用 collect_mode=all_stable，并通过 production_code_from_filename_regex 或 fields.production_code 从每张图片解析生产编号，通过 process_from_filename_regex 或 fields.process_step 解析工序；这样不需要每个生产编号单独配置一个只放一张图的目录。</p>
    </div>
    """, user)


def page_metrics(user, item_id):
    conn = get_conn()
    item = conn.execute("""
        SELECT mi.*, p.production_code FROM measurement_item_config mi
        JOIN production_config p ON p.id=mi.production_id
        WHERE mi.id=?
    """, (item_id,)).fetchone()
    if not item:
        conn.close()
        return base_layout("错误", "<h1>量测项不存在</h1>", user)
    rows = conn.execute("SELECT * FROM metric_config WHERE item_id=? ORDER BY sort_order, id", (item_id,)).fetchall()
    conn.close()
    manage = can_manage_config(user)
    metric_parts = []
    for r in rows:
        actions = ""
        if manage:
            actions = f"""<a class="btn secondary" href="/metric_edit?id={r['id']}">编辑</a><form class="inline-form" method="post" action="/metric_delete" onsubmit="return confirm('确认删除该指标配置？历史采集结果会保留。')"><input type="hidden" name="metric_id" value="{r['id']}"><button class="danger" type="submit">删除</button></form>"""
        metric_parts.append(f"<tr><td>{h(r['metric_name'])}</td><td>{h(r['source_column'])}</td><td>{h(r['unit'])}</td><td>{h(r['data_type'])}</td><td>{h(r['target'])}</td><td>{h(r['lsl'])}</td><td>{h(r['usl'])}</td><td>{h(r['lcl'])}</td><td>{h(r['ucl'])}</td><td>{badge('enabled' if r['enabled'] else 'disabled')}</td><td class='actions'>{actions}</td></tr>")
    rows_html = "".join(metric_parts) or "<tr><td colspan='11'>暂无指标。建议先批量添加：Dx1,Dy1,Dx2,Dy2,Rz</td></tr>"
    manage_form = f"""
      <form method="post" action="/metric_bulk_add" class="form-row">
        <input type="hidden" name="item_id" value="{item_id}">
        <input name="metric_names" style="min-width:420px" value="Dx1,Dy1,Dx2,Dy2,Rz" placeholder="Dx1,Dy1,Dx2,Dy2,Rz">
        <input name="unit" placeholder="单位，可空，例如 um">
        <button type="submit">批量添加指标</button>
        <a class="btn secondary" href="/metric_new?item_id={item_id}">单个新增</a>
        <a class="btn secondary" href="/items?production_id={item['production_id']}">返回量测项</a>
      </form>
      <p class="note">批量添加时，平台指标名和 数据源字段名默认一致，例如 Dx1 ← CSV列 Dx1。</p>
    """ if manage else f'<p><a class="btn secondary" href="/items?production_id={item["production_id"]}">返回量测项</a></p><p class="note">当前账号为只读权限，不能新增或编辑指标。</p>'
    return base_layout("指标配置", f"""
    <h1>指标配置：{h(item['production_code'])} / {h(item['item_name'])}</h1>
    <div class="card">
      {manage_form}
    </div>
    <div class="card"><div class="table-wrap"><table><tr><th>指标名称</th><th>源字段</th><th>单位</th><th>类型</th><th>Target</th><th>MS2下限</th><th>MS2上限</th><th>MS3下限</th><th>MS3上限</th><th>状态</th><th>操作</th></tr>{rows_html}</table></div></div>
    """, user)


def page_metric_form(user, metric_id=None, item_id=None, error=""):
    conn = get_conn()
    metric = None
    if metric_id:
        metric = conn.execute("SELECT * FROM metric_config WHERE id=?", (metric_id,)).fetchone()
        item_id = metric["item_id"] if metric else item_id
    item = conn.execute("SELECT * FROM measurement_item_config WHERE id=?", (item_id,)).fetchone() if item_id else None
    conn.close()
    if not item:
        return base_layout("错误", "<h1>量测项不存在</h1>", user)
    title = "编辑指标" if metric else "新增指标"
    return base_layout(title, f"""
    <h1>{title}</h1>
    <div class="card">
      {f'<div class="error">{h(error)}</div>' if error else ''}
      <form method="post" action="/metric_save">
        <input type="hidden" name="id" value="{h(metric['id'] if metric else '')}">
        <input type="hidden" name="item_id" value="{h(item_id)}">
        <div class="form-grid">
          <label>指标名称 *</label><input name="metric_name" value="{h(metric['metric_name'] if metric else '')}" required placeholder="例如 Dx1">
          <label>源字段名 *</label><input name="source_column" value="{h(metric['source_column'] if metric else '')}" required placeholder="例如 Dx1">
          <label>单位</label><input name="unit" value="{h(metric['unit'] if metric else '')}" placeholder="例如 um / nm / deg">
          <label>数据类型</label><select name="data_type"><option value="number" {'selected' if (metric and metric['data_type']=='number') or not metric else ''}>number</option><option value="text" {'selected' if metric and metric['data_type']=='text' else ''}>text</option></select>
          <label>Target</label><input name="target" value="{h(metric['target'] if metric else '')}">
          <label>MS2下限</label><input name="lsl" value="{h(metric['lsl'] if metric else '')}">
          <label>MS2上限</label><input name="usl" value="{h(metric['usl'] if metric else '')}">
          <label>MS3下限</label><input name="lcl" value="{h(metric['lcl'] if metric else '')}">
          <label>MS3上限</label><input name="ucl" value="{h(metric['ucl'] if metric else '')}">
          <label>排序</label><input name="sort_order" value="{h(metric['sort_order'] if metric else 0)}" type="number">
          <label>状态</label><select name="enabled"><option value="1" {'selected' if (metric and metric['enabled']) or not metric else ''}>启用</option><option value="0" {'selected' if metric and not metric['enabled'] else ''}>停用</option></select>
        </div>
        <br><button type="submit">保存</button> <a class="btn secondary" href="/metrics?item_id={item_id}">返回</a>
      </form>
    </div>
    """, user)


def page_test_collect(user, item_id):
    result = collect_item_with_timeout(item_id, dry_run=True)
    body = f"""
    <h1>测试读取结果</h1>
    <div class="card">
      <p>状态：{badge(result.get('status'))}</p>
      <p>{h(result.get('message'))}</p>
      <p class="note"><b>注意：</b>测试读取只验证路径、表头、生产编号和指标映射，不会写入采集结果。要入库请点击下面的“确认无误，立即采集入库”。如果 metric_preview 为空，说明该量测项还没有添加指标配置。</p>
      <a class="btn" href="/collect_now?item_id={item_id}">确认无误，立即采集入库</a>
      <a class="btn secondary" href="/item_edit?id={item_id}">返回编辑量测项</a>
    </div>
    <div class="card"><h2>识别到的字段</h2><pre>{h(json.dumps(result.get('fieldnames', []), ensure_ascii=False, indent=2))}</pre></div>
    <div class="card"><h2>匹配到的行数 / 采集行数</h2><pre>{h(json.dumps({'matched_rows': result.get('matched_rows'), 'collect_rows': result.get('collect_rows'), 'process_step_column': result.get('process_step_column'), 'metric_preview': result.get('metric_preview')}, ensure_ascii=False, indent=2))}</pre></div>
    <div class="card"><h2>多工序行预览</h2><pre>{h(json.dumps(result.get('row_previews', []), ensure_ascii=False, indent=2))}</pre></div>
    <div class="card"><h2>Image OCR 逐指标明细</h2><pre>{h(json.dumps(result.get('image_ocr_details', []), ensure_ascii=False, indent=2))}</pre></div>
    <div class="card"><h2>Image OCR</h2><pre>{h(json.dumps(result.get('image_ocr', []), ensure_ascii=False, indent=2))}</pre></div>
    """
    return base_layout("测试读取", body, user)


def page_collect_now(user, item_id):
    result = collect_item_with_timeout(item_id, dry_run=False)
    body = f"""
    <h1>立即采集结果</h1>
    <div class="card">
      <p>状态：{badge(result.get('status'))}</p>
      <p>{h(result.get('message'))}</p>
      <a class="btn" href="/results">查看采集结果</a>
      <a class="btn secondary" href="/logs">查看采集日志</a>
      <a class="btn secondary" href="/item_edit?id={item_id}">返回量测项</a>
    </div>
    <div class="card"><h2>详情</h2><pre>{h(json.dumps(result, ensure_ascii=False, indent=2))}</pre></div>
    """
    return base_layout("立即采集", body, user)


def page_results(user, query=None):
    query = query or {}
    production_code = query.get("production_code", [""])[0].strip()
    process_step = query.get("process_step", [""])[0].strip()
    metric_name = query.get("metric_name", [""])[0].strip()
    result_status = query.get("result_status", [""])[0].strip()
    show_voided = query.get("show_voided", ["0"])[0] == "1"
    where, params = [], []
    if not show_voided:
        where.append("COALESCE(is_voided,0)=0")
    if production_code:
        where.append("production_code LIKE ?")
        params.append(f"%{production_code}%")
    if process_step:
        where.append("process_step LIKE ?")
        params.append(f"%{process_step}%")
    if metric_name:
        where.append("metric_name LIKE ?")
        params.append(f"%{metric_name}%")
    if result_status:
        where.append("result_status=?")
        params.append(result_status)
    sql = "SELECT * FROM measurement_result"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY collect_time DESC LIMIT 500"
    conn = get_conn()
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    has_rows = bool(rows)
    status_values = ["MS3_PASS", "MS2_PASS", "MISS_MS2", "PASS", "TEXT"]
    status_options = "".join(f'<option value="{s}" {"selected" if result_status == s else ""}>{display_status(s)}</option>' for s in status_values)
    export_href = "/export_results_xlsx?" + urlencode({
        "production_code": production_code,
        "process_step": process_step,
        "metric_name": metric_name,
        "result_status": result_status,
        "show_voided": "1" if show_voided else "0",
    })
    row_parts = []
    void_allowed = can_void_results(user)
    for r in rows:
        trace_detail = {
            "source_file_path": r["source_file_path"],
            "source_file_mtime": r["source_file_mtime"],
            "ocr_raw_text": r["ocr_raw_text"],
            "ocr_roi_json": r["ocr_roi_json"],
            "ocr_confidence": r["ocr_confidence"],
            "voided_by": r["voided_by"],
            "voided_at": r["voided_at"],
            "void_reason": r["void_reason"],
        }
        action_html = ""
        if void_allowed and not r["is_voided"]:
            action_html = f"""
            <form class="inline-form" method="post" action="/result_delete" onsubmit="return confirm('确认作废这条采集结果？')">
              <input type="hidden" name="result_id" value="{r['id']}">
              <input name="void_reason" value="人工作废" placeholder="作废原因" style="width:100px">
              <button class="danger" type="submit">作废</button>
            </form>
            """
        row_parts.append(f"""
        <tr>
          <td>{f'<input form="bulkVoidResultsForm" type="checkbox" name="result_ids" value="{r["id"]}">' if void_allowed else ''}</td>
          <td>{h(r['collect_time'])}</td><td>{h(r['production_code'])}</td><td>{h(r['measurement_item_name'])}</td><td>{h(r['process_step'])}</td><td>{h(r['execution_time_text'])}</td><td>{h(r['equipment_name'])}</td><td>{h(r['metric_name'])}</td><td>{h(r['metric_value_text'])}</td><td>{h(r['unit'])}</td><td>{badge('VOIDED') if r['is_voided'] else badge(r['result_status'])}</td><td>{h(r['source_file_path'] or r['source_path'])}</td>
          <td><details><summary>详情</summary><pre>{h(json.dumps(trace_detail, ensure_ascii=False, indent=2))}</pre></details></td>
          <td>{action_html}</td>
        </tr>
        """)
    rows_html = "".join(row_parts) or "<tr><td colspan='14'>暂无结果</td></tr>"
    void_tools = """
    <form class="inline-form" method="post" action="/blank_process_results_clear" onsubmit="return confirm('确认作废全部空工序采集结果？')"><button class="danger" type="submit">作废空工序结果</button></form>
    """ if void_allowed and has_rows else ""
    bulk_void_form = """
      <form id="bulkVoidResultsForm" method="post" action="/results_bulk_delete" onsubmit="return confirm('确认作废勾选的采集结果？')">
        <input name="void_reason" value="人工作废" placeholder="作废原因">
        <button class="danger" type="submit">批量作废</button>
      </form>
    """ if void_allowed and has_rows else ("<p class='note'>暂无可作废结果；加载 Demo 数据或完成采集后再显示批量作废工具。</p>" if void_allowed else "<p class='note'>当前账号无作废权限。</p>")
    return base_layout("采集结果", f"""
    <h1>采集结果</h1>
    <div class="card"><form class="form-row" method="get" action="/results"><input name="production_code" value="{h(production_code)}" placeholder="生产编号"><input name="process_step" value="{h(process_step)}" placeholder="工序名"><input name="metric_name" value="{h(metric_name)}" placeholder="指标名"><select name="result_status"><option value="">全部状态</option>{status_options}</select><label><input type="checkbox" name="show_voided" value="1" {'checked' if show_voided else ''}> 显示已作废结果</label><button type="submit">查询</button><a class="btn secondary" href="/results">重置</a><a class="btn secondary" href="{export_href}">导出 Excel</a></form>{void_tools}</div>
    <div class="card">
      {bulk_void_form}
      <div class="table-wrap"><table><tr><th>选择</th><th>采集时间</th><th>生产编号</th><th>量测项</th><th>工序</th><th>执行时间</th><th>设备</th><th>指标</th><th>值</th><th>单位</th><th>状态</th><th>来源文件</th><th>追溯</th><th>操作</th></tr>{rows_html}</table></div>
    </div>
    """, user)


def fetch_result_rows(query=None, limit=5000):
    query = query or {}
    production_code = query.get("production_code", [""])[0].strip()
    process_step = query.get("process_step", [""])[0].strip()
    metric_name = query.get("metric_name", [""])[0].strip()
    result_status = query.get("result_status", [""])[0].strip()
    show_voided = query.get("show_voided", ["0"])[0] == "1"
    where, params = [], []
    if not show_voided:
        where.append("COALESCE(is_voided,0)=0")
    if production_code:
        where.append("production_code LIKE ?")
        params.append(f"%{production_code}%")
    if process_step:
        where.append("process_step LIKE ?")
        params.append(f"%{process_step}%")
    if metric_name:
        where.append("metric_name LIKE ?")
        params.append(f"%{metric_name}%")
    if result_status:
        where.append("result_status=?")
        params.append(result_status)
    sql = "SELECT * FROM measurement_result"
    if where:
        sql += " WHERE " + " AND ".join(where)
    sql += " ORDER BY collect_time DESC LIMIT ?"
    params.append(limit)
    conn = get_conn()
    rows = conn.execute(sql, params).fetchall()
    conn.close()
    return rows


def xml_text(value):
    return html.escape("" if value is None else str(value), quote=False)


def xml_attr(value):
    return html.escape("" if value is None else str(value), quote=True)


def build_xlsx(headers, rows, sheet_name="Results"):
    output = io.BytesIO()
    all_rows = [headers] + rows
    sheet_rows = []
    for r_idx, row in enumerate(all_rows, start=1):
        cells = []
        for c_idx, value in enumerate(row):
            ref = cell_ref(c_idx, r_idx)
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{xml_text(value)}</t></is></c>')
        sheet_rows.append(f'<row r="{r_idx}">{"".join(cells)}</row>')
    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{"".join(sheet_rows)}</sheetData></worksheet>'''
    workbook_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="{xml_attr(sheet_name)[:31]}" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>''')
        zf.writestr("_rels/.rels", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>''')
        zf.writestr("xl/_rels/workbook.xml.rels", '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>''')
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    return output.getvalue()


def export_results_xlsx(query=None):
    rows = fetch_result_rows(query, limit=5000)
    headers = ["采集时间", "生产编号", "量测项", "工序", "执行时间", "设备", "指标", "值", "数值", "单位", "状态", "是否作废", "来源路径", "来源文件", "文件修改时间", "OCR原文", "OCR ROI", "OCR置信度", "作废人", "作废时间", "作废原因"]
    data_rows = [[
        r["collect_time"], r["production_code"], r["measurement_item_name"], r["process_step"], r["execution_time_text"],
        r["equipment_name"], r["metric_name"], r["metric_value_text"], r["metric_value_number"], r["unit"],
        r["result_status"], r["is_voided"], r["source_path"], r["source_file_path"], r["source_file_mtime"],
        r["ocr_raw_text"], r["ocr_roi_json"], r["ocr_confidence"], r["voided_by"], r["voided_at"], r["void_reason"]
    ] for r in rows]
    return build_xlsx(headers, data_rows, "Results")


def void_result(result_id, username="", reason="人工作废"):
    conn = get_conn()
    cur = conn.cursor()
    row = cur.execute("SELECT * FROM measurement_result WHERE id=?", (result_id,)).fetchone()
    if not row:
        conn.close()
        raise ValueError("采集结果不存在")
    cur.execute("""
        UPDATE measurement_result
        SET is_voided=1, voided_by=?, voided_at=?, void_reason=?
        WHERE id=?
    """, (username or "", now_str(), reason or "人工作废", result_id))
    conn.commit()
    conn.close()
    return row["production_code"], row["process_step"], row["metric_name"]


def delete_result(result_id, username="", reason="人工作废"):
    return void_result(result_id, username, reason)


def bulk_void_results(result_ids, username="", reason="人工作废"):
    ids = [safe_int(x) for x in result_ids if safe_int(x) > 0]
    if not ids:
        return 0
    placeholders = ",".join("?" for _ in ids)
    conn = get_conn()
    cur = conn.cursor()
    count = cur.execute(f"SELECT COUNT(*) AS c FROM measurement_result WHERE id IN ({placeholders}) AND COALESCE(is_voided,0)=0", ids).fetchone()["c"]
    cur.execute(f"""
        UPDATE measurement_result
        SET is_voided=1, voided_by=?, voided_at=?, void_reason=?
        WHERE id IN ({placeholders}) AND COALESCE(is_voided,0)=0
    """, [username or "", now_str(), reason or "人工作废"] + ids)
    conn.commit()
    conn.close()
    return count


def bulk_delete_results(result_ids, username="", reason="人工作废"):
    return bulk_void_results(result_ids, username, reason)


def clear_blank_process_results():
    conn = get_conn()
    cur = conn.cursor()
    count = cur.execute("SELECT COUNT(*) AS c FROM measurement_result WHERE COALESCE(process_step,'')='' AND COALESCE(is_voided,0)=0").fetchone()["c"]
    cur.execute("""
        UPDATE measurement_result
        SET is_voided=1, voided_by='system', voided_at=?, void_reason='空工序清理作废'
        WHERE COALESCE(process_step,'')='' AND COALESCE(is_voided,0)=0
    """, (now_str(),))
    conn.commit()
    conn.close()
    return count


def delete_production_config(production_id):
    conn = get_conn()
    cur = conn.cursor()
    prod = cur.execute("SELECT * FROM production_config WHERE id=?", (production_id,)).fetchone()
    if not prod:
        conn.close()
        raise ValueError("生产编号不存在")
    item_ids = [r["id"] for r in cur.execute("SELECT id FROM measurement_item_config WHERE production_id=?", (production_id,)).fetchall()]
    for item_id in item_ids:
        cur.execute("DELETE FROM metric_config WHERE item_id=?", (item_id,))
    cur.execute("DELETE FROM measurement_item_config WHERE production_id=?", (production_id,))
    cur.execute("DELETE FROM production_config WHERE id=?", (production_id,))
    conn.commit()
    conn.close()
    return prod["production_code"]


def delete_item_config(item_id):
    conn = get_conn()
    cur = conn.cursor()
    item = cur.execute("SELECT * FROM measurement_item_config WHERE id=?", (item_id,)).fetchone()
    if not item:
        conn.close()
        raise ValueError("量测项不存在")
    production_id = item["production_id"]
    cur.execute("DELETE FROM metric_config WHERE item_id=?", (item_id,))
    cur.execute("DELETE FROM measurement_item_config WHERE id=?", (item_id,))
    conn.commit()
    conn.close()
    return production_id, item["item_name"]


def delete_metric_config(metric_id):
    conn = get_conn()
    cur = conn.cursor()
    metric = cur.execute("SELECT * FROM metric_config WHERE id=?", (metric_id,)).fetchone()
    if not metric:
        conn.close()
        raise ValueError("指标不存在")
    item_id = metric["item_id"]
    cur.execute("DELETE FROM metric_config WHERE id=?", (metric_id,))
    conn.commit()
    conn.close()
    return item_id, metric["metric_name"]


def clear_collect_logs():
    raise RuntimeError("产线试运行阶段禁止清空 collect_log。请保留历史采集日志用于追溯。")


def clear_orphan_measurement_results():
    conn = get_conn()
    cur = conn.cursor()
    count = cur.execute("""
        SELECT COUNT(*) AS c
        FROM measurement_result
        WHERE COALESCE(is_voided,0)=0
          AND (COALESCE(production_code,'') NOT IN (SELECT production_code FROM production_config)
           OR NOT EXISTS (
                SELECT 1
                FROM measurement_item_config mi
                WHERE mi.id=measurement_result.item_id
           ))
    """).fetchone()["c"]
    cur.execute("""
        UPDATE measurement_result
        SET is_voided=1, voided_by='system', voided_at=?, void_reason='孤儿结果清理作废'
        WHERE COALESCE(is_voided,0)=0
          AND (COALESCE(production_code,'') NOT IN (SELECT production_code FROM production_config)
           OR NOT EXISTS (
                SELECT 1
                FROM measurement_item_config mi
                WHERE mi.id=measurement_result.item_id
           ))
    """, (now_str(),))
    conn.commit()
    conn.close()
    return count


def prune_collect_logs(keep=100):
    raise RuntimeError("产线试运行阶段禁止裁剪 collect_log。请保留完整历史采集日志用于追溯。")


def page_collect_jobs(user):
    conn = get_conn()
    rows = conn.execute("""
        SELECT * FROM collect_job
        ORDER BY COALESCE(created_at, scheduled_at, started_at) DESC, id DESC
        LIMIT 500
    """).fetchall()
    conn.close()
    manage = can_manage_config(user)
    rows_html = "".join(f"""
    <tr>
      <td>#{r['id']}</td><td>{h(r['production_code'])}</td><td>{h(r['measurement_item_name'])}</td>
      <td>{badge(r['status'])}</td><td>{h(r['trigger_type'])}</td><td>{h(r['duration_ms'])}</td>
      <td>{h(r['matched_rows'])}</td><td>{h(r['collect_rows'])}</td><td>{h(r['inserted_count'])}</td>
      <td>{h(r['skipped_count'])}</td><td>{h(r['failed_count'])}</td><td>{h(r['started_at'])}</td><td>{h(r['finished_at'])}</td>
      <td>{h(r['data_source_type'])}</td><td title="{h(r['data_source_path'])}" style="max-width:360px;overflow:hidden;text-overflow:ellipsis">{h(r['data_source_path'])}</td>
      <td>{h(r['error_message'])}</td>
    </tr>
    """ for r in rows) or "<tr><td colspan='16'>暂无采集任务</td></tr>"
    return base_layout("采集任务", f"""
    <h1>采集任务</h1>
    <div class="card note">每次手动采集、定时采集和测试读取都会生成一条任务记录，用于判断任务是否触发、是否运行中、是否超时、是否成功但没有匹配数据。</div>
    <div class="card"><div class="table-wrap"><table>
      <tr><th>任务ID</th><th>生产编号</th><th>量测项</th><th>状态</th><th>触发方式</th><th>耗时ms</th><th>匹配行</th><th>采集行</th><th>新增</th><th>跳过</th><th>失败</th><th>开始时间</th><th>结束时间</th><th>数据源类型</th><th>数据源</th><th>失败原因</th></tr>
      {rows_html}
    </table></div></div>
    """, user)


def page_config_checks(user):
    issues, summary, total = config_check_issues(limit=1000)
    summary_html = "".join(
        f"<tr><td>{h(name)}</td><td>{count}</td></tr>"
        for name, count in sorted(summary.items(), key=lambda kv: (-kv[1], kv[0]))
    ) or "<tr><td colspan='2'>暂无配置问题</td></tr>"
    rows_html = []
    for r in issues:
        if r.get("id"):
            action = f'<a class="btn secondary" href="/item_edit?id={h(r["id"])}">编辑量测项</a>' if can_manage_config(user) else '<span class="note">只读</span>'
        elif r.get("template_id"):
            action = f'<a class="btn secondary" href="/template_detail?id={h(r["template_id"])}">查看模板</a>'
        else:
            action = ""
        rows_html.append(
            f"<tr><td>{h(r['issue'])}</td><td>{h(r['production_code'])}</td><td>{h(r['item_name'])}</td><td>{h(r['suggestion'])}</td><td>{action}</td></tr>"
        )
    detail_html = "".join(rows_html) or "<tr><td colspan='5'>暂无配置问题</td></tr>"
    return base_layout("配置检查", f"""
    <h1>配置检查</h1>
    <div class="card"><p class="note">配置问题总数：{total}。该页面覆盖无工序、无指标、空路径、Excel Sheet、Image OCR JSON、模板无指标和模板生产编号字段缺失。</p></div>
    <div class="dash-grid">
      <div class="card"><h2>问题汇总</h2><div class="table-wrap"><table><tr><th>问题类型</th><th>数量</th></tr>{summary_html}</table></div></div>
      <div class="card"><h2>Top 问题明细</h2><div class="table-wrap"><table><tr><th>问题类型</th><th>生产编号</th><th>相关量测项/模板</th><th>建议修复方式</th><th>处理</th></tr>{detail_html}</table></div></div>
    </div>
    """, user)


def page_logs(user):
    conn = get_conn()
    rows = conn.execute("SELECT * FROM collect_log ORDER BY created_at DESC LIMIT 500").fetchall()
    job_rows = conn.execute("""
        SELECT * FROM collect_job
        ORDER BY COALESCE(created_at, scheduled_at, started_at) DESC, id DESC
        LIMIT 20
    """).fetchall()
    total = conn.execute("SELECT COUNT(*) AS c FROM collect_log").fetchone()["c"]
    orphan_results = conn.execute("""
        SELECT COUNT(*) AS c
        FROM measurement_result
        WHERE COALESCE(production_code,'') NOT IN (SELECT production_code FROM production_config)
           OR NOT EXISTS (
                SELECT 1
                FROM measurement_item_config mi
                JOIN production_config p ON p.id=mi.production_id
                WHERE mi.id=measurement_result.item_id
                  AND p.production_code=measurement_result.production_code
           )
    """).fetchone()["c"]
    conn.close()
    rows_html = "".join(f"""
    <tr><td>{h(r['created_at'])}</td><td>{h(r['production_code'])}</td><td>{h(r['measurement_item_name'])}</td><td>{badge(r['status'])}</td><td>{h(r['matched_rows'])}</td><td>{h(r['inserted_count'])}</td><td>{h(r['skipped_count'])}</td><td>{h(r['message'])}</td><td>{h(r['data_source_path'])}</td></tr>
    """ for r in rows) or "<tr><td colspan='9'>暂无日志</td></tr>"
    jobs_html = "".join(f"""
    <tr><td>#{r['id']}</td><td>{h(r['production_code'])}</td><td>{h(r['measurement_item_name'])}</td><td>{h(r['trigger_type'])}</td><td>{badge(r['status'])}</td><td>{h(r['started_at'])}</td><td>{h(r['finished_at'])}</td><td>{h(r['duration_ms'])}</td><td>{h(r['inserted_count'])}</td><td>{h(r['error_message'])}</td></tr>
    """ for r in job_rows) or "<tr><td colspan='10'>暂无采集任务</td></tr>"
    admin_tools = """
      <p class="note">产线试运行阶段禁止清空或裁剪采集日志。历史日志必须保留，用于排查路径超时、采集失败和权限问题。</p>
      <form class="inline-form" method="post" action="/orphan_results_clear" onsubmit="return confirm('确认作废不属于当前有效生产编号/量测项配置的历史采集结果？')"><button class="danger" type="submit">作废无效历史结果</button></form>
    """ if can_manage_system(user) else ""
    return base_layout("采集日志", f"""
    <h1>采集日志</h1>
    <div class="card">
      <p class="note">当前采集日志共 {total} 条；不属于当前有效生产编号/量测项配置的历史采集结果 {orphan_results} 条。</p>
      {admin_tools}
    </div>
    <div class="card"><h2>最近采集任务</h2><div class="table-wrap"><table><tr><th>任务ID</th><th>生产编号</th><th>量测项</th><th>触发方式</th><th>状态</th><th>开始</th><th>结束</th><th>耗时ms</th><th>新增</th><th>失败原因</th></tr>{jobs_html}</table></div><p class="note"><a href="/collect_jobs">查看全部采集任务</a></p></div>
    <div class="card"><h2>兼容采集日志</h2><div class="table-wrap"><table><tr><th>时间</th><th>生产编号</th><th>量测项</th><th>状态</th><th>匹配行</th><th>新增</th><th>跳过</th><th>信息</th><th>数据源</th></tr>{rows_html}</table></div></div>
    """, user)


def page_audit_logs(user):
    conn = get_conn()
    rows = conn.execute("SELECT * FROM audit_log ORDER BY created_at DESC LIMIT 500").fetchall()
    conn.close()
    rows_html = "".join(
        f"<tr><td>{h(r['created_at'])}</td><td>{h(r['username'])}</td><td>{h(r['action'])}</td><td>{h(r['object_type'])}</td><td>{h(r['object_id'])}</td><td>{h(r['ip_address'])}</td><td>{h(r['detail'])}</td></tr>"
        for r in rows
    ) or "<tr><td colspan='7'>暂无审计日志</td></tr>"
    return base_layout("审计日志", f"""
    <h1>审计日志</h1>
    <div class="card note">记录登录、配置新增/修改、指标配置、配置导入/导出等关键操作。正式部署建议接入公司 AD/LDAP 账号。</div>
    <div class="card"><div class="table-wrap"><table><tr><th>时间</th><th>用户</th><th>动作</th><th>对象类型</th><th>对象ID</th><th>IP</th><th>详情</th></tr>{rows_html}</table></div></div>
    """, user)


def page_users(user, message=""):
    require_permission(user, can_manage_users(user))
    conn = get_conn()
    rows = conn.execute("SELECT id, username, role, enabled, created_at, updated_at FROM admin_user ORDER BY id").fetchall()
    conn.close()
    role_options = "".join(f'<option value="{role}">{role}</option>' for role in VALID_ROLES)
    rows_html = "".join(f"""
    <tr>
      <td>{h(r['username'])}</td><td>{h(r['role'])}</td><td>{badge('enabled' if r['enabled'] else 'disabled')}</td>
      <td>{h(r['created_at'])}</td><td>{h(r['updated_at'])}</td>
      <td class="actions">
        <form class="inline-form" method="post" action="/user_role_save">
          <input type="hidden" name="user_id" value="{r['id']}">
          <select name="role">{''.join(f'<option value="{role}" {"selected" if role == r["role"] else ""}>{role}</option>' for role in VALID_ROLES)}</select>
          <button type="submit">保存角色</button>
        </form>
        <form class="inline-form" method="post" action="/user_reset_password">
          <input type="hidden" name="user_id" value="{r['id']}">
          <input name="password" type="password" placeholder="新密码" required>
          <button class="secondary" type="submit">重置密码</button>
        </form>
        <form class="inline-form" method="post" action="/user_toggle_enabled">
          <input type="hidden" name="user_id" value="{r['id']}">
          <input type="hidden" name="enabled" value="{'0' if r['enabled'] else '1'}">
          <button class="danger" type="submit">{'禁用' if r['enabled'] else '启用'}</button>
        </form>
      </td>
    </tr>
    """ for r in rows) or "<tr><td colspan='6'>暂无用户</td></tr>"
    return base_layout("用户管理", f"""
    <h1>用户管理</h1>
    <div class="card">
      {f'<div class="success">{h(message)}</div>' if message else ''}
      <form method="post" action="/user_create" class="form-row">
        <input name="username" placeholder="用户名" required>
        <input name="password" type="password" placeholder="初始密码" required>
        <select name="role">{role_options}</select>
        <button type="submit">新增用户</button>
      </form>
      <p class="note">viewer 只读；engineer 可维护配置和手动/测试采集；admin 可管理用户和作废结果。</p>
    </div>
    <div class="card"><div class="table-wrap"><table><tr><th>用户名</th><th>角色</th><th>状态</th><th>创建时间</th><th>更新时间</th><th>操作</th></tr>{rows_html}</table></div></div>
    """, user)


def handle_user_create(current, form, ip_address):
    require_permission(current, can_manage_users(current))
    username = form.get("username", [""])[0].strip()
    password = form.get("password", [""])[0]
    role = form.get("role", ["viewer"])[0]
    if role not in VALID_ROLES:
        raise ValueError("角色不合法")
    if not username or not password:
        raise ValueError("用户名和密码不能为空")
    conn = get_conn()
    conn.execute("""
        INSERT INTO admin_user (username, password_hash, role, enabled, created_at, updated_at)
        VALUES (?, ?, ?, 1, ?, ?)
    """, (username, hash_password(password), role, now_str(), now_str()))
    conn.commit()
    conn.close()
    write_audit(current.get("username"), "CREATE_USER", "admin_user", username, f"新增用户 {username}，角色 {role}", ip_address)


def handle_user_role_save(current, form, ip_address):
    require_permission(current, can_manage_users(current))
    user_id = safe_int(form.get("user_id", [0])[0])
    role = form.get("role", ["viewer"])[0]
    if role not in VALID_ROLES:
        raise ValueError("角色不合法")
    conn = get_conn()
    target = conn.execute("SELECT * FROM admin_user WHERE id=?", (user_id,)).fetchone()
    if not target:
        conn.close()
        raise ValueError("用户不存在")
    conn.execute("UPDATE admin_user SET role=?, updated_at=? WHERE id=?", (role, now_str(), user_id))
    conn.commit()
    conn.close()
    write_audit(current.get("username"), "UPDATE_USER_ROLE", "admin_user", target["username"], f"修改用户 {target['username']} 角色为 {role}", ip_address)


def handle_user_reset_password(current, form, ip_address):
    require_permission(current, can_manage_users(current))
    user_id = safe_int(form.get("user_id", [0])[0])
    password = form.get("password", [""])[0]
    if not password:
        raise ValueError("新密码不能为空")
    conn = get_conn()
    target = conn.execute("SELECT * FROM admin_user WHERE id=?", (user_id,)).fetchone()
    if not target:
        conn.close()
        raise ValueError("用户不存在")
    conn.execute("UPDATE admin_user SET password_hash=?, updated_at=? WHERE id=?", (hash_password(password), now_str(), user_id))
    conn.commit()
    conn.close()
    write_audit(current.get("username"), "RESET_USER_PASSWORD", "admin_user", target["username"], f"重置用户 {target['username']} 密码", ip_address)


def handle_user_toggle_enabled(current, form, ip_address):
    require_permission(current, can_manage_users(current))
    user_id = safe_int(form.get("user_id", [0])[0])
    enabled = 1 if safe_int(form.get("enabled", [1])[0], 1) else 0
    conn = get_conn()
    target = conn.execute("SELECT * FROM admin_user WHERE id=?", (user_id,)).fetchone()
    if not target:
        conn.close()
        raise ValueError("用户不存在")
    if target["username"] == current.get("username") and enabled == 0:
        conn.close()
        raise ValueError("不能禁用当前登录用户")
    conn.execute("UPDATE admin_user SET enabled=?, updated_at=? WHERE id=?", (enabled, now_str(), user_id))
    conn.commit()
    conn.close()
    write_audit(current.get("username"), "TOGGLE_USER", "admin_user", target["username"], f"{'启用' if enabled else '禁用'}用户 {target['username']}", ip_address)


def export_config_json(production_id):
    conn = get_conn()
    prod = conn.execute("SELECT * FROM production_config WHERE id=?", (production_id,)).fetchone()
    if not prod:
        conn.close()
        return None
    items = conn.execute("SELECT * FROM measurement_item_config WHERE production_id=? ORDER BY id", (production_id,)).fetchall()
    result = {
        "config_version": "1.0",
        "export_time": now_str(),
        "production": {k: prod[k] for k in prod.keys() if k not in ("id", "created_at", "updated_at")},
        "measurement_items": []
    }
    for item in items:
        metrics = conn.execute("SELECT * FROM metric_config WHERE item_id=? ORDER BY sort_order, id", (item["id"],)).fetchall()
        item_obj = {k: item[k] for k in item.keys() if k not in ("id", "production_id", "created_at", "updated_at", "last_collect_time", "last_collect_status")}
        item_obj["metrics"] = [{k: m[k] for k in m.keys() if k not in ("id", "item_id", "created_at", "updated_at")} for m in metrics]
        result["measurement_items"].append(item_obj)
    conn.close()
    return result


def export_all_config_json():
    conn = get_conn()
    productions = conn.execute("SELECT * FROM production_config ORDER BY id").fetchall()
    result = {
        "config_version": "2.2",
        "export_scope": "all_productions",
        "export_time": now_str(),
        "productions": []
    }
    for prod in productions:
        items = conn.execute("SELECT * FROM measurement_item_config WHERE production_id=? ORDER BY id", (prod["id"],)).fetchall()
        prod_obj = {
            "production": {k: prod[k] for k in prod.keys() if k not in ("id", "created_at", "updated_at")},
            "measurement_items": []
        }
        for item in items:
            metrics = conn.execute("SELECT * FROM metric_config WHERE item_id=? ORDER BY sort_order, id", (item["id"],)).fetchall()
            item_obj = {k: item[k] for k in item.keys() if k not in ("id", "production_id", "created_at", "updated_at", "last_collect_time", "last_collect_status")}
            item_obj["metrics"] = [{k: m[k] for k in m.keys() if k not in ("id", "item_id", "created_at", "updated_at")} for m in metrics]
            prod_obj["measurement_items"].append(item_obj)
        result["productions"].append(prod_obj)
    conn.close()
    return result


def page_import_config(user, message=""):
    return base_layout("导入配置", f"""
    <h1>导入配置 JSON</h1>
    <div class="card">
      {f'<div class="success">{h(message)}</div>' if message else ''}
      <form method="post" action="/import_config">
        <textarea name="config_json" rows="22" style="width:100%" placeholder="粘贴从平台导出的 JSON 配置"></textarea><br><br>
        <button type="submit">导入配置</button>
      </form>
      <p class="note">支持单个生产编号配置 JSON，也支持“导出全部配置”得到的全量 JSON。若生产编号已存在，会阻止重复导入。</p>
    </div>
    """, user)


def insert_config_payload(cur, cfg):
    prod = cfg.get("production") or {}
    code = prod.get("production_code")
    if not code:
        raise ValueError("配置中缺少 production.production_code")
    exists = cur.execute("SELECT id FROM production_config WHERE production_code=?", (code,)).fetchone()
    if exists:
        raise ValueError(f"生产编号已存在：{code}")
    cur.execute("""
    INSERT INTO production_config (production_code, production_name, product_model, process_version, description, status, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (code, prod.get("production_name"), prod.get("product_model"), prod.get("process_version"), prod.get("description"), prod.get("status", "enabled"), now_str()))
    production_id = cur.lastrowid
    for item in cfg.get("measurement_items", []):
        cur.execute("""
        INSERT INTO measurement_item_config (
            production_id, item_name, process_step, process_step_column, execution_time_text, equipment_name,
            data_source_type, data_source_path, excel_sheet_name, image_parse_config_json, header_row_index, csv_encoding, delimiter, production_code_column,
            scan_frequency_seconds, enabled, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            production_id, item.get("item_name"), item.get("process_step"), item.get("process_step_column", ""), item.get("execution_time_text"), item.get("equipment_name"),
            item.get("data_source_type", "auto"), item.get("data_source_path"), item.get("excel_sheet_name"), item.get("image_parse_config_json", ""), item.get("header_row_index", 1), item.get("csv_encoding", "auto"), item.get("delimiter", ","), item.get("production_code_column", "生产编号"),
            item.get("scan_frequency_seconds", 60), item.get("enabled", 1), now_str()
        ))
        item_id = cur.lastrowid
        for m in item.get("metrics", []):
            cur.execute("""
            INSERT INTO metric_config (item_id, metric_name, source_column, unit, data_type, target, lsl, usl, lcl, ucl, enabled, sort_order, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (item_id, m.get("metric_name"), m.get("source_column"), m.get("unit"), m.get("data_type", "number"), m.get("target"), m.get("lsl"), m.get("usl"), m.get("lcl"), m.get("ucl"), m.get("enabled", 1), m.get("sort_order", 0), now_str()))
    return production_id


def import_config(json_text):
    cfg = json.loads(json_text)
    conn = get_conn()
    cur = conn.cursor()
    if cfg.get("export_scope") == "all_productions" or isinstance(cfg.get("productions"), list):
        imported = []
        for payload in cfg.get("productions", []):
            imported.append(insert_config_payload(cur, payload))
        conn.commit()
        conn.close()
        return {"mode": "all", "count": len(imported), "ids": imported}
    production_id = insert_config_payload(cur, cfg)
    conn.commit()
    conn.close()
    return {"mode": "single", "count": 1, "ids": [production_id]}



# ==========================================================
# Template wizard
# ==========================================================

def _decode_template_bytes(data: bytes, encoding: str = "auto"):
    encodings = []
    configured = (encoding or "auto").strip().lower()
    if configured and configured != "auto":
        encodings.append(configured)
    for enc in ["utf-8-sig", "utf-8", "gb18030", "gbk", "cp936", "big5"]:
        if enc not in encodings:
            encodings.append(enc)
    last_error = None
    for enc in encodings:
        try:
            return data.decode(enc, errors="strict"), enc
        except UnicodeDecodeError as ex:
            last_error = ex
            continue
    return data.decode("gb18030", errors="replace"), "gb18030(errors=replace)"


def parse_template_csv_from_text(csv_text: str, delimiter: str = ",", header_row_index: int = 1):
    if delimiter == "\\t":
        delimiter = "\t"
    header_row_index = max(1, safe_int(header_row_index, 1))
    reader = csv.reader(io.StringIO(csv_text), delimiter=delimiter or ",")
    all_rows = list(reader)
    if len(all_rows) < header_row_index:
        raise ValueError(f"文件行数不足，无法读取第 {header_row_index} 行作为表头")
    headers = [str(x).strip() for x in all_rows[header_row_index - 1]]
    if not any(headers):
        raise ValueError("表头为空，请确认表头所在行是否正确")
    if len([x for x in headers if x]) != len(set([x for x in headers if x])):
        raise ValueError("表头存在重复字段，请先处理模板文件")
    preview_rows = []
    for raw in all_rows[header_row_index:header_row_index + 5]:
        if not any(str(x).strip() for x in raw):
            continue
        row = {}
        for idx, name in enumerate(headers):
            if name:
                row[name] = raw[idx] if idx < len(raw) else ""
        preview_rows.append(row)
    return headers, preview_rows


def page_templates(user):
    conn = get_conn()
    templates = conn.execute("""
        SELECT t.*,
               COUNT(tm.id) AS metric_count
        FROM template_config t
        LEFT JOIN template_metric_config tm ON tm.template_id=t.id
        GROUP BY t.id
        ORDER BY t.id DESC
    """).fetchall()
    conn.close()
    manage = can_manage_config(user)
    rows_html = "".join(f"""
    <tr>
      <td>{h(r['template_name'])}</td>
      <td>{h(r['template_version'])}</td>
      <td>{h(r['data_source_type'])}</td>
      <td>{h(r['production_code_column'])}</td>
      <td>{h(r['process_step_column'] if 'process_step_column' in r.keys() else '')}</td>
      <td>{r['metric_count']}</td>
      <td>{h(r['updated_at'] or r['created_at'])}</td>
      <td class="actions">
        <a class="btn secondary" href="/template_detail?id={r['id']}">查看</a>
        {f'<a class="btn secondary" href="/template_edit?id={r["id"]}">编辑</a>' if manage else ''}
        {f'''<form class="inline-form" method="post" action="/template_delete" onsubmit="return confirm('确认删除该模板？已经套用生成的量测项不会被删除。')">
          <input type="hidden" name="template_id" value="{r['id']}">
          <button class="danger" type="submit">删除</button>
        </form>''' if manage else ''}
      </td>
    </tr>
    """ for r in templates) or "<tr><td colspan='8'>暂无模板</td></tr>"
    return base_layout("模板库", f"""
    <h1>模板库</h1>
    <div class="card">
      {f'<a class="btn" href="/template_upload">上传/粘贴模板并生成字段映射</a>' if manage else ''}
      <p class="note">模板用于保存“生产编号字段 + 工序字段 + 量测指标字段”的映射。后续新增生产编号时，可一键套用模板生成量测项和指标配置。</p>
    </div>
    <div class="card">
      <h2>模板列表</h2>
      <div class="table-wrap"><table>
        <tr><th>模板名称</th><th>版本</th><th>数据源类型</th><th>生产编号字段</th><th>工序字段</th><th>指标数量</th><th>更新时间</th><th>操作</th></tr>
        {rows_html}
      </table></div>
    </div>
    """, user)


def page_template_upload(user, message=""):
    sample = "生产编号,贴装工序,Dx,Dy,Rx,Ry,Rz\\nTEST002,PL2toPL3,0.5,0.7,200,20,10\\nTEST002,PL1toPL2,0.2,-0.1,100,50,15\\n"
    return base_layout("创建模板", f"""
    <h1>创建字段映射模板</h1>
    <div class="card">
      {f'<div class="success">{h(message)}</div>' if message else ''}
      <form method="post" action="/template_parse" enctype="multipart/form-data">
        <div class="form-grid">
          <label>模板名称 *</label><input name="template_name" required placeholder="例如 DxDyRz 标准模板">
          <label>模板版本</label><input name="template_version" value="v1.0">
          <label>数据源类型</label><select name="data_source_type"><option value="excel">Excel 多Sheet</option><option value="csv">CSV</option><option value="auto">自动判断</option></select>
          <label>Excel Sheet 名称</label><input name="excel_sheet_name" placeholder="可空；上传 xlsx 后若为空会先让你选择 Sheet">
          <label>CSV 编码</label><select name="encoding"><option value="auto">auto 自动识别</option><option value="utf-8-sig">utf-8-sig</option><option value="utf-8">utf-8</option><option value="gb18030">gb18030</option><option value="gbk">gbk</option></select>
          <label>CSV 分隔符</label><select name="delimiter"><option value=",">逗号 ,</option><option value="\\t">Tab</option><option value=";">分号 ;</option></select>
          <label>表头所在行</label><input name="header_row_index" value="1">
          <label>上传模板文件</label><input type="file" name="template_file" accept=".csv,.txt,.xlsx,.xlsm">
          <label>或粘贴 CSV 内容</label><textarea name="csv_text" rows="7">{h(sample)}</textarea>
          <label>描述</label><textarea name="description" rows="3" placeholder="可填写适用产品、设备、注意事项"></textarea>
        </div>
        <br><button type="submit">读取表头，进入字段选择</button>
      </form>
      <p class="note">支持 CSV 和 Excel xlsx/xlsm。Excel 文件可以包含多个 Sheet；如果未填写 Sheet 名称，系统会先显示 Sheet 列表让你选择。</p>
    </div>
    """, user)


def page_template_sheet_select(user, form, file_info):
    file_bytes = file_info.get("content") if file_info else b""
    if not file_bytes:
        raise ValueError("请上传 Excel 模板文件")
    token, cached_path = save_template_cache_file(file_bytes, file_info.get("filename", "template.xlsx"))
    sheets = xlsx_list_sheets_from_bytes(file_bytes)
    if not sheets:
        raise ValueError("Excel 文件中没有可读取的 Sheet")
    options = "".join(f'<option value="{h(name)}">{h(name)}</option>' for name, _target in sheets)
    hidden = ""
    for key in ["template_name","template_version","data_source_type","encoding","delimiter","header_row_index","description"]:
        hidden += f'<input type="hidden" name="{key}" value="{h(form.get(key, [""])[0])}">'
    return base_layout("选择Sheet", f"""
    <h1>选择 Excel Sheet</h1>
    <div class="card">
      <p class="note">已识别到 {len(sheets)} 个 Sheet。请选择用于建立字段映射模板的 Sheet。</p>
      <form method="post" action="/template_parse">
        {hidden}
        <input type="hidden" name="template_cache_token" value="{h(token)}">
        <input type="hidden" name="data_source_type" value="excel">
        <div class="form-grid">
          <label>Excel Sheet *</label><select name="excel_sheet_name">{options}</select>
        </div>
        <br><button type="submit">读取该 Sheet 表头</button> <a class="btn secondary" href="/template_upload">返回</a>
      </form>
    </div>
    """, user)


def page_template_mapping(user, form, files):
    template_name = form.get("template_name", [""])[0].strip()
    template_version = form.get("template_version", ["v1.0"])[0].strip() or "v1.0"
    data_source_type = form.get("data_source_type", ["auto"])[0].strip() or "auto"
    encoding = form.get("encoding", ["auto"])[0].strip() or "auto"
    delimiter = form.get("delimiter", [","])[0]
    header_row_index = max(1, safe_int(form.get("header_row_index", [1])[0], 1))
    description = form.get("description", [""])[0].strip()
    excel_sheet_name = form.get("excel_sheet_name", [""])[0].strip()
    csv_text = form.get("csv_text", [""])[0]

    file_info = files.get("template_file")
    cache_token = form.get("template_cache_token", [""])[0].strip()

    source_label = "text-input"
    headers = []
    preview_rows = []

    file_bytes = None
    filename = ""
    if file_info and file_info.get("content"):
        file_bytes = file_info["content"]
        filename = file_info.get("filename", "")
    elif cache_token:
        cached_path = get_template_cache_path(cache_token)
        if not cached_path:
            raise ValueError("模板缓存文件不存在，请重新上传。")
        file_bytes = Path(cached_path).read_bytes()
        filename = cached_path

    suffix = Path(filename or "").suffix.lower()
    inferred_type = data_source_type
    if inferred_type == "auto":
        inferred_type = "excel" if suffix in (".xlsx", ".xlsm", ".xls") else "csv"

    if inferred_type == "excel":
        if suffix == ".xls":
            raise ValueError("当前版本支持 .xlsx/.xlsm；旧版 .xls 请先另存为 .xlsx。")
        if not file_bytes:
            raise ValueError("请选择 Excel 模板文件。")
        if not excel_sheet_name:
            return page_template_sheet_select(user, form, file_info)
        headers, preview_rows, source_label = _parse_xlsx_rows_from_bytes(file_bytes, excel_sheet_name, header_row_index)
        data_source_type = "excel"
    else:
        used_encoding = "text-input"
        if file_bytes:
            csv_text, used_encoding = _decode_template_bytes(file_bytes, encoding)
        elif not csv_text.strip():
            raise ValueError("请上传 CSV/Excel 模板文件，或粘贴 CSV 内容")
        headers, preview_rows = parse_template_csv_from_text(csv_text, delimiter, header_row_index)
        source_label = used_encoding
        data_source_type = "csv"

    if not headers:
        raise ValueError("没有识别到表头字段，请确认表头所在行是否正确。")

    options = "".join(f'<option value="{h(col)}" {"selected" if col in ("生产编号","PN","ProductCode","production_code") else ""}>{h(col)}</option>' for col in headers)
    process_options = '<option value="">不使用工序字段</option>' + "".join(
        f'<option value="{h(col)}" {"selected" if col in ("工序","工序名","量测工序","贴装工序","process_step","ProcessStep","Step") else ""}>{h(col)}</option>'
        for col in headers
    )

    metric_checks = ""
    ignore_cols = ("生产编号", "PN", "ProductCode", "production_code", "工序", "工序名", "量测工序", "贴装工序", "process_step", "ProcessStep", "Step", "备注", "时间", "量测时间", "日期", "设备", "人员")
    for col in headers:
        default_checked = "" if col in ignore_cols else "checked"
        metric_checks += f"""
        <label style="display:inline-block;margin:6px 18px 6px 0;">
          <input type="checkbox" name="metric_columns" value="{h(col)}" {default_checked}> {h(col)}
        </label>
        """

    preview_header = "".join(f"<th>{h(x)}</th>" for x in headers)
    preview_body = "".join("<tr>" + "".join(f"<td>{h(row.get(col,''))}</td>" for col in headers) + "</tr>" for row in preview_rows[:10])
    if not preview_body:
        preview_body = f"<tr><td colspan='{len(headers)}'>暂无数据行预览，仅识别到表头</td></tr>"

    hidden_fields = f"""
      <input type="hidden" name="template_name" value="{h(template_name)}">
      <input type="hidden" name="template_version" value="{h(template_version)}">
      <input type="hidden" name="data_source_type" value="{h(data_source_type)}">
      <input type="hidden" name="encoding" value="{h(encoding)}">
      <input type="hidden" name="delimiter" value="{h(delimiter)}">
      <input type="hidden" name="header_row_index" value="{h(header_row_index)}">
      <input type="hidden" name="excel_sheet_name" value="{h(excel_sheet_name)}">
      <input type="hidden" name="description" value="{h(description)}">
      <input type="hidden" name="sample_fields_json" value="{h(json.dumps(headers, ensure_ascii=False))}">
    """

    return base_layout("字段映射确认", f"""
    <h1>字段映射确认</h1>
    <div class="card">
      <p class="note">已识别 {len(headers)} 个字段；数据源类型：{h(data_source_type)}；Sheet/编码：{h(excel_sheet_name or source_label)}。请选择生产编号字段、可选工序字段，以及哪些字段作为量测指标。</p>
      <form method="post" action="/template_save">
        {hidden_fields}
        <div class="form-grid">
          <label>模板名称</label><input value="{h(template_name)}" disabled>
          <label>模板版本</label><input value="{h(template_version)}" disabled>
          <label>数据源类型</label><input value="{h(data_source_type)}" disabled>
          <label>Excel Sheet</label><input value="{h(excel_sheet_name)}" disabled>
          <label>生产编号字段 *</label><select name="production_code_column">{options}</select>
          <label>工序字段</label><select name="process_step_column">{process_options}</select>
          <label>量测指标字段 *</label><div>{metric_checks}</div>
        </div>
        <br><button type="submit">保存模板</button> <a class="btn secondary" href="/template_upload">返回重选</a>
      </form>
    </div>
    <div class="card">
      <h2>数据预览</h2>
      <div class="table-wrap"><table><tr>{preview_header}</tr>{preview_body}</table></div>
    </div>
    """, user)


def page_template_detail(user, template_id):
    conn = get_conn()
    t = conn.execute("SELECT * FROM template_config WHERE id=?", (template_id,)).fetchone()
    if not t:
        conn.close()
        return base_layout("未找到", "<h1>模板不存在</h1>", user)
    metrics = conn.execute("SELECT * FROM template_metric_config WHERE template_id=? ORDER BY sort_order,id", (template_id,)).fetchall()
    apply_logs = conn.execute("""
        SELECT * FROM template_apply_log
        WHERE template_id=?
        ORDER BY applied_at DESC, id DESC
        LIMIT 20
    """, (template_id,)).fetchall()
    conn.close()
    rows_html = "".join(f"""
    <tr>
      <td>{h(m['metric_name'])}</td><td>{h(m['source_column'])}</td><td>{h(m['data_type'])}</td><td>{h(m['unit'])}</td>
      <td>{h(m['target'])}</td><td>{h(m['lsl'])}</td><td>{h(m['usl'])}</td><td>{h(m['lcl'])}</td><td>{h(m['ucl'])}</td>
    </tr>
    """ for m in metrics) or "<tr><td colspan='9'>暂无指标</td></tr>"
    apply_rows = "".join(f"""
    <tr>
      <td>{h(log['applied_at'])}</td><td>{h(log['production_code'])}</td><td>{h(log['item_id'])}</td>
      <td>{h(log['applied_by'])}</td><td>{h(log['template_name'] or t['template_name'])}</td><td>{h(log['template_version'] or t['template_version'])}</td>
      <td><details><summary>查看 snapshot</summary><pre>{h(log['snapshot_json'] or '{}')}</pre></details></td>
    </tr>
    """ for log in apply_logs) or "<tr><td colspan='7'>暂无套用记录</td></tr>"
    fields = json.loads(t["sample_fields_json"] or "[]")
    return base_layout("模板详情", f"""
    <h1>模板详情：{h(t['template_name'])}</h1>
    <div class="card">
      <p><b>版本：</b>{h(t['template_version'])}</p>
      <p><b>数据源类型：</b>{h(t['data_source_type'])}</p>
      <p><b>Excel Sheet：</b>{h(t['excel_sheet_name'] if 'excel_sheet_name' in t.keys() else '')}</p>
      <p><b>生产编号字段：</b>{h(t['production_code_column'])}</p>
      <p><b>工序字段：</b>{h(t['process_step_column'] if 'process_step_column' in t.keys() else '')}</p>
      <p><b>表头字段：</b>{h(', '.join(fields))}</p>
      <p><b>描述：</b>{h(t['description'])}</p>
      {f'<a class="btn secondary" href="/template_edit?id={template_id}">编辑模板</a>' if can_manage_config(user) else ''}
      <a class="btn secondary" href="/templates">返回模板库</a>
    </div>
    <div class="card">
      <h2>指标字段</h2>
      <div class="table-wrap"><table><tr><th>指标名</th><th>源字段名</th><th>类型</th><th>单位</th><th>Target</th><th>MS2下限</th><th>MS2上限</th><th>MS3下限</th><th>MS3上限</th></tr>{rows_html}</table></div>
    </div>
    <div class="card">
      <h2>最近套用记录 Snapshot</h2>
      <div class="table-wrap"><table><tr><th>套用时间</th><th>生产编号</th><th>量测项ID</th><th>操作人</th><th>模板名</th><th>模板版本</th><th>Snapshot</th></tr>{apply_rows}</table></div>
    </div>
    """, user)


def page_template_edit(user, template_id):
    conn = get_conn()
    t = conn.execute("SELECT * FROM template_config WHERE id=?", (template_id,)).fetchone()
    if not t:
        conn.close()
        return base_layout("未找到", "<h1>模板不存在</h1>", user)
    metrics = conn.execute("SELECT * FROM template_metric_config WHERE template_id=? ORDER BY sort_order,id", (template_id,)).fetchall()
    conn.close()
    fields = json.loads(t["sample_fields_json"] or "[]")
    prod_options = "".join(f'<option value="{h(col)}" {"selected" if col == t["production_code_column"] else ""}>{h(col)}</option>' for col in fields)
    process_options = '<option value="">不使用工序字段</option>' + "".join(f'<option value="{h(col)}" {"selected" if col == (t["process_step_column"] if "process_step_column" in t.keys() else "") else ""}>{h(col)}</option>' for col in fields)
    metric_rows = ""
    for m in metrics:
        metric_rows += f"""
        <tr>
          <td><input type="hidden" name="metric_ids" value="{m['id']}"><input name="metric_name_{m['id']}" value="{h(m['metric_name'])}" style="min-width:110px"></td>
          <td><input name="source_column_{m['id']}" value="{h(m['source_column'])}" style="min-width:110px"></td>
          <td><input name="unit_{m['id']}" value="{h(m['unit'])}" style="min-width:80px"></td>
          <td><select name="data_type_{m['id']}"><option value="number" {'selected' if m['data_type']=='number' else ''}>number</option><option value="text" {'selected' if m['data_type']=='text' else ''}>text</option></select></td>
          <td><input name="target_{m['id']}" value="{h(m['target'])}" style="min-width:80px"></td>
          <td><input name="lsl_{m['id']}" value="{h(m['lsl'])}" style="min-width:80px"></td>
          <td><input name="usl_{m['id']}" value="{h(m['usl'])}" style="min-width:80px"></td>
          <td><input name="lcl_{m['id']}" value="{h(m['lcl'])}" style="min-width:80px"></td>
          <td><input name="ucl_{m['id']}" value="{h(m['ucl'])}" style="min-width:80px"></td>
          <td><input name="sort_order_{m['id']}" value="{h(m['sort_order'])}" type="number" style="min-width:70px"></td>
          <td><label><input type="checkbox" name="delete_metric_ids" value="{m['id']}"> 删除</label></td>
        </tr>
        """
    metric_rows = metric_rows or "<tr><td colspan='11'>暂无指标</td></tr>"
    return base_layout("编辑模板", f"""
    <h1>编辑模板：{h(t['template_name'])}</h1>
    <div class="card">
      <form method="post" action="/template_update">
        <input type="hidden" name="template_id" value="{template_id}">
        <div class="form-grid">
          <label>模板名称</label><input name="template_name" value="{h(t['template_name'])}" required>
          <label>模板版本</label><input name="template_version" value="{h(t['template_version'])}">
          <label>生产编号字段</label><select name="production_code_column">{prod_options}</select>
          <label>工序字段</label><select name="process_step_column">{process_options}</select>
          <label>描述</label><textarea name="description" rows="3">{h(t['description'])}</textarea>
        </div>
        <h2>指标配置</h2>
        <div class="table-wrap"><table>
          <tr><th>指标名</th><th>源字段</th><th>单位</th><th>类型</th><th>Target</th><th>MS2下限</th><th>MS2上限</th><th>MS3下限</th><th>MS3上限</th><th>排序</th><th>删除</th></tr>
          {metric_rows}
        </table></div>
        <p class="note">新增指标字段用逗号分隔，字段名会同时作为指标名和源字段名。示例：Dx,Dy,Rz,Rx,Ry,Dz</p>
        <div class="form-row"><input name="new_metric_columns" style="min-width:420px" placeholder="新增指标字段，可空"><button type="submit">保存模板</button><a class="btn secondary" href="/template_detail?id={template_id}">返回</a></div>
      </form>
    </div>
    """, user)


def page_template_apply(user, production_id):
    conn = get_conn()
    prod = conn.execute("SELECT * FROM production_config WHERE id=?", (production_id,)).fetchone()
    templates = conn.execute("""
        SELECT t.*, COUNT(tm.id) AS metric_count
        FROM template_config t
        LEFT JOIN template_metric_config tm ON tm.template_id=t.id
        GROUP BY t.id
        ORDER BY t.id DESC
    """).fetchall()
    conn.close()
    if not prod:
        return base_layout("错误", "<h1>生产编号不存在</h1>", user)
    template_options = "".join(f'<option value="{t["id"]}">{h(t["template_name"])} / {h(t["template_version"])} / {h(t["data_source_type"])} / {t["metric_count"]}个指标</option>' for t in templates)
    if not template_options:
        template_options = '<option value="">暂无模板，请先到模板库创建</option>'
    return base_layout("套用模板", f"""
    <h1>从模板新增量测项：{h(prod['production_code'])}</h1>
    <div class="card">
      <form method="post" action="/template_apply">
        <input type="hidden" name="production_id" value="{production_id}">
        <div class="form-grid">
          <label>选择模板 *</label><select name="template_id" required>{template_options}</select>
          <label>量测项名称</label><input name="item_name" placeholder="不填则使用模板名称">
          <label>固定量测工序</label><input name="process_step" placeholder="模板没有工序字段时使用，例如 MEAS_STEP_01">
          <label>量测执行时间</label><input name="execution_time_text" placeholder="例如 工序完成后 / 每日10:00">
          <label>量测设备</label><input name="equipment_name" placeholder="例如 TOOL01">
          <label>实时数据源路径 *</label><input name="data_source_path" required style="min-width:520px" placeholder="例如 \\\\192.168.1.100\\share\\result.csv">
          <label>抓取频率秒</label><input name="scan_frequency_seconds" value="60">
        </div>
        <br><button type="submit">套用模板并生成量测项</button> <a class="btn secondary" href="/items?production_id={production_id}">返回</a>
      </form>
      <p class="note">套用模板会复制模板中的生产编号字段、工序字段和指标字段到当前生产编号下。若模板配置了工序字段，采集时会按同一生产编号下的多道工序逐行入库。</p>
    </div>
    """, user)


def handle_template_save(user, form, ip_address):
    require_permission(user, can_manage_config(user))
    template_name = form.get("template_name", [""])[0].strip()
    template_version = form.get("template_version", ["v1.0"])[0].strip() or "v1.0"
    data_source_type = form.get("data_source_type", ["csv"])[0].strip() or "csv"
    encoding = form.get("encoding", ["auto"])[0].strip() or "auto"
    delimiter = form.get("delimiter", [","])[0]
    header_row_index = max(1, safe_int(form.get("header_row_index", [1])[0], 1))
    description = form.get("description", [""])[0].strip()
    excel_sheet_name = form.get("excel_sheet_name", [""])[0].strip()
    production_code_column = form.get("production_code_column", [""])[0].strip()
    process_step_column = form.get("process_step_column", [""])[0].strip()
    metric_columns = [x.strip() for x in form.get("metric_columns", []) if x.strip()]
    sample_fields_json = form.get("sample_fields_json", ["[]"])[0]

    if not template_name:
        raise ValueError("模板名称不能为空")
    if not production_code_column:
        raise ValueError("必须选择生产编号字段")
    if not metric_columns:
        raise ValueError("至少选择一个量测指标字段")
    if production_code_column in metric_columns:
        raise ValueError("生产编号字段不能同时作为量测指标")
    if process_step_column and process_step_column in metric_columns:
        raise ValueError("工序字段不能同时作为量测指标")

    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO template_config (
            template_name, template_version, data_source_type, header_row_index, delimiter, encoding, excel_sheet_name,
            production_code_column, process_step_column, sample_fields_json, description, created_at, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        template_name, template_version, data_source_type, header_row_index, delimiter, encoding, excel_sheet_name,
        production_code_column, process_step_column, sample_fields_json, description, now_str(), now_str()
    ))
    template_id = cur.lastrowid
    for idx, col in enumerate(metric_columns):
        cur.execute("""
            INSERT INTO template_metric_config (
                template_id, metric_name, source_column, data_type, unit, sort_order, created_at, updated_at
            ) VALUES (?, ?, ?, 'number', '', ?, ?, ?)
        """, (template_id, col, col, idx, now_str(), now_str()))
    conn.commit()
    conn.close()
    write_audit(user.get("username"), "SAVE_TEMPLATE", "template_config", template_id, f"保存模板 {template_name}，指标：{', '.join(metric_columns)}", ip_address)
    return template_id


def handle_template_update(user, form, ip_address):
    require_permission(user, can_manage_config(user))
    template_id = safe_int(form.get("template_id", [0])[0])
    template_name = form.get("template_name", [""])[0].strip()
    template_version = form.get("template_version", ["v1.0"])[0].strip() or "v1.0"
    production_code_column = form.get("production_code_column", [""])[0].strip()
    process_step_column = form.get("process_step_column", [""])[0].strip()
    description = form.get("description", [""])[0].strip()
    if not template_name:
        raise ValueError("模板名称不能为空")
    if not production_code_column:
        raise ValueError("生产编号字段不能为空")
    if process_step_column and process_step_column == production_code_column:
        raise ValueError("工序字段不能与生产编号字段相同")
    delete_ids = {safe_int(x) for x in form.get("delete_metric_ids", [])}
    conn = get_conn()
    cur = conn.cursor()
    t = cur.execute("SELECT * FROM template_config WHERE id=?", (template_id,)).fetchone()
    if not t:
        conn.close()
        raise ValueError("模板不存在")
    cur.execute("""
        UPDATE template_config
        SET template_name=?, template_version=?, production_code_column=?, process_step_column=?, description=?, updated_at=?
        WHERE id=?
    """, (template_name, template_version, production_code_column, process_step_column, description, now_str(), template_id))
    metric_ids = [safe_int(x) for x in form.get("metric_ids", [])]
    for metric_id in metric_ids:
        if metric_id in delete_ids:
            cur.execute("DELETE FROM template_metric_config WHERE id=? AND template_id=?", (metric_id, template_id))
            continue
        metric_name = form.get(f"metric_name_{metric_id}", [""])[0].strip()
        source_column = form.get(f"source_column_{metric_id}", [""])[0].strip()
        if not metric_name or not source_column:
            continue
        cur.execute("""
            UPDATE template_metric_config
            SET metric_name=?, source_column=?, unit=?, data_type=?, target=?, lsl=?, usl=?, lcl=?, ucl=?, sort_order=?, updated_at=?
            WHERE id=? AND template_id=?
        """, (
            metric_name, source_column,
            form.get(f"unit_{metric_id}", [""])[0].strip(),
            form.get(f"data_type_{metric_id}", ["number"])[0],
            safe_float(form.get(f"target_{metric_id}", [""])[0]),
            safe_float(form.get(f"lsl_{metric_id}", [""])[0]),
            safe_float(form.get(f"usl_{metric_id}", [""])[0]),
            safe_float(form.get(f"lcl_{metric_id}", [""])[0]),
            safe_float(form.get(f"ucl_{metric_id}", [""])[0]),
            safe_int(form.get(f"sort_order_{metric_id}", [0])[0], 0),
            now_str(), metric_id, template_id
        ))
    new_cols = [x.strip() for x in form.get("new_metric_columns", [""])[0].replace("，", ",").split(",") if x.strip()]
    current_max = cur.execute("SELECT COALESCE(MAX(sort_order), -1) AS m FROM template_metric_config WHERE template_id=?", (template_id,)).fetchone()["m"]
    for idx, col in enumerate(new_cols, start=current_max + 1):
        if col == production_code_column or (process_step_column and col == process_step_column):
            continue
        cur.execute("""
            INSERT INTO template_metric_config (template_id, metric_name, source_column, data_type, unit, sort_order, created_at, updated_at)
            VALUES (?, ?, ?, 'number', '', ?, ?, ?)
        """, (template_id, col, col, idx, now_str(), now_str()))
    conn.commit()
    conn.close()
    write_audit(user.get("username"), "UPDATE_TEMPLATE", "template_config", template_id, f"编辑模板 {template_name}", ip_address)
    return template_id


def delete_template_config(template_id):
    conn = get_conn()
    cur = conn.cursor()
    t = cur.execute("SELECT * FROM template_config WHERE id=?", (template_id,)).fetchone()
    if not t:
        conn.close()
        raise ValueError("模板不存在")
    cur.execute("DELETE FROM template_metric_config WHERE template_id=?", (template_id,))
    cur.execute("DELETE FROM template_config WHERE id=?", (template_id,))
    conn.commit()
    conn.close()
    return t["template_name"]


def build_template_apply_snapshot(template_row, metric_rows, applied_item_config):
    return {
        "template_id": template_row["id"],
        "template_name": template_row["template_name"],
        "template_version": template_row["template_version"],
        "data_source_type": template_row["data_source_type"],
        "production_code_column": template_row["production_code_column"],
        "process_step_column": template_row["process_step_column"] if "process_step_column" in template_row.keys() else "",
        "excel_sheet_name": template_row["excel_sheet_name"] if "excel_sheet_name" in template_row.keys() else "",
        "header_row_index": template_row["header_row_index"],
        "metrics": [
            {
                "metric_name": m["metric_name"],
                "source_column": m["source_column"],
                "data_type": m["data_type"],
                "unit": m["unit"],
                "target": m["target"],
                "lsl": m["lsl"],
                "usl": m["usl"],
                "lcl": m["lcl"],
                "ucl": m["ucl"],
                "sort_order": m["sort_order"],
            }
            for m in metric_rows
        ],
        "applied_item_config": applied_item_config,
        "snapshot_time": now_str(),
    }


def handle_template_apply(user, form, ip_address):
    require_permission(user, can_manage_config(user))
    production_id = safe_int(form.get("production_id", [0])[0])
    template_id = safe_int(form.get("template_id", [0])[0])
    item_name_input = form.get("item_name", [""])[0].strip()
    process_step = form.get("process_step", [""])[0].strip()
    execution_time_text = form.get("execution_time_text", [""])[0].strip()
    equipment_name = form.get("equipment_name", [""])[0].strip()
    data_source_path = form.get("data_source_path", [""])[0].strip()
    scan_frequency_seconds = max(10, safe_int(form.get("scan_frequency_seconds", [60])[0], 60))
    if not data_source_path:
        raise ValueError("实时数据源路径不能为空")

    conn = get_conn()
    cur = conn.cursor()
    prod = cur.execute("SELECT * FROM production_config WHERE id=?", (production_id,)).fetchone()
    t = cur.execute("SELECT * FROM template_config WHERE id=?", (template_id,)).fetchone()
    metrics = cur.execute("SELECT * FROM template_metric_config WHERE template_id=? ORDER BY sort_order,id", (template_id,)).fetchall()
    if not prod:
        conn.close()
        raise ValueError("生产编号不存在")
    if not t:
        conn.close()
        raise ValueError("模板不存在")
    if not metrics:
        conn.close()
        raise ValueError("模板下没有指标，无法套用")
    template_process_column = (t["process_step_column"] if "process_step_column" in t.keys() else "") or ""
    template_process_column = template_process_column.strip()
    if not template_process_column and not process_step:
        conn.close()
        raise ValueError("该模板没有工序字段。套用模板时必须填写固定量测工序，否则系统会生成无工序量测项并被禁止采集。")

    item_name = item_name_input or t["template_name"]
    cur.execute("""
        INSERT INTO measurement_item_config (
            production_id, item_name, process_step, process_step_column, execution_time_text, equipment_name,
            data_source_type, data_source_path, excel_sheet_name, header_row_index, csv_encoding, delimiter, production_code_column,
            scan_frequency_seconds, enabled, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
    """, (
        production_id, item_name, process_step, template_process_column, execution_time_text, equipment_name,
        t["data_source_type"] or "auto", data_source_path, t["excel_sheet_name"] if "excel_sheet_name" in t.keys() else "",
        t["header_row_index"] if "header_row_index" in t.keys() else 1, t["encoding"] or "auto", t["delimiter"] or ",", t["production_code_column"],
        scan_frequency_seconds, now_str()
    ))
    item_id = cur.lastrowid

    for idx, m in enumerate(metrics):
        cur.execute("""
            INSERT INTO metric_config (
                item_id, metric_name, source_column, unit, data_type, target, lsl, usl, lcl, ucl, enabled, sort_order, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
        """, (
            item_id, m["metric_name"], m["source_column"], m["unit"], m["data_type"],
            m["target"], m["lsl"], m["usl"], m["lcl"], m["ucl"], idx, now_str()
        ))

    applied_item_config = {
        "item_id": item_id,
        "production_id": production_id,
        "production_code": prod["production_code"],
        "item_name": item_name,
        "process_step": process_step,
        "process_step_column": template_process_column,
        "execution_time_text": execution_time_text,
        "equipment_name": equipment_name,
        "data_source_type": t["data_source_type"] or "auto",
        "data_source_path": data_source_path,
        "excel_sheet_name": t["excel_sheet_name"] if "excel_sheet_name" in t.keys() else "",
        "header_row_index": t["header_row_index"] if "header_row_index" in t.keys() else 1,
        "csv_encoding": t["encoding"] or "auto",
        "delimiter": t["delimiter"] or ",",
        "production_code_column": t["production_code_column"],
        "scan_frequency_seconds": scan_frequency_seconds,
    }
    snapshot = build_template_apply_snapshot(t, metrics, applied_item_config)
    cur.execute("""
        INSERT INTO template_apply_log (
            template_id, production_id, production_code, item_id, applied_by, applied_at,
            template_name, template_version, snapshot_json, detail
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        template_id, production_id, prod["production_code"], item_id, user.get("username"), now_str(),
        t["template_name"], t["template_version"], json.dumps(snapshot, ensure_ascii=False),
        f"套用模板 {t['template_name']} 到量测项 {item_name}"
    ))
    conn.commit()
    conn.close()
    write_audit(user.get("username"), "APPLY_TEMPLATE", "template_config", template_id, f"套用模板到生产编号 {prod['production_code']}，生成量测项 {item_name}", ip_address)
    return item_id


def page_template_apply_bulk(user, production_id):
    conn = get_conn()
    prod = conn.execute("SELECT * FROM production_config WHERE id=?", (production_id,)).fetchone()
    templates = conn.execute("""
        SELECT t.*, COUNT(tm.id) AS metric_count
        FROM template_config t
        LEFT JOIN template_metric_config tm ON tm.template_id=t.id
        GROUP BY t.id
        ORDER BY t.id DESC
    """).fetchall()
    conn.close()
    if not prod:
        return base_layout("错误", "<h1>生产编号不存在</h1>", user)

    template_rows = "".join(f"""
      <label class="template-check-row">
        <input type="checkbox" name="template_ids" value="{t["id"]}">
        <span><b>{h(t["template_name"])}</b> / {h(t["template_version"])} / {h(t["data_source_type"])} / {t["metric_count"]} 个指标</span>
      </label>
    """ for t in templates)
    if not template_rows:
        template_rows = '<p class="note">暂无模板，请先到模板库创建。</p>'

    bulk_controls = """
      <div class="actions" style="margin:8px 0 10px">
        <button type="button" class="secondary" onclick="document.querySelectorAll('input[name=template_ids]').forEach(x=>x.checked=true)">全选模板</button>
        <button type="button" class="secondary" onclick="document.querySelectorAll('input[name=template_ids]').forEach(x=>x.checked=false)">全不选</button>
      </div>
    """ if templates else ""

    return base_layout("套用模板", f"""
    <h1>从模板新增量测项：{h(prod['production_code'])}</h1>
    <div class="card">
      <form method="post" action="/template_apply">
        <input type="hidden" name="production_id" value="{production_id}">
        <style>
          .template-check-list{{border:1px solid var(--line);border-radius:12px;padding:10px;max-height:260px;overflow:auto;background:#fff}}
          .template-check-row{{display:flex;align-items:center;gap:10px;padding:9px 8px;border-bottom:1px solid var(--line);cursor:pointer}}
          .template-check-row:last-child{{border-bottom:none}}
          .template-check-row input{{width:auto}}
        </style>
        <label>选择模板 *</label>
        {bulk_controls}
        <div class="template-check-list">{template_rows}</div>
        <p class="note">可一次勾选多个模板；每个模板会生成一个量测项。勾选多个模板时，“量测项名称”会作为前缀，例如“Demo - 模板A”。</p>
        <br>
        <div class="form-grid">
          <label>量测项名称/前缀</label><input name="item_name" placeholder="不填则使用模板名；多模板时作为前缀">
          <label>固定量测工序</label><input name="process_step" placeholder="模板没有工序字段时使用，例如 MEAS_STEP_01">
          <label>量测执行时间</label><input name="execution_time_text" placeholder="例如 工序完成后 / 每日10:00">
          <label>量测设备</label><input name="equipment_name" placeholder="例如 TOOL01">
          <label>实时数据源路径 *</label><input name="data_source_path" required style="min-width:520px" placeholder="例如 \\\\192.168.1.100\\share\\result.csv">
          <label>抓取频率秒</label><input name="scan_frequency_seconds" value="60" type="number" min="10">
        </div>
        <br><button type="submit">套用选中模板并生成量测项</button> <a class="btn secondary" href="/items?production_id={production_id}">返回</a>
      </form>
      <p class="note">套用模板会复制模板中的生产编号字段、工序字段和指标字段到当前生产编号下。若模板配置了工序字段，采集时会按同一生产编号下的多道工序逐行入库。</p>
    </div>
    """, user)


def selected_template_ids_from_form(form):
    raw_ids = form.get("template_ids", [])
    if not raw_ids:
        raw_ids = form.get("template_id", [])
    result = []
    seen = set()
    for value in raw_ids:
        template_id = safe_int(value)
        if template_id > 0 and template_id not in seen:
            result.append(template_id)
            seen.add(template_id)
    return result


def apply_template_to_production(cur, user, prod, template_id, form_values, total_templates):
    t = cur.execute("SELECT * FROM template_config WHERE id=?", (template_id,)).fetchone()
    metrics = cur.execute("SELECT * FROM template_metric_config WHERE template_id=? ORDER BY sort_order,id", (template_id,)).fetchall()
    if not t:
        raise ValueError(f"模板不存在：{template_id}")
    if not metrics:
        raise ValueError(f"模板 {t['template_name']} 下没有指标，无法套用")

    process_step = form_values["process_step"]
    template_process_column = (t["process_step_column"] if "process_step_column" in t.keys() else "") or ""
    template_process_column = template_process_column.strip()
    if not template_process_column and not process_step:
        raise ValueError(f"模板 {t['template_name']} 没有工序字段。套用时必须填写固定量测工序。")

    item_name_input = form_values["item_name"]
    if item_name_input and total_templates > 1:
        item_name = f"{item_name_input} - {t['template_name']}"
    else:
        item_name = item_name_input or t["template_name"]

    cur.execute("""
        INSERT INTO measurement_item_config (
            production_id, item_name, process_step, process_step_column, execution_time_text, equipment_name,
            data_source_type, data_source_path, excel_sheet_name, header_row_index, csv_encoding, delimiter, production_code_column,
            scan_frequency_seconds, enabled, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?)
    """, (
        prod["id"], item_name, process_step, template_process_column, form_values["execution_time_text"], form_values["equipment_name"],
        t["data_source_type"] or "auto", form_values["data_source_path"], t["excel_sheet_name"] if "excel_sheet_name" in t.keys() else "",
        t["header_row_index"] if "header_row_index" in t.keys() else 1, t["encoding"] or "auto", t["delimiter"] or ",", t["production_code_column"],
        form_values["scan_frequency_seconds"], now_str()
    ))
    item_id = cur.lastrowid

    for idx, m in enumerate(metrics):
        cur.execute("""
            INSERT INTO metric_config (
                item_id, metric_name, source_column, unit, data_type, target, lsl, usl, lcl, ucl, enabled, sort_order, updated_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, ?, ?)
        """, (
            item_id, m["metric_name"], m["source_column"], m["unit"], m["data_type"],
            m["target"], m["lsl"], m["usl"], m["lcl"], m["ucl"], idx, now_str()
        ))

    applied_item_config = {
        "item_id": item_id,
        "production_id": prod["id"],
        "production_code": prod["production_code"],
        "item_name": item_name,
        "process_step": process_step,
        "process_step_column": template_process_column,
        "execution_time_text": form_values["execution_time_text"],
        "equipment_name": form_values["equipment_name"],
        "data_source_type": t["data_source_type"] or "auto",
        "data_source_path": form_values["data_source_path"],
        "excel_sheet_name": t["excel_sheet_name"] if "excel_sheet_name" in t.keys() else "",
        "header_row_index": t["header_row_index"] if "header_row_index" in t.keys() else 1,
        "csv_encoding": t["encoding"] or "auto",
        "delimiter": t["delimiter"] or ",",
        "production_code_column": t["production_code_column"],
        "scan_frequency_seconds": form_values["scan_frequency_seconds"],
    }
    snapshot = build_template_apply_snapshot(t, metrics, applied_item_config)
    cur.execute("""
        INSERT INTO template_apply_log (
            template_id, production_id, production_code, item_id, applied_by, applied_at,
            template_name, template_version, snapshot_json, detail
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        template_id, prod["id"], prod["production_code"], item_id, user.get("username"), now_str(),
        t["template_name"], t["template_version"], json.dumps(snapshot, ensure_ascii=False),
        f"套用模板 {t['template_name']} 到量测项 {item_name}"
    ))
    return {"item_id": item_id, "template_id": template_id, "template_name": t["template_name"], "item_name": item_name}


def handle_template_apply_bulk(user, form, ip_address):
    require_permission(user, can_manage_config(user))
    production_id = safe_int(form.get("production_id", [0])[0])
    template_ids = selected_template_ids_from_form(form)
    form_values = {
        "item_name": form.get("item_name", [""])[0].strip(),
        "process_step": form.get("process_step", [""])[0].strip(),
        "execution_time_text": form.get("execution_time_text", [""])[0].strip(),
        "equipment_name": form.get("equipment_name", [""])[0].strip(),
        "data_source_path": form.get("data_source_path", [""])[0].strip(),
        "scan_frequency_seconds": max(10, safe_int(form.get("scan_frequency_seconds", [60])[0], 60)),
    }
    if not template_ids:
        raise ValueError("请至少选择一个模板")
    if not form_values["data_source_path"]:
        raise ValueError("实时数据源路径不能为空")

    conn = get_conn()
    cur = conn.cursor()
    try:
        prod = cur.execute("SELECT * FROM production_config WHERE id=?", (production_id,)).fetchone()
        if not prod:
            raise ValueError("生产编号不存在")
        applied = [
            apply_template_to_production(cur, user, prod, template_id, form_values, len(template_ids))
            for template_id in template_ids
        ]
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

    item_names = ", ".join(x["item_name"] for x in applied)
    write_audit(user.get("username"), "APPLY_TEMPLATE", "template_config", ",".join(str(x) for x in template_ids), f"套用 {len(applied)} 个模板到生产编号 {prod['production_code']}，生成量测项 {item_names}", ip_address)
    return applied


def page_about(user):
    return base_layout("说明", f"""
    <h1>说明</h1>
    <div class="card">
      <h2>当前版本支持的 CSV / Excel 格式</h2>
      <p>默认支持“每个生产编号一行”的总表结构，CSV 或 Excel 指定 Sheet 均可：</p>
      <pre>生产编号,Dx1,Dy1,Dx2,Dy2,Rz
PROD_A_V1,1.2,2.3,1.1,2.1,0.8
PROD_B_V1,1.5,2.2,1.4,2.0,0.7</pre>
      <p>量测项中配置 <b>生产编号字段名=生产编号</b>，指标中配置 <b>Dx1、Dy1、Dx2、Dy2、Rz</b>，系统会读取当前生产编号对应行。</p>
      <p class="note">V2.6 supports Image OCR data sources for .png/.jpg/.jpeg/.bmp/.tif/.tiff. Configure ROI and regex in the measurement item to extract Rx/Ry/Z from fixed-layout equipment images; shared image folders can route rows by production code parsed from filename or OCR.</p>
    </div>
    <div class="card">
      <h2>共享路径注意事项</h2>
      <p class="note">例如：<code>\\\\192.168.1.100\\share\\result.csv</code>。运行本程序的电脑必须能访问该路径，并且 Windows 当前用户要有共享目录读取权限。若部署在 Linux/Docker，建议把共享目录挂载为本地路径，例如 <code>/mnt/metrology/result.csv</code>。</p>
    </div>
    """, user)


# ==========================================================
# HTTP handler
# ==========================================================

class AppHandler(BaseHTTPRequestHandler):
    def send_bytes(self, status, headers, data):
        self.send_response(status)
        for k, v in headers.items():
            self.send_header(k, v)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_html(self, text, status=200, headers=None):
        data = text.encode("utf-8")
        hds = {"Content-Type": "text/html; charset=utf-8"}
        if headers:
            hds.update(headers)
        self.send_bytes(status, hds, data)

    def require_user(self):
        user = current_user(self)
        if not user:
            status, headers, data = redirect("/login")
            self.send_bytes(status, headers, data)
            return None
        return user

    def ensure_permission(self, user, allowed, object_type="route", object_id=""):
        if allowed:
            return True
        write_audit(
            (user or {}).get("username", "unknown"),
            "PERMISSION_DENIED",
            object_type,
            object_id or self.path,
            "权限不足，拒绝访问",
            self.client_address[0]
        )
        self.send_html(base_layout("权限不足", "<h1>权限不足</h1><div class='card error'>当前账号没有权限执行该操作。</div>", user), status=403)
        return False

    def parse_post_data(self):
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            raw = body.decode("utf-8", errors="replace")
            return parse_qs(raw), {}

        boundary = None
        for part in content_type.split(";"):
            part = part.strip()
            if part.startswith("boundary="):
                boundary = part.split("=", 1)[1].strip().strip('"')
                break
        if not boundary:
            raise ValueError("multipart/form-data 缺少 boundary")

        delimiter = ("--" + boundary).encode("utf-8")
        form = {}
        files = {}
        for part in body.split(delimiter):
            part = part.strip()
            if not part or part == b"--":
                continue
            if part.endswith(b"--"):
                part = part[:-2].strip()
            if b"\r\n\r\n" not in part:
                continue
            header_bytes, data = part.split(b"\r\n\r\n", 1)
            data = data.rstrip(b"\r\n")
            header_text = header_bytes.decode("utf-8", errors="replace")
            name = None
            filename = None
            part_content_type = ""
            for line in header_text.split("\r\n"):
                lower = line.lower()
                if lower.startswith("content-disposition:"):
                    chunks = line.split(";")
                    for chunk in chunks:
                        chunk = chunk.strip()
                        if chunk.startswith("name="):
                            name = chunk.split("=", 1)[1].strip().strip('"')
                        elif chunk.startswith("filename="):
                            filename = chunk.split("=", 1)[1].strip().strip('"')
                elif lower.startswith("content-type:"):
                    part_content_type = line.split(":", 1)[1].strip()
            if not name:
                continue
            if filename:
                files[name] = {"filename": filename, "content": data, "content_type": part_content_type}
            else:
                form.setdefault(name, []).append(data.decode("utf-8", errors="replace"))
        return form, files

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        q = parse_qs(parsed.query)

        if path == "/version":
            self.send_html(f"<h1>Metrology Config App {h(APP_VERSION)}</h1><p>PORT={h(PORT)}</p><p>CSV/XLSX/Image OCR collection + collect_job lifecycle tracking + roles + template snapshots + config checks + dashboard stability cards</p>")
            return

        if path == "/login":
            self.send_html(page_login())
            return
        if path == "/logout":
            sid = parse_cookie(self.headers.get("Cookie")).get("sid")
            if sid in SESSIONS:
                write_audit(SESSIONS[sid].get("username"), "LOGOUT", "session", sid[:8], "管理员退出登录", self.client_address[0])
                del SESSIONS[sid]
            status, headers, data = redirect("/login")
            headers["Set-Cookie"] = "sid=; Path=/; Max-Age=0; HttpOnly; SameSite=Lax"
            self.send_bytes(status, headers, data)
            return

        user = self.require_user()
        if not user:
            return

        if path == "/":
            self.send_html(page_dashboard(user, q))
        elif path == "/productions":
            self.send_html(page_productions(user))
        elif path == "/production_new":
            if not self.ensure_permission(user, can_manage_config(user), "production_config", "new"):
                return
            self.send_html(page_production_form(user))
        elif path == "/production_edit":
            if not self.ensure_permission(user, can_manage_config(user), "production_config", q.get("id", [""])[0]):
                return
            self.send_html(page_production_form(user, safe_int(q.get("id", [0])[0])))
        elif path == "/items":
            self.send_html(page_items(user, safe_int(q.get("production_id", [0])[0])))
        elif path == "/item_new":
            if not self.ensure_permission(user, can_manage_config(user), "measurement_item_config", "new"):
                return
            self.send_html(page_item_form(user, production_id=safe_int(q.get("production_id", [0])[0])))
        elif path == "/item_edit":
            if not self.ensure_permission(user, can_manage_config(user), "measurement_item_config", q.get("id", [""])[0]):
                return
            self.send_html(page_item_form(user, item_id=safe_int(q.get("id", [0])[0])))
        elif path == "/metrics":
            self.send_html(page_metrics(user, safe_int(q.get("item_id", [0])[0])))
        elif path == "/metric_new":
            if not self.ensure_permission(user, can_manage_config(user), "metric_config", "new"):
                return
            self.send_html(page_metric_form(user, item_id=safe_int(q.get("item_id", [0])[0])))
        elif path == "/metric_edit":
            if not self.ensure_permission(user, can_manage_config(user), "metric_config", q.get("id", [""])[0]):
                return
            self.send_html(page_metric_form(user, metric_id=safe_int(q.get("id", [0])[0])))
        elif path == "/test_collect":
            if not self.ensure_permission(user, can_collect(user), "collect_job", "test"):
                return
            self.send_html(page_test_collect(user, safe_int(q.get("item_id", [0])[0])))
        elif path == "/collect_now":
            if not self.ensure_permission(user, can_collect(user), "collect_job", "manual"):
                return
            self.send_html(page_collect_now(user, safe_int(q.get("item_id", [0])[0])))
        elif path == "/templates":
            self.send_html(page_templates(user))
        elif path == "/template_upload":
            if not self.ensure_permission(user, can_manage_config(user), "template_config", "upload"):
                return
            self.send_html(page_template_upload(user))
        elif path == "/template_detail":
            self.send_html(page_template_detail(user, safe_int(q.get("id", [0])[0])))
        elif path == "/template_edit":
            if not self.ensure_permission(user, can_manage_config(user), "template_config", q.get("id", [""])[0]):
                return
            self.send_html(page_template_edit(user, safe_int(q.get("id", [0])[0])))
        elif path == "/template_apply":
            if not self.ensure_permission(user, can_manage_config(user), "template_config", "apply"):
                return
            self.send_html(page_template_apply_bulk(user, safe_int(q.get("production_id", [0])[0])))
        elif path == "/results":
            self.send_html(page_results(user, q))
        elif path == "/collect_jobs":
            self.send_html(page_collect_jobs(user))
        elif path == "/logs":
            self.send_html(page_logs(user))
        elif path == "/config_checks":
            self.send_html(page_config_checks(user))
        elif path == "/audit_logs":
            if not self.ensure_permission(user, can_manage_system(user), "audit_log", "view"):
                return
            self.send_html(page_audit_logs(user))
        elif path == "/export_config":
            production_id = safe_int(q.get("production_id", [0])[0])
            cfg = export_config_json(production_id)
            if not cfg:
                self.send_html(base_layout("错误", "<h1>配置不存在</h1>", user), status=404)
                return
            data = json.dumps(cfg, ensure_ascii=False, indent=2).encode("utf-8")
            filename = f"metrology_config_{cfg['production']['production_code']}.json"
            write_audit(user.get("username"), "EXPORT_CONFIG", "production_config", production_id, f"导出配置 {filename}", self.client_address[0])
            self.send_bytes(200, {
                "Content-Type": "application/json; charset=utf-8",
                "Content-Disposition": f"attachment; filename={filename}"
            }, data)
        elif path == "/export_all_config":
            cfg = export_all_config_json()
            data = json.dumps(cfg, ensure_ascii=False, indent=2).encode("utf-8")
            filename = f"metrology_all_configs_{datetime.now(APP_TZ).strftime('%Y%m%d_%H%M%S')}.json"
            write_audit(user.get("username"), "EXPORT_ALL_CONFIG", "production_config", "all", f"导出全部配置 {filename}", self.client_address[0])
            self.send_bytes(200, {
                "Content-Type": "application/json; charset=utf-8",
                "Content-Disposition": f"attachment; filename={filename}"
            }, data)
        elif path == "/export_results_xlsx":
            data = export_results_xlsx(q)
            filename = f"measurement_results_{datetime.now(APP_TZ).strftime('%Y%m%d_%H%M%S')}.xlsx"
            write_audit(user.get("username"), "EXPORT_RESULTS_XLSX", "measurement_result", "", "导出采集结果 Excel", self.client_address[0])
            self.send_bytes(200, {
                "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Content-Disposition": f"attachment; filename={filename}"
            }, data)
        elif path == "/import_config":
            if not self.ensure_permission(user, can_manage_config(user), "production_config", "import"):
                return
            self.send_html(page_import_config(user))
        elif path == "/users":
            if not self.ensure_permission(user, can_manage_users(user), "admin_user", "view"):
                return
            self.send_html(page_users(user))
        elif path == "/about":
            self.send_html(page_about(user))
        else:
            self.send_html(base_layout("404", "<h1>404 Not Found</h1>", user), status=404)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path
        form, files = self.parse_post_data()

        if path == "/login":
            username = form.get("username", [""])[0].strip()
            password = form.get("password", [""])[0]
            conn = get_conn()
            user = conn.execute("SELECT * FROM admin_user WHERE username=?", (username,)).fetchone()
            conn.close()
            if user and user["password_hash"] == hash_password(password) and user["enabled"]:
                write_audit(username, "LOGIN_SUCCESS", "admin_user", username, "登录成功", self.client_address[0])
                sid = secrets.token_urlsafe(32)
                SESSIONS[sid] = {"username": username, "role": user["role"] or "viewer", "login_time": now_str()}
                status, headers, data = redirect("/")
                headers["Set-Cookie"] = f"sid={sid}; Path=/; HttpOnly; SameSite=Lax"
                self.send_bytes(status, headers, data)
            else:
                detail = "登录失败：账号禁用或密码错误"
                write_audit(username or "unknown", "LOGIN_FAILED", "admin_user", username, detail, self.client_address[0])
                self.send_html(page_login("账号或密码错误"), status=401)
            return

        user = self.require_user()
        if not user:
            return

        try:
            if path == "/production_save":
                self.handle_production_save(form)
            elif path == "/production_delete":
                self.handle_production_delete(form)
            elif path == "/item_save":
                self.handle_item_save(form)
            elif path == "/item_delete":
                self.handle_item_delete(form)
            elif path == "/metric_bulk_add":
                self.handle_metric_bulk_add(form)
            elif path == "/metric_save":
                self.handle_metric_save(form)
            elif path == "/metric_delete":
                self.handle_metric_delete(form)
            elif path == "/template_parse":
                require_permission(user, can_manage_config(user))
                self.send_html(page_template_mapping(user, form, files))
            elif path == "/template_save":
                tid = handle_template_save(user, form, self.client_address[0])
                status, headers, data = redirect(f"/template_detail?id={tid}")
                self.send_bytes(status, headers, data)
            elif path == "/template_update":
                tid = handle_template_update(user, form, self.client_address[0])
                status, headers, data = redirect(f"/template_detail?id={tid}")
                self.send_bytes(status, headers, data)
            elif path == "/template_delete":
                require_permission(user, can_manage_config(user))
                template_id = safe_int(form.get("template_id", [0])[0])
                template_name = delete_template_config(template_id)
                write_audit(user.get("username"), "DELETE_TEMPLATE", "template_config", template_id, f"删除模板 {template_name}", self.client_address[0])
                status, headers, data = redirect("/templates")
                self.send_bytes(status, headers, data)
            elif path == "/template_apply":
                applied_items = handle_template_apply_bulk(user, form, self.client_address[0])
                if len(applied_items) == 1:
                    status, headers, data = redirect(f"/metrics?item_id={applied_items[0]['item_id']}")
                else:
                    production_id = safe_int(form.get("production_id", [0])[0])
                    status, headers, data = redirect(f"/items?production_id={production_id}")
                self.send_bytes(status, headers, data)
            elif path == "/demo_seed":
                result = seed_demo_data(user, self.client_address[0])
                self.send_html(page_demo_ready(user, result))
            elif path == "/import_config":
                require_permission(user, can_manage_config(user))
                config_json = form.get("config_json", [""])[0]
                result = import_config(config_json)
                write_audit(user.get("username"), "IMPORT_CONFIG", "production_config", ",".join(str(x) for x in result["ids"]), f"导入生产编号配置 JSON，共 {result['count']} 个", self.client_address[0])
                self.send_html(page_import_config(user, f"导入成功，共 {result['count']} 个生产编号，ID={result['ids']}"))
            elif path == "/logs_clear":
                require_permission(user, can_manage_system(user))
                write_audit(user.get("username"), "LOG_RETENTION_PROTECTED", "collect_log", "", "产线试运行阶段拒绝清空采集日志", self.client_address[0])
                self.send_html(base_layout("日志保留保护", "<h1>日志保留保护</h1><div class='card error'>产线试运行阶段禁止清空 collect_log，历史采集日志必须保留。</div><p><a class='btn secondary' href='/logs'>返回采集日志</a></p>", user), status=409)
            elif path == "/logs_prune":
                require_permission(user, can_manage_system(user))
                write_audit(user.get("username"), "LOG_RETENTION_PROTECTED", "collect_log", "", "产线试运行阶段拒绝裁剪采集日志", self.client_address[0])
                self.send_html(base_layout("日志保留保护", "<h1>日志保留保护</h1><div class='card error'>产线试运行阶段禁止裁剪 collect_log，历史采集日志必须保留。</div><p><a class='btn secondary' href='/logs'>返回采集日志</a></p>", user), status=409)
            elif path == "/orphan_results_clear":
                require_permission(user, can_void_results(user))
                removed = clear_orphan_measurement_results()
                write_audit(user.get("username"), "VOID_ORPHAN_RESULTS", "measurement_result", "", f"作废不属于当前有效生产编号/量测项配置的历史采集结果 {removed} 条", self.client_address[0])
                status, headers, data = redirect("/logs")
                self.send_bytes(status, headers, data)
            elif path == "/result_delete":
                require_permission(user, can_void_results(user))
                result_id = safe_int(form.get("result_id", [0])[0])
                reason = form.get("void_reason", ["人工作废"])[0].strip() or "人工作废"
                production_code, process_step, metric_name = delete_result(result_id, user.get("username"), reason)
                write_audit(user.get("username"), "VOID_RESULT", "measurement_result", result_id, f"作废采集结果 {production_code}/{process_step}/{metric_name}，原因：{reason}", self.client_address[0])
                status, headers, data = redirect("/results")
                self.send_bytes(status, headers, data)
            elif path == "/results_bulk_delete":
                require_permission(user, can_void_results(user))
                result_ids = form.get("result_ids", [])
                reason = form.get("void_reason", ["人工作废"])[0].strip() or "人工作废"
                removed = bulk_delete_results(result_ids, user.get("username"), reason)
                write_audit(user.get("username"), "BULK_VOID_RESULTS", "measurement_result", ",".join(result_ids), f"批量作废采集结果 {removed} 条，原因：{reason}", self.client_address[0])
                status, headers, data = redirect("/results")
                self.send_bytes(status, headers, data)
            elif path == "/blank_process_results_clear":
                require_permission(user, can_void_results(user))
                removed = clear_blank_process_results()
                write_audit(user.get("username"), "VOID_BLANK_PROCESS_RESULTS", "measurement_result", "", f"作废空工序采集结果 {removed} 条", self.client_address[0])
                status, headers, data = redirect("/results")
                self.send_bytes(status, headers, data)
            elif path == "/user_create":
                handle_user_create(user, form, self.client_address[0])
                status, headers, data = redirect("/users")
                self.send_bytes(status, headers, data)
            elif path == "/user_role_save":
                handle_user_role_save(user, form, self.client_address[0])
                status, headers, data = redirect("/users")
                self.send_bytes(status, headers, data)
            elif path == "/user_reset_password":
                handle_user_reset_password(user, form, self.client_address[0])
                status, headers, data = redirect("/users")
                self.send_bytes(status, headers, data)
            elif path == "/user_toggle_enabled":
                handle_user_toggle_enabled(user, form, self.client_address[0])
                status, headers, data = redirect("/users")
                self.send_bytes(status, headers, data)
            else:
                self.send_html(base_layout("404", "<h1>404 Not Found</h1>", user), status=404)
        except PermissionError as ex:
            write_audit(user.get("username"), "PERMISSION_DENIED", "route", path, str(ex), self.client_address[0])
            self.send_html(base_layout("权限不足", f"<h1>权限不足</h1><div class='card error'>{h(ex)}</div>", user), status=403)
        except Exception as ex:
            request_id = log_request_exception(user, self.command, path, ex, self.client_address[0])
            self.send_html(page_request_error(user, request_id), status=500)

    def handle_production_save(self, form):
        require_permission(current_user(self), can_manage_config(current_user(self)))
        pid = form.get("id", [""])[0].strip()
        vals = (
            form.get("production_code", [""])[0].strip(),
            form.get("production_name", [""])[0].strip(),
            form.get("product_model", [""])[0].strip(),
            form.get("process_version", [""])[0].strip(),
            form.get("description", [""])[0].strip(),
            form.get("status", ["enabled"])[0],
            now_str()
        )
        conn = get_conn()
        cur = conn.cursor()
        if pid:
            cur.execute("""
            UPDATE production_config SET production_code=?, production_name=?, product_model=?, process_version=?, description=?, status=?, updated_at=? WHERE id=?
            """, vals + (safe_int(pid),))
            new_id = safe_int(pid)
        else:
            cur.execute("""
            INSERT INTO production_config (production_code, production_name, product_model, process_version, description, status, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """, vals)
            new_id = cur.lastrowid
        conn.commit()
        conn.close()
        write_audit(current_user(self).get("username"), "SAVE_PRODUCTION", "production_config", new_id, f"保存生产编号 {vals[0]}", self.client_address[0])
        status, headers, data = redirect(f"/items?production_id={new_id}")
        self.send_bytes(status, headers, data)

    def handle_production_delete(self, form):
        require_permission(current_user(self), can_manage_config(current_user(self)))
        production_id = safe_int(form.get("production_id", [0])[0])
        production_code = delete_production_config(production_id)
        write_audit(current_user(self).get("username"), "DELETE_PRODUCTION", "production_config", production_id, f"删除生产编号配置 {production_code}；历史采集结果保留", self.client_address[0])
        status, headers, data = redirect("/productions")
        self.send_bytes(status, headers, data)

    def handle_item_save(self, form):
        require_permission(current_user(self), can_manage_config(current_user(self)))
        item_id = form.get("id", [""])[0].strip()
        production_id = safe_int(form.get("production_id", [0])[0])
        process_step_value = form.get("process_step", [""])[0].strip()
        process_step_column_value = form.get("process_step_column", [""])[0].strip()
        data_source_type_value = form.get("data_source_type", ["auto"])[0]
        image_parse_config_value = form.get("image_parse_config_json", [""])[0].strip()
        if data_source_type_value == "image":
            image_config_value = parse_image_parse_config(image_parse_config_value)
            process_step_column_value = image_config_process_column(image_config_value, process_step_column_value)
        if not process_step_value and not process_step_column_value:
            raise ValueError("量测项必须填写固定量测工序，或填写工序字段名；否则系统无法追溯数据属于哪道工序，且会禁止采集。图片 OCR 可通过 process_from_filename_regex 或 fields.process_step 自动提供工序。")
        vals = (
            production_id,
            form.get("item_name", [""])[0].strip(),
            process_step_value,
            process_step_column_value,
            form.get("execution_time_text", [""])[0].strip(),
            form.get("equipment_name", [""])[0].strip(),
            data_source_type_value,
            form.get("data_source_path", [""])[0].strip(),
            form.get("excel_sheet_name", [""])[0].strip(),
            image_parse_config_value,
            max(1, safe_int(form.get("header_row_index", [1])[0], 1)),
            form.get("csv_encoding", ["auto"])[0],
            form.get("delimiter", [","])[0],
            form.get("production_code_column", ["生产编号"])[0].strip(),
            max(10, safe_int(form.get("scan_frequency_seconds", [60])[0], 60)),
            safe_int(form.get("enabled", [1])[0], 1),
            now_str()
        )
        conn = get_conn()
        cur = conn.cursor()
        if item_id:
            cur.execute("""
            UPDATE measurement_item_config SET production_id=?, item_name=?, process_step=?, process_step_column=?, execution_time_text=?, equipment_name=?, data_source_type=?, data_source_path=?, excel_sheet_name=?, image_parse_config_json=?, header_row_index=?, csv_encoding=?, delimiter=?, production_code_column=?, scan_frequency_seconds=?, enabled=?, updated_at=? WHERE id=?
            """, vals + (safe_int(item_id),))
            new_id = safe_int(item_id)
        else:
            cur.execute("""
            INSERT INTO measurement_item_config (production_id, item_name, process_step, process_step_column, execution_time_text, equipment_name, data_source_type, data_source_path, excel_sheet_name, image_parse_config_json, header_row_index, csv_encoding, delimiter, production_code_column, scan_frequency_seconds, enabled, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, vals)
            new_id = cur.lastrowid
        conn.commit()
        conn.close()
        write_audit(current_user(self).get("username"), "SAVE_MEASUREMENT_ITEM", "measurement_item_config", new_id, f"保存量测项 {vals[1]}", self.client_address[0])
        status, headers, data = redirect(f"/metrics?item_id={new_id}")
        self.send_bytes(status, headers, data)

    def handle_item_delete(self, form):
        require_permission(current_user(self), can_manage_config(current_user(self)))
        item_id = safe_int(form.get("item_id", [0])[0])
        production_id, item_name = delete_item_config(item_id)
        write_audit(current_user(self).get("username"), "DELETE_MEASUREMENT_ITEM", "measurement_item_config", item_id, f"删除量测项配置 {item_name}；历史采集结果保留", self.client_address[0])
        status, headers, data = redirect(f"/items?production_id={production_id}")
        self.send_bytes(status, headers, data)

    def handle_metric_bulk_add(self, form):
        require_permission(current_user(self), can_manage_config(current_user(self)))
        item_id = safe_int(form.get("item_id", [0])[0])
        names = form.get("metric_names", [""])[0]
        unit = form.get("unit", [""])[0].strip()
        metric_names = [x.strip() for x in names.replace("，", ",").split(",") if x.strip()]
        conn = get_conn()
        cur = conn.cursor()
        for idx, name in enumerate(metric_names):
            cur.execute("""
            INSERT INTO metric_config (item_id, metric_name, source_column, unit, data_type, enabled, sort_order, updated_at)
            VALUES (?, ?, ?, ?, 'number', 1, ?, ?)
            """, (item_id, name, name, unit, idx, now_str()))
        conn.commit()
        conn.close()
        write_audit(current_user(self).get("username"), "BULK_ADD_METRICS", "metric_config", item_id, f"批量添加指标：{', '.join(metric_names)}", self.client_address[0])
        status, headers, data = redirect(f"/metrics?item_id={item_id}")
        self.send_bytes(status, headers, data)

    def handle_metric_save(self, form):
        require_permission(current_user(self), can_manage_config(current_user(self)))
        metric_id = form.get("id", [""])[0].strip()
        item_id = safe_int(form.get("item_id", [0])[0])
        vals = (
            item_id,
            form.get("metric_name", [""])[0].strip(),
            form.get("source_column", [""])[0].strip(),
            form.get("unit", [""])[0].strip(),
            form.get("data_type", ["number"])[0],
            safe_float(form.get("target", [""])[0]),
            safe_float(form.get("lsl", [""])[0]),
            safe_float(form.get("usl", [""])[0]),
            safe_float(form.get("lcl", [""])[0]),
            safe_float(form.get("ucl", [""])[0]),
            safe_int(form.get("enabled", [1])[0], 1),
            safe_int(form.get("sort_order", [0])[0], 0),
            now_str()
        )
        conn = get_conn()
        cur = conn.cursor()
        if metric_id:
            cur.execute("""
            UPDATE metric_config SET item_id=?, metric_name=?, source_column=?, unit=?, data_type=?, target=?, lsl=?, usl=?, lcl=?, ucl=?, enabled=?, sort_order=?, updated_at=? WHERE id=?
            """, vals + (safe_int(metric_id),))
        else:
            cur.execute("""
            INSERT INTO metric_config (item_id, metric_name, source_column, unit, data_type, target, lsl, usl, lcl, ucl, enabled, sort_order, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, vals)
        conn.commit()
        conn.close()
        write_audit(current_user(self).get("username"), "SAVE_METRIC", "metric_config", metric_id or "new", f"保存指标 {vals[1]}", self.client_address[0])
        status, headers, data = redirect(f"/metrics?item_id={item_id}")
        self.send_bytes(status, headers, data)

    def handle_metric_delete(self, form):
        require_permission(current_user(self), can_manage_config(current_user(self)))
        metric_id = safe_int(form.get("metric_id", [0])[0])
        item_id, metric_name = delete_metric_config(metric_id)
        write_audit(current_user(self).get("username"), "DELETE_METRIC", "metric_config", metric_id, f"删除指标配置 {metric_name}；历史采集结果保留", self.client_address[0])
        status, headers, data = redirect(f"/metrics?item_id={item_id}")
        self.send_bytes(status, headers, data)

    def log_message(self, fmt, *args):
        print("[%s] %s" % (now_str(), fmt % args))


# ==========================================================
# Main
# ==========================================================

def main():
    init_db()
    t = threading.Thread(target=scheduler_loop, daemon=True)
    t.start()
    server = ThreadingHTTPServer((HOST, PORT), AppHandler)
    print("=" * 78)
    print(APP_TITLE)
    print(f"本机监听：http://{HOST}:{PORT}")
    if HOST == "0.0.0.0":
        print(f"局域网访问：http://{DISPLAY_IP}:{PORT}")
        if not os.environ.get("MDCP_ADMIN_PASSWORD"):
            print("安全警告：当前为局域网监听模式且未设置 MDCP_ADMIN_PASSWORD，请立即设置强密码。")
    else:
        print("当前仅本机可访问。如需局域网访问，请显式设置 MDCP_HOST=0.0.0.0 并设置强密码。")
    print(f"默认账号：{os.environ.get('MDCP_ADMIN_USERNAME', 'admin')}")
    print("默认密码：来自 MDCP_ADMIN_PASSWORD 环境变量；未设置时为 admin123（仅建议本地测试）")
    print(r"数据源路径示例：\\192.168.1.100\share\result.xlsx")
    print("按 Ctrl+C 停止")
    print("=" * 78)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n正在停止服务...")
    finally:
        SCHEDULER_STOP.set()
        server.server_close()


if __name__ == "__main__":
    main()
