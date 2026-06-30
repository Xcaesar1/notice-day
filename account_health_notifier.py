from __future__ import annotations

import argparse
import base64
import hashlib
import hmac
import json
import os
import re
import shutil
import sqlite3
import socket
import subprocess
import sys
import tempfile
import time
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable
from urllib import error as urlerror
from urllib import parse as urlparse
from urllib import request as urlrequest
from zipfile import ZIP_DEFLATED, ZipFile
import html
import xml.etree.ElementTree as ET

import cdp_account_health
import ziniao_webdriver
import ziniao_cdp
import zclaw_account_health


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_STATE_DIR = PROJECT_ROOT / ".local-state" / "account-health-notifier"
DEFAULT_CONFIG_PATH = DEFAULT_STATE_DIR / "config.json"
DEFAULT_DB_PATH = DEFAULT_STATE_DIR / "state.sqlite"
DEFAULT_DWS_CALL = Path(r"Q:\Dingcli\dws-call.cmd")
DEFAULT_SOURCE_DIR = Path(r"C:\Users\god\Desktop\RPA下载结果\账户状况异常明细")
DEFAULT_STORE_LIST = Path(r"F:\店铺清单.xlsx")
DEFAULT_TASK_NAME = "YD-AmazonAccountHealth-DingTalkNotifier"
DEFAULT_CDP_DAEMON_TASK_NAME = "YD-ZiniaoCdpDaemon"

SITE_US = "美国"
SHEET_DETAIL_NAMES = {"异常明细", "寮傚父鏄庣粏"}

DETAIL_HEADERS = [
    "店铺",
    "站点",
    "异常分类",
    "原因",
    "日期",
    "哪些商品会受到影响？",
    "存在销售风险",
    "采取的操作",
    "账户状况评级影响",
]

ITEM_HEADERS = [
    "run_id",
    "店铺",
    "站点",
    "异常分类",
    "ASIN",
    "SKU",
    "原因",
    "日期",
    "存在销售风险",
    "采取的操作",
    "账户状况评级影响",
    "dedupe_key",
    "content_hash",
    "notify_status",
]

RUN_HEADERS = [
    "run_id",
    "started_at",
    "ended_at",
    "source",
    "total_items",
    "notify_candidates",
    "sent_items",
    "status",
    "error",
]

PARSE_SUMMARY_HEADERS = [
    "维度",
    "名称",
    "数量",
]

CDP_TARGET_HEADERS = [
    "run_id",
    "store",
    "site",
    "status",
    "row_count",
    "error",
    "target_port",
    "target_title",
    "target_url",
    "started_at",
    "ended_at",
]

HEADER_ALIASES = {
    "店铺": {"店铺", "搴楅摵"},
    "站点": {"站点", "绔欑偣"},
    "异常分类": {"异常分类", "寮傚父鍒嗙被"},
    "原因": {"原因", "鍘熷洜"},
    "日期": {"日期", "鏃ユ湡"},
    "哪些商品会受到影响？": {
        "哪些商品会受到影响？",
        "哪些商品会受到影响?",
        "鍝簺鍟嗗搧浼氬彈鍒板奖鍝嶏紵",
    },
    "存在销售风险": {"存在销售风险", "瀛樺湪閿€鍞闄?"},
    "采取的操作": {"采取的操作", "閲囧彇鐨勬搷浣?"},
    "账户状况评级影响": {"账户状况评级影响", "璐︽埛鐘跺喌璇勭骇褰卞搷"},
}

_NS = {
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


@dataclass(frozen=True)
class ImpactItem:
    store: str
    site: str
    category: str
    asin: str
    sku: str
    reason: str
    date: str
    impacted_text: str
    sales_risk: str
    action: str
    rating_impact: str
    source_file: str = ""

    @property
    def dedupe_key(self) -> str:
        parts = [
            self.store,
            self.site,
            self.category,
            self.asin,
            self.sku,
            self.date,
            self.reason,
        ]
        return stable_hash("|".join(normalize_for_key(part) for part in parts))

    @property
    def content_hash(self) -> str:
        parts = [
            self.store,
            self.site,
            self.category,
            self.asin,
            self.sku,
            self.reason,
            self.date,
            self.impacted_text,
            self.sales_risk,
            self.action,
            self.rating_impact,
        ]
        return stable_hash("|".join(clean_text(part) for part in parts))

    def to_row(self, run_id: str, notify_status: str) -> dict[str, str]:
        return {
            "run_id": run_id,
            "店铺": self.store,
            "站点": self.site,
            "异常分类": self.category,
            "ASIN": self.asin,
            "SKU": self.sku,
            "原因": self.reason,
            "日期": self.date,
            "存在销售风险": self.sales_risk,
            "采取的操作": self.action,
            "账户状况评级影响": self.rating_impact,
            "dedupe_key": self.dedupe_key,
            "content_hash": self.content_hash,
            "notify_status": notify_status,
        }


class SourceStaleError(Exception):
    def __init__(self, path: Path, message: str):
        super().__init__(message)
        self.path = path


class RunLockError(Exception):
    def __init__(self, path: Path, message: str):
        super().__init__(message)
        self.path = path


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\x00", "")
    text = re.sub(r"[\x01-\x08\x0b\x0c\x0e-\x1f]", "", text)
    return re.sub(r"[ \t\r\f\v]+", " ", text).strip()


def normalize_for_key(value: Any) -> str:
    return re.sub(r"\s+", "", clean_text(value).lower())


def stable_hash(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def run_id_text() -> str:
    return f"{datetime.now().strftime('%Y%m%d_%H%M%S_%f')}_{uuid.uuid4().hex[:8]}"


def default_config() -> dict[str, Any]:
    return {
        "collector": {
            "backend": "zclaw",
        },
        "source": {
            "type": "excel_latest",
            "excel_path": "",
            "excel_dir": str(DEFAULT_SOURCE_DIR),
            "store_list_path": str(DEFAULT_STORE_LIST),
            "site": SITE_US,
            "max_excel_age_hours": 30,
        },
        "ziniao_webdriver": {
            "client_path": "",
            "port": 9515,
            "company": "",
            "username": "",
            "password": "",
            "password_env": "ZINIAO_WEBDRIVER_PASSWORD",
            "request_timeout_seconds": 30,
            "startup_timeout_seconds": 60,
            "update_core_timeout_seconds": 300,
            "update_core_poll_seconds": 2,
            "browser_ready_timeout_seconds": 60,
            "navigation_wait_seconds": 3,
        },
        "ziniao_cdp": {
            "port": 0,
            "port_start": 9222,
            "port_end": 9250,
            "url_contains": "sellercentral.amazon.com",
            "text_limit": 800,
            "collect_page_size": 25,
            "collect_max_pages": 20,
            "watch_duration_seconds": 30,
            "watch_interval_seconds": 2,
            "reconnect_timeout_seconds": 60,
            "daemon_task_name": DEFAULT_CDP_DAEMON_TASK_NAME,
            "daemon_log_path": str(DEFAULT_STATE_DIR / "ziniao-cdp-daemon.log"),
        },
        "dingtalk": {
            "method": "webhook",
            "webhook_url": "",
            "secret": "",
            "dws_call": str(DEFAULT_DWS_CALL),
            "robot_code": "",
            "group_open_conversation_id": "",
            "title_prefix": "亚马逊账号状况异常",
            "send_enabled": False,
        },
        "notify": {
            "dedupe_retention_days": 90,
            "max_items_per_message": 60,
            "require_complete_product_ids_before_send": True,
            "require_all_stores_before_send": True,
        },
        "production": {
            "retry_failed_stores_enabled": True,
            "retry_delay_seconds": 600,
            "send_partial_with_failed_stores": True,
        },
        "schedule": {
            "task_name": DEFAULT_TASK_NAME,
            "interval_hours": 6,
            "command": "production-run",
        },
        "runtime": {
            "primary_host": "",
            "enforce_primary_host_for_send": True,
        },
        "state": {
            "db_path": str(DEFAULT_DB_PATH),
            "result_dir": str(DEFAULT_STATE_DIR / "runs"),
            "run_lock_ttl_minutes": 240,
        },
    }


def load_config(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(f"配置文件不存在: {path}. 先运行 init-config。")
    with path.open("r", encoding="utf-8") as fh:
        loaded = json.load(fh)
    merged = default_config()
    deep_update(merged, loaded)
    return merged


def deep_update(target: dict[str, Any], source: dict[str, Any]) -> None:
    for key, value in source.items():
        if isinstance(value, dict) and isinstance(target.get(key), dict):
            deep_update(target[key], value)
        else:
            target[key] = value


def write_default_config(path: Path, force: bool = False) -> Path:
    if path.exists() and not force:
        raise FileExistsError(f"配置文件已存在: {path}. 如需覆盖, 加 --force。")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(default_config(), fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    return path


def connect_db(path: Path) -> sqlite3.Connection:
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.row_factory = sqlite3.Row
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS notified_items (
            dedupe_key TEXT PRIMARY KEY,
            content_hash TEXT NOT NULL,
            store TEXT NOT NULL,
            site TEXT NOT NULL,
            category TEXT NOT NULL,
            asin TEXT NOT NULL,
            sku TEXT NOT NULL,
            reason TEXT NOT NULL,
            issue_date TEXT NOT NULL,
            first_seen_at TEXT NOT NULL,
            last_seen_at TEXT NOT NULL,
            notified_at TEXT NOT NULL,
            notify_count INTEGER NOT NULL DEFAULT 1,
            payload_json TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS runs (
            run_id TEXT PRIMARY KEY,
            started_at TEXT NOT NULL,
            ended_at TEXT,
            source TEXT NOT NULL,
            total_items INTEGER NOT NULL DEFAULT 0,
            notify_candidates INTEGER NOT NULL DEFAULT 0,
            sent_items INTEGER NOT NULL DEFAULT 0,
            status TEXT NOT NULL,
            error TEXT NOT NULL DEFAULT ""
        );
        CREATE TABLE IF NOT EXISTS notification_attempts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            run_id TEXT NOT NULL,
            attempted_at TEXT NOT NULL,
            dry_run INTEGER NOT NULL,
            status TEXT NOT NULL,
            item_count INTEGER NOT NULL,
            title TEXT NOT NULL,
            error TEXT NOT NULL DEFAULT "",
            dws_stdout TEXT NOT NULL DEFAULT "",
            dws_stderr TEXT NOT NULL DEFAULT ""
        );
        """
    )
    conn.commit()
    return conn


NOTIFIED_ITEM_COLUMNS = [
    "dedupe_key",
    "content_hash",
    "store",
    "site",
    "category",
    "asin",
    "sku",
    "reason",
    "issue_date",
    "first_seen_at",
    "last_seen_at",
    "notified_at",
    "notify_count",
    "payload_json",
]


def table_exists(conn: sqlite3.Connection, table_name: str) -> bool:
    row = conn.execute(
        "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
        (table_name,),
    ).fetchone()
    return row is not None


def insert_notified_item_row(conn: sqlite3.Connection, row: sqlite3.Row | dict[str, Any]) -> None:
    values = [row[column] for column in NOTIFIED_ITEM_COLUMNS]
    conn.execute(
        """
        INSERT INTO notified_items (
            dedupe_key, content_hash, store, site, category, asin, sku, reason,
            issue_date, first_seen_at, last_seen_at, notified_at, notify_count, payload_json
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        values,
    )


def configured_run_lock_ttl_minutes(config: dict[str, Any]) -> float:
    raw = config.get("state", {}).get("run_lock_ttl_minutes", 240)
    if raw in (None, ""):
        return 240.0
    try:
        value = float(raw)
    except Exception:
        raise ValueError("state.run_lock_ttl_minutes 必须是数字")
    if value <= 0:
        raise ValueError("state.run_lock_ttl_minutes 必须大于 0")
    return value


def acquire_run_lock(state_dir: Path, run_id: str, ttl_minutes: float) -> Path:
    state_dir.mkdir(parents=True, exist_ok=True)
    lock_path = state_dir / "run.lock"
    if lock_path.exists():
        age_minutes = max(0.0, (time.time() - lock_path.stat().st_mtime) / 60)
        if age_minutes > ttl_minutes:
            lock_path.unlink()
        else:
            raise RunLockError(lock_path, f"已有运行锁未释放: {lock_path}; 文件年龄 {age_minutes:.1f} 分钟。")
    try:
        fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        raise RunLockError(lock_path, f"已有运行锁未释放: {lock_path}。")
    with os.fdopen(fd, "w", encoding="utf-8") as fh:
        json.dump({"run_id": run_id, "pid": os.getpid(), "acquired_at": now_text()}, fh, ensure_ascii=False)
    return lock_path


def release_run_lock(lock_path: Path | None) -> None:
    if lock_path and lock_path.exists():
        lock_path.unlink()


def prune_old_items(conn: sqlite3.Connection, retention_days: int) -> int:
    cutoff = (datetime.now() - timedelta(days=retention_days)).strftime("%Y-%m-%d %H:%M:%S")
    cur = conn.execute("DELETE FROM notified_items WHERE last_seen_at < ?", (cutoff,))
    conn.commit()
    return cur.rowcount


def _column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in str(cell_ref or "") if ch.isalpha()).upper()
    index = 0
    for ch in letters:
        index = index * 26 + ord(ch) - ord("A") + 1
    return max(index, 1)


def _cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//main:t", _NS))
    value = cell.find("main:v", _NS)
    if value is None or value.text is None:
        return ""
    raw = value.text
    if cell_type == "s":
        try:
            return shared_strings[int(raw)]
        except Exception:
            return raw
    return raw


def read_xlsx_workbook(path: Path) -> dict[str, list[list[str]]]:
    if not path.is_file():
        raise FileNotFoundError(f"Excel 文件不存在: {path}")
    with ZipFile(path) as archive:
        names = set(archive.namelist())
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall("main:si", _NS):
                shared_strings.append("".join(node.text or "" for node in item.findall(".//main:t", _NS)))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in relationships}
        result: dict[str, list[list[str]]] = {}

        for sheet in workbook.findall("main:sheets/main:sheet", _NS):
            sheet_name = sheet.attrib.get("name", "")
            rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
            target = rel_map.get(rid or "")
            if not target:
                continue
            sheet_path = target.lstrip("/")
            if not sheet_path.startswith("xl/"):
                sheet_path = "xl/" + sheet_path
            root = ET.fromstring(archive.read(sheet_path))
            rows: list[list[str]] = []
            for row in root.findall("main:sheetData/main:row", _NS):
                values: list[str] = []
                for cell in row.findall("main:c", _NS):
                    col_index = _column_index(cell.attrib.get("r", "")) - 1
                    while len(values) < col_index:
                        values.append("")
                    values.append(clean_text(_cell_text(cell, shared_strings)))
                rows.append(values)
            result[sheet_name] = rows
        return result


def canonical_header(value: str) -> str:
    normalized = normalize_for_key(value)
    for canonical, aliases in HEADER_ALIASES.items():
        if normalized in {normalize_for_key(alias) for alias in aliases}:
            return canonical
    return clean_text(value)


def rows_to_dicts(rows: list[list[str]]) -> list[dict[str, str]]:
    if not rows:
        return []
    headers = [canonical_header(value) for value in rows[0]]
    dict_rows = []
    for row in rows[1:]:
        item = {}
        for index, header in enumerate(headers):
            if not header:
                continue
            item[header] = clean_text(row[index] if index < len(row) else "")
        if any(item.values()):
            dict_rows.append(item)
    return dict_rows


def find_detail_sheet(workbook: dict[str, list[list[str]]]) -> list[list[str]]:
    for name, rows in workbook.items():
        if clean_text(name) in SHEET_DETAIL_NAMES:
            return rows
    for rows in workbook.values():
        if not rows:
            continue
        headers = {canonical_header(value) for value in rows[0]}
        if {"店铺", "站点", "异常分类"}.issubset(headers):
            return rows
    raise ValueError("未找到账号状况异常明细 Sheet。")


def find_latest_excel(directory: Path) -> Path:
    if not directory.is_dir():
        raise FileNotFoundError(f"结果目录不存在: {directory}")
    candidates = [
        path
        for path in directory.glob("*.xlsx")
        if not path.name.startswith("~$")
        and not path.name.startswith("account-health-")
        and path.is_file()
    ]
    if not candidates:
        raise FileNotFoundError(f"结果目录中没有可作为源数据的 xlsx 文件: {directory}")
    return max(candidates, key=lambda path: path.stat().st_mtime)


def load_expected_stores(path: Path, target_site: str = SITE_US) -> list[str]:
    if not path.is_file():
        return []
    workbook = read_xlsx_workbook(path)
    selected_rows: list[list[str]] | None = None
    for name, rows in workbook.items():
        if clean_text(name) == "店铺执行清单":
            selected_rows = rows
            break
    if selected_rows is None:
        selected_rows = next(iter(workbook.values()), [])
    stores: list[str] = []
    seen: set[str] = set()
    for row in rows_to_dicts(selected_rows):
        store = clean_text(row.get("店铺"))
        site = clean_text(row.get("站点"))
        if not store or site != target_site:
            continue
        key = normalize_for_key(store)
        if key in seen:
            continue
        seen.add(key)
        stores.append(store)
    return stores


def extract_products(text: str) -> list[dict[str, str]]:
    body = clean_text(text)
    body = re.sub(r"([A-Z0-9]{10})(SKU\s*[:：])", r"\1 \2", body, flags=re.IGNORECASE)
    asin_matches = list(
        re.finditer(
            r"ASIN[^A-Z0-9]{0,8}([A-Z0-9]{10})(?=\s*SKU\b|\s|$|[,，;；|])",
            body,
            re.IGNORECASE,
        )
    )
    if not asin_matches:
        sku_match = re.search(r"SKU[^A-Z0-9]{0,8}([A-Z0-9][A-Z0-9._/-]{0,80})", body, re.IGNORECASE)
        return [{"asin": "", "sku": sku_match.group(1).strip() if sku_match else ""}]

    products: list[dict[str, str]] = []
    for index, match in enumerate(asin_matches):
        segment_end = asin_matches[index + 1].start() if index + 1 < len(asin_matches) else len(body)
        segment = body[match.start() : segment_end]
        sku_match = re.search(r"SKU[^A-Z0-9]{0,8}([A-Z0-9][A-Z0-9._/-]{0,80})", segment, re.IGNORECASE)
        products.append(
            {
                "asin": match.group(1).upper(),
                "sku": sku_match.group(1).strip() if sku_match else "",
            }
        )
    return products


def item_from_detail_row(row: dict[str, str], source_file: str) -> list[ImpactItem]:
    site = clean_text(row.get("站点") or SITE_US)
    impacted_text = clean_text(row.get("哪些商品会受到影响？"))
    direct_asin = clean_text(row.get("ASIN")).upper()
    direct_sku = clean_text(row.get("SKU"))
    if direct_asin or direct_sku:
        products = [{"asin": direct_asin, "sku": direct_sku}]
        if not impacted_text:
            impacted_text = clean_text(f"ASIN: {direct_asin} SKU: {direct_sku}")
    else:
        products = extract_products(impacted_text)
    items = []
    for product in products:
        if not clean_text(product.get("asin")) and not clean_text(product.get("sku")):
            continue
        items.append(
            ImpactItem(
                store=clean_text(row.get("店铺")),
                site=site,
                category=clean_text(row.get("异常分类")),
                asin=clean_text(product.get("asin")),
                sku=clean_text(product.get("sku")),
                reason=clean_text(row.get("原因")),
                date=clean_text(row.get("日期")),
                impacted_text=impacted_text,
                sales_risk=clean_text(row.get("存在销售风险")),
                action=clean_text(row.get("采取的操作")),
                rating_impact=clean_text(row.get("账户状况评级影响")),
                source_file=source_file,
            )
        )
    return items


def configured_max_excel_age_hours(config: dict[str, Any]) -> float:
    raw = config.get("source", {}).get("max_excel_age_hours", 0)
    if raw in (None, ""):
        return 0.0
    try:
        value = float(raw)
    except Exception:
        raise ValueError("source.max_excel_age_hours 必须是数字")
    if value < 0:
        raise ValueError("source.max_excel_age_hours 不能小于 0")
    return value


def ensure_excel_source_fresh(config: dict[str, Any], excel_path: Path) -> None:
    max_age_hours = configured_max_excel_age_hours(config)
    if max_age_hours <= 0:
        return
    modified_at = excel_path.stat().st_mtime
    age_hours = max(0.0, (time.time() - modified_at) / 3600)
    if age_hours > max_age_hours:
        modified_text = datetime.fromtimestamp(modified_at).strftime("%Y-%m-%d %H:%M:%S")
        raise SourceStaleError(
            excel_path,
            f"Excel 数据源已过期: {excel_path}; 最后修改时间 {modified_text}, "
            f"文件年龄 {age_hours:.1f} 小时, 超过 source.max_excel_age_hours={max_age_hours:g}。",
        )


def load_items_from_excel(path: Path, target_site: str = SITE_US) -> list[ImpactItem]:
    workbook = read_xlsx_workbook(path)
    rows = rows_to_dicts(find_detail_sheet(workbook))
    items: list[ImpactItem] = []
    for row in rows:
        site = clean_text(row.get("站点") or target_site)
        if site and site != target_site:
            continue
        items.extend(item_from_detail_row(row, str(path)))
    return dedupe_items(items)


def load_covered_stores_from_excel(path: Path, target_site: str = SITE_US) -> list[str]:
    if not path.is_file():
        return []
    workbook = read_xlsx_workbook(path)
    target_rows = workbook.get("Target Results")
    if not target_rows:
        return []
    stores: list[str] = []
    seen: set[str] = set()
    for row in rows_to_dicts(target_rows):
        status = clean_text(row.get("status"))
        store = clean_text(row.get("store") or row.get("店铺"))
        site = clean_text(row.get("site") or row.get("站点") or target_site)
        if status != "success" or not store or not site_matches(site, target_site):
            continue
        key = normalize_for_key(store)
        if key in seen:
            continue
        seen.add(key)
        stores.append(store)
    return stores


def dedupe_items(items: list[ImpactItem]) -> list[ImpactItem]:
    result: list[ImpactItem] = []
    seen: set[tuple[str, str]] = set()
    for item in items:
        signature = (item.dedupe_key, item.content_hash)
        if signature in seen:
            continue
        seen.add(signature)
        result.append(item)
    return result


def load_sample_items() -> list[ImpactItem]:
    return [
        ImpactItem(
            store="BYF",
            site=SITE_US,
            category="食品和商品安全问题",
            asin="B0FPKWYF49",
            sku="BF-9901-BN",
            reason="安全饮用水法案: 食品和商品安全问题",
            date="2026-01-30",
            impacted_text="Bathroom Faucet Brushed Nickel ASIN: B0FPKWYF49 SKU: BF-9901-BN",
            sales_risk="过去 12 个月无销量",
            action="商品已移除",
            rating_impact="无影响",
            source_file="sample",
        )
    ]


def load_items(config: dict[str, Any], args: argparse.Namespace, enforce_freshness: bool = False) -> tuple[list[ImpactItem], str]:
    source = config.get("source", {})
    source_type = args.source_type or source.get("type") or "excel_latest"
    target_site = args.site or source.get("site") or SITE_US
    if args.source_excel:
        source_type = "excel"

    if source_type == "sample":
        return load_sample_items(), "sample"
    if source_type == "excel":
        excel_path = Path(args.source_excel or source.get("excel_path") or "")
        if enforce_freshness:
            ensure_excel_source_fresh(config, excel_path)
        items = load_items_from_excel(excel_path, target_site=target_site)
        return items, str(excel_path)
    if source_type == "excel_latest":
        source_dir = Path(args.source_dir or source.get("excel_dir") or DEFAULT_SOURCE_DIR)
        excel_path = find_latest_excel(source_dir)
        if enforce_freshness:
            ensure_excel_source_fresh(config, excel_path)
        items = load_items_from_excel(excel_path, target_site=target_site)
        return items, str(excel_path)
    if source_type == "bridge":
        raise NotImplementedError("bridge 采集入口已预留, 但当前版本尚未接入跨店铺紫鸟页面自动采集。")
    raise ValueError(f"未知 source.type: {source_type}")


def select_notify_candidates(conn: sqlite3.Connection, items: list[ImpactItem]) -> list[ImpactItem]:
    candidates = []
    seen_at = now_text()
    for item in items:
        existing = conn.execute(
            "SELECT content_hash FROM notified_items WHERE dedupe_key = ?",
            (item.dedupe_key,),
        ).fetchone()
        if existing is None or existing["content_hash"] != item.content_hash:
            candidates.append(item)
        elif existing is not None:
            conn.execute(
                "UPDATE notified_items SET last_seen_at = ? WHERE dedupe_key = ?",
                (seen_at, item.dedupe_key),
            )
    conn.commit()
    return candidates


def mark_notified(conn: sqlite3.Connection, items: list[ImpactItem], notified_at: str) -> None:
    for item in items:
        payload = json.dumps(item.__dict__, ensure_ascii=False, sort_keys=True)
        existing = conn.execute(
            "SELECT notify_count FROM notified_items WHERE dedupe_key = ?",
            (item.dedupe_key,),
        ).fetchone()
        notify_count = int(existing["notify_count"]) + 1 if existing else 1
        conn.execute(
            """
            INSERT INTO notified_items (
                dedupe_key, content_hash, store, site, category, asin, sku, reason,
                issue_date, first_seen_at, last_seen_at, notified_at, notify_count, payload_json
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(dedupe_key) DO UPDATE SET
                content_hash = excluded.content_hash,
                store = excluded.store,
                site = excluded.site,
                category = excluded.category,
                asin = excluded.asin,
                sku = excluded.sku,
                reason = excluded.reason,
                issue_date = excluded.issue_date,
                last_seen_at = excluded.last_seen_at,
                notified_at = excluded.notified_at,
                notify_count = excluded.notify_count,
                payload_json = excluded.payload_json
            """,
            (
                item.dedupe_key,
                item.content_hash,
                item.store,
                item.site,
                item.category,
                item.asin,
                item.sku,
                item.reason,
                item.date,
                notified_at,
                notified_at,
                notified_at,
                notify_count,
                payload,
            ),
        )
    conn.commit()


def chunked(items: list[ImpactItem], size: int) -> list[list[ImpactItem]]:
    size = max(1, int(size or 60))
    return [items[index : index + size] for index in range(0, len(items), size)]


def notification_title(now: datetime | None = None, suffix: str = "") -> str:
    current = now or datetime.now()
    title = f"{current.year}年{current.month}月{current.day}日亚马逊账号状况异常新增通知"
    suffix = clean_text(suffix)
    if suffix:
        return f"{title}{suffix}"
    return title


def ordered_unique_text(values: list[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for value in values:
        text = clean_text(value)
        if not text:
            continue
        key = normalize_for_key(text)
        if key in seen:
            continue
        seen.add(key)
        result.append(text)
    return result


def compact_failure_reason(value: str, limit: int = 120) -> str:
    text = clean_text(value)
    if not text:
        return ""
    text = re.sub(r"\s+", " ", text)
    if len(text) <= limit:
        return text
    return text[: max(0, limit - 1)].rstrip() + "…"


def dedupe_failed_targets(failed_targets: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    for target in failed_targets:
        store = clean_text(target.get("store"))
        site = clean_text(target.get("site") or SITE_US)
        if not store:
            continue
        signature = (normalize_for_key(store), normalize_for_key(site))
        if signature in seen:
            continue
        seen.add(signature)
        result.append(dict(target))
    return result


def _count_by(items: list[ImpactItem], field: str) -> dict[str, int]:
    counts: dict[str, int] = {}
    for item in items:
        value = clean_text(getattr(item, field, "")) or "未识别"
        counts[value] = counts.get(value, 0) + 1
    return dict(sorted(counts.items(), key=lambda pair: (-pair[1], pair[0])))


def _group_items(items: list[ImpactItem]) -> dict[str, dict[str, list[ImpactItem]]]:
    grouped: dict[str, dict[str, list[ImpactItem]]] = {}
    for item in sorted(items, key=lambda value: (value.store, value.category, value.date, value.asin, value.sku)):
        store = item.store or "未识别店铺"
        category = item.category or "未识别问题"
        grouped.setdefault(store, {}).setdefault(category, []).append(item)
    return grouped


DINGTALK_MARKDOWN_TEMPLATE = "field-block-v1"


def render_dingtalk_issue_field_block(
    index: int,
    category: str,
    date_text: str,
    action_text: str,
    asin: str,
    sku: str,
) -> list[str]:
    return [
        f"{index}. 问题类型: {category}  ",
        f"   日期: {date_text}  ",
        f"   当前处理: {action_text}  ",
        f"   ASIN: **{asin}**  ",
        f"   SKU: **{sku}**",
    ]


def render_markdown(
    items: list[ImpactItem],
    title: str,
    chunk_index: int,
    chunk_total: int,
    failed_targets: list[dict[str, Any]] | None = None,
) -> str:
    failed_targets = dedupe_failed_targets(list(failed_targets or []))
    store_counts = _count_by(items, "store")
    category_counts = _count_by(items, "category")
    store_summary = ", ".join(f"{store} {count}条" for store, count in store_counts.items()) or "无"
    category_summary = ", ".join(f"{category} {count}条" for category, count in category_counts.items()) or "无"
    lines = [
        f"### {title}",
        "",
        "**结论**",
        f"- 新增/变化: {len(items)} 条",
        f"- 涉及店铺: {len(store_counts)} 个, {store_summary}",
        f"- 问题类型: {category_summary}",
        f"- 范围: {SITE_US}站, 未解决账号状况异常",
    ]
    if failed_targets:
        failed_store_names = ordered_unique_text([clean_text(item.get("store")) for item in failed_targets])
        failed_store_summary = ", ".join(failed_store_names[:8])
        if len(failed_store_names) > 8:
            failed_store_summary = f"{failed_store_summary} 等 {len(failed_store_names)} 个"
        lines.append(f"- 采集未完成店铺: {len(failed_store_names)} 个, {failed_store_summary}")
    if chunk_total > 1:
        lines.append(f"- 分段: {chunk_index}/{chunk_total}")

    for store, categories in _group_items(items).items():
        store_total = sum(len(group) for group in categories.values())
        category_summary_line = ", ".join(f"{category}: {len(group)} 条" for category, group in categories.items())
        store_summary_line = f"共 {store_total} 条"
        if category_summary_line:
            store_summary_line = f"{store_summary_line} {category_summary_line}"
        lines.extend(["", "---", "", f"#### 店铺: {store}", store_summary_line])
        for category, group in categories.items():
            for index, item in enumerate(group, start=1):
                asin = item.asin or "未识别"
                sku = item.sku or "未识别"
                date_text = item.date or "未识别"
                action_text = item.action or "待确认"
                lines.extend(
                    render_dingtalk_issue_field_block(index, category, date_text, action_text, asin, sku)
                )
            lines.append("")
    if failed_targets:
        lines.extend(["", "---", "", "#### 未完成店铺"])
        for index, target in enumerate(failed_targets, start=1):
            store = clean_text(target.get("store")) or "未识别店铺"
            site = clean_text(target.get("site") or SITE_US)
            status = clean_text(target.get("status")) or "failed"
            error = compact_failure_reason(
                clean_text(target.get("error"))
                or clean_text(target.get("coverage_error"))
                or clean_text(target.get("reason"))
                or "补采仍未成功"
            )
            lines.append(f"{index}. 店铺: {store} ({site})  ")
            lines.append(f"   状态: {status}  ")
            if error:
                lines.append(f"   原因: {error}")
            lines.append("")
    return "\n".join(lines).strip() + "\n"


def run_dws_call(dws_call: Path, args: list[str], state_dir: Path, timeout: int = 60) -> subprocess.CompletedProcess[str]:
    if not dws_call.is_file():
        raise FileNotFoundError(f"DWS 调用入口不存在: {dws_call}")
    args_dir = state_dir / "dws-args"
    args_dir.mkdir(parents=True, exist_ok=True)
    args_file = args_dir / f"args-{run_id_text()}.json"
    with args_file.open("w", encoding="utf-8") as fh:
        json.dump({"args": args}, fh, ensure_ascii=False)
    return subprocess.run(
        [str(dws_call), str(args_file)],
        cwd=str(PROJECT_ROOT),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        timeout=timeout,
    )


def dingtalk_send_method(dingtalk: dict[str, Any]) -> str:
    method = clean_text(dingtalk.get("method")).lower()
    if method:
        return method
    if clean_text(dingtalk.get("webhook_url")):
        return "webhook"
    return "dws"


def build_dingtalk_signed_url(webhook_url: str, secret: str, timestamp_ms: int | None = None) -> str:
    timestamp = str(timestamp_ms if timestamp_ms is not None else round(time.time() * 1000))
    string_to_sign = f"{timestamp}\n{secret}"
    hmac_code = hmac.new(
        secret.encode("utf-8"),
        string_to_sign.encode("utf-8"),
        digestmod=hashlib.sha256,
    ).digest()
    sign = urlparse.quote_plus(base64.b64encode(hmac_code))
    separator = "&" if "?" in webhook_url else "?"
    return f"{webhook_url}{separator}timestamp={timestamp}&sign={sign}"


def send_dingtalk_webhook(
    webhook_url: str,
    secret: str,
    title: str,
    markdown: str,
    dry_run: bool,
    timeout: int = 30,
) -> subprocess.CompletedProcess[str]:
    if not webhook_url:
        if not dry_run:
            raise ValueError("缺少 dingtalk.webhook_url。")
        webhook_url = "https://oapi.dingtalk.com/robot/send?access_token=DRY_RUN"
    if not secret:
        if not dry_run:
            raise ValueError("缺少 dingtalk.secret。")
        secret = "DRY_RUN_SECRET"

    if dry_run:
        stdout = json.dumps(
            {
                "ok": True,
                "dry_run": True,
                "method": "webhook",
                "title": title,
                "markdown_chars": len(markdown),
            },
            ensure_ascii=False,
        )
        return subprocess.CompletedProcess(["dingtalk-webhook", "--dry-run"], 0, stdout, "")

    payload = {
        "msgtype": "markdown",
        "markdown": {
            "title": title,
            "text": markdown,
        },
    }
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urlrequest.Request(
        build_dingtalk_signed_url(webhook_url, secret),
        data=body,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    try:
        with urlrequest.urlopen(request, timeout=timeout) as response:
            response_text = response.read().decode("utf-8", errors="replace")
    except urlerror.HTTPError as exc:
        response_text = exc.read().decode("utf-8", errors="replace")
        return subprocess.CompletedProcess(["dingtalk-webhook", "--send"], 1, response_text, f"HTTP {exc.code}")
    except Exception as exc:
        return subprocess.CompletedProcess(["dingtalk-webhook", "--send"], 1, "", str(exc))

    try:
        response_payload = json.loads(response_text)
    except json.JSONDecodeError:
        return subprocess.CompletedProcess(["dingtalk-webhook", "--send"], 1, response_text, "invalid_json_response")

    errcode = response_payload.get("errcode")
    if errcode == 0:
        return subprocess.CompletedProcess(["dingtalk-webhook", "--send"], 0, response_text, "")
    return subprocess.CompletedProcess(
        ["dingtalk-webhook", "--send"],
        1,
        response_text,
        clean_text(response_payload.get("errmsg")) or f"errcode={errcode}",
    )


def send_dingtalk_markdown(
    config: dict[str, Any],
    state_dir: Path,
    title: str,
    markdown: str,
    dry_run: bool,
) -> subprocess.CompletedProcess[str]:
    dingtalk = config.get("dingtalk", {})
    method = dingtalk_send_method(dingtalk)
    if method == "webhook":
        return send_dingtalk_webhook(
            clean_text(dingtalk.get("webhook_url")),
            clean_text(dingtalk.get("secret")),
            title,
            markdown,
            dry_run=dry_run,
        )
    if method != "dws":
        raise ValueError(f"不支持的 dingtalk.method: {method}")

    robot_code = clean_text(dingtalk.get("robot_code"))
    group_id = clean_text(dingtalk.get("group_open_conversation_id"))
    if not robot_code or not group_id:
        if not dry_run:
            raise ValueError("缺少 dingtalk.robot_code 或 dingtalk.group_open_conversation_id。")
        robot_code = robot_code or "DRY_RUN_ROBOT"
        group_id = group_id or "DRY_RUN_GROUP"

    message_dir = state_dir / "messages"
    message_dir.mkdir(parents=True, exist_ok=True)
    message_file = message_dir / f"message-{run_id_text()}.md"
    message_file.write_text(markdown, encoding="utf-8")

    args = [
        "chat",
        "message",
        "send-by-bot",
        "--robot-code",
        robot_code,
        "--group",
        group_id,
        "--title",
        title,
        "--text",
        f"@{message_file}",
        "-f",
        "json",
        "--yes",
    ]
    if dry_run:
        args.append("--dry-run")
    return run_dws_call(Path(dingtalk.get("dws_call") or DEFAULT_DWS_CALL), args, state_dir)


def record_attempt(
    conn: sqlite3.Connection,
    run_id: str,
    dry_run: bool,
    status: str,
    item_count: int,
    title: str,
    error: str = "",
    dws_stdout: str = "",
    dws_stderr: str = "",
) -> None:
    conn.execute(
        """
        INSERT INTO notification_attempts (
            run_id, attempted_at, dry_run, status, item_count, title, error, dws_stdout, dws_stderr
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (run_id, now_text(), int(dry_run), status, item_count, title, error, dws_stdout, dws_stderr),
    )
    conn.commit()


def write_run_start(conn: sqlite3.Connection, run_id: str, started_at: str, source: str) -> None:
    conn.execute(
        "INSERT OR REPLACE INTO runs (run_id, started_at, source, status) VALUES (?, ?, ?, ?)",
        (run_id, started_at, source, "running"),
    )
    conn.commit()


def write_run_end(
    conn: sqlite3.Connection,
    run_id: str,
    total_items: int,
    notify_candidates: int,
    sent_items: int,
    status: str,
    error: str,
) -> None:
    conn.execute(
        """
        UPDATE runs
        SET ended_at = ?, total_items = ?, notify_candidates = ?, sent_items = ?, status = ?, error = ?
        WHERE run_id = ?
        """,
        (now_text(), total_items, notify_candidates, sent_items, status, clean_text(error), run_id),
    )
    conn.commit()


def _column_letter(index: int) -> str:
    letters = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters or "A"


def _xml_text(value: Any) -> str:
    return html.escape(clean_text(value), quote=False)


def _worksheet_xml(headers: list[str], rows: list[dict[str, Any]]) -> str:
    table_rows = [headers] + [[row.get(header, "") for header in headers] for row in rows]
    row_xml = []
    for row_index, row in enumerate(table_rows, start=1):
        cells = []
        for col_index, value in enumerate(row, start=1):
            ref = f"{_column_letter(col_index)}{row_index}"
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{_xml_text(value)}</t></is></c>')
        row_xml.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    dimension = f"A1:{_column_letter(max(len(headers), 1))}{max(len(table_rows), 1)}"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<dimension ref="{dimension}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'<sheetData>{"".join(row_xml)}</sheetData>'
        f'<autoFilter ref="{dimension}"/>'
        '</worksheet>'
    )


def write_run_xlsx(path: Path, item_rows: list[dict[str, Any]], run_rows: list[dict[str, Any]]) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    created = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            "</Types>",
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            "</Relationships>",
        )
        archive.writestr(
            "docProps/core.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            "<dc:creator>YD-MCP</dc:creator>"
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>'
            "</cp:coreProperties>",
        )
        archive.writestr(
            "docProps/app.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>YD-MCP</Application></Properties>",
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="待通知明细" sheetId="1" r:id="rId1"/>'
            '<sheet name="执行结果" sheetId="2" r:id="rId2"/></sheets>'
            "</workbook>",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            "</Relationships>",
        )
        archive.writestr(
            "xl/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            "</styleSheet>",
        )
        archive.writestr("xl/worksheets/sheet1.xml", _worksheet_xml(ITEM_HEADERS, item_rows))
        archive.writestr("xl/worksheets/sheet2.xml", _worksheet_xml(RUN_HEADERS, run_rows))
    return str(path)


def update_run_artifacts(
    config: dict[str, Any],
    run_id: str,
    item_rows: list[dict[str, Any]],
    run_row: dict[str, Any],
) -> Path:
    result_dir = Path(config.get("state", {}).get("result_dir") or DEFAULT_STATE_DIR / "runs")
    result_path = result_dir / f"account-health-notifier-{run_id}.xlsx"
    write_run_xlsx(result_path, item_rows, [run_row])
    return result_path


def issue(code: str, message: str, severity: str = "error") -> dict[str, str]:
    return {"code": code, "severity": severity, "message": message}


class PrimaryHostError(Exception):
    pass


def current_host_name() -> str:
    return clean_text(os.environ.get("COMPUTERNAME") or os.environ.get("HOSTNAME") or socket.gethostname())


def primary_host_policy(config: dict[str, Any]) -> dict[str, Any]:
    runtime = config.get("runtime", {})
    primary_host = clean_text(runtime.get("primary_host"))
    enforce = bool(runtime.get("enforce_primary_host_for_send", True))
    current_host = current_host_name()
    allowed = (not enforce) or (primary_host and normalize_for_key(primary_host) == normalize_for_key(current_host))
    return {
        "current_host": current_host,
        "primary_host": primary_host,
        "enforce": enforce,
        "allowed": allowed,
    }


def ensure_primary_host_allowed(config: dict[str, Any]) -> None:
    policy = primary_host_policy(config)
    if not policy["enforce"]:
        return
    if not policy["primary_host"]:
        raise PrimaryHostError("缺少 runtime.primary_host, 无法确认当前机器是否允许真实发送。")
    if not policy["allowed"]:
        raise PrimaryHostError(
            f"当前主机 {policy['current_host']} 不允许真实发送; 仅允许 {policy['primary_host']} 执行生产通知。"
        )


def schedule_command_name(config: dict[str, Any]) -> str:
    return clean_text(config.get("schedule", {}).get("command") or "production-run") or "production-run"


def should_require_source_excel_for_runtime(config: dict[str, Any]) -> bool:
    return schedule_command_name(config) == "run"


def collector_backend_name(config: dict[str, Any], override: str = "") -> str:
    backend = clean_text(override or config.get("collector", {}).get("backend") or "webdriver").lower()
    return backend or "webdriver"


def should_require_collector_for_runtime(config: dict[str, Any]) -> bool:
    return schedule_command_name(config) == "production-run"


def validate_runtime_config(
    config_path: Path,
    config: dict[str, Any],
    state_dir: Path,
    require_send_ready: bool = False,
    require_source_excel: bool = True,
    require_collector_ready: bool = False,
) -> dict[str, Any]:
    issues: list[dict[str, str]] = []
    collector_backend = collector_backend_name(config)
    if collector_backend not in {"webdriver", "zclaw"}:
        issues.append(issue("collector_backend_invalid", "collector.backend must be webdriver or zclaw"))
    if require_collector_ready and collector_backend == "webdriver":
        webdriver_config = config.get("ziniao_webdriver", {})
        client_path_text = clean_text(webdriver_config.get("client_path"))
        client_path = Path(client_path_text or "")
        if not client_path_text:
            issues.append(issue("webdriver_client_path_missing", "ziniao_webdriver.client_path is required"))
        elif not client_path.is_file():
            issues.append(issue("webdriver_client_path_missing", f"WebDriver client not found: {client_path}"))
        try:
            if int(webdriver_config.get("port") or 0) <= 0:
                issues.append(issue("webdriver_port_invalid", "ziniao_webdriver.port must be > 0"))
        except Exception:
            issues.append(issue("webdriver_port_invalid", "ziniao_webdriver.port must be an integer"))
        if not clean_text(webdriver_config.get("company")):
            issues.append(issue("webdriver_company_missing", "ziniao_webdriver.company is required"))
        if not clean_text(webdriver_config.get("username")):
            issues.append(issue("webdriver_username_missing", "ziniao_webdriver.username is required"))
        password = clean_text(webdriver_config.get("password"))
        password_env = clean_text(webdriver_config.get("password_env"))
        env_password = clean_text(os.environ.get(password_env)) if password_env else ""
        if not password and not env_password:
            issues.append(
                issue(
                    "webdriver_password_missing",
                    "ziniao_webdriver.password is required, or ziniao_webdriver.password_env must point to an existing environment variable",
                )
            )
    if not config_path.is_file():
        issues.append(issue("config_missing", f"配置文件不存在: {config_path}"))

    source = config.get("source", {})
    source_type = source.get("type") or "excel_latest"
    if require_source_excel:
        if source_type == "excel":
            excel_path = Path(source.get("excel_path") or "")
            if not excel_path.is_file():
                issues.append(issue("source_excel_missing", f"指定 Excel 不存在: {excel_path}"))
        elif source_type == "excel_latest":
            excel_dir = Path(source.get("excel_dir") or "")
            if not excel_dir.is_dir():
                issues.append(issue("source_excel_dir_missing", f"结果目录不存在: {excel_dir}"))
        elif source_type == "bridge":
            issues.append(issue("bridge_not_implemented", "bridge 采集入口当前版本尚未接入"))

    try:
        configured_max_excel_age_hours(config)
    except Exception as exc:
        issues.append(issue("source_max_excel_age_hours_invalid", str(exc)))

    if config.get("notify", {}).get("require_all_stores_before_send", False):
        store_list_path = Path(source.get("store_list_path") or "")
        if not store_list_path.is_file():
            issues.append(issue("store_list_missing", f"店铺清单不存在: {store_list_path}"))
        else:
            try:
                expected_stores = load_expected_stores(store_list_path, source.get("site") or SITE_US)
                if not expected_stores:
                    issues.append(issue("store_list_empty", f"店铺清单没有目标站点店铺: {store_list_path}"))
            except Exception as exc:
                issues.append(issue("store_list_invalid", f"店铺清单读取失败: {exc}"))

    dingtalk = config.get("dingtalk", {})
    method = dingtalk_send_method(dingtalk)
    if method == "dws":
        dws_call = Path(dingtalk.get("dws_call") or DEFAULT_DWS_CALL)
        if not dws_call.is_file():
            issues.append(issue("dws_call_missing", f"DWS 调用入口不存在: {dws_call}"))
    elif method != "webhook":
        issues.append(issue("dingtalk_method_invalid", "dingtalk.method 必须是 webhook 或 dws"))

    notify = config.get("notify", {})
    try:
        if int(notify.get("max_items_per_message") or 0) <= 0:
            issues.append(issue("max_items_per_message_invalid", "max_items_per_message 必须大于 0"))
    except Exception:
        issues.append(issue("max_items_per_message_invalid", "max_items_per_message 必须是正整数"))
    try:
        if int(notify.get("dedupe_retention_days") or 0) <= 0:
            issues.append(issue("dedupe_retention_days_invalid", "dedupe_retention_days 必须大于 0"))
    except Exception:
        issues.append(issue("dedupe_retention_days_invalid", "dedupe_retention_days 必须是正整数"))

    schedule = config.get("schedule", {})
    try:
        if int(schedule.get("interval_hours") or 0) <= 0:
            issues.append(issue("interval_hours_invalid", "interval_hours 必须大于 0"))
    except Exception:
        issues.append(issue("interval_hours_invalid", "interval_hours 必须是正整数"))

    primary_host = primary_host_policy(config)
    if require_send_ready and primary_host["enforce"]:
        if not primary_host["primary_host"]:
            issues.append(issue("primary_host_missing", "缺少 runtime.primary_host, 单主机生产模式无法确认唯一发送机器"))
        elif not primary_host["allowed"]:
            issues.append(
                issue(
                    "primary_host_mismatch",
                    f"当前主机 {primary_host['current_host']} 与 runtime.primary_host={primary_host['primary_host']} 不一致",
                )
            )

    try:
        configured_run_lock_ttl_minutes(config)
    except Exception as exc:
        issues.append(issue("run_lock_ttl_minutes_invalid", str(exc)))

    db_path = Path(config.get("state", {}).get("db_path") or state_dir / "state.sqlite")
    result_dir = Path(config.get("state", {}).get("result_dir") or DEFAULT_STATE_DIR / "runs")

    if require_send_ready:
        if method == "webhook":
            if not clean_text(dingtalk.get("webhook_url")):
                issues.append(issue("webhook_url_missing", "缺少 dingtalk.webhook_url"))
            if not clean_text(dingtalk.get("secret")):
                issues.append(issue("webhook_secret_missing", "缺少 dingtalk.secret"))
        elif method == "dws":
            if not clean_text(dingtalk.get("robot_code")):
                issues.append(issue("robot_code_missing", "缺少 dingtalk.robot_code"))
            if not clean_text(dingtalk.get("group_open_conversation_id")):
                issues.append(issue("group_open_conversation_id_missing", "缺少 dingtalk.group_open_conversation_id"))
        if not bool(dingtalk.get("send_enabled", False)):
            issues.append(issue("send_enabled_false", "定时真实通知需要 dingtalk.send_enabled=true"))

    ok = not any(item["severity"] == "error" for item in issues)
    return {
        "ok": ok,
        "require_send_ready": require_send_ready,
        "config_path": str(config_path),
        "state_dir": str(state_dir),
        "db_path": str(db_path),
        "result_dir": str(result_dir),
        "issues": issues,
    }


def summarize_items(items: list[ImpactItem], expected_stores: list[str] | None = None) -> dict[str, Any]:
    by_store: dict[str, int] = {}
    by_category: dict[str, int] = {}
    by_store_category: dict[str, dict[str, int]] = {}
    missing_asin = 0
    missing_sku = 0
    for item in items:
        store = item.store or "未识别店铺"
        category = item.category or "未识别分类"
        by_store[store] = by_store.get(store, 0) + 1
        by_category[category] = by_category.get(category, 0) + 1
        by_store_category.setdefault(store, {})
        by_store_category[store][category] = by_store_category[store].get(category, 0) + 1
        if not item.asin:
            missing_asin += 1
        if not item.sku:
            missing_sku += 1
    expected_stores = expected_stores or []
    parsed_store_keys = {normalize_for_key(store) for store in by_store}
    missing_stores = [
        store for store in expected_stores if normalize_for_key(store) not in parsed_store_keys
    ]
    extra_stores = [
        store
        for store in by_store
        if normalize_for_key(store) not in {normalize_for_key(item) for item in expected_stores}
    ] if expected_stores else []
    coverage_ok = not expected_stores or not missing_stores
    return {
        "total_items": len(items),
        "store_count": len(by_store),
        "expected_store_count": len(expected_stores),
        "coverage_ok": coverage_ok,
        "missing_stores": missing_stores,
        "extra_stores": extra_stores,
        "category_count": len(by_category),
        "missing_asin": missing_asin,
        "missing_sku": missing_sku,
        "by_store": dict(sorted(by_store.items())),
        "by_category": dict(sorted(by_category.items())),
        "by_store_category": {
            store: dict(sorted(categories.items()))
            for store, categories in sorted(by_store_category.items())
        },
    }


def build_coverage_summary(config: dict[str, Any], args: argparse.Namespace, items: list[ImpactItem]) -> dict[str, Any]:
    source_config = config.get("source", {})
    store_list_path = Path(getattr(args, "store_list", "") or source_config.get("store_list_path") or "")
    target_site = getattr(args, "site", "") or source_config.get("site") or SITE_US
    expected_stores = load_expected_stores(store_list_path, target_site)
    summary = summarize_items(items, expected_stores=expected_stores)
    source_excel_text = clean_text(getattr(args, "source_excel", "") or source_config.get("excel_path") or "")
    covered_stores = (
        load_covered_stores_from_excel(Path(source_excel_text), target_site)
        if source_excel_text
        else []
    )
    if expected_stores and covered_stores:
        covered_keys = {normalize_for_key(store) for store in covered_stores}
        expected_keys = {normalize_for_key(store) for store in expected_stores}
        summary["missing_stores"] = [
            store for store in expected_stores if normalize_for_key(store) not in covered_keys
        ]
        summary["extra_stores"] = [
            store for store in covered_stores if normalize_for_key(store) not in expected_keys
        ]
        summary["coverage_ok"] = not summary["missing_stores"]
        summary["covered_store_count"] = len(covered_stores)
        summary["covered_stores"] = covered_stores
    summary["store_list_path"] = str(store_list_path)
    summary["target_site"] = target_site
    return summary


def fallback_coverage_summary(
    config: dict[str, Any],
    args: argparse.Namespace,
    items: list[ImpactItem],
    error: str,
) -> dict[str, Any]:
    source_config = config.get("source", {})
    store_list_path = Path(getattr(args, "store_list", "") or source_config.get("store_list_path") or "")
    target_site = getattr(args, "site", "") or source_config.get("site") or SITE_US
    summary = summarize_items(items, expected_stores=[])
    summary["coverage_ok"] = False
    summary["store_list_path"] = str(store_list_path)
    summary["target_site"] = target_site
    summary["error"] = clean_text(error)
    return summary


def selected_source_type(config: dict[str, Any], args: argparse.Namespace) -> str:
    source_config = config.get("source", {})
    if getattr(args, "source_excel", ""):
        return "excel"
    return getattr(args, "source_type", "") or source_config.get("type") or "excel_latest"


def should_require_store_coverage(config: dict[str, Any], args: argparse.Namespace) -> bool:
    if getattr(args, "skip_store_coverage", False):
        return False
    if getattr(args, "require_all_stores", False):
        return True
    if selected_source_type(config, args) == "sample":
        return False
    return bool(config.get("notify", {}).get("require_all_stores_before_send", False))


def should_require_complete_product_ids(config: dict[str, Any]) -> bool:
    return bool(config.get("notify", {}).get("require_complete_product_ids_before_send", True))


def coverage_failure_message(summary: dict[str, Any]) -> str:
    expected_count = int(summary.get("expected_store_count") or 0)
    if expected_count <= 0:
        return f"店铺清单为空或无法读取, 无法确认全店铺覆盖: {summary.get('store_list_path', '')}"
    missing_stores = summary.get("missing_stores", [])
    if missing_stores:
        shown = ", ".join(missing_stores[:20])
        suffix = "" if len(missing_stores) <= 20 else f" 等 {len(missing_stores)} 个"
        return f"店铺覆盖不完整, 缺失 {len(missing_stores)} 个美国站店铺: {shown}{suffix}"
    return ""


def product_quality_failure_message(summary: dict[str, Any]) -> str:
    missing_asin = int(summary.get("missing_asin") or 0)
    missing_sku = int(summary.get("missing_sku") or 0)
    if missing_asin <= 0 and missing_sku <= 0:
        return ""
    missing_parts = []
    if missing_asin:
        missing_parts.append(f"缺少 ASIN {missing_asin} 条")
    if missing_sku:
        missing_parts.append(f"缺少 SKU {missing_sku} 条")
    return f"商品标识不完整, {', '.join(missing_parts)}; 请先修复采集或解析结果再通知。"


def summary_rows(summary: dict[str, Any]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if summary.get("error"):
        rows.append({"维度": "解析错误", "名称": summary.get("error", ""), "数量": 0})
    for store, count in summary.get("by_store", {}).items():
        rows.append({"维度": "店铺", "名称": store, "数量": count})
    for category, count in summary.get("by_category", {}).items():
        rows.append({"维度": "异常分类", "名称": category, "数量": count})
    for store, categories in summary.get("by_store_category", {}).items():
        for category, count in categories.items():
            rows.append({"维度": f"店铺/异常分类:{store}", "名称": category, "数量": count})
    for store in summary.get("missing_stores", []):
        rows.append({"维度": "覆盖缺失店铺", "名称": store, "数量": 0})
    for store in summary.get("extra_stores", []):
        rows.append({"维度": "清单外店铺", "名称": store, "数量": summary.get("by_store", {}).get(store, 0)})
    rows.append({"维度": "质量", "名称": "缺少ASIN", "数量": summary.get("missing_asin", 0)})
    rows.append({"维度": "质量", "名称": "缺少SKU", "数量": summary.get("missing_sku", 0)})
    return rows


def write_parse_xlsx(path: Path, item_rows: list[dict[str, Any]], summary: dict[str, Any]) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    created = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            "</Types>",
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            "</Relationships>",
        )
        archive.writestr(
            "docProps/core.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            "<dc:creator>YD-MCP</dc:creator>"
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>'
            "</cp:coreProperties>",
        )
        archive.writestr(
            "docProps/app.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>YD-MCP</Application></Properties>",
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="全店铺明细" sheetId="1" r:id="rId1"/>'
            '<sheet name="解析汇总" sheetId="2" r:id="rId2"/></sheets>'
            "</workbook>",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            "</Relationships>",
        )
        archive.writestr(
            "xl/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            "</styleSheet>",
        )
        archive.writestr("xl/worksheets/sheet1.xml", _worksheet_xml(ITEM_HEADERS, item_rows))
        archive.writestr("xl/worksheets/sheet2.xml", _worksheet_xml(PARSE_SUMMARY_HEADERS, summary_rows(summary)))
    return str(path)


def write_collect_open_xlsx(
    path: Path,
    item_rows: list[dict[str, Any]],
    target_rows: list[dict[str, Any]],
    summary: dict[str, Any],
) -> str:
    path.parent.mkdir(parents=True, exist_ok=True)
    created = datetime.now().strftime("%Y-%m-%dT%H:%M:%S")
    sheets = [
        ("CDP Items", ITEM_HEADERS, item_rows),
        ("Target Results", CDP_TARGET_HEADERS, target_rows),
        ("Coverage", PARSE_SUMMARY_HEADERS, summary_rows(summary)),
    ]
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        overrides = "".join(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for index in range(1, len(sheets) + 1)
        )
        archive.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            f"{overrides}"
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            "</Types>",
        )
        archive.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            "</Relationships>",
        )
        archive.writestr(
            "docProps/core.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" '
            'xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            "<dc:creator>YD-MCP</dc:creator>"
            f'<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>'
            f'<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>'
            "</cp:coreProperties>",
        )
        archive.writestr(
            "docProps/app.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>YD-MCP</Application></Properties>",
        )
        sheet_nodes = "".join(
            f'<sheet name="{html.escape(name, quote=True)}" sheetId="{index}" r:id="rId{index}"/>'
            for index, (name, _headers, _rows) in enumerate(sheets, start=1)
        )
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f"<sheets>{sheet_nodes}</sheets>"
            "</workbook>",
        )
        rel_nodes = "".join(
            f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
            for index in range(1, len(sheets) + 1)
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{rel_nodes}"
            f'<Relationship Id="rId{len(sheets) + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            "</Relationships>",
        )
        archive.writestr(
            "xl/styles.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border/></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            "</styleSheet>",
        )
        for index, (_name, headers, rows) in enumerate(sheets, start=1):
            archive.writestr(f"xl/worksheets/sheet{index}.xml", _worksheet_xml(headers, rows))
    return str(path)


def execute_parse(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    run_id = run_id_text()
    result_dir = Path(args.output_dir or config.get("state", {}).get("result_dir") or DEFAULT_STATE_DIR / "runs")
    artifact = result_dir / f"account-health-parse-{run_id}.xlsx"
    source = getattr(args, "source_excel", "") or getattr(args, "source_dir", "") or config.get("source", {}).get("excel_path") or config.get("source", {}).get("excel_dir") or "unknown"
    try:
        items, source = load_items(config, args)
        item_rows = [item.to_row(run_id, "parsed") for item in items]
        try:
            summary = build_coverage_summary(config, args, items)
            coverage_error = coverage_failure_message(summary) if args.require_all_stores else ""
            status = "coverage_failed" if coverage_error else "success"
            error = ""
        except Exception as exc:
            error = f"store_list_invalid: {exc}"
            summary = fallback_coverage_summary(config, args, items, error)
            coverage_error = error
            status = "failed"
    except Exception as exc:
        item_rows = []
        error = str(exc)
        summary = fallback_coverage_summary(config, args, [], error)
        coverage_error = ""
        status = "failed"
    write_parse_xlsx(artifact, item_rows, summary)
    ok = status == "success"
    return {
        "ok": ok,
        "status": status,
        "run_id": run_id,
        "source": source,
        "artifact": str(artifact),
        "coverage_error": coverage_error,
        "error": error,
        **summary,
    }


def execute_run(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path)
    state_dir = Path(args.state_dir or config_path.parent or DEFAULT_STATE_DIR)
    db_path = Path(config.get("state", {}).get("db_path") or state_dir / "state.sqlite")
    conn = connect_db(db_path)

    run_id = run_id_text()
    started_at = now_text()
    source = ""
    item_rows: list[dict[str, Any]] = []
    sent_items = 0
    candidates: list[ImpactItem] = []
    failed_targets = dedupe_failed_targets(list(getattr(args, "failed_targets", []) or []))
    status = "success"
    error = ""
    coverage_summary: dict[str, Any] = {}
    lock_path: Path | None = None

    try:
        lock_path = acquire_run_lock(state_dir, run_id, configured_run_lock_ttl_minutes(config))
        try:
            retention_days = int(config.get("notify", {}).get("dedupe_retention_days") or 90)
            if retention_days <= 0:
                raise ValueError
        except Exception:
            raise ValueError("dedupe_retention_days 必须是正整数")
        try:
            max_items = int(config.get("notify", {}).get("max_items_per_message") or 60)
            if max_items <= 0:
                raise ValueError
        except Exception:
            raise ValueError("max_items_per_message 必须是正整数")
        prune_old_items(conn, retention_days)
        items, source = load_items(config, args, enforce_freshness=True)
        write_run_start(conn, run_id, started_at, source)
        coverage_summary = build_coverage_summary(config, args, items)
        coverage_error = coverage_failure_message(coverage_summary) if should_require_store_coverage(config, args) else ""
        if coverage_error:
            status = "coverage_failed"
            error = coverage_error
            candidates = []
        elif should_require_complete_product_ids(config) and (quality_error := product_quality_failure_message(coverage_summary)):
            status = "quality_failed"
            error = quality_error
            candidates = []
        else:
            candidates = select_notify_candidates(conn, items)
            dry_run = bool(args.dry_run)
            if args.send:
                dry_run = False
            if not args.send and not config.get("dingtalk", {}).get("send_enabled", False):
                dry_run = True
            if not dry_run:
                ensure_primary_host_allowed(config)

            chunks = chunked(candidates, max_items) if candidates else ([[]] if failed_targets else [])
            title_base = notification_title(suffix="（部分成功）" if failed_targets else "")
            for index, chunk in enumerate(chunks, start=1):
                title = title_base if len(chunks) == 1 else f"{title_base} ({index}/{len(chunks)})"
                markdown = render_markdown(
                    chunk,
                    title,
                    index,
                    len(chunks),
                    failed_targets=failed_targets if index == 1 else [],
                )
                try:
                    result = send_dingtalk_markdown(config, state_dir, title, markdown, dry_run=dry_run)
                except Exception as exc:
                    status = "send_failed"
                    error = str(exc)
                    record_attempt(conn, run_id, dry_run, "failed", len(chunk), title, error=error)
                    break
                if result.returncode == 0:
                    record_attempt(conn, run_id, dry_run, "dry_run" if dry_run else "sent", len(chunk), title, dws_stdout=result.stdout, dws_stderr=result.stderr)
                    if not dry_run:
                        mark_notified(conn, chunk, now_text())
                        sent_items += len(chunk)
                else:
                    status = "send_failed"
                    error = f"exit={result.returncode}"
                    record_attempt(conn, run_id, dry_run, "failed", len(chunk), title, error=error, dws_stdout=result.stdout, dws_stderr=result.stderr)
                    break

            if status != "send_failed":
                if failed_targets:
                    status = "partial_notice_dry_run" if dry_run else "partial_notice_sent"
                elif not candidates:
                    status = "no_new_items"
                elif dry_run:
                    status = "dry_run"
    except SourceStaleError as exc:
        status = "source_stale"
        error = str(exc)
        source = str(exc.path)
        write_run_start(conn, run_id, started_at, source)
    except RunLockError as exc:
        status = "run_locked"
        error = str(exc)
        source = str(exc.path)
        write_run_start(conn, run_id, started_at, source)
    except PrimaryHostError as exc:
        status = "host_blocked"
        error = str(exc)
        if not source:
            source = "unknown"
        write_run_start(conn, run_id, started_at, source)
    except Exception as exc:
        status = "failed"
        error = str(exc)
        if not source:
            source = "unknown"
        write_run_start(conn, run_id, started_at, source)
    finally:
        try:
            total_items = 0
            try:
                total_items = len(items)  # type: ignore[name-defined]
            except Exception:
                total_items = 0
            write_run_end(conn, run_id, total_items, len(candidates), sent_items, status, error)
            for item in candidates:
                notify_status = "sent" if sent_items and item in candidates[:sent_items] else status
                item_rows.append(item.to_row(run_id, notify_status))
            run_row = {
                "run_id": run_id,
                "started_at": started_at,
                "ended_at": now_text(),
                "source": source,
                "total_items": total_items,
                "notify_candidates": len(candidates),
                "sent_items": sent_items,
                "status": status,
                "error": error,
            }
            artifact = update_run_artifacts(config, run_id, item_rows, run_row)
        finally:
            conn.close()
            release_run_lock(lock_path)

    return {
        "ok": status in {"success", "no_new_items", "dry_run", "partial_notice_sent", "partial_notice_dry_run"},
        "status": status,
        "run_id": run_id,
        "source": source,
        "total_items": total_items,
        "notify_candidates": len(candidates),
        "sent_items": sent_items,
        "artifact": str(artifact),
        "error": error,
        "coverage": coverage_summary,
        "failed_targets": failed_targets,
    }


def execute_doctor(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config_exists = config_path.is_file()
    config = load_config(config_path) if config_exists else default_config()
    state_dir = Path(args.state_dir or config_path.parent or DEFAULT_STATE_DIR)
    dingtalk = config.get("dingtalk", {})
    method = dingtalk_send_method(dingtalk)
    dws_call = Path(dingtalk.get("dws_call") or DEFAULT_DWS_CALL)
    db_path = Path(config.get("state", {}).get("db_path") or state_dir / "state.sqlite")
    checks = {
        "config_exists": config_exists,
        "config_path": str(config_path),
        "db_path": str(db_path),
        "dingtalk_method": method,
        "webhook_configured": bool(clean_text(dingtalk.get("webhook_url"))),
        "webhook_secret_configured": bool(clean_text(dingtalk.get("secret"))),
        "dws_call_exists": dws_call.is_file(),
        "robot_configured": bool(clean_text(dingtalk.get("robot_code"))),
        "group_configured": bool(clean_text(dingtalk.get("group_open_conversation_id"))),
    }
    checks.update(
        {
            "current_host": current_host_name(),
            "primary_host": primary_host_policy(config).get("primary_host", ""),
            "enforce_primary_host_for_send": primary_host_policy(config).get("enforce", True),
            "host_allowed_for_send": primary_host_policy(config).get("allowed", False),
        }
    )
    dws_status: dict[str, Any] = {}
    if method == "dws" and dws_call.is_file():
        try:
            result = run_dws_call(dws_call, ["auth", "status", "-f", "json"], state_dir, timeout=60)
            dws_status = {
                "exit_code": result.returncode,
                "stdout": result.stdout.strip(),
                "stderr": result.stderr.strip(),
            }
        except Exception as exc:
            dws_status = {"error": str(exc)}
    validation = validate_runtime_config(
        config_path,
        config,
        state_dir,
        require_send_ready=False,
        require_source_excel=should_require_source_excel_for_runtime(config),
        require_collector_ready=should_require_collector_for_runtime(config),
    )
    ok = validation["ok"]
    return {"ok": ok, "checks": checks, "validation": validation, "dws_auth_status": dws_status}


def execute_validate_config(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    state_dir = Path(args.state_dir or config_path.parent or DEFAULT_STATE_DIR)
    return validate_runtime_config(
        config_path,
        config,
        state_dir,
        require_send_ready=bool(args.require_send_ready),
        require_source_excel=should_require_source_excel_for_runtime(config),
        require_collector_ready=should_require_collector_for_runtime(config),
    )


def execute_merge_state(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    state_dir = Path(args.state_dir or config_path.parent or DEFAULT_STATE_DIR)
    source_db = Path(args.source_db).expanduser()
    target_db = Path(args.target_db or config.get("state", {}).get("db_path") or state_dir / "state.sqlite").expanduser()
    dry_run = bool(getattr(args, "dry_run", False))

    if not source_db.is_file():
        return {
            "ok": False,
            "status": "source_missing",
            "source_db": str(source_db),
            "target_db": str(target_db),
            "error": f"Source database not found: {source_db}",
            "issues": [issue("source_db_missing", f"找不到源状态库: {source_db}")],
        }
    if source_db.resolve() == target_db.resolve():
        return {
            "ok": False,
            "status": "invalid_args",
            "source_db": str(source_db),
            "target_db": str(target_db),
            "error": "Source and target database paths are the same.",
            "issues": [issue("state_db_same_path", "源状态库和目标状态库不能是同一路径")],
        }

    source_conn = sqlite3.connect(source_db)
    source_conn.row_factory = sqlite3.Row
    target_conn = connect_db(target_db)
    inserted = 0
    skipped_existing = 0
    by_store_inserted: dict[str, int] = {}

    try:
        if not table_exists(source_conn, "notified_items"):
            return {
                "ok": False,
                "status": "invalid_source",
                "source_db": str(source_db),
                "target_db": str(target_db),
                "error": "Source database does not contain notified_items table.",
                "issues": [issue("source_table_missing", "源状态库缺少 notified_items 表")],
            }
        source_rows = source_conn.execute(
            f"SELECT {', '.join(NOTIFIED_ITEM_COLUMNS)} FROM notified_items ORDER BY notified_at, dedupe_key"
        ).fetchall()
        target_before = int(target_conn.execute("SELECT COUNT(*) FROM notified_items").fetchone()[0])
        for row in source_rows:
            exists = target_conn.execute(
                "SELECT 1 FROM notified_items WHERE dedupe_key = ?",
                (row["dedupe_key"],),
            ).fetchone()
            if exists is not None:
                skipped_existing += 1
                continue
            inserted += 1
            store_name = clean_text(row["store"]) or "UNKNOWN"
            by_store_inserted[store_name] = by_store_inserted.get(store_name, 0) + 1
            if not dry_run:
                insert_notified_item_row(target_conn, row)
        if not dry_run:
            target_conn.commit()
        target_after = target_before + inserted if dry_run else int(target_conn.execute("SELECT COUNT(*) FROM notified_items").fetchone()[0])
    finally:
        source_conn.close()
        target_conn.close()

    return {
        "ok": True,
        "status": "dry_run" if dry_run else "success",
        "source_db": str(source_db),
        "target_db": str(target_db),
        "source_count": len(source_rows),
        "target_count_before": target_before,
        "inserted": inserted,
        "skipped_existing": skipped_existing,
        "target_count_after": target_after,
        "by_store_inserted": by_store_inserted,
    }


def execute_cdp_smoke(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    probe_args = build_cdp_args(config, args)
    return ziniao_cdp.probe_payload(probe_args)


def build_cdp_args(config: dict[str, Any], args: argparse.Namespace) -> argparse.Namespace:
    cdp_config = config.get("ziniao_cdp", {})
    port = int(args.port or cdp_config.get("port") or 0)
    port_start = int(args.port_start or cdp_config.get("port_start") or ziniao_cdp.DEFAULT_PORT_START)
    port_end = int(args.port_end or cdp_config.get("port_end") or ziniao_cdp.DEFAULT_PORT_END)
    url_contains = args.url_contains or cdp_config.get("url_contains") or ziniao_cdp.DEFAULT_URL_CONTAINS
    text_limit = int(args.text_limit or cdp_config.get("text_limit") or 800)
    duration_seconds = float(
        getattr(args, "duration_seconds", 0) or cdp_config.get("watch_duration_seconds") or ziniao_cdp.DEFAULT_WATCH_DURATION_SECONDS
    )
    interval_seconds = float(
        getattr(args, "interval_seconds", 0) or cdp_config.get("watch_interval_seconds") or ziniao_cdp.DEFAULT_WATCH_INTERVAL_SECONDS
    )
    reconnect_timeout_seconds = float(
        getattr(args, "reconnect_timeout_seconds", 0)
        or cdp_config.get("reconnect_timeout_seconds")
        or ziniao_cdp.DEFAULT_RECONNECT_TIMEOUT_SECONDS
    )
    return argparse.Namespace(
        port=port,
        port_start=port_start,
        port_end=port_end,
        url_contains=url_contains,
        text_limit=text_limit,
        include_body_sample=bool(getattr(args, "include_body_sample", False)),
        duration_seconds=duration_seconds,
        interval_seconds=interval_seconds,
        reconnect_timeout_seconds=reconnect_timeout_seconds,
        close_target=bool(getattr(args, "close_target", False)),
        wait_reopen=bool(getattr(args, "wait_reopen", False)),
        disconnect_wait_seconds=float(getattr(args, "disconnect_wait_seconds", 0) or 3),
    )


def execute_cdp_doctor(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    return ziniao_cdp.doctor_payload(build_cdp_args(config, args))


def execute_cdp_watch(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    return ziniao_cdp.watch_payload(build_cdp_args(config, args))


def execute_cdp_lifecycle_test(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    return ziniao_cdp.lifecycle_payload(build_cdp_args(config, args))


def impact_item_from_cdp_row(row: dict[str, Any], source_file: str) -> ImpactItem:
    return ImpactItem(
        store=clean_text(row.get("store")),
        site=clean_text(row.get("site")) or SITE_US,
        category=clean_text(row.get("category")),
        asin=clean_text(row.get("asin")),
        sku=clean_text(row.get("sku")),
        reason=clean_text(row.get("reason")),
        date=clean_text(row.get("date")),
        impacted_text=clean_text(row.get("impacted_text")),
        sales_risk=clean_text(row.get("sales_risk")),
        action=clean_text(row.get("action")),
        rating_impact=clean_text(row.get("rating_impact")),
        source_file=source_file,
    )


def execute_cdp_collect_current(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    cdp_args = build_cdp_args(config, args)
    start_date = clean_text(args.start_date)
    end_date = clean_text(args.end_date)
    if not start_date or not end_date:
        start_date, end_date = cdp_account_health.current_month_range()
    cdp_config = config.get("ziniao_cdp", {})
    page_size = int(args.page_size or cdp_config.get("collect_page_size") or cdp_account_health.DEFAULT_PAGE_SIZE)
    max_pages = int(args.max_pages or cdp_config.get("collect_max_pages") or cdp_account_health.DEFAULT_MAX_PAGES)
    categories = cdp_account_health.selected_categories(args.categories)
    collection = cdp_account_health.collect_current_account_health(
        cdp_args,
        start_date=start_date,
        end_date=end_date,
        categories=categories,
        page_size=page_size,
        max_pages=max_pages,
        store_override=args.store,
        site_override=args.site,
    )
    run_id = run_id_text()
    source_file = f"cdp:{collection.get('target', {}).get('url', '')}"
    items = [impact_item_from_cdp_row(row, source_file) for row in collection.get("rows", [])]
    item_rows = [item.to_row(run_id, "cdp_collected") for item in items]
    summary = summarize_items(items, expected_stores=[])
    summary["target_site"] = collection.get("site") or args.site or SITE_US
    result_dir = Path(args.output_dir or config.get("state", {}).get("result_dir") or DEFAULT_STATE_DIR / "runs")
    artifact = result_dir / f"account-health-cdp-current-{run_id}.xlsx"
    write_parse_xlsx(artifact, item_rows, summary)
    json_artifact = result_dir / f"account-health-cdp-current-{run_id}.json"
    json_artifact.parent.mkdir(parents=True, exist_ok=True)
    json_artifact.write_text(
        json.dumps(
            {
                "run_id": run_id,
                "start_date": start_date,
                "end_date": end_date,
                "store": collection.get("store"),
                "site": collection.get("site"),
                "row_count": len(items),
                "page_reports": collection.get("page_reports", []),
                "rows": collection.get("rows", []),
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    payload = {
        "ok": True,
        "status": "success",
        "run_id": run_id,
        "store": collection.get("store"),
        "site": collection.get("site"),
        "start_date": start_date,
        "end_date": end_date,
        "row_count": len(items),
        "artifact": str(artifact),
        "json_artifact": str(json_artifact),
        "page_reports": collection.get("page_reports", []),
        **summary,
    }
    if args.include_items:
        payload["items"] = [item.to_row(run_id, "cdp_collected") for item in items]
    else:
        payload["items_preview"] = [item.to_row(run_id, "cdp_collected") for item in items[:20]]
    return payload


def cdp_target_public_dict(target: dict[str, Any]) -> dict[str, Any]:
    return {
        "port": target.get("port", ""),
        "id": target.get("id", ""),
        "type": target.get("type", ""),
        "title": target.get("title", ""),
        "url": target.get("url", ""),
    }


def sanitized_target_result(result: dict[str, Any]) -> dict[str, Any]:
    sanitized = dict(result)
    sanitized["target"] = cdp_target_public_dict(dict(result.get("target") or {}))
    return sanitized


def cdp_target_result_rows(run_id: str, target_results: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for result in target_results:
        target = dict(result.get("target") or {})
        rows.append(
            {
                "run_id": run_id,
                "store": result.get("store", ""),
                "site": result.get("site", ""),
                "status": result.get("status", ""),
                "row_count": result.get("row_count", 0),
                "error": result.get("error", ""),
                "target_port": target.get("port", ""),
                "target_title": target.get("title", ""),
                "target_url": target.get("url", ""),
                "started_at": result.get("started_at", ""),
                "ended_at": result.get("ended_at", ""),
            }
        )
    return rows


def make_failed_target_result(
    store: str,
    site: str,
    status: str,
    error: str,
    *,
    started_at: str = "",
    ended_at: str = "",
) -> dict[str, Any]:
    return {
        "ok": False,
        "status": clean_text(status) or "failed",
        "store": clean_text(store),
        "site": clean_text(site) or SITE_US,
        "row_count": 0,
        "error": clean_text(error),
        "target": {"port": "", "id": "", "type": "", "title": "", "url": ""},
        "page_reports": [],
        "page_snapshot": {},
        "debugging_port": 0,
        "started_at": clean_text(started_at),
        "ended_at": clean_text(ended_at),
    }


def payload_target_results(payload: dict[str, Any]) -> list[dict[str, Any]]:
    return [sanitized_target_result(result) for result in payload.get("target_results", [])]


def payload_items(payload: dict[str, Any], source_file: str) -> list[ImpactItem]:
    rows = payload.get("items") or []
    items: list[ImpactItem] = []
    for row in rows:
        if isinstance(row, dict):
            items.extend(item_from_detail_row(row, source_file))
    return dedupe_items(items)


def collection_retry_store_names(payload: dict[str, Any]) -> list[str]:
    stores: list[str] = []
    for result in payload_target_results(payload):
        if result.get("ok"):
            continue
        stores.append(clean_text(result.get("store")))
    for store in payload.get("missing_stores", []):
        stores.append(clean_text(store))
    return ordered_unique_text(stores)


def collection_failed_targets(payload: dict[str, Any]) -> list[dict[str, Any]]:
    failed_targets = [result for result in payload_target_results(payload) if not result.get("ok")]
    target_site = clean_text(payload.get("target_site")) or SITE_US
    covered_keys = {
        normalize_for_key(clean_text(result.get("store")))
        for result in failed_targets
        if clean_text(result.get("store"))
    }
    for store in payload.get("missing_stores", []):
        store_name = clean_text(store)
        if not store_name:
            continue
        key = normalize_for_key(store_name)
        if key in covered_keys:
            continue
        failed_targets.append(
            make_failed_target_result(
                store_name,
                target_site,
                "missing",
                "店铺未完成采集",
            )
        )
        covered_keys.add(key)
    return dedupe_failed_targets(failed_targets)


def retry_failed_stores_enabled(config: dict[str, Any]) -> bool:
    return bool(config.get("production", {}).get("retry_failed_stores_enabled", True))


def production_retry_delay_seconds(config: dict[str, Any]) -> int:
    raw = config.get("production", {}).get("retry_delay_seconds", 600)
    try:
        value = int(raw)
    except Exception:
        value = 600
    return max(0, value)


def send_partial_with_failed_stores(config: dict[str, Any]) -> bool:
    return bool(config.get("production", {}).get("send_partial_with_failed_stores", True))


def build_retry_target_results(
    retry_payload: dict[str, Any],
    retry_stores: list[str],
) -> list[dict[str, Any]]:
    results = payload_target_results(retry_payload)
    existing_keys = {
        normalize_for_key(clean_text(result.get("store")))
        for result in results
        if clean_text(result.get("store"))
    }
    target_site = clean_text(retry_payload.get("target_site")) or SITE_US
    for store in retry_payload.get("missing_stores", []):
        store_name = clean_text(store)
        if not store_name:
            continue
        key = normalize_for_key(store_name)
        if key in existing_keys:
            continue
        results.append(
            make_failed_target_result(
                store_name,
                target_site,
                "missing",
                "店铺未完成采集",
            )
        )
        existing_keys.add(key)
    by_key = {
        normalize_for_key(clean_text(result.get("store"))): result
        for result in results
        if clean_text(result.get("store"))
    }
    fallback_error = clean_text(
        retry_payload.get("coverage_error")
        or retry_payload.get("store_list_error")
        or retry_payload.get("error")
        or "补采仍未成功"
    )
    ordered_results: list[dict[str, Any]] = []
    for store in retry_stores:
        key = normalize_for_key(store)
        result = by_key.get(key)
        if result:
            ordered_results.append(result)
        else:
            ordered_results.append(
                make_failed_target_result(
                    store,
                    target_site,
                    "retry_failed",
                    fallback_error,
                )
            )
    return ordered_results


def merge_collection_payloads(
    primary: dict[str, Any],
    retry: dict[str, Any],
    retry_stores: list[str],
    result_dir: Path,
) -> dict[str, Any]:
    run_id = run_id_text()
    backend = clean_text(primary.get("backend") or retry.get("backend") or "webdriver")
    target_site = clean_text(primary.get("target_site") or retry.get("target_site") or SITE_US) or SITE_US
    expected_stores = ordered_unique_text(
        list(primary.get("opened_stores", []))
        + list(primary.get("missing_stores", []))
        + list(retry.get("opened_stores", []))
        + list(retry.get("missing_stores", []))
    )
    retry_store_keys = {normalize_for_key(store) for store in retry_stores if clean_text(store)}
    primary_target_results = payload_target_results(primary)
    merged_target_results = [
        result
        for result in primary_target_results
        if normalize_for_key(clean_text(result.get("store"))) not in retry_store_keys
    ]
    merged_target_results.extend(build_retry_target_results(retry, retry_stores))

    primary_items = payload_items(primary, str(primary.get("artifact") or "primary-collection"))
    retry_items = payload_items(retry, str(retry.get("artifact") or "retry-collection"))
    merged_items = dedupe_items(
        [
            item
            for item in primary_items
            if normalize_for_key(item.store) not in retry_store_keys
        ]
        + retry_items
    )

    summary = cdp_open_summary(
        merged_items,
        expected_stores=expected_stores,
        target_results=merged_target_results,
        target_site=target_site,
        store_list_path=clean_text(primary.get("store_list_path") or retry.get("store_list_path")),
        store_list_error=clean_text(primary.get("store_list_error") or retry.get("store_list_error")),
    )
    target_failures = [result for result in merged_target_results if not result.get("ok")]
    if not merged_target_results:
        status = "no_targets"
    elif target_failures:
        status = "partial_failed"
    elif summary.get("missing_stores"):
        status = "partial"
    else:
        status = "success"
    ok = bool(merged_target_results) and not target_failures and not summary.get("missing_stores")

    artifact = result_dir / f"account-health-{backend}-stores-merged-{run_id}.xlsx"
    item_rows = [item.to_row(run_id, f"{backend}_collected") for item in merged_items]
    target_rows = cdp_target_result_rows(run_id, merged_target_results)
    write_collect_open_xlsx(artifact, item_rows, target_rows, summary)

    json_artifact = result_dir / f"account-health-{backend}-stores-merged-{run_id}.json"
    json_artifact.parent.mkdir(parents=True, exist_ok=True)
    json_payload = {
        "run_id": run_id,
        "status": status,
        "backend": backend,
        "start_date": primary.get("start_date") or retry.get("start_date") or "",
        "end_date": primary.get("end_date") or retry.get("end_date") or "",
        "target_site": target_site,
        "row_count": len(merged_items),
        "target_count": len(merged_target_results),
        "target_results": merged_target_results,
        "missing_stores": summary.get("missing_stores", []),
        "opened_stores": summary.get("opened_stores", []),
        "rows": item_rows,
    }
    json_artifact.write_text(json.dumps(json_payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    payload = {
        "ok": ok,
        "status": status,
        "backend": backend,
        "run_id": run_id,
        "start_date": primary.get("start_date") or retry.get("start_date") or "",
        "end_date": primary.get("end_date") or retry.get("end_date") or "",
        "target_site": target_site,
        "row_count": len(merged_items),
        "target_count": len(merged_target_results),
        "artifact": str(artifact),
        "json_artifact": str(json_artifact),
        "target_results": merged_target_results,
        **summary,
        "items": item_rows,
    }
    return payload


def load_cdp_expected_stores(
    config: dict[str, Any],
    args: argparse.Namespace,
    target_site: str,
) -> tuple[list[str], str, str]:
    if getattr(args, "skip_store_list", False):
        return [], "", ""
    source_config = config.get("source", {})
    store_list_path = Path(getattr(args, "store_list", "") or source_config.get("store_list_path") or "")
    if not store_list_path.is_file():
        return [], str(store_list_path), f"store_list_missing: {store_list_path}"
    return load_expected_stores(store_list_path, target_site), str(store_list_path), ""


def site_matches(value: str, target_site: str) -> bool:
    return normalize_for_key(value or target_site) == normalize_for_key(target_site)


def cdp_open_summary(
    items: list[ImpactItem],
    expected_stores: list[str],
    target_results: list[dict[str, Any]],
    target_site: str,
    store_list_path: str,
    store_list_error: str = "",
) -> dict[str, Any]:
    summary = summarize_items(items, expected_stores=[])
    opened_stores: list[str] = []
    opened_keys: set[str] = set()
    failed_targets = 0
    for result in target_results:
        if not result.get("ok"):
            failed_targets += 1
            continue
        store = clean_text(result.get("store"))
        site = clean_text(result.get("site")) or target_site
        if not store or not site_matches(site, target_site):
            continue
        key = normalize_for_key(store)
        if key in opened_keys:
            continue
        opened_keys.add(key)
        opened_stores.append(store)
    missing_stores = [
        store for store in expected_stores if normalize_for_key(store) not in opened_keys
    ]
    extra_stores = [
        store for store in opened_stores if normalize_for_key(store) not in {normalize_for_key(item) for item in expected_stores}
    ] if expected_stores else []
    summary.update(
        {
            "expected_store_count": len(expected_stores),
            "opened_store_count": len(opened_stores),
            "opened_stores": opened_stores,
            "missing_stores": missing_stores,
            "extra_stores": extra_stores,
            "coverage_ok": not expected_stores or not missing_stores,
            "target_site": target_site,
            "store_list_path": store_list_path,
            "store_list_error": store_list_error,
            "target_count": len(target_results),
            "failed_target_count": failed_targets,
        }
    )
    return summary


def execute_cdp_collect_open(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    cdp_args = build_cdp_args(config, args)
    start_date = clean_text(args.start_date)
    end_date = clean_text(args.end_date)
    if not start_date or not end_date:
        start_date, end_date = cdp_account_health.current_month_range()
    cdp_config = config.get("ziniao_cdp", {})
    source_config = config.get("source", {})
    target_site = clean_text(args.site) or clean_text(source_config.get("site")) or SITE_US
    page_size = int(args.page_size or cdp_config.get("collect_page_size") or cdp_account_health.DEFAULT_PAGE_SIZE)
    max_pages = int(args.max_pages or cdp_config.get("collect_max_pages") or cdp_account_health.DEFAULT_MAX_PAGES)
    categories = cdp_account_health.selected_categories(args.categories)
    expected_stores, store_list_path, store_list_error = load_cdp_expected_stores(config, args, target_site)

    collection = cdp_account_health.collect_open_account_health(
        cdp_args,
        start_date=start_date,
        end_date=end_date,
        categories=categories,
        page_size=page_size,
        max_pages=max_pages,
    )
    run_id = run_id_text()
    target_results = [sanitized_target_result(result) for result in collection.get("target_results", [])]
    source_file = "cdp-open"
    items = [
        impact_item_from_cdp_row(row, source_file)
        for row in collection.get("rows", [])
        if site_matches(clean_text(row.get("site")), target_site)
    ]
    item_rows = [item.to_row(run_id, "cdp_collected") for item in items]
    target_rows = cdp_target_result_rows(run_id, target_results)
    summary = cdp_open_summary(
        items,
        expected_stores=expected_stores,
        target_results=target_results,
        target_site=target_site,
        store_list_path=store_list_path,
        store_list_error=store_list_error,
    )

    coverage_error = ""
    if store_list_error and getattr(args, "require_all_open", False):
        coverage_error = store_list_error
    elif getattr(args, "require_all_open", False):
        coverage_error = coverage_failure_message(summary)

    target_failures = [result for result in target_results if not result.get("ok")]
    if not target_results:
        status = "no_targets"
    elif target_failures:
        status = "partial_failed"
    elif coverage_error:
        status = "coverage_failed"
    elif summary.get("missing_stores"):
        status = "partial"
    else:
        status = "success"
    ok = bool(target_results) and not target_failures and not coverage_error

    result_dir = Path(args.output_dir or config.get("state", {}).get("result_dir") or DEFAULT_STATE_DIR / "runs")
    artifact = result_dir / f"account-health-cdp-open-{run_id}.xlsx"
    write_collect_open_xlsx(artifact, item_rows, target_rows, summary)
    json_artifact = result_dir / f"account-health-cdp-open-{run_id}.json"
    json_artifact.parent.mkdir(parents=True, exist_ok=True)
    json_payload = {
        "run_id": run_id,
        "status": status,
        "start_date": start_date,
        "end_date": end_date,
        "target_site": target_site,
        "row_count": len(items),
        "target_count": len(target_results),
        "target_results": target_results,
        "missing_stores": summary.get("missing_stores", []),
        "opened_stores": summary.get("opened_stores", []),
        "coverage_error": coverage_error,
        "store_list_error": store_list_error,
        "rows": [item.to_row(run_id, "cdp_collected") for item in items],
    }
    json_artifact.write_text(json.dumps(json_payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    payload = {
        "ok": ok,
        "status": status,
        "run_id": run_id,
        "start_date": start_date,
        "end_date": end_date,
        "target_site": target_site,
        "row_count": len(items),
        "target_count": len(target_results),
        "artifact": str(artifact),
        "json_artifact": str(json_artifact),
        "coverage_error": coverage_error,
        "store_list_error": store_list_error,
        "target_results": target_results,
        **summary,
    }
    if args.include_items:
        payload["items"] = [item.to_row(run_id, "cdp_collected") for item in items]
    else:
        payload["items_preview"] = [item.to_row(run_id, "cdp_collected") for item in items[:20]]
    return payload


def execute_zclaw_collect_stores(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    start_date = clean_text(args.start_date)
    end_date = clean_text(args.end_date)
    if not start_date or not end_date:
        start_date, end_date = zclaw_account_health.current_month_range()
    source_config = config.get("source", {})
    target_site = clean_text(args.site) or clean_text(source_config.get("site")) or SITE_US
    categories = zclaw_account_health.selected_categories(args.categories)
    zclaw_args = zclaw_account_health.build_args(config, args)
    collection = zclaw_account_health.collect_visible_account_health(
        start_date=start_date,
        end_date=end_date,
        categories=categories,
        stores=args.stores,
        limit=int(args.limit or 0),
        page_size=zclaw_args["page_size"],
        max_pages=zclaw_args["max_pages"],
        close_after=not args.keep_open,
        open_timeout=zclaw_args["open_timeout"],
        nav_timeout=zclaw_args["nav_timeout"],
        exec_timeout=zclaw_args["exec_timeout"],
    )

    run_id = run_id_text()
    target_results = [sanitized_target_result(result) for result in collection.get("target_results", [])]
    source_file = "zclaw-stores"
    items = [
        impact_item_from_cdp_row(row, source_file)
        for row in collection.get("rows", [])
        if site_matches(clean_text(row.get("site")), target_site)
    ]
    item_rows = [item.to_row(run_id, "zclaw_collected") for item in items]
    target_rows = cdp_target_result_rows(run_id, target_results)

    expected_stores = [clean_text(store) for store in collection.get("selected_stores", []) if clean_text(store)]
    summary = cdp_open_summary(
        items,
        expected_stores=expected_stores,
        target_results=target_results,
        target_site=target_site,
        store_list_path="ziniao-cli store list",
        store_list_error="",
    )
    target_failures = [result for result in target_results if not result.get("ok")]
    if not target_results:
        status = "no_targets"
    elif target_failures:
        status = "partial_failed"
    elif summary.get("missing_stores"):
        status = "partial"
    else:
        status = "success"
    ok = bool(target_results) and not target_failures and not summary.get("missing_stores")

    result_dir = Path(args.output_dir or config.get("state", {}).get("result_dir") or DEFAULT_STATE_DIR / "runs")
    artifact = result_dir / f"account-health-zclaw-stores-{run_id}.xlsx"
    write_collect_open_xlsx(artifact, item_rows, target_rows, summary)
    json_artifact = result_dir / f"account-health-zclaw-stores-{run_id}.json"
    json_artifact.parent.mkdir(parents=True, exist_ok=True)
    json_payload = {
        "run_id": run_id,
        "status": status,
        "start_date": start_date,
        "end_date": end_date,
        "target_site": target_site,
        "row_count": len(items),
        "target_count": len(target_results),
        "target_results": target_results,
        "missing_stores": summary.get("missing_stores", []),
        "opened_stores": summary.get("opened_stores", []),
        "rows": [item.to_row(run_id, "zclaw_collected") for item in items],
    }
    json_artifact.write_text(json.dumps(json_payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    payload = {
        "ok": ok,
        "status": status,
        "run_id": run_id,
        "start_date": start_date,
        "end_date": end_date,
        "target_site": target_site,
        "row_count": len(items),
        "target_count": len(target_results),
        "artifact": str(artifact),
        "json_artifact": str(json_artifact),
        "target_results": target_results,
        **summary,
    }
    if args.include_items:
        payload["items"] = [item.to_row(run_id, "zclaw_collected") for item in items]
    else:
        payload["items_preview"] = [item.to_row(run_id, "zclaw_collected") for item in items[:20]]
    return payload


def execute_webdriver_collect_stores(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    start_date = clean_text(args.start_date)
    end_date = clean_text(args.end_date)
    if not start_date or not end_date:
        start_date, end_date = ziniao_webdriver.current_month_range()
    source_config = config.get("source", {})
    target_site = clean_text(args.site) or clean_text(source_config.get("site")) or SITE_US
    categories = ziniao_webdriver.selected_categories(args.categories)
    webdriver_args = ziniao_webdriver.build_args(config, args)
    try:
        collection = ziniao_webdriver.collect_visible_account_health(
            client_path=webdriver_args["client_path"],
            webdriver_port=webdriver_args["port"],
            credentials=webdriver_args["credentials"],
            start_date=start_date,
            end_date=end_date,
            categories=categories,
            stores=args.stores,
            limit=int(args.limit or 0),
            page_size=webdriver_args["page_size"],
            max_pages=webdriver_args["max_pages"],
            close_after=not args.keep_open,
            startup_timeout=webdriver_args["startup_timeout"],
            update_core_timeout=webdriver_args["update_core_timeout"],
            update_core_poll_seconds=webdriver_args["update_core_poll_seconds"],
            request_timeout=webdriver_args["request_timeout"],
            browser_ready_timeout=webdriver_args["browser_ready_timeout"],
            navigation_wait_seconds=webdriver_args["navigation_wait_seconds"],
        )
    except Exception as exc:
        return {
            "ok": False,
            "status": "collect_failed",
            "backend": "webdriver",
            "run_id": run_id_text(),
            "start_date": start_date,
            "end_date": end_date,
            "target_site": target_site,
            "row_count": 0,
            "target_count": 0,
            "artifact": "",
            "json_artifact": "",
            "target_results": [],
            "items_preview": [],
            "error": clean_text(exc),
        }

    run_id = run_id_text()
    target_results = [sanitized_target_result(result) for result in collection.get("target_results", [])]
    source_file = "webdriver-stores"
    items = [
        impact_item_from_cdp_row(row, source_file)
        for row in collection.get("rows", [])
        if site_matches(clean_text(row.get("site")), target_site)
    ]
    item_rows = [item.to_row(run_id, "webdriver_collected") for item in items]
    target_rows = cdp_target_result_rows(run_id, target_results)

    expected_stores = [clean_text(store) for store in collection.get("selected_stores", []) if clean_text(store)]
    summary = cdp_open_summary(
        items,
        expected_stores=expected_stores,
        target_results=target_results,
        target_site=target_site,
        store_list_path="ziniao-webdriver getBrowserList",
        store_list_error="",
    )
    target_failures = [result for result in target_results if not result.get("ok")]
    if not target_results:
        status = "no_targets"
    elif target_failures:
        status = "partial_failed"
    elif summary.get("missing_stores"):
        status = "partial"
    else:
        status = "success"
    ok = bool(target_results) and not target_failures and not summary.get("missing_stores")

    result_dir = Path(args.output_dir or config.get("state", {}).get("result_dir") or DEFAULT_STATE_DIR / "runs")
    artifact = result_dir / f"account-health-webdriver-stores-{run_id}.xlsx"
    write_collect_open_xlsx(artifact, item_rows, target_rows, summary)
    json_artifact = result_dir / f"account-health-webdriver-stores-{run_id}.json"
    json_artifact.parent.mkdir(parents=True, exist_ok=True)
    json_payload = {
        "run_id": run_id,
        "status": status,
        "backend": "webdriver",
        "start_date": start_date,
        "end_date": end_date,
        "target_site": target_site,
        "row_count": len(items),
        "target_count": len(target_results),
        "target_results": target_results,
        "missing_stores": summary.get("missing_stores", []),
        "opened_stores": summary.get("opened_stores", []),
        "rows": [item.to_row(run_id, "webdriver_collected") for item in items],
    }
    json_artifact.write_text(json.dumps(json_payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    payload = {
        "ok": ok,
        "status": status,
        "backend": "webdriver",
        "run_id": run_id,
        "start_date": start_date,
        "end_date": end_date,
        "target_site": target_site,
        "row_count": len(items),
        "target_count": len(target_results),
        "artifact": str(artifact),
        "json_artifact": str(json_artifact),
        "target_results": target_results,
        **summary,
    }
    if args.include_items:
        payload["items"] = [item.to_row(run_id, "webdriver_collected") for item in items]
    else:
        payload["items_preview"] = [item.to_row(run_id, "webdriver_collected") for item in items[:20]]
    return payload


def execute_production_run(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(getattr(args, "config", str(DEFAULT_CONFIG_PATH)))
    config = load_config(config_path) if config_path.is_file() else default_config()
    backend = collector_backend_name(config, getattr(args, "backend", ""))
    result_dir = Path(
        getattr(args, "output_dir", "")
        or config.get("state", {}).get("result_dir")
        or DEFAULT_STATE_DIR / "runs"
    )
    collect_args = argparse.Namespace(
        config=str(config_path),
        state_dir=getattr(args, "state_dir", ""),
        start_date=getattr(args, "start_date", ""),
        end_date=getattr(args, "end_date", ""),
        categories=getattr(args, "categories", "all"),
        stores=getattr(args, "stores", ""),
        limit=int(getattr(args, "limit", 0) or 0),
        site=getattr(args, "site", ""),
        page_size=int(getattr(args, "page_size", 0) or 0),
        max_pages=int(getattr(args, "max_pages", 0) or 0),
        open_timeout=int(getattr(args, "open_timeout", 0) or 0),
        nav_timeout=int(getattr(args, "nav_timeout", 0) or 0),
        exec_timeout=int(getattr(args, "exec_timeout", 0) or 0),
        keep_open=bool(getattr(args, "keep_open", False)),
        output_dir=getattr(args, "output_dir", ""),
        include_items=True,
    )
    collector_fn: Callable[[argparse.Namespace], dict[str, Any]]
    if backend == "webdriver":
        collector_fn = execute_webdriver_collect_stores
    elif backend == "zclaw":
        collector_fn = execute_zclaw_collect_stores
    else:
        return {
            "ok": False,
            "status": "collect_failed",
            "collection_status": "invalid_backend",
            "notification_status": "skipped",
            "backend": backend,
            "collection_run_id": "",
            "start_date": "",
            "end_date": "",
            "target_site": "",
            "target_count": 0,
            "row_count": 0,
            "source_excel": "",
            "collect_artifact": "",
            "collect_json_artifact": "",
            "notify_artifact": "",
            "total_items": 0,
            "notify_candidates": 0,
            "sent_items": 0,
            "error": f"Unsupported collector backend: {backend}",
            "collection": {},
        }
    initial_collection = collector_fn(collect_args)
    final_collection = initial_collection
    retry_collection: dict[str, Any] = {}
    retry_stores = collection_retry_store_names(initial_collection)
    retry_info = {
        "attempted": False,
        "retry_delay_seconds": 0,
        "retry_stores": retry_stores,
        "retry_status": "skipped",
        "remaining_failed_stores": retry_stores,
    }
    if retry_failed_stores_enabled(config) and retry_stores:
        retry_info["attempted"] = True
        retry_delay_seconds = production_retry_delay_seconds(config)
        retry_info["retry_delay_seconds"] = retry_delay_seconds
        if retry_delay_seconds > 0:
            time.sleep(retry_delay_seconds)
        retry_args = argparse.Namespace(**vars(collect_args))
        retry_args.stores = ",".join(retry_stores)
        retry_args.limit = 0
        retry_collection = collector_fn(retry_args)
        retry_info["retry_status"] = clean_text(retry_collection.get("status")) or "unknown"
        final_collection = merge_collection_payloads(initial_collection, retry_collection, retry_stores, result_dir)
        retry_info["remaining_failed_stores"] = collection_retry_store_names(final_collection)

    source_excel = clean_text(final_collection.get("artifact"))
    final_failed_targets = collection_failed_targets(final_collection)
    allow_partial_send = bool(source_excel) and bool(final_failed_targets) and send_partial_with_failed_stores(config)
    if (not final_collection.get("ok") and not allow_partial_send) or not source_excel:
        return {
            "ok": False,
            "status": "collect_failed",
            "collection_status": final_collection.get("status", "unknown"),
            "notification_status": "skipped",
            "backend": backend,
            "collection_run_id": final_collection.get("run_id", ""),
            "start_date": final_collection.get("start_date", ""),
            "end_date": final_collection.get("end_date", ""),
            "target_site": final_collection.get("target_site", ""),
            "target_count": final_collection.get("target_count", 0),
            "row_count": final_collection.get("row_count", 0),
            "source_excel": source_excel,
            "collect_artifact": source_excel,
            "collect_json_artifact": final_collection.get("json_artifact", ""),
            "notify_artifact": "",
            "total_items": 0,
            "notify_candidates": 0,
            "sent_items": 0,
            "error": final_collection.get("coverage_error") or final_collection.get("error") or "Collection did not cover all target stores.",
            "collection": final_collection,
            "initial_collection": initial_collection,
            "retry_collection": retry_collection,
            "retry": retry_info,
        }

    notify_args = argparse.Namespace(
        config=getattr(args, "config", str(DEFAULT_CONFIG_PATH)),
        state_dir=getattr(args, "state_dir", ""),
        source_type="excel",
        source_excel=source_excel,
        source_dir="",
        site=final_collection.get("target_site") or getattr(args, "site", "") or SITE_US,
        store_list="",
        require_all_stores=not allow_partial_send,
        skip_store_coverage=allow_partial_send,
        dry_run=bool(getattr(args, "dry_run", False)),
        send=bool(getattr(args, "send", False)),
        failed_targets=final_failed_targets if allow_partial_send else [],
    )
    notification = execute_run(notify_args)
    notification_status = clean_text(notification.get("status")) or "unknown"
    ok = bool(notification.get("ok"))
    if ok and allow_partial_send:
        status = "partial_success"
    else:
        status = notification_status if ok else f"notify_{notification_status}"
    return {
        "ok": ok,
        "status": status,
        "backend": backend,
        "collection_status": final_collection.get("status", "unknown"),
        "notification_status": notification_status,
        "collection_run_id": final_collection.get("run_id", ""),
        "notification_run_id": notification.get("run_id", ""),
        "start_date": final_collection.get("start_date", ""),
        "end_date": final_collection.get("end_date", ""),
        "target_site": final_collection.get("target_site", ""),
        "target_count": final_collection.get("target_count", 0),
        "row_count": final_collection.get("row_count", 0),
        "source_excel": source_excel,
        "collect_artifact": source_excel,
        "collect_json_artifact": final_collection.get("json_artifact", ""),
        "notify_artifact": notification.get("artifact", ""),
        "total_items": notification.get("total_items", 0),
        "notify_candidates": notification.get("notify_candidates", 0),
        "sent_items": notification.get("sent_items", 0),
        "error": notification.get("error", ""),
        "coverage": notification.get("coverage", {}),
        "collection": final_collection,
        "initial_collection": initial_collection,
        "retry_collection": retry_collection,
        "retry": retry_info,
        "notification": notification,
    }


def execute_send_test(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path)
    state_dir = Path(args.state_dir or config_path.parent or DEFAULT_STATE_DIR)
    dry_run = not args.send
    if not dry_run:
        validation = validate_runtime_config(
            config_path,
            config,
            state_dir,
            require_send_ready=True,
            require_source_excel=False,
            require_collector_ready=False,
        )
        if not validation["ok"]:
            return {
                "ok": False,
                "status": "preflight_failed",
                "dry_run": False,
                "issues": validation["issues"],
            }
    title = f"{config.get('dingtalk', {}).get('title_prefix', '亚马逊账号状况异常')}测试消息"
    markdown = "### 亚马逊账号状况异常测试消息\n\n- 这是一条钉钉机器人连通性测试。\n- 如果你看到这条消息, 说明机器人发群链路可用。\n"
    result = send_dingtalk_markdown(config, state_dir, title, markdown, dry_run=dry_run)
    return {
        "ok": result.returncode == 0,
        "dry_run": dry_run,
        "exit_code": result.returncode,
        "stdout": result.stdout,
        "stderr": result.stderr,
    }


def execute_install_schedule(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH).resolve()
    config = load_config(config_path)
    state_dir = Path(args.state_dir or config_path.parent or DEFAULT_STATE_DIR)
    schedule = config.get("schedule", {})
    task_name = args.task_name or schedule.get("task_name") or DEFAULT_TASK_NAME
    schedule_command = clean_text(getattr(args, "schedule_command", "") or schedule.get("command") or "production-run")
    if schedule_command not in {"production-run", "run"}:
        return {
            "ok": False,
            "status": "preflight_failed",
            "dry_run": bool(args.dry_run),
            "issues": [issue("schedule_command_invalid", "schedule.command 必须是 production-run 或 run")],
        }
    interval_raw = args.interval_hours or schedule.get("interval_hours") or 6
    try:
        interval = int(interval_raw)
        if interval <= 0:
            raise ValueError
    except Exception:
        return {
            "ok": False,
            "status": "preflight_failed",
            "dry_run": bool(args.dry_run),
            "issues": [issue("interval_hours_invalid", "interval_hours 必须是正整数")],
        }
    python_exe = Path(args.python or sys.executable).resolve()
    script = Path(__file__).resolve()
    command = f'"{python_exe}" "{script}" {schedule_command} --config "{config_path}"'
    schtasks_args = [
        "schtasks",
        "/Create",
        "/TN",
        task_name,
        "/SC",
        "HOURLY",
        "/MO",
        str(interval),
        "/TR",
        command,
        "/F",
    ]
    validation = validate_runtime_config(
        config_path,
        config,
        state_dir,
        require_send_ready=True,
        require_source_excel=schedule_command == "run",
        require_collector_ready=schedule_command == "production-run",
    )
    if not validation["ok"] and not getattr(args, "skip_preflight", False):
        return {
            "ok": False,
            "status": "preflight_failed",
            "dry_run": bool(args.dry_run),
            "command": schtasks_args,
            "issues": validation["issues"],
        }
    if args.dry_run:
        return {"ok": True, "status": "dry_run", "dry_run": True, "command": schtasks_args, "issues": validation["issues"]}
    result = subprocess.run(schtasks_args, text=True, encoding="utf-8", errors="replace", capture_output=True)
    return {
        "ok": result.returncode == 0,
        "dry_run": False,
        "exit_code": result.returncode,
        "stdout": result.stdout,
        "stderr": result.stderr,
        "command": schtasks_args,
    }


def execute_install_cdp_daemon(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path) if config_path.is_file() else default_config()
    cdp_config = config.get("ziniao_cdp", {})
    task_name = args.task_name or cdp_config.get("daemon_task_name") or DEFAULT_CDP_DAEMON_TASK_NAME
    python_exe = Path(args.python or sys.executable).resolve()
    daemon_script = Path(__file__).resolve().parent / "tools" / "ziniao_cdp_daemon.py"
    log_path = Path(args.log_file or cdp_config.get("daemon_log_path") or DEFAULT_STATE_DIR / "ziniao-cdp-daemon.log")
    port_start = int(args.port_start or cdp_config.get("port_start") or ziniao_cdp.DEFAULT_PORT_START)
    port_end = int(args.port_end or cdp_config.get("port_end") or ziniao_cdp.DEFAULT_PORT_END)
    if not daemon_script.is_file():
        return {"ok": False, "status": "preflight_failed", "issues": [issue("daemon_missing", str(daemon_script))]}
    command = (
        f'"{python_exe}" "{daemon_script}" '
        f'--port-start {port_start} --port-end {port_end} --log-file "{log_path}"'
    )
    schtasks_args = [
        "schtasks",
        "/Create",
        "/TN",
        task_name,
        "/SC",
        "ONLOGON",
        "/TR",
        command,
        "/F",
    ]
    if args.dry_run:
        return {"ok": True, "status": "dry_run", "dry_run": True, "command": schtasks_args}
    result = subprocess.run(schtasks_args, text=True, encoding="utf-8", errors="replace", capture_output=True)
    return {
        "ok": result.returncode == 0,
        "status": "installed" if result.returncode == 0 else "install_failed",
        "dry_run": False,
        "exit_code": result.returncode,
        "stdout": result.stdout,
        "stderr": result.stderr,
        "command": schtasks_args,
    }


def execute_self_test(args: argparse.Namespace) -> dict[str, Any]:
    state_dir = Path(args.state_dir or DEFAULT_STATE_DIR / "self-test")
    if state_dir.exists():
        resolved = state_dir.resolve()
        allowed_root = DEFAULT_STATE_DIR.resolve()
        if allowed_root not in resolved.parents and resolved != allowed_root:
            raise ValueError(f"拒绝清理非默认自测目录: {state_dir}")
        shutil.rmtree(state_dir)
    config_path = state_dir / "config.json"
    config = default_config()
    config["source"]["type"] = "sample"
    config["dingtalk"]["send_enabled"] = False
    config["state"]["db_path"] = str(state_dir / "state.sqlite")
    config["state"]["result_dir"] = str(state_dir / "runs")
    config_path.parent.mkdir(parents=True, exist_ok=True)
    config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    first_args = argparse.Namespace(
        config=str(config_path),
        state_dir=str(state_dir),
        source_type="sample",
        source_excel=None,
        source_dir=None,
        site=SITE_US,
        dry_run=True,
        send=False,
    )
    first = execute_run(first_args)

    conn = connect_db(Path(config["state"]["db_path"]))
    try:
        sample_items = load_sample_items()
        mark_notified(conn, sample_items, now_text())
    finally:
        conn.close()

    second = execute_run(first_args)

    changed = config.copy()
    changed_path = state_dir / "changed-config.json"
    changed_path.write_text(json.dumps(changed, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    changed_items = [
        ImpactItem(
            store="BYF",
            site=SITE_US,
            category="食品和商品安全问题",
            asin="B0FPKWYF49",
            sku="BF-9901-BN",
            reason="安全饮用水法案: 食品和商品安全问题",
            date="2026-01-30",
            impacted_text="Bathroom Faucet Brushed Nickel ASIN: B0FPKWYF49 SKU: BF-9901-BN",
            sales_risk="过去 12 个月无销量",
            action="商品存在销售风险",
            rating_impact="无影响",
            source_file="sample-changed",
        )
    ]
    conn = connect_db(Path(config["state"]["db_path"]))
    try:
        changed_candidates = select_notify_candidates(conn, changed_items)
    finally:
        conn.close()

    return {
        "ok": first["notify_candidates"] == 1 and second["notify_candidates"] == 0 and len(changed_candidates) == 1,
        "state_dir": str(state_dir),
        "first_run_candidates": first["notify_candidates"],
        "second_run_candidates": second["notify_candidates"],
        "changed_candidates": len(changed_candidates),
        "first_artifact": first["artifact"],
        "second_artifact": second["artifact"],
    }


def print_result(payload: dict[str, Any], as_json: bool) -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")
    if as_json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        for key, value in payload.items():
            print(f"{key}: {value}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Amazon account health ASIN/SKU DingTalk notifier.")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG_PATH), help="配置文件路径")
    parser.add_argument("--state-dir", default="", help="运行状态目录")
    parser.add_argument("--json", action="store_true", help="输出 JSON")
    sub = parser.add_subparsers(dest="command", required=True)

    def add_common(child: argparse.ArgumentParser) -> None:
        child.add_argument("--config", default=argparse.SUPPRESS, help="配置文件路径")
        child.add_argument("--state-dir", default=argparse.SUPPRESS, help="运行状态目录")
        child.add_argument("--json", action="store_true", default=argparse.SUPPRESS, help="输出 JSON")

    init_config = sub.add_parser("init-config", help="生成默认配置文件")
    add_common(init_config)
    init_config.add_argument("--force", action="store_true", help="覆盖已有配置")

    doctor = sub.add_parser("doctor", help="检查配置和 DWS 状态")
    add_common(doctor)
    doctor.set_defaults(func=execute_doctor)

    validate_config = sub.add_parser("validate-config", help="检查本地配置是否可用于运行或定时通知")
    add_common(validate_config)
    validate_config.add_argument("--require-send-ready", action="store_true", help="要求机器人和群配置已满足真实通知")
    validate_config.set_defaults(func=execute_validate_config)

    merge_state = sub.add_parser("merge-state", help="合并旧机器或旧目录中的已通知历史库")
    add_common(merge_state)
    merge_state.add_argument("--source-db", required=True, help="源状态库路径")
    merge_state.add_argument("--target-db", default="", help="目标状态库路径, 默认使用当前配置 state.db_path")
    merge_state.add_argument("--dry-run", action="store_true", help="只预览将导入多少条历史记录")
    merge_state.set_defaults(func=execute_merge_state)

    run = sub.add_parser("run", help="执行采集、去重和通知")
    add_common(run)
    run.add_argument("--source-type", choices=["excel", "excel_latest", "sample", "bridge"], default="")
    run.add_argument("--source-excel", default="", help="指定账号状况结果 Excel")
    run.add_argument("--source-dir", default="", help="指定结果 Excel 目录")
    run.add_argument("--site", default="", help="目标站点")
    run.add_argument("--store-list", default="", help="用于校验店铺覆盖率的店铺清单")
    run.add_argument("--require-all-stores", action="store_true", help="强制要求店铺清单全部覆盖")
    run.add_argument("--skip-store-coverage", action="store_true", help="跳过店铺覆盖校验, 仅用于 sample 或排障")
    run.add_argument("--dry-run", action="store_true", help="不真实发送钉钉消息")
    run.add_argument("--send", action="store_true", help="允许真实发送钉钉消息")
    run.set_defaults(func=execute_run)

    parse = sub.add_parser("parse", help="只解析 Excel 并生成全店铺汇总报告, 不发送钉钉")
    add_common(parse)
    parse.add_argument("--source-type", choices=["excel", "excel_latest", "sample", "bridge"], default="")
    parse.add_argument("--source-excel", default="", help="指定账号状况结果 Excel")
    parse.add_argument("--source-dir", default="", help="指定结果 Excel 目录")
    parse.add_argument("--site", default="", help="目标站点")
    parse.add_argument("--store-list", default="", help="用于校验店铺覆盖率的店铺清单")
    parse.add_argument("--output-dir", default="", help="解析报告输出目录")
    parse.add_argument("--require-all-stores", action="store_true", help="店铺覆盖不完整时返回失败")
    parse.set_defaults(func=execute_parse)

    cdp_smoke = sub.add_parser("cdp-smoke", help="检查紫鸟 CDP 端口和 Seller Central 页面只读访问")
    add_common(cdp_smoke)
    cdp_smoke.add_argument("--port", type=int, default=0, help="指定 CDP 端口, 不指定则扫描")
    cdp_smoke.add_argument("--port-start", type=int, default=0, help="扫描起始端口")
    cdp_smoke.add_argument("--port-end", type=int, default=0, help="扫描结束端口")
    cdp_smoke.add_argument("--url-contains", default="", help="目标页面 URL 片段")
    cdp_smoke.add_argument("--text-limit", type=int, default=0, help="页面正文摘要长度")
    cdp_smoke.set_defaults(func=execute_cdp_smoke)

    cdp_doctor = sub.add_parser("cdp-doctor", help="诊断紫鸟 CDP、daemon、端口、target 和页面可读性")
    add_common(cdp_doctor)
    cdp_doctor.add_argument("--port", type=int, default=0, help="指定 CDP 端口, 不指定则扫描")
    cdp_doctor.add_argument("--port-start", type=int, default=0, help="扫描起始端口")
    cdp_doctor.add_argument("--port-end", type=int, default=0, help="扫描结束端口")
    cdp_doctor.add_argument("--url-contains", default="", help="目标页面 URL 片段")
    cdp_doctor.add_argument("--text-limit", type=int, default=0, help="页面正文摘要长度")
    cdp_doctor.add_argument("--include-body-sample", action="store_true", help="输出页面正文摘要")
    cdp_doctor.set_defaults(func=execute_cdp_doctor)

    cdp_watch = sub.add_parser("cdp-watch", help="连续检查 CDP 心跳并记录断开/重连状态")
    add_common(cdp_watch)
    cdp_watch.add_argument("--port", type=int, default=0, help="指定 CDP 端口, 不指定则扫描")
    cdp_watch.add_argument("--port-start", type=int, default=0, help="扫描起始端口")
    cdp_watch.add_argument("--port-end", type=int, default=0, help="扫描结束端口")
    cdp_watch.add_argument("--url-contains", default="", help="目标页面 URL 片段")
    cdp_watch.add_argument("--text-limit", type=int, default=0, help="页面正文摘要长度")
    cdp_watch.add_argument("--duration-seconds", type=float, default=0, help="观察时长")
    cdp_watch.add_argument("--interval-seconds", type=float, default=0, help="心跳间隔")
    cdp_watch.set_defaults(func=execute_cdp_watch)

    cdp_lifecycle = sub.add_parser("cdp-lifecycle-test", help="验证当前 target 可读、断开和重连边界")
    add_common(cdp_lifecycle)
    cdp_lifecycle.add_argument("--port", type=int, default=0, help="指定 CDP 端口, 不指定则扫描")
    cdp_lifecycle.add_argument("--port-start", type=int, default=0, help="扫描起始端口")
    cdp_lifecycle.add_argument("--port-end", type=int, default=0, help="扫描结束端口")
    cdp_lifecycle.add_argument("--url-contains", default="", help="目标页面 URL 片段")
    cdp_lifecycle.add_argument("--text-limit", type=int, default=0, help="页面正文摘要长度")
    cdp_lifecycle.add_argument("--close-target", action="store_true", help="关闭当前 Seller Central target")
    cdp_lifecycle.add_argument("--wait-reopen", action="store_true", help="关闭后等待重开窗口并重连")
    cdp_lifecycle.add_argument("--disconnect-wait-seconds", type=float, default=0, help="关闭后等待断开秒数")
    cdp_lifecycle.add_argument("--reconnect-timeout-seconds", type=float, default=0, help="重连等待上限")
    cdp_lifecycle.set_defaults(func=execute_cdp_lifecycle_test)

    cdp_collect = sub.add_parser("cdp-collect-current", help="Collect current-month account health issues through Ziniao CDP")
    add_common(cdp_collect)
    cdp_collect.add_argument("--port", type=int, default=0, help="CDP port")
    cdp_collect.add_argument("--port-start", type=int, default=0, help="CDP scan start port")
    cdp_collect.add_argument("--port-end", type=int, default=0, help="CDP scan end port")
    cdp_collect.add_argument("--url-contains", default="", help="target page URL fragment")
    cdp_collect.add_argument("--text-limit", type=int, default=0, help="page text limit")
    cdp_collect.add_argument("--start-date", default="", help="start date YYYY-MM-DD, defaults to current month start")
    cdp_collect.add_argument("--end-date", default="", help="end date YYYY-MM-DD, defaults to today")
    cdp_collect.add_argument("--categories", default="all", help="all or comma-separated policy keys")
    cdp_collect.add_argument("--page-size", type=int, default=0, help="API page size")
    cdp_collect.add_argument("--max-pages", type=int, default=0, help="maximum pages per category")
    cdp_collect.add_argument("--store", default="", help="override store name")
    cdp_collect.add_argument("--site", default="", help="override site")
    cdp_collect.add_argument("--output-dir", default="", help="output directory")
    cdp_collect.add_argument("--include-items", action="store_true", help="include full item rows in JSON output")
    cdp_collect.set_defaults(func=execute_cdp_collect_current)

    cdp_collect_open = sub.add_parser("cdp-collect-open", help="Collect account health issues from all open Ziniao Seller Central CDP targets")
    add_common(cdp_collect_open)
    cdp_collect_open.add_argument("--port", type=int, default=0, help="CDP port")
    cdp_collect_open.add_argument("--port-start", type=int, default=0, help="CDP scan start port")
    cdp_collect_open.add_argument("--port-end", type=int, default=0, help="CDP scan end port")
    cdp_collect_open.add_argument("--url-contains", default="", help="target page URL fragment")
    cdp_collect_open.add_argument("--text-limit", type=int, default=0, help="page text limit")
    cdp_collect_open.add_argument("--start-date", default="", help="start date YYYY-MM-DD, defaults to current month start")
    cdp_collect_open.add_argument("--end-date", default="", help="end date YYYY-MM-DD, defaults to today")
    cdp_collect_open.add_argument("--categories", default="all", help="all or comma-separated policy keys")
    cdp_collect_open.add_argument("--page-size", type=int, default=0, help="API page size")
    cdp_collect_open.add_argument("--max-pages", type=int, default=0, help="maximum pages per category")
    cdp_collect_open.add_argument("--site", default="", help="target site")
    cdp_collect_open.add_argument("--store-list", default="", help="store list used for US-site coverage")
    cdp_collect_open.add_argument("--skip-store-list", action="store_true", help="skip store-list coverage")
    cdp_collect_open.add_argument("--require-all-open", action="store_true", help="fail when expected US stores are not open")
    cdp_collect_open.add_argument("--output-dir", default="", help="output directory")
    cdp_collect_open.add_argument("--include-items", action="store_true", help="include full item rows in JSON output")
    cdp_collect_open.set_defaults(func=execute_cdp_collect_open)

    zclaw_collect = sub.add_parser("zclaw-collect-stores", help="Collect account health issues through Ziniao ZClaw store/page commands")
    add_common(zclaw_collect)
    zclaw_collect.add_argument("--start-date", default="", help="start date YYYY-MM-DD, defaults to current month start")
    zclaw_collect.add_argument("--end-date", default="", help="end date YYYY-MM-DD, defaults to today")
    zclaw_collect.add_argument("--categories", default="all", help="all or comma-separated policy keys")
    zclaw_collect.add_argument("--stores", default="", help="comma-separated store names or store IDs; defaults to all visible Amazon US stores")
    zclaw_collect.add_argument("--limit", type=int, default=0, help="limit selected stores for smoke testing")
    zclaw_collect.add_argument("--site", default="", help="target site")
    zclaw_collect.add_argument("--page-size", type=int, default=0, help="API page size")
    zclaw_collect.add_argument("--max-pages", type=int, default=0, help="maximum pages per category")
    zclaw_collect.add_argument("--open-timeout", type=int, default=0, help="store open timeout seconds")
    zclaw_collect.add_argument("--nav-timeout", type=int, default=0, help="page navigation timeout seconds")
    zclaw_collect.add_argument("--exec-timeout", type=int, default=0, help="page script timeout seconds")
    zclaw_collect.add_argument("--keep-open", action="store_true", help="do not close store windows after collection")
    zclaw_collect.add_argument("--output-dir", default="", help="output directory")
    zclaw_collect.add_argument("--include-items", action="store_true", help="include full item rows in JSON output")
    zclaw_collect.set_defaults(func=execute_zclaw_collect_stores)

    webdriver_collect = sub.add_parser("webdriver-collect-stores", help="Collect account health issues through official Ziniao WebDriver + CDP")
    add_common(webdriver_collect)
    webdriver_collect.add_argument("--start-date", default="", help="start date YYYY-MM-DD, defaults to current month start")
    webdriver_collect.add_argument("--end-date", default="", help="end date YYYY-MM-DD, defaults to today")
    webdriver_collect.add_argument("--categories", default="all", help="all or comma-separated policy keys")
    webdriver_collect.add_argument("--stores", default="", help="comma-separated store names or browserOauth values; defaults to all Amazon US stores")
    webdriver_collect.add_argument("--limit", type=int, default=0, help="limit selected stores for smoke testing")
    webdriver_collect.add_argument("--site", default="", help="target site")
    webdriver_collect.add_argument("--page-size", type=int, default=0, help="API page size")
    webdriver_collect.add_argument("--max-pages", type=int, default=0, help="maximum pages per category")
    webdriver_collect.add_argument("--keep-open", action="store_true", help="do not close store windows after collection")
    webdriver_collect.add_argument("--output-dir", default="", help="output directory")
    webdriver_collect.add_argument("--include-items", action="store_true", help="include full item rows in JSON output")
    webdriver_collect.set_defaults(func=execute_webdriver_collect_stores)

    production_run = sub.add_parser("production-run", help="Collect all stores, dedupe, and send DingTalk")
    add_common(production_run)
    production_run.add_argument("--start-date", default="", help="start date YYYY-MM-DD, defaults to current month start")
    production_run.add_argument("--end-date", default="", help="end date YYYY-MM-DD, defaults to today")
    production_run.add_argument("--categories", default="all", help="all or comma-separated policy keys")
    production_run.add_argument("--stores", default="", help="comma-separated store names or store IDs; defaults to all visible Amazon US stores")
    production_run.add_argument("--limit", type=int, default=0, help="limit selected stores for smoke testing")
    production_run.add_argument("--backend", choices=["webdriver", "zclaw"], default="", help="override collector backend")
    production_run.add_argument("--site", default="", help="target site")
    production_run.add_argument("--page-size", type=int, default=0, help="API page size")
    production_run.add_argument("--max-pages", type=int, default=0, help="maximum pages per category")
    production_run.add_argument("--keep-open", action="store_true", help="do not close store windows after collection")
    production_run.add_argument("--output-dir", default="", help="collection output directory")
    production_run.add_argument("--dry-run", action="store_true", help="do not send DingTalk messages")
    production_run.add_argument("--send", action="store_true", help="allow real DingTalk sending")
    production_run.set_defaults(func=execute_production_run)

    send_test = sub.add_parser("send-test", help="发送或 dry-run 一条测试消息")
    add_common(send_test)
    send_test.add_argument("--send", action="store_true", help="真实发送测试消息")
    send_test.set_defaults(func=execute_send_test)

    schedule = sub.add_parser("install-schedule", help="安装 Windows 任务计划")
    add_common(schedule)
    schedule.add_argument("--dry-run", action="store_true", help="只显示 schtasks 命令")
    schedule.add_argument("--task-name", default="", help="任务计划名称")
    schedule.add_argument("--interval-hours", default="", help="执行间隔小时数")
    schedule.add_argument("--python", default="", help="Python 可执行文件")
    schedule.add_argument("--schedule-command", choices=["production-run", "run"], default="", help="任务计划执行命令")
    schedule.add_argument("--skip-preflight", action="store_true", help="跳过配置预检, 仅用于排障")
    schedule.set_defaults(func=execute_install_schedule)

    cdp_daemon = sub.add_parser("install-cdp-daemon", help="安装紫鸟 CDP daemon 开机登录自启任务")
    add_common(cdp_daemon)
    cdp_daemon.add_argument("--dry-run", action="store_true", help="只显示 schtasks 命令")
    cdp_daemon.add_argument("--task-name", default="", help="任务计划名称")
    cdp_daemon.add_argument("--python", default="", help="Python 可执行文件")
    cdp_daemon.add_argument("--log-file", default="", help="daemon 日志路径")
    cdp_daemon.add_argument("--port-start", type=int, default=0, help="扫描起始端口")
    cdp_daemon.add_argument("--port-end", type=int, default=0, help="扫描结束端口")
    cdp_daemon.set_defaults(func=execute_install_cdp_daemon)

    self_test = sub.add_parser("self-test", help="运行本地去重和 dry-run 自测")
    add_common(self_test)
    self_test.set_defaults(func=execute_self_test)
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command == "init-config":
        try:
            path = write_default_config(Path(args.config), force=args.force)
            print_result({"ok": True, "config": str(path)}, args.json)
            return 0
        except Exception as exc:
            print_result({"ok": False, "error": str(exc)}, args.json)
            return 1
    try:
        payload = args.func(args)
        print_result(payload, args.json)
        return 0 if payload.get("ok") else 1
    except Exception as exc:
        print_result({"ok": False, "error": str(exc)}, args.json)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
