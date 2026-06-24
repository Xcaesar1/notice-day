from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import tempfile
import time
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile
import html
import xml.etree.ElementTree as ET


PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_STATE_DIR = PROJECT_ROOT / ".local-state" / "account-health-notifier"
DEFAULT_CONFIG_PATH = DEFAULT_STATE_DIR / "config.json"
DEFAULT_DB_PATH = DEFAULT_STATE_DIR / "state.sqlite"
DEFAULT_DWS_CALL = Path(r"Q:\Dingcli\dws-call.cmd")
DEFAULT_SOURCE_DIR = Path(r"C:\Users\god\Desktop\RPA下载结果\账户状况异常明细")
DEFAULT_STORE_LIST = Path(r"F:\店铺清单.xlsx")
DEFAULT_TASK_NAME = "YD-AmazonAccountHealth-DingTalkNotifier"

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
        "source": {
            "type": "excel_latest",
            "excel_path": "",
            "excel_dir": str(DEFAULT_SOURCE_DIR),
            "store_list_path": str(DEFAULT_STORE_LIST),
            "site": SITE_US,
        },
        "dingtalk": {
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
        "schedule": {
            "task_name": DEFAULT_TASK_NAME,
            "interval_hours": 6,
        },
        "state": {
            "db_path": str(DEFAULT_DB_PATH),
            "result_dir": str(DEFAULT_STATE_DIR / "runs"),
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


def load_items(config: dict[str, Any], args: argparse.Namespace) -> tuple[list[ImpactItem], str]:
    source = config.get("source", {})
    source_type = args.source_type or source.get("type") or "excel_latest"
    target_site = args.site or source.get("site") or SITE_US
    if args.source_excel:
        source_type = "excel"

    if source_type == "sample":
        return load_sample_items(), "sample"
    if source_type == "excel":
        excel_path = Path(args.source_excel or source.get("excel_path") or "")
        items = load_items_from_excel(excel_path, target_site=target_site)
        return items, str(excel_path)
    if source_type == "excel_latest":
        source_dir = Path(args.source_dir or source.get("excel_dir") or DEFAULT_SOURCE_DIR)
        excel_path = find_latest_excel(source_dir)
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


def render_markdown(items: list[ImpactItem], title: str, chunk_index: int, chunk_total: int) -> str:
    lines = [
        f"### {title}",
        "",
        f"- 本次新增/变化: {len(items)} 条",
        f"- 范围: {SITE_US}站, 未解决账号状况异常",
    ]
    if chunk_total > 1:
        lines.append(f"- 分段: {chunk_index}/{chunk_total}")
    lines.append("")

    for item in items:
        asin = item.asin or "未识别"
        sku = item.sku or "未识别"
        lines.append(f"#### {item.store} / {item.category}")
        lines.append(f"- ASIN: `{asin}`")
        lines.append(f"- SKU: `{sku}`")
        if item.date:
            lines.append(f"- 日期: {item.date}")
        if item.reason:
            lines.append(f"- 原因: {item.reason}")
        if item.sales_risk:
            lines.append(f"- 销售风险: {item.sales_risk}")
        if item.action:
            lines.append(f"- 操作: {item.action}")
        if item.rating_impact:
            lines.append(f"- 评级影响: {item.rating_impact}")
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


def send_dingtalk_markdown(
    config: dict[str, Any],
    state_dir: Path,
    title: str,
    markdown: str,
    dry_run: bool,
) -> subprocess.CompletedProcess[str]:
    dingtalk = config.get("dingtalk", {})
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


def validate_runtime_config(
    config_path: Path,
    config: dict[str, Any],
    state_dir: Path,
    require_send_ready: bool = False,
) -> dict[str, Any]:
    issues: list[dict[str, str]] = []
    if not config_path.is_file():
        issues.append(issue("config_missing", f"配置文件不存在: {config_path}"))

    source = config.get("source", {})
    source_type = source.get("type") or "excel_latest"
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
    dws_call = Path(dingtalk.get("dws_call") or DEFAULT_DWS_CALL)
    if not dws_call.is_file():
        issues.append(issue("dws_call_missing", f"DWS 调用入口不存在: {dws_call}"))

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

    db_path = Path(config.get("state", {}).get("db_path") or state_dir / "state.sqlite")
    result_dir = Path(config.get("state", {}).get("result_dir") or DEFAULT_STATE_DIR / "runs")

    if require_send_ready:
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
    status = "success"
    error = ""
    coverage_summary: dict[str, Any] = {}

    try:
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
        items, source = load_items(config, args)
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

            chunks = chunked(candidates, max_items)
            title_prefix = clean_text(config.get("dingtalk", {}).get("title_prefix") or "亚马逊账号状况异常")
            for index, chunk in enumerate(chunks, start=1):
                title = f"{title_prefix}新增通知 {run_id}"
                markdown = render_markdown(chunk, title, index, len(chunks))
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
                if not candidates:
                    status = "no_new_items"
                elif dry_run:
                    status = "dry_run"
    except Exception as exc:
        status = "failed"
        error = str(exc)
        if not source:
            source = "unknown"
        write_run_start(conn, run_id, started_at, source)
    finally:
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
        conn.close()

    return {
        "ok": status in {"success", "no_new_items", "dry_run"},
        "status": status,
        "run_id": run_id,
        "source": source,
        "total_items": total_items,
        "notify_candidates": len(candidates),
        "sent_items": sent_items,
        "artifact": str(artifact),
        "error": error,
        "coverage": coverage_summary,
    }


def execute_doctor(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config_exists = config_path.is_file()
    config = load_config(config_path) if config_exists else default_config()
    state_dir = Path(args.state_dir or config_path.parent or DEFAULT_STATE_DIR)
    dws_call = Path(config.get("dingtalk", {}).get("dws_call") or DEFAULT_DWS_CALL)
    db_path = Path(config.get("state", {}).get("db_path") or state_dir / "state.sqlite")
    checks = {
        "config_exists": config_exists,
        "config_path": str(config_path),
        "db_path": str(db_path),
        "dws_call_exists": dws_call.is_file(),
        "robot_configured": bool(clean_text(config.get("dingtalk", {}).get("robot_code"))),
        "group_configured": bool(clean_text(config.get("dingtalk", {}).get("group_open_conversation_id"))),
    }
    dws_status: dict[str, Any] = {}
    if dws_call.is_file():
        try:
            result = run_dws_call(dws_call, ["auth", "status", "-f", "json"], state_dir, timeout=60)
            dws_status = {
                "exit_code": result.returncode,
                "stdout": result.stdout.strip(),
                "stderr": result.stderr.strip(),
            }
        except Exception as exc:
            dws_status = {"error": str(exc)}
    validation = validate_runtime_config(config_path, config, state_dir, require_send_ready=False)
    ok = checks["dws_call_exists"] and validation["ok"]
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
    )


def execute_send_test(args: argparse.Namespace) -> dict[str, Any]:
    config_path = Path(args.config or DEFAULT_CONFIG_PATH)
    config = load_config(config_path)
    state_dir = Path(args.state_dir or config_path.parent or DEFAULT_STATE_DIR)
    dry_run = not args.send
    if not dry_run:
        validation = validate_runtime_config(config_path, config, state_dir, require_send_ready=True)
        if not validation["ok"]:
            return {
                "ok": False,
                "status": "preflight_failed",
                "dry_run": False,
                "issues": validation["issues"],
            }
    title = f"{config.get('dingtalk', {}).get('title_prefix', '亚马逊账号状况异常')}测试消息"
    markdown = "### 亚马逊账号状况异常测试消息\n\n- 这是一条 DWS 应用机器人连通性测试。\n- 如果你看到这条消息, 说明机器人发群链路可用。\n"
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
    command = f'"{python_exe}" "{script}" run --config "{config_path}"'
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
    validation = validate_runtime_config(config_path, config, state_dir, require_send_ready=True)
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
    schedule.add_argument("--skip-preflight", action="store_true", help="跳过配置预检, 仅用于排障")
    schedule.set_defaults(func=execute_install_schedule)

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
