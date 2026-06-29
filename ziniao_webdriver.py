from __future__ import annotations

import argparse
import json
import os
import subprocess
import time
import uuid
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any
from urllib import error as urlerror
from urllib import request as urlrequest

import cdp_account_health
import ziniao_cdp


WEBDRIVER_URL_TEMPLATE = "http://127.0.0.1:{port}"
PERFORMANCE_DASHBOARD_URL = "https://sellercentral.amazon.com/performance/dashboard"
DEFAULT_PORT = 9515
DEFAULT_STARTUP_TIMEOUT = 60
DEFAULT_REQUEST_TIMEOUT = 30
DEFAULT_UPDATE_CORE_TIMEOUT = 300
DEFAULT_UPDATE_CORE_POLL_SECONDS = 2.0
DEFAULT_BROWSER_READY_TIMEOUT = 60
DEFAULT_NAVIGATION_WAIT_SECONDS = 3.0
DEFAULT_STOP_TIMEOUT = 30

STATUS_RETRYABLE_UPDATE_CORE = {10000}
STATUS_FATAL = {-10002, -10003, -10004, -10013}
STATUS_RETRYABLE_START = {-10006}


class ZiniaoWebDriverError(RuntimeError):
    pass


@dataclass(frozen=True)
class Credentials:
    company: str
    username: str
    password: str


@dataclass(frozen=True)
class Store:
    browser_oauth: str
    browser_name: str
    site_id: str
    platform_id: str
    platform_name: str
    browser_id: str = ""
    browser_ip: str = ""
    is_expired: bool = False


@dataclass(frozen=True)
class StartedBrowser:
    browser_oauth: str
    debugging_port: int
    launcher_page: str
    core_version: str
    core_type: str
    ip: str = ""
    browser_path: str = ""
    download_path: str = ""
    user_data: str = ""
    duplicate: int = 0


def clean_text(value: Any) -> str:
    return cdp_account_health.clean_text(value)


def current_month_range(today: date | None = None) -> tuple[str, str]:
    return cdp_account_health.current_month_range(today)


def selected_categories(value: str = "all") -> list[cdp_account_health.PolicyCategory]:
    return cdp_account_health.selected_categories(value)


def normalize_for_key(value: Any) -> str:
    return "".join(clean_text(value).lower().split())


def build_args(config: dict[str, Any], args: argparse.Namespace) -> dict[str, Any]:
    cdp_config = config.get("ziniao_cdp", {}) if isinstance(config, dict) else {}
    webdriver_config = config.get("ziniao_webdriver", {}) if isinstance(config, dict) else {}
    password = clean_text(webdriver_config.get("password"))
    password_env = clean_text(webdriver_config.get("password_env"))
    if not password and password_env:
        password = clean_text(os.environ.get(password_env))
    return {
        "client_path": clean_text(webdriver_config.get("client_path")),
        "port": int(webdriver_config.get("port") or DEFAULT_PORT),
        "request_timeout": int(webdriver_config.get("request_timeout_seconds") or DEFAULT_REQUEST_TIMEOUT),
        "startup_timeout": int(webdriver_config.get("startup_timeout_seconds") or DEFAULT_STARTUP_TIMEOUT),
        "update_core_timeout": int(webdriver_config.get("update_core_timeout_seconds") or DEFAULT_UPDATE_CORE_TIMEOUT),
        "update_core_poll_seconds": float(
            webdriver_config.get("update_core_poll_seconds") or DEFAULT_UPDATE_CORE_POLL_SECONDS
        ),
        "browser_ready_timeout": int(
            webdriver_config.get("browser_ready_timeout_seconds") or DEFAULT_BROWSER_READY_TIMEOUT
        ),
        "navigation_wait_seconds": float(
            webdriver_config.get("navigation_wait_seconds") or DEFAULT_NAVIGATION_WAIT_SECONDS
        ),
        "page_size": int(
            getattr(args, "page_size", 0)
            or cdp_config.get("collect_page_size")
            or cdp_account_health.DEFAULT_PAGE_SIZE
        ),
        "max_pages": int(
            getattr(args, "max_pages", 0)
            or cdp_config.get("collect_max_pages")
            or cdp_account_health.DEFAULT_MAX_PAGES
        ),
        "credentials": Credentials(
            company=clean_text(webdriver_config.get("company")),
            username=clean_text(webdriver_config.get("username")),
            password=password,
        ),
    }


def request_id(prefix: str) -> str:
    return f"{prefix}-{uuid.uuid4().hex[:12]}"


def base_url(port: int) -> str:
    return WEBDRIVER_URL_TEMPLATE.format(port=int(port))


def _status_code(payload: dict[str, Any]) -> int:
    try:
        return int(payload.get("statusCode") or 0)
    except Exception:
        return -99999


def _response_detail(payload: dict[str, Any]) -> str:
    parts = [
        clean_text(payload.get("err")),
        clean_text(payload.get("msg")),
        clean_text(payload.get("statusMsg")),
        clean_text(payload.get("LastError")),
    ]
    detail = " | ".join(part for part in parts if part)
    return detail or clean_text(payload)


def _post_json(port: int, payload: dict[str, Any], timeout: int | float = DEFAULT_REQUEST_TIMEOUT) -> dict[str, Any]:
    body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
    request = urlrequest.Request(
        base_url(port),
        data=body,
        headers={
            "Accept": "application/json",
            "Content-Type": "application/json; charset=utf-8",
        },
        method="POST",
    )
    try:
        with urlrequest.urlopen(request, timeout=timeout) as response:
            raw = response.read()
    except urlerror.HTTPError as exc:
        raw = exc.read()
    except (urlerror.URLError, TimeoutError) as exc:
        raise ZiniaoWebDriverError(f"WebDriver request failed on {base_url(port)}: {clean_text(exc)}") from exc
    text = raw.decode("utf-8", errors="replace").strip()
    if not text:
        return {}
    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise ZiniaoWebDriverError(f"WebDriver returned non-JSON response: {text[:500]}") from exc


def _credential_payload(credentials: Credentials, action: str, **extra: Any) -> dict[str, Any]:
    return {
        "company": credentials.company,
        "username": credentials.username,
        "password": credentials.password,
        "action": action,
        "requestId": request_id(action),
        **extra,
    }


def _start_browser_payloads(store: Store) -> list[dict[str, Any]]:
    payloads: list[dict[str, Any]] = []
    browser_id = clean_text(store.browser_id)
    browser_oauth = clean_text(store.browser_oauth)
    if browser_id:
        payloads.append({"browserId": browser_id})
    if browser_oauth:
        payloads.append({"browserOauth": browser_oauth})
    return payloads


def get_running_info(port: int, timeout: int | float = DEFAULT_REQUEST_TIMEOUT) -> dict[str, Any]:
    payload = {"action": "getRunningInfo", "requestId": request_id("getRunningInfo")}
    return _post_json(port, payload, timeout=timeout)


def ensure_webdriver_server(
    client_path: str,
    port: int,
    startup_timeout: int = DEFAULT_STARTUP_TIMEOUT,
) -> dict[str, Any]:
    try:
        payload = get_running_info(port, timeout=3)
        return {"started": False, "probe": payload}
    except Exception:
        pass
    if not client_path:
        raise ZiniaoWebDriverError("ziniao_webdriver.client_path is empty")
    executable = Path(client_path)
    if not executable.is_file():
        raise ZiniaoWebDriverError(f"WebDriver client not found: {executable}")
    subprocess.Popen(
        [str(executable), "--run_type=web_driver", "--ipc_type=http", "--port", str(int(port))],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        stdin=subprocess.DEVNULL,
    )
    deadline = time.time() + max(int(startup_timeout), 1)
    last_error = ""
    while time.time() < deadline:
        try:
            payload = get_running_info(port, timeout=3)
            return {"started": True, "probe": payload}
        except Exception as exc:
            last_error = clean_text(exc)
            time.sleep(1)
    raise ZiniaoWebDriverError(
        f"WebDriver server did not become ready on port {port} within {startup_timeout}s: {last_error}"
    )


def update_core(
    port: int,
    credentials: Credentials,
    timeout: int = DEFAULT_UPDATE_CORE_TIMEOUT,
    poll_seconds: float = DEFAULT_UPDATE_CORE_POLL_SECONDS,
    request_timeout: int | float = DEFAULT_REQUEST_TIMEOUT,
) -> dict[str, Any]:
    deadline = time.time() + max(int(timeout), 1)
    last_payload: dict[str, Any] = {}
    while time.time() < deadline:
        try:
            payload = _post_json(
                port,
                _credential_payload(credentials, "updateCore"),
                timeout=request_timeout,
            )
        except Exception as exc:
            last_payload = {"err": clean_text(exc)}
            time.sleep(max(float(poll_seconds), 0.5))
            continue
        last_payload = payload
        status = _status_code(payload)
        if status == 0:
            return payload
        if status in STATUS_FATAL:
            raise ZiniaoWebDriverError(
                f"updateCore failed status={status}: {_response_detail(payload)}"
            )
        if status not in STATUS_RETRYABLE_UPDATE_CORE and _response_detail(payload):
            time.sleep(max(float(poll_seconds), 0.5))
        else:
            time.sleep(max(float(poll_seconds), 0.5))
    raise ZiniaoWebDriverError(
        f"updateCore timed out after {timeout}s: {_response_detail(last_payload)}"
    )


def get_browser_list(
    port: int,
    credentials: Credentials,
    request_timeout: int | float = DEFAULT_REQUEST_TIMEOUT,
) -> list[Store]:
    payload = _post_json(
        port,
        _credential_payload(credentials, "getBrowserList"),
        timeout=request_timeout,
    )
    status = _status_code(payload)
    if status != 0:
        raise ZiniaoWebDriverError(
            f"getBrowserList failed status={status}: {_response_detail(payload)}"
        )
    stores: list[Store] = []
    for item in payload.get("browserList") or []:
        if not isinstance(item, dict):
            continue
        browser_oauth = clean_text(item.get("browserOauth"))
        browser_name = clean_text(item.get("browserName"))
        if not browser_oauth or not browser_name:
            continue
        stores.append(
            Store(
                browser_oauth=browser_oauth,
                browser_name=browser_name,
                site_id=clean_text(item.get("siteId")),
                platform_id=clean_text(item.get("platform_id")),
                platform_name=clean_text(item.get("platform_name")),
                browser_id=clean_text(item.get("browserId")),
                browser_ip=clean_text(item.get("browserIp")),
                is_expired=bool(item.get("isExpired")),
            )
        )
    return stores


def filter_us_amazon_stores(stores: list[Store]) -> list[Store]:
    selected: list[Store] = []
    for store in stores:
        platform = store.platform_name.lower()
        site_id = clean_text(store.site_id)
        is_amazon = "amazon" in platform or "亚马逊" in store.platform_name or store.platform_id == "1"
        is_us = site_id == "1" or "美国" in store.platform_name or "united states" in platform or platform.endswith("us")
        if is_amazon and is_us:
            selected.append(store)
    return selected


def _parse_store_filter(value: str) -> set[str]:
    return {normalize_for_key(part) for part in str(value or "").split(",") if clean_text(part)}


def select_stores(stores: list[Store], names_or_ids: str = "", limit: int = 0) -> list[Store]:
    selected = filter_us_amazon_stores(stores)
    filters = _parse_store_filter(names_or_ids)
    if filters:
        selected = [
            store
            for store in selected
            if normalize_for_key(store.browser_name) in filters
            or normalize_for_key(store.browser_oauth) in filters
            or normalize_for_key(store.browser_id) in filters
        ]
    if limit > 0:
        selected = selected[:limit]
    return selected


def start_browser(
    port: int,
    credentials: Credentials,
    store: Store,
    request_timeout: int | float = DEFAULT_REQUEST_TIMEOUT,
    attempts: int = 2,
) -> StartedBrowser:
    attempts = max(int(attempts), 1)
    identifier_payloads = _start_browser_payloads(store)
    if not identifier_payloads:
        raise ZiniaoWebDriverError(f"startBrowser missing browserId/browserOauth for {store.browser_name}")
    last_payload: dict[str, Any] = {}
    last_error = ""
    last_identifier = ""
    for identifier_payload in identifier_payloads:
        identifier_name = "browserId" if "browserId" in identifier_payload else "browserOauth"
        identifier_value = clean_text(identifier_payload.get(identifier_name))
        last_identifier = f"{identifier_name}={identifier_value}"
        for attempt in range(attempts):
            try:
                payload = _post_json(
                    port,
                    _credential_payload(credentials, "startBrowser", **identifier_payload),
                    timeout=request_timeout,
                )
            except Exception as exc:
                last_error = clean_text(exc)
                if attempt + 1 < attempts:
                    time.sleep(5)
                    continue
                break
            last_payload = payload
            status = _status_code(payload)
            if status == 0:
                debugging_port = int(payload.get("debuggingPort") or 0)
                if debugging_port <= 0:
                    raise ZiniaoWebDriverError(
                        f"startBrowser returned invalid debuggingPort for {store.browser_name} ({identifier_name}={identifier_value}): {payload}"
                    )
                return StartedBrowser(
                    browser_oauth=store.browser_oauth,
                    debugging_port=debugging_port,
                    launcher_page=clean_text(payload.get("launcherPage")),
                    core_version=clean_text(payload.get("core_version") or payload.get("coreVersion")),
                    core_type=clean_text(payload.get("core_type") or payload.get("coreType")),
                    ip=clean_text(payload.get("ip")),
                    browser_path=clean_text(payload.get("browserPath")),
                    download_path=clean_text(payload.get("downloadPath")),
                    user_data=clean_text(payload.get("userData")),
                    duplicate=int(payload.get("duplicate") or 0),
                )
            if status in STATUS_RETRYABLE_START and attempt + 1 < attempts:
                time.sleep(5)
                continue
            last_error = f"status={status}: {_response_detail(payload)}"
            break
    detail = last_error or _response_detail(last_payload) or "unknown"
    raise ZiniaoWebDriverError(
        f"startBrowser failed store={store.browser_name} ({last_identifier}): {detail}"
    )


def stop_browser(
    port: int,
    credentials: Credentials,
    store: Store,
    request_timeout: int | float = DEFAULT_STOP_TIMEOUT,
) -> dict[str, Any]:
    identifier_payloads = _start_browser_payloads(store)
    if not identifier_payloads:
        raise ZiniaoWebDriverError(f"stopBrowser missing browserId/browserOauth for {store.browser_name}")
    identifier_payload = identifier_payloads[0]
    payload = _post_json(
        port,
        _credential_payload(credentials, "stopBrowser", **identifier_payload),
        timeout=request_timeout,
    )
    status = _status_code(payload)
    if status != 0:
        raise ZiniaoWebDriverError(
            f"stopBrowser failed store={store.browser_name} status={status}: {_response_detail(payload)}"
        )
    return payload


def wait_for_page_target(debugging_port: int, timeout: int = DEFAULT_BROWSER_READY_TIMEOUT) -> ziniao_cdp.CdpTarget:
    deadline = time.time() + max(int(timeout), 1)
    last_error = ""
    while time.time() < deadline:
        try:
            targets = ziniao_cdp.list_targets(debugging_port)
            pages = [
                target
                for target in targets
                if target.type == "page" and not clean_text(target.url).startswith("chrome-extension://")
            ]
            if pages:
                return pages[0]
        except Exception as exc:
            last_error = clean_text(exc)
        time.sleep(1)
    raise ZiniaoWebDriverError(
        f"no page target found on debuggingPort={debugging_port} within {timeout}s: {last_error}"
    )


def navigate_target(
    target: ziniao_cdp.CdpTarget,
    url: str,
    wait_seconds: float = DEFAULT_NAVIGATION_WAIT_SECONDS,
) -> ziniao_cdp.CdpTarget:
    with ziniao_cdp.CdpSession(target.web_socket_debugger_url, timeout=20) as session:
        session.call("Page.enable")
        session.call("Page.navigate", {"url": url})
    time.sleep(max(float(wait_seconds), 0.0))
    return wait_for_page_target(target.port, timeout=max(int(wait_seconds) + 10, 10))


def probe_target(target: ziniao_cdp.CdpTarget, text_limit: int = 600) -> dict[str, Any]:
    try:
        probe = ziniao_cdp.probe_target(target, text_limit=text_limit)
        return probe.safe_dict(include_body_sample=True)
    except Exception as exc:
        return {"error": clean_text(exc), "target": target.safe_dict()}


def _looks_like_login_or_verification(snapshot: dict[str, Any]) -> bool:
    url = clean_text(snapshot.get("url")).lower()
    text = clean_text(snapshot.get("body_sample"))
    if "/ap/signin" in url or "/ap/mfa" in url:
        return True
    if "切换账户" in text and "添加账户" in text:
        return True
    if "switch accounts" in text.lower() and "add account" in text.lower():
        return True
    if "输入手机号或邮箱" in text and "继续" in text:
        return True
    if "验证码" in text or "two-step verification" in text.lower():
        return True
    return False


def collect_store_account_health(
    store: Store,
    credentials: Credentials,
    webdriver_port: int,
    start_date: str,
    end_date: str,
    categories: list[cdp_account_health.PolicyCategory],
    page_size: int = cdp_account_health.DEFAULT_PAGE_SIZE,
    max_pages: int = cdp_account_health.DEFAULT_MAX_PAGES,
    close_after: bool = True,
    request_timeout: int | float = DEFAULT_REQUEST_TIMEOUT,
    browser_ready_timeout: int = DEFAULT_BROWSER_READY_TIMEOUT,
    navigation_wait_seconds: float = DEFAULT_NAVIGATION_WAIT_SECONDS,
) -> dict[str, Any]:
    started_at = time.strftime("%Y-%m-%d %H:%M:%S")
    rows: list[dict[str, str]] = []
    target: ziniao_cdp.CdpTarget | None = None
    page_reports: list[dict[str, Any]] = []
    page_snapshot: dict[str, Any] = {}
    started_browser: StartedBrowser | None = None
    status = "success"
    error = ""
    try:
        started_browser = start_browser(
            webdriver_port,
            credentials,
            store,
            request_timeout=request_timeout,
        )
        target = wait_for_page_target(started_browser.debugging_port, timeout=browser_ready_timeout)
        initial_url = clean_text(target.url).lower()
        if initial_url == "about:blank" or "sellercentral.amazon.com" not in initial_url:
            warmup_url = started_browser.launcher_page or PERFORMANCE_DASHBOARD_URL
            target = navigate_target(target, warmup_url, wait_seconds=navigation_wait_seconds)
            target = navigate_target(target, PERFORMANCE_DASHBOARD_URL, wait_seconds=navigation_wait_seconds)
        page_snapshot = probe_target(target, text_limit=800)
        if _looks_like_login_or_verification(page_snapshot):
            raise ZiniaoWebDriverError(
                f"Amazon login or verification is required for {store.browser_name}: "
                f"{clean_text(page_snapshot.get('url'))} | {clean_text(page_snapshot.get('body_sample'))[:300]}"
            )
        result = cdp_account_health.collect_target_account_health(
            target,
            start_date=start_date,
            end_date=end_date,
            categories=categories,
            page_size=page_size,
            max_pages=max_pages,
            store_override=store.browser_name,
            site_override="美国",
        )
        rows = result.get("rows", [])
        page_reports = result.get("page_reports", [])
        target_payload = result.get("target", target.safe_dict())
        ended_at = time.strftime("%Y-%m-%d %H:%M:%S")
        return {
            "ok": True,
            "status": "success",
            "store": store.browser_name,
            "site": "美国",
            "row_count": len(rows),
            "error": "",
            "target": target_payload,
            "page_reports": page_reports,
            "page_snapshot": page_snapshot,
            "debugging_port": started_browser.debugging_port,
            "started_at": started_at,
            "ended_at": ended_at,
            "rows": rows,
        }
    except Exception as exc:
        status = "failed"
        error = clean_text(exc)
        ended_at = time.strftime("%Y-%m-%d %H:%M:%S")
        return {
            "ok": False,
            "status": status,
            "store": store.browser_name,
            "site": "美国",
            "row_count": 0,
            "error": error,
            "target": target.safe_dict() if target else {},
            "page_reports": page_reports,
            "page_snapshot": page_snapshot,
            "debugging_port": started_browser.debugging_port if started_browser else 0,
            "started_at": started_at,
            "ended_at": ended_at,
            "rows": rows,
        }
    finally:
        if close_after and started_browser is not None:
            try:
                stop_browser(webdriver_port, credentials, store, request_timeout=DEFAULT_STOP_TIMEOUT)
            except Exception:
                pass


def collect_visible_account_health(
    client_path: str,
    webdriver_port: int,
    credentials: Credentials,
    start_date: str,
    end_date: str,
    categories: list[cdp_account_health.PolicyCategory],
    stores: str = "",
    limit: int = 0,
    page_size: int = cdp_account_health.DEFAULT_PAGE_SIZE,
    max_pages: int = cdp_account_health.DEFAULT_MAX_PAGES,
    close_after: bool = True,
    startup_timeout: int = DEFAULT_STARTUP_TIMEOUT,
    update_core_timeout: int = DEFAULT_UPDATE_CORE_TIMEOUT,
    update_core_poll_seconds: float = DEFAULT_UPDATE_CORE_POLL_SECONDS,
    request_timeout: int | float = DEFAULT_REQUEST_TIMEOUT,
    browser_ready_timeout: int = DEFAULT_BROWSER_READY_TIMEOUT,
    navigation_wait_seconds: float = DEFAULT_NAVIGATION_WAIT_SECONDS,
) -> dict[str, Any]:
    ensure_webdriver_server(client_path, webdriver_port, startup_timeout=startup_timeout)
    update_core(
        webdriver_port,
        credentials,
        timeout=update_core_timeout,
        poll_seconds=update_core_poll_seconds,
        request_timeout=request_timeout,
    )
    all_stores = get_browser_list(webdriver_port, credentials, request_timeout=request_timeout)
    selected = select_stores(all_stores, names_or_ids=stores, limit=limit)
    target_results: list[dict[str, Any]] = []
    rows: list[dict[str, str]] = []
    for store in selected:
        result = collect_store_account_health(
            store,
            credentials=credentials,
            webdriver_port=webdriver_port,
            start_date=start_date,
            end_date=end_date,
            categories=categories,
            page_size=page_size,
            max_pages=max_pages,
            close_after=close_after,
            request_timeout=request_timeout,
            browser_ready_timeout=browser_ready_timeout,
            navigation_wait_seconds=navigation_wait_seconds,
        )
        target_results.append({key: value for key, value in result.items() if key != "rows"})
        rows.extend(result.get("rows", []))
    ok = bool(target_results) and all(item.get("ok") for item in target_results)
    return {
        "ok": ok,
        "status": "success" if ok else ("no_targets" if not target_results else "partial"),
        "start_date": start_date,
        "end_date": end_date,
        "selected_store_count": len(selected),
        "selected_stores": [store.browser_name for store in selected],
        "all_store_count": len(all_stores),
        "rows": rows,
        "row_count": len(rows),
        "target_results": target_results,
    }
