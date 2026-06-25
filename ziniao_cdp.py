from __future__ import annotations

import argparse
import importlib.util
import json
import re
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
from dataclasses import asdict, dataclass
from datetime import datetime
from typing import Any

import websocket


DEFAULT_PORT_START = 9222
DEFAULT_PORT_END = 9250
DEFAULT_HTTP_TIMEOUT = 2.0
DEFAULT_WS_TIMEOUT = 8.0
DEFAULT_URL_CONTAINS = "sellercentral.amazon.com"
DEFAULT_WATCH_DURATION_SECONDS = 30.0
DEFAULT_WATCH_INTERVAL_SECONDS = 2.0
DEFAULT_RECONNECT_TIMEOUT_SECONDS = 60.0
DEFAULT_BACKOFF_SECONDS = (1.0, 2.0, 5.0, 10.0)

STATUS_CONNECTED = "CONNECTED"
STATUS_NO_CDP_PORT = "NO_CDP_PORT"
STATUS_TARGET_GONE = "TARGET_GONE"
STATUS_CONTEXT_LOST = "CONTEXT_LOST"
STATUS_WS_CLOSED = "WS_CLOSED"
STATUS_PORT_GONE = "PORT_GONE"
STATUS_BUSINESS_NOT_READY = "BUSINESS_NOT_READY"
STATUS_CDP_ERROR = "CDP_ERROR"


class ZiniaoCdpError(RuntimeError):
    pass


def configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


@dataclass(frozen=True)
class PortOwner:
    port: int
    local_address: str = ""
    pid: int = 0
    name: str = ""
    command_line: str = ""

    @property
    def is_ziniao_browser(self) -> bool:
        return self.name.lower() == "ziniaobrowser.exe"

    @property
    def is_localhost(self) -> bool:
        return self.local_address in ("127.0.0.1", "::1", "localhost")

    def safe_dict(self) -> dict[str, Any]:
        return {
            "port": self.port,
            "local_address": self.local_address,
            "pid": self.pid,
            "name": self.name,
            "is_ziniao_browser": self.is_ziniao_browser,
            "is_localhost": self.is_localhost,
            "remote_debugging_port": parse_remote_debugging_port(self.command_line),
            "user_data_dir": parse_user_data_dir(self.command_line),
        }


@dataclass(frozen=True)
class CdpBrowser:
    port: int
    browser: str
    protocol_version: str
    web_socket_debugger_url: str
    owner: PortOwner | None = None

    @property
    def preferred_score(self) -> tuple[int, int, int]:
        if self.owner and self.owner.is_ziniao_browser and self.owner.is_localhost:
            return (0, 0, self.port)
        if self.owner and self.owner.is_ziniao_browser:
            return (1, 0, self.port)
        return (2, 0 if self.owner else 1, self.port)

    def safe_dict(self) -> dict[str, Any]:
        payload = {
            "port": self.port,
            "browser": self.browser,
            "protocol_version": self.protocol_version,
            "web_socket_debugger_url": self.web_socket_debugger_url,
            "preferred": self.preferred_score[0] == 0,
        }
        if self.owner:
            payload["owner"] = self.owner.safe_dict()
        return payload


@dataclass(frozen=True)
class CdpTarget:
    port: int
    id: str
    type: str
    title: str
    url: str
    web_socket_debugger_url: str

    def safe_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class PageProbe:
    target: CdpTarget
    title: str
    url: str
    ready_state: str
    text_length: int
    body_sample: str
    has_account_health_text: bool
    has_us_site_text: bool
    asin_sku_matches: list[str]

    @property
    def business_ready(self) -> bool:
        return self.ready_state == "complete" and self.has_account_health_text

    def safe_dict(self, include_body_sample: bool = True) -> dict[str, Any]:
        payload: dict[str, Any] = {
            "target": self.target.safe_dict(),
            "title": self.title,
            "url": self.url,
            "ready_state": self.ready_state,
            "text_length": self.text_length,
            "has_account_health_text": self.has_account_health_text,
            "has_us_site_text": self.has_us_site_text,
            "business_ready": self.business_ready,
            "asin_sku_matches": self.asin_sku_matches,
        }
        if include_body_sample:
            payload["body_sample"] = self.body_sample
        return payload


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def parse_remote_debugging_port(command_line: str) -> int:
    match = re.search(r"--remote-debugging-port=(\d+)", command_line or "")
    return int(match.group(1)) if match else 0


def parse_user_data_dir(command_line: str) -> str:
    match = re.search(r'--user-data-dir="?([^"\s]+)"?', command_line or "")
    return match.group(1) if match else ""


def _read_json(url: str, timeout: float = DEFAULT_HTTP_TIMEOUT) -> Any:
    request = urllib.request.Request(url, headers={"Accept": "application/json"})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            charset = response.headers.get_content_charset() or "utf-8"
            return json.loads(response.read().decode(charset, errors="replace"))
    except (urllib.error.URLError, TimeoutError, socket.timeout, json.JSONDecodeError) as exc:
        raise ZiniaoCdpError(f"CDP HTTP request failed: {url}: {exc}") from exc


def is_tcp_open(port: int, timeout: float = 0.08) -> bool:
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=timeout):
            return True
    except OSError:
        return False


def _run_powershell_json(script: str, timeout: float = 10.0) -> Any:
    completed = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        text=True,
        encoding="utf-8",
        errors="replace",
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=timeout,
        check=False,
    )
    if completed.returncode != 0:
        raise ZiniaoCdpError(clean_error(completed.stderr or completed.stdout))
    text = completed.stdout.strip()
    if not text:
        return []
    return json.loads(text)


def clean_error(text: str) -> str:
    return re.sub(r"\s+", " ", str(text or "")).strip()


def list_port_owners(port_start: int = DEFAULT_PORT_START, port_end: int = DEFAULT_PORT_END) -> dict[int, PortOwner]:
    ports = ",".join(str(port) for port in range(port_start, port_end + 1))
    script = f"""
$ports = @({ports})
$rows = Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue |
  Where-Object {{ $ports -contains $_.LocalPort }} |
  ForEach-Object {{
    $proc = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue
    [pscustomobject]@{{
      port = [int]$_.LocalPort
      local_address = [string]$_.LocalAddress
      pid = [int]$_.OwningProcess
      name = if ($proc) {{ [string]($proc.ProcessName + '.exe') }} else {{ '' }}
      command_line = ''
    }}
  }}
$rows | ConvertTo-Json -Depth 5
"""
    try:
        data = _run_powershell_json(script)
    except Exception:
        return {}
    if isinstance(data, dict):
        data = [data]
    owners = {}
    for item in data or []:
        try:
            port = int(item.get("port") or 0)
        except Exception:
            continue
        owners[port] = PortOwner(
            port=port,
            local_address=str(item.get("local_address") or ""),
            pid=int(item.get("pid") or 0),
            name=str(item.get("name") or ""),
            command_line=str(item.get("command_line") or ""),
        )
    return owners


def list_process_summary() -> dict[str, Any]:
    script = r"""
$rows = Get-CimInstance Win32_Process |
  Where-Object { $_.Name -match 'ziniaobrowser|ziniao|env-kit|python' } |
  ForEach-Object {
    $cmd = [string]$_.CommandLine
    [pscustomobject]@{
      name = [string]$_.Name
      pid = [int]$_.ProcessId
      has_remote_debugging_port = $cmd.Contains('--remote-debugging-port')
      remote_debugging_port = if ($cmd -match '--remote-debugging-port=(\d+)') { [int]$Matches[1] } else { 0 }
      user_data_dir = if ($cmd -match '--user-data-dir="?([^"\s]+)"?') { [string]$Matches[1] } else { '' }
      is_cdp_daemon = $cmd.Contains('ziniao_cdp_daemon')
    }
  }
$rows | ConvertTo-Json -Depth 5
"""
    try:
        data = _run_powershell_json(script)
    except Exception as exc:
        return {"ok": False, "error": str(exc), "processes": []}
    if isinstance(data, dict):
        data = [data]
    processes = data or []
    return {
        "ok": True,
        "processes": processes,
        "env_kit_alive": any(item.get("name") == "env-kit.exe" for item in processes),
        "ziniao_alive": any(item.get("name") == "ziniao.exe" for item in processes),
        "ziniao_browser_alive": any(item.get("name") == "ziniaobrowser.exe" for item in processes),
        "daemon_alive": any(item.get("is_cdp_daemon") for item in processes),
    }


def get_browser(port: int, timeout: float = DEFAULT_HTTP_TIMEOUT, owner: PortOwner | None = None) -> CdpBrowser:
    data = _read_json(f"http://127.0.0.1:{port}/json/version", timeout=timeout)
    return CdpBrowser(
        port=port,
        browser=str(data.get("Browser") or ""),
        protocol_version=str(data.get("Protocol-Version") or ""),
        web_socket_debugger_url=str(data.get("webSocketDebuggerUrl") or ""),
        owner=owner,
    )


def list_targets(port: int, timeout: float = DEFAULT_HTTP_TIMEOUT) -> list[CdpTarget]:
    data = _read_json(f"http://127.0.0.1:{port}/json/list", timeout=timeout)
    if not isinstance(data, list):
        raise ZiniaoCdpError(f"CDP /json/list returned non-list payload on port {port}")
    targets: list[CdpTarget] = []
    for item in data:
        if not isinstance(item, dict):
            continue
        targets.append(
            CdpTarget(
                port=port,
                id=str(item.get("id") or ""),
                type=str(item.get("type") or ""),
                title=str(item.get("title") or ""),
                url=str(item.get("url") or ""),
                web_socket_debugger_url=str(item.get("webSocketDebuggerUrl") or ""),
            )
        )
    return targets


def scan_ports(
    port_start: int = DEFAULT_PORT_START,
    port_end: int = DEFAULT_PORT_END,
    timeout: float = 0.4,
    include_owners: bool = True,
) -> list[CdpBrowser]:
    owners = list_port_owners(port_start, port_end) if include_owners else {}
    browsers: list[CdpBrowser] = []
    for port in range(port_start, port_end + 1):
        if not is_tcp_open(port):
            continue
        try:
            browser = get_browser(port, timeout=timeout, owner=owners.get(port))
        except ZiniaoCdpError:
            continue
        if browser.web_socket_debugger_url:
            browsers.append(browser)
    return sorted(browsers, key=lambda item: item.preferred_score)


def find_targets(
    url_contains: str = DEFAULT_URL_CONTAINS,
    port: int | None = None,
    port_start: int = DEFAULT_PORT_START,
    port_end: int = DEFAULT_PORT_END,
    timeout: float = DEFAULT_HTTP_TIMEOUT,
) -> list[CdpTarget]:
    ports = [port] if port else [browser.port for browser in scan_ports(port_start, port_end)]
    matches: list[CdpTarget] = []
    seen: set[tuple[str, str]] = set()
    needle = url_contains.lower()
    for candidate_port in ports:
        try:
            targets = list_targets(candidate_port, timeout=timeout)
        except ZiniaoCdpError:
            continue
        for target in targets:
            key = (target.id, target.url)
            if key in seen:
                continue
            if target.type == "page" and needle in target.url.lower() and target.web_socket_debugger_url:
                matches.append(target)
                seen.add(key)
    return matches


class CdpSession:
    def __init__(self, ws_url: str, timeout: float = DEFAULT_WS_TIMEOUT) -> None:
        self._ws_url = ws_url
        self._timeout = timeout
        self._next_id = 0
        self._ws: websocket.WebSocket | None = None

    def __enter__(self) -> "CdpSession":
        self._ws = websocket.create_connection(self._ws_url, timeout=self._timeout, suppress_origin=True)
        self._ws.settimeout(self._timeout)
        return self

    def __exit__(self, exc_type: object, exc: object, tb: object) -> None:
        if self._ws is not None:
            self._ws.close()
            self._ws = None

    def call(self, method: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        if self._ws is None:
            raise ZiniaoCdpError("CDP session is not connected")
        self._next_id += 1
        message_id = self._next_id
        payload = {"id": message_id, "method": method, "params": params or {}}
        try:
            self._ws.send(json.dumps(payload, ensure_ascii=False))
            deadline = time.monotonic() + self._timeout
            while True:
                if time.monotonic() > deadline:
                    raise ZiniaoCdpError(f"CDP call timed out: {method}")
                raw = self._ws.recv()
                message = json.loads(raw)
                if message.get("id") != message_id:
                    continue
                if "error" in message:
                    raise ZiniaoCdpError(f"CDP call failed: {method}: {message['error']}")
                return dict(message.get("result") or {})
        except websocket.WebSocketConnectionClosedException as exc:
            raise ZiniaoCdpError(f"{STATUS_WS_CLOSED}: {exc}") from exc
        except (websocket.WebSocketTimeoutException, socket.timeout) as exc:
            raise ZiniaoCdpError(f"{STATUS_WS_CLOSED}: timeout during {method}") from exc


def evaluate_json(target: CdpTarget, expression: str, timeout: float = DEFAULT_WS_TIMEOUT) -> Any:
    with CdpSession(target.web_socket_debugger_url, timeout=timeout) as session:
        result = session.call(
            "Runtime.evaluate",
            {
                "expression": expression,
                "returnByValue": True,
                "awaitPromise": True,
            },
        )
    if "exceptionDetails" in result:
        details = result["exceptionDetails"]
        if isinstance(details, dict):
            exception = details.get("exception") or {}
            text = (
                exception.get("description")
                or exception.get("value")
                or exception.get("className")
                or details.get("text")
                or str(details)
            )
        else:
            text = str(details)
        raise ZiniaoCdpError(f"CDP Runtime.evaluate exception: {text}")
    value = result.get("result", {}).get("value")
    if isinstance(value, str):
        return json.loads(value)
    return value


def probe_sellercentral_page(
    url_contains: str = DEFAULT_URL_CONTAINS,
    port: int | None = None,
    port_start: int = DEFAULT_PORT_START,
    port_end: int = DEFAULT_PORT_END,
    text_limit: int = 800,
) -> PageProbe:
    targets = find_targets(
        url_contains=url_contains,
        port=port,
        port_start=port_start,
        port_end=port_end,
    )
    if not targets:
        raise ZiniaoCdpError(
            f"{STATUS_TARGET_GONE}: no CDP page target matched {url_contains!r}; "
            "confirm Ziniao browser was opened after CDP daemon injection."
        )
    return probe_target(targets[0], text_limit=text_limit)


def probe_target(target: CdpTarget, text_limit: int = 800) -> PageProbe:
    expression = f"""
JSON.stringify({{
  title: document.title,
  url: location.href,
  readyState: document.readyState,
  textLength: document.body ? document.body.innerText.length : 0,
  bodySample: document.body ? document.body.innerText.slice(0, {int(text_limit)}) : '',
  hasAccountHealthText: document.body ? /账户状况|Account Health|政策合规性|食品和商品安全问题/.test(document.body.innerText) : false,
  hasUsSiteText: document.body ? /美国|United States|US\\b/.test(document.body.innerText) : false,
  asinSkuMatches: document.body ? (document.body.innerText.match(/\\bASIN[:：]?\\s*[A-Z0-9]{{8,12}}|\\bSKU[:：]?\\s*[^\\s]+/g) || []).slice(0, 20) : []
}})
"""
    data = evaluate_json(target, expression)
    return PageProbe(
        target=target,
        title=str(data.get("title") or ""),
        url=str(data.get("url") or ""),
        ready_state=str(data.get("readyState") or ""),
        text_length=int(data.get("textLength") or 0),
        body_sample=str(data.get("bodySample") or ""),
        has_account_health_text=bool(data.get("hasAccountHealthText")),
        has_us_site_text=bool(data.get("hasUsSiteText")),
        asin_sku_matches=[str(item) for item in data.get("asinSkuMatches") or []],
    )


def classify_error(error: BaseException | str, previous_had_browser: bool = False) -> str:
    text = str(error)
    lowered = text.lower()
    if STATUS_WS_CLOSED in text or "websocket" in lowered or "handshake" in lowered:
        return STATUS_WS_CLOSED
    if "execution context was destroyed" in lowered or "cannot find context" in lowered:
        return STATUS_CONTEXT_LOST
    if STATUS_TARGET_GONE in text or "no cdp page target" in lowered:
        return STATUS_TARGET_GONE
    if "connection refused" in lowered or "actively refused" in lowered or "timed out" in lowered:
        return STATUS_PORT_GONE if previous_had_browser else STATUS_NO_CDP_PORT
    return STATUS_CDP_ERROR


def diagnose_once(
    url_contains: str = DEFAULT_URL_CONTAINS,
    port: int | None = None,
    port_start: int = DEFAULT_PORT_START,
    port_end: int = DEFAULT_PORT_END,
    text_limit: int = 800,
    include_body_sample: bool = False,
    previous_had_browser: bool = False,
    include_process: bool = True,
    include_owners: bool = True,
) -> dict[str, Any]:
    started_at = now_iso()
    process_summary = list_process_summary() if include_process else {}
    scan_start = port if port else port_start
    scan_end = port if port else port_end
    browsers = scan_ports(scan_start, scan_end, include_owners=include_owners)
    browser_payload = [browser.safe_dict() for browser in browsers]
    if not browsers:
        status = STATUS_PORT_GONE if previous_had_browser else STATUS_NO_CDP_PORT
        hint = (
            "env-kit.exe is alive but no CDP port is readable; start the CDP daemon, then reopen the store browser."
            if process_summary.get("env_kit_alive")
            else "Ziniao env-kit is not alive; open Ziniao first, start the CDP daemon, then launch a store browser."
        )
        return {
            "ok": False,
            "status": status,
            "started_at": started_at,
            "process": process_summary,
            "browsers": browser_payload,
            "targets": [],
            "error": hint,
        }
    try:
        needle = url_contains.lower()
        targets: list[CdpTarget] = []
        seen: set[tuple[str, str]] = set()
        for browser in browsers:
            try:
                for target in list_targets(browser.port):
                    key = (target.id, target.url)
                    if key in seen:
                        continue
                    if target.type == "page" and needle in target.url.lower() and target.web_socket_debugger_url:
                        targets.append(target)
                        seen.add(key)
            except ZiniaoCdpError:
                continue
        if not targets:
            return {
                "ok": False,
                "status": STATUS_TARGET_GONE,
                "started_at": started_at,
                "process": process_summary,
                "browsers": browser_payload,
                "targets": [],
                "error": f"No Seller Central target matched {url_contains!r}.",
            }
        probe = probe_target(targets[0], text_limit=text_limit)
        status = STATUS_CONNECTED if probe.business_ready else STATUS_BUSINESS_NOT_READY
        return {
            "ok": status == STATUS_CONNECTED,
            "status": status,
            "started_at": started_at,
            "process": process_summary,
            "browsers": browser_payload,
            "targets": [target.safe_dict() for target in targets],
            "probe": probe.safe_dict(include_body_sample=include_body_sample),
        }
    except Exception as exc:
        return {
            "ok": False,
            "status": classify_error(exc, previous_had_browser=bool(browsers)),
            "started_at": started_at,
            "process": process_summary,
            "browsers": browser_payload,
            "targets": [],
            "error": clean_error(str(exc)),
        }


def wait_for_reconnect(args: argparse.Namespace) -> dict[str, Any]:
    deadline = time.monotonic() + float(args.reconnect_timeout_seconds)
    attempts = []
    delay_index = 0
    previous_had_browser = False
    while True:
        attempt = diagnose_once(
            url_contains=args.url_contains,
            port=args.port or None,
            port_start=args.port_start,
            port_end=args.port_end,
            text_limit=args.text_limit,
            previous_had_browser=previous_had_browser,
            include_process=False,
            include_owners=False,
        )
        previous_had_browser = bool(attempt.get("browsers"))
        attempts.append(
            {
                "at": now_iso(),
                "status": attempt.get("status"),
                "ok": attempt.get("ok"),
                "target": attempt.get("probe", {}).get("target", {}),
                "error": attempt.get("error", ""),
            }
        )
        if attempt.get("ok"):
            return {"ok": True, "status": STATUS_CONNECTED, "attempts": attempts, "result": attempt}
        if time.monotonic() >= deadline:
            return {"ok": False, "status": "RECONNECT_TIMEOUT", "attempts": attempts, "result": attempt}
        delay = DEFAULT_BACKOFF_SECONDS[min(delay_index, len(DEFAULT_BACKOFF_SECONDS) - 1)]
        delay_index += 1
        time.sleep(min(delay, max(0.1, deadline - time.monotonic())))


def probe_payload(args: argparse.Namespace) -> dict[str, Any]:
    browsers = scan_ports(args.port_start, args.port_end)
    probe = probe_sellercentral_page(
        url_contains=args.url_contains,
        port=args.port or None,
        port_start=args.port_start,
        port_end=args.port_end,
        text_limit=args.text_limit,
    )
    return {
        "ok": True,
        "browsers": [browser.safe_dict() for browser in browsers],
        "probe": probe.safe_dict(include_body_sample=True),
    }


def doctor_payload(args: argparse.Namespace) -> dict[str, Any]:
    checks = {
        "websocket_client_available": importlib.util.find_spec("websocket") is not None,
        "frida_available": importlib.util.find_spec("frida") is not None,
    }
    result = diagnose_once(
        url_contains=args.url_contains,
        port=args.port or None,
        port_start=args.port_start,
        port_end=args.port_end,
        text_limit=args.text_limit,
        include_body_sample=bool(args.include_body_sample),
    )
    result["checks"] = checks
    result["ok"] = bool(result.get("ok") and all(checks.values()))
    return result


def watch_payload(args: argparse.Namespace) -> dict[str, Any]:
    duration = float(args.duration_seconds)
    interval = float(args.interval_seconds)
    deadline = time.monotonic() + duration
    events = []
    previous_status = ""
    previous_had_browser = False
    while True:
        result = diagnose_once(
            url_contains=args.url_contains,
            port=args.port or None,
            port_start=args.port_start,
            port_end=args.port_end,
            text_limit=args.text_limit,
            previous_had_browser=previous_had_browser,
            include_process=False,
            include_owners=False,
        )
        previous_had_browser = bool(result.get("browsers"))
        event = {
            "at": now_iso(),
            "status": result.get("status"),
            "ok": result.get("ok"),
            "changed": result.get("status") != previous_status,
            "target": result.get("probe", {}).get("target", {}),
            "error": result.get("error", ""),
        }
        previous_status = str(result.get("status") or "")
        events.append(event)
        if time.monotonic() >= deadline:
            break
        time.sleep(interval)
    return {
        "ok": bool(events and events[-1].get("ok")),
        "status": events[-1]["status"] if events else STATUS_CDP_ERROR,
        "events": events,
    }


def close_target(target: CdpTarget) -> dict[str, Any]:
    browser = get_browser(target.port)
    if not browser.web_socket_debugger_url:
        raise ZiniaoCdpError(f"Browser WebSocket URL missing on port {target.port}")
    with CdpSession(browser.web_socket_debugger_url) as session:
        result = session.call("Target.closeTarget", {"targetId": target.id})
    return {"ok": bool(result.get("success")), "target": target.safe_dict(), "result": result}


def lifecycle_payload(args: argparse.Namespace) -> dict[str, Any]:
    initial = diagnose_once(
        url_contains=args.url_contains,
        port=args.port or None,
        port_start=args.port_start,
        port_end=args.port_end,
        text_limit=args.text_limit,
        include_body_sample=False,
    )
    payload: dict[str, Any] = {"ok": bool(initial.get("ok")), "initial": initial, "steps": []}
    if not initial.get("ok"):
        return payload
    target_payload = initial.get("probe", {}).get("target", {})
    target = CdpTarget(
        port=int(target_payload.get("port") or 0),
        id=str(target_payload.get("id") or ""),
        type=str(target_payload.get("type") or ""),
        title=str(target_payload.get("title") or ""),
        url=str(target_payload.get("url") or ""),
        web_socket_debugger_url=str(target_payload.get("web_socket_debugger_url") or ""),
    )
    if not args.close_target:
        payload["steps"].append(
            {
                "action": "close_target",
                "skipped": True,
                "reason": "Pass --close-target to close the current Seller Central target for a real disconnect test.",
            }
        )
        return payload
    closed = close_target(target)
    payload["steps"].append({"action": "close_target", "result": closed})
    time.sleep(float(args.disconnect_wait_seconds))
    after_close = diagnose_once(
        url_contains=args.url_contains,
        port=args.port or None,
        port_start=args.port_start,
        port_end=args.port_end,
        text_limit=args.text_limit,
        previous_had_browser=True,
    )
    payload["steps"].append({"action": "after_close_diagnose", "result": after_close})
    if args.wait_reopen:
        payload["steps"].append({"action": "wait_for_reconnect", "result": wait_for_reconnect(args)})
    payload["ok"] = True
    return payload


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Read Ziniao-launched pages through Chrome DevTools Protocol.")
    parser.add_argument("command", choices=["probe", "doctor", "watch", "lifecycle-test"], help="Command to run")
    parser.add_argument("--json", action="store_true", help="Print JSON")
    parser.add_argument("--port", type=int, default=0, help="Use a known CDP port instead of scanning")
    parser.add_argument("--port-start", type=int, default=DEFAULT_PORT_START, help="First CDP port to scan")
    parser.add_argument("--port-end", type=int, default=DEFAULT_PORT_END, help="Last CDP port to scan")
    parser.add_argument("--url-contains", default=DEFAULT_URL_CONTAINS, help="Target URL substring")
    parser.add_argument("--text-limit", type=int, default=800, help="Body sample character limit")
    parser.add_argument("--include-body-sample", action="store_true", help="Include body sample in doctor output")
    parser.add_argument("--duration-seconds", type=float, default=DEFAULT_WATCH_DURATION_SECONDS)
    parser.add_argument("--interval-seconds", type=float, default=DEFAULT_WATCH_INTERVAL_SECONDS)
    parser.add_argument("--reconnect-timeout-seconds", type=float, default=DEFAULT_RECONNECT_TIMEOUT_SECONDS)
    parser.add_argument("--close-target", action="store_true", help="Close current Seller Central target during lifecycle test")
    parser.add_argument("--wait-reopen", action="store_true", help="Wait for a reopened Seller Central target")
    parser.add_argument("--disconnect-wait-seconds", type=float, default=3.0)
    return parser


def main(argv: list[str] | None = None) -> int:
    configure_stdio()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "probe":
            payload = probe_payload(args)
        elif args.command == "doctor":
            payload = doctor_payload(args)
        elif args.command == "watch":
            payload = watch_payload(args)
        elif args.command == "lifecycle-test":
            payload = lifecycle_payload(args)
        else:
            raise ZiniaoCdpError(f"Unknown command: {args.command}")
    except Exception as exc:
        payload = {"ok": False, "status": classify_error(exc), "error": clean_error(str(exc))}
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            print(f"error: {exc}", file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(f"ok: {payload.get('ok')}")
        print(f"status: {payload.get('status')}")
    return 0 if payload.get("ok") else 1


if __name__ == "__main__":
    raise SystemExit(main())
