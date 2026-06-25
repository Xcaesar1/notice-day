from __future__ import annotations

import argparse
import json
import socket
import sys
import time
import urllib.error
import urllib.request
from dataclasses import asdict, dataclass
from typing import Any

import websocket


DEFAULT_PORT_START = 9222
DEFAULT_PORT_END = 9250
DEFAULT_HTTP_TIMEOUT = 2.0
DEFAULT_WS_TIMEOUT = 8.0
DEFAULT_URL_CONTAINS = "sellercentral.amazon.com"


class ZiniaoCdpError(RuntimeError):
    pass


def configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


@dataclass(frozen=True)
class CdpBrowser:
    port: int
    browser: str
    protocol_version: str
    web_socket_debugger_url: str


@dataclass(frozen=True)
class CdpTarget:
    port: int
    id: str
    type: str
    title: str
    url: str
    web_socket_debugger_url: str


@dataclass(frozen=True)
class PageProbe:
    target: CdpTarget
    title: str
    url: str
    ready_state: str
    text_length: int
    body_sample: str
    has_account_health_text: bool
    asin_sku_matches: list[str]


def _read_json(url: str, timeout: float = DEFAULT_HTTP_TIMEOUT) -> Any:
    request = urllib.request.Request(url, headers={"Accept": "application/json"})
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            charset = response.headers.get_content_charset() or "utf-8"
            return json.loads(response.read().decode(charset, errors="replace"))
    except (urllib.error.URLError, TimeoutError, socket.timeout, json.JSONDecodeError) as exc:
        raise ZiniaoCdpError(f"CDP HTTP request failed: {url}: {exc}") from exc


def get_browser(port: int, timeout: float = DEFAULT_HTTP_TIMEOUT) -> CdpBrowser:
    data = _read_json(f"http://127.0.0.1:{port}/json/version", timeout=timeout)
    return CdpBrowser(
        port=port,
        browser=str(data.get("Browser") or ""),
        protocol_version=str(data.get("Protocol-Version") or ""),
        web_socket_debugger_url=str(data.get("webSocketDebuggerUrl") or ""),
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
) -> list[CdpBrowser]:
    browsers: list[CdpBrowser] = []
    for port in range(port_start, port_end + 1):
        try:
            browser = get_browser(port, timeout=timeout)
        except ZiniaoCdpError:
            continue
        if browser.web_socket_debugger_url:
            browsers.append(browser)
    return browsers


def find_targets(
    url_contains: str = DEFAULT_URL_CONTAINS,
    port: int | None = None,
    port_start: int = DEFAULT_PORT_START,
    port_end: int = DEFAULT_PORT_END,
    timeout: float = DEFAULT_HTTP_TIMEOUT,
) -> list[CdpTarget]:
    ports = [port] if port else [browser.port for browser in scan_ports(port_start, port_end)]
    matches: list[CdpTarget] = []
    needle = url_contains.lower()
    for candidate_port in ports:
        try:
            targets = list_targets(candidate_port, timeout=timeout)
        except ZiniaoCdpError:
            continue
        for target in targets:
            if target.type == "page" and needle in target.url.lower() and target.web_socket_debugger_url:
                matches.append(target)
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
            f"No CDP page target matched {url_contains!r}; confirm Ziniao browser was opened after CDP daemon injection."
        )
    target = targets[0]
    expression = f"""
JSON.stringify({{
  title: document.title,
  url: location.href,
  readyState: document.readyState,
  textLength: document.body ? document.body.innerText.length : 0,
  bodySample: document.body ? document.body.innerText.slice(0, {int(text_limit)}) : '',
  hasAccountHealthText: document.body ? /账户状况|Account Health|政策合规性|食品和商品安全问题/.test(document.body.innerText) : false,
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
        asin_sku_matches=[str(item) for item in data.get("asinSkuMatches") or []],
    )


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
        "browsers": [asdict(browser) for browser in browsers],
        "probe": asdict(probe),
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Read Ziniao-launched pages through Chrome DevTools Protocol.")
    parser.add_argument("command", choices=["probe"], help="Command to run")
    parser.add_argument("--json", action="store_true", help="Print JSON")
    parser.add_argument("--port", type=int, default=0, help="Use a known CDP port instead of scanning")
    parser.add_argument("--port-start", type=int, default=DEFAULT_PORT_START, help="First CDP port to scan")
    parser.add_argument("--port-end", type=int, default=DEFAULT_PORT_END, help="Last CDP port to scan")
    parser.add_argument("--url-contains", default=DEFAULT_URL_CONTAINS, help="Target URL substring")
    parser.add_argument("--text-limit", type=int, default=800, help="Body sample character limit")
    return parser


def main(argv: list[str] | None = None) -> int:
    configure_stdio()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.command == "probe":
            payload = probe_payload(args)
        else:
            raise ZiniaoCdpError(f"Unknown command: {args.command}")
    except Exception as exc:
        payload = {"ok": False, "error": str(exc)}
        if args.json:
            print(json.dumps(payload, ensure_ascii=False, indent=2))
        else:
            print(f"error: {exc}", file=sys.stderr)
        return 1
    if args.json:
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(f"ok: {payload['ok']}")
        print(f"browsers: {len(payload['browsers'])}")
        print(f"title: {payload['probe']['title']}")
        print(f"url: {payload['probe']['url']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
