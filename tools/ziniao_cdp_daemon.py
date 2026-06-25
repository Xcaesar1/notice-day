from __future__ import annotations

import argparse
import json
import logging
import sys
import time
import urllib.request
from pathlib import Path
from typing import Any

import frida


PROJECT_ROOT = Path(__file__).resolve().parents[1]
HOOK_PATH = Path(__file__).with_name("frida_port0.js")
DEFAULT_LOG_PATH = PROJECT_ROOT / ".local-state" / "account-health-notifier" / "ziniao-cdp-daemon.log"
DEFAULT_PORT_START = 9222
DEFAULT_PORT_END = 9250

SESSIONS: dict[int, frida.core.Session] = {}
SCRIPTS: dict[int, frida.core.Script] = {}


def configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="replace")


def setup_logging(log_path: Path) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    handlers: list[logging.Handler] = [
        logging.StreamHandler(),
        logging.FileHandler(log_path, encoding="utf-8"),
    ]
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=handlers,
    )


def load_hook_code() -> str:
    return HOOK_PATH.read_text(encoding="utf-8")


def hook_env_kit(pid: int, hook_code: str) -> bool:
    if pid in SESSIONS:
        return False
    try:
        session = frida.attach(pid)
        script = session.create_script(hook_code)

        def on_message(message: dict[str, Any], data: bytes | None) -> None:
            payload = message.get("payload") or message.get("description") or message.get("type")
            logging.info("env-kit PID=%s hook message=%s", pid, payload)

        script.on("message", on_message)
        script.load()
        SESSIONS[pid] = session
        SCRIPTS[pid] = script
        logging.info("hooked env-kit.exe PID=%s", pid)
        return True
    except Exception as exc:
        logging.warning("failed to hook env-kit.exe PID=%s error=%s", pid, exc)
        return False


def discover_cdp_ports(port_start: int, port_end: int) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for port in range(port_start, port_end + 1):
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=0.5) as response:
                version = json.loads(response.read().decode("utf-8", errors="replace"))
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=0.5) as response:
                targets = json.loads(response.read().decode("utf-8", errors="replace"))
        except Exception:
            continue
        pages = [
            {
                "type": item.get("type"),
                "title": item.get("title"),
                "url": item.get("url"),
            }
            for item in targets
            if isinstance(item, dict) and item.get("type") == "page"
        ]
        results.append(
            {
                "port": port,
                "browser": version.get("Browser"),
                "page_count": len(pages),
                "pages": pages,
            }
        )
    return results


def tick(args: argparse.Namespace, hook_code: str) -> dict[str, Any]:
    device = frida.get_local_device()
    processes = device.enumerate_processes()
    env_kits = [process for process in processes if process.name == "env-kit.exe"]
    current_pids = {process.pid for process in env_kits}
    for stale_pid in set(SESSIONS) - current_pids:
        try:
            SESSIONS[stale_pid].detach()
        except Exception:
            pass
        SESSIONS.pop(stale_pid, None)
        SCRIPTS.pop(stale_pid, None)
        logging.info("detached stale env-kit.exe PID=%s", stale_pid)
    hooked_now = []
    for process in env_kits:
        if hook_env_kit(process.pid, hook_code):
            hooked_now.append(process.pid)
    ports = discover_cdp_ports(args.port_start, args.port_end)
    if ports:
        logging.info("active CDP ports=%s", [item["port"] for item in ports])
    return {
        "ok": True,
        "env_kit_pids": sorted(current_pids),
        "hooked_pids": sorted(SESSIONS),
        "hooked_now": hooked_now,
        "cdp_ports": ports,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Keep Ziniao env-kit hooked so new store browsers expose CDP.")
    parser.add_argument("--interval-seconds", type=float, default=5.0)
    parser.add_argument("--port-start", type=int, default=DEFAULT_PORT_START)
    parser.add_argument("--port-end", type=int, default=DEFAULT_PORT_END)
    parser.add_argument("--log-file", default=str(DEFAULT_LOG_PATH))
    parser.add_argument("--once", action="store_true", help="Run one hook/discovery tick and exit")
    parser.add_argument("--json", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    configure_stdio()
    args = build_parser().parse_args(argv)
    setup_logging(Path(args.log_file))
    hook_code = load_hook_code()
    logging.info("Ziniao CDP daemon started hook=%s", HOOK_PATH)
    while True:
        try:
            result = tick(args, hook_code)
            if args.json:
                print(json.dumps(result, ensure_ascii=False, indent=2))
            if args.once:
                return 0 if result["ok"] else 1
            time.sleep(args.interval_seconds)
        except KeyboardInterrupt:
            logging.info("Ziniao CDP daemon stopped")
            return 0
        except Exception as exc:
            logging.error("daemon tick failed error=%s", exc)
            if args.once:
                if args.json:
                    print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2))
                return 1
            time.sleep(args.interval_seconds)


if __name__ == "__main__":
    raise SystemExit(main())
