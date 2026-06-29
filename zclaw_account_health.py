from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import time
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import cdp_account_health


PRODUCT_POLICIES_URL = cdp_account_health.PRODUCT_POLICIES_URL
DEFAULT_OPEN_TIMEOUT = 60
DEFAULT_NAV_TIMEOUT = 45
DEFAULT_EXEC_TIMEOUT = 45


class ZclawError(Exception):
    pass


@dataclass(frozen=True)
class Store:
    store_id: str
    store_name: str
    platform_name: str
    ip: str = ""


def clean_text(value: Any) -> str:
    return cdp_account_health.clean_text(value)


def current_month_range(today: date | None = None) -> tuple[str, str]:
    return cdp_account_health.current_month_range(today)


def selected_categories(value: str = "all") -> list[cdp_account_health.PolicyCategory]:
    return cdp_account_health.selected_categories(value)


def _run_cli(args: list[str], timeout: int = 30) -> dict[str, Any]:
    command = [*_ziniao_cli_command(), *args]
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=timeout,
    )
    stdout = (completed.stdout or "").strip()
    stderr = (completed.stderr or "").strip()
    if completed.returncode != 0:
        detail = stderr or stdout or f"exit={completed.returncode}"
        raise ZclawError(f"ziniao-cli {' '.join(args[:2])} failed: {detail[:800]}")
    if not stdout:
        return {"ok": True, "data": None}
    try:
        return json.loads(stdout)
    except json.JSONDecodeError:
        return {"ok": True, "data": stdout}


def _ziniao_cli_command() -> list[str]:
    node = Path(r"Q:\Node.js\node.exe")
    script = Path(r"Q:\Node.js\node_global\node_modules\@ziniao-open\cli\scripts\run.js")
    if node.is_file() and script.is_file():
        return [str(node), str(script)]
    for name in ("ziniao-cli.cmd", "ziniao-cli.exe", "ziniao-cli"):
        path = shutil.which(name)
        if path:
            return [path]
    fallback = r"Q:\Node.js\node_global\ziniao-cli.cmd"
    return [fallback if Path(fallback).is_file() else "ziniao-cli"]


def prepare_agent(timeout: int = 30) -> dict[str, Any]:
    return _run_cli(["store", "prepare-agent"], timeout=timeout)


def _list_stores_once(timeout: int = 30) -> list[Store]:
    payload = _run_cli(["store", "list", "--all", "--format", "json"], timeout=timeout)
    data = payload.get("data") if isinstance(payload, dict) else {}
    items = data.get("items") if isinstance(data, dict) else []
    stores: list[Store] = []
    for item in items if isinstance(items, list) else []:
        if not isinstance(item, dict):
            continue
        stores.append(
            Store(
                store_id=clean_text(item.get("storeId")),
                store_name=clean_text(item.get("storeName")),
                platform_name=clean_text(item.get("platformName")),
                ip=clean_text(item.get("ip")),
            )
        )
    return [store for store in stores if store.store_id and store.store_name]


def list_stores(timeout: int = 30) -> list[Store]:
    for attempt in range(2):
        try:
            prepare_agent(timeout=min(timeout, 30))
        except Exception:
            pass
        stores = _list_stores_once(timeout=timeout)
        if stores or attempt == 1:
            return stores
        time.sleep(1)
    return []


def filter_us_amazon_stores(stores: list[Store]) -> list[Store]:
    selected: list[Store] = []
    for store in stores:
        platform = store.platform_name.lower()
        if ("亚马逊" in store.platform_name or "amazon" in platform) and (
            "美国" in store.platform_name or "us" in platform or "united states" in platform
        ):
            selected.append(store)
    return selected


def _parse_store_filter(value: str) -> set[str]:
    return {clean_text(part).lower() for part in value.split(",") if clean_text(part)}


def select_stores(stores: list[Store], names_or_ids: str = "", limit: int = 0) -> list[Store]:
    selected = filter_us_amazon_stores(stores)
    filters = _parse_store_filter(names_or_ids)
    if filters:
        selected = [
            store
            for store in selected
            if store.store_name.lower() in filters or store.store_id.lower() in filters
        ]
    if limit > 0:
        selected = selected[:limit]
    return selected


def _open_store_once(store: Store, url: str = PRODUCT_POLICIES_URL, timeout: int = DEFAULT_OPEN_TIMEOUT) -> dict[str, Any]:
    args = ["store", "open", "--name", store.store_name, "--expected-name", store.store_name]
    if url:
        args.extend(["--url", url])
    try:
        return _run_cli(args, timeout=timeout)
    except ZclawError as name_exc:
        fallback = ["store", "open", "--id", store.store_id]
        if url:
            fallback.extend(["--url", url])
        try:
            return _run_cli(fallback, timeout=timeout)
        except ZclawError as id_exc:
            raise ZclawError(
                f"store open failed by name and id for {store.store_name}: "
                f"name={clean_text(name_exc)}; id={clean_text(id_exc)}"
            ) from id_exc


def open_store(store: Store, url: str = PRODUCT_POLICIES_URL, timeout: int = DEFAULT_OPEN_TIMEOUT) -> dict[str, Any]:
    last_error: Exception | None = None
    for attempt in range(2):
        if attempt:
            try:
                prepare_agent(timeout=min(timeout, 30))
            except Exception:
                pass
            time.sleep(1)
        try:
            return _open_store_once(store, url=url, timeout=timeout)
        except Exception as exc:
            last_error = exc
    if last_error:
        raise last_error
    raise ZclawError(f"store open failed for {store.store_name}")


def close_store(store: Store, timeout: int = 20) -> dict[str, Any]:
    return _run_cli(["store", "close", "--id", store.store_id], timeout=timeout)


def visit_page(
    store: Store,
    url: str = PRODUCT_POLICIES_URL,
    timeout: int = DEFAULT_NAV_TIMEOUT,
    target_id: str = "",
) -> dict[str, Any]:
    args = [
        "page",
        "visit",
        "--store-id",
        store.store_id,
        "--url",
        url,
        "--wait-until",
        "domcontentloaded",
        "--timeout",
        str(timeout * 1000),
    ]
    if target_id:
        args.extend(["--target-id", target_id])
    return _run_cli(args, timeout=timeout + 10)


def page_exec(store: Store, script: str, timeout: int = DEFAULT_EXEC_TIMEOUT, target_id: str = "") -> Any:
    args = [
        "page",
        "exec",
        "--store-id",
        store.store_id,
        "--script",
        script,
        "--timeout",
        str(timeout * 1000),
    ]
    if target_id:
        args.extend(["--target-id", target_id])
    payload = _run_cli(args, timeout=timeout + 10)
    data = payload.get("data") if isinstance(payload, dict) else {}
    inner = data.get("data") if isinstance(data, dict) else {}
    if not isinstance(inner, dict):
        return inner
    if inner.get("exceptionDetails"):
        raise ZclawError(f"page exec exception: {clean_text(inner.get('exceptionDetails'))[:800]}")
    result = inner.get("result")
    if isinstance(result, dict) and "value" in result:
        result = result.get("value")
    if isinstance(result, str):
        try:
            return json.loads(result)
        except json.JSONDecodeError:
            return result
    return result


def page_content(store: Store, timeout: int = 15, target_id: str = "") -> dict[str, Any]:
    args = [
        "page",
        "content",
        "--store-id",
        store.store_id,
        "--content-format",
        "text",
        "--timeout",
        str(timeout * 1000),
    ]
    if target_id:
        args.extend(["--target-id", target_id])
    payload = _run_cli(args, timeout=timeout + 10)
    data = payload.get("data") if isinstance(payload, dict) else {}
    inner = data.get("data") if isinstance(data, dict) else {}
    return inner if isinstance(inner, dict) else {}


def page_state(store: Store, text_limit: int = 1500, timeout: int = 10, target_id: str = "") -> dict[str, Any]:
    script = (
        "JSON.stringify({"
        "title:document.title||'',"
        "url:location.href||'',"
        "ready:document.readyState||'',"
        f"text:document.body?document.body.innerText.slice(0,{int(text_limit)}):''"
        "})"
    )
    data = page_exec(store, script, timeout=timeout, target_id=target_id)
    state = data if isinstance(data, dict) else {}
    if target_id:
        state["targetId"] = target_id
    return state


def _looks_like_second_verification(state: dict[str, Any]) -> bool:
    url = clean_text(state.get("url")).lower()
    text = clean_text(state.get("text"))
    if "sellercentral.amazon.com" in url and "账户状况" in text:
        return False
    has_continue = any(marker in text for marker in ("继续", "Continue", "CONTINUE"))
    has_challenge = any(
        marker in text
        for marker in (
            "验证码",
            "验证",
            "二次",
            "两步",
            "安全",
            "One Time Password",
            "Two-Step Verification",
            "Authentication required",
        )
    )
    return has_continue and (has_challenge or "amazon.com/ap/" in url or "signin" in url)


def _looks_like_account_switcher(state: dict[str, Any]) -> bool:
    text = clean_text(state.get("text"))
    url = clean_text(state.get("url")).lower()
    return (
        "amazon.com/ap/signin" in url
        and (
            ("切换账户" in text and "添加账户" in text)
            or ("Switch accounts" in text and "Add account" in text)
        )
    )


def _looks_like_password_prompt(state: dict[str, Any]) -> bool:
    text = clean_text(state.get("text"))
    url = clean_text(state.get("url")).lower()
    return (
        "amazon.com/ap/signin" in url
        and ("密码" in text or "Password" in text)
        and ("登录" in text or "Sign in" in text)
    )


def _looks_like_email_prompt(state: dict[str, Any]) -> bool:
    text = clean_text(state.get("text"))
    url = clean_text(state.get("url")).lower()
    return (
        "amazon.com/ap/signin" in url
        and (
            ("输入手机号码或邮箱" in text and "继续" in text)
            or ("Email or mobile phone number" in text and "Continue" in text)
        )
    )


def _looks_like_mfa_prompt(state: dict[str, Any]) -> bool:
    text = clean_text(state.get("text"))
    url = clean_text(state.get("url")).lower()
    return (
        "amazon.com/ap/mfa" in url
        or ("两步验证" in text and ("验证码" in text or "OTP" in text))
    )


def _is_seller_business_page(state: dict[str, Any]) -> bool:
    url = clean_text(state.get("url")).lower()
    if "sellercentral.amazon.com" not in url:
        return False
    if "/ap/signin" in url or "/ap/mfa" in url:
        return False
    return True


def page_click(store: Store, selector: str, timeout: int = 10, target_id: str = "") -> dict[str, Any]:
    args = [
        "page",
        "click",
        "--store-id",
        store.store_id,
        "--selector",
        selector,
        "--timeout",
        str(timeout * 1000),
    ]
    if target_id:
        args.extend(["--target-id", target_id])
    _run_cli(args, timeout=timeout + 10)
    return {"clicked": True, "selector": selector}


def click_existing_account_if_present(store: Store, timeout: int = 10, target_id: str = "") -> dict[str, Any]:
    script = r"""
(() => {
  const target = document.querySelector('a.cvf-widget-btn-verify-account-switcher');
  if (!target) return JSON.stringify({clicked:false, reason:'selector not found'});
  const text = (target.innerText || target.textContent || '').trim();
  target.dispatchEvent(new MouseEvent('mouseover', {bubbles:true, cancelable:true, view:window}));
  target.dispatchEvent(new MouseEvent('mousedown', {bubbles:true, cancelable:true, view:window}));
  target.dispatchEvent(new MouseEvent('mouseup', {bubbles:true, cancelable:true, view:window}));
  target.click();
  return JSON.stringify({clicked:true, text});
})()
"""
    data = page_exec(
        store,
        " ".join(line.strip() for line in script.splitlines() if line.strip()),
        timeout=timeout,
        target_id=target_id,
    )
    return data if isinstance(data, dict) else {"clicked": False, "result": data}


def click_email_continue_if_ready(store: Store, timeout: int = 10, target_id: str = "") -> dict[str, Any]:
    return page_click(store, "input#continue", timeout=timeout, target_id=target_id)


def click_password_login_if_ready(store: Store, timeout: int = 10, target_id: str = "") -> dict[str, Any]:
    return page_click(store, "input#signInSubmit", timeout=timeout, target_id=target_id)


def click_mfa_login_if_ready(store: Store, timeout: int = 10, target_id: str = "") -> dict[str, Any]:
    script = r"""
(() => {
  const otp = document.querySelector('#auth-mfa-otpcode,input[name=otpCode],input[type=tel]');
  const length = otp ? (otp.value || '').length : 0;
  if (length < 6) return JSON.stringify({clicked:false, reason:'otp not ready', otpLength:length});
  const btn = document.querySelector('#auth-signin-button') || Array.from(document.querySelectorAll('button,input[type=submit]')).find(el => /^(登录|Sign in|Continue|继续)$/i.test((el.innerText || el.value || '').trim()));
  if (!btn) return JSON.stringify({clicked:false, reason:'mfa submit not found', otpLength:length});
  btn.click();
  return JSON.stringify({clicked:true, otpLength:length, id:btn.id || ''});
})()
"""
    data = page_exec(
        store,
        " ".join(line.strip() for line in script.splitlines() if line.strip()),
        timeout=timeout,
        target_id=target_id,
    )
    return data if isinstance(data, dict) else {"clicked": False, "result": data}


def click_continue_if_present(store: Store, timeout: int = 10, target_id: str = "") -> dict[str, Any]:
    script = r"""
(() => {
  const nodes = Array.from(document.querySelectorAll('button,input[type=submit],input[type=button],a'));
  const visible = (el) => {
    const style = window.getComputedStyle(el);
    const rect = el.getBoundingClientRect();
    return style.visibility !== 'hidden' && style.display !== 'none' && rect.width > 0 && rect.height > 0;
  };
  const textOf = (el) => (el.innerText || el.value || el.getAttribute('aria-label') || el.textContent || '').trim();
  const target = nodes.find((el) => visible(el) && !el.disabled && /^(继续|Continue)$/i.test(textOf(el)));
  if (target) {
    const text = textOf(target);
    target.click();
    return JSON.stringify({clicked: true, text});
  }
  return JSON.stringify({
    clicked: false,
    candidates: nodes.map(textOf).filter(Boolean).slice(0, 20)
  });
})()
"""
    data = page_exec(
        store,
        " ".join(line.strip() for line in script.splitlines() if line.strip()),
        timeout=timeout,
        target_id=target_id,
    )
    return data if isinstance(data, dict) else {"clicked": False, "result": data}


def wait_for_seller_page_with_verification(store: Store, timeout: int = 90) -> dict[str, Any]:
    deadline = time.time() + max(timeout, 1)
    last: dict[str, Any] = {}
    clicked = False
    while time.time() < deadline:
        try:
            content = page_content(store, timeout=10)
            target_id = clean_text(content.get("targetId"))
            if content:
                state = {
                    "title": content.get("title", ""),
                    "url": content.get("url", ""),
                    "ready": "complete",
                    "text": content.get("content", ""),
                    "targetId": target_id,
                }
            else:
                state = page_state(store, timeout=10, target_id=target_id)
            if state:
                last = state
                if _looks_like_account_switcher(state):
                    click_result = click_existing_account_if_present(store, timeout=10, target_id=target_id)
                    clicked = clicked or bool(click_result.get("clicked"))
                    time.sleep(5)
                    continue
                if _looks_like_email_prompt(state):
                    click_result = click_email_continue_if_ready(store, timeout=10, target_id=target_id)
                    clicked = clicked or bool(click_result.get("clicked"))
                    time.sleep(5)
                    continue
                if _looks_like_password_prompt(state):
                    click_result = click_password_login_if_ready(store, timeout=10, target_id=target_id)
                    clicked = clicked or bool(click_result.get("clicked"))
                    time.sleep(5)
                    continue
                if _looks_like_mfa_prompt(state):
                    click_result = click_mfa_login_if_ready(store, timeout=10, target_id=target_id)
                    clicked = clicked or bool(click_result.get("clicked"))
                    time.sleep(5)
                    continue
                if _looks_like_second_verification(state):
                    click_result = click_continue_if_present(store, timeout=10, target_id=target_id)
                    clicked = clicked or bool(click_result.get("clicked"))
                    time.sleep(5)
                    continue
                if state.get("ready") == "complete" and _is_seller_business_page(state):
                    return state
        except Exception as exc:
            last = {"error": clean_text(exc)}
        time.sleep(2)
    suffix = " after clicking continue" if clicked else ""
    raise ZclawError(f"Seller Central page not ready{suffix} for {store.store_name}: {clean_text(last)[:500]}")


def wait_for_seller_page(store: Store, timeout: int = 30) -> dict[str, Any]:
    deadline = time.time() + max(timeout, 1)
    last: dict[str, Any] = {}
    while time.time() < deadline:
        try:
            content = page_content(store, timeout=10)
            if content:
                data = {
                    "title": content.get("title", ""),
                    "url": content.get("url", ""),
                    "ready": "complete",
                    "text": content.get("content", ""),
                    "targetId": content.get("targetId", ""),
                }
                last = data
                if _is_seller_business_page(data):
                    return data
        except Exception:
            pass
        time.sleep(2)
    raise ZclawError(f"Seller Central page not ready for {store.store_name}: {clean_text(last)[:500]}")


def fetch_policy_page(
    store: Store,
    category: cdp_account_health.PolicyCategory,
    start_date: str,
    end_date: str,
    page_size: int,
    offset: int,
    next_page_token: str = "",
    timeout: int = DEFAULT_EXEC_TIMEOUT,
    target_id: str = "",
) -> dict[str, Any]:
    api_path = cdp_account_health.policy_api_url(
        category,
        start_date=start_date,
        end_date=end_date,
        page_size=page_size,
        offset=offset,
        next_page_token=next_page_token,
    )
    script = (
        "(() => {"
        "const xhr=new XMLHttpRequest();"
        f"xhr.open('POST',{json.dumps(api_path)},false);"
        "xhr.setRequestHeader('content-type','application/json');"
        "let body=null;"
        "let thrown='';"
        "try{xhr.send('{}');}"
        "catch(err){thrown=String(err);}"
        "try{body=JSON.parse(xhr.responseText);}"
        "catch(err){body={parseError:String(err),sample:xhr.responseText.slice(0,500)};}"
        "return JSON.stringify({status:xhr.status,statusText:xhr.statusText,thrown,body});"
        "})()"
    )
    last_data: dict[str, Any] = {}
    for attempt in range(1, 6):
        data = page_exec(store, script, timeout=timeout, target_id=target_id)
        if not isinstance(data, dict):
            raise ZclawError(f"policy API returned non-object exec result for {store.store_name}/{category.key}")
        last_data = data
        status = int(data.get("status") or 0)
        body = data.get("body")
        if 200 <= status < 300 and isinstance(body, dict):
            return body
        if status == 0 and attempt < 5:
            if target_id:
                try:
                    wait_for_seller_page_with_verification(store, timeout=30)
                except Exception:
                    pass
            time.sleep(5)
            continue
        raise ZclawError(
            f"policy API failed store={store.store_name} category={category.key} status={status} "
            f"thrown={clean_text(data.get('thrown'))[:300]} body={clean_text(body)[:500]}"
        )
    raise ZclawError(f"policy API failed store={store.store_name} category={category.key}: {clean_text(last_data)[:500]}")


def _defects_from_body(body: dict[str, Any]) -> list[dict[str, Any]]:
    for key in ("violations", "defects"):
        value = body.get(key)
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
    return []


def collect_store_account_health(
    store: Store,
    start_date: str,
    end_date: str,
    categories: list[cdp_account_health.PolicyCategory],
    page_size: int = cdp_account_health.DEFAULT_PAGE_SIZE,
    max_pages: int = cdp_account_health.DEFAULT_MAX_PAGES,
    close_after: bool = True,
    open_timeout: int = DEFAULT_OPEN_TIMEOUT,
    nav_timeout: int = DEFAULT_NAV_TIMEOUT,
    exec_timeout: int = DEFAULT_EXEC_TIMEOUT,
) -> dict[str, Any]:
    started_at = time.strftime("%Y-%m-%d %H:%M:%S")
    rows: list[dict[str, str]] = []
    page_reports: list[dict[str, Any]] = []
    status = "success"
    error = ""
    try:
        open_store(store, timeout=open_timeout)
        time.sleep(2)
        visit_page(store, PRODUCT_POLICIES_URL, timeout=nav_timeout)
        state = wait_for_seller_page_with_verification(store, timeout=max(nav_timeout, 90))
        target_id = clean_text(state.get("targetId"))
        if "performance/account/health/product-policies" not in clean_text(state.get("url")):
            visit_page(store, PRODUCT_POLICIES_URL, timeout=nav_timeout, target_id=target_id)
            state = wait_for_seller_page(store, timeout=nav_timeout)
            target_id = clean_text(state.get("targetId")) or target_id
        for category in categories:
            next_token = ""
            for page_index in range(max_pages):
                body = fetch_policy_page(
                    store,
                    category,
                    start_date=start_date,
                    end_date=end_date,
                    page_size=page_size,
                    offset=page_index * page_size,
                    next_page_token=next_token,
                    timeout=exec_timeout,
                    target_id=target_id,
                )
                defects = _defects_from_body(body)
                for defect in defects:
                    rows.extend(cdp_account_health.rows_from_violation(defect, category, store.store_name, "美国"))
                page_reports.append(
                    {
                        "category": category.key,
                        "metric_name": category.metric_name,
                        "page": page_index + 1,
                        "defects": len(defects),
                        "next_page_token": bool(body.get("nextPageToken")),
                        "marketplace_id": clean_text(body.get("marketplaceId")),
                    }
                )
                next_token = clean_text(body.get("nextPageToken"))
                if not next_token or not defects:
                    break
    except Exception as exc:
        status = "failed"
        error = clean_text(exc)
    finally:
        if close_after:
            try:
                close_store(store)
            except Exception:
                pass
    ended_at = time.strftime("%Y-%m-%d %H:%M:%S")
    return {
        "ok": status == "success",
        "status": status,
        "store": store.store_name,
        "store_id": store.store_id,
        "site": "美国",
        "row_count": len(rows),
        "error": error,
        "target": {
            "port": "",
            "id": store.store_id,
            "type": "zclaw-store",
            "title": store.store_name,
            "url": PRODUCT_POLICIES_URL,
        },
        "page_reports": page_reports,
        "started_at": started_at,
        "ended_at": ended_at,
        "rows": rows,
    }


def collect_visible_account_health(
    start_date: str,
    end_date: str,
    categories: list[cdp_account_health.PolicyCategory],
    stores: str = "",
    limit: int = 0,
    page_size: int = cdp_account_health.DEFAULT_PAGE_SIZE,
    max_pages: int = cdp_account_health.DEFAULT_MAX_PAGES,
    close_after: bool = True,
    open_timeout: int = DEFAULT_OPEN_TIMEOUT,
    nav_timeout: int = DEFAULT_NAV_TIMEOUT,
    exec_timeout: int = DEFAULT_EXEC_TIMEOUT,
) -> dict[str, Any]:
    all_stores = list_stores()
    selected = select_stores(all_stores, names_or_ids=stores, limit=limit)
    target_results: list[dict[str, Any]] = []
    rows: list[dict[str, str]] = []
    for store in selected:
        result = collect_store_account_health(
            store,
            start_date=start_date,
            end_date=end_date,
            categories=categories,
            page_size=page_size,
            max_pages=max_pages,
            close_after=close_after,
            open_timeout=open_timeout,
            nav_timeout=nav_timeout,
            exec_timeout=exec_timeout,
        )
        target_results.append({key: value for key, value in result.items() if key != "rows"})
        rows.extend(result.get("rows", []))
    return {
        "ok": all(item.get("ok") for item in target_results) if target_results else False,
        "status": "success" if target_results and all(item.get("ok") for item in target_results) else "partial",
        "start_date": start_date,
        "end_date": end_date,
        "selected_store_count": len(selected),
        "selected_stores": [store.store_name for store in selected],
        "all_store_count": len(all_stores),
        "rows": rows,
        "row_count": len(rows),
        "target_results": target_results,
    }


def build_args(config: dict[str, Any], args: argparse.Namespace) -> dict[str, Any]:
    cdp_config = config.get("ziniao_cdp", {}) if isinstance(config, dict) else {}
    return {
        "page_size": int(getattr(args, "page_size", 0) or cdp_config.get("collect_page_size") or cdp_account_health.DEFAULT_PAGE_SIZE),
        "max_pages": int(getattr(args, "max_pages", 0) or cdp_config.get("collect_max_pages") or cdp_account_health.DEFAULT_MAX_PAGES),
        "open_timeout": int(getattr(args, "open_timeout", 0) or DEFAULT_OPEN_TIMEOUT),
        "nav_timeout": int(getattr(args, "nav_timeout", 0) or DEFAULT_NAV_TIMEOUT),
        "exec_timeout": int(getattr(args, "exec_timeout", 0) or DEFAULT_EXEC_TIMEOUT),
    }
