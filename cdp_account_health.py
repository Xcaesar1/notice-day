from __future__ import annotations

import argparse
import json
import re
import time
import urllib.parse
from dataclasses import dataclass
from datetime import date
from typing import Any

import ziniao_cdp


PRODUCT_POLICIES_URL = "https://sellercentral.amazon.com/performance/account/health/product-policies"
DEFAULT_PAGE_SIZE = 25
DEFAULT_MAX_PAGES = 20


@dataclass(frozen=True)
class PolicyCategory:
    key: str
    label: str
    metric_name: str


POLICY_CATEGORIES: tuple[PolicyCategory, ...] = (
    PolicyCategory("safe", "食品和商品安全问题", "ProductSafety"),
    PolicyCategory("restricted", "违反受限商品政策", "RESTRICTED_PRODUCTS"),
    PolicyCategory("list", "上架政策违规", "ListingPolicy"),
    PolicyCategory("abuse", "违反买家商品评论政策", "PRODUCT_REVIEW_ABUSE"),
    PolicyCategory("auth", "商品真实性买家投诉", "ProductAuthenticity"),
    PolicyCategory("condition", "商品状况买家投诉", "ProductCondition"),
    PolicyCategory("intel", "知识产权投诉", "IntellectualProperty"),
    PolicyCategory("brand-protection", "涉嫌侵犯知识产权", "AUTOMATED_BRAND_PROTECTION"),
    PolicyCategory("regulatory-compliance", "监管合规性", "REGULATORY_COMPLIANCE"),
)


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = re.sub(r"<[^>]+>", " ", str(value).replace("\x00", " "))
    return re.sub(r"\s+", " ", text).strip()


def current_month_range(today: date | None = None) -> tuple[str, str]:
    today = today or date.today()
    return today.replace(day=1).isoformat(), today.isoformat()


def selected_categories(value: str = "all") -> list[PolicyCategory]:
    lookup = {item.key: item for item in POLICY_CATEGORIES}
    if not value or value == "all":
        return list(POLICY_CATEGORIES)
    selected: list[PolicyCategory] = []
    unknown: list[str] = []
    for raw in value.split(","):
        key = raw.strip()
        if not key:
            continue
        item = lookup.get(key)
        if item:
            selected.append(item)
        else:
            unknown.append(key)
    if unknown:
        raise ValueError(f"Unsupported policy categories: {', '.join(unknown)}")
    return selected


def find_target(args: argparse.Namespace) -> ziniao_cdp.CdpTarget:
    targets = ziniao_cdp.find_targets(
        url_contains=args.url_contains,
        port=args.port or None,
        port_start=args.port_start,
        port_end=args.port_end,
    )
    if not targets:
        raise ziniao_cdp.ZiniaoCdpError(f"{ziniao_cdp.STATUS_TARGET_GONE}: no Seller Central target found")
    return targets[0]


def navigate_to_product_policies(target: ziniao_cdp.CdpTarget, category_key: str = "safe", wait_seconds: float = 6.0) -> None:
    url = f"{PRODUCT_POLICIES_URL}?t={urllib.parse.quote(category_key)}"
    with ziniao_cdp.CdpSession(target.web_socket_debugger_url, timeout=20) as session:
        session.call("Page.enable")
        session.call("Page.navigate", {"url": url})
    time.sleep(max(wait_seconds, 0.0))


def page_context(target: ziniao_cdp.CdpTarget) -> dict[str, str]:
    expression = r"""
JSON.stringify((() => {
  const lines = (document.body ? document.body.innerText : '')
    .split(/\n+/).map(x => x.trim()).filter(Boolean);
  let store = '';
  let site = '';
  for (let i = 1; i < lines.length; i++) {
    if (/^(美国|United States|US)$/.test(lines[i]) && lines[i - 1] && !/账户状况|管理账户状况/.test(lines[i - 1])) {
      store = lines[i - 1];
      site = lines[i];
      break;
    }
  }
  return {title: document.title, url: location.href, store, site};
})())
"""
    data = ziniao_cdp.evaluate_json(target, expression, timeout=10)
    return {key: clean_text(value) for key, value in dict(data or {}).items()}


def policy_api_url(
    category: PolicyCategory,
    start_date: str,
    end_date: str,
    page_size: int,
    offset: int,
    next_page_token: str = "",
) -> str:
    params = {
        "metricNames": category.metric_name,
        "pageSize": str(page_size),
        "duration": "30",
        "offset": str(offset),
        "startDate": start_date,
        "endDate": end_date,
        "useCustomDateRange": "true",
        "nextPageToken": next_page_token,
        "statuses": "Open",
        "sortField": "CREATION_DATE",
        "sortByOrder": "DESC",
        "vendorCode": "",
        "searchValues": "",
        "searchMap": "{}",
        "policyGroups": "",
        "tags": "",
        "platform": "SELLER_CENTRAL",
        "excludeSkuDeleteIssues": "true",
    }
    return "/performance/api/product/policy/defects/pagination?" + urllib.parse.urlencode(params)


def fetch_policy_page(
    target: ziniao_cdp.CdpTarget,
    category: PolicyCategory,
    start_date: str,
    end_date: str,
    page_size: int,
    offset: int,
    next_page_token: str = "",
) -> dict[str, Any]:
    api = policy_api_url(category, start_date, end_date, page_size, offset, next_page_token)
    expression = f"""
(async () => {{
  const res = await fetch({json.dumps(api)}, {{
    method: 'POST',
    credentials: 'include',
    headers: {{'content-type': 'application/json'}},
    body: '{{}}'
  }});
  const text = await res.text();
  let body = null;
  try {{ body = JSON.parse(text); }} catch (err) {{ body = {{parseError: String(err), sample: text.slice(0, 500)}}; }}
  return JSON.stringify({{ok: res.ok, status: res.status, body}});
}})()
"""
    data = ziniao_cdp.evaluate_json(target, expression, timeout=30)
    if not data.get("ok"):
        raise ziniao_cdp.ZiniaoCdpError(
            f"policy API failed category={category.key} status={data.get('status')} body={clean_text(data.get('body'))[:200]}"
        )
    body = data.get("body")
    if not isinstance(body, dict):
        raise ziniao_cdp.ZiniaoCdpError(f"policy API returned non-object body for {category.key}")
    return body


def _list_value(value: Any) -> list[str]:
    if isinstance(value, list):
        return [clean_text(item) for item in value if clean_text(item)]
    text = clean_text(value)
    return [text] if text else []


def _issue_asins_skus(violation: dict[str, Any]) -> tuple[list[str], list[str]]:
    asins: list[str] = []
    skus: list[str] = []
    affected = violation.get("affectedEntity") if isinstance(violation.get("affectedEntity"), dict) else {}
    asins.extend(_list_value(affected.get("asins")))
    skus.extend(_list_value(affected.get("skus")))
    for issue in violation.get("issueList") or []:
        if not isinstance(issue, dict):
            continue
        source = issue.get("source") if isinstance(issue.get("source"), dict) else {}
        target = issue.get("target") if isinstance(issue.get("target"), dict) else {}
        params = issue.get("parameters") if isinstance(issue.get("parameters"), dict) else {}
        if source.get("target") == "ASIN":
            asins.extend(_list_value(source.get("artifactId")))
        if target.get("target") == "SKU":
            skus.extend(_list_value(target.get("artifactId")))
        asins.extend(_list_value(params.get("asin")))
    return list(dict.fromkeys(asins)), list(dict.fromkeys(skus))


def rows_from_violation(
    violation: dict[str, Any],
    category: PolicyCategory,
    store: str,
    site: str,
) -> list[dict[str, str]]:
    reason = violation.get("reason") if isinstance(violation.get("reason"), dict) else {}
    affected = violation.get("affectedEntity") if isinstance(violation.get("affectedEntity"), dict) else {}
    action_taken = violation.get("actionTaken") if isinstance(violation.get("actionTaken"), dict) else {}
    impact_date = violation.get("impactDate") if isinstance(violation.get("impactDate"), dict) else {}
    view_details = violation.get("viewDetails") if isinstance(violation.get("viewDetails"), dict) else {}
    detail_list = view_details.get("contentList") if isinstance(view_details.get("contentList"), list) else []
    asins, skus = _issue_asins_skus(violation)
    if not asins:
        asins = [""]
    if not skus:
        skus = [""]
    count = max(len(asins), len(skus))
    title = clean_text(affected.get("title"))
    rows: list[dict[str, str]] = []
    for index in range(count):
        asin = asins[index] if index < len(asins) else asins[0]
        sku = skus[index] if index < len(skus) else skus[0]
        impacted = "\n".join(part for part in [title, f"ASIN: {asin}" if asin else "", f"SKU: {sku}" if sku else ""] if part)
        rows.append(
            {
                "store": store,
                "site": site,
                "category": category.label,
                "asin": asin,
                "sku": sku,
                "reason": clean_text(reason.get("reason")) or category.label,
                "date": clean_text(impact_date.get("formattedDate")),
                "impacted_text": impacted,
                "sales_risk": clean_text(violation.get("gmsImpact")),
                "action": clean_text(action_taken.get("text")),
                "rating_impact": clean_text(violation.get("ahrImpact")),
                "detail": clean_text(detail_list[0]) if detail_list else "",
                "violation_id": clean_text(violation.get("violationId")),
            }
        )
    return rows


def collect_current_account_health(
    args: argparse.Namespace,
    start_date: str,
    end_date: str,
    categories: list[PolicyCategory],
    page_size: int = DEFAULT_PAGE_SIZE,
    max_pages: int = DEFAULT_MAX_PAGES,
    store_override: str = "",
    site_override: str = "",
) -> dict[str, Any]:
    target = find_target(args)
    navigate_to_product_policies(target, categories[0].key if categories else "safe")
    target = find_target(args)
    context = page_context(target)
    store = clean_text(store_override) or context.get("store") or "UNKNOWN_STORE"
    site = clean_text(site_override) or context.get("site") or "美国"
    rows: list[dict[str, str]] = []
    page_reports: list[dict[str, Any]] = []
    for category in categories:
        next_token = ""
        category_total = 0
        for page_index in range(max_pages):
            body = fetch_policy_page(
                target,
                category,
                start_date=start_date,
                end_date=end_date,
                page_size=page_size,
                offset=page_index * page_size,
                next_page_token=next_token,
            )
            violations = body.get("violations") or []
            if not isinstance(violations, list):
                violations = []
            for violation in violations:
                if isinstance(violation, dict):
                    rows.extend(rows_from_violation(violation, category, store, site))
            category_total += len(violations)
            page_reports.append(
                {
                    "category": category.key,
                    "metric_name": category.metric_name,
                    "page": page_index + 1,
                    "violations": len(violations),
                    "next_page_token": bool(body.get("nextPageToken")),
                }
            )
            next_token = clean_text(body.get("nextPageToken"))
            if not next_token or not violations:
                break
    return {
        "ok": True,
        "store": store,
        "site": site,
        "start_date": start_date,
        "end_date": end_date,
        "rows": rows,
        "row_count": len(rows),
        "target": target.safe_dict(),
        "page_context": context,
        "page_reports": page_reports,
    }
