from __future__ import annotations

import argparse
import json
import os
import shutil
import sys
import tempfile
import time
import unittest
import uuid
from unittest import mock
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import account_health_notifier as notifier  # noqa: E402
import cdp_account_health  # noqa: E402
import ziniao_cdp  # noqa: E402


def write_single_sheet_xlsx(path: Path, sheet_name: str, headers: list[str], rows: list[dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<sheets><sheet name="{sheet_name}" sheetId="1" r:id="rId1"/></sheets>'
            "</workbook>",
        )
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            "</Relationships>",
        )
        archive.writestr("xl/worksheets/sheet1.xml", notifier._worksheet_xml(headers, rows))


class ProductExtractionTests(unittest.TestCase):
    def test_extracts_asin_and_sku_when_keywords_are_adjacent(self) -> None:
        text = "Product title ASIN：B0G7VV33CTSKU: HG-8901ORB2"

        products = notifier.extract_products(text)

        self.assertEqual(products, [{"asin": "B0G7VV33CT", "sku": "HG-8901ORB2"}])

    def test_placeholder_filter_text_does_not_create_product(self) -> None:
        row = {
            "店铺": "BYF",
            "站点": "美国",
            "异常分类": "违反受限商品政策",
            "原因": "政策合规性提供反馈",
            "日期": "重置 申请日期",
            "哪些商品会受到影响？": "ASIN、SKU、品牌或原因 - 开始输入以筛选",
            "存在销售风险": "",
            "采取的操作": "",
            "账户状况评级影响": "哪些商品会受到影响？",
        }

        items = notifier.item_from_detail_row(row, "test.xlsx")

        self.assertEqual(items, [])

    def test_extracts_asin_without_sku(self) -> None:
        text = "Restricted product ASIN：B0F6NCL5SH"

        products = notifier.extract_products(text)

        self.assertEqual(products, [{"asin": "B0F6NCL5SH", "sku": ""}])


class CoverageSummaryTests(unittest.TestCase):
    def test_summary_reports_missing_expected_stores(self) -> None:
        items = [
            notifier.ImpactItem(
                store="BYF",
                site="美国",
                category="食品和商品安全问题",
                asin="B0FPKWYF49",
                sku="BF-9901-BN",
                reason="reason",
                date="2026-01-30",
                impacted_text="ASIN：B0FPKWYF49 SKU: BF-9901-BN",
                sales_risk="none",
                action="removed",
                rating_impact="none",
            )
        ]

        summary = notifier.summarize_items(items, expected_stores=["BYF", "Hangoro"])

        self.assertFalse(summary["coverage_ok"])
        self.assertEqual(summary["expected_store_count"], 2)
        self.assertEqual(summary["store_count"], 1)
        self.assertEqual(summary["missing_stores"], ["Hangoro"])
        self.assertEqual(summary["missing_asin"], 0)


class ZiniaoCdpTests(unittest.TestCase):
    def test_find_targets_returns_only_matching_seller_pages(self) -> None:
        with mock.patch.object(
            ziniao_cdp,
            "scan_ports",
            return_value=[
                ziniao_cdp.CdpBrowser(
                    port=9222,
                    browser="Chrome/138.0.7204.252",
                    protocol_version="1.3",
                    web_socket_debugger_url="ws://127.0.0.1:9222/devtools/browser/x",
                )
            ],
        ), mock.patch.object(
            ziniao_cdp,
            "list_targets",
            return_value=[
                ziniao_cdp.CdpTarget(
                    port=9222,
                    id="seller",
                    type="page",
                    title="新的亚马逊销售体验",
                    url="https://sellercentral.amazon.com/amazonsell/business",
                    web_socket_debugger_url="ws://127.0.0.1:9222/devtools/page/seller",
                ),
                ziniao_cdp.CdpTarget(
                    port=9222,
                    id="extension",
                    type="page",
                    title="BYF|亚马逊-美国",
                    url="chrome-extension://dmgckiokdaggcmhfbagamdbkhflkdnah/index.html",
                    web_socket_debugger_url="ws://127.0.0.1:9222/devtools/page/extension",
                ),
            ],
        ):
            targets = ziniao_cdp.find_targets()

        self.assertEqual([target.id for target in targets], ["seller"])

    def test_cdp_smoke_uses_config_defaults(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config = notifier.default_config()
            config["ziniao_cdp"]["port"] = 9333
            config["ziniao_cdp"]["url_contains"] = "sellercentral.amazon.com"
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                port=0,
                port_start=0,
                port_end=0,
                url_contains="",
                text_limit=0,
            )
            with mock.patch.object(ziniao_cdp, "probe_payload", return_value={"ok": True, "probe": {}}) as mocked:
                result = notifier.execute_cdp_smoke(args)

        self.assertTrue(result["ok"])
        passed = mocked.call_args.args[0]
        self.assertEqual(passed.port, 9333)
        self.assertEqual(passed.url_contains, "sellercentral.amazon.com")

    def test_classifies_websocket_and_context_failures(self) -> None:
        self.assertEqual(ziniao_cdp.classify_error("WS_CLOSED: closed"), ziniao_cdp.STATUS_WS_CLOSED)
        self.assertEqual(
            ziniao_cdp.classify_error("Execution context was destroyed"),
            ziniao_cdp.STATUS_CONTEXT_LOST,
        )

    def test_cdp_doctor_uses_runtime_config(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config = notifier.default_config()
            config["ziniao_cdp"]["port_start"] = 9330
            config["ziniao_cdp"]["port_end"] = 9340
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                port=0,
                port_start=0,
                port_end=0,
                url_contains="",
                text_limit=0,
                include_body_sample=False,
            )
            with mock.patch.object(ziniao_cdp, "doctor_payload", return_value={"ok": True}) as mocked:
                result = notifier.execute_cdp_doctor(args)

        self.assertTrue(result["ok"])
        passed = mocked.call_args.args[0]
        self.assertEqual(passed.port_start, 9330)
        self.assertEqual(passed.port_end, 9340)

    def test_install_cdp_daemon_dry_run_builds_onlogon_task(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config = notifier.default_config()
            config["ziniao_cdp"]["daemon_task_name"] = "Test-ZiniaoCdpDaemon"
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                dry_run=True,
                task_name="",
                python=sys.executable,
                log_file="",
                port_start=0,
                port_end=0,
            )

            result = notifier.execute_install_cdp_daemon(args)

        self.assertTrue(result["ok"])
        self.assertEqual(result["status"], "dry_run")
        self.assertIn("ONLOGON", result["command"])
        self.assertIn("Test-ZiniaoCdpDaemon", result["command"])


class CdpAccountHealthCollectTests(unittest.TestCase):
    def test_ziniao_extension_context_reads_store_and_us_site_from_extension_title(self) -> None:
        with mock.patch.object(
            ziniao_cdp,
            "list_targets",
            return_value=[
                ziniao_cdp.CdpTarget(
                    port=9222,
                    id="extension",
                    type="page",
                    title="BYF|亚马逊-美国",
                    url="chrome-extension://example/index.html",
                    web_socket_debugger_url="ws://127.0.0.1:9222/devtools/page/extension",
                )
            ],
        ):
            context = cdp_account_health.ziniao_extension_context(9222)

        self.assertEqual(context, {"store": "BYF", "site": notifier.SITE_US})

    def test_rows_from_violation_extracts_asin_sku_and_business_fields(self) -> None:
        category = cdp_account_health.PolicyCategory("safe", "食品和商品安全问题", "ProductSafety")
        violation = {
            "reason": {"reason": "安全饮用水法案: 食品和商品安全问题"},
            "impactDate": {"formattedDate": "2026年6月10日"},
            "affectedEntity": {
                "title": "Bathroom Faucet",
                "asins": ["B0H4QD9255"],
                "skus": ["3314-BN-BASIC2"],
            },
            "gmsImpact": "过去 12 个月无销量",
            "actionTaken": {"text": "商品已移除"},
            "ahrImpact": "无影响",
            "viewDetails": {"contentList": ["需要提交安全饮用水法案文件。"]},
            "violationId": "V1",
        }

        rows = cdp_account_health.rows_from_violation(violation, category, "Soebiz", "美国")

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["store"], "Soebiz")
        self.assertEqual(rows[0]["asin"], "B0H4QD9255")
        self.assertEqual(rows[0]["sku"], "3314-BN-BASIC2")
        self.assertEqual(rows[0]["category"], "食品和商品安全问题")
        self.assertIn("Bathroom Faucet", rows[0]["impacted_text"])

    def test_execute_cdp_collect_current_writes_artifacts_from_mocked_collection(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config = notifier.default_config()
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                port=9222,
                port_start=0,
                port_end=0,
                url_contains="",
                text_limit=0,
                start_date="2026-06-01",
                end_date="2026-06-25",
                categories="safe",
                page_size=25,
                max_pages=2,
                store="",
                site="",
                output_dir="",
                include_items=True,
            )
            collection = {
                "ok": True,
                "store": "Soebiz",
                "site": "美国",
                "target": {"url": "https://sellercentral.amazon.com/performance/account/health/product-policies"},
                "page_reports": [{"category": "safe", "page": 1, "violations": 1}],
                "rows": [
                    {
                        "store": "Soebiz",
                        "site": "美国",
                        "category": "食品和商品安全问题",
                        "asin": "B0H4QD9255",
                        "sku": "3314-BN-BASIC2",
                        "reason": "安全饮用水法案",
                        "date": "2026年6月10日",
                        "impacted_text": "Bathroom Faucet\nASIN: B0H4QD9255\nSKU: 3314-BN-BASIC2",
                        "sales_risk": "过去 12 个月无销量",
                        "action": "商品已移除",
                        "rating_impact": "无影响",
                    }
                ],
            }
            with mock.patch.object(cdp_account_health, "collect_current_account_health", return_value=collection):
                result = notifier.execute_cdp_collect_current(args)

            self.assertTrue(result["ok"])
            self.assertEqual(result["row_count"], 1)
            self.assertTrue(Path(result["artifact"]).is_file())
            self.assertTrue(Path(result["json_artifact"]).is_file())
            self.assertEqual(result["items"][0]["ASIN"], "B0H4QD9255")

    def test_execute_cdp_collect_open_writes_target_results_and_missing_open_stores(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            store_list = root / "stores.xlsx"
            write_single_sheet_xlsx(
                store_list,
                "店铺执行清单",
                ["店铺", "站点"],
                [
                    {"店铺": "BYF", "站点": notifier.SITE_US},
                    {"店铺": "Hangoro", "站点": notifier.SITE_US},
                ],
            )
            config = notifier.default_config()
            config["source"]["store_list_path"] = str(store_list)
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                port=9222,
                port_start=0,
                port_end=0,
                url_contains="",
                text_limit=0,
                start_date="2026-06-01",
                end_date="2026-06-25",
                categories="safe",
                page_size=25,
                max_pages=2,
                site=notifier.SITE_US,
                store_list="",
                skip_store_list=False,
                require_all_open=False,
                output_dir="",
                include_items=True,
            )
            collection = {
                "ok": True,
                "status": "success",
                "start_date": "2026-06-01",
                "end_date": "2026-06-25",
                "target_count": 1,
                "rows": [
                    {
                        "store": "BYF",
                        "site": notifier.SITE_US,
                        "category": "食品和商品安全问题",
                        "asin": "B0FPKWYF49",
                        "sku": "BF-9901-BN",
                        "reason": "安全饮用水法案",
                        "date": "2026年6月10日",
                        "impacted_text": "Product\nASIN: B0FPKWYF49\nSKU: BF-9901-BN",
                        "sales_risk": "过去 12 个月无销量",
                        "action": "商品已移除",
                        "rating_impact": "无影响",
                    }
                ],
                "target_results": [
                    {
                        "ok": True,
                        "status": "success",
                        "store": "BYF",
                        "site": notifier.SITE_US,
                        "row_count": 1,
                        "error": "",
                        "target": {
                            "port": 9222,
                            "id": "seller",
                            "type": "page",
                            "title": "亚马逊",
                            "url": "https://sellercentral.amazon.com/performance/account/health/product-policies",
                            "web_socket_debugger_url": "ws://127.0.0.1:9222/devtools/page/seller",
                        },
                        "page_reports": [{"category": "safe", "page": 1, "violations": 1}],
                        "started_at": "2026-06-25 10:00:00",
                        "ended_at": "2026-06-25 10:00:05",
                    }
                ],
            }
            with mock.patch.object(cdp_account_health, "collect_open_account_health", return_value=collection):
                result = notifier.execute_cdp_collect_open(args)

            self.assertTrue(result["ok"])
            self.assertEqual(result["status"], "partial")
            self.assertEqual(result["row_count"], 1)
            self.assertEqual(result["opened_stores"], ["BYF"])
            self.assertEqual(result["missing_stores"], ["Hangoro"])
            self.assertNotIn("web_socket_debugger_url", result["target_results"][0]["target"])
            self.assertTrue(Path(result["artifact"]).is_file())
            self.assertTrue(Path(result["json_artifact"]).is_file())
            workbook = notifier.read_xlsx_workbook(Path(result["artifact"]))
            self.assertIn("Target Results", workbook)
            self.assertEqual(notifier.rows_to_dicts(workbook["Target Results"])[0]["store"], "BYF")


class ExcelEndToEndTests(unittest.TestCase):
    def test_parse_generated_excel_filters_site_dedupes_and_writes_summary(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_excel = root / "source.xlsx"
            store_list = root / "stores.xlsx"
            output_dir = root / "out"
            rows = [
                {
                    "店铺": "BYF",
                    "站点": notifier.SITE_US,
                    "异常分类": "食品和商品安全问题",
                    "ASIN": "",
                    "SKU": "",
                    "原因": "安全饮用水法案",
                    "日期": "2026-01-30",
                    "哪些商品会受到影响？": "Product A ASIN: B0FPKWYF49 SKU: BF-9901-BN Product B ASIN: B0G7VV33CTSKU: HG-8901ORB2",
                    "存在销售风险": "过去 12 个月无销量",
                    "采取的操作": "商品已移除",
                    "账户状况评级影响": "无影响",
                },
                {
                    "店铺": "BYF",
                    "站点": notifier.SITE_US,
                    "异常分类": "食品和商品安全问题",
                    "ASIN": "",
                    "SKU": "",
                    "原因": "安全饮用水法案",
                    "日期": "2026-01-30",
                    "哪些商品会受到影响？": "Product A ASIN: B0FPKWYF49 SKU: BF-9901-BN Product B ASIN: B0G7VV33CTSKU: HG-8901ORB2",
                    "存在销售风险": "过去 12 个月无销量",
                    "采取的操作": "商品已移除",
                    "账户状况评级影响": "无影响",
                },
                {
                    "店铺": "BYF",
                    "站点": "加拿大",
                    "异常分类": "食品和商品安全问题",
                    "原因": "不应纳入美国站",
                    "日期": "2026-01-30",
                    "哪些商品会受到影响？": "ASIN: B0F6NCL5SH SKU: CA-SKU",
                },
            ]
            detail_headers = [
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
            write_single_sheet_xlsx(source_excel, "异常明细", detail_headers, rows)
            write_single_sheet_xlsx(
                store_list,
                "店铺执行清单",
                ["店铺", "站点"],
                [
                    {"店铺": "BYF", "站点": notifier.SITE_US},
                    {"店铺": "Hangoro", "站点": notifier.SITE_US},
                    {"店铺": "CanadaShop", "站点": "加拿大"},
                ],
            )
            args = argparse.Namespace(
                config=str(root / "missing-config.json"),
                state_dir=str(root),
                source_type="excel",
                source_excel=str(source_excel),
                source_dir="",
                site=notifier.SITE_US,
                store_list=str(store_list),
                output_dir=str(output_dir),
                require_all_stores=True,
            )

            result = notifier.execute_parse(args)

            self.assertFalse(result["ok"])
            self.assertEqual(result["total_items"], 2)
            self.assertEqual(result["missing_stores"], ["Hangoro"])
            self.assertEqual(result["missing_asin"], 0)
            self.assertEqual(result["missing_sku"], 0)
            artifact = Path(result["artifact"])
            self.assertTrue(artifact.is_file())
            workbook = notifier.read_xlsx_workbook(artifact)
            detail_rows = notifier.rows_to_dicts(workbook["全店铺明细"])
            summary_rows = notifier.rows_to_dicts(workbook["解析汇总"])
            self.assertEqual(len(detail_rows), 2)
            self.assertIn("Hangoro", {row["名称"] for row in summary_rows if row["维度"] == "覆盖缺失店铺"})

    def test_find_latest_excel_ignores_generated_notifier_reports(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source = root / "sellercentral-account-health.xlsx"
            generated = root / "account-health-parse-20260625_120000_000000_deadbeef.xlsx"
            write_single_sheet_xlsx(source, "异常明细", ["店铺", "站点", "异常分类"], [])
            write_single_sheet_xlsx(generated, "全店铺明细", ["店铺", "站点", "异常分类"], [])
            os.utime(source, (1000, 1000))
            os.utime(generated, (2000, 2000))

            selected = notifier.find_latest_excel(root)

            self.assertEqual(selected, source)

    def test_parse_allows_stale_excel_for_diagnostics(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_excel = root / "source.xlsx"
            output_dir = root / "out"
            write_single_sheet_xlsx(
                source_excel,
                "异常明细",
                ["店铺", "站点", "异常分类", "原因", "日期", "哪些商品会受到影响？"],
                [
                    {
                        "店铺": "BYF",
                        "站点": notifier.SITE_US,
                        "异常分类": "食品和商品安全问题",
                        "原因": "安全饮用水法案",
                        "日期": "2026-01-30",
                        "哪些商品会受到影响？": "ASIN: B0STALE000 SKU: SKU-0000",
                    }
                ],
            )
            stale_time = time.time() - 7200
            os.utime(source_excel, (stale_time, stale_time))
            config = notifier.default_config()
            config["source"]["type"] = "excel"
            config["source"]["excel_path"] = str(source_excel)
            config["source"]["max_excel_age_hours"] = 1
            config["state"]["result_dir"] = str(output_dir)
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                source_type="excel",
                source_excel=str(source_excel),
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                output_dir=str(output_dir),
                require_all_stores=False,
            )

            result = notifier.execute_parse(args)

            self.assertTrue(result["ok"])
            self.assertEqual(result["status"], "success")
            self.assertEqual(result["total_items"], 1)

    def test_parse_missing_excel_returns_failure_report_without_exception(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            output_dir = root / "out"
            args = argparse.Namespace(
                config=str(root / "missing-config.json"),
                state_dir=str(root),
                source_type="excel",
                source_excel=str(root / "missing.xlsx"),
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                output_dir=str(output_dir),
                require_all_stores=False,
            )

            result = notifier.execute_parse(args)

            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "failed")
            self.assertEqual(result["total_items"], 0)
            self.assertIn("missing.xlsx", result["error"])
            self.assertTrue(Path(result["artifact"]).is_file())

    def test_parse_invalid_store_list_keeps_items_and_returns_failure_report(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_excel = root / "source.xlsx"
            store_list = root / "stores.xlsx"
            output_dir = root / "out"
            write_single_sheet_xlsx(
                source_excel,
                "异常明细",
                [
                    "店铺",
                    "站点",
                    "异常分类",
                    "原因",
                    "日期",
                    "哪些商品会受到影响？",
                ],
                [
                    {
                        "店铺": "BYF",
                        "站点": notifier.SITE_US,
                        "异常分类": "食品和商品安全问题",
                        "原因": "安全饮用水法案",
                        "日期": "2026-01-30",
                        "哪些商品会受到影响？": "ASIN: B0FPKWYF49 SKU: BF-9901-BN",
                    }
                ],
            )
            store_list.write_text("not an xlsx", encoding="utf-8")
            args = argparse.Namespace(
                config=str(root / "missing-config.json"),
                state_dir=str(root),
                source_type="excel",
                source_excel=str(source_excel),
                source_dir="",
                site=notifier.SITE_US,
                store_list=str(store_list),
                output_dir=str(output_dir),
                require_all_stores=True,
            )

            result = notifier.execute_parse(args)

            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "failed")
            self.assertEqual(result["total_items"], 1)
            self.assertIn("store_list", result["error"])
            workbook = notifier.read_xlsx_workbook(Path(result["artifact"]))
            detail_rows = notifier.rows_to_dicts(workbook["全店铺明细"])
            self.assertEqual(len(detail_rows), 1)


class RunIdTests(unittest.TestCase):
    def test_run_id_has_microsecond_precision(self) -> None:
        first = notifier.run_id_text()
        second = notifier.run_id_text()

        self.assertRegex(first, r"^\d{8}_\d{6}_\d{6}_[0-9a-f]{8}$")
        self.assertNotEqual(first, second)


class RunCoverageGuardTests(unittest.TestCase):
    def test_run_all_14_us_stores_allows_dry_run_without_marking_notified(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_excel = root / "source.xlsx"
            store_list = root / "stores.xlsx"
            dws_call = root / "dws.cmd"
            stores = [
                "BYF",
                "Hangoro",
                "Naukwan",
                "Winkear",
                "Wowkk",
                "Lexdale",
                "Kruzoo",
                "Taucet",
                "Soebiz",
                "Wintap",
                "Jabbol",
                "Hanallx",
                "Qinkell",
                "Artiqua",
            ]
            detail_rows = [
                {
                    "店铺": store,
                    "站点": notifier.SITE_US,
                    "异常分类": "食品和商品安全问题",
                    "原因": "安全饮用水法案",
                    "日期": "2026-01-30",
                    "哪些商品会受到影响？": f"ASIN: B0TEST{index:04d} SKU: SKU-{index:04d}",
                }
                for index, store in enumerate(stores)
            ]
            write_single_sheet_xlsx(
                source_excel,
                "异常明细",
                ["店铺", "站点", "异常分类", "原因", "日期", "哪些商品会受到影响？"],
                detail_rows,
            )
            write_single_sheet_xlsx(
                store_list,
                "店铺执行清单",
                ["店铺", "站点"],
                [{"店铺": store, "站点": notifier.SITE_US} for store in stores]
                + [{"店铺": "CanadaShop", "站点": "加拿大"}],
            )
            dws_call.write_text("@echo off\r\necho {\"ok\":true}\r\nexit /b 0\r\n", encoding="utf-8")
            config = notifier.default_config()
            config["source"]["type"] = "excel"
            config["source"]["excel_path"] = str(source_excel)
            config["source"]["store_list_path"] = str(store_list)
            config["dingtalk"]["dws_call"] = str(dws_call)
            config["dingtalk"]["send_enabled"] = False
            config["notify"]["require_all_stores_before_send"] = True
            config["notify"]["max_items_per_message"] = 20
            config["state"]["db_path"] = str(root / "state.sqlite")
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                source_type="",
                source_excel="",
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                require_all_stores=False,
                skip_store_coverage=False,
                dry_run=True,
                send=False,
            )

            result = notifier.execute_run(args)
            second = notifier.execute_run(args)

            self.assertTrue(result["ok"])
            self.assertEqual(result["status"], "dry_run")
            self.assertEqual(result["total_items"], 14)
            self.assertEqual(result["notify_candidates"], 14)
            self.assertEqual(result["sent_items"], 0)
            self.assertTrue(result["coverage"]["coverage_ok"])
            self.assertEqual(result["coverage"]["expected_store_count"], 14)
            self.assertEqual(result["coverage"]["missing_stores"], [])
            self.assertEqual(second["notify_candidates"], 14)
            conn = notifier.connect_db(root / "state.sqlite")
            try:
                dry_run_attempts = conn.execute(
                    "SELECT COUNT(*) FROM notification_attempts WHERE status = 'dry_run' AND item_count = 14"
                ).fetchone()[0]
                notified_count = conn.execute("SELECT COUNT(*) FROM notified_items").fetchone()[0]
            finally:
                conn.close()
            self.assertEqual(dry_run_attempts, 2)
            self.assertEqual(notified_count, 0)

    def test_run_chunks_full_store_dry_run_without_marking_notified(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_excel = root / "source.xlsx"
            store_list = root / "stores.xlsx"
            dws_call = root / "dws.cmd"
            stores = [
                "BYF",
                "Hangoro",
                "Naukwan",
                "Winkear",
                "Wowkk",
                "Lexdale",
                "Kruzoo",
                "Taucet",
                "Soebiz",
                "Wintap",
                "Jabbol",
                "Hanallx",
                "Qinkell",
                "Artiqua",
            ]
            detail_rows = [
                {
                    "店铺": store,
                    "站点": notifier.SITE_US,
                    "异常分类": "食品和商品安全问题",
                    "原因": "安全饮用水法案",
                    "日期": "2026-01-30",
                    "哪些商品会受到影响？": f"ASIN: B0TEST{index:04d} SKU: SKU-{index:04d}",
                }
                for index, store in enumerate(stores)
            ]
            write_single_sheet_xlsx(
                source_excel,
                "异常明细",
                ["店铺", "站点", "异常分类", "原因", "日期", "哪些商品会受到影响？"],
                detail_rows,
            )
            write_single_sheet_xlsx(
                store_list,
                "店铺执行清单",
                ["店铺", "站点"],
                [{"店铺": store, "站点": notifier.SITE_US} for store in stores]
                + [{"店铺": "CanadaShop", "站点": "加拿大"}],
            )
            dws_call.write_text("@echo off\r\necho {\"ok\":true}\r\nexit /b 0\r\n", encoding="utf-8")
            config = notifier.default_config()
            config["source"]["type"] = "excel"
            config["source"]["excel_path"] = str(source_excel)
            config["source"]["store_list_path"] = str(store_list)
            config["dingtalk"]["dws_call"] = str(dws_call)
            config["dingtalk"]["send_enabled"] = False
            config["notify"]["require_all_stores_before_send"] = True
            config["notify"]["max_items_per_message"] = 5
            config["state"]["db_path"] = str(root / "state.sqlite")
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                source_type="",
                source_excel="",
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                require_all_stores=False,
                skip_store_coverage=False,
                dry_run=True,
                send=False,
            )

            result = notifier.execute_run(args)

            self.assertTrue(result["ok"])
            self.assertEqual(result["status"], "dry_run")
            self.assertEqual(result["total_items"], 14)
            self.assertEqual(result["notify_candidates"], 14)
            self.assertEqual(result["sent_items"], 0)
            self.assertTrue(result["coverage"]["coverage_ok"])
            conn = notifier.connect_db(root / "state.sqlite")
            try:
                attempts = conn.execute(
                    "SELECT item_count, status, dry_run FROM notification_attempts ORDER BY id"
                ).fetchall()
                notified_count = conn.execute("SELECT COUNT(*) FROM notified_items").fetchone()[0]
            finally:
                conn.close()
            self.assertEqual([row["item_count"] for row in attempts], [5, 5, 4])
            self.assertEqual([row["status"] for row in attempts], ["dry_run", "dry_run", "dry_run"])
            self.assertEqual([row["dry_run"] for row in attempts], [1, 1, 1])
            self.assertEqual(notified_count, 0)

    def test_run_blocks_notification_when_product_ids_are_incomplete(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_excel = root / "source.xlsx"
            store_list = root / "stores.xlsx"
            dws_call = root / "dws.cmd"
            stores = [
                "BYF",
                "Hangoro",
                "Naukwan",
                "Winkear",
                "Wowkk",
                "Lexdale",
                "Kruzoo",
                "Taucet",
                "Soebiz",
                "Wintap",
                "Jabbol",
                "Hanallx",
                "Qinkell",
                "Artiqua",
            ]
            detail_rows = []
            for index, store in enumerate(stores):
                impacted = f"ASIN: B0MISS{index:04d}"
                if index != 0:
                    impacted += f" SKU: SKU-{index:04d}"
                detail_rows.append(
                    {
                        "店铺": store,
                        "站点": notifier.SITE_US,
                        "异常分类": "食品和商品安全问题",
                        "原因": "安全饮用水法案",
                        "日期": "2026-01-30",
                        "哪些商品会受到影响？": impacted,
                    }
                )
            write_single_sheet_xlsx(
                source_excel,
                "异常明细",
                ["店铺", "站点", "异常分类", "原因", "日期", "哪些商品会受到影响？"],
                detail_rows,
            )
            write_single_sheet_xlsx(
                store_list,
                "店铺执行清单",
                ["店铺", "站点"],
                [{"店铺": store, "站点": notifier.SITE_US} for store in stores],
            )
            dws_call.write_text("@echo off\r\necho {\"ok\":true}\r\nexit /b 0\r\n", encoding="utf-8")
            config = notifier.default_config()
            config["source"]["type"] = "excel"
            config["source"]["excel_path"] = str(source_excel)
            config["source"]["store_list_path"] = str(store_list)
            config["dingtalk"]["dws_call"] = str(dws_call)
            config["dingtalk"]["send_enabled"] = False
            config["notify"]["require_all_stores_before_send"] = True
            config["state"]["db_path"] = str(root / "state.sqlite")
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                source_type="",
                source_excel="",
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                require_all_stores=False,
                skip_store_coverage=False,
                dry_run=True,
                send=False,
            )

            result = notifier.execute_run(args)

            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "quality_failed")
            self.assertEqual(result["notify_candidates"], 0)
            self.assertTrue(result["coverage"]["coverage_ok"])
            self.assertEqual(result["coverage"]["missing_sku"], 1)
            self.assertIn("SKU", result["error"])
            conn = notifier.connect_db(root / "state.sqlite")
            try:
                attempt_count = conn.execute("SELECT COUNT(*) FROM notification_attempts").fetchone()[0]
                notified_count = conn.execute("SELECT COUNT(*) FROM notified_items").fetchone()[0]
            finally:
                conn.close()
            self.assertEqual(attempt_count, 0)
            self.assertEqual(notified_count, 0)

    def test_run_blocks_notification_when_excel_source_is_stale(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            source_excel = root / "source.xlsx"
            store_list = root / "stores.xlsx"
            dws_call = root / "dws.cmd"
            stores = [
                "BYF",
                "Hangoro",
                "Naukwan",
                "Winkear",
                "Wowkk",
                "Lexdale",
                "Kruzoo",
                "Taucet",
                "Soebiz",
                "Wintap",
                "Jabbol",
                "Hanallx",
                "Qinkell",
                "Artiqua",
            ]
            detail_rows = [
                {
                    "店铺": store,
                    "站点": notifier.SITE_US,
                    "异常分类": "食品和商品安全问题",
                    "原因": "安全饮用水法案",
                    "日期": "2026-01-30",
                    "哪些商品会受到影响？": f"ASIN: B0STALE{index:03d} SKU: SKU-{index:04d}",
                }
                for index, store in enumerate(stores)
            ]
            write_single_sheet_xlsx(
                source_excel,
                "异常明细",
                ["店铺", "站点", "异常分类", "原因", "日期", "哪些商品会受到影响？"],
                detail_rows,
            )
            stale_time = time.time() - 7200
            os.utime(source_excel, (stale_time, stale_time))
            write_single_sheet_xlsx(
                store_list,
                "店铺执行清单",
                ["店铺", "站点"],
                [{"店铺": store, "站点": notifier.SITE_US} for store in stores],
            )
            dws_call.write_text("@echo off\r\necho {\"ok\":true}\r\nexit /b 0\r\n", encoding="utf-8")
            config = notifier.default_config()
            config["source"]["type"] = "excel"
            config["source"]["excel_path"] = str(source_excel)
            config["source"]["store_list_path"] = str(store_list)
            config["source"]["max_excel_age_hours"] = 1
            config["dingtalk"]["dws_call"] = str(dws_call)
            config["dingtalk"]["send_enabled"] = False
            config["notify"]["require_all_stores_before_send"] = True
            config["state"]["db_path"] = str(root / "state.sqlite")
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                source_type="",
                source_excel="",
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                require_all_stores=False,
                skip_store_coverage=False,
                dry_run=True,
                send=False,
            )

            result = notifier.execute_run(args)

            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "source_stale")
            self.assertEqual(result["notify_candidates"], 0)
            self.assertIn("Excel", result["error"])
            self.assertIn("过期", result["error"])
            conn = notifier.connect_db(root / "state.sqlite")
            try:
                attempt_count = conn.execute("SELECT COUNT(*) FROM notification_attempts").fetchone()[0]
                notified_count = conn.execute("SELECT COUNT(*) FROM notified_items").fetchone()[0]
            finally:
                conn.close()
            self.assertEqual(attempt_count, 0)
            self.assertEqual(notified_count, 0)

    def test_run_skips_when_another_run_lock_is_active(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            lock_path = root / "run.lock"
            lock_path.write_text('{"run_id":"previous"}', encoding="utf-8")
            dws_call = root / "dws.cmd"
            dws_call.write_text("@echo off\r\nexit /b 1\r\n", encoding="utf-8")
            config = notifier.default_config()
            config["source"]["type"] = "sample"
            config["dingtalk"]["dws_call"] = str(dws_call)
            config["dingtalk"]["send_enabled"] = False
            config["notify"]["require_all_stores_before_send"] = False
            config["state"]["db_path"] = str(root / "state.sqlite")
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                source_type="",
                source_excel="",
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                require_all_stores=False,
                skip_store_coverage=True,
                dry_run=True,
                send=False,
            )

            result = notifier.execute_run(args)

            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "run_locked")
            self.assertEqual(result["notify_candidates"], 0)
            self.assertIn("run.lock", result["error"])
            self.assertTrue(lock_path.is_file())
            self.assertFalse((root / "messages").exists())
            self.assertTrue(Path(result["artifact"]).is_file())
            conn = notifier.connect_db(root / "state.sqlite")
            try:
                attempt_count = conn.execute("SELECT COUNT(*) FROM notification_attempts").fetchone()[0]
                notified_count = conn.execute("SELECT COUNT(*) FROM notified_items").fetchone()[0]
            finally:
                conn.close()
            self.assertEqual(attempt_count, 0)
            self.assertEqual(notified_count, 0)

    def test_run_blocks_notification_when_store_coverage_is_incomplete(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            store_list = root / "stores.xlsx"
            notifier.write_parse_xlsx(
                store_list,
                [
                    {"店铺": "BYF", "站点": notifier.SITE_US},
                    {"店铺": "Hangoro", "站点": notifier.SITE_US},
                ],
                {},
            )
            config = notifier.default_config()
            config["source"]["type"] = "sample"
            config["source"]["store_list_path"] = str(store_list)
            config["dingtalk"]["dws_call"] = str(root / "missing-dws.cmd")
            config["dingtalk"]["send_enabled"] = True
            config["notify"]["require_all_stores_before_send"] = True
            config["state"]["db_path"] = str(root / "state.sqlite")
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                source_type="",
                source_excel="",
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                require_all_stores=True,
                skip_store_coverage=False,
                dry_run=False,
                send=True,
            )

            result = notifier.execute_run(args)

            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "coverage_failed")
            self.assertEqual(result["notify_candidates"], 0)
            self.assertIn("Hangoro", result["error"])

    def test_run_reports_invalid_retention_without_exception(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config = notifier.default_config()
            config["source"]["type"] = "sample"
            config["dingtalk"]["send_enabled"] = False
            config["notify"]["require_all_stores_before_send"] = False
            config["notify"]["dedupe_retention_days"] = "abc"
            config["state"]["db_path"] = str(root / "state.sqlite")
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                source_type="",
                source_excel="",
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                require_all_stores=False,
                skip_store_coverage=True,
                dry_run=True,
                send=False,
            )

            result = notifier.execute_run(args)

            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "failed")
            self.assertIn("dedupe_retention_days", result["error"])
            self.assertEqual(result["notify_candidates"], 0)
            self.assertTrue(Path(result["artifact"]).is_file())

    def test_run_reports_invalid_max_items_without_exception(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config = notifier.default_config()
            config["source"]["type"] = "sample"
            config["dingtalk"]["send_enabled"] = False
            config["notify"]["require_all_stores_before_send"] = False
            config["notify"]["max_items_per_message"] = "abc"
            config["state"]["db_path"] = str(root / "state.sqlite")
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                source_type="",
                source_excel="",
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                require_all_stores=False,
                skip_store_coverage=True,
                dry_run=True,
                send=False,
            )

            result = notifier.execute_run(args)

            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "failed")
            self.assertIn("max_items_per_message", result["error"])
            self.assertEqual(result["notify_candidates"], 0)
            self.assertTrue(Path(result["artifact"]).is_file())
            conn = notifier.connect_db(root / "state.sqlite")
            try:
                attempt_count = conn.execute("SELECT COUNT(*) FROM notification_attempts").fetchone()[0]
                notified_count = conn.execute("SELECT COUNT(*) FROM notified_items").fetchone()[0]
            finally:
                conn.close()
            self.assertEqual(attempt_count, 0)
            self.assertEqual(notified_count, 0)


class ConfigValidationTests(unittest.TestCase):
    def _base_config_path(self, root: Path) -> Path:
        source_dir = root / "source"
        source_dir.mkdir()
        store_list = root / "stores.xlsx"
        dws_call = root / "dws.cmd"
        dws_call.write_text("@echo off\r\necho {}\r\nexit /b 0\r\n", encoding="utf-8")
        write_single_sheet_xlsx(
            store_list,
            "搴楅摵鎵ц娓呭崟",
            ["搴楅摵", "绔欑偣"],
            [{"搴楅摵": "BYF", "绔欑偣": notifier.SITE_US}],
        )
        config = notifier.default_config()
        config["source"]["excel_dir"] = str(source_dir)
        config["source"]["store_list_path"] = str(store_list)
        config["dingtalk"]["dws_call"] = str(dws_call)
        config["dingtalk"]["send_enabled"] = False
        config["state"]["db_path"] = str(root / "state.sqlite")
        config["state"]["result_dir"] = str(root / "runs")
        config_path = root / "config.json"
        config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
        return config_path

    def test_validate_config_requires_robot_group_and_send_enabled_for_send_ready(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config_path = self._base_config_path(root)
            args = argparse.Namespace(config=str(config_path), state_dir=str(root), require_send_ready=True)

            result = notifier.execute_validate_config(args)

            codes = {issue["code"] for issue in result["issues"]}
            self.assertFalse(result["ok"])
            self.assertIn("robot_code_missing", codes)
            self.assertIn("group_open_conversation_id_missing", codes)
            self.assertIn("send_enabled_false", codes)

    def test_validate_config_reports_missing_config_without_exception(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            args = argparse.Namespace(config=str(root / "missing.json"), state_dir=str(root), require_send_ready=False)

            result = notifier.execute_validate_config(args)

            self.assertFalse(result["ok"])
            self.assertIn("config_missing", {issue["code"] for issue in result["issues"]})

    def test_validate_config_reports_invalid_excel_age_limit(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config_path = self._base_config_path(root)
            config = json.loads(config_path.read_text(encoding="utf-8"))
            config["source"]["max_excel_age_hours"] = -1
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(config=str(config_path), state_dir=str(root), require_send_ready=False)

            result = notifier.execute_validate_config(args)

            self.assertFalse(result["ok"])
            self.assertIn("source_max_excel_age_hours_invalid", {issue["code"] for issue in result["issues"]})

    def test_validate_config_reports_invalid_run_lock_ttl(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config_path = self._base_config_path(root)
            config = json.loads(config_path.read_text(encoding="utf-8"))
            config["state"]["run_lock_ttl_minutes"] = 0
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(config=str(config_path), state_dir=str(root), require_send_ready=False)

            result = notifier.execute_validate_config(args)

            self.assertFalse(result["ok"])
            self.assertIn("run_lock_ttl_minutes_invalid", {issue["code"] for issue in result["issues"]})

    def test_install_schedule_dry_run_refuses_incomplete_send_config(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config_path = self._base_config_path(root)
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                dry_run=True,
                task_name="",
                interval_hours="",
                python="",
                skip_preflight=False,
            )

            result = notifier.execute_install_schedule(args)

            self.assertFalse(result["ok"])
            self.assertTrue(result["dry_run"])
            self.assertEqual(result["status"], "preflight_failed")
            self.assertIn("command", result)
            self.assertIn("robot_code_missing", {issue["code"] for issue in result["issues"]})

    def test_install_schedule_reports_invalid_interval_without_exception(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config_path = self._base_config_path(root)
            config = json.loads(config_path.read_text(encoding="utf-8"))
            config["dingtalk"]["robot_code"] = "robot"
            config["dingtalk"]["group_open_conversation_id"] = "group"
            config["dingtalk"]["send_enabled"] = True
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                dry_run=True,
                task_name="",
                interval_hours="abc",
                python="",
                skip_preflight=False,
            )

            result = notifier.execute_install_schedule(args)

            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "preflight_failed")
            self.assertIn("interval_hours_invalid", {issue["code"] for issue in result["issues"]})

    def test_send_test_send_requires_send_ready_config_without_calling_dws(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config_path = self._base_config_path(root)
            args = argparse.Namespace(config=str(config_path), state_dir=str(root), send=True)

            result = notifier.execute_send_test(args)

            codes = {issue["code"] for issue in result["issues"]}
            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "preflight_failed")
            self.assertFalse(result["dry_run"])
            self.assertIn("robot_code_missing", codes)
            self.assertIn("group_open_conversation_id_missing", codes)
            self.assertIn("send_enabled_false", codes)
            self.assertFalse((root / "messages").exists())


class RuntimeFileSafetyTests(unittest.TestCase):
    def test_run_lock_is_released_when_artifact_write_fails(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config = notifier.default_config()
            config["source"]["type"] = "sample"
            config["dingtalk"]["send_enabled"] = False
            config["notify"]["require_all_stores_before_send"] = False
            config["state"]["db_path"] = str(root / "state.sqlite")
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                source_type="",
                source_excel="",
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                require_all_stores=False,
                skip_store_coverage=True,
                dry_run=True,
                send=False,
            )
            original_update = notifier.update_run_artifacts

            def raise_artifact_error(*_args, **_kwargs):
                raise RuntimeError("artifact write failed")

            try:
                notifier.update_run_artifacts = raise_artifact_error  # type: ignore[assignment]
                with self.assertRaises(RuntimeError):
                    notifier.execute_run(args)
            finally:
                notifier.update_run_artifacts = original_update  # type: ignore[assignment]

            self.assertFalse((root / "run.lock").exists())

    def test_run_dws_call_uses_unique_argument_files_within_same_millisecond(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            dws_call = root / "dws.cmd"
            dws_call.write_text("@echo off\r\nexit /b 0\r\n", encoding="utf-8")
            original_time = notifier.time.time
            try:
                notifier.time.time = lambda: 123456.789  # type: ignore[assignment]
                notifier.run_dws_call(dws_call, ["first"], root)
                notifier.run_dws_call(dws_call, ["second"], root)
            finally:
                notifier.time.time = original_time  # type: ignore[assignment]

            arg_files = sorted((root / "dws-args").glob("args-*.json"))
            payloads = [json.loads(path.read_text(encoding="utf-8"))["args"][0] for path in arg_files]
            self.assertEqual(len(arg_files), 2)
            self.assertEqual(payloads, ["first", "second"])

    def test_send_dingtalk_uses_unique_message_files_within_same_millisecond(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            dws_call = root / "dws.cmd"
            dws_call.write_text("@echo off\r\nexit /b 0\r\n", encoding="utf-8")
            config = notifier.default_config()
            config["dingtalk"]["dws_call"] = str(dws_call)
            original_time = notifier.time.time
            try:
                notifier.time.time = lambda: 123456.789  # type: ignore[assignment]
                notifier.send_dingtalk_markdown(config, root, "t1", "first message", dry_run=True)
                notifier.send_dingtalk_markdown(config, root, "t2", "second message", dry_run=True)
            finally:
                notifier.time.time = original_time  # type: ignore[assignment]

            message_files = sorted((root / "messages").glob("message-*.md"))
            payloads = [path.read_text(encoding="utf-8") for path in message_files]
            self.assertEqual(len(message_files), 2)
            self.assertEqual(payloads, ["first message", "second message"])

    def test_self_test_can_run_twice_without_leaving_locked_state(self) -> None:
        root = notifier.DEFAULT_STATE_DIR / f"self-test-unit-{uuid.uuid4().hex}"
        args = argparse.Namespace(state_dir=str(root))

        try:
            first = notifier.execute_self_test(args)
            second = notifier.execute_self_test(args)

            self.assertTrue(first["ok"])
            self.assertTrue(second["ok"])
            self.assertTrue((root / "state.sqlite").is_file())
        finally:
            if root.exists():
                shutil.rmtree(root)


class SendFailureTests(unittest.TestCase):
    def _write_config(self, root: Path, dws_call: Path) -> Path:
        config = notifier.default_config()
        config["source"]["type"] = "sample"
        config["dingtalk"]["dws_call"] = str(dws_call)
        config["dingtalk"]["robot_code"] = "robot"
        config["dingtalk"]["group_open_conversation_id"] = "group"
        config["dingtalk"]["send_enabled"] = True
        config["notify"]["require_all_stores_before_send"] = False
        config["state"]["db_path"] = str(root / "state.sqlite")
        config["state"]["result_dir"] = str(root / "runs")
        config_path = root / "config.json"
        config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
        return config_path

    def _run_args(self, root: Path, config_path: Path, *, dry_run: bool, send: bool) -> argparse.Namespace:
        return argparse.Namespace(
            config=str(config_path),
            state_dir=str(root),
            source_type="",
            source_excel="",
            source_dir="",
            site=notifier.SITE_US,
            store_list="",
            require_all_stores=False,
            skip_store_coverage=True,
            dry_run=dry_run,
            send=send,
        )

    def test_dry_run_dws_failure_is_not_reported_as_success(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            dws_call = root / "fail-dws.cmd"
            dws_call.write_text("@echo off\r\necho failed 1>&2\r\nexit /b 9\r\n", encoding="utf-8")
            config_path = self._write_config(root, dws_call)

            result = notifier.execute_run(self._run_args(root, config_path, dry_run=True, send=False))

            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "send_failed")
            self.assertEqual(result["notify_candidates"], 1)
            self.assertEqual(result["sent_items"], 0)
            conn = notifier.connect_db(root / "state.sqlite")
            try:
                attempt_count = conn.execute("SELECT COUNT(*) FROM notification_attempts WHERE status = 'failed'").fetchone()[0]
                notified_count = conn.execute("SELECT COUNT(*) FROM notified_items").fetchone()[0]
            finally:
                conn.close()
            self.assertEqual(attempt_count, 1)
            self.assertEqual(notified_count, 0)

    def test_missing_dws_is_recorded_as_send_failure_without_marking_notified(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config_path = self._write_config(root, root / "missing-dws.cmd")

            result = notifier.execute_run(self._run_args(root, config_path, dry_run=False, send=True))

            self.assertFalse(result["ok"])
            self.assertEqual(result["status"], "send_failed")
            self.assertIn("DWS", result["error"])
            conn = notifier.connect_db(root / "state.sqlite")
            try:
                attempt_count = conn.execute("SELECT COUNT(*) FROM notification_attempts WHERE status = 'failed'").fetchone()[0]
                notified_count = conn.execute("SELECT COUNT(*) FROM notified_items").fetchone()[0]
            finally:
                conn.close()
            self.assertEqual(attempt_count, 1)
            self.assertEqual(notified_count, 0)


class SendSuccessTests(unittest.TestCase):
    def test_successful_fake_send_marks_notified_and_second_run_has_no_candidates(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            dws_call = root / "success-dws.cmd"
            dws_call.write_text("@echo off\r\necho {\"ok\":true}\r\nexit /b 0\r\n", encoding="utf-8")
            config = notifier.default_config()
            config["source"]["type"] = "sample"
            config["dingtalk"]["dws_call"] = str(dws_call)
            config["dingtalk"]["robot_code"] = "robot"
            config["dingtalk"]["group_open_conversation_id"] = "group"
            config["dingtalk"]["send_enabled"] = True
            config["notify"]["require_all_stores_before_send"] = False
            config["state"]["db_path"] = str(root / "state.sqlite")
            config["state"]["result_dir"] = str(root / "runs")
            config_path = root / "config.json"
            config_path.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
            args = argparse.Namespace(
                config=str(config_path),
                state_dir=str(root),
                source_type="",
                source_excel="",
                source_dir="",
                site=notifier.SITE_US,
                store_list="",
                require_all_stores=False,
                skip_store_coverage=True,
                dry_run=False,
                send=True,
            )

            first = notifier.execute_run(args)
            second = notifier.execute_run(args)

            self.assertTrue(first["ok"])
            self.assertEqual(first["status"], "success")
            self.assertEqual(first["sent_items"], 1)
            self.assertTrue(second["ok"])
            self.assertEqual(second["status"], "no_new_items")
            self.assertEqual(second["notify_candidates"], 0)
            conn = notifier.connect_db(root / "state.sqlite")
            try:
                notified_count = conn.execute("SELECT COUNT(*) FROM notified_items").fetchone()[0]
                sent_attempts = conn.execute("SELECT COUNT(*) FROM notification_attempts WHERE status = 'sent'").fetchone()[0]
            finally:
                conn.close()
            self.assertEqual(notified_count, 1)
            self.assertEqual(sent_attempts, 1)


if __name__ == "__main__":
    unittest.main()
