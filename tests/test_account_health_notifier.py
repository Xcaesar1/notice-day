from __future__ import annotations

import argparse
import json
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import account_health_notifier as notifier  # noqa: E402


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


class RunIdTests(unittest.TestCase):
    def test_run_id_has_microsecond_precision(self) -> None:
        first = notifier.run_id_text()
        second = notifier.run_id_text()

        self.assertRegex(first, r"^\d{8}_\d{6}_\d{6}_[0-9a-f]{8}$")
        self.assertNotEqual(first, second)


class RunCoverageGuardTests(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
