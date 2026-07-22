import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import amazon_frontend_check as monitor


class FakeRetryPage:
    def set_default_timeout(self, _value):
        pass

    def set_default_navigation_timeout(self, _value):
        pass


class FakeLocator:
    def __init__(self, visible=False):
        self.visible = visible

    @property
    def first(self):
        return self

    def count(self):
        return int(self.visible)

    def is_visible(self, timeout=None):
        return self.visible


class FakeBuyingOptionsPage:
    def locator(self, selector):
        return FakeLocator(selector in {"#desktop_buybox", "#buybox"})


class FakeRetryContext:
    def __init__(self):
        self.closed = False

    def new_page(self):
        return FakeRetryPage()

    def close(self):
        self.closed = True


class FakeBrowser:
    def __init__(self):
        self.contexts = []

    def new_context(self, **_kwargs):
        context = FakeRetryContext()
        self.contexts.append(context)
        return context


def base_page(browser):
    return SimpleNamespace(context=SimpleNamespace(browser=browser))


class FakeBatchPage(FakeRetryPage):
    def close(self):
        pass


class FakeBatchContext(FakeRetryContext):
    def new_page(self):
        return FakeBatchPage()


class FakeBatchBrowser:
    def new_context(self, **_kwargs):
        return FakeBatchContext()

    def close(self):
        pass


class FakePlaywrightManager:
    def __enter__(self):
        self.chromium = SimpleNamespace(launch=lambda **_kwargs: FakeBatchBrowser())
        return self

    def __exit__(self, *_args):
        pass


class AmazonFrontendCheckTests(unittest.TestCase):
    def test_default_postal_code_is_10043(self):
        self.assertEqual(monitor.DEFAULT_POSTAL_CODE, "10043")
        self.assertEqual(monitor.POSTAL_CODES, ["10043"])

    def test_project_start_resumes_without_completed_projects(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            for name in ["A00眼镜", "A27小家纺", "A28太阳能灯"]:
                (root / name).mkdir()

            def fake_project_paths(child):
                return SimpleNamespace(root=child, name=child.name)

            with (
                patch.object(monitor, "ROOT", root),
                patch.object(monitor, "PROJECT_FILTER", ""),
                patch.object(monitor, "PROJECT_START", "A27小家纺"),
                patch.object(monitor, "desktop_input_files", return_value=[]),
                patch.object(monitor, "project_paths_from_existing", side_effect=fake_project_paths),
            ):
                actual = monitor.discover_projects()

        self.assertEqual([project.name for project in actual], ["A27小家纺", "A28太阳能灯"])

    def test_available_buybox_does_not_claim_zip_dependency(self):
        result = {"status": "OK", "buybox": "有购物车"}
        with patch.object(monitor, "extract_product", return_value=result):
            actual = monitor.extract_with_zip_fallback(base_page(FakeBrowser()), {}, None)

        self.assertEqual(actual["delivery_zip_checked"], "")
        self.assertFalse(actual["buybox_zip_dependent"])
        self.assertEqual(actual["buybox_zip_success"], "")

    def test_zip_fallback_records_full_path_and_success_zip(self):
        browser = FakeBrowser()
        initial = {"status": "OK", "buybox": "不可售/无购物车"}
        recovered = {"status": "OK", "buybox": "有Buy Now"}
        with (
            patch.object(monitor, "POSTAL_CODES", ["90012", "10001", "91748"]),
            patch.object(monitor, "extract_product", side_effect=[initial, recovered]),
            patch.object(monitor, "probe_buybox", side_effect=["不可售/无购物车", "有Buy Now"]),
            patch.object(monitor, "set_delivery_zip"),
        ):
            actual = monitor.extract_with_zip_fallback(base_page(browser), {}, None)

        self.assertEqual(actual["delivery_zip_checked"], "90012 / 10001")
        self.assertTrue(actual["buybox_zip_dependent"])
        self.assertEqual(actual["buybox_zip_success"], "10001")
        self.assertEqual(len(browser.contexts), 2)
        self.assertTrue(all(context.closed for context in browser.contexts))

    def test_failed_zip_attempt_does_not_abort_remaining_fallbacks(self):
        browser = FakeBrowser()
        initial = {"status": "OK", "buybox": "不可售/无购物车"}
        recovered = {"status": "OK", "buybox": "有购物车"}
        with (
            patch.object(monitor, "POSTAL_CODES", ["90012", "10001"]),
            patch.object(monitor, "extract_product", side_effect=[initial, recovered]),
            patch.object(monitor, "probe_buybox", side_effect=[RuntimeError("transient retry failure"), "有购物车"]),
            patch.object(monitor, "set_delivery_zip"),
        ):
            actual = monitor.extract_with_zip_fallback(base_page(browser), {}, None)

        self.assertEqual(actual["delivery_zip_checked"], "90012 / 10001")
        self.assertEqual(actual["buybox_zip_success"], "10001")

    def test_all_failed_zip_probes_do_not_repeat_full_product_capture(self):
        browser = FakeBrowser()
        initial = {"status": "OK", "buybox": "不可售/无购物车", "current_price": ""}
        with (
            patch.object(monitor, "POSTAL_CODES", ["90012", "10001", "91748"]),
            patch.object(monitor, "extract_product", return_value=initial) as extract_product,
            patch.object(monitor, "probe_buybox", return_value="不可售/无购物车"),
            patch.object(monitor, "set_delivery_zip"),
        ):
            actual = monitor.extract_with_zip_fallback(base_page(browser), {}, None)

        self.assertEqual(extract_product.call_count, 1)
        self.assertEqual(actual["delivery_zip_checked"], "90012 / 10001 / 91748")
        self.assertFalse(actual["buybox_zip_dependent"])

    def test_failed_zip_change_is_not_used_for_buybox_probe(self):
        browser = FakeBrowser()
        initial = {"status": "OK", "buybox": "不可售/无购物车", "current_price": ""}
        with (
            patch.object(monitor, "POSTAL_CODES", ["10043"]),
            patch.object(monitor, "extract_product", return_value=initial),
            patch.object(monitor, "set_delivery_zip", return_value=False),
            patch.object(monitor, "probe_buybox") as probe_buybox,
        ):
            actual = monitor.extract_with_zip_fallback(base_page(browser), {"子ASIN网址": "https://www.amazon.com/dp/X"}, None)

        probe_buybox.assert_not_called()
        self.assertEqual(actual["delivery_zip_checked"], "10043")
        self.assertFalse(actual["buybox_zip_dependent"])

    def test_current_price_does_not_fall_back_to_other_variant_price(self):
        def fake_first_text(_page, selectors):
            if any("twister" in selector or "variation" in selector for selector in selectors):
                return "$135.56"
            return ""

        with (
            patch.object(monitor, "first_text", side_effect=fake_first_text),
            patch.object(monitor, "visible_money_from_selectors", return_value=""),
            patch.object(monitor, "visible_text_containing", return_value=""),
        ):
            actual = monitor.extract_price_details(object())

        self.assertEqual(actual["current_price"], "")

    def test_typical_price_alone_does_not_create_strike_or_discount(self):
        price_html = """
        <div>Typical Price $14.99</div>
        <span class="a-price-whole">14.</span><span class="a-price-fraction">24</span>
        <span>-5%</span>
        """
        with patch.object(monitor, "html_first_block", return_value=price_html):
            html_result = monitor.html_detect_price("ignored")

        self.assertEqual(html_result["list_price"], "")
        self.assertEqual(html_result["typical_price"], "$14.99")
        self.assertEqual(html_result["has_strike"], "否")
        self.assertEqual(html_result["discount"], "")

        with (
            patch.object(monitor, "first_text", return_value="Typical Price: $14.99 $14.24 -5%"),
            patch.object(monitor, "visible_money_from_selectors", return_value=""),
            patch.object(monitor, "visible_text_containing", return_value="-5%"),
        ):
            dom_result = monitor.extract_price_details(object())

        self.assertEqual(dom_result["list_price"], "")
        self.assertEqual(dom_result["typical_price"], "$14.99")
        self.assertEqual(dom_result["has_strike"], "否")
        self.assertEqual(dom_result["discount"], "")

    def test_prime_member_price_is_offer_and_regular_price_is_buybox_price(self):
        price_text = """
        Prime Member Price $56.99 This price is exclusively for Amazon Prime members.
        Regular Price $79.99
        Typical Price $59.99 Exclusive Prime price
        """
        with patch.object(monitor, "html_first_block", return_value=price_text):
            html_result = monitor.html_detect_price("ignored")

        self.assertEqual(html_result["current_price"], "$79.99")
        self.assertEqual(html_result["prime_offer"], "Prime会员专享折扣：$56.99")
        self.assertEqual(html_result["typical_price"], "$59.99")

        def fake_first_text(_page, selectors):
            if "#desktop_buybox" in selectors:
                return "Prime Member Price $56.99 exclusively for Amazon Prime members Regular Price $79.99"
            return "$56.99 Typical Price $59.99 Exclusive Prime price"

        with (
            patch.object(monitor, "first_text", side_effect=fake_first_text),
            patch.object(monitor, "visible_money_from_selectors", return_value=""),
            patch.object(monitor, "visible_text_containing", return_value=""),
        ):
            dom_result = monitor.extract_price_details(object())

        self.assertEqual(dom_result["current_price"], "$79.99")
        self.assertEqual(dom_result["prime_offer"], "Prime会员专享折扣：$56.99")

    def test_see_all_buying_options_is_not_confirmed_buybox(self):
        html = '<div id="desktop_buybox">See All Buying Options</div>'
        self.assertEqual(monitor.html_detect_buybox(html), "无购物车/仅购买选项")

        with patch.object(monitor, "first_text", return_value="See All Buying Options"):
            actual = monitor.detect_buybox(FakeBuyingOptionsPage(), "")

        self.assertEqual(actual, "无购物车/仅购买选项")
        self.assertTrue(monitor.is_cart_lost(actual))

    def test_persistent_missing_buybox_is_always_an_exception(self):
        current = {
            "status": "OK",
            "title": "Mirror",
            "rating": "4.6",
            "reviews": "462",
            "category": "Home & Kitchen",
            "rank": "#53 in Floor & Full Length Mirrors",
            "buybox": "无购物车/仅购买选项",
            "other_sellers": "无明显跟卖",
        }
        previous = {
            "B0CPLW4G8V": {
                "title": "Mirror",
                "rating": "4.6",
                "reviews": "462",
                "category": "Home & Kitchen",
                "rank": "#53 in Floor & Full Length Mirrors",
                "buybox": "购物车区域可见",
            }
        }

        issues, _ = monitor.compare(
            {"父ASIN": "B0CPLW39FW", "子ASIN": "B0CPLW4G8V"},
            current,
            previous,
        )

        cart_issues = [issue for issue in issues if issue["问题模块"] == "购物车丢失"]
        self.assertEqual(len(cart_issues), 1)
        self.assertIn("未检测到 Add to Cart/Buy Now", cart_issues[0]["问题摘要"])

    def test_unavailable_offer_clears_hidden_price_and_promotions(self):
        actual = monitor.normalize_unavailable_offer_fields({
            "buybox": "不可售/无购物车",
            "list_price": "$199.99",
            "typical_price": "$149.99",
            "current_price": "$135.56",
            "discount": "-10%",
            "prime": "Prime Member Price",
            "coupon": "Apply 5% coupon",
            "multi_buy": "Shop items",
            "has_strike": "是",
        })

        self.assertEqual(actual["current_price"], "")
        self.assertEqual(actual["multi_buy"], "")
        self.assertEqual(actual["has_strike"], "否")

    def test_embedded_offer_json_is_not_reported_as_promotion(self):
        value = 'Eligible\\":true,\\"offerListingId\\":\\"hidden\\"'
        self.assertEqual(monitor.concise_offer_text(value), "")

    def test_multi_buy_clean_extracts_select_item_discounts(self):
        text = (
            "Exclusive Prime price Save 5% on 2 select item(s) Shop items "
            "Save 7% on 3 select item(s) Shop items "
            "Save 10% on 5 select item(s) Shop items"
        )

        self.assertEqual(
            monitor.multi_buy_clean(text),
            "Save 5% on 2 select item(s)；Save 7% on 3 select item(s)；Save 10% on 5 select item(s)",
        )

    def test_parent_review_split_is_not_duplicated_in_notes(self):
        current = {
            "status": "OK",
            "title": "Product",
            "rating": "4.5",
            "reviews": "10",
            "category": "Home",
            "rank": "#1 in Home",
            "buybox": "有购物车",
            "parent_review_split": "同父体评论数不一致：A=10，B=20",
            "other_sellers": "无明显跟卖",
        }
        issues, notes = monitor.compare({"父ASIN": "PARENT", "子ASIN": "CHILD"}, current, {})

        self.assertTrue(any(issue["问题模块"] == "父体评论拆分" for issue in issues))
        self.assertNotIn("同父体评论数不一致", notes)

    def test_issue_notes_use_blank_lines_and_remove_repeated_change_text(self):
        issues = [
            {"问题模块": "排名变化", "问题摘要": "大类排名上升超过10%：#409,124 -> #161,462（60.5%）"},
            {"问题模块": "排名变化", "问题摘要": "小类排名上升超过10%：#292 -> #132（54.8%）"},
            {"问题模块": "跟卖", "问题摘要": "疑似跟卖：New (2) from"},
            {"问题模块": "排名变化", "问题摘要": "大类排名上升超过10%：#409,124 -> #161,462（60.5%）"},
        ]
        actual = monitor.readable_issue_notes(
            issues,
            ["大类排名上升超过10%：#409,124 -> #161,462（60.5%）；小类排名上升超过10%：#292 -> #132（54.8%）"],
            ["邮编检查：10043；购物车依赖邮编：10043"],
        )

        self.assertEqual(actual.count("大类排名上升超过10%"), 1)
        self.assertEqual(actual.count("小类排名上升超过10%"), 1)
        self.assertIn("\n\n跟卖：疑似跟卖：New (2) from\n\n", actual)
        self.assertTrue(actual.endswith("邮编检查：10043；购物车依赖邮编：10043"))

    def test_page_not_found_is_valid_link_exception_not_startup_failure(self):
        record = {
            "current": {
                "status": "ERROR",
                "not_found": True,
                "title": "Page Not Found",
                "error": "Amazon returned Page Not Found",
            }
        }

        self.assertEqual(monitor.capture_failure_reason(record), "")

    def test_progress_timestamp_and_zip_dependency_self_check(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            cache_dir = Path(temp_dir)
            paths = SimpleNamespace(cache_dir=cache_dir, name="test")
            progress = {"checked_at": "stale"}
            monitor.persist_run_progress(paths, progress)
            persisted = json.loads((cache_dir / "run_progress.json").read_text(encoding="utf-8"))
            self.assertNotEqual(persisted["checked_at"], "stale")
            self.assertEqual(persisted["checked_at"], persisted["updated_at"])

            row = [""] * len(monitor.REPORT_HEADERS)
            row[monitor.REPORT_HEADERS.index("子ASIN")] = "B000000001"
            row[monitor.REPORT_HEADERS.index("购物车")] = "有购物车"
            row[monitor.REPORT_HEADERS.index("当前/Buy Box价格")] = "$19.99"
            row[monitor.REPORT_HEADERS.index("备注")] = "邮编检查：90012 / 10001；购物车依赖邮编：10001"
            unavailable = [""] * len(monitor.REPORT_HEADERS)
            unavailable[monitor.REPORT_HEADERS.index("子ASIN")] = "B000000002"
            unavailable[monitor.REPORT_HEADERS.index("购物车")] = "不可售/无购物车"
            unavailable[monitor.REPORT_HEADERS.index("备注")] = "邮编检查：90012 / 10001 / 91748"
            typical_only = [""] * len(monitor.REPORT_HEADERS)
            typical_only[monitor.REPORT_HEADERS.index("子ASIN")] = "B000000003"
            typical_only[monitor.REPORT_HEADERS.index("购物车")] = "有购物车"
            typical_only[monitor.REPORT_HEADERS.index("当前/Buy Box价格")] = "$14.24"
            typical_only[monitor.REPORT_HEADERS.index("Typical Price")] = "$14.99"
            typical_only[monitor.REPORT_HEADERS.index("是否有划线")] = "是"
            typical_only[monitor.REPORT_HEADERS.index("划线百分比")] = "-5%"
            monitor.write_self_check_report(paths, [row, unavailable, typical_only], [], {"run_mode": "filtered"})
            report_path = next((cache_dir / "self_checks").glob("self_check_*.json"))
            report = json.loads(report_path.read_text(encoding="utf-8"))
            self.assertEqual(report["zipcode_dependent_buybox_count"], 1)
            self.assertEqual(
                report["zipcode_dependent_buybox"],
                [{"child_asin": "B000000001", "zipcode": "10001"}],
            )
            self.assertEqual(report["missing_current_price"], 0)
            self.assertEqual(report["unavailable_buybox_rows"], 1)
            self.assertEqual(report["zipcode_checked_without_buybox"], ["B000000002"])
            self.assertEqual(report["typical_only_marked_strike"], ["B000000003"])

    def test_batches_are_isolated_and_completed_batches_resume(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_file = root / "input.xlsx"
            input_file.touch()
            paths = monitor.ProjectPaths(
                name="batch-test",
                root=root,
                input_dir=root,
                output_dir=root / "output",
                cache_dir=root / "cache",
                input_file=input_file,
                total_book=root / "output" / "total.xlsx",
                exception_book=root / "output" / "exceptions.xlsx",
                snapshot_file=root / "cache" / "latest_snapshot.json",
                screenshot_dir=root / "cache" / "screenshots",
                script_cache_dir=root / "cache" / "scripts",
            )
            items = [
                {"父ASIN": "PARENT", "子ASIN": "B000000001"},
                {"父ASIN": "PARENT", "子ASIN": "B000000002"},
            ]
            with patch.object(monitor, "ASIN_BATCH_SIZE", 1):
                batches = monitor.create_batches(items, root / "cache" / "batch_runs" / "run")

            screenshot_dirs = []

            def fake_extract(_page, item, batch_paths):
                screenshot_dirs.append(batch_paths.screenshot_dir)
                return {"status": "OK", "captcha": False, "blocked": False, "title": item["子ASIN"], "reviews": "1"}

            with (
                patch.object(monitor, "sync_playwright", side_effect=lambda: FakePlaywrightManager()),
                patch.object(monitor, "extract_with_zip_fallback", side_effect=fake_extract),
            ):
                with monitor.ThreadPoolExecutor(max_workers=2) as executor:
                    futures = [executor.submit(monitor.run_batch, paths, batch, lambda *_args: None) for batch in batches]
                    results = [future.result() for future in futures]

            self.assertEqual([result["status"] for result in results], ["completed", "completed"])
            self.assertEqual(len(set(screenshot_dirs)), 2)
            self.assertTrue(all(path.parent.name.startswith("batch_") for path in screenshot_dirs))
            self.assertEqual(monitor.completed_batch_records(batches[0])[0]["index"], 0)

    def test_resource_aware_workers_honor_memory_cap(self):
        with (
            patch.object(monitor, "ASIN_BATCH_WORKERS", 3),
            patch.object(monitor, "available_memory_gb", return_value=3.1),
            patch.object(monitor, "cpu_count", return_value=16),
        ):
            workers, decision = monitor.determine_batch_workers(8)

        self.assertEqual(workers, 1)
        self.assertEqual(decision["memory_limit"], 1)

    def test_project_run_reuses_completed_batches(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_file = root / "input.xlsx"
            input_file.touch()
            paths = monitor.ProjectPaths(
                name="resume-test",
                root=root,
                input_dir=root / "input",
                output_dir=root / "output",
                cache_dir=root / "cache",
                input_file=input_file,
                total_book=root / "output" / "total.xlsx",
                exception_book=root / "output" / "exceptions.xlsx",
                snapshot_file=root / "cache" / "latest_snapshot.json",
                screenshot_dir=root / "cache" / "screenshots",
                script_cache_dir=root / "cache" / "scripts",
            )
            items = [
                {"父ASIN": "PARENT", "子ASIN": "B000000001", "子SKU": "S1"},
                {"父ASIN": "PARENT", "子ASIN": "B000000002", "子SKU": "S2"},
            ]
            capture_count = 0

            def fake_extract(_page, item, _batch_paths):
                nonlocal capture_count
                capture_count += 1
                return {
                    "status": "OK", "captcha": False, "blocked": False, "title": item["子ASIN"],
                    "reviews": "10", "rating": "4.5", "category": "Home", "rank": "#1 in Home",
                    "buybox": "有购物车", "other_sellers": "无明显跟卖", "aplus_visible": "是",
                }

            with (
                patch.object(monitor, "read_items", return_value=items),
                patch.object(monitor, "load_previous", return_value={}),
                patch.object(monitor, "ASIN_BATCH_SIZE", 1),
                patch.object(monitor, "ASIN_BATCH_WORKERS", 2),
                patch.object(monitor, "ASIN_BATCH_RESUME", True),
                patch.object(monitor, "available_memory_gb", return_value=8.0),
                patch.object(monitor, "sync_playwright", side_effect=lambda: FakePlaywrightManager()),
                patch.object(monitor, "extract_with_zip_fallback", side_effect=fake_extract),
                patch.object(monitor, "write_link_check_workbook", return_value=paths.total_book),
                patch.object(monitor, "write_exception_summary_workbook", return_value=paths.exception_book),
                patch.object(monitor, "write_self_check_report"),
            ):
                first = monitor.run_project(paths)
                second = monitor.run_project(paths)

            progress = json.loads((paths.cache_dir / "batch_progress.json").read_text(encoding="utf-8"))
            self.assertEqual(capture_count, 2)
            self.assertEqual(first["processed_count"], 2)
            self.assertEqual(second["resumed_batch_count"], 2)
            self.assertEqual(progress["status"], "completed")
            self.assertEqual(progress["completed_batch_count"], 2)

    def test_failed_batch_does_not_discard_successful_batch(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_file = root / "input.xlsx"
            input_file.touch()
            paths = monitor.ProjectPaths(
                name="failure-test", root=root, input_dir=root, output_dir=root / "output",
                cache_dir=root / "cache", input_file=input_file, total_book=root / "output" / "total.xlsx",
                exception_book=root / "output" / "exceptions.xlsx", snapshot_file=root / "cache" / "latest.json",
                screenshot_dir=root / "cache" / "screenshots", script_cache_dir=root / "cache" / "scripts",
            )
            items = [
                {"父ASIN": "PARENT", "子ASIN": "B000000001"},
                {"父ASIN": "PARENT", "子ASIN": "B000000002"},
            ]
            with patch.object(monitor, "ASIN_BATCH_SIZE", 1):
                batches = monitor.create_batches(items, root / "cache" / "batch_runs" / "run")

            def fail_playwright():
                raise RuntimeError("browser startup failed")

            with patch.object(monitor, "sync_playwright", side_effect=fail_playwright):
                failed = monitor.run_batch(paths, batches[0], lambda *_args: None)
            with (
                patch.object(monitor, "sync_playwright", side_effect=lambda: FakePlaywrightManager()),
                patch.object(monitor, "extract_with_zip_fallback", return_value={"status": "OK", "captcha": False, "blocked": False}),
            ):
                successful = monitor.run_batch(paths, batches[1], lambda *_args: None)

            self.assertEqual(failed["status"], "failed")
            self.assertEqual(successful["status"], "completed")
            self.assertTrue(batches[0]["result_file"].exists())
            self.assertTrue(batches[1]["result_file"].exists())


if __name__ == "__main__":
    unittest.main()
