import json
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import amazon_frontend_check as monitor
from openpyxl import Workbook


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


class FakeClickableLocator(FakeLocator):
    def __init__(self, visible=False, value=""):
        super().__init__(visible)
        self.value = value
        self.clicked = False

    def click(self, **_kwargs):
        self.clicked = True

    def get_attribute(self, name):
        return self.value if name == "value" else None


class FakeBuyingOptionsPage:
    def locator(self, selector):
        return FakeLocator(selector in {"#desktop_buybox", "#buybox"})


class FakeVisibleCartPage:
    def locator(self, selector):
        return FakeLocator(selector == "#add-to-cart-button")


class FakeBodyTextLocator:
    def __init__(self, text):
        self.text = text

    def inner_text(self, timeout=None):
        return self.text


class FakeNotFoundPage:
    def __init__(self):
        self.goto_calls = 0

    def goto(self, *_args, **_kwargs):
        self.goto_calls += 1

    def wait_for_timeout(self, _value):
        pass

    def locator(self, selector):
        if selector == "body":
            return FakeBodyTextLocator("Sorry! We couldn't find that page. Page Not Found")
        return FakeLocator(False)


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
    def __init__(self):
        self.launch_count = 0

    def __enter__(self):
        def launch(**_kwargs):
            self.launch_count += 1
            return FakeBatchBrowser()

        self.chromium = SimpleNamespace(launch=launch)
        return self

    def __exit__(self, *_args):
        pass


class FakeReadinessPage:
    def __init__(self, fail=False):
        self.fail = fail
        self.wait_timeouts = []
        self.evaluations = []

    def wait_for_function(self, _expression, timeout=None):
        self.wait_timeouts.append(timeout)
        if self.fail:
            raise RuntimeError("not ready")

    def evaluate(self, expression):
        self.evaluations.append(expression)

    def wait_for_timeout(self, timeout):
        self.wait_timeouts.append(timeout)


class AmazonFrontendCheckTests(unittest.TestCase):
    def test_confirmed_not_found_page_short_circuits_detail_capture(self):
        page = FakeNotFoundPage()
        with (
            patch.object(monitor, "was_previously_not_found", return_value=False),
            patch.object(monitor, "dismiss_amazon_continue", return_value=False),
            patch.object(monitor, "wait_for_product_header_ready"),
            patch.object(monitor, "first_text", return_value="Page Not Found"),
            patch.object(monitor, "load_detail_sections") as load_details,
            patch.object(monitor, "screenshot_price") as price_screenshot,
            patch.object(monitor, "screenshot_customer_reviews_card") as review_screenshot,
        ):
            current = monitor.extract_product(
                page,
                {"子ASIN": "B0GVKFK8S3", "子ASIN网址": "https://www.amazon.com/dp/B0GVKFK8S3"},
                None,
            )

        self.assertEqual(current["status"], "ERROR")
        self.assertTrue(current["not_found"])
        self.assertEqual(current["error"], "Amazon returned Page Not Found")
        load_details.assert_not_called()
        price_screenshot.assert_not_called()
        review_screenshot.assert_not_called()

    def test_previous_not_found_curl_confirmation_skips_browser_navigation(self):
        page = FakeNotFoundPage()
        with (
            patch.object(monitor, "was_previously_not_found", return_value=True),
            patch.object(
                monitor,
                "curl_fetch_html",
                return_value="<html><title>Page Not Found</title><body>The Dogs of Amazon</body></html>",
            ),
        ):
            current = monitor.extract_product(
                page,
                {"子ASIN": "B0GVKFK8S3", "子ASIN网址": "https://www.amazon.com/dp/B0GVKFK8S3"},
                None,
            )

        self.assertEqual(page.goto_calls, 0)
        self.assertTrue(current["not_found"])
        self.assertEqual(current["fetch_method"], "curl_not_found_precheck")

    def test_project_watchdog_retries_have_hard_cap(self):
        self.assertLessEqual(monitor.ASIN_PROJECT_STALL_RETRIES, 3)

    def test_discovery_places_deferred_container_projects_last(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            regular = root / "A10"
            deferred_container = root / "暂不做检查"
            deferred = deferred_container / "A09"
            regular.mkdir()
            deferred.mkdir(parents=True)

            def fake_existing(path):
                if path in {regular, deferred}:
                    return SimpleNamespace(name=path.name, root=path)
                return None

            with (
                patch.object(monitor, "ROOT", root),
                patch.object(monitor, "INCLUDE_DESKTOP_INPUTS", False),
                patch.object(monitor, "PROJECT_EXCLUDES", set()),
                patch.object(monitor, "PROJECT_PREFIXES", ()),
                patch.object(monitor, "PROJECT_FILTER", ""),
                patch.object(monitor, "PROJECT_START", ""),
                patch.object(monitor, "project_paths_from_existing", side_effect=fake_existing),
            ):
                projects = monitor.discover_projects()

            self.assertEqual([project.name for project in projects], ["A10", "A09"])

    def test_existing_project_accepts_numbered_input_copy(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            project_root = Path(temp_dir) / "A00眼镜"
            input_dir = project_root / "1_输入需求信息"
            (project_root / "2_输出信息").mkdir(parents=True)
            (project_root / "3_系统运行缓存").mkdir()
            input_dir.mkdir()
            numbered_input = input_dir / "A00眼镜-ASIN检查基础信息(1).xlsx"
            numbered_input.touch()

            with patch.object(monitor, "ensure_legacy_project_entry"):
                paths = monitor.project_paths_from_existing(project_root)

            self.assertIsNotNone(paths)
            self.assertEqual(paths.input_file, numbered_input)

    def test_top_screenshot_waits_for_render_complete_images(self):
        page = FakeReadinessPage()

        self.assertTrue(monitor.wait_for_top_screenshot_ready(page, timeout_ms=4321))
        self.assertEqual(page.evaluations, ["window.scrollTo(0, 0)"])
        self.assertEqual(page.wait_timeouts, [4321, 900])

    def test_rank_capture_retries_only_after_empty_first_pass(self):
        page = object()
        with (
            patch.object(monitor, "_extract_rank_text_once", side_effect=["", "#12 in Home"]),
            patch.object(monitor, "wait_for_rank_ready", return_value=True) as wait_ready,
        ):
            rank = monitor.extract_rank_text(page)

        self.assertEqual(rank, "#12 in Home")
        wait_ready.assert_called_once_with(page, timeout_ms=4000)

    def test_rating_cell_text_shows_change_with_readable_spacing(self):
        self.assertEqual(
            monitor.rating_cell_text("4.5 out of 5 stars", "4.4 out of 5 stars"),
            "4.5 out of 5 stars\n\n较上次：4.4，上升 0.1",
        )
        self.assertEqual(
            monitor.rating_cell_text("4.3 out of 5 stars", "4.5 out of 5 stars"),
            "4.3 out of 5 stars\n\n较上次：4.5，下降 0.2",
        )
        self.assertEqual(
            monitor.rating_cell_text("4.5 out of 5 stars", "4.5 out of 5 stars"),
            "4.5 out of 5 stars\n\n较上次：4.5，无变化",
        )

    def test_rating_cell_text_without_valid_baseline_keeps_frontend_text(self):
        self.assertEqual(monitor.rating_cell_text("4.5 out of 5 stars", ""), "4.5 out of 5 stars")

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
                patch.object(monitor, "INCLUDE_DESKTOP_INPUTS", False),
                patch.object(monitor, "PROJECT_PREFIXES", ()),
                patch.object(monitor, "PROJECT_EXCLUDES", set()),
                patch.object(monitor, "PROJECT_FILTER", ""),
                patch.object(monitor, "PROJECT_START", "A27小家纺"),
                patch.object(monitor, "desktop_input_files", return_value=[]),
                patch.object(monitor, "project_paths_from_existing", side_effect=fake_project_paths),
            ):
                actual = monitor.discover_projects()

        self.assertEqual([project.name for project in actual], ["A27小家纺", "A28太阳能灯"])

    def test_project_exclude_omits_named_folder(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            for name in ["A00眼镜", "A28太阳能灯", "暂不做检查"]:
                (root / name).mkdir()

            def fake_project_paths(child):
                return SimpleNamespace(root=child, name=child.name)

            with (
                patch.object(monitor, "ROOT", root),
                patch.object(monitor, "INCLUDE_DESKTOP_INPUTS", False),
                patch.object(monitor, "PROJECT_PREFIXES", ()),
                patch.object(monitor, "PROJECT_EXCLUDES", {"暂不做检查"}),
                patch.object(monitor, "PROJECT_FILTER", ""),
                patch.object(monitor, "PROJECT_START", ""),
                patch.object(monitor, "desktop_input_files", return_value=[]),
                patch.object(monitor, "project_paths_from_existing", side_effect=fake_project_paths),
            ):
                actual = monitor.discover_projects()

        self.assertEqual([project.name for project in actual], ["A00眼镜", "A28太阳能灯"])

    def test_project_prefixes_include_only_matching_folders(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            for name in ["A29SAKA", "A29YHD", "A31SE", "A32北耕", "A39小家电"]:
                (root / name).mkdir()

            def fake_project_paths(child):
                return SimpleNamespace(root=child, name=child.name)

            with (
                patch.object(monitor, "ROOT", root),
                patch.object(monitor, "INCLUDE_DESKTOP_INPUTS", False),
                patch.object(monitor, "PROJECT_PREFIXES", ("A29", "A32")),
                patch.object(monitor, "PROJECT_EXCLUDES", set()),
                patch.object(monitor, "PROJECT_FILTER", ""),
                patch.object(monitor, "PROJECT_START", ""),
                patch.object(monitor, "project_paths_from_existing", side_effect=fake_project_paths),
            ):
                actual = monitor.discover_projects()

        self.assertEqual([project.name for project in actual], ["A29SAKA", "A29YHD", "A32北耕"])

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

    def test_postal_initialization_retries_with_fresh_page(self):
        class PostalPage:
            def __init__(self):
                self.closed = False

            def goto(self, *_args, **_kwargs):
                pass

            def set_default_timeout(self, _value):
                pass

            def set_default_navigation_timeout(self, _value):
                pass

            def close(self):
                self.closed = True

        class PostalContext:
            def __init__(self):
                self.pages = []
                self.storage_calls = 0

            def new_page(self):
                page = PostalPage()
                self.pages.append(page)
                return page

            def storage_state(self, **_kwargs):
                self.storage_calls += 1

        context = PostalContext()
        with (
            patch.object(monitor, "POSTAL_CODES", ["10043"]),
            patch.object(monitor, "POSTAL_INIT_ATTEMPTS", 3),
            patch.object(monitor, "set_delivery_zip", side_effect=[False, True]) as set_zip,
            patch.object(monitor, "first_text", return_value=""),
            patch.object(monitor, "console_log") as log,
            patch.object(monitor, "sleep"),
        ):
            actual = monitor.initialize_context_postal_code(context, {})

        self.assertEqual(actual, "10043")
        self.assertEqual(set_zip.call_count, 2)
        self.assertEqual(len(context.pages), 2)
        self.assertTrue(all(page.closed for page in context.pages))
        self.assertEqual(context.storage_calls, 1)
        self.assertTrue(any(call.args[2] == "WARNING" for call in log.call_args_list))

    def test_postal_initialization_raises_after_all_attempts_fail(self):
        class PostalPage:
            def __init__(self):
                self.closed = False
                self.screenshots = 0

            def goto(self, *_args, **_kwargs):
                pass

            def set_default_timeout(self, _value):
                pass

            def set_default_navigation_timeout(self, _value):
                pass

            def screenshot(self, **_kwargs):
                self.screenshots += 1

            def close(self):
                self.closed = True

        class PostalContext:
            def __init__(self):
                self.pages = []

            def new_page(self):
                page = PostalPage()
                self.pages.append(page)
                return page

        context = PostalContext()
        with (
            patch.object(monitor, "POSTAL_CODES", ["10043"]),
            patch.object(monitor, "POSTAL_INIT_ATTEMPTS", 3),
            patch.object(monitor, "set_delivery_zip", return_value=False) as set_zip,
            patch.object(monitor, "first_text", return_value=""),
            patch.object(monitor, "console_log") as log,
            patch.object(monitor, "sleep"),
        ):
            with self.assertRaisesRegex(RuntimeError, "10043"):
                monitor.initialize_context_postal_code(context, {})

        self.assertEqual(set_zip.call_count, 3)
        self.assertEqual(len(context.pages), 3)
        self.assertTrue(all(page.closed for page in context.pages))
        self.assertEqual(context.pages[-1].screenshots, 1)
        self.assertEqual([call.args[2] for call in log.call_args_list], ["WARNING", "WARNING", "ERROR"])

    def test_postal_initialization_uses_stable_home_page(self):
        class SetupPage:
            def goto(self, *_args, **_kwargs):
                pass

            def set_default_timeout(self, _value):
                pass

            def set_default_navigation_timeout(self, _value):
                pass

            def close(self):
                pass

        page = SetupPage()
        context = SimpleNamespace(new_page=lambda: page, storage_state=lambda **_kwargs: None)
        with (
            patch.object(monitor, "POSTAL_CODES", ["10043"]),
            patch.object(monitor, "set_delivery_zip", return_value=True) as set_zip,
        ):
            actual = monitor.initialize_context_postal_code(
                context,
                {"子ASIN网址": "https://www.amazon.com/dp/DOESNOTEXIST"},
            )

        self.assertEqual(actual, "10043")
        set_zip.assert_called_once_with(page, "10043", "https://www.amazon.com/")

    def test_open_batch_browser_retries_without_cached_state(self):
        class Context:
            def __init__(self):
                self.closed = False

            def close(self):
                self.closed = True

        class Browser:
            def __init__(self):
                self.kwargs = []
                self.contexts = []
                self.closed = False

            def new_context(self, **kwargs):
                self.kwargs.append(kwargs)
                context = Context()
                self.contexts.append(context)
                return context

            def close(self):
                self.closed = True

        browser = Browser()
        playwright = SimpleNamespace(chromium=SimpleNamespace(launch=lambda **_kwargs: browser))
        with tempfile.TemporaryDirectory() as temp_dir:
            state_file = Path(temp_dir) / "postal_state.json"
            state_file.write_text("{}", encoding="utf-8")
            with (
                patch.object(monitor, "POSTAL_STATE_FILE", state_file),
                patch.object(
                    monitor,
                    "initialize_context_postal_code",
                    side_effect=[RuntimeError("cached state failed"), "10043"],
                ),
            ):
                actual_browser, actual_context, postal_code = monitor.open_batch_browser(playwright, {})

        self.assertIs(actual_browser, browser)
        self.assertIs(actual_context, browser.contexts[1])
        self.assertEqual(postal_code, "10043")
        self.assertTrue(browser.contexts[0].closed)
        self.assertFalse(browser.closed)
        self.assertIn("storage_state", browser.kwargs[0])
        self.assertNotIn("storage_state", browser.kwargs[1])

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

    def test_html_price_does_not_use_hidden_buybox_subscription_price(self):
        def fake_block(_html, marker):
            if marker in {"desktop_buybox", "buybox"}:
                return '<div id="desktop_buybox"><span class="a-price"><span class="a-price-whole">19</span><span class="a-price-fraction">79</span></span></div>'
            return ""

        with patch.object(monitor, "html_first_block", side_effect=fake_block):
            actual = monitor.html_detect_price("ignored")

        self.assertEqual(actual["current_price"], "")

    def test_regular_price_offer_is_activated_before_buybox_check(self):
        regular = FakeClickableLocator(True)

        class Page:
            def locator(self, selector):
                return regular if "Regular Price" in selector else FakeClickableLocator(False)

            def wait_for_timeout(self, _value):
                pass

        self.assertTrue(monitor.activate_regular_price_offer(Page()))
        self.assertTrue(regular.clicked)

    def test_asin_identity_mismatch_clears_sibling_offer(self):
        actual = monitor.apply_asin_identity_guard(
            {
                "buybox": "有购物车",
                "current_price": "$19.99",
                "prime": "Prime会员专享折扣：$17.99",
                "coupon": "Apply 5% coupon",
                "multi_buy": "Save 5% on 2 select item(s)",
                "has_strike": "否",
            },
            "B0H6HPZRLW",
            "B0F8Q5Y5CH",
        )

        self.assertTrue(actual["asin_mismatch"])
        self.assertIn("目标ASIN未加载", actual["buybox"])
        self.assertEqual(actual["current_price"], "")
        self.assertEqual(actual["multi_buy"], "")
        self.assertEqual(actual["redirected_buybox"], "有购物车")
        self.assertEqual(actual["redirected_current_price"], "$19.99")
        self.assertFalse(monitor.is_cart_lost(actual["buybox"]))

    def test_unrecognized_redirect_suppresses_derivative_change_issues(self):
        current = monitor.apply_asin_identity_guard(
            {
                "status": "OK",
                "title": "Sibling variation",
                "rating": "4.1",
                "reviews": "3",
                "category": "Home",
                "rank": "#9 in Home",
                "buybox": "有购物车",
                "current_price": "$19.99",
                "aplus_visible": "否",
                "other_sellers": "无明显跟卖",
            },
            "B000000001",
            "B000000002",
        )
        previous = {
            "B000000001": {
                "title": "Requested variation",
                "rating": "4.8",
                "reviews": "236",
                "category": "Home",
                "rank": "#1 in Home",
                "buybox": "有购物车",
                "current_price": "$29.99",
                "aplus_visible": "是",
            }
        }

        issues, notes = monitor.compare(
            {"父ASIN": "PARENT", "子ASIN": "B000000001"},
            current,
            previous,
        )

        self.assertEqual([issue["问题模块"] for issue in issues], ["ASIN无法识别"])
        self.assertNotIn("评论数变化", notes)
        self.assertFalse(any(issue["问题模块"] == "购物车丢失" for issue in issues))

    def test_required_primary_zip_is_set_before_capture(self):
        current = {"status": "OK", "buybox": "有购物车"}
        with (
            patch.object(monitor, "extract_product", return_value=current),
        ):
            actual = monitor.extract_with_zip_fallback(
                object(),
                {"子ASIN网址": "https://www.amazon.com/dp/B000000001"},
                None,
                "10043",
            )

        self.assertEqual(actual["delivery_zip_checked"], "10043")
        self.assertEqual(actual["location_zip"], "10043")

    def test_typical_price_creates_annotated_strike_and_discount(self):
        price_html = """
        <div>Typical Price $14.99</div>
        <span class="a-price-whole">14.</span><span class="a-price-fraction">24</span>
        <span>-5%</span>
        """
        with patch.object(monitor, "html_first_block", return_value=price_html):
            html_result = monitor.html_detect_price("ignored")

        self.assertEqual(html_result["list_price"], "")
        self.assertEqual(html_result["typical_price"], "$14.99")
        self.assertEqual(html_result["has_strike"], "是\n\n划线来源：Typical Price")
        self.assertEqual(html_result["discount"], "-5%")

        with (
            patch.object(monitor, "first_text", return_value="$14.24 -5% Typical Price: $14.99"),
            patch.object(monitor, "visible_money_from_selectors", return_value=""),
            patch.object(monitor, "visible_text_containing", return_value="-5%"),
        ):
            dom_result = monitor.extract_price_details(object())

        self.assertEqual(dom_result["list_price"], "")
        self.assertEqual(dom_result["typical_price"], "$14.99")
        self.assertEqual(dom_result["has_strike"], "是\n\n划线来源：Typical Price")
        self.assertEqual(dom_result["discount"], "-5%")

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

    def test_visible_cart_button_wins_over_unrelated_unavailable_body_text(self):
        def fake_first_text(_page, selectors):
            if selectors == ["body"]:
                return "Currently unavailable for another variation"
            return ""

        with patch.object(monitor, "first_text", side_effect=fake_first_text):
            actual = monitor.detect_buybox(FakeVisibleCartPage(), "")

        self.assertEqual(actual, "有购物车")

    def test_price_fallback_includes_visible_buybox_regions(self):
        seen_selectors = []

        def fake_visible_money(_page, selectors):
            seen_selectors.extend(selectors)
            return "$79.99" if "#rightCol" in selectors else ""

        with (
            patch.object(monitor, "first_text", return_value=""),
            patch.object(monitor, "visible_money_from_selectors", side_effect=fake_visible_money),
            patch.object(monitor, "visible_text_containing", return_value=""),
        ):
            actual = monitor.extract_price_details(object())

        self.assertEqual(actual["current_price"], "$79.99")
        self.assertIn("#desktop_buybox", seen_selectors)
        self.assertIn("#rightCol", seen_selectors)

    def test_selected_variation_price_is_used_when_primary_price_is_missing(self):
        with (
            patch.object(monitor, "first_text", return_value=""),
            patch.object(monitor, "visible_money_from_selectors", side_effect=["", "", "$13.49"]),
            patch.object(monitor, "visible_text_containing", return_value=""),
        ):
            actual = monitor.extract_price_details(object())

        self.assertEqual(actual["current_price"], "$13.49")
        self.assertEqual(actual["price_source"], "selected_variation")

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
            typical_only[monitor.REPORT_HEADERS.index("是否有划线")] = "是\n\n划线来源：Typical Price"
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
            self.assertEqual(report["typical_price_strike_rows"], ["B000000003"])

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

    def test_dual_project_mode_limits_each_project_to_one_batch_worker(self):
        with (
            patch.object(monitor, "ASIN_PROJECT_WORKERS", 2),
            patch.object(monitor, "ASIN_BATCH_WORKERS", 3),
            patch.object(monitor, "available_memory_gb", return_value=16.0),
            patch.object(monitor, "cpu_count", return_value=16),
        ):
            workers, decision = monitor.determine_batch_workers(8)

        self.assertEqual(workers, 1)
        self.assertEqual(decision["project_concurrency_limit"], 1)

    def test_dual_project_scheduler_runs_same_wave_concurrently(self):
        projects = [
            SimpleNamespace(name="project-a", output_dir=Path("output-a")),
            SimpleNamespace(name="project-b", output_dir=Path("output-b")),
        ]
        waves = []

        def fake_parallel_wave(wave):
            waves.append([paths.name for paths in wave])
            return {
                paths.name: {"project": paths.name, "project_success": True}
                for paths in wave
            }, []

        with (
            patch.object(monitor, "ASIN_PROJECT_WORKERS", 2),
            patch.object(monitor, "run_parallel_project_wave", side_effect=fake_parallel_wave),
            patch.object(monitor, "console_log"),
        ):
            summaries = monitor.run_projects(projects)

        self.assertEqual(waves, [["project-a", "project-b"]])
        self.assertEqual([summary["project"] for summary in summaries], ["project-a", "project-b"])

    def test_single_project_scheduler_still_uses_watchdog_process(self):
        projects = [SimpleNamespace(name="project-a", output_dir=Path("output-a"))]
        waves = []

        def fake_parallel_wave(wave):
            waves.append([paths.name for paths in wave])
            return {
                paths.name: {"project": paths.name, "project_success": True}
                for paths in wave
            }, []

        with (
            patch.object(monitor, "ASIN_PROJECT_WORKERS", 1),
            patch.object(monitor, "run_parallel_project_wave", side_effect=fake_parallel_wave),
            patch.object(monitor, "console_log"),
        ):
            summaries = monitor.run_projects(projects)

        self.assertEqual(waves, [["project-a"]])
        self.assertEqual([summary["project"] for summary in summaries], ["project-a"])

    def test_watchdog_failure_retries_only_failed_project(self):
        projects = [SimpleNamespace(name="project-a", output_dir=Path("output-a"))]
        waves = []

        def fake_parallel_wave(wave):
            waves.append([paths.name for paths in wave])
            if len(waves) == 1:
                return {}, [{
                    "project": "project-a",
                    "stage": "project_watchdog",
                    "error": "stalled",
                }]
            return {
                "project-a": {"project": "project-a", "project_success": True}
            }, []

        with (
            patch.object(monitor, "ASIN_PROJECT_WORKERS", 1),
            patch.object(monitor, "ASIN_PROJECT_STALL_RETRIES", 2),
            patch.object(monitor, "run_parallel_project_wave", side_effect=fake_parallel_wave),
            patch.object(monitor, "console_log"),
        ):
            summaries = monitor.run_projects(projects)

        self.assertEqual(waves, [["project-a"], ["project-a"]])
        self.assertEqual([summary["project"] for summary in summaries], ["project-a"])

    def test_failed_project_wave_does_not_start_next_wave(self):
        projects = [
            SimpleNamespace(name="project-a", output_dir=Path("output-a")),
            SimpleNamespace(name="project-b", output_dir=Path("output-b")),
            SimpleNamespace(name="project-c", output_dir=Path("output-c")),
        ]
        waves = []

        def fake_parallel_wave(wave):
            waves.append([paths.name for paths in wave])
            return {
                "project-b": {"project": "project-b", "project_success": True}
            }, [{"project": "project-a", "stage": "project_run", "error": "startup gate failed"}]

        with (
            patch.object(monitor, "ASIN_PROJECT_WORKERS", 2),
            patch.object(monitor, "run_parallel_project_wave", side_effect=fake_parallel_wave),
            patch.object(monitor, "write_root_summary"),
            patch.object(monitor, "console_log"),
        ):
            with self.assertRaisesRegex(RuntimeError, "startup gate failed"):
                monitor.run_projects(projects)

        self.assertEqual(waves, [["project-a", "project-b"]])

    def test_watchdog_marks_running_batch_interrupted(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            cache_dir = root / "cache"
            batch_dir = cache_dir / "batch_runs" / "run-key" / "batch_0001"
            batch_dir.mkdir(parents=True)
            progress_file = batch_dir / "progress.json"
            progress_file.write_text(json.dumps({"status": "running", "error": ""}), encoding="utf-8")
            aggregate = {
                "run_key": "run-key",
                "status": "running",
                "active_batch_count": 1,
                "batches": {
                    "batch_0001": {
                        "status": "running",
                        "progress_file": str(progress_file),
                        "error": "",
                    }
                },
            }
            (cache_dir / "batch_progress.json").write_text(json.dumps(aggregate), encoding="utf-8")
            paths = SimpleNamespace(name="project-a", cache_dir=cache_dir)

            monitor.mark_project_interrupted(paths, "stalled")

            latest = json.loads((cache_dir / "batch_progress.json").read_text(encoding="utf-8"))
            batch = json.loads(progress_file.read_text(encoding="utf-8"))
            self.assertEqual(latest["status"], "interrupted")
            self.assertEqual(latest["active_batch_count"], 0)
            self.assertEqual(latest["batches"]["batch_0001"]["status"], "interrupted")
            self.assertEqual(batch["status"], "interrupted")

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
                patch.object(monitor, "write_exception_summary_workbook", return_value=paths.exception_book) as exception_writer,
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
            self.assertEqual(first["exception_workbook"], "")
            exception_writer.assert_not_called()

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

    def test_browser_session_failure_restarts_from_batch_checkpoint(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            input_file = root / "input.xlsx"
            input_file.touch()
            paths = monitor.ProjectPaths(
                name="session-retry-test", root=root, input_dir=root, output_dir=root / "output",
                cache_dir=root / "cache", input_file=input_file, total_book=root / "output" / "total.xlsx",
                exception_book=root / "output" / "exceptions.xlsx", snapshot_file=root / "cache" / "latest.json",
                screenshot_dir=root / "cache" / "screenshots", script_cache_dir=root / "cache" / "scripts",
            )
            items = [
                {"鐖禔SIN": "PARENT", "瀛怉SIN": "B000000001"},
                {"鐖禔SIN": "PARENT", "瀛怉SIN": "B000000002"},
                {"鐖禔SIN": "PARENT", "瀛怉SIN": "B000000003"},
            ]
            with patch.object(monitor, "ASIN_BATCH_SIZE", 2):
                batch = monitor.create_batches(items, root / "cache" / "batch_runs" / "run")[1]

            calls = []
            failed_once = False

            def flaky_extract(_page, item, _batch_paths):
                nonlocal failed_once
                child = item["瀛怉SIN"]
                calls.append(child)
                if child == "B000000003" and not failed_once:
                    failed_once = True
                    raise RuntimeError("BrowserContext.new_page: Target page, context or browser has been closed")
                return {"status": "OK", "captcha": False, "blocked": False, "title": child, "reviews": "1"}

            with (
                patch.object(monitor, "ASIN_BATCH_SESSION_RETRIES", 1),
                patch.object(monitor, "sync_playwright", side_effect=lambda: FakePlaywrightManager()),
                patch.object(monitor, "extract_with_zip_fallback", side_effect=flaky_extract),
                patch.object(monitor, "sleep"),
            ):
                result = monitor.run_batch(paths, batch, lambda *_args: None)

            self.assertEqual(result["status"], "completed")
            self.assertEqual(result["processed_count"], 2)
            self.assertEqual(result["session_restart_count"], 1)
            self.assertEqual(calls, ["B000000002", "B000000003", "B000000003"])
            self.assertEqual([record["index"] for record in result["records"]], [1, 2])

    def test_technical_capture_error_restarts_without_persisting_error(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            paths = monitor.ProjectPaths(
                name="capture-retry-test", root=root, input_dir=root, output_dir=root / "output",
                cache_dir=root / "cache", input_file=root / "input.xlsx",
                total_book=root / "output" / "total.xlsx",
                exception_book=root / "output" / "exceptions.xlsx",
                snapshot_file=root / "cache" / "latest.json",
                screenshot_dir=root / "cache" / "screenshots",
                script_cache_dir=root / "cache" / "scripts",
            )
            items = [
                {"父ASIN": "PARENT", "子ASIN": "B000000001"},
                {"父ASIN": "PARENT", "子ASIN": "B000000002"},
            ]
            with patch.object(monitor, "ASIN_BATCH_SIZE", 1):
                batch = monitor.create_batches(items, root / "cache" / "batch_runs" / "run")[1]

            calls = []

            def flaky_extract(_page, item, _batch_paths):
                calls.append(item["子ASIN"])
                if len(calls) == 1:
                    return {
                        "status": "ERROR",
                        "captcha": False,
                        "blocked": False,
                        "error": "temporary parser failure",
                    }
                return {
                    "status": "OK",
                    "captcha": False,
                    "blocked": False,
                    "title": item["子ASIN"],
                }

            with (
                patch.object(monitor, "ASIN_BATCH_SESSION_RETRIES", 1),
                patch.object(monitor, "sync_playwright", side_effect=lambda: FakePlaywrightManager()),
                patch.object(monitor, "extract_with_zip_fallback", side_effect=flaky_extract),
                patch.object(monitor, "sleep"),
            ):
                result = monitor.run_batch(paths, batch, lambda *_args: None)

            self.assertEqual(result["status"], "completed")
            self.assertEqual(result["processed_count"], 1)
            self.assertEqual(result["error_count"], 0)
            self.assertEqual(result["session_restart_count"], 1)
            self.assertEqual(calls, ["B000000002", "B000000002"])
            self.assertEqual(result["records"][0]["current"]["status"], "OK")

    def test_page_not_found_is_not_a_retryable_capture_error(self):
        self.assertTrue(monitor.is_retryable_capture_result({"status": "ERROR", "error": "temporary failure"}))
        self.assertFalse(
            monitor.is_retryable_capture_result(
                {"status": "ERROR", "not_found": True, "error": "Amazon returned Page Not Found"}
            )
        )

    def test_transient_amazon_service_error_is_detected(self):
        self.assertTrue(
            monitor.is_transient_amazon_error_page(
                "Amazon.com",
                "Sorry, something went wrong on our end. Please go back and try again.",
            )
        )
        self.assertTrue(monitor.is_transient_amazon_error_page("503 Service Unavailable", ""))
        self.assertFalse(monitor.is_transient_amazon_error_page("Product title", "In Stock Add to Cart"))

    def test_invalid_star_distribution_is_rejected(self):
        self.assertEqual(
            monitor.parse_star_distribution("5 star 69% 4 star 69% 3 star 69% 2 star 69% 1 star 3%"),
            {},
        )
        self.assertEqual(
            monitor.parse_star_distribution("5 star 70% 4 star 20% 3 star 5% 2 star 3% 1 star 2%"),
            {"5星": 70.0, "4星": 20.0, "3星": 5.0, "2星": 3.0, "1星": 2.0},
        )

    def test_unrecognized_asin_is_terminal_and_not_a_capture_failure(self):
        current = monitor.unrecognized_asin_result("BAD-ASIN", "输入格式无效")

        self.assertEqual(current["status"], "ERROR")
        self.assertTrue(current["asin_unrecognized"])
        self.assertIn("ASIN无法识别", current["error"])
        self.assertFalse(monitor.is_retryable_capture_result(current))
        self.assertEqual(monitor.capture_failure_reason({"current": current}), "")

    def test_identity_guard_marks_missing_identity_terminal(self):
        current = monitor.apply_asin_identity_guard(
            {"status": "ERROR", "title": "", "captcha": False, "not_found": False},
            "B000000001",
            "",
        )

        self.assertTrue(current["asin_unrecognized"])
        self.assertFalse(monitor.is_retryable_capture_result(current))

    def test_aplus_image_module_is_visible_even_without_text(self):
        class Locator:
            def count(self):
                return 1

            def scroll_into_view_if_needed(self, timeout=None):
                return None

            def evaluate(self, _script):
                return {"visible": True, "text": "", "meaningful": True}

        class LocatorGroup:
            first = Locator()

        class Page:
            def locator(self, selector):
                return LocatorGroup() if selector == "#aplus" else type("Empty", (), {"first": type("L", (), {"count": lambda self: 0})()})()

            def wait_for_timeout(self, _timeout):
                return None

        text, visible = monitor.extract_aplus_state(Page())

        self.assertEqual(text, "")
        self.assertTrue(visible)

    def test_aplus_css_component_text_is_not_reported_as_copy(self):
        css_text = "data-csa-c-content-id box-sizing: border-box .aplus-module { list-style: none; }"

        self.assertEqual(monitor.sanitize_aplus_text(css_text), "")
        self.assertEqual(monitor.sanitize_aplus_text("Product description Waterproof solar lights"), "Product description Waterproof solar lights")

    def test_html_image_only_aplus_is_visible_without_copy_text(self):
        raw_html = '''
        <div id="aplus_feature_div">
          <div class="aplus-v2 aplus-module">
            <img data-src="https://images.example.test/aplus.jpg" alt="">
          </div>
        </div>
        '''

        self.assertEqual(monitor.html_detect_aplus(raw_html), "")
        self.assertTrue(monitor.html_has_meaningful_aplus(raw_html))

    def test_html_empty_aplus_placeholder_is_not_visible(self):
        raw_html = '''
        <div id="aplus_feature_div"></div>
        <div data-feature-name="customerReviews"><img src="review.jpg"></div>
        '''

        self.assertFalse(monitor.html_has_meaningful_aplus(raw_html))

    def test_rank_wait_does_not_reference_product_bullets(self):
        class Page:
            def wait_for_function(self, _script, timeout=None):
                return None

        with (
            patch.object(monitor, "load_detail_sections"),
            patch.object(monitor, "extract_bullets_from_page", side_effect=AssertionError("unexpected bullet extraction")),
        ):
            self.assertTrue(monitor.wait_for_rank_ready(Page()))

    def test_batch_rotates_browser_before_session_item_limit(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            paths = monitor.ProjectPaths(
                name="rotation-test", root=root, input_dir=root, output_dir=root / "output",
                cache_dir=root / "cache", input_file=root / "input.xlsx",
                total_book=root / "output" / "total.xlsx",
                exception_book=root / "output" / "exceptions.xlsx",
                snapshot_file=root / "cache" / "latest.json",
                screenshot_dir=root / "cache" / "screenshots",
                script_cache_dir=root / "cache" / "scripts",
            )
            items = [
                {"閻栫SIN": "PARENT", "鐎涙€塖IN": f"B00000000{index}"}
                for index in range(1, 7)
            ]
            with patch.object(monitor, "ASIN_BATCH_SIZE", 5):
                batch = monitor.create_batches(items, root / "cache" / "batch_runs" / "run")[1]
            manager = FakePlaywrightManager()

            with (
                patch.object(monitor, "ASIN_BROWSER_SESSION_ITEM_LIMIT", 2),
                patch.object(monitor, "sync_playwright", return_value=manager),
                patch.object(monitor, "initialize_context_postal_code", return_value="10043"),
                patch.object(
                    monitor,
                    "extract_with_zip_fallback",
                    return_value={"status": "OK", "captcha": False, "blocked": False, "title": "item"},
                ),
            ):
                result = monitor.run_batch(paths, batch, lambda *_args: None)

            self.assertEqual(result["status"], "completed")
            self.assertEqual(result["processed_count"], 5)
            self.assertEqual(manager.launch_count, 3)

    def test_marketplace_domains_are_detected(self):
        self.assertEqual(monitor.marketplace_from_url("https://www.amazon.com/dp/B000000001"), "US")
        self.assertEqual(monitor.marketplace_from_url("https://www.amazon.co.uk/dp/B000000001"), "UK")
        self.assertEqual(monitor.marketplace_from_url("https://www.amazon.de/dp/B000000001"), "DE")
        self.assertEqual(monitor.marketplace_from_url("https://www.amazon.com.au/dp/B000000001"), "AU")
        self.assertEqual(monitor.marketplace_from_url("https://www.amazon.fr/dp/B000000001"), "FR")

    def test_legacy_workbook_defaults_to_us_without_manual_edit(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_dir = Path(temp_dir) / "LegacyProject" / "1_输入需求信息"
            input_dir.mkdir(parents=True)
            path = input_dir / "LegacyProject-ASIN检查基础信息.xlsx"
            wb = Workbook()
            ws = wb.active
            ws.title = "ASIN清单"
            ws.append(["父ASIN", "子ASIN", "子SKU", "父ASIN网址", "子ASIN网址", "是否启用检查", "备注"])
            ws.append(["B000000001", "B000000002", "SKU", "", "", "是", ""])
            wb.save(path)
            item = monitor.read_items(path)[0]
            self.assertEqual(item["站点"], "US")
            self.assertEqual(item["子ASIN网址"], "https://www.amazon.com/dp/B000000002")

    def test_url_domain_overrides_legacy_us_default(self):
        item = {"站点": "UK", "子ASIN": "B000000001"}
        self.assertEqual(monitor.marketplace_profile(item)["domain"], "www.amazon.co.uk")
        self.assertEqual(monitor.item_postal_codes(item), ["M1 1AE"])

    def test_snapshots_are_isolated_by_marketplace(self):
        us = {"站点": "US", "子ASIN": "B000000001"}
        de = {"站点": "DE", "子ASIN": "B000000001"}
        self.assertEqual(monitor.snapshot_key(us), "US:B000000001")
        self.assertEqual(monitor.snapshot_key(de), "DE:B000000001")
        self.assertNotEqual(monitor.snapshot_key(us), monitor.snapshot_key(de))

    def test_european_currency_values_are_parsed(self):
        self.assertEqual(monitor.parse_money_value("19,99 €"), 19.99)
        self.assertEqual(monitor.parse_money_value("1.299,99 €"), 1299.99)
        self.assertEqual(monitor.parse_money_value("$1,299.99"), 1299.99)
        self.assertEqual(monitor.first_money("Preis: 19,99 €"), "19,99€")

    def test_german_multi_buy_discount_is_parsed(self):
        text = (
            "Spare 5% bei 2 ausgewählten Artikeln Weitere Artikel "
            "Spare 8% bei 3 ausgewählten Artikeln Weitere Artikel "
            "Spare 10% bei 5 ausgewählten Artikeln"
        )
        self.assertEqual(
            monitor.multi_buy_clean(text),
            "Spare 5% bei 2 ausgewählten Artikeln；Spare 8% bei 3 ausgewählten Artikeln；Spare 10% bei 5 ausgewählten Artikeln",
        )

    def test_german_star_distribution_is_parsed(self):
        text = "5 Sterne 70 % 4 Sterne 20 % 3 Sterne 5 % 2 Sterne 3 % 1 Stern 2 %"
        self.assertEqual(
            monitor.parse_star_distribution(text),
            {"5星": 70.0, "4星": 20.0, "3星": 5.0, "2星": 3.0, "1星": 2.0},
        )

    def test_visual_strike_evidence_requires_reference_above_current(self):
        class Page:
            def evaluate(self, _script):
                return {"reference_price": "244,85 €", "source": "visual_line_through", "crossed": True}

        evidence = monitor.extract_visual_strike_evidence(Page(), "234,18 €")
        self.assertTrue(evidence["visual_strike_detected"])
        self.assertEqual(evidence["reference_price"], "244,85 €")

        invalid = monitor.extract_visual_strike_evidence(Page(), "300,00 €")
        self.assertFalse(invalid["visual_strike_detected"])

    def test_relationship_extraction_uses_explicit_variation_map(self):
        raw_html = '''
        <script>
        {"parentAsin":"B0PARENT01","dimensionValuesDisplayData":{
          "B0CHILD001":["Black"],"B0CHILD002":["White"]}}
        </script>
        <div data-asin="B0RECOMM01">recommendation</div>
        '''
        result = monitor.extract_relationship_from_html(raw_html, "B0CHILD001")
        self.assertEqual(result["observed_parent_asin"], "B0PARENT01")
        self.assertEqual(result["observed_variant_asins"], ["B0CHILD001", "B0CHILD002"])
        self.assertTrue(result["variant_set_complete"])
        self.assertNotIn("B0RECOMM01", result["observed_variant_asins"])

    def test_compare_reports_parent_and_variant_changes(self):
        item = {"站点": "DE", "父ASIN": "B0PARENT02", "子ASIN": "B0CHILD001", "备注": ""}
        previous = {
            "DE:B0CHILD001": {
                "parent": "B0PARENT01",
                "observed_parent_asin": "B0PARENT01",
                "observed_variant_asins": ["B0CHILD001", "B0CHILD002"],
                "variant_set_complete": True,
            }
        }
        current = {
            "status": "OK", "title": "Mirror", "rating": "4.5", "reviews": "10",
            "category": "Home", "rank": "#1 in Home", "buybox": "有购物车",
            "observed_parent_asin": "B0PARENT02",
            "observed_variant_asins": ["B0CHILD001", "B0CHILD003"],
            "variant_set_complete": True,
        }
        issues, _ = monitor.compare(item, current, previous)
        modules = [row["问题模块"] for row in issues]
        self.assertIn("父子体关系变化", modules)
        self.assertIn("子体关联变化", modules)


if __name__ == "__main__":
    unittest.main()
