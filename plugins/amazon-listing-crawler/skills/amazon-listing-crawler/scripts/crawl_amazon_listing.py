#!/usr/bin/env python3
"""Crawl one Amazon.com product page with an explicit US browser context."""

from __future__ import annotations

import argparse
import asyncio
import json
import random
import re
import sys
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse


ASIN_RE = re.compile(r"^[A-Z0-9]{10}$", re.IGNORECASE)
ASIN_IN_URL_RE = re.compile(r"/(?:dp|gp/product|product)/([A-Z0-9]{10})(?:[/?]|$)", re.IGNORECASE)
BLOCK_PATTERNS = {
    "captcha": (
        "enter the characters you see below",
        "type the characters you see in this image",
        "click the button below to continue shopping",
    ),
    "robot_check": ("sorry, we just need to make sure you're not a robot", "robot check"),
    "access_denied": ("access denied", "api-services-support@amazon.com"),
    "request_blocked": ("your request has been blocked", "automated access to amazon data"),
}


@dataclass
class CrawlConfig:
    target: str
    output_dir: Path
    postal_code: str = "10001"
    timezone_id: str = "America/New_York"
    locale: str = "en-US"
    currency: str = "USD"
    timeout_ms: int = 45_000
    retries: int = 2
    headed: bool = False
    screenshot: bool = True
    proxy: str | None = None
    user_data_dir: Path | None = None
    browser_channel: str | None = "chrome"
    manual_challenge_timeout: int = 0


def parse_target(target: str) -> tuple[str, str]:
    value = target.strip()
    if ASIN_RE.fullmatch(value):
        asin = value.upper()
        return asin, f"https://www.amazon.com/dp/{asin}"

    parsed = urlparse(value)
    if parsed.scheme not in {"http", "https"}:
        raise ValueError("Target must be a 10-character ASIN or an http(s) Amazon product URL")
    host = (parsed.hostname or "").lower()
    if host != "amazon.com" and not host.endswith(".amazon.com"):
        raise ValueError("Only amazon.com URLs are accepted; pass an ASIN to normalize automatically")
    match = ASIN_IN_URL_RE.search(parsed.path + ("/" if not parsed.path.endswith("/") else ""))
    if not match:
        raise ValueError("Could not find a 10-character ASIN in the Amazon URL")
    asin = match.group(1).upper()
    return asin, f"https://www.amazon.com/dp/{asin}"


def asin_from_url(url: str) -> str | None:
    parsed = urlparse(url)
    match = ASIN_IN_URL_RE.search(parsed.path + ("/" if not parsed.path.endswith("/") else ""))
    return match.group(1).upper() if match else None


def clean_text(value: str | None) -> str | None:
    if value is None:
        return None
    cleaned = re.sub(r"\s+", " ", value).strip()
    return cleaned or None


def detect_block(title: str, body: str, url: str) -> str | None:
    sample = f"{title}\n{body[:30_000]}\n{url}".lower()
    for reason, patterns in BLOCK_PATTERNS.items():
        if any(pattern in sample for pattern in patterns):
            return reason
    if "/errors/validatecaptcha" in url.lower():
        return "captcha"
    return None


async def current_block_reason(page: Any) -> str | None:
    try:
        title = await page.title()
        body = await page.locator("body").inner_text(timeout=5_000)
        return detect_block(title, body, page.url)
    except Exception:
        return "page_unavailable"


async def wait_for_manual_unblock(page: Any, timeout_seconds: int) -> bool:
    """Wait for the user to complete Amazon's challenge in a headed browser."""
    deadline = asyncio.get_running_loop().time() + timeout_seconds
    while asyncio.get_running_loop().time() < deadline:
        if await current_block_reason(page) is None:
            return True
        await page.wait_for_timeout(1_000)
    return False


async def first_text(page: Any, selectors: list[str]) -> str | None:
    for selector in selectors:
        locator = page.locator(selector).first
        try:
            if await locator.count() and await locator.is_visible(timeout=750):
                value = clean_text(await locator.inner_text(timeout=1_500))
                if value:
                    return value
        except Exception:
            continue
    return None


async def all_text(page: Any, selector: str) -> list[str]:
    try:
        values = await page.locator(selector).all_inner_texts()
    except Exception:
        return []
    return [text for value in values if (text := clean_text(value))]


async def attr(page: Any, selectors: list[str], name: str) -> str | None:
    for selector in selectors:
        try:
            value = await page.locator(selector).first.get_attribute(name, timeout=1_500)
            if value:
                return value
        except Exception:
            continue
    return None


async def resolved_page_asin(page: Any, fallback: str) -> str:
    selected = await attr(page, ["input#ASIN", "input[name='ASIN']"], "value")
    if selected and ASIN_RE.fullmatch(selected):
        return selected.upper()
    data_asin = await attr(
        page,
        ["#averageCustomerReviews[data-asin]", "#all-offers-display-params[data-asin]"],
        "data-asin",
    )
    if data_asin and ASIN_RE.fullmatch(data_asin):
        return data_asin.upper()
    return asin_from_url(page.url) or fallback


async def json_ld(page: Any) -> list[dict[str, Any]]:
    documents: list[dict[str, Any]] = []
    try:
        blocks = await page.locator('script[type="application/ld+json"]').all_text_contents()
    except Exception:
        return documents
    for block in blocks:
        try:
            parsed = json.loads(block)
        except (json.JSONDecodeError, TypeError):
            continue
        items = parsed if isinstance(parsed, list) else [parsed]
        for item in items:
            if not isinstance(item, dict):
                continue
            graph = item.get("@graph")
            if isinstance(graph, list):
                documents.extend(node for node in graph if isinstance(node, dict))
            documents.append(item)
    return documents


def find_product_schema(documents: list[dict[str, Any]]) -> dict[str, Any]:
    for document in documents:
        schema_type = document.get("@type")
        types = schema_type if isinstance(schema_type, list) else [schema_type]
        if "Product" in types:
            return document
    return {}


async def set_postal_code(page: Any, postal_code: str) -> bool:
    """Set the Amazon delivery ZIP through the public page UI."""
    try:
        trigger = page.locator("#nav-global-location-popover-link").first
        if not await trigger.count():
            return False
        await trigger.click(timeout=5_000)
        zip_input = page.locator("#GLUXZipUpdateInput, input[data-action='GLUXPostalInputAction']").first
        await zip_input.wait_for(state="visible", timeout=5_000)
        await zip_input.fill(postal_code)
        submit = page.locator("#GLUXZipUpdate input[type='submit'], #GLUXZipUpdate").first
        await submit.click(timeout=5_000)
        try:
            await page.locator("#GLUXConfirmClose, button[name='glowDoneButton']").first.click(timeout=3_000)
        except Exception:
            pass
        await page.wait_for_timeout(1_000)
        return True
    except Exception:
        return False


async def extract_product(page: Any, asin: str) -> dict[str, Any]:
    schema = find_product_schema(await json_ld(page))
    offers = schema.get("offers") if isinstance(schema.get("offers"), dict) else {}
    aggregate = schema.get("aggregateRating") if isinstance(schema.get("aggregateRating"), dict) else {}
    brand = schema.get("brand")
    if isinstance(brand, dict):
        brand = brand.get("name")

    price = await first_text(page, [
        "#corePrice_feature_div .a-offscreen",
        "#corePriceDisplay_desktop_feature_div .a-offscreen",
        "#priceblock_ourprice",
        "#priceblock_dealprice",
        ".apexPriceToPay .a-offscreen",
    ])
    rating_text = await first_text(page, ["#acrPopover", "span[data-hook='rating-out-of-text']"])
    review_text = await first_text(page, ["#acrCustomerReviewText", "span[data-hook='total-review-count']"])
    image = await attr(page, ["#landingImage", "#imgBlkFront"], "src")
    bullets = await all_text(page, "#feature-bullets li span.a-list-item")
    unavailable = await first_text(page, ["#availability", "#outOfStock"])

    return {
        "asin": asin,
        "title": await first_text(page, ["#productTitle", "h1.a-size-large"]) or clean_text(schema.get("name")),
        "brand": await first_text(page, ["#bylineInfo", "a#brand"]) or (clean_text(str(brand)) if brand else None),
        "price": price or clean_text(str(offers.get("price"))) if offers.get("price") is not None else price,
        "currency": offers.get("priceCurrency"),
        "availability": unavailable or clean_text(str(offers.get("availability", "")).rsplit("/", 1)[-1]),
        "rating": rating_text or clean_text(str(aggregate.get("ratingValue"))) if aggregate else rating_text,
        "review_count": review_text or clean_text(str(aggregate.get("reviewCount"))) if aggregate else review_text,
        "bullet_points": bullets,
        "primary_image": image or schema.get("image"),
        "seller": await first_text(page, ["#sellerProfileTriggerId", "#merchant-info"]),
        "buy_box": await first_text(page, ["#desktop_buybox", "#buybox"]),
    }


async def runtime_context(page: Any, config: CrawlConfig, postal_applied: bool) -> dict[str, Any]:
    browser_values = await page.evaluate("""() => ({
        timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
        locale: Intl.DateTimeFormat().resolvedOptions().locale,
        date: new Date().toString(),
        language: navigator.language,
        languages: navigator.languages,
        userAgent: navigator.userAgent
    })""")
    return {
        **browser_values,
        "requested_timezone": config.timezone_id,
        "requested_locale": config.locale,
        "requested_currency": config.currency,
        "postal_code": config.postal_code,
        "postal_code_applied": postal_applied,
        "proxy_configured": bool(config.proxy),
    }


async def save_evidence(page: Any, directory: Path, screenshot: bool) -> None:
    directory.mkdir(parents=True, exist_ok=True)
    try:
        (directory / "page.html").write_text(await page.content(), encoding="utf-8")
    except Exception:
        pass
    if screenshot:
        try:
            await page.screenshot(path=str(directory / "page.png"), full_page=True, timeout=15_000)
        except Exception:
            pass


async def crawl(config: CrawlConfig) -> tuple[dict[str, Any], int]:
    try:
        from playwright.async_api import async_playwright
    except ImportError as exc:
        raise RuntimeError(
            "Playwright is required. Run: python -m pip install playwright && "
            "python -m playwright install chromium"
        ) from exc

    asin, canonical_url = parse_target(config.target)
    config.output_dir.mkdir(parents=True, exist_ok=True)
    result: dict[str, Any] = {
        "status": "failed",
        "asin": asin,
        "requested_url": config.target,
        "canonical_url": canonical_url,
        "started_at": datetime.now(timezone.utc).isoformat(),
        "attempts": [],
        "blocked": False,
        "block_reason": None,
    }

    async with async_playwright() as playwright:
        launch_args: dict[str, Any] = {
            "headless": not config.headed,
            "args": ["--lang=en-US"],
        }
        if config.browser_channel:
            launch_args["channel"] = config.browser_channel
        if config.proxy:
            launch_args["proxy"] = {"server": config.proxy}

        context_args: dict[str, Any] = {
            "locale": config.locale,
            "timezone_id": config.timezone_id,
            "geolocation": {"latitude": 40.7506, "longitude": -73.9972},
            "permissions": ["geolocation"],
            "viewport": {"width": 1440, "height": 1100},
            "screen": {"width": 1440, "height": 1100},
            "color_scheme": "light",
            "extra_http_headers": {
                "Accept-Language": "en-US,en;q=0.9",
                "DNT": "1",
            },
        }

        if config.user_data_dir:
            context = await playwright.chromium.launch_persistent_context(
                str(config.user_data_dir), **launch_args, **context_args
            )
            browser = None
        else:
            browser = await playwright.chromium.launch(**launch_args)
            context = await browser.new_context(**context_args)

        await context.add_cookies([
            {"name": "lc-main", "value": "en_US", "domain": ".amazon.com", "path": "/"},
            {"name": "i18n-prefs", "value": config.currency, "domain": ".amazon.com", "path": "/"},
        ])
        page = context.pages[0] if context.pages else await context.new_page()
        page.set_default_timeout(config.timeout_ms)
        page.set_default_navigation_timeout(config.timeout_ms)

        try:
            postal_applied = False
            for attempt in range(1, config.retries + 2):
                attempt_dir = config.output_dir / f"attempt-{attempt}"
                record: dict[str, Any] = {"attempt": attempt, "started_at": datetime.now(timezone.utc).isoformat()}
                try:
                    response = await page.goto(canonical_url, wait_until="domcontentloaded")
                    record["http_status"] = response.status if response else None
                    await page.wait_for_timeout(1_500)

                    title = await page.title()
                    body = await page.locator("body").inner_text(timeout=5_000)
                    block_reason = detect_block(title, body, page.url)
                    record.update({"final_url": page.url, "title": title, "block_reason": block_reason})
                    await save_evidence(page, attempt_dir, config.screenshot)

                    if block_reason:
                        result.update({"blocked": True, "block_reason": block_reason})
                        result["runtime_context"] = await runtime_context(page, config, postal_applied)
                        result["session"] = {
                            "persistent_profile": bool(config.user_data_dir),
                            "browser_channel": config.browser_channel or "bundled-chromium",
                            "manual_challenge_enabled": bool(config.manual_challenge_timeout),
                        }
                        if config.manual_challenge_timeout:
                            print(
                                "Amazon challenge detected. Complete it in the open browser; "
                                f"waiting up to {config.manual_challenge_timeout} seconds...",
                                file=sys.stderr,
                                flush=True,
                            )
                            resolved = await wait_for_manual_unblock(page, config.manual_challenge_timeout)
                            record["manual_challenge_resolved"] = resolved
                            if resolved:
                                result.update({"blocked": False, "block_reason": None})
                                await page.goto(canonical_url, wait_until="domcontentloaded")
                                await page.wait_for_timeout(1_000)
                                followup_reason = await current_block_reason(page)
                                if followup_reason:
                                    result.update({"blocked": True, "block_reason": followup_reason})
                                    record["followup_block_reason"] = followup_reason
                                else:
                                    await save_evidence(page, config.output_dir / "manual-unblocked", config.screenshot)
                            if result["blocked"]:
                                result["attempts"].append(record)
                                result["status"] = "blocked"
                                break
                        else:
                            result["attempts"].append(record)
                            result["status"] = "blocked"
                            break

                    if not postal_applied:
                        postal_applied = await set_postal_code(page, config.postal_code)
                        if postal_applied:
                            await page.goto(canonical_url, wait_until="domcontentloaded")
                            await page.wait_for_timeout(1_000)

                    product = await extract_product(page, asin)
                    resolved_asin = await resolved_page_asin(page, asin)
                    product["asin"] = resolved_asin
                    required = ["title", "price", "primary_image"]
                    missing = [field for field in required if not product.get(field)]
                    result.update({
                        "status": "ok" if not missing else "partial",
                        "resolved_asin": resolved_asin,
                        "product": product,
                        "missing_fields": missing,
                        "final_url": page.url,
                        "runtime_context": await runtime_context(page, config, postal_applied),
                    })
                    result["attempts"].append(record)
                    break
                except Exception as exc:
                    record.update({"error_type": type(exc).__name__, "error": str(exc)})
                    result["attempts"].append(record)
                    await save_evidence(page, attempt_dir, config.screenshot)
                    if attempt <= config.retries:
                        await asyncio.sleep((2 ** (attempt - 1)) + random.uniform(0.5, 1.5))
                        continue
                    result["error"] = record["error"]
        finally:
            await save_evidence(page, config.output_dir, config.screenshot)
            await context.close()
            if browser:
                await browser.close()

    result["finished_at"] = datetime.now(timezone.utc).isoformat()
    output_path = config.output_dir / "result.json"
    output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    exit_code = {"ok": 0, "partial": 2, "blocked": 3}.get(result["status"], 1)
    return result, exit_code


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("target", help="10-character ASIN or amazon.com product URL")
    parser.add_argument("--output-dir", type=Path, default=Path("amazon-listing-output"))
    parser.add_argument("--postal-code", default="10001", help="US delivery ZIP code (default: 10001)")
    parser.add_argument("--timezone", default="America/New_York", help="IANA timezone (default: America/New_York)")
    parser.add_argument("--timeout-ms", type=int, default=45_000)
    parser.add_argument("--retries", type=int, default=2)
    parser.add_argument("--headed", action="store_true", help="Show Chromium for interactive diagnosis")
    parser.add_argument("--no-screenshot", action="store_true")
    parser.add_argument("--proxy", help="Authorized proxy URL, for example http://host:port")
    parser.add_argument(
        "--user-data-dir",
        type=Path,
        default=Path.home() / ".amazon-listing-crawler" / "chrome-profile",
        help="Dedicated persistent profile directory",
    )
    parser.add_argument(
        "--browser-channel",
        choices=["chrome", "msedge", "chromium"],
        default="chrome",
        help="Installed browser channel (default: chrome)",
    )
    parser.add_argument(
        "--manual-challenge-timeout",
        type=int,
        default=0,
        metavar="SECONDS",
        help="Wait in a headed browser for manual challenge completion",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if not re.fullmatch(r"\d{5}(?:-\d{4})?", args.postal_code):
        print("error: --postal-code must be a US ZIP code", file=sys.stderr)
        return 1
    if args.retries < 0 or args.retries > 5:
        print("error: --retries must be between 0 and 5", file=sys.stderr)
        return 1
    if args.manual_challenge_timeout < 0 or args.manual_challenge_timeout > 900:
        print("error: --manual-challenge-timeout must be between 0 and 900", file=sys.stderr)
        return 1
    if args.manual_challenge_timeout and not args.headed:
        print("error: --manual-challenge-timeout requires --headed", file=sys.stderr)
        return 1
    if args.manual_challenge_timeout and not args.user_data_dir:
        print("error: manual challenge handling requires --user-data-dir to preserve the session", file=sys.stderr)
        return 1

    config = CrawlConfig(
        target=args.target,
        output_dir=args.output_dir.resolve(),
        postal_code=args.postal_code,
        timezone_id=args.timezone,
        timeout_ms=args.timeout_ms,
        retries=args.retries,
        headed=args.headed,
        screenshot=not args.no_screenshot,
        proxy=args.proxy,
        user_data_dir=args.user_data_dir.resolve() if args.user_data_dir else None,
        browser_channel=args.browser_channel,
        manual_challenge_timeout=args.manual_challenge_timeout,
    )
    try:
        result, exit_code = asyncio.run(crawl(config))
    except Exception as exc:
        config.output_dir.mkdir(parents=True, exist_ok=True)
        failure = {
            "status": "failed",
            "requested_url": config.target,
            "blocked": False,
            "error_type": type(exc).__name__,
            "error": str(exc),
            "finished_at": datetime.now(timezone.utc).isoformat(),
        }
        (config.output_dir / "result.json").write_text(
            json.dumps(failure, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        print(f"error: {exc}", file=sys.stderr)
        print(f"diagnostic: {config.output_dir / 'result.json'}", file=sys.stderr)
        return 1
    print(json.dumps({
        "status": result["status"],
        "asin": result["asin"],
        "output": str(config.output_dir / "result.json"),
        "blocked": result["blocked"],
    }, ensure_ascii=False))
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
