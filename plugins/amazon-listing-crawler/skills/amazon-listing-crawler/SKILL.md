---
name: amazon-listing-crawler
description: Crawl Amazon.com product detail pages with a Playwright browser configured for the United States, including America/New_York timezone, en-US locale, USD preferences, a US delivery ZIP code, persistent sessions, and user-assisted challenge recovery. Use when Codex needs to collect an Amazon Listing by ASIN or URL, recover from recurring blocked or CAPTCHA responses without bypassing them, diagnose crawl failures, or save structured product data together with HTML and screenshot evidence.
---

# Amazon Listing Crawler

Use the bundled script for deterministic crawling. Treat Amazon responses as dynamic: keep evidence, report blocked states explicitly, and never claim missing fields are empty when the page was challenged.

## Run a crawl

Install Playwright once if it is unavailable:

```powershell
python -m pip install playwright
python -m playwright install chromium
```

Run from the skill directory. The script uses installed Chrome and a dedicated persistent profile by default:

```powershell
python scripts/crawl_amazon_listing.py B0EXAMPLE01 --output-dir output
```

Pass either an ASIN or an `amazon.com` product URL. Use `--postal-code` to select another US delivery area. Use `--headed` only for interactive diagnosis; keep scheduled crawls headless. If Chrome is unavailable, pass `--browser-channel chromium` after installing Playwright Chromium.

## Recover a blocked session

When `result.json` reports `blocked`, open a dedicated persistent Chrome session and let the user complete Amazon's challenge once:

```powershell
python scripts/crawl_amazon_listing.py B0EXAMPLE01 `
  --headed `
  --manual-challenge-timeout 300 `
  --output-dir output
```

Keep the terminal and browser open. Complete the visible Amazon prompt manually. The script detects the cleared challenge, reloads the Listing, and continues extraction. It stores cookies and session state in `~/.amazon-listing-crawler/chrome-profile` for later runs.

Reuse that profile for subsequent headless crawls:

```powershell
python scripts/crawl_amazon_listing.py B0EXAMPLE01 `
  --output-dir output
```

Use a dedicated automation profile. Do not point Playwright at a Chrome profile that is currently open.

## Interpret outputs

Read `result.json` first. Check these fields before consuming product data:

1. Require `status` to be `ok` or `partial`.
2. Check `blocked` and `block_reason`. Do not retry CAPTCHA indefinitely.
3. Verify `runtime_context.timezone` is `America/New_York`, `locale` starts with `en-US`, and `postal_code_applied` reflects the requested ZIP code.
4. Use `missing_fields` to distinguish partial extraction from transport failure.
5. Compare the requested `asin` with `resolved_asin`; Amazon may redirect aliases to another detail ASIN.
6. Inspect `page.html`, `page.png`, and per-attempt evidence when diagnosis is required.

Exit codes are `0` for a complete result, `2` for partial extraction, `3` for a challenge/block, and `1` for other failures.

## Reliability rules

- Crawl at low frequency and comply with Amazon terms, robots directives, and applicable law.
- Prefer a persistent browser profile with `--user-data-dir` when consent or delivery preferences must survive runs.
- Use `--proxy` only with an authorized US endpoint; timezone and locale do not change IP geolocation.
- Do not retry challenge pages automatically. Use the manual headed flow once and reuse its dedicated persistent profile.
- Preserve raw evidence before changing selectors.
- Never implement CAPTCHA bypassing or stealth circumvention.

## Batch usage

Invoke the script once per ASIN from an external queue, add jitter between items, and isolate each item's output directory. Do not launch high-concurrency page floods.
