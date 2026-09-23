---
name: amazon-asin-monitor
description: Deploy, upgrade, run, schedule, or audit the Amazon ASIN frontend monitor for US, UK, Germany, Australia, and France. Use for ASIN frontend inspections, Buy Box checks, localized price and promotion detection, parent-child variation monitoring, Amazon Excel reports, or Codex scheduled ASIN checks.
---

# Amazon ASIN Monitor

Use the packaged Python collector as the source of truth. Never replace a failed scripted run with subjective page inspection.

## Required references

Read these before changing or running the workflow:

- `references/deployment-spec.md` for paths, input format, execution, and scheduling.
- `references/check-rules.md` for field sourcing and self-check rules.

## Packaged resources

- Main collector: `scripts/amazon_frontend_check.py`
- Unit tests: `scripts/test_amazon_frontend_check.py`
- Legacy Windows Task Scheduler cleanup guard: `scripts/install_amazon_asin_monitor_task.ps1`
- Five-market workbook template: `../../assets/Amazon-ASIN检查基础信息模板-五站点.xlsx`

## Workflow

1. Deploy or upgrade the packaged collector to `D:\Codex\amazon_frontend_check.py` without overwriting unrelated project files.
2. Run the packaged unit tests before a production inspection.
3. Run the deployed collector with hidden-window `Start-Process`, redirecting stdout and stderr to the required log files.
4. Read `D:\Codex\last_run_summary.json`, the newest output workbook, and each row's self-check results.
5. Report the workbook path, newest sheet, anomaly count, `_updated` status, marketplace and postal code, relationship changes, and any failed self-checks.
6. Use Codex automation for weekly scheduling. Do not create a Windows Task Scheduler task.

## Non-negotiable rules

- Buy Box must be decided from independent Add to Cart, Buy Now, and buybox signals; never infer it from price.
- Category nodes come only from the top breadcrumb selectors.
- Best Sellers Rank is used only for major-category and subcategory rankings.
- Parent and child relationship changes must be compared with the prior valid baseline.
- Strike-through price requires localized labels or visual strike-price markup and a reference price greater than the current price.
- Embed both the star-distribution image and the price screenshot in Excel when evidence is available.
- Do not overwrite a good relationship baseline with blocked, consent-only, or incomplete page data.

## Failure behavior

If the script fails, report only the failure stage and reason. Do not fabricate results or replace collection with manual browsing.

## Final report

State what ran, which marketplace and postal code were used, the output workbook and sheet, anomaly and self-check totals, relationship changes, Buy Box results, strike-price and promotion results, embedded-image status, and whether an `_updated` file was created.
