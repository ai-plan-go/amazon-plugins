---
name: amazon-asin-monitor
description: Deploy, run, resume, or audit the Amazon ASIN frontend monitor and its Excel reports. Use when the user asks for Amazon ASIN巡检, 前台监控, Buy Box或购物车检查, ASIN项目批量检查, 排名与折扣核验, Excel截图证据, PDCA复盘, or the daily 08:00 Codex automation.
---

# Amazon ASIN Monitor

Use the bundled script for deterministic collection. Do not replace a failed script run with subjective browser inspection.

## Read first

- Read `references/deployment-spec.md` before deploying or discovering projects.
- Read `references/check-rules.md` before interpreting fields or auditing a workbook.

## Deploy

From the plugin root, deploy the packaged files:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/deploy_amazon_asin_monitor.ps1
```

The deployment target is `D:\Codex\amazon_frontend_check.py`. The plugin does not create a Windows Task Scheduler task and does not include WeCom notification or Webhook code.

## Run

Run the fixed production script:

```powershell
python D:\Codex\amazon_frontend_check.py
```

Prefer the script's preflight, project discovery, progress logs, resume files, and quality gates. Projects run sequentially: a project must pass before the next project starts. Batches inside one project may use bounded concurrency when configured.

Useful environment controls:

```powershell
$env:AMAZON_POSTAL_CODES = "10043"
$env:ASIN_PROJECT_START = "A27小家纺"
$env:ASIN_PROJECT_FILTER = "A27小家纺"
$env:ASIN_PARENT_FILTER = "B0EXAMPLE01"
$env:ASIN_CHILD_FILTER = "B0EXAMPLE02,B0EXAMPLE03"
$env:ASIN_BATCH_SIZE = "25"
$env:ASIN_BATCH_WORKERS = "2"
```

Do not rerun completed projects when the user asks to continue. Use `ASIN_PROJECT_START`, filtering, and existing batch progress to resume only pending work.

## Check

After a run, read `D:\Codex\last_run_summary.json` if present. For every project, report:

- Output workbook path under `D:\Codex\各项目链接检查\<项目名称>\2_输出信息`.
- Latest sheet name and exception count.
- Whether an `_updated` workbook was generated.
- Any ASIN whose Buy Box result depended on postal code `10043`.
- Failure stage and reason if the script stopped.

Confirm field provenance explicitly: category nodes come only from top breadcrumb selectors; big and small category ranks come only from Best Sellers Rank; Excel embeds star-distribution evidence and price screenshots.

## Failure behavior

If preflight or collection cannot access Amazon, stop immediately and report the script failure stage. Do not fabricate results, do not continue with ineffective jobs, and do not use AI-driven page opening as a substitute for script output.

## Scheduling

Use Codex automation for a daily 08:00 `Asia/Shanghai` run. Never create Windows Task Scheduler tasks.
