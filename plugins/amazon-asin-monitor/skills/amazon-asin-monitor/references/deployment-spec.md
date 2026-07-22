# Deployment Specification

## Fixed paths

- Workspace: `D:\Codex`
- Main script: `D:\Codex\amazon_frontend_check.py`
- Project root: `D:\Codex\各项目链接检查`
- Root run summary: `D:\Codex\last_run_summary.json`

Each project uses exactly:

- `D:\Codex\各项目链接检查\<项目名称>\1_输入需求信息`
- `D:\Codex\各项目链接检查\<项目名称>\2_输出信息`
- `D:\Codex\各项目链接检查\<项目名称>\3_系统运行缓存`

Human-readable workbooks must only be written to `2_输出信息`. Baselines, raw results, batch checkpoints, and resume state belong in `3_系统运行缓存`.

## Input discovery

Discover `*-ASIN检查基础信息.xlsx` files from configured desktop locations and existing project input directories. The preferred input sheet is `ASIN清单`; otherwise use the first sheet. Missing child URLs are composed as `https://www.amazon.com/dp/<子ASIN>`.

## Execution controls

- Default delivery postal code: `10043`.
- Run projects sequentially and require the current project's quality gate to pass before starting the next.
- Stop during preflight when Amazon is inaccessible.
- Stop a project after the configured consecutive blocked-page threshold.
- Persist batch progress so interrupted runs can resume without repeating completed batches.
- Allow bounded batch concurrency inside one project only; the script caps workers at 3.
- Emit timestamped console and `D:\Codex\amazon_monitor.log` progress messages.

## Output behavior

The monitor writes dated detail and exception sheets, embeds screenshot evidence, and saves an `_updated_<timestamp>.xlsx` variant when the normal workbook is locked by Excel. The run summary records workbook paths, latest sheets, counts, postal-code-dependent Buy Box results, batch timing, failures, and quality-gate status.

## Scheduling and notifications

Schedule daily 08:00 runs through Codex automation using timezone `Asia/Shanghai`. Windows Task Scheduler is prohibited. WeCom notification and workbook upload are intentionally excluded because the delivery channel's file-size limit is not suitable for these reports.
