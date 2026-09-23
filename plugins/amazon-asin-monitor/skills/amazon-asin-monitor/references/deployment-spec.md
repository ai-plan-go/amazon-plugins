# Deployment and execution specification

## Fixed paths

- Runtime root: `D:\Codex`
- Main script: `D:\Codex\amazon_frontend_check.py`
- Standard output: `D:\Codex\amazon_frontend_check.log`
- Standard error: `D:\Codex\amazon_frontend_check_error.log`
- Summary: `D:\Codex\last_run_summary.json`
- Project input: `D:\Codex\各项目链接检查\<项目名称>\1_基础信息`
- Project output: `D:\Codex\各项目链接检查\<项目名称>\2_输出信息`
- Baseline cache: `D:\Codex\各项目链接检查\<项目名称>\3_系统运行缓存`

Do not write project output or cache into another directory.

## Input workbook

The preferred columns are:

1. 站点
2. 父ASIN
3. 子ASIN
4. 子SKU
5. 子ASIN网址
6. 是否启用检查
7. 备注

The parent URL column is obsolete and should not appear in new templates. The parser may continue reading legacy workbooks containing that column. A blank marketplace defaults to the United States.

Supported marketplace values include Chinese names, country codes, and Amazon domains for US, UK, DE, AU, and FR.

## Hidden execution

Run with PowerShell `Start-Process`, `-WindowStyle Hidden`, and `-Wait`. Redirect both streams and inspect the exit code.

```powershell
$process = Start-Process -FilePath "python" `
  -ArgumentList @("D:\Codex\amazon_frontend_check.py") `
  -WorkingDirectory "D:\Codex" `
  -WindowStyle Hidden `
  -Wait `
  -PassThru `
  -RedirectStandardOutput "D:\Codex\amazon_frontend_check.log" `
  -RedirectStandardError "D:\Codex\amazon_frontend_check_error.log"
```

Background execution does not mean discarding the result. Always read the summary and output workbook after completion.

## Scheduling

Use a Codex automation for Monday 09:15 Asia/Shanghai. Audit existing automations before creating or changing one so duplicate daily checks are not left active. Do not create a Windows Task Scheduler task.

## Upgrade verification

1. Compile the main script.
2. Run all packaged unit tests.
3. Run a controlled workbook inspection.
4. Confirm the summary references the expected workbook.
5. Confirm the newest sheet and row count.
6. Confirm star and price evidence images are embedded where available.
7. Confirm every enabled row has a self-check result.
