param(
    [string]$TargetRoot = "D:\Codex"
)

$ErrorActionPreference = "Stop"
$source = Join-Path $PSScriptRoot "amazon_frontend_check.py"
$target = Join-Path $TargetRoot "amazon_frontend_check.py"

New-Item -ItemType Directory -Force -Path $TargetRoot | Out-Null
Copy-Item -LiteralPath $source -Destination $target -Force
python -m py_compile $target

Write-Host "Deployed: $target"
Write-Host "Scheduling is handled by Codex automation. Windows Task Scheduler was not modified."
Write-Host "WeCom push is not included."
