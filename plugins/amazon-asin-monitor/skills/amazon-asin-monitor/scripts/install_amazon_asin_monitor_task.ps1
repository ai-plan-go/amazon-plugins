$ErrorActionPreference = "Stop"

$taskName = "Amazon ASIN Frontend Monitor 0800"
$scriptPath = "D:\Codex\run_asin_check.ps1"
$taskRun = "powershell.exe -ExecutionPolicy Bypass -File `"$scriptPath`""

schtasks.exe /Create /TN $taskName /SC DAILY /ST 08:00 /TR $taskRun /F | Out-Null
Write-Host "Installed scheduled task: $taskName"
