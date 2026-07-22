# Amazon ASIN Monitor Scripts

- `amazon_frontend_check.py`: production monitor deployed to `D:\Codex\amazon_frontend_check.py`.
- `deploy_amazon_asin_monitor.ps1`: copies and compiles the packaged monitor.
- `install_amazon_asin_monitor_task.ps1`: guard script that refuses Windows Task Scheduler installation.

Runtime dependencies:

```powershell
python -m pip install openpyxl pillow playwright
python -m playwright install chromium
```

Run tests from the plugin root:

```powershell
$env:PYTHONPATH = (Resolve-Path scripts)
python -m unittest discover -s tests -v
```
