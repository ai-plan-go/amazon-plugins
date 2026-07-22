# Amazon Plugins

面向 Amazon 运营场景的 Codex 插件集合。

## Amazon Listing Crawler

`amazon-listing-crawler` 用于根据 ASIN 或 Amazon.com 商品链接抓取 Listing 页面，并生成结构化商品数据与诊断证据。

主要能力：

- 使用系统 Chrome 与专用持久化浏览器配置，降低临时无状态浏览器导致的异常。
- 固定 `America/New_York` 时区、`en-US` 语言、USD 币种和美国配送邮编。
- 提取标题、品牌、价格、库存、评分、评论数、Bullet Points、主图、卖家和 Buy Box。
- 保存 `result.json`、页面 HTML 和完整页面截图，便于检查和复盘。
- 识别 CAPTCHA、Robot Check、Access Denied 和 Amazon 软阻断页面。
- 支持用户在可视 Chrome 中人工完成挑战，放行后自动继续抓取并复用会话。
- 记录请求 ASIN 与实际落地 ASIN，避免 Amazon 重定向造成数据归属错误。

### 安装

添加 Marketplace 并安装插件：

```powershell
codex plugin marketplace add https://github.com/ai-plan-go/amazon-plugins
codex plugin add amazon-listing-crawler@amazon-plugins
```

安装后新建一个 Codex 任务，通过 ASIN 或 Amazon.com 商品链接调用：

```text
$amazon-listing-crawler B0H6XH37MV
```

### 首次运行依赖

脚本需要系统已安装 Chrome，并安装 Python Playwright：

```powershell
python -m pip install playwright
```

### 处理 blocked

如果 `result.json` 返回 `status: blocked`，使用可视模式启动一次：

```powershell
python scripts/crawl_amazon_listing.py B0H6XH37MV `
  --headed `
  --manual-challenge-timeout 300 `
  --output-dir output
```

在打开的 Chrome 中人工完成 Amazon 提示。脚本检测到挑战解除后会自动重新加载 Listing 并继续提取。会话默认保存在 `~/.amazon-listing-crawler/chrome-profile`，后续任务会自动复用。

插件不会自动绕过 CAPTCHA，也不会对挑战页持续重试。时区和语言设置不能改变出口 IP 地区；如果 Amazon 持续按 IP 阻断，应使用合规且经过授权的美国网络出口。

### 输出

- `result.json`：抓取状态、商品字段、实际 ASIN、运行时区和异常诊断。
- `page.html`：最终页面原始 HTML。
- `page.png`：最终页面完整截图。
- `attempt-N/`：每次请求的页面证据。

状态码：`ok` 表示完整成功，`partial` 表示部分字段缺失，`blocked` 表示 Amazon 挑战或访问阻断，`failed` 表示浏览器或网络异常。

## English

The `amazon-listing-crawler` plugin collects structured Amazon.com Listing data by ASIN or URL. It uses a persistent Chrome session configured for `America/New_York`, `en-US`, USD, and a US delivery ZIP code. It saves JSON, HTML, and screenshot evidence and supports user-assisted recovery from Amazon challenge pages without bypassing them.

Install the marketplace and plugin:

```powershell
codex plugin marketplace add https://github.com/ai-plan-go/amazon-plugins
codex plugin add amazon-listing-crawler@amazon-plugins
```

Then start a new Codex task and invoke `$amazon-listing-crawler` with an ASIN or Amazon.com product URL.

The bundled Python script requires Playwright and an installed Chrome browser:

```powershell
python -m pip install playwright
```

## Amazon ASIN Monitor

`amazon-asin-monitor` 用于从 Excel 项目清单执行 Amazon ASIN 前台巡检，并生成内嵌截图证据的检查工作簿。

主要能力：

- 按项目顺序执行，当前项目通过质量门禁后才启动下一个项目。
- 支持项目续跑、父体/子体过滤、分批检查、断点恢复和受控批次并发。
- 默认按美国邮编 `10043` 独立检查 Add to Cart、Buy Now 和 Buy Box 信号，不从价格推断购物车。
- 区分类目 breadcrumb 与 Best Sellers Rank 的字段来源。
- 识别 Prime 会员专享价、Regular Price、Coupon 和多买折扣。
- 在 Excel 中嵌入星级占比图和价格截图，并输出异常汇总与运行进度日志。
- 不包含企业微信推送，不创建 Windows Task Scheduler 任务。

安装：

```powershell
codex plugin marketplace add https://github.com/ai-plan-go/amazon-plugins
codex plugin add amazon-asin-monitor@amazon-plugins
```

安装后可在新任务中调用：

```text
$amazon-asin-monitor 运行全部待检查项目
```

依赖 Python、Playwright、openpyxl 和 Pillow。生产脚本默认部署到 `D:\Codex\amazon_frontend_check.py`，项目数据位于 `D:\Codex\各项目链接检查`。
