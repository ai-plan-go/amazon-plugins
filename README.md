# Amazon Plugins

Codex plugins for Amazon operations.

## Amazon Listing Crawler

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

