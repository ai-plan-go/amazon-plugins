# Collection and self-check rules

## Marketplace isolation

Resolve domain, language, currency, and postal code per row. Do not reuse US page assumptions for UK, Germany, Australia, or France. Preserve a separate session context where the collector requires one.

## Buy Box

Check Add to Cart, Buy Now, and buybox container signals independently. Price presence is not evidence of a Buy Box. Record the marketplace postal code used for any location-dependent decision.

## Category and ranking

- `类目节点`: top breadcrumb only, using `#wayfinding-breadcrumbs_feature_div`, `#wayfinding-breadcrumbs_container`, or `.a-breadcrumb`.
- `大类排名` and `小类排名`: Best Sellers Rank only.
- Never derive category nodes from Best Sellers Rank.

## Current, reference, and strike-through prices

Parse marketplace-aware decimal and thousands separators. Require the reference price to be greater than the current price.

Recognize both localized labels and visual markup, including:

- English: `List Price`, `Was`, `RRP`
- German: `Statt`, `UVP`
- French: `Prix conseillé`, `Ancien prix`, `Prix de référence`
- Visual selectors and attributes such as `.a-text-price`, line-through styling, and `data-a-strike`

The `是否有划线` result must agree with the visible red reference-price evidence in the screenshot. If text extraction and visual evidence conflict, fail the row self-check instead of silently reporting `否`.

## Promotions

Capture coupons, Prime-exclusive discounts, and multi-buy offers separately. German multi-buy wording such as `Spare 5 % bei 2 ausgewählten Artikeln` is a promotion even when no coupon badge exists.

## Reviews and evidence

Open the rating popover and support English and localized labels, including German `5 Sterne`, `4 Sterne`, and equivalents. Create stable resized evidence images before embedding. The output workbook should contain both star-distribution and price screenshots when the page exposes them.

## Variation relationships

Compare the configured parent ASIN, the parent or family observed on the page, and the current child set against the previous valid baseline. Detect and report:

- parent ASIN changed;
- child moved to another parent;
- child added to or removed from a family;
- configured parent differs from observed parent;
- child-set association changed.

Use structured page data such as variation maps or `dimensionValuesDisplayData` when available. Do not replace a valid baseline with incomplete data from a blocked, consent-only, or failed page.

## Row self-check

For every enabled row, validate at least:

- marketplace, domain, currency, and postal code are consistent;
- current price and reference price are numerically plausible;
- strike-through result agrees with collected visual or localized evidence;
- Buy Box result came from purchase controls rather than price;
- breadcrumb and ranking fields use their required independent sources;
- relationship data is complete enough to update the baseline;
- expected screenshots are present and embedded.

Expose failures in the output and summary so a wrong-looking row cannot pass silently.
