# Check Rules

## Source-of-truth rules

- `类目节点` comes only from the top breadcrumb area using `#wayfinding-breadcrumbs_feature_div`, `#wayfinding-breadcrumbs_container`, or `.a-breadcrumb` selectors.
- Best Sellers Rank is used only for `大类排名` and `小类排名`; never infer category nodes from it.
- Buy Box is independent from price. Check Add to Cart, Buy Now, desktop buybox, buybox, and buying-options signals. Never infer a cart from a displayed price.
- The default postal-code check is `10043`, and remarks must record when this postal code changes the Buy Box result.
- Excel must embed the star-distribution evidence image and the price screenshot, not merely store file paths.

## Price and promotion rules

- `是否有划线` is true only when a supported main-price module explicitly contains `List Price` or `Was`.
- `Typical Price` is recorded separately and does not by itself set `是否有划线=是`.
- Recognize explicit Prime Member Price or exclusive Prime pricing as `Prime专享折扣`; ordinary Prime shipping text is not a discount.
- When a Prime member offer and a selectable `Regular Price` coexist, use the regular offer as the ordinary cart price and keep the member amount in the Prime field.
- Normalize coupon text without surrounding page noise.
- Capture quantity promotions such as `Save 5% on 2 select item(s)`, `Save 8% on 5 select item(s)`, or `Buy any 10, Save 10%` in `买赠/多买折扣`.

## Review, seller, and ranking rules

- Generate stable star evidence from DOM-extracted rating, review count, and star percentages.
- Reuse parent review evidence only when later child ASINs have the same review count.
- Treat estimated increases in 1-star through 3-star reviews as new negative-review exceptions.
- Only explicit new-offer text such as `New (2) from` can trigger suspected follow selling; ignore Used, Resale, and Renewed offers.
- Highlight big or small rank changes greater than 10 percent.

## Report readability

- Separate distinct exception explanations with a blank line.
- Deduplicate equivalent rank-change or issue descriptions before writing cells.
- Keep each exception readable and avoid repeating the same fact in both the summary and detail text.

## Stability and quality gates

- Retry transient navigation failures up to 3 times.
- Preserve evidence and report CAPTCHA, Robot Check, Access Denied, invalid links, and capture failures explicitly.
- Do not overwrite the latest baseline when fewer than half the planned items are valid or CAPTCHA is at least 40 percent.
- A project cannot be marked successful when rows, batches, links, or captures fail its quality gate.
- If a project fails, stop before starting the next project.

## Required report

After execution, report the script path, project workbook paths, latest sheet names, exception counts, `_updated` files, postal-code-dependent Buy Box ASINs, and failure stage when applicable. Explicitly confirm breadcrumb, Best Sellers Rank, and embedded-image provenance.
