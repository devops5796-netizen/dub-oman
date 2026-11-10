# Monitor metrics

## Website-published ads

`website_published_ads` counts unique Dubizzle Oman ads whose website
`createdAt` timestamp falls on the report's `inspect_date`. The website search
API is queried directly; timestamps are interpreted in UTC, matching the
monitor's existing report-date convention. IDs are deduplicated across pages
and top-level ad categories.

The source is `dubizzle_search_api`. The scan covers the 13 public ad category
parents used by the site's ad scraper, including their subcategories. The
separate new-car catalog is excluded because its feed does not provide a
verified ad publication date. This is not active-listing inventory and does
not use the R2 `unique_ads` count, which remains unchanged.

The website search request does not use an `Authorization` header. If the
source is unavailable, the report stores a null count, status `error`, a short
error, and scan diagnostics; it never reports a partial total or falls back to
R2. Historical dates are counted only if the website search API still returns
their records.

The site monitor stores these fields at the top level of its R2 `report.json`:
`website_published_ads`, `website_published_ads_date`,
`website_published_ads_source`, and `website_published_ads_status`. The hub
aggregation, `site_daily` schema, and dashboard are maintained outside this
repository and must map these report fields to expose the metric there.
