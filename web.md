# Daily ads count from the website — implementation prompt

Use the prompt below when adding a **website-sourced daily published-ad count**
to another site's monitor. This is separate from the existing R2-based
`unique_ads` metric.

## Copy/paste prompt

```text
Add a website-sourced daily ad count to this site's existing monitor and
dashboard pipeline.

Goal
----
Count unique ads published on the public website on the listing date being
monitored (normally yesterday). Do not calculate this value from R2 files,
scraper JSON summaries, Excel/CSV exports, or stored R2 reports.

Before coding
-------------
1. Inspect the website and this repository's existing monitoring/report flow.
2. Identify a reliable public page/API and the site's actual publication-date
   field. Verify the response shape against live or captured website responses.
3. Determine how pagination and category/subcategory coverage work so the
   count covers all relevant ads, not just the homepage or first page.
4. If the public site does not provide a reliable source for a complete count,
   report that limitation instead of fabricating a number or substituting an
   R2 count.

Implementation
--------------
1. Count ads whose website publication date matches the monitor's listing
   date. Deduplicate using the site's stable listing/ad ID across categories
   and pages.
2. Fetch data directly from the website. Use reasonable timeouts and request
   volume. Validate every page/API response and surface failures.
3. If any required category or page fails, do not report a partial count as a
   successful total. Log the error and store a NULL count with status "error".
   A successful scan with no matching ads is count 0 with status "ok".
4. Add these site-level fields to the daily report consumed by the monitor hub:

   {
     "website_published_ads": 123,
     "website_published_ads_date": "YYYY-MM-DD",
     "website_published_ads_source": "short_source_identifier",
     "website_published_ads_status": "ok"
   }

   Include scan diagnostics (for example category/page counts) when useful.
   On failure, use null for `website_published_ads`, set status to "error", and
   include a short `website_published_ads_error` for diagnosis.
5. Carry the fields through the hub aggregation and site_daily export/table
   schema, then add a dashboard tile/column labelled clearly as website data.
   Preserve the existing R2-based `unique_ads` field and its behavior.
6. Add tests for date filtering, ID deduplication, pagination/category coverage,
   empty-but-successful results, and source failure. Do not require live website
   access for unit tests.
7. Update directly related documentation with the metric definition, source,
   date/timezone convention, and known limitations.

Important semantics
-------------------
- The daily value means ads whose website `date_published` (or equivalent)
  falls on the selected listing date; it is not the current active-listing
  inventory and is not the R2 upload/file count.
- Keep the count site-level; do not sum across sites unless the product
  explicitly requests a hub-wide total and defines duplicate handling.
- For historical backfills, only emit a value when the website can still
  reliably provide that historical date. Never silently use today's page
  inventory or R2 data as a fallback.
- Use the existing report date convention and make the date/timezone explicit.

Finish by running the smallest relevant test/type/lint checks and summarize
the implemented website source, metric semantics, limitations, and validation.
```
