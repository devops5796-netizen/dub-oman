import unittest
from datetime import datetime, timezone
from unittest.mock import patch

from website_ads import (
    WEBSITE_AD_CATEGORIES,
    WebsiteAdsScanError,
    scan_website_published_ads,
)


def _epoch(date_text):
    return datetime.strptime(date_text, "%Y-%m-%d").replace(tzinfo=timezone.utc).timestamp()


def _hit(ad_id, created_at, sort_id=None):
    return {
        "_source": {"id": ad_id, "createdAt": created_at},
        "sort": [created_at, sort_id if sort_id is not None else ad_id],
    }


def _response(hits):
    return {"responses": [{"hits": {"hits": hits}}]}


class WebsiteAdsScanTests(unittest.TestCase):
    def test_date_filter_deduplication_pagination_and_category_coverage(self):
        target = _epoch("2026-11-08")
        previous_day = _epoch("2026-11-07")
        calls = []

        def fetch_page(category, start, end, search_after, authorization):
            calls.append((category, start, end, search_after, authorization))
            page_number = sum(1 for call in calls if call[0] == category) - 1
            pages = {
                "vehicles": [
                    [_hit("one", target), _hit("wrong-day", previous_day)],
                    [_hit("one", target), _hit("two", target)],
                    [],
                ],
                "properties": [
                    [_hit("two", target), _hit("three", target)],
                    [],
                ],
            }
            return _response(pages[category][page_number])

        with patch("website_ads.PAGE_SIZE", 2):
            result = scan_website_published_ads(
                "2026-11-08",
                "Basic test-token",
                categories=("vehicles", "properties"),
                fetch_page=fetch_page,
            )

        self.assertEqual(result.count, 3)
        self.assertEqual(result.categories_scanned, 2)
        self.assertEqual(result.pages_scanned, 5)
        self.assertEqual(result.categories_completed, ["vehicles", "properties"])
        self.assertEqual([call[0] for call in calls], [
            "vehicles", "vehicles", "vehicles", "properties", "properties",
        ])
        self.assertEqual(calls[0][1], _epoch("2026-11-08"))
        self.assertEqual(calls[0][2], _epoch("2026-11-09"))
        self.assertEqual(calls[0][4], "Basic test-token")

    def test_default_category_coverage_matches_site_ad_categories(self):
        calls = []

        def fetch_page(category, _start, _end, _cursor, _authorization):
            calls.append(category)
            return _response([])

        result = scan_website_published_ads(
            "2026-11-08",
            "Basic test-token",
            fetch_page=fetch_page,
        )

        self.assertEqual(tuple(calls), WEBSITE_AD_CATEGORIES)
        self.assertEqual(result.categories_scanned, len(WEBSITE_AD_CATEGORIES))
        self.assertEqual(result.count, 0)

    def test_empty_success_is_zero(self):
        result = scan_website_published_ads(
            "2026-11-08",
            "Basic test-token",
            categories=("vehicles", "properties"),
            fetch_page=lambda *_args: _response([]),
        )

        self.assertEqual(result.count, 0)
        self.assertEqual(result.pages_scanned, 2)

    def test_source_failure_raises_with_partial_scan_diagnostics(self):
        def fetch_page(category, *_args):
            if category == "properties":
                raise WebsiteAdsScanError("HTTP 503")
            return _response([_hit("one", _epoch("2026-11-08"))])

        with self.assertRaises(WebsiteAdsScanError) as raised:
            scan_website_published_ads(
                "2026-11-08",
                "Basic test-token",
                categories=("vehicles", "properties"),
                fetch_page=fetch_page,
            )

        self.assertIn("properties page 1", str(raised.exception))
        self.assertEqual(
            raised.exception.diagnostics(),
            {
                "categories_scanned": 1,
                "pages_scanned": 2,
                "categories_completed": ["vehicles"],
            },
        )

    def test_missing_authorization_fails_without_request(self):
        with self.assertRaisesRegex(WebsiteAdsScanError, "AUTHORIZATION"):
            scan_website_published_ads("2026-11-08", None, fetch_page=lambda *_args: {})


if __name__ == "__main__":
    unittest.main()
