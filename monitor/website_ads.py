"""Count Dubizzle Oman ads published on a selected UTC listing date."""

from __future__ import annotations

import http.client
import json
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, timezone
from typing import Any, Callable, Dict, List, Optional, Sequence, Tuple

SEARCH_URL = "https://search.mena.sector.run/_msearch"
SEARCH_INDEX = "olx-om-production-ads-ar"
LOCATION_ID = "0-1"
PAGE_SIZE = 100
REQUEST_TIMEOUT_SECONDS = 30
MAX_PAGES_PER_CATEGORY = 1000
WEBSITE_PUBLISHED_ADS_SOURCE = "dubizzle_search_api"

# Top-level ad categories used by the site's production scraper. Each query
# includes all subcategories under that parent.
WEBSITE_AD_CATEGORIES = (
    "vehicles",
    "properties",
    "mobile-phones-accessories",
    "electronics-home-appliances",
    "home-garden",
    "fashion-beauty",
    "pets",
    "kids-babies",
    "sporting-goods-bikes",
    "hobbies-music-art-books",
    "jobs-services",
    "business-industrial",
    "services",
)


class WebsiteAdsScanError(RuntimeError):
    def __init__(
        self,
        message: str,
        *,
        categories_scanned: int = 0,
        pages_scanned: int = 0,
        categories_completed: Optional[Sequence[str]] = None,
    ) -> None:
        super().__init__(message)
        self.categories_scanned = categories_scanned
        self.pages_scanned = pages_scanned
        self.categories_completed = list(categories_completed or ())

    def diagnostics(self) -> Dict[str, Any]:
        return {
            "categories_scanned": self.categories_scanned,
            "pages_scanned": self.pages_scanned,
            "categories_completed": self.categories_completed,
        }


@dataclass(frozen=True)
class WebsiteAdsScanResult:
    count: int
    categories_scanned: int
    pages_scanned: int
    categories_completed: List[str]

    def diagnostics(self) -> Dict[str, Any]:
        return {
            "categories_scanned": self.categories_scanned,
            "pages_scanned": self.pages_scanned,
            "categories_completed": self.categories_completed,
        }


FetchPage = Callable[
    [str, float, float, Optional[List[Any]], str],
    Dict[str, Any],
]


def _request_search_page(
    category: str,
    start_timestamp: float,
    end_timestamp: float,
    search_after: Optional[List[Any]],
    authorization: str,
) -> Dict[str, Any]:
    query: Dict[str, Any] = {
        "size": PAGE_SIZE,
        "query": {
            "bool": {
                "must": [
                    {"term": {"category.slug": category}},
                    {"term": {"location.externalID": LOCATION_ID}},
                    {
                        "range": {
                            "createdAt": {
                                "gte": start_timestamp,
                                "lt": end_timestamp,
                            }
                        }
                    },
                ]
            }
        },
        "sort": [
            {"createdAt": {"order": "desc"}},
            {"id": {"order": "desc"}},
        ],
        "_source": ["id", "createdAt"],
    }
    if search_after is not None:
        query["search_after"] = search_after

    body = (
        json.dumps({"index": SEARCH_INDEX})
        + "\n"
        + json.dumps(query)
        + "\n"
    ).encode("utf-8")
    request = urllib.request.Request(
        SEARCH_URL
        + "?"
        + urllib.parse.urlencode(
            {
                "filter_path": (
                    "responses.hits.hits._source,"
                    "responses.hits.hits.sort,responses.error"
                )
            }
        ),
        data=body,
        headers={
            "Accept": "application/json",
            "Authorization": authorization,
            "Content-Type": "application/x-ndjson",
            "Origin": "https://www.dubizzle.com.om",
            "Referer": "https://www.dubizzle.com.om/",
            "User-Agent": "dub-oman-monitor/1.0",
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
            return json.loads(response.read().decode("utf-8"))
    except (
        http.client.HTTPException,
        urllib.error.URLError,
        TimeoutError,
        UnicodeDecodeError,
        json.JSONDecodeError,
    ) as exc:
        raise WebsiteAdsScanError(f"search request failed: {exc}") from exc


def _utc_date(value: Any) -> date:
    if isinstance(value, bool):
        raise ValueError("createdAt is not a timestamp")
    try:
        timestamp = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError("createdAt is missing or invalid") from exc
    return datetime.fromtimestamp(timestamp, timezone.utc).date()


def _parse_hits(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    if not isinstance(payload, dict):
        raise ValueError("invalid search response envelope")
    responses = payload.get("responses")
    if not isinstance(responses, list) or len(responses) != 1:
        raise ValueError("unexpected search response envelope")
    response = responses[0]
    if not isinstance(response, dict):
        raise ValueError("invalid search response")
    if response.get("error") is not None:
        raise ValueError("search API returned an error")

    hits_object = response.get("hits")
    hits = hits_object.get("hits") if isinstance(hits_object, dict) else None
    if not isinstance(hits, list):
        raise ValueError("search response is missing hits")
    return hits


def scan_website_published_ads(
    listing_date: str,
    authorization: Optional[str],
    *,
    categories: Sequence[str] = WEBSITE_AD_CATEGORIES,
    fetch_page: Optional[FetchPage] = None,
) -> WebsiteAdsScanResult:
    """Count unique ads whose website createdAt falls on listing_date in UTC."""
    try:
        day = datetime.strptime(listing_date, "%Y-%m-%d").date()
    except ValueError as exc:
        raise WebsiteAdsScanError("listing date must use YYYY-MM-DD") from exc
    if day.strftime("%Y-%m-%d") != listing_date:
        raise WebsiteAdsScanError("listing date must use YYYY-MM-DD")
    if not authorization or not authorization.strip():
        raise WebsiteAdsScanError("AUTHORIZATION is not configured")
    authorization = authorization.strip()
    if not categories:
        raise WebsiteAdsScanError("no website ad categories are configured")

    start = datetime.combine(day, time.min, tzinfo=timezone.utc)
    end = start + timedelta(days=1)
    start_timestamp = start.timestamp()
    end_timestamp = end.timestamp()
    page_fetcher = fetch_page or _request_search_page
    unique_ids: set[str] = set()
    pages_scanned = 0
    completed: List[str] = []

    for category in categories:
        search_after: Optional[List[Any]] = None
        seen_cursors: set[Tuple[str, ...]] = set()
        category_pages = 0

        while True:
            if category_pages >= MAX_PAGES_PER_CATEGORY:
                raise WebsiteAdsScanError(
                    f"{category} exceeded the page safety limit",
                    categories_scanned=len(completed),
                    pages_scanned=pages_scanned,
                    categories_completed=completed,
                )

            category_pages += 1
            pages_scanned += 1
            try:
                payload = page_fetcher(
                    category,
                    start_timestamp,
                    end_timestamp,
                    search_after,
                    authorization,
                )
                hits = _parse_hits(payload)
            except (WebsiteAdsScanError, ValueError) as exc:
                raise WebsiteAdsScanError(
                    f"{category} page {category_pages}: {exc}",
                    categories_scanned=len(completed),
                    pages_scanned=pages_scanned,
                    categories_completed=completed,
                ) from exc

            for hit in hits:
                if not isinstance(hit, dict) or not isinstance(hit.get("_source"), dict):
                    raise WebsiteAdsScanError(
                        f"{category} page {category_pages}: invalid ad record",
                        categories_scanned=len(completed),
                        pages_scanned=pages_scanned,
                        categories_completed=completed,
                    )
                source = hit["_source"]
                try:
                    published_date = _utc_date(source.get("createdAt"))
                except (OverflowError, OSError, ValueError) as exc:
                    raise WebsiteAdsScanError(
                        f"{category} page {category_pages}: invalid createdAt",
                        categories_scanned=len(completed),
                        pages_scanned=pages_scanned,
                        categories_completed=completed,
                    ) from exc

                if published_date != day:
                    continue
                ad_id = source.get("id")
                if ad_id is None or not str(ad_id).strip():
                    raise WebsiteAdsScanError(
                        f"{category} page {category_pages}: ad is missing its stable id",
                        categories_scanned=len(completed),
                        pages_scanned=pages_scanned,
                        categories_completed=completed,
                    )
                unique_ids.add(str(ad_id).strip())

            if len(hits) < PAGE_SIZE:
                break

            cursor = hits[-1].get("sort")
            if not isinstance(cursor, list) or not cursor:
                raise WebsiteAdsScanError(
                    f"{category} page {category_pages}: missing pagination cursor",
                    categories_scanned=len(completed),
                    pages_scanned=pages_scanned,
                    categories_completed=completed,
                )
            cursor_key = tuple(str(value) for value in cursor)
            if cursor_key in seen_cursors:
                raise WebsiteAdsScanError(
                    f"{category} page {category_pages}: repeated pagination cursor",
                    categories_scanned=len(completed),
                    pages_scanned=pages_scanned,
                    categories_completed=completed,
                )
            seen_cursors.add(cursor_key)
            search_after = cursor

        completed.append(category)

    return WebsiteAdsScanResult(
        count=len(unique_ids),
        categories_scanned=len(completed),
        pages_scanned=pages_scanned,
        categories_completed=completed,
    )
