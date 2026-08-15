"""
request_metrics.py
==================
Collect HTTP request throughput and error-rate metrics per scraper for the hub dashboard.

Reads daily JSON summaries under summary/ (same location as ads_counter).
Scrapers should emit request_metrics in their summary JSON after each run.
"""

from __future__ import annotations

import json
import logging
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

log = logging.getLogger("monitor")

REQUEST_TOTAL_KEYS = (
    "requests_total",
    "total_http_requests",
    "http_requests",
    "scrape_do_requests",
    "request_count",
)
REQUEST_FAILED_KEYS = (
    "requests_failed",
    "failed_requests",
    "requests_errors",
    "errors_count",
    "http_errors",
)
DURATION_KEYS = (
    "duration_sec",
    "elapsed_seconds",
    "scrape_duration_sec",
    "runtime_sec",
)
RPM_KEYS = (
    "requests_per_min",
    "req_per_min",
    "avg_requests_per_min",
    "rpm",
)
CACHE_KEYS = ("cache_hits", "cache_hits_free")


def _first_int(sources: List[Dict], keys: Tuple[str, ...]) -> Optional[int]:
    for src in sources:
        for key in keys:
            val = src.get(key)
            if isinstance(val, (int, float)) and val >= 0:
                return int(val)
    return None


def _first_float(sources: List[Dict], keys: Tuple[str, ...]) -> Optional[float]:
    for src in sources:
        for key in keys:
            val = src.get(key)
            if isinstance(val, (int, float)) and val >= 0:
                return float(val)
    return None


def _metric_sources(data: Dict) -> List[Dict]:
    """Collect dict layers that may hold request metrics."""
    sources: List[Dict] = [data]
    nested = data.get("request_metrics")
    if isinstance(nested, dict):
        sources.insert(0, nested)
    stats = data.get("stats")
    if isinstance(stats, dict):
        sources.append(stats)
    return sources


def _collect_failed_items(data: Dict) -> List[Dict]:
    """Normalize failed scrape targets from common JSON shapes."""
    items: List[Dict] = []
    seen: set = set()

    def add_item(name: str, errors: Any, detail: Optional[str] = None) -> None:
        if not name:
            return
        try:
            err_count = int(errors) if errors is not None else 0
        except (TypeError, ValueError):
            err_count = 1 if errors else 0
        if err_count <= 0 and not detail:
            return
        key = (name, err_count, detail or "")
        if key in seen:
            return
        seen.add(key)
        entry: Dict[str, Any] = {"name": name, "errors": err_count}
        if detail:
            entry["detail"] = detail
        items.append(entry)

    for src in _metric_sources(data):
        for raw in src.get("failed_items") or src.get("failed_scrapers") or []:
            if not isinstance(raw, dict):
                continue
            name = (
                raw.get("name")
                or raw.get("scraper")
                or raw.get("slug")
                or raw.get("category")
                or "unknown"
            )
            detail = raw.get("detail") or raw.get("reason") or raw.get("message")
            add_item(str(name), raw.get("errors") or raw.get("error_count") or 1, detail)

        for raw in src.get("errors") or []:
            if isinstance(raw, str):
                add_item(raw, 1)
            elif isinstance(raw, dict):
                name = raw.get("name") or raw.get("scraper") or raw.get("slug") or "error"
                detail = raw.get("detail") or raw.get("message") or raw.get("reason")
                add_item(str(name), raw.get("count") or 1, detail)

    for raw in data.get("subcategories") or data.get("categories") or []:
        if not isinstance(raw, dict):
            continue
        err = raw.get("errors") or raw.get("error_count")
        if err:
            name = (
                raw.get("name_en")
                or raw.get("name_ar")
                or raw.get("name")
                or raw.get("slug")
                or "subcategory"
            )
            add_item(str(name), err)

    return items


def format_failed_items_summary(failed_items: List[Dict], max_len: int = 400) -> Optional[str]:
    if not failed_items:
        return None
    parts: List[str] = []
    for item in failed_items[:12]:
        name = item.get("name", "?")
        count = item.get("errors", 0)
        detail = item.get("detail")
        bit = f"{name}: {count} error(s)"
        if detail:
            bit += f" ({detail})"
        parts.append(bit)
    text = "; ".join(parts)
    if len(failed_items) > 12:
        text += f"; +{len(failed_items) - 12} more"
    return text[:max_len] if len(text) > max_len else text


def extract_request_metrics(data: Any) -> Optional[Dict[str, Any]]:
    """
    Parse a JSON summary into normalized request metrics.

    Returns None when no request counters are present.
    """
    if not isinstance(data, dict):
        return None

    sources = _metric_sources(data)
    requests_total = _first_int(sources, REQUEST_TOTAL_KEYS)
    requests_failed = _first_int(sources, REQUEST_FAILED_KEYS)
    duration_sec = _first_float(sources, DURATION_KEYS)
    requests_per_min = _first_float(sources, RPM_KEYS)
    cache_hits = _first_int(sources, CACHE_KEYS)
    failed_items = _collect_failed_items(data)

    if requests_total is None and requests_failed is None and not failed_items:
        return None

    if requests_failed is None:
        requests_failed = sum(int(i.get("errors") or 0) for i in failed_items) or 0

    if requests_total is None and requests_failed:
        requests_total = requests_failed

    error_rate_pct: Optional[float] = None
    if requests_total and requests_total > 0:
        error_rate_pct = round((requests_failed or 0) / requests_total * 100.0, 2)

    if requests_per_min is None and requests_total and duration_sec and duration_sec > 0:
        requests_per_min = round(requests_total / (duration_sec / 60.0), 2)

    return {
        "requests_total": requests_total,
        "requests_failed": requests_failed or 0,
        "error_rate_pct": error_rate_pct,
        "requests_per_min": requests_per_min,
        "duration_sec": duration_sec,
        "cache_hits": cache_hits,
        "failed_items": failed_items,
        "failed_items_summary": format_failed_items_summary(failed_items),
        "metrics_source": "json_summary",
    }


def _metrics_completeness(metrics: Dict[str, Any]) -> int:
    score = 0
    if metrics.get("requests_total"):
        score += 4
    if metrics.get("requests_per_min") is not None:
        score += 2
    if metrics.get("error_rate_pct") is not None:
        score += 1
    if metrics.get("failed_items"):
        score += 1
    return score


def _request_metrics_scan_targets(
    base: str,
    category: Optional[str],
    dt: datetime,
) -> List[Tuple[str, bool]]:
    """
    (prefix, summary_filename_only) pairs for request-metrics discovery.

    Standard scrapers write to summary/. Legacy uploads used the
    partition root (e.g. dalil_summary_YYYYMMDD.json).
    """
    base = base.strip("/")
    date_part = f"year={dt.year}/month={dt.month:02d}/day={dt.day:02d}"

    if category:
        partition = f"{base}/{date_part}/{category.strip('/')}/"
    else:
        partition = f"{base}/{date_part}/"

    targets: List[Tuple[str, bool]] = [
        (f"{partition}summary/", False),
        (partition, True),  # legacy: summary file directly under partition
    ]
    return targets


def load_json_request_metrics(
    client,
    bucket: str,
    r2_base: str,
    category: Optional[str],
    partition_dt: datetime,
) -> Tuple[Optional[Dict[str, Any]], Optional[str]]:
    """
    Scan summary/ (and legacy partition-root summaries) for request metrics.

    Returns (metrics_dict, r2_key).
    """
    best: Optional[Dict[str, Any]] = None
    best_key: Optional[str] = None
    best_score = -1

    for prefix, summary_only in _request_metrics_scan_targets(r2_base, category, partition_dt):
        try:
            paginator = client.get_paginator("list_objects_v2")
            for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
                for obj in page.get("Contents", []):
                    key = obj["Key"]
                    if not key.lower().endswith(".json"):
                        continue
                    if summary_only:
                        basename = key.rsplit("/", 1)[-1].lower()
                        if "summary" not in basename:
                            continue
                    try:
                        resp = client.get_object(Bucket=bucket, Key=key)
                        data = json.loads(resp["Body"].read().decode("utf-8"))
                    except Exception as exc:
                        log.debug(f"Could not read JSON metrics {key}: {exc}")
                        continue

                    metrics = extract_request_metrics(data)
                    if metrics is None:
                        continue
                    score = _metrics_completeness(metrics)
                    if score > best_score:
                        best = dict(metrics)
                        best_key = key
                        best_score = score
        except Exception as exc:
            log.debug(f"JSON listing under {prefix}: {exc}")

    if best is not None:
        best["json_summary_key"] = best_key
    return best, best_key


def empty_scraper_metrics() -> Dict[str, Any]:
    return {
        "requests_total": None,
        "requests_failed": None,
        "error_rate_pct": None,
        "requests_per_min": None,
        "duration_sec": None,
        "cache_hits": None,
        "failed_items": [],
        "failed_items_summary": None,
        "metrics_source": "none",
        "json_summary_key": None,
    }


def count_scraper_request_metrics(
    client,
    bucket: str,
    r2_base: str,
    category: Optional[str],
    partition_dt: datetime,
) -> Dict[str, Any]:
    """Load request metrics for one scraper partition."""
    metrics, _key = load_json_request_metrics(
        client, bucket, r2_base, category, partition_dt
    )
    if metrics is None:
        return empty_scraper_metrics()
    return metrics


def aggregate_site_request_metrics(scraper_results: List[Dict]) -> Dict[str, Any]:
    """Roll up per-scraper HTTP metrics to site-level totals."""
    totals = {
        "requests_total": 0,
        "requests_failed": 0,
        "duration_sec": 0.0,
        "scrapers_with_metrics": 0,
        "has_metrics": False,
    }
    rpm_values: List[float] = []

    for row in scraper_results:
        total = row.get("requests_total")
        if total is None:
            continue
        totals["has_metrics"] = True
        totals["scrapers_with_metrics"] += 1
        totals["requests_total"] += int(total)
        totals["requests_failed"] += int(row.get("requests_failed") or 0)
        duration = row.get("duration_sec")
        if duration:
            totals["duration_sec"] += float(duration)
        rpm = row.get("requests_per_min")
        if rpm is not None:
            rpm_values.append(float(rpm))

    if not totals["has_metrics"]:
        return {
            "requests_total": None,
            "requests_failed": None,
            "error_rate_pct": None,
            "requests_per_min": None,
            "scrapers_with_metrics": 0,
        }

    error_rate_pct: Optional[float] = None
    if totals["requests_total"] > 0:
        error_rate_pct = round(totals["requests_failed"] / totals["requests_total"] * 100.0, 2)

    requests_per_min: Optional[float] = None
    if rpm_values:
        requests_per_min = round(sum(rpm_values) / len(rpm_values), 2)
    elif totals["duration_sec"] > 0:
        requests_per_min = round(
            totals["requests_total"] / (totals["duration_sec"] / 60.0), 2
        )

    return {
        "requests_total": totals["requests_total"],
        "requests_failed": totals["requests_failed"],
        "error_rate_pct": error_rate_pct,
        "requests_per_min": requests_per_min,
        "scrapers_with_metrics": totals["scrapers_with_metrics"],
    }


def build_run_error_summary(
    scraper_results: List[Dict],
    alerts: List[Dict],
) -> Dict[str, Any]:
    """
    Per-run error summary: validation failures, HTTP errors, who failed.
    """
    failed_scrapers: List[Dict] = []
    for row in scraper_results:
        if row.get("all_passed"):
            continue
        reasons: List[str] = []
        if row.get("files_found", 0) == 0 and not row.get("files_optional"):
            reasons.append("no Excel files")
        failed_checks = []
        for fr in row.get("file_results") or []:
            for check in fr.get("validation", {}).get("checks", []):
                if not check.get("passed"):
                    failed_checks.append(check.get("name") or "check")
        if failed_checks:
            reasons.append("validation: " + ", ".join(sorted(set(failed_checks))[:5]))
        http_failed = row.get("requests_failed")
        if http_failed:
            reasons.append(f"{http_failed} HTTP error(s)")
        summary = row.get("failed_items_summary")
        if summary:
            reasons.append(summary)

        failed_scrapers.append({
            "scraper": row.get("scraper"),
            "reason": "; ".join(reasons) if reasons else "validation failed",
            "requests_failed": http_failed,
            "error_rate_pct": row.get("error_rate_pct"),
        })

    site_http = aggregate_site_request_metrics(scraper_results)
    scrapers_total = len(scraper_results)
    scrapers_failed = len(failed_scrapers)

    return {
        "scrapers_total": scrapers_total,
        "scrapers_failed": scrapers_failed,
        "scrapers_passed": scrapers_total - scrapers_failed,
        "validation_fail_rate_pct": (
            round(scrapers_failed / scrapers_total * 100.0, 1) if scrapers_total else None
        ),
        "failed_scrapers": failed_scrapers,
        "alert_count": len(alerts),
        "http": {
            "requests_total": site_http.get("requests_total"),
            "requests_failed": site_http.get("requests_failed"),
            "error_rate_pct": site_http.get("error_rate_pct"),
            "requests_per_min": site_http.get("requests_per_min"),
            "scrapers_reporting": site_http.get("scrapers_with_metrics", 0),
        },
    }