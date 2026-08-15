"""
Shared R2 paths and loaders for the multi-site monitor hub.

Layout in bucket (default prefix: monitor-sites/):
  monitor-sites/registry.yml
  monitor-sites/{folder}/site.yml          ← site identity (4sale, boshamlan, …)
  monitor-sites/hub/{YYYY-MM-DD}/all-sites.json

Each site's scraper data + validation artifacts stay under its data prefix:
  {r2_prefix}/monitor/websites-config.yml
  {r2_prefix}/monitor/monitor_stats.yml
  {r2_prefix}/monitor/{partition-date}/report.json   ← partition date = listing date + 1 day
"""

from __future__ import annotations

import logging
import os
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple

import boto3
import yaml

log = logging.getLogger("monitor")

MONITOR_SITES_ROOT = os.environ.get("MONITOR_SITES_PREFIX", "monitor-sites").strip("/")


def build_r2_client() -> Tuple:
    """Return a boto3 S3 client pointed at Cloudflare R2."""
    access_key  = os.environ["CF_R2_ACCESS_KEY_ID"]
    secret_key  = os.environ["CF_R2_SECRET_ACCESS_KEY"]
    endpoint    = os.environ["CF_R2_ENDPOINT_URL"].rstrip("/")
    bucket_name = os.environ["CF_R2_BUCKET_NAME"]

    if endpoint.endswith("/" + bucket_name):
        endpoint = endpoint[: -len("/" + bucket_name)]

    client = boto3.client(
        "s3",
        endpoint_url=endpoint,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        region_name="auto",
    )
    return client, bucket_name


def registry_r2_key(root: str = MONITOR_SITES_ROOT) -> str:
    return f"{root}/registry.yml"


def site_config_r2_key(folder: str, root: str = MONITOR_SITES_ROOT) -> str:
    return f"{root}/{folder.strip('/')}/site.yml"


def hub_merged_r2_key(run_date: str, root: str = MONITOR_SITES_ROOT) -> str:
    return f"{root}/hub/{run_date}/all-sites.json"


def hub_tables_base(root: str = MONITOR_SITES_ROOT) -> str:
    """Prefix for flat Parquet exports (MotherDuck ingest)."""
    return f"{root.strip('/')}/hub/tables"


def hub_table_parquet_key(table: str, partition_date: str, root: str = MONITOR_SITES_ROOT) -> str:
    """R2 key for one flat table partition, e.g. hub_daily/2026-06-17.parquet."""
    return f"{hub_tables_base(root)}/{table.strip('/')}/{partition_date}.parquet"


def list_hub_partition_dates(
    client,
    bucket: str,
    root: str = MONITOR_SITES_ROOT,
) -> List[str]:
    """YYYY-MM-DD partition folders under hub/ that contain all-sites.json."""
    prefix = f"{root.strip('/')}/hub/"
    seen: set = set()
    paginator = client.get_paginator("list_objects_v2")
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key = obj["Key"]
            if not key.endswith("/all-sites.json"):
                continue
            rest = key[len(prefix):]
            date_part = rest.split("/", 1)[0]
            if len(date_part) == 10 and date_part[4] == "-" and date_part[7] == "-":
                seen.add(date_part)
    return sorted(seen)


def monitor_data_keys(site: Dict) -> Dict[str, str]:
    """Paths under the site's data prefix (excel/csv storage area)."""
    base = f"{site.get('r2_prefix', '').strip('/')}/monitor"
    return {
        "base":   base,
        "config": f"{base}/websites-config.yml",
        "stats":  f"{base}/monitor_stats.yml",
    }


def report_r2_key(site: Dict, partition_date: str) -> str:
    """R2 key for daily report — *partition_date* matches scraper save folders (listing + 1 day)."""
    return f"{monitor_data_keys(site)['base']}/{partition_date}/report.json"


def partition_date_for_listing(listing_dt: datetime) -> datetime:
    """R2 folder uses save_date = listing date + 1 day."""
    return listing_dt + timedelta(days=1)


def partition_date_str_for_listing(listing_date: str) -> str:
    """YYYY-MM-DD partition path for a listing-date string."""
    listing_dt = datetime.strptime(listing_date, "%Y-%m-%d")
    return partition_date_for_listing(listing_dt).strftime("%Y-%m-%d")


_NON_DAILY_SCHEDULES = frozenset({
    "monthly",
    "quarterly",
    "every_2_days",
    "weekly",
    "biweekly",
})

_SCHEDULE_LOOKBACK_DAYS = {
    "monthly": 31,
    "quarterly": 120,
    # Sheeel often skips a day; keep enough runway past weekends / partition offset.
    "every_2_days": 7,
    "weekly": 7,
    "biweekly": 14,
}

# When registry.yml omits schedule, infer from monitor-sites folder slug.
_FOLDER_DEFAULT_SCHEDULE = {
    "motorgy": "monthly",
    "kcsb": "quarterly",
    "sheeel": "every_2_days",
    "sheel": "every_2_days",
}


def _normalize_schedule(site: Dict) -> str:
    raw = site.get("schedule")
    if raw:
        return str(raw).lower().replace(" ", "_").replace("-", "_")
    folder = (site.get("folder") or site.get("site_id") or "").lower()
    return _FOLDER_DEFAULT_SCHEDULE.get(folder, "daily")


def site_allows_report_fallback(site: Dict) -> bool:
    """Non-daily sites (monthly, quarterly, …) may reuse their latest report in the hub."""
    fb = site.get("report_fallback")
    if fb is True or fb == "latest":
        return True
    if fb is False:
        return False
    return _normalize_schedule(site) in _NON_DAILY_SCHEDULES


def report_lookback_days(site: Dict) -> int:
    explicit = site.get("report_lookback_days")
    if explicit is not None:
        return int(explicit)
    return _SCHEDULE_LOOKBACK_DAYS.get(_normalize_schedule(site), 120)


def list_report_partition_dates(client, bucket: str, site: Dict) -> List[str]:
    """YYYY-MM-DD folders under {r2_prefix}/monitor/ that contain report.json."""
    base = monitor_data_keys(site)["base"]
    prefix = f"{base}/"
    seen: set = set()
    paginator = client.get_paginator("list_objects_v2")
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get("Contents", []):
            key = obj["Key"]
            if not key.endswith("/report.json"):
                continue
            rest = key[len(prefix):]
            date_part = rest.split("/", 1)[0]
            if len(date_part) == 10 and date_part[4] == "-" and date_part[7] == "-":
                seen.add(date_part)
    return sorted(seen)


def resolve_site_folder(explicit: Optional[str] = None) -> str:
    """Folder slug under monitor-sites/ (from env or CLI --site-slug)."""
    if explicit:
        return explicit.strip("/")
    env_slug = os.environ.get("MONITOR_SITE_SLUG", "").strip()
    if env_slug:
        return env_slug
    raise EnvironmentError(
        "MONITOR_SITE_SLUG is required (e.g. 4sale, boshamlan, motorgy, bleems, kcsb, sheeel). "
        "Set it in the repo's monitor.yml workflow env block."
    )


def fetch_yaml_object(client, bucket: str, key: str) -> Dict:
    resp = client.get_object(Bucket=bucket, Key=key)
    return yaml.safe_load(resp["Body"].read().decode("utf-8")) or {}


def put_yaml_object(client, bucket: str, key: str, data: Dict, header: str = "") -> None:
    body = (header + yaml.dump(data, allow_unicode=True, default_flow_style=False, sort_keys=False)).encode("utf-8")
    client.put_object(Bucket=bucket, Key=key, Body=body, ContentType="text/yaml")
    log.info(f"Uploaded → r2://{bucket}/{key}")


def put_bytes(client, bucket: str, key: str, body: bytes, content_type: str) -> None:
    client.put_object(Bucket=bucket, Key=key, Body=body, ContentType=content_type)
    log.info(f"Uploaded → r2://{bucket}/{key}")


def load_site_config_from_r2(
    client,
    bucket: str,
    folder: Optional[str] = None,
    root: str = MONITOR_SITES_ROOT,
) -> Dict:
    slug = resolve_site_folder(folder)
    key  = site_config_r2_key(slug, root)
    try:
        site = fetch_yaml_object(client, bucket, key)
        site["folder"] = slug
        log.info(f"Loaded site config from r2://{bucket}/{key}")
        return site
    except client.exceptions.NoSuchKey:
        raise FileNotFoundError(
            f"Site config not found at r2://{bucket}/{key}. "
            f"Create monitor-sites/{slug}/site.yml in R2 (Cloudflare dashboard or aws s3 cp)."
        ) from None


def load_registry_from_r2(
    client,
    bucket: str,
    root: str = MONITOR_SITES_ROOT,
) -> Dict:
    key = registry_r2_key(root)
    try:
        reg = fetch_yaml_object(client, bucket, key)
        log.info(f"Loaded registry from r2://{bucket}/{key}")
        return reg
    except client.exceptions.NoSuchKey:
        raise FileNotFoundError(
            f"Registry not found at r2://{bucket}/{key}. "
            f"Create monitor-sites/registry.yml in R2."
        ) from None