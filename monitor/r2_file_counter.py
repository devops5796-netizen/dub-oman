import logging
from datetime import datetime
from typing import Dict

log = logging.getLogger("monitor")


def count_r2_objects(client, bucket: str, prefix: str) -> int:
    """
    Count all objects under *prefix* using paginated list_objects_v2.

    Skips zero-byte folder marker keys ending with '/'.
    """
    normalized = prefix.strip("/")
    list_prefix = f"{normalized}/" if normalized else ""

    count = 0
    paginator = client.get_paginator("list_objects_v2")
    try:
        for page in paginator.paginate(Bucket=bucket, Prefix=list_prefix):
            for obj in page.get("Contents", []):
                key = obj["Key"]
                if key.endswith("/"):
                    continue
                count += 1
    except Exception as exc:
        log.warning(f"R2 object count failed for prefix {list_prefix!r}: {exc}")
        return 0

    return count

def count_site_r2_files(client, bucket: str, r2_prefix: str) -> int:
    """Total objects under the site's data prefix (all scrapers + monitor artifacts)."""
    prefix = r2_prefix.strip("/")
    if not prefix:
        return 0
    total = count_r2_objects(client, bucket, prefix)
    log.info(f"Site R2 inventory ({prefix}): {total} object(s)")
    return total


def _list_prefix(prefix: str) -> str:
    normalized = prefix.strip("/")
    return f"{normalized}/" if normalized else ""


def sum_r2_bytes(client, bucket: str, prefix: str) -> int:
    """Sum Size (bytes) of all objects under *prefix*."""
    list_prefix = _list_prefix(prefix)
    total = 0
    paginator = client.get_paginator("list_objects_v2")
    try:
        for page in paginator.paginate(Bucket=bucket, Prefix=list_prefix):
            for obj in page.get("Contents", []):
                if obj["Key"].endswith("/"):
                    continue
                total += obj.get("Size", 0)
    except Exception as exc:
        log.warning(f"R2 byte sum failed for prefix {list_prefix!r}: {exc}")
        return 0
    return total


def sum_site_r2_bytes(client, bucket: str, r2_prefix: str) -> int:
    """Total storage bytes under the site's data prefix."""
    prefix = r2_prefix.strip("/")
    if not prefix:
        return 0
    total = sum_r2_bytes(client, bucket, prefix)
    log.info(f"Site R2 storage ({prefix}): {total} byte(s)")
    return total


def date_partition_prefix(base: str, dt: datetime) -> str:
    base = base.strip("/")
    date_part = f"year={dt.year}/month={dt.month:02d}/day={dt.day:02d}"
    return f"{base}/{date_part}/"


def scraper_date_prefix(base: str, category: str, dt: datetime) -> str:
    """All objects for one scraper on one date (excel, summary, images, etc.)."""
    base = base.strip("/")
    cat = category.strip("/")
    date_part = f"year={dt.year}/month={dt.month:02d}/day={dt.day:02d}"
    return f"{base}/{date_part}/{cat}/"


def sum_scraper_daily_r2_bytes(
    client, bucket: str, base: str, category: str, dt: datetime
) -> int:
    return sum_r2_bytes(client, bucket, scraper_date_prefix(base, category, dt))


def collect_scraper_r2_sizes(
    client,
    bucket: str,
    base: str,
    scraper_categories: Dict[str, str],
) -> Dict[str, int]:
    """
    Single listing pass over base/. Assign each object to the longest
    matching scraper category path (supports nested paths like vehicles/cars-for-sale).
    """
    sizes = {name: 0 for name in scraper_categories}
    if not scraper_categories:
        return sizes

    prefix = _list_prefix(base)
    cat_entries = sorted(
        ((name, cat.strip("/")) for name, cat in scraper_categories.items()),
        key=lambda item: len(item[1]),
        reverse=True,
    )

    paginator = client.get_paginator("list_objects_v2")
    try:
        for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
            for obj in page.get("Contents", []):
                key = obj["Key"]
                if key.endswith("/"):
                    continue
                size = obj.get("Size", 0)
                for name, cat in cat_entries:
                    if f"/{cat}/" in key:
                        sizes[name] += size
                        break
    except Exception as exc:
        log.warning(f"R2 scraper size collection failed for {prefix!r}: {exc}")

    return sizes


def collect_daily_scraper_r2_sizes(
    client,
    bucket: str,
    base: str,
    scraper_categories: Dict[str, str],
    dt: datetime,
) -> Dict[str, int]:
    """Sum bytes per scraper for a single date partition."""
    return {
        name: sum_scraper_daily_r2_bytes(client, bucket, base, cat, dt)
        for name, cat in scraper_categories.items()
    }