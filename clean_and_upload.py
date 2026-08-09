import argparse
import ast
import io
import json
import os
import re
from datetime import datetime, timedelta, timezone
import random
import time
import pandas as pd
import requests as req
from PIL import Image
import glob
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth
from text_utils import clean_text, sanitize_filename
from contact_info_fetcher import build_ad_url, fetch_contact_info, EMPTY_CONTACT_INFO
from r2_uploader import upload_buffer

THUMB_URL_TEMPLATE = "https://images.dubizzle.com.om/thumbnails/{photo_id}-800x600.webp"

COLUMNS_TO_DROP = ['geo_point', 'price', 'title_l1', 'description_l1', 'slug_l1', 'coverPhoto',
                   'external_link', 'external_link_l1', 'documentsTags', 'videoCount', 'documentCount'
                   'panoramaCount', 'location.lvl0','location.lvl1','location.lvl2', 'category.lvl0',
                   'category.lvl1','category.lvl2']

VEHICLE_MANUFACTURER_SPLIT_SLUGS = {"cars-for-sale", "cars-for-rent"}

# Some top-level scraped categories should be grouped under a shared parent
# folder in R2, with their own slug as the subfolder:
#   properties-for-rent -> properties/properties-for-rent
#   properties-for-sale -> properties/properties-for-sale
CATEGORY_GROUP_OVERRIDES = {
    "properties-for-rent": "properties",
    "properties-for-sale": "properties",
}


def resolve_category_r2_path(cat0_slug: str, subcat_slug: str | None = None) -> str:
    """
    Slug-based R2 folder path for a top-level category, optionally nested
    under a subcategory slug.

    - properties-for-rent / properties-for-sale -> "properties/<own-slug>"
    - vehicles + a subcat slug                  -> "vehicles/<subcat-slug>"
    - everything else                           -> the category's own slug
    """
    cat0_slug = cat0_slug or "uncategorized"
    if cat0_slug in CATEGORY_GROUP_OVERRIDES:
        return f"{CATEGORY_GROUP_OVERRIDES[cat0_slug]}/{cat0_slug}"
    if cat0_slug == "vehicles" and subcat_slug:
        return f"vehicles/{subcat_slug}"
    return cat0_slug

TIMESTAMP_FIELDS = ("createdAt", "updatedAt", "timestamp")

# =============================================================================
# Helper functions (unchanged)
# =============================================================================

def parse_formatted_extra_fields(record) -> dict:
    field = record.get("formattedExtraFields")
    if isinstance(field, str):
        try:
            field = ast.literal_eval(field)
        except (ValueError, SyntaxError):
            field = []
    if not isinstance(field, list):
        return {}
    result = {}
    for item in field:
        if isinstance(item, dict):
            attr = item.get("attribute")
            val = item.get("formattedValue_l1") or item.get("formattedValue")
            if attr and val is not None:
                result[attr] = val
    return result


def parse_category(cat_field):
    if isinstance(cat_field, list):
        cats = cat_field
    elif isinstance(cat_field, str):
        try:
            cats = ast.literal_eval(cat_field)
        except (ValueError, SyntaxError):
            cats = []
    else:
        cats = []
    by_level = {c.get("level"): c for c in cats if isinstance(c, dict)}
    return by_level.get(0), by_level.get(1), by_level.get(2)


def sheet_name_for(cat1: dict | None, cat2: dict | None) -> str:
    if cat1 is None:
        name = "Uncategorized"
    else:
        name = cat1.get("name_l1") or cat1.get("name") or "Uncategorized"
        if cat2:
            sub = cat2.get("name_l1") or cat2.get("name")
            if sub:
                name = f"{name} ({sub})"
    name = clean_text(name)
    name = re.sub(r"[:\\/?*\[\]]", "-", name)
    return name[:31] or "Uncategorized"


def photo_urls(photos_field) -> list:
    if isinstance(photos_field, str):
        try:
            photos_field = ast.literal_eval(photos_field)
        except (ValueError, SyntaxError):
            photos_field = []
    if not photos_field or not isinstance(photos_field, list):
        return []
    urls = []
    for p in photos_field:
        pid = p.get("id") if isinstance(p, dict) else None
        if pid:
            urls.append(THUMB_URL_TEMPLATE.format(photo_id=pid))
    return urls


def format_timestamp(value):
    if value is None or value == "":
        return value
    try:
        ts = float(value)
    except (TypeError, ValueError):
        return value
    try:
        dt = datetime.fromtimestamp(ts, tz=timezone.utc)
    except (OverflowError, OSError, ValueError):
        return value
    return dt.strftime("%Y-%m-%dT%H:%M:%SZ")


def clean_timestamp_fields(record: dict) -> dict:
    for field in TIMESTAMP_FIELDS:
        if field in record:
            record[field] = format_timestamp(record[field])
    return record


def download_images(images: list, id_prod: str, category_display: str, dt: datetime = None) -> list:
    r2_paths = []
    uploaded = 0
    failed = 0
    if not images:
        return r2_paths
    file_prefix = id_prod or "unknown"
    for idx, img_url in enumerate(images, start=1):
        filename = f"{file_prefix}-{idx}.webp"
        try:
            r = req.get(img_url, timeout=15)
            if r.status_code == 200:
                img = Image.open(io.BytesIO(r.content)).convert("RGB")
                buf = io.BytesIO()
                img.save(buf, format="WEBP", quality=75, method=6)
                buf.seek(0)
                r2_key = upload_buffer(
                    buf,
                    filename=filename,
                    category_display=category_display,
                    file_type="images",
                    content_type="image/webp",
                    dt=dt,
                )
                if r2_key:
                    r2_paths.append(r2_key)
                    uploaded += 1
                else:
                    failed += 1
            else:
                failed += 1
        except Exception as e:
            print(f"    [ERROR] {filename} image {idx}: {e}")
            failed += 1
    if uploaded or failed:
        print(f"    {file_prefix}: {uploaded} uploaded, {failed} failed out of {len(images)}")
    return r2_paths


def load_raw(csv_path: str) -> pd.DataFrame | None:
    if not os.path.exists(csv_path) or os.path.getsize(csv_path) == 0:
        return None
    return pd.read_csv(csv_path)


def clean_and_group(df: pd.DataFrame, page=None, dt: datetime = None):
    sheets: dict[str, list] = {}
    all_records = []
    cat0_name_l1 = None
    cat0_name_ar = None
    cat0_slug = None
    for _, row in df.iterrows():
        cat0, cat1, cat2 = parse_category(row.get("category"))
        if cat0 is None:
            continue
        if cat0_name_l1 is None:
            cat0_name_l1 = cat0.get("name_l1")
            cat0_name_ar = cat0.get("name")
            cat0_slug = cat0.get("slug")
        sheet = sheet_name_for(cat1, cat2)
        urls = photo_urls(row.get("photos"))
        ad_id = str(row.get("id") or row.get("externalID") or "")

        if cat0_slug == "vehicles":
            subcat_slug = (cat1.get("slug") if cat1 else None) or "uncategorized"
            r2_category_path = resolve_category_r2_path(cat0_slug, subcat_slug)
        else:
            r2_category_path = resolve_category_r2_path(cat0_slug)

        image_r2_paths = download_images(urls, id_prod=ad_id, category_display=r2_category_path, dt=dt)
        record = row.to_dict()
        record = clean_timestamp_fields(record)
        record["image_r2_paths"] = image_r2_paths
        record["photo_urls"] = urls
        record.pop("photos", None)
        record["image_r2_paths"] = image_r2_paths
        if page is not None:
            ad_url = build_ad_url(record)
            if ad_url:
                record["contact_info"] = fetch_contact_info(page, ad_url)
                time.sleep(random.uniform(2, 5))
            else:
                record["contact_info"] = dict(EMPTY_CONTACT_INFO)
        else:
            record["contact_info"] = dict(EMPTY_CONTACT_INFO)
        sheets.setdefault(sheet, []).append(record)
        all_records.append(record)
    return cat0_name_l1, cat0_name_ar, cat0_slug, sheets, all_records


def _stringify_complex_columns(sheet_df: pd.DataFrame) -> pd.DataFrame:
    for col in sheet_df.columns:
        sheet_df[col] = sheet_df[col].apply(
            lambda v: json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else v
        )
    return sheet_df


def safe_sheet_name(name: str, used: set) -> str:
    name = clean_text(name)
    name = re.sub(r"[:\\/?*\[\]]", "-", name)[:31] or "Sheet"
    candidate = name
    n = 1
    while candidate in used:
        suffix = f"~{n}"
        candidate = name[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def build_excel(groups: dict) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)
    used_names: set = set()
    for name, rows in groups.items():
        ws = wb.create_sheet(title=safe_sheet_name(name, used_names))
        sheet_df = _stringify_complex_columns(pd.DataFrame(rows))
        for r in dataframe_to_rows(sheet_df, index=False, header=True):
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def group_by_make_model(records: list) -> dict:
    by_make: dict[str, dict[str, list]] = {}
    for record in records:
        extra = parse_formatted_extra_fields(record)
        make = sanitize_filename(extra.get("make"))
        model = clean_text(extra.get("model"))
        by_make.setdefault(make, {}).setdefault(model, []).append(record)
    return by_make


def split_vehicle_records(records: list) -> tuple[dict, dict]:
    manufacturer_groups: dict[str, dict] = {}
    other_groups: dict[str, dict] = {}
    for record in records:
        _, cat1, _ = parse_category(record.get("category"))
        if cat1 is None:
            slug = "uncategorized"
            name = "Uncategorized"
        else:
            slug = cat1.get("slug") or "uncategorized"
            name = clean_text(cat1.get("name_l1") or cat1.get("name") or "Uncategorized")
        if slug in VEHICLE_MANUFACTURER_SPLIT_SLUGS:
            group = manufacturer_groups.setdefault(slug, {"name": name, "records": []})
            group["records"].append(record)
        else:
            group = other_groups.setdefault(slug, {"name": name, "records": []})
            group["records"].append(record)
    return manufacturer_groups, other_groups


def group_by_subsubcategory(records: list) -> dict[str, list]:
    """Group a single subcategory's own records by their cat2 (sub-subcategory),
    falling back to cat1's own name when there's no deeper level."""
    groups: dict[str, list] = {}
    for record in records:
        _, cat1, cat2 = parse_category(record.get("category"))
        if cat2:
            name = clean_text(cat2.get("name_l1") or cat2.get("name") or "Other")
        elif cat1:
            name = clean_text(cat1.get("name_l1") or cat1.get("name") or "Other")
        else:
            name = "Other"
        groups.setdefault(name, []).append(record)
    return groups


def has_any_subsubcategory(records: list) -> bool:
    for record in records:
        _, _, cat2 = parse_category(record.get("category"))
        if cat2:
            return True
    return False


def build_subcategory_files(records: list) -> dict[str, dict[str, list]]:
    files: dict[str, dict[str, list]] = {}
    for record in records:
        _, cat1, cat2 = parse_category(record.get("category"))
        subcat_slug = (cat1.get("slug") if cat1 else None) or "uncategorized"
        if cat2:
            sheet_name = clean_text(cat2.get("name_l1") or cat2.get("name") or subcat_slug)
        elif cat1:
            sheet_name = clean_text(cat1.get("name_l1") or cat1.get("name") or subcat_slug)
        else:
            sheet_name = "Uncategorized"
        files.setdefault(subcat_slug, {}).setdefault(sheet_name, []).append(record)
    return files


def remove_category_column(groups):
    for _, rows in groups.items():
        for record in rows:
            record.pop("category", None)


# =============================================================================
# Build complete summary (combines request_stats + failed_items)
# =============================================================================

def format_failed_summary(failed_items: list, max_len: int = 400) -> str | None:
    """Format failed items into a short summary string."""
    if not failed_items:
        return None
    parts = []
    for item in failed_items[:12]:
        name = item.get("name", "?")
        count = item.get("errors", 0)
        detail = item.get("detail", "")
        bit = f"{name}: {count} error(s)"
        if detail:
            bit += f" ({detail})"
        parts.append(bit)
    text = "; ".join(parts)
    if len(failed_items) > 12:
        text += f"; +{len(failed_items) - 12} more"
    return text[:max_len]


def build_complete_summary(records: list, cat0_name_l1: str, cat0_slug: str, dt: datetime, cat0_name_ar: str = None) -> dict:
    """Build a complete summary.json containing all metrics."""
    
    # 1. Build basic summary
    groups: dict[str, dict] = {}
    for record in records:
        _, cat1, cat2 = parse_category(record.get("category"))
        if cat1 is None:
            key = "uncategorized"
            name_en = "Uncategorized"
            name_ar = "غير مصنف"
            slug = "uncategorized"
        else:
            slug = cat1.get("slug") or "uncategorized"
            key = slug
            name_en = cat1.get("name_l1") or cat1.get("name") or "Uncategorized"
            name_ar = cat1.get("name") or name_en

        group = groups.setdefault(key, {
            "name_ar": name_ar,
            "name_en": name_en,
            "slug": slug,
            "listings_count": 0,
            "_sub_seen": set(),
            "subcategories": [],
        })
        group["listings_count"] += 1

        if cat2:
            sub_name = cat2.get("name_l1") or cat2.get("name")
            if sub_name and sub_name not in group["_sub_seen"]:
                group["_sub_seen"].add(sub_name)
                group["subcategories"].append(sub_name)

    subcategories = [
        {
            "name_ar": g["name_ar"],
            "name_en": g["name_en"],
            "slug": g["slug"],
            "listings_count": g["listings_count"],
            "has_subcategories": bool(g["subcategories"]),
            "subcategories": g["subcategories"],
        }
        for g in groups.values()
    ]

    # 2. Read request_stats.json
    stats_file = f"request_stats_{cat0_slug}.json"
    request_metrics = {}
    requests_duration_sec = None
    
    if os.path.exists(stats_file):
        with open(stats_file, "r", encoding="utf-8") as f:
            stats_data = json.load(f)
        
        # Get actual request duration from stats (in minutes, convert to seconds)
        duration_min = stats_data.get("total_duration_min", 0)
        if duration_min:
            requests_duration_sec = duration_min * 60
        
        request_metrics = {
            "requests_total": stats_data.get("total_requests", 0),
            "requests_failed": 0,
            "duration_sec": stats_data.get("total_duration", 0),
            "requests_per_min": stats_data.get("total_req_per_min", 0),
            "requests_duration_sec": requests_duration_sec,
        }
    
    # 3. Read failed_pages.json
    failed_file = f"failed_pages_{cat0_slug}.json"
    failed_items = []
    total_failed = 0
    if os.path.exists(failed_file):
        with open(failed_file, "r", encoding="utf-8") as f:
            failed_data = json.load(f)
        total_failed = failed_data.get("total_failed", 0)
        request_metrics["requests_failed"] = total_failed
        for page in failed_data.get("failed_pages", []):
            failed_items.append({
                "name": f"{page.get('category', 'unknown')}-page-{page.get('page', '?')}",
                "errors": 1,
                "detail": page.get("error", "Unknown error")
            })
    
    # 4. Calculate error_rate_pct
    total_requests = request_metrics.get("requests_total", 0)
    if total_requests > 0:
        request_metrics["error_rate_pct"] = round(total_failed / total_requests * 100, 2)
    else:
        request_metrics["error_rate_pct"] = None
    
    # 5. Calculate requests_per_min from actual request duration
    if requests_duration_sec and requests_duration_sec > 0:
        request_metrics["requests_per_min"] = round(
            request_metrics["requests_total"] / (requests_duration_sec / 60.0), 2
        )
    
    # 6. Build final summary
    return {
        "scraped_at": datetime.now(timezone.utc).isoformat(),
        "data_scraped_date": dt.strftime("%Y-%m-%d"),
        "saved_to_R2_date": dt.strftime("%Y-%m-%d"),
        "category": {
            "name_ar": cat0_name_ar or cat0_name_l1,
            "name_en": cat0_name_l1,
            "slug": cat0_slug,
            "r2_path": resolve_category_r2_path(cat0_slug),
        },
        "workflow_name": "doman",
        "total_subcategories": len(subcategories),
        "total_listings": len(records),
        "subcategories": subcategories,
        "request_metrics": request_metrics,
        "failed_items": failed_items,
        "failed_items_summary": format_failed_summary(failed_items),
    }


# =============================================================================
# Main run function
# =============================================================================

def run(csv_path: str, date_str: str | None = None, skip_summary: bool = False):
    dt = (
        datetime.strptime(date_str, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        if date_str
        else datetime.now(timezone.utc)
    )
    
    # Get workflow global start time from GitHub Actions
    workflow_global_start = os.getenv("WORKFLOW_GLOBAL_START")
    workflow_global_duration = None
    
    if workflow_global_start:
        try:
            start_ts = float(workflow_global_start)
            now_ts = time.time()
            workflow_global_duration = round(now_ts - start_ts, 2)
            print(f"✅ Global workflow duration: {workflow_global_duration}s")
        except (ValueError, TypeError):
            print("⚠️ Warning: Could not parse WORKFLOW_GLOBAL_START")
    
    # Load raw data
    df = load_raw(csv_path)
    if df is None or df.empty:
        print(f"{csv_path} is missing or empty -- nothing to clean or upload.")
        return

    # Drop unwanted columns
    existing_cols = [c for c in COLUMNS_TO_DROP if c in df.columns]
    if existing_cols:
        df = df.drop(columns=existing_cols)
        print(f"  Dropped columns: {existing_cols}")

    # Clean data and fetch contact info
    with Stealth().use_sync(sync_playwright()) as p:
        browser = p.chromium.launch(headless=True, channel="chrome")
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            locale="en-US",
            timezone_id="Asia/Riyadh",
        )
        page = context.new_page()
        try:
            cat0_name_l1, cat0_name_ar, cat0_slug, sheets, records = clean_and_group(df, page=page, dt=dt)
        finally:
            browser.close()

    if not cat0_name_l1:
        print(f"No usable category data found in {csv_path}")
        return

    print(f"Category: {cat0_name_l1} ({cat0_slug}) -- {len(sheets)} sheet(s), {len(records)} ad(s)")
    for name, rows in sheets.items():
        print(f"  - {name}: {len(rows)}")

    # Upload files based on category type
    if cat0_slug == "vehicles":
        manufacturer_groups, other_groups = split_vehicle_records(records)

        # cars-for-sale / cars-for-rent: split by brand into files, each
        # file split into sheets by model -> vehicles/<slug>/excel|json/
        for slug, group in manufacturer_groups.items():
            r2_path = resolve_category_r2_path(cat0_slug, slug)
            by_make = group_by_make_model(group["records"])
            for make, models in by_make.items():
                total_ads = sum(len(rows) for rows in models.values())
                print(f"    - {make}: {len(models)} model(s), {total_ads} ad(s)")

                excel_buf = build_excel(models)
                excel_key = upload_buffer(
                    excel_buf,
                    filename=f"{sanitize_filename(make)}.xlsx",
                    category_display=r2_path,
                    file_type="excel",
                    content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    dt=dt,
                )
                print(f"      Excel -> {excel_key}")

                json_bytes = json.dumps(models, ensure_ascii=False, indent=2, default=str).encode("utf-8")
                json_key = upload_buffer(
                    io.BytesIO(json_bytes),
                    filename=f"{sanitize_filename(make)}.json",
                    category_display=r2_path,
                    file_type="json",
                    content_type="application/json",
                    dt=dt,
                )
                print(f"      JSON  -> {json_key}")

        # every other vehicle subcategory (boats, motorcycles, trucks,
        # spare-parts, vip-car-plates, other-vehicles, car-accessories, ...)
        # gets its own folder -> vehicles/<slug>/excel|json/
        for slug, group in other_groups.items():
            r2_path = resolve_category_r2_path(cat0_slug, slug)
            sub_sheets = group_by_subsubcategory(group["records"])
            total_ads = sum(len(rows) for rows in sub_sheets.values())
            print(f"    - {slug}: {len(sub_sheets)} sheet(s), {total_ads} ad(s)")

            excel_buf = build_excel(sub_sheets)
            excel_key = upload_buffer(
                excel_buf,
                filename=f"{slug}.xlsx",
                category_display=r2_path,
                file_type="excel",
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                dt=dt,
            )
            print(f"      Excel -> {excel_key}")

            json_bytes = json.dumps(sub_sheets, ensure_ascii=False, indent=2, default=str).encode("utf-8")
            json_key = upload_buffer(
                io.BytesIO(json_bytes),
                filename=f"{slug}.json",
                category_display=r2_path,
                file_type="json",
                content_type="application/json",
                dt=dt,
            )
            print(f"      JSON  -> {json_key}")

    elif has_any_subsubcategory(records):
        r2_path = resolve_category_r2_path(cat0_slug)
        subcat_files = build_subcategory_files(records)
        for subcat_slug, sheets in subcat_files.items():
            total_ads = sum(len(rows) for rows in sheets.values())
            print(f"  {subcat_slug}: {len(sheets)} sheet(s), {total_ads} ad(s)")
            
            excel_buf = build_excel(sheets)
            excel_key = upload_buffer(
                excel_buf,
                filename=f"{subcat_slug}.xlsx",
                category_display=r2_path,
                file_type="excel",
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                dt=dt,
            )
            print(f"    Excel -> {excel_key}")
            
            json_bytes = json.dumps(sheets, ensure_ascii=False, indent=2, default=str).encode("utf-8")
            json_key = upload_buffer(
                io.BytesIO(json_bytes),
                filename=f"{subcat_slug}.json",
                category_display=r2_path,
                file_type="json",
                content_type="application/json",
                dt=dt,
            )
            print(f"    JSON  -> {json_key}")

    else:
        r2_path = resolve_category_r2_path(cat0_slug)
        excel_buf = build_excel(sheets)
        excel_key = upload_buffer(
            excel_buf,
            filename=f"{cat0_slug}.xlsx",
            category_display=r2_path,
            file_type="excel",
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            dt=dt,
        )
        print(f"Excel -> {excel_key}")
        
        json_bytes = json.dumps(sheets, ensure_ascii=False, indent=2, default=str).encode("utf-8")
        json_key = upload_buffer(
            io.BytesIO(json_bytes),
            filename=f"{cat0_slug}.json",
            category_display=r2_path,
            file_type="json",
            content_type="application/json",
            dt=dt,
        )
        print(f"JSON  -> {json_key}")

    # ========================================================================
    # Build summary but don't upload if skip_summary is True
    # ========================================================================
    summary = build_complete_summary(records, cat0_name_l1, cat0_slug, dt, cat0_name_ar)
    summary_r2_path = resolve_category_r2_path(cat0_slug)
    
    if skip_summary:
        # ✅ Save placeholder locally (to be finalized later)
        placeholder_path = f"summary_placeholder_{cat0_slug}.json"
        with open(placeholder_path, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        print(f"✅ Summary placeholder saved: {placeholder_path}")
        print(f"  (Will be finalized with workflow duration later)")
    else:
        # ✅ Upload directly (old behavior)
        summary_bytes = json.dumps(summary, ensure_ascii=False, indent=2).encode("utf-8")
        summary_key = upload_buffer(
            io.BytesIO(summary_bytes),
            filename="summary.json",
            category_display=summary_r2_path,
            file_type="summary",
            content_type="application/json",
            dt=dt,
        )

        print(f"Summary -> {summary_key}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("csv_path", help="Path to the raw scraped CSV")
    parser.add_argument("--date", default=None, help="YYYY-MM-DD data date")
    parser.add_argument("--skip-summary", action="store_true", 
                        help="Skip uploading summary, save placeholder instead")
    args = parser.parse_args()
    run(args.csv_path, args.date, args.skip_summary)