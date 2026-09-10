import argparse
import glob
import io
import json
import os
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, timedelta

import pandas as pd
import requests as req
from PIL import Image
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

from r2_uploader import upload_buffer

COLUMNS_TO_DROP: list[str] = ['content', 'description', 'content_l1', 'description_l1', 'seo_links',
                              'mapped_model_id', 'mapped_make_id']

MAX_IMAGE_DIMENSION = 1280
WEBP_QUALITY = 75


# =============================================================================
# Helper functions (unchanged)
# =============================================================================

def clean_text(value) -> str:
    if value is None:
        return "Unknown"
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or "Unknown"


def sanitize_name(value) -> str:
    text = clean_text(value)
    text = re.sub(r'[\\/:*?"<>|]', "-", text)
    return text or "Unknown"


def find_col(df: pd.DataFrame, name: str) -> str | None:
    if name in df.columns:
        return name
    for c in df.columns:
        if c == name or c.endswith(f".{name}"):
            return c
    return None


def parse_images_field(value) -> list[dict]:
    if isinstance(value, list):
        return value
    if isinstance(value, str) and value:
        try:
            parsed = json.loads(value)
            return parsed if isinstance(parsed, list) else []
        except (json.JSONDecodeError, TypeError):
            return []
    return []


def extract_image_urls(images_field) -> list[str]:
    items = parse_images_field(images_field)
    urls = []
    for item in items:
        if isinstance(item, dict) and item.get("url"):
            urls.append(item["url"])
    return urls


def download_and_upload_images(urls: list[str], car_ref: str, dt: datetime) -> list[str]:
    r2_paths = []
    uploaded = 0
    failed = 0
    if not urls:
        return r2_paths
    prefix = car_ref or "unknown"
    for idx, img_url in enumerate(urls, start=1):
        filename = f"{prefix}-{idx}.webp"
        try:
            r = req.get(img_url, timeout=15)
            if r.status_code == 200:
                img = Image.open(io.BytesIO(r.content)).convert("RGB")
                buf = io.BytesIO()
                img.save(buf, format="WEBP", quality=WEBP_QUALITY, method=6)
                buf.seek(0)
                r2_key = upload_buffer(buf, filename=filename, category_display='motors', file_type="images",
                                        content_type="image/webp", dt=dt)
                if r2_key:
                    r2_paths.append(r2_key)
                    uploaded += 1
                else:
                    failed += 1
            else:
                failed += 1
        except Exception as e:
            print(f"    [ERROR] {filename}: {e}")
            failed += 1
    if uploaded or failed:
        print(f"    {prefix}: {uploaded} uploaded, {failed} failed out of {len(urls)}")
    return r2_paths


def load_raw(input_path: str) -> pd.DataFrame | None:
    if not os.path.exists(input_path) or os.path.getsize(input_path) == 0:
        return None
    if input_path.endswith(".json"):
        return pd.read_json(input_path)
    return pd.read_csv(input_path)


# =============================================================================
# Build complete summary for motors
# =============================================================================

def format_failed_summary(failed_items: list, max_len: int = 400) -> str | None:
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


def build_complete_summary_motors(by_make: dict, dt: datetime, stats_data: dict = None, failed_data: dict = None) -> dict:
    """Build complete summary for motors with request_metrics."""
    
    # 1. Build basic summary
    subcategories = []
    total_listings = 0
    for make_slug, models in by_make.items():
        listings_count = sum(len(rows) for rows in models.values())
        total_listings += listings_count
        model_names = sorted(models.keys())
        subcategories.append({
            "name_ar": "",
            "name_en": make_slug,
            "slug": make_slug,
            "listings_count": listings_count,
            "has_subcategories": len(model_names) > 1,
            "subcategories": model_names,
        })
    
    # 2. Build request_metrics
    request_metrics = {}
    requests_duration_sec = None
    
    if stats_data:
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
    
    # 3. Add failed_items
    failed_items = []
    total_failed = 0
    if failed_data:
        total_failed = failed_data.get("total_failed", 0)
        request_metrics["requests_failed"] = total_failed
        for url in failed_data.get("failed_urls", []):
            failed_items.append({
                "name": url,
                "errors": 1,
                "detail": "Failed to scrape details"
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
    
    # 6. Final summary
    return {
        "scraped_at": datetime.now(timezone.utc).isoformat(),
        "data_scraped_date": dt.strftime("%Y-%m-%d"),
        "saved_to_R2_date": dt.strftime("%Y-%m-%d"),
        "category": {
            "name_ar": "سيارات جديدة",
            "name_en": "Motors",
            "slug": "motors",
            "r2_path": "motors",
        },
        "workflow_name": "motors",
        "total_subcategories": len(subcategories),
        "total_listings": total_listings,
        "subcategories": subcategories,
        "request_metrics": request_metrics,
        "failed_items": failed_items,
        "failed_items_summary": format_failed_summary(failed_items),
    }


# =============================================================================
# Mode: images
# =============================================================================

def run_images_chunk(input_path: str, start: int, end: int, workers: int, output_csv: str):
    df = load_raw(input_path)
    if df is None or df.empty:
        print(f"{input_path} is missing or empty -- nothing to do.")
        pd.DataFrame(columns=["_row_id", "images_r2_paths"]).to_csv(output_csv, index=False)
        return
    if "_row_id" not in df.columns:
        raise SystemExit("Input file is missing '_row_id' -- run merge_car_details.py first.")
    chunk = df[(df["_row_id"] >= start) & (df["_row_id"] < end)].reset_index(drop=True)
    if chunk.empty:
        print(f"No rows in range [{start}:{end}).")
        pd.DataFrame(columns=["_row_id", "images_r2_paths"]).to_csv(output_csv, index=False)
        return
    images_col = find_col(chunk, "images")
    version_col = find_col(chunk, "version_id")
    data_date = datetime.now(timezone.utc)
    n = len(chunk)
    results = [None] * n

    def worker(pos: int, raw_images, car_ref: str) -> tuple:
        urls = extract_image_urls(raw_images)
        r2_paths = download_and_upload_images(urls, car_ref, data_date)
        return pos, r2_paths

    print(f"Downloading images for {n} cars [{start}:{end}) using {workers} workers...")
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {}
        for pos, row in chunk.iterrows():
            raw_images = row.get(images_col) if images_col else None
            row_id = row["_row_id"]
            car_ref = str(row.get(version_col) or row_id)
            futures[executor.submit(worker, pos, raw_images, car_ref)] = pos
        completed = 0
        for future in as_completed(futures):
            try:
                pos, r2_paths = future.result(timeout=180)
                results[pos] = r2_paths
            except Exception as e:
                pos = futures[future]
                print(f"    [ERROR] row {pos} failed: {e}")
                results[pos] = []
            completed += 1
            if completed % 50 == 0 or completed == n:
                print(f"    Progress: {completed}/{n}")
    out_df = pd.DataFrame({
        "_row_id": chunk["_row_id"],
        "images_r2_paths": [json.dumps(r) for r in results],
    })
    out_df.to_csv(output_csv, index=False)
    print(f"Saved: {output_csv} ({len(out_df)} rows)")


# =============================================================================
# Mode: finalize
# =============================================================================

def clean_and_split_prebuilt(df: pd.DataFrame) -> dict[str, dict[str, list]]:
    make_col = find_col(df, "make_slug")
    model_col = find_col(df, "model_slug")
    if make_col is None or model_col is None:
        raise ValueError("Could not find make_slug / model_slug columns in the data.")
    raw_images_col = find_col(df, "images")
    by_make: dict[str, dict[str, list]] = {}
    for _, row in df.iterrows():
        record = row.to_dict()
        # if raw_images_col:
        #     record.pop(raw_images_col, None)
        # r2_paths = record.pop("images_r2_paths", None)
        # record["images"] = r2_paths if isinstance(r2_paths, list) else []
        # Images step is disabled -- keep the original scraped image URLs
        # instead of R2 paths.
        raw_images_value = record.pop(raw_images_col, None) if raw_images_col else None
        record["images"] = extract_image_urls(raw_images_value)
        record.pop("images_r2_paths", None)
        record.pop("_row_id", None)
        make_slug = sanitize_name(record.get(make_col) or "unknown")
        model_slug = sanitize_name(record.get(model_col) or "unknown")
        by_make.setdefault(make_slug, {}).setdefault(model_slug, []).append(record)
    return by_make


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


def _stringify_complex_columns(sheet_df: pd.DataFrame) -> pd.DataFrame:
    for col in sheet_df.columns:
        sheet_df[col] = sheet_df[col].apply(
            lambda v: json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else v
        )
    return sheet_df


def build_excel(models: dict[str, list]) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)
    used_names: set = set()
    for model_slug, rows in models.items():
        ws = wb.create_sheet(title=safe_sheet_name(model_slug, used_names))
        sheet_df = _stringify_complex_columns(pd.DataFrame(rows))
        for r in dataframe_to_rows(sheet_df, index=False, header=True):
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def upload_by_make(by_make: dict[str, dict[str, list]], dt: datetime, local_dir: str | None = None) -> None:
    if local_dir:
        os.makedirs(local_dir, exist_ok=True)
    for make_slug, models in by_make.items():
        total_ads = sum(len(rows) for rows in models.values())
        print(f"  - {make_slug}: {len(models)} model(s), {total_ads} car(s)")
        excel_buf = build_excel(models)
        excel_key = upload_buffer(
            excel_buf, filename=f"{make_slug}.xlsx", category_display='motors', file_type="excel",
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            dt=dt,
        )
        print(f"      Excel -> {excel_key}")
        if local_dir:
            excel_buf.seek(0)
            with open(os.path.join(local_dir, f"{make_slug}.xlsx"), "wb") as f:
                f.write(excel_buf.read())
        json_bytes = json.dumps(models, ensure_ascii=False, indent=2, default=str).encode("utf-8")
        json_key = upload_buffer(
            io.BytesIO(json_bytes), filename=f"{make_slug}.json", category_display='motors', file_type="json",
            content_type="application/json", dt=dt,
        )
        print(f"      JSON  -> {json_key}")
        if local_dir:
            with open(os.path.join(local_dir, f"{make_slug}.json"), "wb") as f:
                f.write(json_bytes)


def run_finalize(input_path: str, images_dir: str, skip_summary: bool = False, local_output_dir: str | None = None):
    dt = datetime.now(timezone.utc)
    data_date = dt
    
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
    
    df = load_raw(input_path)
    if df is None or df.empty:
        print(f"{input_path} is missing or empty -- nothing to clean or upload.")
        return

    existing_cols = [c for c in COLUMNS_TO_DROP if c in df.columns]
    if existing_cols:
        df = df.drop(columns=existing_cols)
        print(f"  Dropped columns: {existing_cols}")

    # Images step is disabled -- no R2 image upload, no image-chunk merge.
    # Original scraped image URLs are kept as-is in clean_and_split_prebuilt.
    # if "_row_id" in df.columns:
    #     image_files = glob.glob(os.path.join(images_dir, "images_*.csv"))
    #     if image_files:
    #         image_parts = [pd.read_csv(f) for f in image_files if os.path.getsize(f) > 0]
    #         image_parts = [p for p in image_parts if not p.empty]
    #         if image_parts:
    #             images_merged = pd.concat(image_parts, ignore_index=True)
    #             images_merged["images_r2_paths"] = images_merged["images_r2_paths"].apply(
    #                 lambda v: json.loads(v) if pd.notna(v) and v else []
    #             )
    #             df = df.merge(images_merged, on="_row_id", how="left")
    #             print(f"  Merged image paths for {images_merged['images_r2_paths'].apply(bool).sum()}/{len(df)} rows")
    #         else:
    #             print("  No non-empty image chunk files found -- proceeding without images.")
    #     else:
    #         print(f"  No image chunk files found under {images_dir} -- proceeding without images.")

    by_make = clean_and_split_prebuilt(df)
    print(f"Split into {len(by_make)} make(s)")
    upload_by_make(by_make, data_date, local_dir=local_output_dir)

    # Read stats and failed data
    stats_data = None
    if os.path.exists("request_stats_motors.json"):
        with open("request_stats_motors.json", "r", encoding="utf-8") as f:
            stats_data = json.load(f)
            if workflow_global_duration is not None:
                stats_data["total_duration"] = workflow_global_duration
    
    failed_data = None
    if os.path.exists("failed_urls_motors.json"):
        with open("failed_urls_motors.json", "r", encoding="utf-8") as f:
            failed_data = json.load(f)

    # Build summary
    summary = build_complete_summary_motors(by_make, data_date, stats_data, failed_data)
    
    if local_output_dir:
        os.makedirs(local_output_dir, exist_ok=True)
        with open(os.path.join(local_output_dir, "summary.json"), "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)

    if skip_summary:
        # ✅ Save placeholder locally (to be finalized later)
        placeholder_path = "summary_placeholder_motors.json"
        with open(placeholder_path, "w", encoding="utf-8") as f:
            json.dump(summary, f, ensure_ascii=False, indent=2)
        print(f"✅ Summary placeholder saved: {placeholder_path}")
        print(f"  (Will be finalized with workflow duration later)")
    else:
        # ✅ Upload directly
        summary_bytes = json.dumps(summary, ensure_ascii=False, indent=2).encode("utf-8")
        summary_key = upload_buffer(
            io.BytesIO(summary_bytes),
            filename="summary.json",
            category_display='motors',
            file_type="summary",
            content_type="application/json",
            dt=data_date,
        )
        print(f"Summary -> {summary_key}")
    


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("input_path", nargs="?")
    parser.add_argument("--mode", choices=["images", "finalize"], default="finalize")
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--end", type=int, default=None)
    parser.add_argument("--workers", type=int, default=8)
    parser.add_argument("--output-csv", default=None)
    parser.add_argument("--images-dir", default="images_parts")
    parser.add_argument("--skip-summary", action="store_true",
                        help="Skip uploading summary, save placeholder instead")
    parser.add_argument("--local-output-dir", default=None,
                        help="If set, also save a local copy of each make's Excel/JSON and the "
                             "summary here (e.g. for a GitHub Actions artifact upload)")
    args = parser.parse_args()

    if not args.input_path:
        parser.error("input_path is required")

    if args.mode == "images":
        if args.end is None:
            parser.error("--end is required for --mode images")
        output_csv = args.output_csv or f"images_{args.start}_{args.end}.csv"
        run_images_chunk(args.input_path, args.start, args.end, args.workers, output_csv)
    else:
        run_finalize(args.input_path, args.images_dir, args.skip_summary, args.local_output_dir)