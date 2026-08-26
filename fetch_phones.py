#!/usr/bin/env python3
"""
Fetch phone numbers from R2 JSON files and re-upload updated JSON + Excel.
Runs SEQUENTIALLY (one file at a time) with delays between requests.
"""

import argparse
import json
import os
import io
import time
import random
import re
from datetime import datetime, timezone

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from playwright.sync_api import sync_playwright
from playwright_stealth import Stealth

from contact_info_fetcher import build_ad_url, fetch_contact_info, EMPTY_CONTACT_INFO, has_valid_phone
from r2_uploader import get_r2_client, BUCKET_NAME


def safe_sheet_name(name: str, used: set) -> str:
    name = re.sub(r"[:\\/?*\[\]]", "-", name)[:31] or "Sheet"
    candidate = name
    n = 1
    while candidate in used:
        suffix = f"~{n}"
        candidate = name[: 31 - len(suffix)] + suffix
        n += 1
    used.add(candidate)
    return candidate


def list_today_json_files(client, bucket, date_str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    prefix = f"DOMAN/year={dt.year}/month={dt.strftime('%m')}/day={dt.strftime('%d')}/"

    files = []
    paginator = client.get_paginator('list_objects_v2')
    for page in paginator.paginate(Bucket=bucket, Prefix=prefix):
        for obj in page.get('Contents', []):
            key = obj['Key']
            if '/json/' in key and key.endswith('.json'):
                files.append(key)
    return files


def download_json(client, bucket, key):
    resp = client.get_object(Bucket=bucket, Key=key)
    return json.loads(resp['Body'].read().decode('utf-8'))


def upload_json(client, bucket, key, data):
    json_bytes = json.dumps(data, ensure_ascii=False, indent=2, default=str).encode('utf-8')
    client.put_object(Bucket=bucket, Key=key, Body=json_bytes, ContentType='application/json')


def build_excel_from_json(data: dict) -> io.BytesIO:
    wb = Workbook()
    wb.remove(wb.active)
    used_names = set()
    for sheet_name, rows in data.items():
        ws = wb.create_sheet(title=safe_sheet_name(sheet_name, used_names))
        if not rows:
            continue
        sheet_df = pd.DataFrame(rows)
        for col in sheet_df.columns:
            sheet_df[col] = sheet_df[col].apply(
                lambda v: json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else v
            )
        for r in dataframe_to_rows(sheet_df, index=False, header=True):
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def upload_excel(client, bucket, key, data):
    buf = build_excel_from_json(data)
    client.put_object(
        Bucket=bucket, 
        Key=key, 
        Body=buf,
        ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


def process_file(client, bucket, json_key, page, delay_min, delay_max, skip_existing=True):
    print(f"\n{'='*60}")
    print(f"📄 Processing: {json_key}")

    data = download_json(client, bucket, json_key)

    total_records = sum(len(records) for records in data.values())
    processed = 0
    success = 0
    failed = 0
    skipped = 0

    for sheet_name, records in data.items():
        for record in records:
            processed += 1
            ad_url = build_ad_url(record)
            if not ad_url:
                record['contact_info'] = dict(EMPTY_CONTACT_INFO)
                failed += 1
                continue

            # ⏩ Skip if already has valid phone (for re-runs)
            if skip_existing:
                existing = record.get('contact_info')
                if isinstance(existing, dict) and has_valid_phone(existing):
                    skipped += 1
                    if processed % 10 == 0:
                        print(f"  ⏩ [{processed}/{total_records}] Skipped (already has phone)")
                    continue

            contact = fetch_contact_info(page, ad_url, max_retries=3)
            if contact and has_valid_phone(contact):
                record['contact_info'] = contact
                success += 1
                print(f"  ✅ [{processed}/{total_records}] {ad_url}")
            else:
                record['contact_info'] = dict(EMPTY_CONTACT_INFO) if contact is None else contact
                failed += 1
                print(f"  ❌ [{processed}/{total_records}] {ad_url}")

            # ⏳ Delay between requests (within file)
            if processed < total_records:
                delay = random.uniform(delay_min, delay_max)
                time.sleep(delay)

    # ☁️ Replace JSON
    upload_json(client, bucket, json_key, data)
    print(f"  ☁️  Re-uploaded JSON: {json_key}")

    # ☁️ Replace Excel (rebuilt from updated JSON)
    excel_key = json_key.replace('/json/', '/excel/').replace('.json', '.xlsx')
    upload_excel(client, bucket, excel_key, data)
    print(f"  ☁️  Re-uploaded Excel: {excel_key}")

    print(f"  📊 Stats: {success} success, {failed} failed, {skipped} skipped / {total_records}")
    return success, failed, skipped, total_records


def run(date_str: str, delay_min: float = 5.0, delay_max: float = 10.0, 
        max_files: int = None, categories: list = None, skip_existing: bool = True):
    client = get_r2_client()
    if not client or not BUCKET_NAME:
        print("❌ Failed to initialize R2 client")
        return

    bucket = BUCKET_NAME
    files = list_today_json_files(client, bucket, date_str)

    if categories:
        files = [f for f in files if any(cat in f for cat in categories)]

    if max_files:
        files = files[:max_files]

    print(f"🔍 Found {len(files)} JSON files for {date_str}")
    if not files:
        print("No files to process.")
        return

    total_success = 0
    total_failed = 0
    total_skipped = 0
    total_records = 0

    with Stealth().use_sync(sync_playwright()) as p:
        browser = p.chromium.launch(headless=True, channel="chrome")
        context = browser.new_context(
            viewport={"width": 1920, "height": 1080},
            locale="en-US",
            timezone_id="Asia/Riyadh",
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/150.0.0.0 Safari/537.36"
            ),
        )
        page = context.new_page()

        try:
            for idx, json_key in enumerate(files, 1):
                print(f"\n🗂️  File {idx}/{len(files)}")
                s, f, sk, t = process_file(
                    client, bucket, json_key, page, 
                    delay_min, delay_max, skip_existing
                )
                total_success += s
                total_failed += f
                total_skipped += sk
                total_records += t

                # ⏳ Delay between files
                if idx < len(files):
                    file_delay = random.uniform(delay_min * 2, delay_max * 2)
                    print(f"\n⏳ File delay: {file_delay:.1f}s before next file...")
                    time.sleep(file_delay)
        finally:
            browser.close()

    print(f"\n{'='*60}")
    print(f"🏁 DONE!")
    print(f"   Total records: {total_records}")
    print(f"   Success:       {total_success}")
    print(f"   Failed:        {total_failed}")
    print(f"   Skipped:       {total_skipped}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Fetch phone numbers from R2 files and re-upload")
    parser.add_argument("--date", required=True, help="YYYY-MM-DD data date")
    parser.add_argument("--delay-min", type=float, default=5.0, help="Min delay between requests (sec)")
    parser.add_argument("--delay-max", type=float, default=10.0, help="Max delay between requests (sec)")
    parser.add_argument("--max-files", type=int, default=None, help="Limit number of files to process")
    parser.add_argument("--categories", nargs="+", default=None, help="Filter by category slugs")
    parser.add_argument("--no-skip-existing", action="store_true", help="Re-fetch even if phone exists")
    args = parser.parse_args()

    run(
        args.date, 
        args.delay_min, 
        args.delay_max, 
        args.max_files, 
        args.categories,
        skip_existing=not args.no_skip_existing
    )