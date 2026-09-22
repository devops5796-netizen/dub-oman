#!/usr/bin/env python3
"""
R2 Contact Info Extractor — Oman (dubizzle.com.om / DOMAN)
============================================================
Usage:
  python r2_doman_contact_extractor.py <YYYY-MM-DD>                 # single day
  python r2_doman_contact_extractor.py <START_YYYY-MM-DD> <END_YYYY-MM-DD>   # backfill range, inclusive

Direct port of the DKSA extractor -- same dedup/merge logic, just:
  - BASE_PREFIX changed from "DKSA/" to "DOMAN/"
  - phone normalization adapted for Oman (+968, 8-digit local mobile)
    instead of Saudi (+966). Verify this against a real contact_info
    sample once you run it -- I don't have real Oman raw numbers to
    validate the exact format against.

Dedup logic (unchanged):
  - Group by name
  - If same name + same phone   → keep newest (last occurrence)
  - If same name + diff phone   → keep BOTH

Day-over-day accumulation (unchanged): each day reads the PREVIOUS day's
merged agent-agency file, merges in today's newly-found contacts, and
writes a fresh merged file under today's own day= folder. This means a
backfill MUST run every day in the range in chronological order -- you
can't skip straight to the last day, or you lose everything in between.
That's exactly what the two-date range mode below does.
"""

import os
import sys
import json
import ast
import re
import boto3
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
from collections import defaultdict

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
CF_R2_ACCESS_KEY = os.getenv("CF_R2_ACCESS_KEY_ID")
CF_R2_SECRET_KEY = os.getenv("CF_R2_SECRET_ACCESS_KEY")
CF_R2_ENDPOINT_URL = os.getenv("CF_R2_ENDPOINT_URL")
BUCKET_NAME = os.getenv("CF_R2_BUCKET_NAME", "")
BASE_PREFIX     = "DOMAN/"
OUTPUT_SUBDIR   = "agent-agency"
OUTPUT_FILENAME = "agent-agency.xlsx"

if not all([CF_R2_ENDPOINT_URL, CF_R2_ACCESS_KEY, CF_R2_SECRET_KEY, BUCKET_NAME]):
    print("ERROR: Set CF_R2_ENDPOINT_URL, CF_R2_ACCESS_KEY_ID, CF_R2_SECRET_ACCESS_KEY, CF_R2_BUCKET_NAME")
    sys.exit(1)

s3 = boto3.client(
    "s3",
    endpoint_url=CF_R2_ENDPOINT_URL,
    aws_access_key_id=CF_R2_ACCESS_KEY,
    aws_secret_access_key=CF_R2_SECRET_KEY,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def parse_date(date_str: str):
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return f"{dt.year:04d}", f"{dt.month:02d}", f"{dt.day:02d}"


def get_day_prefix(year: str, month: str, day: str):
    return f"{BASE_PREFIX}year={year}/month={month}/day={day}/"


def get_prev_day_prefix(year: str, month: str, day: str):
    dt = datetime(int(year), int(month), int(day)) - timedelta(days=1)
    return get_day_prefix(f"{dt.year:04d}", f"{dt.month:02d}", f"{dt.day:02d}")


def normalize_category(cat: str) -> str:
    cat = cat.lower().strip()
    cat = re.sub(r'[\s&]+', '-', cat)
    cat = re.sub(r'-+', '-', cat)
    cat = cat.strip('-')
    return cat


def list_folders(prefix: str):
    folders = []
    paginator = s3.get_paginator("list_objects_v2")
    for page in paginator.paginate(Bucket=BUCKET_NAME, Prefix=prefix, Delimiter="/"):
        for cp in page.get("CommonPrefixes", []):
            folders.append(cp["Prefix"])
    return sorted(folders)


def list_all_excel_keys(prefix: str):
    keys = []
    paginator = s3.get_paginator("list_objects_v2")
    for page in paginator.paginate(Bucket=BUCKET_NAME, Prefix=prefix):
        for obj in page.get("Contents", []):
            k = obj["Key"]
            if k.endswith(".xlsx") and not k.split("/")[-1].startswith("~$"):
                keys.append(k)
    return keys


def read_excel_sheets(key: str):
    try:
        resp = s3.get_object(Bucket=BUCKET_NAME, Key=key)
        data = resp["Body"].read()
        xl = pd.ExcelFile(BytesIO(data))
        return {name: xl.parse(name) for name in xl.sheet_names}
    except Exception as e:
        print(f"  ⚠️  Failed to read {key}: {e}")
        return {}


def safe_parse_dict(raw: str):
    raw = str(raw).strip()
    if not raw or raw.lower() in ("contact_info", "nan", "none", "nan", "null", ""):
        return None
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass
    try:
        return ast.literal_eval(raw)
    except (ValueError, SyntaxError):
        pass
    return None


# ---------------------------------------------------------------------------
# Phone validation & normalization (Oman)
# ---------------------------------------------------------------------------

def normalize_phone_number(phone: str) -> str:
    """Omani numbers: 8-digit local mobile -> 968XXXXXXXX.
    Oman doesn't use a trunk '0' prefix domestically, but we strip one
    defensively in case a source ever includes it. VERIFY this against a
    real contact_info sample -- adjust if the actual raw format differs."""
    digits = re.sub(r'\D', '', str(phone))
    if digits.startswith('968') and len(digits) == 11:
        return digits
    if digits.startswith('0') and len(digits) == 9:
        digits = digits[1:]
    if len(digits) == 8:
        digits = '968' + digits
    return digits


def has_valid_phone(contact: dict) -> bool:
    if not isinstance(contact, dict):
        return False
    for key in ("mobile", "whatsapp", "proxyMobile"):
        val = contact.get(key)
        if val is None:
            continue
        val_str = str(val).strip()
        if val_str and val_str.lower() not in ("n/a", "null", "none", "nan", ""):
            digits = normalize_phone_number(val_str)
            if len(digits) >= 7:
                return True

    mobile_numbers = contact.get("mobileNumbers")
    if isinstance(mobile_numbers, list):
        for num in mobile_numbers:
            digits = normalize_phone_number(str(num))
            if len(digits) >= 7:
                return True
    return False


def get_primary_phone(contact: dict) -> str | None:
    for key in ("mobile", "whatsapp", "proxyMobile"):
        val = contact.get(key)
        if val:
            digits = normalize_phone_number(str(val))
            if len(digits) >= 7:
                return digits

    mobile_numbers = contact.get("mobileNumbers") or []
    if isinstance(mobile_numbers, list):
        for num in mobile_numbers:
            digits = normalize_phone_number(str(num))
            if len(digits) >= 7:
                return digits
    return None


def is_valid_contact(contact: dict) -> bool:
    if not isinstance(contact, dict):
        return False
    name = str(contact.get("name") or "").strip()
    if not name or name.lower() in ("n/a", "null", "none", "nan"):
        return False
    return has_valid_phone(contact)


# ---------------------------------------------------------------------------
# Extraction
# ---------------------------------------------------------------------------

def extract_contacts_from_sheets(sheets: dict):
    contacts = []
    skipped = 0
    for sheet_name, df in sheets.items():
        if df.empty:
            continue

        contact_col = None
        for col in df.columns:
            if str(col).strip() == 'contact_info':
                contact_col = col
                break
        if contact_col is None:
            for col in df.columns:
                if str(col).strip().lower() == 'contactinfo':
                    contact_col = col
                    break
        if contact_col is None:
            continue

        for raw in df[contact_col].dropna().astype(str):
            obj = safe_parse_dict(raw)
            if isinstance(obj, dict) and obj.get("name") is not None:
                if is_valid_contact(obj):
                    contacts.append(obj)
                else:
                    skipped += 1
    if skipped:
        print(f"    ⚠️  Skipped {skipped} contact(s) with missing/invalid phone")
    return contacts


# ---------------------------------------------------------------------------
# Dedup by name, then by phone under same name
# ---------------------------------------------------------------------------

def dedup_by_name(contacts: list):
    """
    - Same name + same phone  → keep newest (last occurrence)
    - Same name + diff phone  → keep BOTH
    """
    by_name = defaultdict(list)
    for c in contacts:
        name = str(c.get("name") or "").strip()
        if not name:
            continue
        by_name[name].append(c)

    result = []
    for name, group in by_name.items():
        if len(group) == 1:
            result.append(group[0])
            continue

        # Group by normalized phone under this name
        by_phone = defaultdict(list)
        for c in group:
            phone = get_primary_phone(c)
            by_phone[phone or "__no_phone__"].append(c)

        # For each unique phone under this name, keep the LAST (newest)
        for phone, phone_group in by_phone.items():
            result.append(phone_group[-1])

    return result


# ---------------------------------------------------------------------------
# Previous day reader
# ---------------------------------------------------------------------------

def read_previous_contacts(prev_key: str):
    try:
        resp = s3.get_object(Bucket=BUCKET_NAME, Key=prev_key)
        data = resp["Body"].read()
        df = pd.read_excel(BytesIO(data))
        if df.empty:
            return []

        contacts = []
        for _, row in df.iterrows():
            record = {}
            for col in df.columns:
                val = row[col]
                if pd.isna(val):
                    record[col] = None
                else:
                    record[col] = val

            for list_col in ["mobileNumbers", "roles"]:
                if list_col in record and isinstance(record[list_col], str):
                    try:
                        record[list_col] = json.loads(record[list_col].replace("'", '"'))
                    except Exception:
                        record[list_col] = []
                elif list_col in record and record[list_col] is None:
                    record[list_col] = []

            contacts.append(record)
        return contacts
    except s3.exceptions.NoSuchKey:
        return []
    except Exception as e:
        print(f"  ⚠️  Failed to read previous {prev_key}: {e}")
        return []


# ---------------------------------------------------------------------------
# Writer
# ---------------------------------------------------------------------------

def write_contacts_excel(contacts: list, output_key: str):
    if not contacts:
        df = pd.DataFrame(columns=["name", "mobile", "whatsapp", "proxyMobile",
                                    "mobileNumbers", "roles"])
    else:
        df = pd.json_normalize(contacts)

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Contacts")
    buf.seek(0)
    s3.put_object(Bucket=BUCKET_NAME, Key=output_key, Body=buf.getvalue())
    return len(df)


# ---------------------------------------------------------------------------
# Core logic
# ---------------------------------------------------------------------------

def get_categories(day_prefix: str):
    cats = []
    for p in list_folders(day_prefix):
        cat = p.replace(day_prefix, "").strip("/")
        if cat and cat != OUTPUT_SUBDIR:
            cats.append(cat)
    return sorted(cats)


def process_category(day_prefix: str, category: str, prev_day_prefix: str):
    norm_cat = normalize_category(category)
    cat_prefix = f"{day_prefix}{category}/"
    excel_keys = list_all_excel_keys(cat_prefix)

    # 1. Extract today's valid contacts (recurses into any subfolders, e.g.
    #    vehicles/cars-for-sale/excel/ or properties/properties-for-rent/excel/)
    day_contacts = []
    for key in excel_keys:
        sheets = read_excel_sheets(key)
        day_contacts.extend(extract_contacts_from_sheets(sheets))

    # 2. Deduplicate TODAY by name+phone
    day_unique = dedup_by_name(day_contacts)

    # 3. Read previous day's merged file
    prev_key = f"{prev_day_prefix}{OUTPUT_SUBDIR}/{norm_cat}/{OUTPUT_FILENAME}"
    prev_contacts = read_previous_contacts(prev_key)
    prev_contacts = [c for c in prev_contacts if is_valid_contact(c)]

    # 4. Merge: previous first, then today
    merged = dedup_by_name(prev_contacts + day_unique)

    # 5. Write to R2
    output_key = f"{day_prefix}{OUTPUT_SUBDIR}/{norm_cat}/{OUTPUT_FILENAME}"
    n_rows = write_contacts_excel(merged, output_key)

    print(f"    → {category}: {len(day_unique)} new valid | {len(prev_contacts)} prev valid | {len(merged)} total unique")
    return len(merged)


def main(target_date: str):
    year, month, day = parse_date(target_date)
    day_prefix = get_day_prefix(year, month, day)
    prev_day_prefix = get_prev_day_prefix(year, month, day)

    print(f"📅 Target:  {day_prefix}")
    print(f"📅 Previous: {prev_day_prefix}")

    categories = get_categories(day_prefix)
    if not categories:
        print("   (no categories found)")
        return

    for cat in categories:
        process_category(day_prefix, cat, prev_day_prefix)

    print("\n✅ Done.")


def run_range(start_date: str, end_date: str):
    """Backfill mode: run every day from start_date to end_date, inclusive,
    STRICTLY in chronological order. Each day depends on the previous day's
    merged output already being written, so this can't be parallelized or
    run out of order."""
    start = datetime.strptime(start_date, "%Y-%m-%d").date()
    end = datetime.strptime(end_date, "%Y-%m-%d").date()
    if end < start:
        print(f"ERROR: end date {end_date} is before start date {start_date}")
        sys.exit(1)

    total_days = (end - start).days + 1
    current = start
    day_num = 0
    while current <= end:
        day_num += 1
        date_str = current.isoformat()
        print(f"\n{'=' * 70}")
        print(f"[{day_num}/{total_days}] {date_str}")
        print('=' * 70)
        main(date_str)
        current += timedelta(days=1)

    print(f"\n✅ Backfill complete: {start.isoformat()} -> {end.isoformat()} ({total_days} day(s)).")


if __name__ == "__main__":
    if len(sys.argv) == 2:
        main(sys.argv[1])
    elif len(sys.argv) == 3:
        run_range(sys.argv[1], sys.argv[2])
    else:
        print("Usage:")
        print("  python r2_doman_contact_extractor.py <YYYY-MM-DD>                      # single day")
        print("  python r2_doman_contact_extractor.py <START_YYYY-MM-DD> <END_YYYY-MM-DD>  # backfill range")
        sys.exit(1)