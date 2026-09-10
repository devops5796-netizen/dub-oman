import argparse
import json
import random
import time

import pandas as pd
import requests
from request_tracker import tracker

BASE_URL = "https://content.dubizzle.com.om/api/new-cars/all-new-cars"
DETAILS_BASE = "https://content.dubizzle.com.om/api"


HEADERS = {
    "accept": "application/json, text/plain, */*",
    "origin": "https://www.dubizzle.com.om",
    "referer": "https://www.dubizzle.com.om/",
    "user-agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/150.0.0.0 Safari/537.36"
    ),
}


def get_all_motors():
    all_cars = []

    response = requests.get(BASE_URL, params={"page": 1}, headers=HEADERS, timeout=30)
    response.raise_for_status()
    tracker.log_request(source="listing_pages", success=True)

    data = response.json()
    total_pages = data["page_info"]["total_pages"]
    print(f"Total pages: {total_pages}")
    all_cars.extend(data["cars"])

    for page in range(2, total_pages + 1):
        print(f"Fetching page {page}/{total_pages}")
        try:
            response = requests.get(BASE_URL, params={"page": page}, headers=HEADERS, timeout=30)
            response.raise_for_status()
            tracker.log_request(source="listing_pages", success=True)
            data = response.json()
            all_cars.extend(data["cars"])
        except Exception as e:
            tracker.log_request(source="listing_pages", success=False)
            print(f"  [ERROR] page {page}: {e}")

        delay = random.uniform(3, 6)
        print(f"  Waiting {delay:.2f}s before next request...")
        time.sleep(delay)

    print(f"Total cars collected: {len(all_cars)}")
    df = pd.DataFrame(all_cars)
    df.to_csv("new_cars.csv", index=False, encoding="utf-8-sig")
    print("Saved to new_cars.csv")
    return df


def cars_details(slug_url, max_retries: int = 3):
    url = f"{DETAILS_BASE}{slug_url}"

    data = None
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.get(url, headers=HEADERS, timeout=30)
            response.raise_for_status()
            tracker.log_request(source="car_details", success=True)
            print("URL:", response.url)
            print("STATUS:", response.status_code)
            print("CONTENT-TYPE:", response.headers.get("content-type"))
            print("RESPONSE:", response.text[:300])
            data = response.json()
            break
        except Exception as e:
            tracker.log_request(source="car_details", success=False)
            print(f"  [Attempt {attempt}/{max_retries}] {slug_url} failed: {e}")
            if attempt < max_retries:
                time.sleep(attempt * 2)

    if data is None:
        return None

    car_data = data.get('data', {})
    row_data = {}
    for key, value in car_data.items():
        if value is None:
            row_data[key] = ''
        elif isinstance(value, (str, int, float, bool)):
            row_data[key] = value
        elif isinstance(value, (dict, list)):
            row_data[key] = json.dumps(value, ensure_ascii=False)
        else:
            row_data[key] = str(value)

    return pd.DataFrame([row_data])


def _flatten_overview_data(df: pd.DataFrame) -> pd.DataFrame:
    col = "overview_data"
    if col not in df.columns:
        return df

    df[col] = df[col].apply(lambda x: json.loads(x) if pd.notna(x) and x else {})

    rows = []
    for item in df[col]:
        row = {}
        for key, value in item.items():
            if isinstance(value, (dict, list)):
                row[f"{col}.{key}"] = json.dumps(value, ensure_ascii=False)
            else:
                row[f"{col}.{key}"] = value
        rows.append(row)

    overview_df = pd.DataFrame(rows)
    return pd.concat([df.drop(columns=[col]), overview_df], axis=1)


def run_list():
    print('Scraping all motors new cars (listing pages only) ....')
    all_cars_df = get_all_motors()
    print(f"Listing phase done: {len(all_cars_df)} cars found. Run --mode details next.")


def run_details(start: int, end: int):
    all_cars_df = pd.read_csv("new_cars.csv")
    slug_url_list = all_cars_df['url'].values.tolist()[start:end]

    print(f"Scraping details for cars [{start}:{end}] ({len(slug_url_list)} cars)")

    dfs = []
    failed_urls = []

    for i, slug_url in enumerate(slug_url_list):
        print(f'Scraping [{i + 1}/{len(slug_url_list)}] (global index {start + i})')
        df = cars_details(slug_url)
        if df is not None:
            dfs.append(df)
        else:
            failed_urls.append(slug_url)
            print(f"  [FAILED] Skipping: {slug_url}")

        if i < len(slug_url_list) - 1:
            delay = random.uniform(3, 6)
            print(f"  Waiting {delay:.2f}s before next request...")
            time.sleep(delay)

    print('Finished getting details for this chunk.')

    total_attempted = len(slug_url_list)
    total_failed = len(failed_urls)
    failed_pct = round(total_failed / total_attempted * 100, 2) if total_attempted > 0 else 0

    with open(f"failed_urls_motors_{start}_{end}.json", "w", encoding="utf-8") as f:
        json.dump({
            "total_failed": total_failed,
            "failed_percentage": failed_pct,
            "failed_urls": failed_urls,
        }, f, ensure_ascii=False, indent=2)
    print(f"Failed: {total_failed}/{total_attempted} ({failed_pct}%)")

    stats = tracker.save(f"request_stats_motors_{start}_{end}.json")
    print(f"\n--- Chunk Request Stats ---")
    print(f"Total: {stats['total_requests']} req | {stats['total_req_per_min']} req/min")

    if not dfs:
        print("No car details scraped in this chunk -- nothing to save.")
        # still write empty outputs so the merge step has consistent files to skip
        pd.DataFrame().to_csv(f"car_details_{start}_{end}.csv", index=False, encoding="utf-8-sig")
        return

    chunk_df = pd.concat(dfs, ignore_index=True)
    chunk_df = _flatten_overview_data(chunk_df)

    chunk_df.to_csv(f"car_details_{start}_{end}.csv", index=False, encoding="utf-8-sig")
    print(f"Saved: car_details_{start}_{end}.csv ({len(chunk_df)} rows)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=["list", "details"], required=True)
    parser.add_argument("--start", type=int, default=0)
    parser.add_argument("--end", type=int, default=None)
    args = parser.parse_args()

    if args.mode == "list":
        run_list()
    else:
        if args.end is None:
            parser.error("--end is required for --mode details")
        run_details(args.start, args.end)