import argparse
import json
import os
import time
import random
from datetime import datetime, timezone, timedelta
import pandas as pd
import requests
from request_tracker import tracker
from dotenv import load_dotenv
load_dotenv()

URL = "https://search.mena.sector.run/_msearch"
#AUTHORIZATION = os.getenv("AUTHORIZATION")
AUTHORIZATION = "Basic b2x4LW9tLXByb2R1Y3Rpb24tc2VhcmNoOmg1PWl9alNnYSFGa1k2P0Y1NVZ0S0p6JFYkKkY1UT49"
INDEX = "olx-om-production-ads-ar"
LOCATION_ID = "0-1"

PAGE_SIZE = 100
MAX_RETRIES = 10

TARGET_DATE = (datetime.now(timezone.utc) - timedelta(days=1)).date()

CATEGORY_SLUGS = [
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
    "services"
]

headers = {
    "accept": "*/*",
    "authorization": AUTHORIZATION,
    "content-type": "application/x-ndjson",
    "origin": "https://www.dubizzle.com.om",
    "referer": "https://www.dubizzle.com.om/",
    "user-agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/150.0.0.0 Safari/537.36"
    ),
}

SORT = [
    {"extraFields.seller_verified": {"order": "desc"}},
    {"productScore": {"order": "desc"}},
    {"timestamp": {"order": "desc"}},
    {"id": {"order": "desc"}},
]

params = {
    "filter_path": (
        "took,"
        "*.hits.total.*,"
        "*.hits.hits._source.*,"
        "*.hits.hits.sort,"
        "*.error"
    )
}


def send_query(queries):
    payload = ""

    for query in queries:
        payload += json.dumps({"index": INDEX}) + "\n"
        payload += json.dumps(query) + "\n"

    for attempt in range(MAX_RETRIES):
        try:
            response = requests.post(
                URL,
                params=params,
                headers=headers,
                data=payload.encode("utf-8"),
                timeout=60,
            )

            if response.status_code in [429, 500, 502, 503, 504]:
                raise Exception(f"HTTP {response.status_code}")

            response.raise_for_status()
            return response.json()

        except Exception as e:
            print(f"Attempt {attempt + 1}/{MAX_RETRIES}: {e}")

            if attempt == MAX_RETRIES - 1:
                raise

            wait = (attempt + 1) * 5
            print(f"Retrying in {wait} seconds...")
            time.sleep(wait)


def build_query(category_slug, search_after=None, product=None):
    must = [
        {"term": {"category.slug": category_slug}},
        {"term": {"location.externalID": LOCATION_ID}},
    ]

    must_not = []

    if product is None:
        must_not.append({"terms": {"product": ["featured", "elite"]}})
    else:
        must.append({"term": {"product": product}})

    query = {
        "size": PAGE_SIZE,
        "track_total_hits": 200000,
        "query": {
            "bool": {
                "must": must,
                "must_not": must_not,
            }
        },
        "sort": [
            {"extraFields.seller_verified": {"order": "desc"}},
            {"productScore": {"order": "desc"}},
            {"timestamp": {"order": "desc"}},
            {"id": {"order": "desc"}},
        ],
        "timeout": "10s",
    }

    if search_after is not None:
        query["search_after"] = search_after

    return query


def scrape(category_slug):
    print(f"\n========== {category_slug} ==========")

    states = {
        "elite": {
            "records": [],
            "search_after": None,
            "finished": False,
        },
        "featured": {
            "records": [],
            "search_after": None,
            "finished": False,
        },
        "normal": {
            "records": [],
            "search_after": None,
            "finished": False,
        },
    }

    failed_pages = []
    page = 0

    while True:
        page += 1

        queries = []
        mapping = []

        for product in ["elite", "featured", None]:

            key = "normal" if product is None else product

            if states[key]["finished"]:
                continue

            queries.append(
                build_query(
                    category_slug,
                    search_after=states[key]["search_after"],
                    product=product,
                )
            )

            mapping.append(key)

        if not queries:
            break

        try:
            data = send_query(queries)
            tracker.log_request(source="scraping_pages", success=True)

        except Exception as e:

            tracker.log_request(source="scraping_pages", success=False)

            failed_pages.append(
                {
                    "category": category_slug,
                    "page": page,
                    "error": str(e),
                }
            )

            break

        responses = data.get("responses", [])

        for key, response in zip(mapping, responses):

            if "error" in response:

                failed_pages.append(
                    {
                        "category": category_slug,
                        "product": key,
                        "page": page,
                        "error": json.dumps(response["error"]),
                    }
                )

                states[key]["finished"] = True
                continue

            hits_obj = response.get("hits", {})

            total = hits_obj.get("total", {}).get("value", 0)
            hits = hits_obj.get("hits", [])

            print(
                f"{key.upper():9}"
                f"{len(hits):4} "
                f"Collected={len(states[key]['records'])}"
                f" Total={total}"
            )

            if not hits:
                states[key]["finished"] = True
                continue

            states[key]["records"].extend(
                hit["_source"]
                for hit in hits
            )

            if len(hits) < PAGE_SIZE:
                states[key]["finished"] = True

            else:
                states[key]["search_after"] = hits[-1]["sort"]

        delay = random.uniform(0.5, 2.0)
        time.sleep(delay)

    print(
        f"Normal={len(states['normal']['records'])} | "
        f"Featured={len(states['featured']['records'])} | "
        f"Elite={len(states['elite']['records'])}"
    )

    return (
        states["normal"]["records"],
        states["featured"]["records"],
        states["elite"]["records"],
        failed_pages,
    )


def filter_yesterday_hits(hits):
    filtered = []

    for hit in hits:
        created_at = hit.get("createdAt")

        if created_at is None:
            continue

        try:
            dt = datetime.fromtimestamp(float(created_at), tz=timezone.utc)

            if dt.date() == TARGET_DATE:
                filtered.append(hit)

        except (ValueError, TypeError):
            pass

    return filtered


def run(category_slug: str, out_dir: str = "."):
    normal, featured, elite, failed_pages = scrape(category_slug)

    all_records = normal + featured + elite
    print(f"Before yesterday filter ({TARGET_DATE}):", len(all_records))
    all_records = filter_yesterday_hits(all_records)
    print("After yesterday filter:", len(all_records))

    df = pd.DataFrame(all_records)

    print("Before dedup:", len(df))
    if "id" in df.columns:
        df = df.drop_duplicates(subset=["id"], keep="first")
    print("After dedup:", len(df))

    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"{category_slug}.csv")

    stats_file = f"request_stats_{category_slug}.json"
    stats = tracker.save(stats_file)

    # Written locally so clean_and_upload.py can pick it up and upload it
    # next to summary.json for this same category.
    failed_file = f"failed_pages_{category_slug}.json"
    with open(failed_file, "w", encoding="utf-8") as f:
        json.dump({
            "category": category_slug,
            "total_failed": len(failed_pages),
            "failed_pages": failed_pages,
        }, f, ensure_ascii=False, indent=2)
    print(f"Failed pages -> {failed_file} ({len(failed_pages)} failed)")

    print(f"\n--- Combined Request Stats ---")
    print(f"Total: {stats['total_requests']} req | {stats['total_req_per_min']} req/min")
    print(f"By source: {stats['per_source']}")

    if df.empty:
        print(f"Nothing to write for {category_slug} -- skipping CSV output.")
        return None

    df.to_csv(out_path, index=False, encoding="utf-8-sig")
    print(f"Done! -> {out_path}")
    return out_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--category", required=True, choices=CATEGORY_SLUGS,
        help="Top-level category slug to scrape",
    )
    parser.add_argument("--out-dir", default="data")
    args = parser.parse_args()
    run(args.category, args.out_dir)
