import random
import re
from text_utils import clean_text
from request_tracker import tracker

AD_URL_TEMPLATE = "https://www.dubizzle.com.om/en/ad/{slug}-ID{externalID}.html"

CONTACT_BUTTON_SELECTORS = [
    'button:has-text("Show phone number")',
    'button:has-text("Show Phone Number")',
    'button:has-text("Show Number")',
    'button:has-text("Call")',
    'button:has-text("اتصل")',
    'button:has-text("عرض")',
    'button:has-text("Phone")',
    '[data-testid*="phone" i]',
    '[data-testid*="show-phone" i]',
    '[data-testid="call-cta-button"]',
    'button[class*="phone"]',
    'a[class*="phone"]',
    '[class*="contact"] button',
    '[class*="contact"] a',
]

EMPTY_CONTACT_INFO = {}


def build_ad_url(record: dict) -> str | None:
    """
    Ad pages look like:
    https://www.dubizzle.com.om/en/ad/{slug}-ID{externalID}.html

    Builds it from the record's own `externalID` + `slug` fields. Verify these column
    names match your raw CSV -- adjust if the ES source uses different keys
    (e.g. externalID instead of id).
    """
    ad_id = record.get("externalID")
    slug = record.get("slug")
    if not ad_id or not slug:
        return None
    slug = re.sub(r"[^a-zA-Z0-9\-]+", "-", clean_text(slug)).strip("-").lower()
    return AD_URL_TEMPLATE.format(slug=slug or "ad", externalID=ad_id)


def has_valid_phone(data: dict) -> bool:
    """Return True if data has a real phone number (not null/N/A/empty)."""
    if not isinstance(data, dict):
        return False
    for key in ("mobile", "whatsapp", "proxyMobile"):
        val = data.get(key)
        if val is None:
            continue
        val_str = str(val).strip()
        if val_str and val_str.lower() not in ("n/a", "null", "none", "nan", ""):
            digits = re.sub(r'\D', '', val_str)
            if len(digits) >= 7:
                return True
    
    # Check mobileNumbers list (Oman API returns this)
    mobile_numbers = data.get("mobileNumbers")
    if isinstance(mobile_numbers, list):
        for num in mobile_numbers:
            digits = re.sub(r'\D', '', str(num))
            if len(digits) >= 7:
                return True
    return False


def _call_api_directly(page, listing_id: str, ad_url: str):
    api_url = f"https://www.dubizzle.com.om/api/listing/{listing_id}/contactInfo/"
    try:
        resp = page.request.get(
            api_url,
            headers={
                "Accept": "application/json",
                "X-Requested-With": "XMLHttpRequest",
                "Referer": ad_url,
            },
            timeout=10000,
        )
        if resp.status == 200:
            return resp.json()
    except Exception as e:
        print(f"      [API-DIRECT] Failed: {e}")
    return None


def _try_fetch_once(page, ad_url: str, listing_id: str):
    page.goto(ad_url, wait_until="domcontentloaded", timeout=30000)
    tracker.log_request(source="scraping_phone_num")
    page.wait_for_timeout(random.uniform(1500, 2500))

    # 1) Try API directly first
    data = _call_api_directly(page, listing_id, ad_url)
    if has_valid_phone(data):
        return data

    # 2) Look for phone button
    call_button = None
    for selector in CONTACT_BUTTON_SELECTORS:
        loc = page.locator(selector).first
        try:
            if loc.is_visible(timeout=2000):
                call_button = loc
                break
        except Exception:
            continue

    if call_button is None:
        return {"_no_phone": True}

    call_button.scroll_into_view_if_needed()
    page.wait_for_timeout(300)
    call_button.click(force=True)
    page.wait_for_timeout(3000)

    # 3) Try API again after click
    data = _call_api_directly(page, listing_id, ad_url)
    if has_valid_phone(data):
        return data

    # 4) If API returned name but no valid phone, treat as failure
    if isinstance(data, dict) and data.get("name") is not None:
        print(f"  [EMPTY-PHONE] {ad_url} | Name: {data.get('name')} (no valid number)")
        return None

    return None


def fetch_contact_info(page, ad_url: str, max_retries: int = 2) -> dict | None:
    match = re.search(r"ID(\d+)\.html", ad_url or "")
    if not match:
        print(f"  [PARSE-FAIL] {ad_url}")
        return None
    listing_id = match.group(1)

    for attempt in range(1, max_retries + 1):
        try:
            data = _try_fetch_once(page, ad_url, listing_id)
        except Exception as e:
            if "Timeout" in str(e) or "net::" in str(e):
                if attempt < max_retries:
                    wait = random.uniform(1, 3)
                    print(f"    [RETRY] network error (attempt {attempt}): {e}")
                    page.wait_for_timeout(wait * 1000)
                    continue
            print(f"  [NETWORK-FAIL] {ad_url} | {e}")
            return None

        if isinstance(data, dict) and data.get("_no_phone"):
            print(f"  [NO-BUTTON] {ad_url}")
            return None

        if isinstance(data, dict) and has_valid_phone(data):
            mobile = data.get("mobile") or data.get("whatsapp") or data.get("proxyMobile")
            #print(f"  [SUCCESS] {ad_url} | Name: {data['name']} | Mobile: {mobile}")
            print(f"  [SUCCESS] {ad_url}")
            return data

        if data is not None:
            print(f"  [EMPTY-API] {ad_url}")
            return None

        if attempt < max_retries:
            wait = random.uniform(1, 3)
            print(f"    [RETRY] empty response (attempt {attempt}), waiting {wait:.1f}s...")
            page.wait_for_timeout(wait * 1000)

    print(f"  [FAILED] {ad_url}")
    return None