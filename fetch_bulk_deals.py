"""
Project Shadow — Daily Bulk Deal Fetcher
==========================================
Fetches NSE bulk deal data once a day, filters it down to the 12 tracked
"operator-linked" accounts, and appends new rows into a year-wise Excel
file (e.g. data/activity_2026.xlsx). Never overwrites history — only
appends rows that aren't already present.

Run manually:   python scripts/fetch_bulk_deals.py
Run by CI:       triggered daily by .github/workflows/daily-bulk-deal-fetch.yml

IMPORTANT — READ BEFORE FIRST RUN
----------------------------------
NSE's website actively blocks plain scripted requests. This script first
"warms up" a session by visiting the NSE homepage (to collect cookies),
then calls the bulk-deals API using that session — this is the standard
workaround, but NSE changes its site/API occasionally without notice.

If this script fails with a 401/403 error, or the JSON keys below don't
match what NSE actually returns, that means NSE has changed something on
their end. Run once locally, look at the printed raw response, and send
the output back so the field-mapping below can be corrected.
"""

import datetime
import hashlib
import json
import logging
import os
import sys
import time
from pathlib import Path

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
LOG_DIR = Path(__file__).resolve().parent.parent / "logs"
DATA_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

IST = datetime.timezone(datetime.timedelta(hours=5, minutes=30))

# NSE bulk deals endpoint. VERIFY THIS against the live site — NSE has
# changed this URL/shape before. Community-documented endpoint as of last
# check:
NSE_HOME_URL = "https://www.nseindia.com/report-detail/display-bulk-and-block-deals"
NSE_API_URL = "https://www.nseindia.com/api/historical/bulk-deals"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/report-detail/display-bulk-and-block-deals",
}

# NOTE: All accounts/clients are stored, not just the 12 tracked ones —
# filtering to the watchlist happens on the website (tracker.html), not
# here, so you can widen or narrow your watchlist any time without losing
# data for accounts you didn't track yet.

RETRY_ATTEMPTS = 3
RETRY_WAIT_SECONDS = 15 * 60  # 15 minutes between retries

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_DIR / "fetch_log.txt"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("bulk_deal_fetcher")


def now_ist() -> datetime.datetime:
    return datetime.datetime.now(IST)


def get_session() -> requests.Session:
    """Warm up a session against NSE to obtain cookies required by the API."""
    session = requests.Session()
    session.headers.update(HEADERS)
    resp = session.get(NSE_HOME_URL, timeout=15)
    resp.raise_for_status()
    return session


def fetch_bulk_deals(session: requests.Session, from_date: datetime.date, to_date: datetime.date) -> list[dict]:
    """
    Calls the NSE bulk deals API for the given date range and returns the
    raw list of deal records. Field names below (BD_DT_DATE, BD_SYMBOL,
    etc.) are NSE's documented convention as of last check — verify against
    a live response and adjust FIELD_MAP if NSE has renamed anything.
    """
    params = {
        "from": from_date.strftime("%d-%m-%Y"),
        "to": to_date.strftime("%d-%m-%Y"),
    }
    resp = session.get(NSE_API_URL, params=params, headers=HEADERS, timeout=20)
    if resp.status_code != 200:
        log.error("NSE API returned status %s: %s", resp.status_code, resp.text[:300])
        resp.raise_for_status()

    try:
        payload = resp.json()
    except json.JSONDecodeError:
        log.error("Could not parse NSE response as JSON. First 300 chars:\n%s", resp.text[:300])
        raise

    records = payload.get("data", payload if isinstance(payload, list) else [])
    log.info("Fetched %d raw records from NSE for %s to %s", len(records), from_date, to_date)
    return records


# Maps our normalized column name -> list of possible NSE field names to try.
FIELD_MAP = {
    "Date": ["BD_DT_DATE", "mTIMESTAMP", "date"],
    "Symbol": ["BD_SYMBOL", "symbol"],
    "SecurityName": ["BD_SCRIP_NAME", "scripName"],
    "Client": ["BD_CLIENT_NAME", "clientName"],
    "ActivityType": ["BD_BUY_SELL", "buySell"],
    "Quantity": ["BD_QTY_TRD", "qty"],
    "Price": ["BD_TP_WATP", "watp", "price"],
}


def normalize_record(raw: dict) -> dict | None:
    row = {}
    for norm_key, candidates in FIELD_MAP.items():
        value = None
        for c in candidates:
            if c in raw and raw[c] not in (None, ""):
                value = raw[c]
                break
        row[norm_key] = value

    if not row["Client"] or not row["Symbol"]:
        return None
    return row


def row_hash(row: dict) -> str:
    """
    Stable fingerprint of a row, used to prevent duplicate entries.
    Price is intentionally excluded — the historical CSV importer and this
    live API fetch can round price slightly differently, and we don't want
    that to cause the same real-world deal to be stored twice.
    """
    key = "|".join(str(row.get(k, "")) for k in ["Date", "Symbol", "Client", "ActivityType", "Quantity"])
    return hashlib.sha256(key.encode("utf-8")).hexdigest()


def load_existing(year: int) -> pd.DataFrame:
    path = DATA_DIR / f"activity_{year}.xlsx"
    if path.exists():
        return pd.read_excel(path)
    return pd.DataFrame(columns=list(FIELD_MAP.keys()) + ["_hash"])


def save_year(year: int, df: pd.DataFrame) -> None:
    path = DATA_DIR / f"activity_{year}.xlsx"
    df.to_excel(path, index=False)
    log.info("Saved %d total rows to %s", len(df), path)


def append_rows(rows: list[dict]) -> int:
    """Groups rows by year, dedupes against existing data, appends new ones."""
    added_total = 0
    by_year: dict[int, list[dict]] = {}

    for row in rows:
        date_val = row.get("Date")
        if not date_val:
            continue
        try:
            parsed = pd.to_datetime(date_val, dayfirst=True)
        except Exception:
            log.warning("Skipping row with unparseable date: %r", date_val)
            continue
        row["Date"] = parsed.strftime("%Y-%m-%d")
        row["_hash"] = row_hash(row)
        by_year.setdefault(parsed.year, []).append(row)

    for year, year_rows in by_year.items():
        existing = load_existing(year)
        existing_hashes = set(existing["_hash"]) if "_hash" in existing.columns else set()

        new_rows = [r for r in year_rows if r["_hash"] not in existing_hashes]
        if not new_rows:
            log.info("Year %d: no new rows (all %d already recorded)", year, len(year_rows))
            continue

        new_df = pd.DataFrame(new_rows)
        combined = pd.concat([existing, new_df], ignore_index=True)
        save_year(year, combined)
        added_total += len(new_rows)
        log.info("Year %d: added %d new rows", year, len(new_rows))

    return added_total


def run_once() -> int:
    """Single fetch attempt. Returns number of new rows added."""
    today = now_ist().date()

    # Look back a few days too, so if a previous day's run failed or NSE
    # published late, we naturally backfill the gap without manual effort.
    from_date = today - datetime.timedelta(days=5)

    session = get_session()
    raw_records = fetch_bulk_deals(session, from_date, today)
    normalized = [normalize_record(r) for r in raw_records]
    normalized = [r for r in normalized if r is not None]
    log.info("%d of %d raw records were valid and usable", len(normalized), len(raw_records))

    added = append_rows(normalized)

    # Check whether today's date actually made it into the data — if not,
    # NSE likely hasn't published yet.
    today_str = today.strftime("%Y-%m-%d")
    got_today = any(r["Date"] == today_str for r in normalized)
    if not got_today:
        log.warning("No data found for today (%s) yet — NSE may not have published.", today_str)

    return added, got_today


def main():
    log.info("=== Bulk deal fetch run started (%s IST) ===", now_ist().isoformat())

    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            added, got_today = run_once()
        except Exception:
            log.exception("Fetch attempt %d/%d failed with an error", attempt, RETRY_ATTEMPTS)
            added, got_today = 0, False

        if got_today:
            log.info("Success: today's data was found and processed (%d new rows).", added)
            break

        if attempt < RETRY_ATTEMPTS:
            log.warning(
                "Today's data not yet available (attempt %d/%d). Retrying in %d minutes...",
                attempt, RETRY_ATTEMPTS, RETRY_WAIT_SECONDS // 60,
            )
            time.sleep(RETRY_WAIT_SECONDS)
        else:
            log.warning(
                "Gave up after %d attempts — today's data still not published. "
                "Tomorrow's run will backfill it automatically.",
                RETRY_ATTEMPTS,
            )

    log.info("=== Bulk deal fetch run finished ===\n")


if __name__ == "__main__":
    main()
