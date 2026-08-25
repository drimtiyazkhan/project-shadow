"""
Project Shadow — Daily Bulk Deal Fetcher (v2)
================================================
Fetches NSE's daily bulk deal CSV archive once a day and appends any new
rows into a year-wise Excel file (e.g. data/activity_2026.xlsx). Never
overwrites history — only appends rows that aren't already present.

WHAT CHANGED FROM v1:
    v1 guessed at NSE's JSON API (nseindia.com/api/historical/bulk-deals),
    which needs cookie/session tricks and turned out to be the wrong
    approach entirely.

    v2 instead fetches NSE's actual static daily archive file:
        https://nsearchives.nseindia.com/content/equities/bulk.csv
    This is the same CSV format NSE gives you when you download it
    manually from the website — same column headers, same date/number
    formatting — just automated. It updates once a day, typically
    available by 6:30–7:00 PM IST.

Run manually:   python scripts/fetch_bulk_deals.py
Run by CI:      triggered daily by .github/workflows/daily-bulk-deal-fetch.yml
"""

import datetime
import hashlib
import io
import logging
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

# NSE's static daily bulk deal archive — a plain CSV, replaced once a day.
NSE_CSV_URL = "https://nsearchives.nseindia.com/content/equities/bulk.csv"
NSE_HOME_URL = "https://www.nseindia.com/all-reports"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept": "text/csv,application/vnd.ms-excel,*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": "https://www.nseindia.com/all-reports",
}

RETRY_ATTEMPTS = 4
RETRY_WAIT_SECONDS = 20 * 60  # 20 minutes between retries

# Same schema/column-mapping used by scripts/convert_historical_csv.py, so
# a row fetched live and a row imported from a manually downloaded CSV
# produce an identical hash and never get double-counted.
COLUMN_MAP = {
    "Date": "Date",
    "Symbol": "Symbol",
    "Security Name": "SecurityName",
    "Client Name": "Client",
    "Buy / Sell": "ActivityType",
    "Quantity Traded": "Quantity",
    "Trade Price / Wght. Avg. Price": "Price",
    "Remarks": "Remarks",
}

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


def clean_number(val):
    if pd.isna(val):
        return 0.0
    s = str(val).strip()
    if s in ("", "-"):
        return 0.0
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def row_hash(row: dict) -> str:
    # Price intentionally excluded — a live fetch and a manually-imported
    # historical file can round price slightly differently; we don't want
    # that to create a false duplicate/non-duplicate for the same deal.
    key = "|".join(str(row.get(k, "")) for k in ["Date", "Symbol", "Client", "ActivityType", "Quantity"])
    return hashlib.sha256(key.encode("utf-8")).hexdigest()


def get_session() -> requests.Session:
    """
    Warm up a session against NSE's main site to collect cookies, as a
    safety net. The static archive file usually doesn't require this, but
    NSE's edge/CDN occasionally rejects requests with no prior cookie —
    warming up costs one extra request and avoids that failure mode.
    """
    session = requests.Session()
    session.headers.update(HEADERS)
    try:
        session.get(NSE_HOME_URL, timeout=15)
    except requests.RequestException as e:
        log.warning("Session warm-up request failed (continuing anyway): %s", e)
    return session


def fetch_csv_text(session: requests.Session) -> str:
    resp = session.get(NSE_CSV_URL, headers=HEADERS, timeout=20)
    log.info("GET %s -> HTTP %s (%d bytes)", NSE_CSV_URL, resp.status_code, len(resp.content))
    if resp.status_code != 200:
        log.error("Unexpected status. First 300 chars of response:\n%s", resp.text[:300])
        resp.raise_for_status()
    return resp.text


def parse_csv(text: str) -> pd.DataFrame:
    df = pd.read_csv(io.StringIO(text))
    df.columns = [c.strip() for c in df.columns]

    missing = [c for c in COLUMN_MAP if c not in df.columns]
    if missing:
        log.error("Expected columns not found in the fetched CSV: %s", missing)
        log.error("Columns actually present: %s", list(df.columns))
        raise ValueError(f"Unexpected CSV format from NSE — missing columns: {missing}")

    df = df.rename(columns=COLUMN_MAP)
    df = df[[v for v in COLUMN_MAP.values()]]

    df["Date"] = pd.to_datetime(df["Date"].astype(str).str.strip(), format="%d-%b-%Y", errors="coerce").dt.strftime("%Y-%m-%d")
    bad_dates = df["Date"].isna().sum()
    if bad_dates:
        log.warning("%d rows had an unparseable date and will be dropped.", bad_dates)
        df = df.dropna(subset=["Date"])

    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    df["Client"] = df["Client"].astype(str).str.strip()
    df["ActivityType"] = df["ActivityType"].astype(str).str.strip().str.upper()
    df["Quantity"] = df["Quantity"].apply(clean_number)
    df["Price"] = df["Price"].apply(clean_number)

    return df


def load_existing(year: int) -> pd.DataFrame:
    path = DATA_DIR / f"activity_{year}.xlsx"
    if path.exists():
        return pd.read_excel(path)
    return pd.DataFrame(columns=list(COLUMN_MAP.values()) + ["_hash"])


def save_year(year: int, df: pd.DataFrame) -> None:
    path = DATA_DIR / f"activity_{year}.xlsx"
    df.to_excel(path, index=False)
    log.info("Saved %d total rows to %s", len(df), path)


def append_rows(df: pd.DataFrame) -> int:
    """Groups rows by year, dedupes against existing data, appends new ones."""
    if df.empty:
        return 0

    df = df.copy()
    df["_year"] = pd.to_datetime(df["Date"]).dt.year
    df["_hash"] = df.apply(lambda r: row_hash(r.to_dict()), axis=1)

    added_total = 0
    for year, year_df in df.groupby("_year"):
        year_df = year_df.drop(columns=["_year"])
        existing = load_existing(int(year))
        existing_hashes = set(existing["_hash"]) if "_hash" in existing.columns else set()

        new_rows = year_df[~year_df["_hash"].isin(existing_hashes)]
        if new_rows.empty:
            log.info("Year %d: no new rows (all %d already recorded)", year, len(year_df))
            continue

        combined = pd.concat([existing, new_rows], ignore_index=True)
        save_year(int(year), combined)
        added_total += len(new_rows)
        log.info("Year %d: added %d new rows", year, len(new_rows))

    return added_total


def run_once():
    """Single fetch attempt. Returns (rows_added, got_todays_data)."""
    today = now_ist().date()
    today_str = today.strftime("%Y-%m-%d")

    session = get_session()
    csv_text = fetch_csv_text(session)
    df = parse_csv(csv_text)
    log.info("Parsed %d rows from NSE's archive file", len(df))

    if df.empty:
        return 0, False

    latest_date_in_file = df["Date"].max()
    log.info("Most recent date present in the fetched file: %s (today is %s)", latest_date_in_file, today_str)

    added = append_rows(df)
    got_today = latest_date_in_file == today_str
    return added, got_today


def main():
    log.info("=== Bulk deal fetch run started (%s IST) ===", now_ist().isoformat())

    for attempt in range(1, RETRY_ATTEMPTS + 1):
        try:
            added, got_today = run_once()
        except Exception:
            log.exception("Fetch attempt %d/%d raised an error", attempt, RETRY_ATTEMPTS)
            added, got_today = 0, False

        if got_today:
            log.info("Success: today's data was found and processed (%d new rows).", added)
            break

        if attempt < RETRY_ATTEMPTS:
            log.warning(
                "Today's data not yet reflected in NSE's file (attempt %d/%d). Retrying in %d minutes...",
                attempt, RETRY_ATTEMPTS, RETRY_WAIT_SECONDS // 60,
            )
            time.sleep(RETRY_WAIT_SECONDS)
        else:
            log.warning(
                "Gave up after %d attempts — NSE's file still doesn't show today's date. "
                "This is expected on weekends/holidays. If it's a trading day, check "
                "%s manually.",
                RETRY_ATTEMPTS, NSE_CSV_URL,
            )

    log.info("=== Bulk deal fetch run finished ===\n")


if __name__ == "__main__":
    main()
