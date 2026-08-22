"""
Project Shadow — Historical CSV to Tracker XLSX Converter
============================================================
Run this locally (on your own computer) to convert an NSE bulk-deal CSV
export into the .xlsx format the website's tracker page reads.

USAGE:
    python convert_historical_csv.py "Bulk-Deals-01-01-2024-to-31-12-2024.csv" data/activity_2024.xlsx
    python convert_historical_csv.py "Bulk-Deals-01-01-2025-to-31-12-2025.csv" data/activity_2025.xlsx
    python convert_historical_csv.py "Bulk-Deals-01-01-2026-to-21-08-2026.csv" data/activity_2026.xlsx

REQUIREMENTS (install once):
    pip install pandas openpyxl

WHAT IT DOES:
    - Reads the raw NSE CSV (handles the trailing spaces in headers,
      the "7,48,608" comma-formatted numbers, and DD-MON-YYYY dates)
    - Renames columns to the standard schema: Date, Symbol, SecurityName,
      Client, ActivityType, Quantity, Price, Remarks
    - Converts dates to YYYY-MM-DD
    - Keeps ALL clients/accounts (no filtering) — filtering by tracked
      accounts happens later, on the website itself, not here
    - Adds a "_hash" column so the daily automated fetch script can tell
      what's already recorded and never creates duplicates
    - If the output file already exists, merges in only new rows instead
      of overwriting it (safe to re-run)
"""

import argparse
import hashlib
import sys
from pathlib import Path

import pandas as pd

# Standard schema used across the whole site (historical + daily fetch)
SCHEMA = ["Date", "Symbol", "SecurityName", "Client", "ActivityType", "Quantity", "Price", "Remarks"]

# Raw NSE CSV column name -> our standard column name
# (NSE headers have trailing spaces, e.g. "Date " — we strip those first)
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


def clean_number(val):
    """Turns '7,48,608' or '1,171.15' into a plain float. Leaves blanks/'-' as None."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s in ("", "-"):
        return None
    s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def row_hash(row) -> str:
    # Price intentionally excluded from the hash — different data sources
    # (CSV export vs live API) can round price slightly differently, and
    # we don't want that to create false duplicate/non-duplicate rows.
    key = "|".join(str(row.get(k, "")) for k in ["Date", "Symbol", "Client", "ActivityType", "Quantity"])
    return hashlib.sha256(key.encode("utf-8")).hexdigest()


def convert(input_csv: Path, output_xlsx: Path):
    print(f"Reading {input_csv} ...")
    df = pd.read_csv(input_csv, encoding="utf-8-sig")  # utf-8-sig strips the BOM NSE includes

    # Strip whitespace from header names (NSE exports them as "Date ", "Symbol ", etc.)
    df.columns = [c.strip() for c in df.columns]

    missing = [c for c in COLUMN_MAP if c not in df.columns]
    if missing:
        print(f"ERROR: Expected columns not found in the CSV: {missing}")
        print(f"Columns actually present: {list(df.columns)}")
        sys.exit(1)

    df = df.rename(columns=COLUMN_MAP)
    df = df[[v for v in COLUMN_MAP.values()]]

    # Normalize date: '17-AUG-2026' -> '2026-08-17'
    df["Date"] = pd.to_datetime(df["Date"].str.strip(), format="%d-%b-%Y", errors="coerce").dt.strftime("%Y-%m-%d")

    bad_dates = df["Date"].isna().sum()
    if bad_dates:
        print(f"WARNING: {bad_dates} rows had an unparseable date and will be dropped.")
        df = df.dropna(subset=["Date"])

    df["Symbol"] = df["Symbol"].astype(str).str.strip().str.upper()
    df["Client"] = df["Client"].astype(str).str.strip()
    df["ActivityType"] = df["ActivityType"].astype(str).str.strip().str.upper()
    df["Quantity"] = df["Quantity"].apply(clean_number)
    df["Price"] = df["Price"].apply(clean_number)

    df["_hash"] = df.apply(row_hash, axis=1)

    output_xlsx.parent.mkdir(parents=True, exist_ok=True)

    if output_xlsx.exists():
        print(f"{output_xlsx} already exists — merging in only new rows.")
        existing = pd.read_excel(output_xlsx)
        existing_hashes = set(existing["_hash"]) if "_hash" in existing.columns else set()
        new_rows = df[~df["_hash"].isin(existing_hashes)]
        combined = pd.concat([existing, new_rows], ignore_index=True)
        added = len(new_rows)
    else:
        combined = df
        added = len(df)

    combined.to_excel(output_xlsx, index=False)
    print(f"Done. {added} new rows added. Total rows in {output_xlsx.name}: {len(combined)}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert an NSE bulk deal CSV export into the tracker's xlsx format.")
    parser.add_argument("input_csv", help="Path to the raw NSE CSV file")
    parser.add_argument("output_xlsx", help="Path to write, e.g. data/activity_2024.xlsx")
    args = parser.parse_args()

    convert(Path(args.input_csv), Path(args.output_xlsx))
