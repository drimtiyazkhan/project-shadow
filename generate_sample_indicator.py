"""
Project Shadow — Sample Indicator Generator (Trailing 1 Month)
==================================================================
Automatically regenerates a "trailing 1 month" trial Pine Script sample,
reading directly from this repo's own data/activity_<year>.xlsx files —
no manual CSV upload needed.

This is the same generation logic as the original generate_hft_pine.py
(Colab script), adapted to:
  - Read data/activity_<year>.xlsx (this site's own database) instead of
    a manually uploaded bulk-deals CSV
  - Match tracked accounts by keyword (same approach as tracker.html),
    tolerant of minor legal-name formatting differences, instead of
    requiring an exact full legal name match
  - Run automatically in CI, chained to fire only after the daily fetch
    workflow succeeds (see .github/workflows/generate-sample-indicator.yml)

Run manually:  python scripts/generate_sample_indicator.py
Run by CI:     triggered by .github/workflows/generate-sample-indicator.yml,
               which itself only runs after Daily Bulk Deal Fetch succeeds.
"""

from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------

REPO_ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = REPO_ROOT / "data"
OUTPUT_PATH = REPO_ROOT / "downloads" / "OPERATOR_Sample_Trailing1Month.pine"

# (keyword to match against Client column, short code, display name)
# Keyword matching (not exact full-name matching) so minor legal-name
# formatting differences in the data never cause a tracked account to be
# silently dropped — mirrors the same approach used in tracker.html.
TRACKED_ACCOUNTS = [
    ("JUMP TRADING", "J", "Jump Trading Financial India Private Limited"),
    ("QE SECURITIES", "Q", "QE Securities LLP"),
    ("JUNOMONETA", "N", "Junomoneta Finsol Private Limited"),
    ("NK SECURITIES", "K", "NK Securities Research Private Limited"),
    ("HRTI", "H", "HRTI Private Limited"),
    ("GRAVITON", "G", "Graviton Research Capital LLP"),
    ("SHARE INDIA", "S", "Share India Securities Limited"),
    ("ELIXIR WEALTH", "E", "Elixir Wealth Management Private Limited"),
    ("D3 STOCK VISION", "D", "D3 Stock Vision LLP"),
    ("GOLDMINE STOCKS", "M", "Goldmine Stocks Private Limited"),
    ("MICROCURVES", "C", "Microcurves Trading Private Limited"),
    ("MUSIGMA", "MS", "Musigma Securities"),
]

TRAILING_DAYS = 30      # "1 month" window, trailing from the latest date found
CHUNK_SIZE = 30         # entries per load_N() function (keeps Pine local blocks small)


# ---------------------------------------------------------------------------
# STEP 1 — Load from this site's own database (not an uploaded CSV)
# ---------------------------------------------------------------------------

def match_tracked_account(client_name: str):
    """Returns (code, display_name) if client_name matches a tracked
    account by keyword, else None."""
    name = (client_name or "").upper()
    for keyword, code, display_name in TRACKED_ACCOUNTS:
        if keyword in name:
            return code, display_name
    return None


def load_database_rows() -> pd.DataFrame:
    """Loads every data/activity_<year>.xlsx file found in the repo and
    concatenates them — covers year-boundary cases where the trailing
    30-day window spans two files (e.g. early January)."""
    frames = []
    for path in sorted(DATA_DIR.glob("activity_*.xlsx")):
        try:
            df = pd.read_excel(path)
            frames.append(df)
        except Exception as e:
            print(f"WARNING: could not read {path.name}: {e}")

    if not frames:
        raise FileNotFoundError(
            f"No data/activity_*.xlsx files found in {DATA_DIR}. "
            "Run the daily fetch at least once before generating a sample."
        )

    combined = pd.concat(frames, ignore_index=True)
    combined["Date"] = pd.to_datetime(combined["Date"]).dt.strftime("%Y-%m-%d")
    return combined


def filter_tracked_rows(df: pd.DataFrame) -> pd.DataFrame:
    matches = df["Client"].apply(match_tracked_account)
    df = df.copy()
    df["_code"] = matches.apply(lambda m: m[0] if m else None)
    df["_display_name"] = matches.apply(lambda m: m[1] if m else None)
    return df[df["_code"].notna()]


# ---------------------------------------------------------------------------
# STEP 2 — Determine trailing window
# ---------------------------------------------------------------------------

def get_trailing_window(df: pd.DataFrame, trailing_days=TRAILING_DAYS):
    dates = pd.to_datetime(df["Date"])
    window_end = dates.max()
    window_start = window_end - timedelta(days=trailing_days - 1)
    return window_start, window_end


# ---------------------------------------------------------------------------
# STEP 3 — Aggregate by symbol/date/trader with qty & price
# ---------------------------------------------------------------------------

def aggregate(df: pd.DataFrame, window_start, window_end):
    grouped = defaultdict(lambda: defaultdict(dict))
    dt_series = pd.to_datetime(df["Date"])
    mask = (dt_series >= window_start) & (dt_series <= window_end)
    windowed = df[mask]

    for _, row in windowed.iterrows():
        dt = pd.to_datetime(row["Date"])
        sym = str(row["Symbol"]).strip().upper()
        dkey = dt.strftime("%Y%m%d")
        code = row["_code"]
        qty = row["Quantity"]
        price = row["Price"]
        bs = str(row["ActivityType"]).strip().upper()
        if bs not in ("BUY", "SELL"):
            continue
        grouped[(sym, dkey)][code][bs] = (qty, price)

    return grouped


def build_entries(grouped):
    entries = []
    for (sym, dkey), traders in sorted(grouped.items()):
        parts = []
        for code in sorted(traders.keys()):
            qp = traders[code]
            name = next(dn for kw, c, dn in TRACKED_ACCOUNTS if c == code)
            bq, bp = qp.get("BUY", (0, 0))
            sq, sp = qp.get("SELL", (0, 0))
            parts.append(f"{name}@{bq}@{bp}@{sq}@{sp}")
        val = "|".join(parts)
        entries.append(f'map.put(_m, "{sym}:{dkey}", "{val}")')
    return entries


# ---------------------------------------------------------------------------
# STEP 4 — Build Pine Script text (unchanged from the original generator)
# ---------------------------------------------------------------------------

PINE_HEADER_TEMPLATE = '''//@version=5
indicator("Project Shadow - Operator Activity (Trailing 1M Sample)", overlay=true, max_labels_count=500)

// SAMPLE VERSION — trailing 1 month: {start_date} to {end_date}
// Tracks bulk-deal activity by the 12 operator-linked accounts we track.
// Includes quantity + trade price detail per account per deal.
// This sample regenerates automatically every day from live data.
// Full-year indicators available at the Downloads section of the site.

col_arrow = input.color(#00E676, "Arrow color", group="Style")
col_glow  = input.color(color.new(#00E676, 85), "Candle glow", group="Style")
col_bg    = input.color(color.new(#001a0a, 15), "Label bg", group="Style")
show_glow = input.bool(true, "Show candle glow", group="Style")
show_tip  = input.bool(true, "Show names in tooltip", group="Style")
arr_off   = input.float(0.4, "Arrow gap (%)", minval=0.05, maxval=5.0, step=0.05, group="Style")

var map<string,string> _m = map.new<string,string>()

'''

PINE_FOOTER = '''
sym = str.upper(syminfo.ticker)
bdt = str.format_time(time, "yyyyMMdd", "Asia/Kolkata")
v   = map.get(_m, sym + ":" + bdt)
bc  = na(v) ? 0 : array.size(str.split(v, "|"))
bn  = na(v) ? "" : v

xp(s) =>
    string[] parts = str.split(s, "|")
    string result = ""
    for i = 0 to array.size(parts) - 1
        string[] f = str.split(array.get(parts, i), "@")
        nm = array.get(f, 0)
        result := result + nm + "\\n"
    result

sz = bc <= 2 ? "tiny" : bc <= 4 ? "small" : bc <= 6 ? "normal" : bc <= 8 ? "large" : bc <= 10 ? "huge" : "enormous"

if bc > 0
    y = low * (1.0 - arr_off / 100.0)
    tip = show_tip ? "OPERATOR ACTIVITY\\nStock: " + sym + "\\nDate: " + str.format_time(time, "yyyy-MM-dd", "Asia/Kolkata") + "\\nAccounts: " + str.tostring(bc) + " of 12\\n\\n" + xp(bn) : "Operator: " + str.tostring(bc) + " accounts"
    label.new(x=bar_index, y=y, text="▲" + str.tostring(bc), style=label.style_label_up, color=col_bg, textcolor=col_arrow, size=sz, tooltip=tip, xloc=xloc.bar_index, yloc=yloc.price)

bgcolor(show_glow and bc > 0 ? col_glow : na, title="Operator activity glow")
alertcondition(bc > 0, title="Operator Activity", message="Operator-linked activity on {{ticker}}")
'''


def build_functions_block(entries, chunk_size=CHUNK_SIZE):
    chunks = [entries[i:i + chunk_size] for i in range(0, len(entries), chunk_size)]
    func_blocks = []
    var_decls = []
    for idx, chunk in enumerate(chunks, start=1):
        body = "\n".join("    " + line for line in chunk)
        func_blocks.append(f"load_{idx}() =>\n{body}\n    true\n")
        var_decls.append(f"var _l{idx} = load_{idx}()")
    return "\n".join(func_blocks) + "\n" + "\n".join(var_decls) + "\n"


# ---------------------------------------------------------------------------
# STEP 5 — Orchestration
# ---------------------------------------------------------------------------

def generate_sample_indicator(output_path=OUTPUT_PATH, trailing_days=TRAILING_DAYS, chunk_size=CHUNK_SIZE):
    df = load_database_rows()
    tracked = filter_tracked_rows(df)
    if tracked.empty:
        raise ValueError(
            "No rows matched any of the 12 tracked accounts. Check that the "
            "database's 'Client' column contains recognizable account names."
        )

    window_start, window_end = get_trailing_window(tracked, trailing_days)
    grouped = aggregate(tracked, window_start, window_end)
    entries = build_entries(grouped)

    header = PINE_HEADER_TEMPLATE.format(
        start_date=window_start.strftime("%d %b %Y"),
        end_date=window_end.strftime("%d %b %Y"),
    )
    functions_block = build_functions_block(entries, chunk_size)
    pine_script = header + functions_block + PINE_FOOTER

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(pine_script, encoding="utf-8")

    char_count = len(pine_script)
    print(f"Generated: {output_path}")
    print(f"Trailing window   : {window_start.date()} to {window_end.date()} ({trailing_days} days)")
    print(f"Stock/date entries: {len(entries)}")
    print(f"Character count   : {char_count} / 85000 ({char_count / 85000 * 100:.1f}%)")
    if char_count > 85000:
        print("WARNING: exceeds TradingView's 85,000 character limit. Reduce trailing_days.")

    return pine_script


if __name__ == "__main__":
    generate_sample_indicator()
