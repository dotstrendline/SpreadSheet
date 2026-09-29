#!/usr/bin/env python3
"""
NIFTY Tracker -> Google Sheets  (H1 / AVG_H1 breadth + Option Score)
======================================================================

One script, no HTML. Every cycle (5 min) it:

  1. H1 BREADTH  - fetches the NIFTY 50 constituents, classifies each as
                   H1 (ceiling) / L1 (floor), computes H1%, appends
                   time, H1, AVG_H1 to today's CSV  (data/H1_avg/).
  2. OPTION SCORE - fetches the NIFTY option chain (nearest expiry) and the
                   NIFTY 50 open, computes the Option Score per strike within
                   open +/- band, and updates today's Strike x Time pivot CSV
                   (data/option_score/).
  3. GOOGLE SHEET - pushes both of today's CSVs into your Google Spreadsheet:
                       tab  H1            (time | H1 | AVG_H1)
                       tab  OptionScore   (Strike price x HH:MM grid)
                       tab  Nifty50_Data  (latest snapshot of all 50 stocks)
                   Only these three tabs are used. They are rewritten in full
                   from today's CSVs each cycle, so a new trading day replaces
                   yesterday's view (history stays in the CSVs), and a failed
                   write is repaired automatically by the next run.

CORE LOGIC IS UNCHANGED from the previous two trackers:
  H1 : ltp >= dayHigh - 20% of range ; L1 : ltp <= dayLow + 20% of range
       (H1 wins ties; stocks in neither bucket are ignored)
       H1 = round(H1_count / (H1_count + L1_count) * 100)
       AVG_H1 = round(mean of every H1 logged today, including this one)
  Option Score = (Put %chg OI - Call %chg OI) / 10, where
       prev OI = OI - change in OI ; %chg = change / prev OI * 100
       rounded to the nearest whole number, ties away from zero (Excel-style)

USAGE
-----
    pip install -r requirements.txt
    python nifty_to_sheets.py --once        # one cycle (what GitHub Actions runs)
    python nifty_to_sheets.py               # local PC: loops every 5 minutes
    python nifty_to_sheets.py --once --no-sheets   # CSV only, skip Google

GOOGLE SETUP (details in README.md)
-----------------------------------
    SHEET_WEBHOOK_URL   Web app URL from your Google Sheet's Apps Script
    SHEET_TOKEN         the password you set inside apps_script.gs
"""

import argparse
import csv
import json
import os
import sys
import time
from datetime import datetime
from urllib.parse import quote

import requests

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
ROOT = os.path.dirname(os.path.abspath(__file__))
H1_AVG_DIR = os.path.join(ROOT, "data", "H1_avg")
OPTION_DIR = os.path.join(ROOT, "data", "option_score")

REFRESH_SECONDS = 300  # 5 minutes
FETCH_RETRIES = 3
R20_FRACTION = 0.20

INDEX_NAME = "NIFTY 50"
OPTION_SYMBOL = "NIFTY"
BAND = 1000  # +/- points around today's open

NSE_BASE = "https://www.nseindia.com"
NSE_LIVE_PAGE = f"{NSE_BASE}/market-data/live-equity-market"
NSE_OPTION_CHAIN_PAGE = f"{NSE_BASE}/option-chain"

# NSE has changed the endpoint behind the live-equity-market table before.
# Try the current one first, then fall back to the older one.
INDEX_API_NEW = f"{NSE_BASE}/api/NextApi/apiClient/marketWatchApi?functionName=getIndicesData&symbol="
INDEX_API_OLD = f"{NSE_BASE}/api/equity-stockIndices?index="

# option-chain-v3 needs an explicit expiry; contract-info lists valid expiries.
OPTION_CHAIN_CONTRACT_INFO_API = f"{NSE_BASE}/api/option-chain-contract-info?symbol="
OPTION_CHAIN_V3_API = f"{NSE_BASE}/api/option-chain-v3?type=Indices&symbol={{}}&expiry={{}}"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
    ),
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Connection": "keep-alive",
}


# ----------------------------------------------------------------------------
# Shared helpers
# ----------------------------------------------------------------------------
def new_session(referer: str = None) -> requests.Session:
    s = requests.Session()
    s.headers.update(HEADERS)
    if referer:
        s.headers["Referer"] = referer
    return s


def _ci_get(row: dict, *names):
    """Case/format-insensitive field lookup (dayHigh vs DAYHIGH vs DAY_HIGH)."""
    norm_map = {str(k).upper().replace(" ", "").replace("_", ""): row[k] for k in row}
    for n in names:
        key = n.upper().replace(" ", "").replace("_", "")
        if key in norm_map:
            return norm_map[key]
    return None


def _safe_float(val):
    if val is None:
        return None
    s = str(val).replace(",", "").strip()
    if s in ("", "-", "N/A", "nan", "None"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _extract_rows(payload: dict):
    """Old endpoint: {"data": [..]}   New endpoint: {"data": {"data": [..]}}"""
    d = payload.get("data")
    if isinstance(d, dict):
        return d.get("data", [])
    return d or []


def round_half_up(x):
    """Round to nearest whole number, ties away from zero (2.5 -> 3, -2.5 -> -3).
    Matches Excel ROUND(); Python's round() rounds ties to even."""
    if x >= 0:
        return int(x + 0.5)
    return -int(-x + 0.5)


def _fetch_index_payload(retries: int = FETCH_RETRIES) -> list:
    """Rows of the NIFTY 50 live-equity-market table (constituents + index row).
    Tries the newer endpoint first, then the legacy one, with retries."""
    session = new_session(NSE_LIVE_PAGE)
    quoted = quote(INDEX_NAME)

    last_err = None
    for attempt in range(1, retries + 1):
        for url in (INDEX_API_NEW + quoted, INDEX_API_OLD + quoted):
            try:
                # NSE only serves the API once these cookies exist on the session
                session.get(NSE_BASE, timeout=15)
                session.get(NSE_LIVE_PAGE, timeout=15)
                resp = session.get(url, timeout=20)
                resp.raise_for_status()
                rows = _extract_rows(resp.json())
                if rows:
                    return rows
                last_err = RuntimeError(f"{url.split('?')[0]} returned no rows")
            except Exception as exc:  # noqa: BLE001
                last_err = exc
                print(f"  attempt {attempt} {url.split('?')[0]} failed: {exc}", file=sys.stderr)
        time.sleep(2 * attempt)
    raise RuntimeError(f"Could not fetch live NSE data after {retries} attempts: {last_err}")


# ----------------------------------------------------------------------------
# 1) H1 breadth  (logic unchanged)
# ----------------------------------------------------------------------------
def stock_bucket(ltp: float, day_high: float, day_low: float) -> str:
    """'H1' if ltp >= dayHigh - 20% of range; else 'L1' if ltp <= dayLow + 20%
    of range; else ''. H1 is checked first, so it wins the zero-range tie."""
    r20 = R20_FRACTION * (day_high - day_low)
    if ltp >= day_high - r20:
        return "H1"
    if ltp <= day_low + r20:
        return "L1"
    return ""


def classify_h1_l1(rows: list):
    """Returns (h1_count, l1_count, total_fetched). The aggregate index row
    (symbol starting with 'NIFTY') is skipped. A stock in neither bucket is
    not counted toward Total_Active."""
    h1 = l1 = total_fetched = 0

    for row in rows:
        symbol = _ci_get(row, "SYMBOL")
        if not symbol or str(symbol).strip().upper().startswith("NIFTY"):
            continue
        total_fetched += 1

        day_high = _safe_float(_ci_get(row, "DAYHIGH", "HIGH"))
        day_low = _safe_float(_ci_get(row, "DAYLOW", "LOW"))
        ltp = _safe_float(_ci_get(row, "LASTPRICE", "LTP", "LAST"))
        if day_high is None or day_low is None or ltp is None:
            continue

        bucket = stock_bucket(ltp, day_high, day_low)
        if bucket == "H1":
            h1 += 1
        elif bucket == "L1":
            l1 += 1

    return h1, l1, total_fetched


H1_CSV_FIELDS = ["time", "H1", "AVG_H1"]

# Latest NIFTY 50 constituents fetched in THIS run (feeds the Nifty50_Data tab).
LATEST_NIFTY50 = {"rows": None, "time": None}


def h1_csv_path(for_date: datetime = None) -> str:
    d = (for_date or datetime.now()).strftime("%Y-%m-%d")
    return os.path.join(H1_AVG_DIR, f"nifty_h1_percentage_log_{d}.csv")


def write_h1_percentage_csv(h1_pct: float, checked_at: str) -> tuple:
    """Append time, H1 (whole number) and AVG_H1 (average of all H1 logged
    today INCLUDING this one, whole number) to today's CSV. A new day starts a
    new file, so the average only ever reflects today."""
    os.makedirs(H1_AVG_DIR, exist_ok=True)
    path = h1_csv_path()

    h1_rounded = round(h1_pct)

    previous = []
    if os.path.isfile(path):
        with open(path, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                try:
                    previous.append(float(row["H1"]))
                except (KeyError, ValueError, TypeError):
                    pass

    values = previous + [h1_rounded]
    avg_rounded = round(sum(values) / len(values))

    exists = os.path.isfile(path)
    with open(path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=H1_CSV_FIELDS)
        if not exists:
            writer.writeheader()
        writer.writerow({"time": checked_at, "H1": h1_rounded, "AVG_H1": avg_rounded})

    return h1_rounded, avg_rounded


def run_h1_cycle() -> bool:
    """Fetch -> classify -> append to today's CSV. Returns True if a new row
    was logged. On NSE failure nothing is written."""
    checked_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_time = datetime.now().strftime("%H:%M:%S")

    try:
        rows = _fetch_index_payload()
        h1, l1, total_fetched = classify_h1_l1(rows)
    except Exception as exc:  # noqa: BLE001
        print(f"[{log_time}] [H1 ERROR] Live NSE fetch failed: {exc} - will retry next cycle.")
        return False

    LATEST_NIFTY50["rows"], LATEST_NIFTY50["time"] = rows, checked_at

    active = h1 + l1
    h1_pct = round(h1 / active * 100) if active else 0
    h1_rounded, avg_rounded = write_h1_percentage_csv(h1_pct, checked_at)

    print(f"[{log_time}] H1 OK - fetched {total_fetched} | active {active} | "
          f"H1 {h1} / L1 {l1} | H1% {h1_rounded} | AVG_H1 {avg_rounded}")
    return True


# ----------------------------------------------------------------------------
# 2) Option score  (logic unchanged)
# ----------------------------------------------------------------------------
def fetch_index_open(index_name: str = INDEX_NAME, retries: int = FETCH_RETRIES) -> float:
    """Today's OPEN value for `index_name`."""
    last_err = None
    for attempt in range(1, retries + 1):
        session = new_session(NSE_LIVE_PAGE)
        for url in (INDEX_API_NEW + quote(index_name), INDEX_API_OLD + quote(index_name)):
            try:
                session.get(NSE_BASE, timeout=15)
                session.get(NSE_LIVE_PAGE, timeout=15)
                resp = session.get(url, timeout=20)
                resp.raise_for_status()
                for row in _extract_rows(resp.json()):
                    sym = _ci_get(row, "SYMBOL", "INDEX")
                    if sym and str(sym).strip().upper() == index_name.upper():
                        open_v = _safe_float(_ci_get(row, "OPEN"))
                        if open_v is not None:
                            return open_v
                last_err = RuntimeError(f"'{index_name}' row not found / no OPEN field")
            except Exception as exc:  # noqa: BLE001
                last_err = exc
                print(f"  index attempt {attempt} {url.split('?')[0]} failed: {exc}", file=sys.stderr)
        time.sleep(2 * attempt)
    raise RuntimeError(f"Could not fetch index open after {retries} attempts: {last_err}")


def _get_nearest_expiry(session, symbol):
    resp = session.get(OPTION_CHAIN_CONTRACT_INFO_API + quote(symbol), timeout=20)
    resp.raise_for_status()
    data = resp.json()
    dates = data.get("expiryDates") or data.get("records", {}).get("expiryDates")
    if not dates:
        raise RuntimeError("No expiry dates returned by option-chain-contract-info")
    return dates[0]


def fetch_option_chain_raw(symbol: str = OPTION_SYMBOL, retries: int = FETCH_RETRIES):
    """List of (strike, call_oi, call_chng_oi, put_oi, put_chng_oi) for the
    nearest expiry."""
    last_err = None

    for attempt in range(1, retries + 1):
        try:
            session = new_session(NSE_OPTION_CHAIN_PAGE)
            session.get(NSE_BASE, timeout=15)
            session.get(NSE_OPTION_CHAIN_PAGE, timeout=15)

            expiry = _get_nearest_expiry(session, symbol)

            url = OPTION_CHAIN_V3_API.format(quote(symbol), quote(expiry))
            resp = session.get(url, timeout=20)
            if resp.status_code == 401:  # cookies expired mid-flow: refresh once
                session = new_session(NSE_OPTION_CHAIN_PAGE)
                session.get(NSE_BASE, timeout=15)
                session.get(NSE_OPTION_CHAIN_PAGE, timeout=15)
                resp = session.get(url, timeout=20)
            resp.raise_for_status()
            records = resp.json().get("records", {}).get("data", [])

            rows = []
            for rec in records:
                rec_expiry = rec.get("expiryDates") or rec.get("expiryDate")
                if rec_expiry and rec_expiry != expiry:
                    continue

                ce, pe = rec.get("CE"), rec.get("PE")
                strike = rec.get("strikePrice")
                if strike is None and ce:
                    strike = ce.get("strikePrice")
                if strike is None and pe:
                    strike = pe.get("strikePrice")
                if strike is None or not ce or not pe:
                    continue

                rows.append((
                    float(strike),
                    float(ce.get("openInterest", 0) or 0),
                    float(ce.get("changeinOpenInterest", 0) or 0),
                    float(pe.get("openInterest", 0) or 0),
                    float(pe.get("changeinOpenInterest", 0) or 0),
                ))

            if rows:
                return rows
            last_err = RuntimeError(f"option-chain-v3 returned no strikes for expiry {expiry}")
        except Exception as exc:  # noqa: BLE001
            last_err = exc
            print(f"  option-chain attempt {attempt} failed: {exc}", file=sys.stderr)
        time.sleep(2 * attempt)

    raise RuntimeError(f"Could not fetch option chain after {retries} attempts: {last_err}")


def compute_rows(raw_rows, open_price, band):
    """Strike -> Option Score for strikes within open +/- band."""
    low, high = open_price - band, open_price + band
    rows = []
    for strike, call_oi, call_chng, put_oi, put_chng in raw_rows:
        if not (low <= strike <= high):
            continue

        call_prev = call_oi - call_chng
        put_prev = put_oi - put_chng

        call_pct = (call_chng / call_prev * 100) if call_prev else 0.0
        put_pct = (put_chng / put_prev * 100) if put_prev else 0.0

        rows.append({"strike": strike, "option_score": round_half_up((put_pct - call_pct) / 10)})
    rows.sort(key=lambda r: r["strike"])
    return rows


def option_csv_path(date_str: str) -> str:
    os.makedirs(OPTION_DIR, exist_ok=True)
    return os.path.join(OPTION_DIR, f"option_score_{date_str}.csv")


def load_score_grid(path):
    """Read a pivot CSV into {strike: {time_label: score}}."""
    grid = {}
    if not os.path.exists(path):
        return grid
    with open(path, newline="", encoding="utf-8") as f:
        reader = list(csv.reader(f))
    if not reader:
        return grid

    times = reader[0][1:]
    for row in reader[1:]:
        if not row or row[0].strip() == "":
            continue
        try:
            strike = float(row[0].strip())
        except ValueError:
            continue
        for j, tlabel in enumerate(times, start=1):
            if j >= len(row):
                continue
            v = row[j].strip()
            if v != "":
                try:
                    grid.setdefault(strike, {})[tlabel] = float(v)
                except ValueError:
                    pass
    return grid


def save_score_grid(path, grid):
    """Strike price in column 1, times across the header, rounded scores."""
    strikes = sorted(grid.keys())
    times = sorted({t for scores in grid.values() for t in scores})
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Strike price"] + times)
        for strike in strikes:
            row = [f"{strike:g}"]
            for t in times:
                v = grid[strike].get(t)
                row.append("" if v is None else round_half_up(v))
            writer.writerow(row)


def update_score_grid(rows, now: datetime) -> str:
    """Merge this cycle's scores into today's pivot CSV (a re-run of the same
    HH:MM slot overwrites it). Returns the CSV path."""
    path = option_csv_path(now.strftime("%Y-%m-%d"))
    time_label = now.strftime("%H:%M")

    grid = load_score_grid(path)
    for r in rows:
        grid.setdefault(r["strike"], {})[time_label] = round_half_up(r["option_score"])
    save_score_grid(path, grid)
    return path


def run_option_cycle() -> bool:
    """Fetch -> compute -> update today's pivot CSV. Returns True on success."""
    log_time = datetime.now().strftime("%H:%M:%S")
    try:
        open_price = fetch_index_open()
        raw_rows = fetch_option_chain_raw()
        rows = compute_rows(raw_rows, open_price, BAND)
    except Exception as exc:  # noqa: BLE001
        print(f"[{log_time}] [OPTION ERROR] Live NSE fetch failed: {exc} - will retry next cycle.")
        return False

    if not rows:
        print(f"[{log_time}] [OPTION ERROR] No strikes inside open +/- {BAND:g} - nothing logged.")
        return False

    path = update_score_grid(rows, datetime.now())
    print(f"[{log_time}] OPTION OK - open={open_price:.2f} strikes={len(rows)} -> {path}")
    return True


# ----------------------------------------------------------------------------
# 3) Google Sheets
# ----------------------------------------------------------------------------
def h1_sheet_values(path: str) -> list:
    """Today's H1 CSV as a table: header + rows, H1/AVG_H1 as numbers."""
    if not os.path.isfile(path):
        return []
    values = [H1_CSV_FIELDS[:]]
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.DictReader(f):
            try:
                values.append([row["time"], int(float(row["H1"])), int(float(row["AVG_H1"]))])
            except (KeyError, ValueError, TypeError):
                continue
    return values if len(values) > 1 else []


def option_sheet_values(path: str) -> list:
    """Today's option-score pivot CSV as a table: header + one row per strike."""
    grid = load_score_grid(path)
    if not grid:
        return []
    strikes = sorted(grid)
    times = sorted({t for scores in grid.values() for t in scores})
    values = [["Strike price"] + times]
    for strike in strikes:
        row = [int(strike) if strike == int(strike) else strike]
        for t in times:
            v = grid[strike].get(t)
            row.append("" if v is None else round_half_up(v))
        values.append(row)
    return values


class SheetsWriter:
    """Sends each table to a small Google Apps Script web app (see
    apps_script.gs) that lives inside your Google Sheet and writes the tab.
    No Google Cloud project / service account needed - just a URL + a token."""

    def __init__(self):
        self.url = os.environ.get("SHEET_WEBHOOK_URL", "").strip()
        self.token = os.environ.get("SHEET_TOKEN", "").strip()
        if not self.url:
            sys.exit("SHEET_WEBHOOK_URL is not set (the Web app URL from Apps Script -> Deploy).")
        if not self.token:
            sys.exit("SHEET_TOKEN is not set (the same password you put in apps_script.gs).")

    def write_table(self, title: str, values: list, freeze_col=False,
                    color_scores=False, text_first_col=False):
        if not values:
            return
        payload = {
            "token": self.token,
            "title": title,
            "values": values,
            "freeze_col": freeze_col,          # freeze strike column (option tab)
            "color_scores": color_scores,      # green >0 / red <0 (option tab)
            "text_first_col": text_first_col,  # keep timestamps as text (H1 tab)
        }
        last_err = None
        for attempt in range(1, 4):
            try:
                # Apps Script answers a POST with a redirect; requests follows it.
                resp = requests.post(self.url, data=json.dumps(payload),
                                     headers={"Content-Type": "text/plain"}, timeout=90)
                resp.raise_for_status()
                try:
                    result = resp.json()
                except ValueError:
                    raise RuntimeError("Web app did not return JSON - check the deployment is "
                                       "'Execute as: Me' and 'Who has access: Anyone'")
                if not result.get("ok"):
                    raise RuntimeError(f"Apps Script error: {result.get('error')}")
                return
            except Exception as exc:  # noqa: BLE001
                last_err = exc
                print(f"  sheet write '{title}' attempt {attempt} failed: {exc}", file=sys.stderr)
                time.sleep(3 * attempt)
        raise RuntimeError(f"Could not write tab '{title}': {last_err}")


H1_TAB = "H1"
OPTION_TAB = "OptionScore"
NIFTY50_TAB = "Nifty50_Data"

NIFTY50_HEADER = ["Symbol", "High", "Low", "LTP", "% Change"]


def nifty50_sheet_values(rows: list) -> list:
    """NIFTY 50 constituents from this run's NSE call, one row each (A-Z):
    Symbol, High, Low, LTP, % Change. % Change is NSE's own pChange; if NSE
    didn't send it, it is worked out from the previous close when available."""
    body = []
    for row in rows:
        symbol = _ci_get(row, "SYMBOL")
        if not symbol or str(symbol).strip().upper().startswith("NIFTY"):
            continue  # skip the aggregate index row
        ltp = _safe_float(_ci_get(row, "LASTPRICE", "LTP", "LAST"))
        high = _safe_float(_ci_get(row, "DAYHIGH", "HIGH"))
        low = _safe_float(_ci_get(row, "DAYLOW", "LOW"))

        pct = _safe_float(_ci_get(row, "PCHANGE", "PERCHANGE", "PERCENTCHANGE"))
        if pct is None and ltp is not None:
            prev = _safe_float(_ci_get(row, "PREVIOUSCLOSE", "PREVCLOSE"))
            if prev:
                pct = (ltp - prev) / prev * 100
        if pct is not None:
            pct = round(pct, 2)

        body.append([str(symbol).strip()] + ["" if v is None else v for v in (high, low, ltp, pct)])
    body.sort(key=lambda r: r[0])
    return [NIFTY50_HEADER] + body if body else []


def push_to_sheets(writer) -> bool:
    """Rewrite the two fixed tabs (H1, OptionScore) from TODAY's CSVs. Because
    each tab is rewritten in full from today's file, a new trading day simply
    replaces yesterday's data - the CSVs in data/ keep the full history."""
    today = datetime.now().strftime("%Y-%m-%d")
    log_time = datetime.now().strftime("%H:%M:%S")
    ok = True

    h1_values = h1_sheet_values(h1_csv_path())
    opt_values = option_sheet_values(option_csv_path(today))
    if opt_values:
        opt_values[0][0] = f"Strike price ({today})"  # shows which day the grid is for

    # Nifty50_Data is only refreshed when THIS run fetched fresh constituents;
    # if the NSE fetch failed, the last good snapshot in the sheet stays put.
    n50_values = []
    if LATEST_NIFTY50["rows"]:
        n50_values = nifty50_sheet_values(LATEST_NIFTY50["rows"])

    for title, values, kwargs in (
        (H1_TAB, h1_values, {"text_first_col": True}),
        (OPTION_TAB, opt_values, {"freeze_col": True, "color_scores": True}),
        (NIFTY50_TAB, n50_values, {"freeze_col": True}),
    ):
        if not values:
            print(f"[{log_time}] SHEET SKIP - '{title}' (no fresh data this cycle, left unchanged)")
            continue
        try:
            writer.write_table(title, values, **kwargs)
            print(f"[{log_time}] SHEET OK - '{title}' ({len(values) - 1} rows)")
        except Exception as exc:  # noqa: BLE001
            ok = False
            print(f"[{log_time}] [SHEET ERROR] {exc}")
    return ok


# ----------------------------------------------------------------------------
# Runner
# ----------------------------------------------------------------------------
def run_once(writer) -> bool:
    """One full cycle. H1 and Option Score are independent: if one NSE call
    fails the other still runs. Returns False only if a Google write failed."""
    run_h1_cycle()
    run_option_cycle()
    if writer is None:
        return True
    return push_to_sheets(writer)


def main():
    parser = argparse.ArgumentParser(description="NIFTY H1 + Option Score -> Google Sheets")
    parser.add_argument("--once", action="store_true",
                        help="Run one cycle and exit (GitHub Actions). Omit to loop every 5 min.")
    parser.add_argument("--no-sheets", action="store_true",
                        help="Write the CSVs only; skip Google Sheets.")
    args = parser.parse_args()

    writer = None if args.no_sheets else SheetsWriter()

    if args.once:
        sys.exit(0 if run_once(writer) else 1)

    print("NIFTY H1 + Option Score -> Google Sheets")
    print(f"  CSV folders : {H1_AVG_DIR} | {OPTION_DIR}")
    print(f"  Refresh     : every {REFRESH_SECONDS}s | Press Ctrl+C to stop.\n")
    while True:
        cycle_start = time.monotonic()
        try:
            run_once(writer)
        except Exception as exc:  # noqa: BLE001
            print(f"[ERROR] Unexpected failure in update cycle: {exc}")
        remaining = REFRESH_SECONDS - (time.monotonic() - cycle_start)
        if remaining > 0:
            time.sleep(remaining)
        else:
            print("[WARN] Cycle overran the refresh interval - running next cycle immediately.")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped.")
        sys.exit(0)
