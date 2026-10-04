#!/usr/bin/env python3
"""
Fetch the last N fiscal years of dividend history from ShareSansar
for a list of mutual-fund symbols.

Usage:
    pip install requests pandas lxml openpyxl
    python sharesansar_dividends.py                 # your 50 companies (default)
    python sharesansar_dividends.py --group funds   # the mutual funds
    python sharesansar_dividends.py --group all
    python sharesansar_dividends.py --symbols SLCF SAGF
    python sharesansar_dividends.py --years 3 --debug

Output: dividend_history.csv and dividend_history.xlsx
"""
import argparse
import io
import json
import re
import sys
import time

import pandas as pd
import requests

BASE = "https://www.sharesansar.com"

SYMBOLS = {
    "H8020": "Himalayan 80-20", "SFEF": "Sunrise Focused Equity Fund",
    "C30MF": "Citizens Super 30 Mutual Fund", "SAGF": "Sanima Growth Fund",
    "GIBF1": "Global IME Balanced Fund - 1", "RMF2": "RBB Mutual Fund 2",
    "GSY": "Garima Samriddhi Yojana", "MNMF1": "Muktinath Mutual Fund 1",
    "NSIF2": "NMB Sulav Investment Fund - II", "SFMF": "Sunrise First Mutual Fund",
    "PRSF": "Prabhu Smart Fund", "SIGS2": "Siddhartha Investment Growth Scheme - 2",
    "SIGS3": "Siddhartha Investment Growth Scheme 3", "NMB50": "NMB 50",
    "SLCF": "Sanima Large Cap Fund", "SFF": "Sanima Flexi Fund",
    "LVF2": "Laxmi Value Fund 2", "LUK": "Laxmi Unnati Kosh",
    "RSY": "Reliable Samriddhi Yojana", "PSF": "Prabhu Select Fund",
    "NBF2": "NABIL BALANCED FUND-2", "SLK": "Shubha Laxmi Kosh",
    "SEF": "Siddhartha Equity Fund", "CMF2": "CITIZENS MUTUAL FUND 2",
    "NBF3": "Nabil Balance Fund III", "MBLEF": "MBL Equity Fund",
    "NMBHF2": "NMB Hybrid Fund L-2", "CSBY": "Citizens Sadabahar Yojana",
    "NICGF2": "Nic Asia Growth Fund 2", "KDBY": "Kumari Dhanabriddhi Yojana",
    "NIBLSF": "NIBL Sahabhagita Fund", "NMBSBF": "NMB Saral Bachat Fund - E",
    "NICFC": "NIC Asia Flexi Cap Fund", "NI31": "NI 31",
    "NICBF": "NIC Asia Balanced Fund", "NFCF": "Nabil Flexi Cap Fund",
    "SSIS": "Siddhartha Systematic Investment Scheme", "KEF": "Kumari Equity Fund",
    "SBCF": "Sunrise Bluechip Fund", "KSLY": "Kumari Sunaulo Lagani Yojana",
    "KSY": "Kumari Sabal Yojana", "GBIMESY2": "Global IME Samunnat Yojana - II",
    "NIBLGF": "NIBL Growth Fund", "NADDF": "NIC Asia Dynamic Debt Fund",
    "PSIS": "Prabhu Systematic Investment Scheme", "NIBSF2": "NIBL Samriddhi Fund - II",
    "NICAELIS": "NIC Asia Equity Linked Investment Scheme", "RBBF40": "RBB FOCUS 40",
    "MSIP": "Machhapuchchhre SIP Yojana", "NIBLSTF": "NIBL STABLE FUND",
    "CSY": "Citizens Santulit Yojana", "NSY": "Nepal Life Samriddhi Lagani Yojana",
    "NICSF": "NIC Asia Select - 30 (Index Fund)", "MMF1": "Mega Mutual Fund 1",
    "RMF1": "RBB Mutual Fund 1", "GSYA": "Garima Subarna Yojana",
    "HLICF": "HLI Large Cap Fund",
}

COMPANIES = """ACLBSL ALBSL ANLB AVYAN CBBL CYCL DDBL DLBS FMDBL FOWAD GBLBS GILB GLBSL GMFBS
HLBSL ILBS JBLB JSLBB KMCDB LLBS MATRI MERO MLBBL MLBS MLBSL MSLB NADEP NESDO NICLBSL
NMBMF NMFBS NMLBBL NUBL RSDC SHLB SKBBL SLBBL SLBSL SMATA SMB SMFBS SMPDA SWASTIK
SWBBL SWMF ULBSL UNLB USLB VLBS""".split()

GROUPS = {
    "companies": {s: "" for s in COMPANIES},   # the microfinance/other share list
    "funds": SYMBOLS,                          # the mutual-fund list
}
GROUPS["all"] = {**GROUPS["funds"], **GROUPS["companies"]}

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept-Language": "en-US,en;q=0.9",
}

# Endpoints to try, in order, for the "Dividend History" tab data.
DIVIDEND_ENDPOINTS = ["/company-dividend", "/company-dividend-history"]


def get_token_and_company_id(session, symbol, debug=False):
    r = session.get(f"{BASE}/company/{symbol.lower()}", headers=HEADERS, timeout=30)
    r.raise_for_status()
    html = r.text
    if debug:
        open(f"debug_{symbol}_page.html", "w", encoding="utf-8").write(html)

    token = None
    for pat in (r'<meta[^>]+name=["\']_token["\'][^>]+content=["\']([^"\']+)',
                r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+name=["\']_token["\']',
                r'name=["\']_token["\'][^>]*value=["\']([^"\']+)'):
        m = re.search(pat, html)
        if m:
            token = m.group(1)
            break

    company_id = None
    for pat in (r'id=["\']companyid["\'][^>]*>\s*(\d+)',
                r'id=["\']company_id["\'][^>]*value=["\'](\d+)',
                r'company\s*[:=]\s*["\']?(\d+)["\']?',
                r'(\d+)\s*</[^>]+>\s*<[^>]+>\s*' + re.escape(symbol.upper())):
        m = re.search(pat, html, flags=re.I)
        if m:
            company_id = m.group(1)
            break
    return token, company_id, html


def to_dataframe(resp):
    """Turn an endpoint response (JSON or HTML) into a DataFrame."""
    text = resp.text.strip()
    # JSON (DataTables style: {"data": [...]}) or plain list
    try:
        js = json.loads(text)
        rows = js.get("data", js.get("aaData")) if isinstance(js, dict) else js
        if rows:
            df = pd.DataFrame(rows)
            # cells may contain HTML fragments; strip tags
            df = df.map(lambda v: re.sub(r"<[^>]+>", "", v).strip() if isinstance(v, str) else v)
            return df
        return pd.DataFrame()
    except ValueError:
        pass
    # HTML table
    try:
        tables = pd.read_html(io.StringIO(text))
        return tables[0] if tables else pd.DataFrame()
    except ValueError:
        return pd.DataFrame()


def fetch_dividends(session, symbol, debug=False):
    token, company_id, html = get_token_and_company_id(session, symbol, debug)
    if not company_id:
        raise RuntimeError("could not find company id on page")

    ajax_headers = dict(HEADERS)
    ajax_headers.update({
        "X-Requested-With": "XMLHttpRequest",
        "Referer": f"{BASE}/company/{symbol.lower()}",
    })
    if token:
        ajax_headers["X-CSRF-TOKEN"] = token

    payload = {
        "draw": 1, "start": 0, "length": 100,
        "company": company_id, "_token": token or "",
    }

    last_err = None
    for ep in DIVIDEND_ENDPOINTS:
        try:
            resp = session.post(BASE + ep, data=payload, headers=ajax_headers, timeout=30)
            if debug:
                open(f"debug_{symbol}{ep.replace('/', '_')}.txt", "w", encoding="utf-8").write(
                    f"status={resp.status_code}\n{resp.text[:5000]}")
            if resp.status_code != 200:
                last_err = f"{ep} -> HTTP {resp.status_code}"
                continue
            df = to_dataframe(resp)
            if not df.empty:
                return df
            last_err = f"{ep} -> empty"
        except requests.RequestException as e:
            last_err = f"{ep} -> {e}"
    raise RuntimeError(last_err or "no data")


def normalise(df):
    """Rename columns to a standard set and add a numeric fiscal-year sort key."""
    rename = {}
    for c in df.columns:
        lc = str(c).lower()
        if "bonus" in lc and "list" not in lc:
            rename[c] = "Bonus Share (%)"
        elif "cash" in lc:
            rename[c] = "Cash Dividend (%)"
        elif "total" in lc:
            rename[c] = "Total Dividend (%)"
        elif "announce" in lc:
            rename[c] = "Announcement Date"
        elif "book" in lc:
            rename[c] = "Book Close Date"
        elif "distribution" in lc:
            rename[c] = "Distribution Date"
        elif "listing" in lc:
            rename[c] = "Bonus Listing Date"
        elif "year" in lc or lc in ("fy", "fiscal_year"):
            rename[c] = "Fiscal Year"
    df = df.rename(columns=rename)
    # drop serial-number column if present
    df = df[[c for c in df.columns if str(c).lower() not in ("s.n.", "sn", "#", "0")
             or c in rename.values()]]
    return df


def fiscal_key(v):
    m = re.match(r"\s*(\d{4})", str(v))
    return int(m.group(1)) if m else -1


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--symbols", nargs="*", help="specific symbols (overrides --group)")
    ap.add_argument("--group", choices=list(GROUPS), default="companies",
                    help="which list to run (default: companies)")
    ap.add_argument("--years", type=int, default=5, help="fiscal years to keep")
    ap.add_argument("--delay", type=float, default=1.5, help="seconds between requests")
    ap.add_argument("--debug", action="store_true", help="save raw responses to disk")
    args = ap.parse_args()

    names = {**GROUPS["all"]}
    symbols = [x.upper() for x in args.symbols] if args.symbols else list(GROUPS[args.group])
    session = requests.Session()
    frames, failures = [], []

    for i, sym in enumerate(symbols, 1):
        print(f"[{i}/{len(symbols)}] {sym} ...", end=" ", flush=True)
        try:
            df = normalise(fetch_dividends(session, sym, args.debug))
            if "Fiscal Year" in df.columns:
                df["_k"] = df["Fiscal Year"].map(fiscal_key)
                df = df.sort_values("_k", ascending=False).head(args.years).drop(columns="_k")
            else:
                df = df.head(args.years)
            df.insert(0, "Name", names.get(sym, ""))
            df.insert(0, "Symbol", sym)
            frames.append(df)
            print(f"{len(df)} rows")
        except Exception as e:
            failures.append((sym, str(e)))
            print(f"FAILED ({e})")
        time.sleep(args.delay)

    if frames:
        out = pd.concat(frames, ignore_index=True)
        for col in ("Bonus Share (%)", "Cash Dividend (%)", "Total Dividend (%)"):
            if col in out.columns:
                out[col] = pd.to_numeric(
                    out[col].astype(str).str.replace("%", "").str.strip(), errors="coerce")
        front = ["Symbol", "Name", "Fiscal Year", "Bonus Share (%)",
                 "Cash Dividend (%)", "Total Dividend (%)"]
        cols = [c for c in front if c in out.columns] + \
               [c for c in out.columns if c not in front]
        out = out[cols]
        out.to_csv("dividend_history.csv", index=False, encoding="utf-8-sig")
        with pd.ExcelWriter("dividend_history.xlsx") as xw:
            out.to_excel(xw, sheet_name="All", index=False)
            if "Fiscal Year" in out.columns:
                for col, sheet in (("Cash Dividend (%)", "Cash_Pivot"),
                                   ("Bonus Share (%)", "Bonus_Pivot")):
                    if col in out.columns:
                        pv = out.pivot_table(index="Symbol", columns="Fiscal Year",
                                             values=col, aggfunc="first")
                        pv = pv[sorted(pv.columns, key=fiscal_key, reverse=True)]
                        pv.to_excel(xw, sheet_name=sheet)
        print(f"\nSaved {len(out)} rows to dividend_history.csv / .xlsx")
        print(out.head(15).to_string(index=False))
    if failures:
        print("\nFailed symbols:")
        for s, e in failures:
            print(f"  {s}: {e}")
        print("Re-run with --debug --symbols <SYMBOL> and check the debug_* files.")
    sys.exit(0 if frames else 1)


if __name__ == "__main__":
    main()
