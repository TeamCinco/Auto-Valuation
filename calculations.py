"""
calculations.py
Processes raw data for each ticker and computes valuation ratios.
"""

import os
import numpy as np
from data_loader import load_market_cap, load_cash_flow


def discover_tickers(input_dir):
    """Scan flat folder for TICKER-market-cap.* and TICKER-cash-flow-statement-ttm.* pairs."""
    files = [f for f in os.listdir(input_dir) if not f.startswith('.')]
    mcap_files = {}
    cf_files = {}

    for f in files:
        fl = f.lower()
        if not (fl.endswith('.csv') or fl.endswith('.xlsx')):
            continue
        if '-market-cap.' in fl:
            ticker = f[:fl.index('-market-cap.')].upper()
            mcap_files[ticker] = os.path.join(input_dir, f)
        elif '-cash-flow-statement-ttm.' in fl:
            ticker = f[:fl.index('-cash-flow-statement-ttm.')].upper()
            cf_files[ticker] = os.path.join(input_dir, f)

    paired = sorted(set(mcap_files.keys()) & set(cf_files.keys()))
    mcap_only = set(mcap_files.keys()) - set(cf_files.keys())
    cf_only = set(cf_files.keys()) - set(mcap_files.keys())

    if mcap_only:
        print(f"WARNING: Market cap found but no cash flow for: {', '.join(sorted(mcap_only))}")
    if cf_only:
        print(f"WARNING: Cash flow found but no market cap for: {', '.join(sorted(cf_only))}")

    return [(t, mcap_files[t], cf_files[t]) for t in paired]


def process_ticker(ticker, mcap_file, cf_file):
    """Process one ticker -> dict with quarterly data and computed ratios."""
    print(f"  Market cap:  {os.path.basename(mcap_file)}")
    quarterly_mcap = load_market_cap(mcap_file)

    print(f"  Cash flow:   {os.path.basename(cf_file)}")
    cf_data = load_cash_flow(cf_file)

    mcap_dates = sorted(quarterly_mcap.index)
    cf_dates = sorted(cf_data.keys())
    print(f"  Market cap:  {len(mcap_dates)} quarters ({mcap_dates[0].strftime('%Y-%m-%d') if mcap_dates else 'NONE'} -> {mcap_dates[-1].strftime('%Y-%m-%d') if mcap_dates else 'NONE'})")
    print(f"  Cash flow:   {len(cf_dates)} quarters ({cf_dates[0].strftime('%Y-%m-%d') if cf_dates else 'NONE'} -> {cf_dates[-1].strftime('%Y-%m-%d') if cf_dates else 'NONE'})")

    mcap_missing = set(cf_dates) - set(mcap_dates)
    if mcap_missing:
        print(f"  WARNING: {len(mcap_missing)} cash flow quarters have NO market cap data")

    all_quarters = sorted(set(list(quarterly_mcap.index) + list(cf_data.keys())))

    result = {
        'ticker': ticker,
        'quarters': [],
        'net_income': [],
        'fcf': [],
        'avg_mcap': [],
        'avg_pe': [],
        'avg_pfcf': []
    }

    for q in all_quarters:
        ni = cf_data.get(q, {}).get('net_income')
        fcf = cf_data.get(q, {}).get('fcf')
        mcap = quarterly_mcap.get(q)

        if ni is None and fcf is None:
            continue

        result['quarters'].append(q)
        result['net_income'].append(ni)
        result['fcf'].append(fcf)
        result['avg_mcap'].append(round(mcap, 0) if mcap is not None and not np.isnan(mcap) else None)

        pe = None
        pfcf = None
        if mcap is not None and not np.isnan(mcap):
            if ni is not None and ni != 0:
                pe = round(float(mcap) / float(ni), 2)
            if fcf is not None and fcf != 0:
                pfcf = round(float(mcap) / float(fcf), 2)

        result['avg_pe'].append(pe)
        result['avg_pfcf'].append(pfcf)

    return result