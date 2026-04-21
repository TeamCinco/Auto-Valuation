"""
calculations.py
Processes raw data for each ticker and computes valuation ratios.
"""

from __future__ import annotations

import os
import logging
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from data_loader import load_market_cap, load_cash_flow

logger = logging.getLogger(__name__)


@dataclass
class QuarterData:
    """Single quarter's valuation data for one ticker."""
    date: pd.Timestamp
    net_income: float | None = None
    fcf: float | None = None
    avg_mcap: float | None = None

    @property
    def pe(self) -> float | None:
        if self.avg_mcap is None or self.net_income is None or self.net_income == 0:
            return None
        return round(self.avg_mcap / self.net_income, 2)

    @property
    def pfcf(self) -> float | None:
        if self.avg_mcap is None or self.fcf is None or self.fcf == 0:
            return None
        return round(self.avg_mcap / self.fcf, 2)


@dataclass
class TickerResult:
    """Full quarterly valuation history for one ticker."""
    ticker: str
    quarters: list[QuarterData] = field(default_factory=list)

    # Convenience accessors for the excel writer
    @property
    def dates(self):       return [q.date for q in self.quarters]
    @property
    def net_income(self):  return [q.net_income for q in self.quarters]
    @property
    def fcf(self):         return [q.fcf for q in self.quarters]
    @property
    def avg_mcap(self):    return [q.avg_mcap for q in self.quarters]
    @property
    def avg_pe(self):      return [q.pe for q in self.quarters]
    @property
    def avg_pfcf(self):    return [q.pfcf for q in self.quarters]


def discover_tickers(input_dir: str) -> list[tuple[str, str, str]]:
    """Scan flat folder for TICKER-market-cap.* and TICKER-cash-flow-statement-ttm.* pairs.

    Returns list of (ticker, mcap_path, cf_path) tuples.
    """
    VALID_EXT = {'.csv', '.xlsx'}
    MCAP_MARKER = '-market-cap.'
    CF_MARKER = '-cash-flow-statement-ttm.'

    mcap_files: dict[str, str] = {}
    cf_files: dict[str, str] = {}

    for f in os.listdir(input_dir):
        fl = f.lower()
        if not any(fl.endswith(ext) for ext in VALID_EXT):
            continue

        if MCAP_MARKER in fl:
            ticker = f[:fl.index(MCAP_MARKER)].upper()
            mcap_files[ticker] = os.path.join(input_dir, f)
        elif CF_MARKER in fl:
            ticker = f[:fl.index(CF_MARKER)].upper()
            cf_files[ticker] = os.path.join(input_dir, f)

    paired = sorted(mcap_files.keys() & cf_files.keys())

    for orphan in sorted(mcap_files.keys() - cf_files.keys()):
        logger.warning("Market cap found but no cash flow for: %s", orphan)
    for orphan in sorted(cf_files.keys() - mcap_files.keys()):
        logger.warning("Cash flow found but no market cap for: %s", orphan)

    return [(t, mcap_files[t], cf_files[t]) for t in paired]


def _log_coverage(ticker: str, mcap_dates: list, cf_dates: list):
    """Log date range diagnostics for a ticker."""
    def _fmt(dates):
        if not dates:
            return "NONE"
        return f"{len(dates)} quarters ({dates[0]:%Y-%m-%d} -> {dates[-1]:%Y-%m-%d})"

    logger.info("  %s market cap:  %s", ticker, _fmt(mcap_dates))
    logger.info("  %s cash flow:   %s", ticker, _fmt(cf_dates))

    missing = set(cf_dates) - set(mcap_dates)
    if missing:
        logger.warning("  %s: %d cash flow quarters have NO market cap data", ticker, len(missing))


def _safe_round(val, decimals=0) -> float | None:
    """Round a value, returning None if it's None or NaN."""
    if val is None:
        return None
    try:
        if np.isnan(val):
            return None
    except (TypeError, ValueError):
        pass
    return round(float(val), decimals)


def process_ticker(ticker: str, mcap_file: str, cf_file: str) -> TickerResult:
    """Load data for one ticker and compute quarterly valuation ratios."""
    quarterly_mcap = load_market_cap(mcap_file)
    cf_data = load_cash_flow(cf_file)

    _log_coverage(ticker, sorted(quarterly_mcap.index), sorted(cf_data.keys()))

    all_quarters = sorted(set(quarterly_mcap.index) | set(cf_data.keys()))
    result = TickerResult(ticker=ticker)

    for q in all_quarters:
        cf = cf_data.get(q, {})
        ni = cf.get('net_income')
        fcf = cf.get('fcf')

        if ni is None and fcf is None:
            continue

        result.quarters.append(QuarterData(
            date=q,
            net_income=ni,
            fcf=fcf,
            avg_mcap=_safe_round(quarterly_mcap.get(q), 0),
        ))

    return result