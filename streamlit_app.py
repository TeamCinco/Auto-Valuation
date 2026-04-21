"""
streamlit_app.py
Web UI for the Valuation Sheet Builder.

Workflow:
  1. User downloads files from stockanalysis.com into a local folder.
  2. User enters the folder path here.
  3. App scans the folder, shows what was discovered, processes everything
     in memory, and offers the styled xlsx as a download.

No files are persisted server-side. The Excel workbook is written to a
BytesIO buffer that lives only for the duration of the session.
"""

from __future__ import annotations  # Enables PEP 604 union syntax on Python 3.9

import io
import os
import logging
from pathlib import Path

import streamlit as st

# Project modules (must be in the same directory or on PYTHONPATH)
from calculations import discover_tickers, process_ticker
from excel_writer import build_workbook


# ── Page config ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="Valuation Sheet Builder",
    page_icon="📊",
    layout="centered",
)

# Quiet third-party loggers; capture our own diagnostics
logging.basicConfig(level=logging.INFO, format="%(message)s")


# ── Helpers ──────────────────────────────────────────────────────────
def _validate_folder(path_str: str) -> tuple[Path | None, str | None]:
    """Return (resolved_path, error_message). One of them is None."""
    if not path_str or not path_str.strip():
        return None, "Enter a folder path above to get started."

    path = Path(path_str.strip()).expanduser()

    if not path.exists():
        return None, f"Path does not exist: `{path}`"
    if not path.is_dir():
        return None, f"Path is not a folder: `{path}`"

    return path, None


def _scan_folder(folder: Path) -> dict:
    """Categorize files in the folder. Returns a summary dict."""
    files = [f for f in os.listdir(folder) if not f.startswith(".")]

    mcap, cf, other = [], [], []
    for f in files:
        fl = f.lower()
        if not (fl.endswith(".csv") or fl.endswith(".xlsx")):
            other.append(f)
        elif "-market-cap." in fl:
            mcap.append(f)
        elif "-cash-flow-statement-ttm." in fl:
            cf.append(f)
        else:
            other.append(f)

    return {
        "total": len(files),
        "market_cap": sorted(mcap),
        "cash_flow": sorted(cf),
        "other": sorted(other),
    }


def _generate_workbook(folder: Path) -> tuple[bytes | None, list[str], str | None]:
    """Run the full pipeline. Returns (xlsx_bytes, tickers_processed, error)."""
    try:
        pairs = discover_tickers(str(folder))
    except Exception as e:
        return None, [], f"Failed to scan folder: {e}"

    if not pairs:
        return None, [], (
            "No valid ticker pairs found. Each ticker needs BOTH a "
            "`TICKER-market-cap.csv/xlsx` and a "
            "`TICKER-cash-flow-statement-ttm.csv/xlsx` file."
        )

    tickers_data = []
    skipped = []
    for ticker, mcap_file, cf_file in pairs:
        try:
            result = process_ticker(ticker, mcap_file, cf_file)
            if result and result.dates:
                tickers_data.append(result)
            else:
                skipped.append(ticker)
        except Exception as e:
            skipped.append(f"{ticker} ({e})")

    if not tickers_data:
        return None, [], "All tickers failed to process. Check your input files."

    # Write xlsx to an in-memory buffer — no disk I/O, no cleanup needed
    buffer = io.BytesIO()
    try:
        build_workbook(tickers_data, buffer)
        buffer.seek(0)
    except Exception as e:
        return None, [], f"Workbook generation failed: {e}"

    processed = [td.ticker for td in tickers_data]
    if skipped:
        processed.append(f"(skipped: {', '.join(skipped)})")

    return buffer.getvalue(), processed, None


# ── UI ───────────────────────────────────────────────────────────────
st.title("📊 Valuation Sheet Builder")
st.caption(
    "Build a styled valuation comp sheet from stockanalysis.com data. "
    "Files stay on your machine — this app only reads them."
)

with st.expander("How to use", expanded=False):
    st.markdown(
        """
        **1. Download the input data from [stockanalysis.com](https://stockanalysis.com)**

        For every ticker in your peer group, save two files into one folder:

        - `TICKER-market-cap.csv` (or `.xlsx`)
        - `TICKER-cash-flow-statement-ttm.csv` (or `.xlsx`)

        Example:
        ```
        /Users/you/Desktop/comps/
          META-market-cap.csv
          META-cash-flow-statement-ttm.xlsx
          GOOG-market-cap.csv
          GOOG-cash-flow-statement-ttm.csv
        ```

        **2. Paste the folder path below and click Generate.**

        The parser tolerates raw stockanalysis.com output — abbreviated numbers
        (`11.98B`, `847.5M`), parenthesized negatives, mixed date formats, and
        the `% Change` column are all handled automatically. No cleanup needed.
        """
    )

# ── Folder input ─────────────────────────────────────────────────────
default_path = st.session_state.get("folder_path", "")
folder_str = st.text_input(
    "Input folder path",
    value=default_path,
    placeholder="/Users/you/Desktop/comps",
    help="Absolute path to the folder containing your downloaded CSV/XLSX files.",
)
st.session_state["folder_path"] = folder_str

folder, err = _validate_folder(folder_str)

if err:
    st.info(err)
    st.stop()

# ── Folder preview ───────────────────────────────────────────────────
summary = _scan_folder(folder)

col1, col2, col3 = st.columns(3)
col1.metric("Market cap files", len(summary["market_cap"]))
col2.metric("Cash flow files", len(summary["cash_flow"]))
col3.metric("Other / unmatched", len(summary["other"]))

# Pair-matching preview
mcap_tickers = {
    f[: f.lower().index("-market-cap.")].upper() for f in summary["market_cap"]
}
cf_tickers = {
    f[: f.lower().index("-cash-flow-statement-ttm.")].upper() for f in summary["cash_flow"]
}
paired = sorted(mcap_tickers & cf_tickers)
orphans_mcap = sorted(mcap_tickers - cf_tickers)
orphans_cf = sorted(cf_tickers - mcap_tickers)

if paired:
    st.success(f"**{len(paired)} ticker(s) ready:** {', '.join(paired)}")
else:
    st.warning("No complete ticker pairs found in this folder.")

if orphans_mcap:
    st.warning(
        f"Market cap file with no matching cash flow: "
        f"{', '.join(orphans_mcap)}"
    )
if orphans_cf:
    st.warning(
        f"Cash flow file with no matching market cap: "
        f"{', '.join(orphans_cf)}"
    )

# ── Generate button ──────────────────────────────────────────────────
st.divider()

output_filename = st.text_input(
    "Output filename",
    value="valuation_output.xlsx",
    help="Filename for the download. Must end in `.xlsx`.",
)
if not output_filename.lower().endswith(".xlsx"):
    output_filename = output_filename + ".xlsx"

if st.button("Generate workbook", type="primary", disabled=len(paired) == 0):
    with st.spinner(f"Processing {len(paired)} ticker(s)..."):
        xlsx_bytes, processed, err = _generate_workbook(folder)

    if err:
        st.error(err)
    else:
        st.success(f"Done. Processed: {', '.join(processed)}")
        st.download_button(
            label="⬇ Download workbook",
            data=xlsx_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.caption(
            "The file is generated in memory and streamed directly to your "
            "browser. Nothing is saved on this server."
        )

# ── Footer ───────────────────────────────────────────────────────────
st.divider()
st.caption(
    "Your input files never leave your machine. This app reads the folder "
    "path you provide, processes the data in memory, and discards everything "
    "when you close the tab."
)