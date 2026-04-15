"""
excel_writer.py
Builds the output valuation Excel workbook.
Layout and styling matches the original valuation sheet exactly.
"""

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from styles import (
    FILL_TICKER_ROW, FILL_NONE,
    FONT_SECTION_TITLE, FONT_HEADER, FONT_TICKER_SUMMARY, FONT_TICKER_DETAIL,
    FONT_DATE, FONT_ROW_LABEL, FONT_DATA, FONT_INSTRUCTION_BOLD, FONT_INSTRUCTION,
    ALIGN_CENTER, ALIGN_RIGHT, ALIGN_LEFT,
    BORDER_BOX, BORDER_NONE,
    FMT_INT, FMT_DEC2, FMT_PRICE,
    COL_WIDTH_LABEL, COL_WIDTH_DATA, COL_WIDTH_SUMMARY, COL_WIDTH_TICKER, COL_WIDTH_GAP,
    apply_cell,
)


def _write_summary_tables(ws, tickers):
    """Write PE Valuation and P FCF Valuation summary tables.
    Only ticker names filled; data columns left empty for manual entry.

    Layout (1-indexed columns):
      PE Valuation:    B(2)-G(7)     → Ticker, Price, EPS, PE, 5Y, 10Y
      Gap:             H(8)-J(10)
      P FCF Valuation: K(11)-P(16)   → Ticker, Market cap, TTM FCF, P FCF, 5Y, 10Y
    """
    # ── PE Valuation title (row 1, col B) ──
    apply_cell(ws.cell(1, 2), 'PE Valuation', FONT_SECTION_TITLE)

    # ── PE headers (row 2) ──
    pe_headers = ['Ticker', 'Price', 'EPS', 'PE', '5Y', '10Y']
    for j, h in enumerate(pe_headers):
        apply_cell(ws.cell(2, 2 + j), h, FONT_HEADER, border=BORDER_BOX)

    # ── PE ticker rows ──
    for i, td in enumerate(tickers):
        row = 3 + i
        apply_cell(ws.cell(row, 2), td.ticker, FONT_TICKER_SUMMARY, border=BORDER_BOX)
        for c in range(3, 8):
            apply_cell(ws.cell(row, c), border=BORDER_BOX)

    # ── P FCF Valuation title (row 1, col K) ──
    apply_cell(ws.cell(1, 11), 'P FCF Valuation', FONT_SECTION_TITLE)

    # ── P FCF headers (row 2) ──
    pfcf_headers = ['Ticker', 'Market cap', 'TTM FCF', 'P FCF', '5Y', '10Y']
    for j, h in enumerate(pfcf_headers):
        apply_cell(ws.cell(2, 11 + j), h, FONT_HEADER, border=BORDER_BOX)

    # ── P FCF ticker rows ──
    for i, td in enumerate(tickers):
        row = 3 + i
        apply_cell(ws.cell(row, 11), td.ticker, FONT_TICKER_SUMMARY, border=BORDER_BOX)
        for c in range(12, 17):
            apply_cell(ws.cell(row, c), border=BORDER_BOX)


    return 3 + len(tickers)


def _write_detail_block(ws, td, start_row):
    """Write one ticker's detail block: yellow ticker row + 5 data rows + blank separator.

    Layout per block:
      Row 0: TICKER | 2016-03-31 | 2016-06-30 | ...    (yellow fill, bold)
      Row 1: Net Income TTM | val | val | ...
      Row 2: FCF TTM | val | val | ...
      Row 3: Average market cap | val | val | ...
      Row 4: Average PE | val | val | ...
      Row 5: Average P FCF | val | val | ...
      Row 6: (blank separator)
    """
    quarters = td.dates
    n_q = len(quarters)
    r = start_row

    # ── Ticker header row (yellow background) ──
    apply_cell(ws.cell(r, 1), td.ticker, FONT_TICKER_DETAIL, FILL_TICKER_ROW, ALIGN_LEFT)
    for j, q in enumerate(quarters):
        apply_cell(ws.cell(r, 2 + j), q.strftime('%Y-%m-%d'), FONT_DATE, FILL_TICKER_ROW, ALIGN_CENTER)

    # ── Data rows (no fill, no borders on data cells) ──
    row_defs = [
        ('Net Income TTM',      td.net_income, FMT_INT),
        ('FCF TTM',             td.fcf,        FMT_INT),
        ('Average market cap',  td.avg_mcap,   FMT_INT),
        ('Average PE',          td.avg_pe,     FMT_DEC2),
        ('Average P FCF',       td.avg_pfcf,   FMT_DEC2),
    ]

    for k, (label, values, fmt) in enumerate(row_defs):
        dr = r + 1 + k
        apply_cell(ws.cell(dr, 1), label, FONT_ROW_LABEL, alignment=ALIGN_LEFT)

        for j, val in enumerate(values):
            apply_cell(ws.cell(dr, 2 + j), val, FONT_DATA, alignment=ALIGN_RIGHT, number_format=fmt)

    return r + 7  # ticker row + 5 data rows + 1 blank


def build_workbook(tickers_data, output_path):
    """Build the complete valuation workbook and save it."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Valuation'

    # ── Summary tables ──
    last_summary_row = _write_summary_tables(ws, tickers_data)

    # ── Detail blocks (start after a blank row) ──
    current_row = last_summary_row + 2

    for td in tickers_data:
        if not td.quarters:
            continue
        current_row = _write_detail_block(ws, td, current_row)

    # ── Column widths ──
    ws.column_dimensions['A'].width = COL_WIDTH_LABEL

    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH_SUMMARY
    for c in range(8, 11):
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH_GAP
    for c in range(11, 17):
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH_SUMMARY

    max_data_col = 2
    for td in tickers_data:
        max_data_col = max(max_data_col, 2 + len(td.quarters))
    for c in range(2, max_data_col + 1):
        cur = ws.column_dimensions[get_column_letter(c)].width or 0
        if cur < COL_WIDTH_DATA:
            ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH_DATA

    ws.column_dimensions[get_column_letter(43)].width = 18
    ws.column_dimensions[get_column_letter(44)].width = 25

    ws.freeze_panes = 'B1'

    wb.save(output_path)
    print(f"\nSaved: {output_path}")