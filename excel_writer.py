"""
excel_writer.py
Builds the output valuation Excel workbook.
Layout matches the original valuation.csv structure exactly.
"""

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from styles import (
    FILL_HEADER, FILL_TICKER_ROW, FILL_DATA_ALT, FILL_NONE,
    FONT_SECTION_TITLE, FONT_HEADER, FONT_TICKER_SUMMARY, FONT_TICKER_DETAIL,
    FONT_DATE, FONT_ROW_LABEL, FONT_DATA, FONT_INSTRUCTION_BOLD, FONT_INSTRUCTION,
    ALIGN_CENTER, ALIGN_RIGHT, ALIGN_LEFT,
    BORDER_ALL, BORDER_BOTTOM,
    FMT_INT, FMT_DEC2, FMT_PRICE,
    COL_WIDTH_LABEL, COL_WIDTH_DATA, COL_WIDTH_TICKER, COL_WIDTH_GAP,
    apply_cell,
)


def _write_summary_tables(ws, tickers):
    """Write the PE Valuation and P FCF Valuation summary tables (rows 1-N).
    Only ticker names are filled; data columns left empty for manual entry.
    """
    # ── PE Valuation (cols B-G, i.e. 2-7) ──
    apply_cell(ws.cell(1, 2), 'PE Valuation', FONT_SECTION_TITLE, FILL_HEADER, ALIGN_CENTER)
    for c in range(3, 8):
        apply_cell(ws.cell(1, c), fill=FILL_HEADER)

    pe_headers = ['Ticker', 'Price', 'EPS', 'PE', '5Y', '10Y']
    for j, h in enumerate(pe_headers):
        apply_cell(ws.cell(2, 2 + j), h, FONT_HEADER, FILL_HEADER, ALIGN_CENTER, BORDER_ALL)

    for i, td in enumerate(tickers):
        row = 3 + i
        apply_cell(ws.cell(row, 2), td['ticker'], FONT_TICKER_SUMMARY, border=BORDER_ALL)
        for c in range(3, 8):
            apply_cell(ws.cell(row, c), border=BORDER_ALL)

    # ── Gap columns H, I, J (cols 8-10) ──

    # ── P FCF Valuation (cols K-P, i.e. 11-16) ──
    apply_cell(ws.cell(1, 11), 'P FCF Valuation', FONT_SECTION_TITLE, FILL_HEADER, ALIGN_CENTER)
    for c in range(12, 17):
        apply_cell(ws.cell(1, c), fill=FILL_HEADER)

    pfcf_headers = ['Ticker', 'Market cap', 'TTM FCF', 'P FCF', '5Y', '10Y']
    for j, h in enumerate(pfcf_headers):
        apply_cell(ws.cell(2, 11 + j), h, FONT_HEADER, FILL_HEADER, ALIGN_CENTER, BORDER_ALL)

    for i, td in enumerate(tickers):
        row = 3 + i
        apply_cell(ws.cell(row, 11), td['ticker'], FONT_TICKER_SUMMARY, border=BORDER_ALL)
        for c in range(12, 17):
            apply_cell(ws.cell(row, c), border=BORDER_ALL)

    # ── Instructions (cols AQ-AR, i.e. 43-44) ──
    apply_cell(ws.cell(3, 43), 'Step 1', FONT_INSTRUCTION_BOLD)
    apply_cell(ws.cell(3, 44), 'Fill out valuation area', FONT_INSTRUCTION)
    apply_cell(ws.cell(4, 43), 'Stock Analysis, TTM CF statement', FONT_INSTRUCTION)
    apply_cell(ws.cell(5, 43), 'Average market cap', FONT_INSTRUCTION)
    apply_cell(ws.cell(7, 43), 'Step 2', FONT_INSTRUCTION_BOLD)

    return 3 + len(tickers)  # last summary row


def _write_detail_block(ws, td, start_row):
    """Write one ticker's detail block (ticker row + 5 data rows).
    Returns the next available row (after blank separator).
    """
    quarters = td['quarters']
    n_q = len(quarters)
    r = start_row

    # ── Ticker header row with quarter dates ──
    apply_cell(ws.cell(r, 1), td['ticker'], FONT_TICKER_DETAIL, FILL_TICKER_ROW, ALIGN_LEFT, BORDER_ALL)
    for j, q in enumerate(quarters):
        label = q.strftime('%Y-%m-%d')
        apply_cell(ws.cell(r, 2 + j), label, FONT_DATE, FILL_TICKER_ROW, ALIGN_CENTER, BORDER_ALL)

    # ── Data rows ──
    row_defs = [
        ('Net Income TTM', td['net_income'], FMT_INT),
        ('FCF TTM', td['fcf'], FMT_INT),
        ('Average market cap', td['avg_mcap'], FMT_INT),
        ('Average PE', td['avg_pe'], FMT_DEC2),
        ('Average P FCF', td['avg_pfcf'], FMT_DEC2),
    ]

    for k, (label, values, fmt) in enumerate(row_defs):
        dr = r + 1 + k
        use_zebra = (k % 2 == 0)
        fill = FILL_DATA_ALT if use_zebra else FILL_NONE

        apply_cell(ws.cell(dr, 1), label, FONT_ROW_LABEL, fill, ALIGN_LEFT, BORDER_ALL)

        for j, val in enumerate(values):
            apply_cell(
                ws.cell(dr, 2 + j), val, FONT_DATA, fill, ALIGN_RIGHT, BORDER_ALL, fmt
            )

    return r + 7  # ticker row + 5 data rows + 1 blank separator


def build_workbook(tickers_data, output_path):
    """Build the complete valuation workbook and save it."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Valuation'

    # ── Summary tables ──
    last_summary_row = _write_summary_tables(ws, tickers_data)

    # ── Detail blocks ──
    current_row = last_summary_row + 2  # blank row after summary

    for td in tickers_data:
        if not td['quarters']:
            continue
        current_row = _write_detail_block(ws, td, current_row)

    # ── Column widths ──
    ws.column_dimensions['A'].width = COL_WIDTH_LABEL
    # Summary PE table
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH_TICKER if c == 2 else 12
    # Gap
    for c in range(8, 11):
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH_GAP
    # Summary P FCF table
    for c in range(11, 17):
        ws.column_dimensions[get_column_letter(c)].width = 14 if c > 11 else COL_WIDTH_TICKER
    # Data columns in detail blocks
    for c in range(2, 42):
        if ws.column_dimensions[get_column_letter(c)].width < COL_WIDTH_DATA:
            ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH_DATA
    # Instructions
    ws.column_dimensions[get_column_letter(43)].width = 18
    ws.column_dimensions[get_column_letter(44)].width = 25

    # ── Freeze panes: freeze column A ──
    ws.freeze_panes = 'B1'

    wb.save(output_path)
    print(f"\nSaved: {output_path}")