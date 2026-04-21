"""
excel_writer.py
Builds the output valuation Excel workbook.

Color spec (exact):
  Section titles ............. #FFFF00 primary yellow
  Summary column headers ..... #FFF2CC soft header yellow
  Summary body cells ......... #F2F2F2 table body gray
  EPS column (inputs) ........ #FF0000 input red
  PE / P FCF columns ......... #FFFF00 primary yellow (focus metrics)
  Company detail headers ..... #D9D9D9 divider gray (META, GOOG, ...)
  Time-series data cells ..... #FFFFFF white
  Subtle borders ............. #BFBFBF gridline gray

All ticker detail blocks are aligned to a unified date spine — every ticker
has the same quarter columns, with missing quarters left blank.
"""

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from styles import (
    FILL_SECTION_TITLE, FILL_HIGHLIGHT, FILL_SUMMARY_HEADER, FILL_INPUT,
    FILL_TABLE_BODY, FILL_WHITE, FILL_DIVIDER, FILL_NONE,
    FONT_SECTION_TITLE, FONT_HEADER, FONT_TICKER_SUMMARY, FONT_SUMMARY_DATA,
    FONT_HIGHLIGHT, FONT_INPUT, FONT_TICKER_DETAIL, FONT_DATE,
    FONT_ROW_LABEL, FONT_DATA,
    ALIGN_CENTER, ALIGN_RIGHT, ALIGN_LEFT,
    BORDER_BOX, BORDER_NONE,
    FMT_INT, FMT_DEC2, FMT_PRICE,
    COL_WIDTH_LABEL, COL_WIDTH_DATA, COL_WIDTH_SUMMARY, COL_WIDTH_TICKER, COL_WIDTH_GAP,
    ROW_HEIGHT_DEFAULT, ROW_HEIGHT_TITLE,
    apply_cell,
)


def _build_date_spine(tickers_data):
    """Sorted union of every ticker's quarter-end dates → shared column axis."""
    all_dates = set()
    for td in tickers_data:
        all_dates.update(td.dates)
    return sorted(all_dates)


def _align_to_spine(td, spine):
    """Pad a ticker's parallel lists to the unified spine; missing → None."""
    index = {d: i for i, d in enumerate(td.dates)}

    ni, fcf, mcap, pe, pfcf = [], [], [], [], []
    for d in spine:
        i = index.get(d)
        if i is None:
            ni.append(None); fcf.append(None); mcap.append(None)
            pe.append(None); pfcf.append(None)
        else:
            ni.append(td.net_income[i])
            fcf.append(td.fcf[i])
            mcap.append(td.avg_mcap[i])
            pe.append(td.avg_pe[i])
            pfcf.append(td.avg_pfcf[i])
    return ni, fcf, mcap, pe, pfcf


def _write_summary_tables(ws, tickers):
    """Write PE Valuation and P FCF Valuation summary tables.

    Layout (1-indexed columns):
      PE Valuation:    B(2)-G(7)     → Ticker, Price, EPS, PE, 5Y, 10Y
      Gap:             H(8)-J(10)
      P FCF Valuation: K(11)-P(16)   → Ticker, Market cap, TTM FCF, P FCF, 5Y, 10Y

    Colors per spec:
      Row 1 section titles → primary yellow (#FFFF00)
      Row 2 headers        → soft header yellow (#FFF2CC)
      Body cells           → table body gray (#F2F2F2)
      EPS column           → input red (#FF0000), white bold text
      PE / P FCF columns   → primary yellow (#FFFF00), black bold text
    """
    # ── Row 1: Section titles (merged, primary yellow, centered, bold) ──
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=7)
    apply_cell(
        ws.cell(1, 2), 'PE Valuation',
        FONT_SECTION_TITLE, FILL_SECTION_TITLE, ALIGN_CENTER, BORDER_BOX,
    )
    for c in range(3, 8):
        ws.cell(1, c).fill = FILL_SECTION_TITLE
        ws.cell(1, c).border = BORDER_BOX

    ws.merge_cells(start_row=1, start_column=11, end_row=1, end_column=16)
    apply_cell(
        ws.cell(1, 11), 'P FCF Valuation',
        FONT_SECTION_TITLE, FILL_SECTION_TITLE, ALIGN_CENTER, BORDER_BOX,
    )
    for c in range(12, 17):
        ws.cell(1, c).fill = FILL_SECTION_TITLE
        ws.cell(1, c).border = BORDER_BOX

    ws.row_dimensions[1].height = ROW_HEIGHT_TITLE

    # ── Row 2: Column headers (soft header yellow) ──
    pe_headers = ['Ticker', 'Price', 'EPS', 'PE', '5Y', '10Y']
    for j, h in enumerate(pe_headers):
        apply_cell(
            ws.cell(2, 2 + j), h,
            FONT_HEADER, FILL_SUMMARY_HEADER, ALIGN_CENTER, BORDER_BOX,
        )

    pfcf_headers = ['Ticker', 'Market cap', 'TTM FCF', 'P FCF', '5Y', '10Y']
    for j, h in enumerate(pfcf_headers):
        apply_cell(
            ws.cell(2, 11 + j), h,
            FONT_HEADER, FILL_SUMMARY_HEADER, ALIGN_CENTER, BORDER_BOX,
        )

    # ── Rows 3+: Body rows ──
    # PE table:    col 2=Ticker, 3=Price, 4=EPS(RED), 5=PE(YELLOW), 6=5Y, 7=10Y
    # P FCF table: col 11=Ticker, 12=MCap, 13=TTM FCF, 14=P FCF(YELLOW), 15=5Y, 16=10Y
    for i, td in enumerate(tickers):
        row = 3 + i

        # ── PE row ──
        apply_cell(
            ws.cell(row, 2), td.ticker,
            FONT_TICKER_SUMMARY, FILL_TABLE_BODY, ALIGN_LEFT, BORDER_BOX,
        )
        # Price (body gray)
        apply_cell(
            ws.cell(row, 3),
            font=FONT_SUMMARY_DATA, fill=FILL_TABLE_BODY,
            alignment=ALIGN_RIGHT, border=BORDER_BOX,
            number_format=FMT_PRICE,
        )
        # EPS — INPUT column: red fill, white bold text
        apply_cell(
            ws.cell(row, 4),
            font=FONT_INPUT, fill=FILL_INPUT,
            alignment=ALIGN_RIGHT, border=BORDER_BOX,
            number_format=FMT_DEC2,
        )
        # PE — HIGHLIGHT column: primary yellow, bold
        apply_cell(
            ws.cell(row, 5),
            font=FONT_HIGHLIGHT, fill=FILL_HIGHLIGHT,
            alignment=ALIGN_RIGHT, border=BORDER_BOX,
            number_format=FMT_DEC2,
        )
        # 5Y, 10Y (body gray)
        for c in (6, 7):
            apply_cell(
                ws.cell(row, c),
                font=FONT_SUMMARY_DATA, fill=FILL_TABLE_BODY,
                alignment=ALIGN_RIGHT, border=BORDER_BOX,
                number_format=FMT_DEC2,
            )

        # ── P FCF row ──
        apply_cell(
            ws.cell(row, 11), td.ticker,
            FONT_TICKER_SUMMARY, FILL_TABLE_BODY, ALIGN_LEFT, BORDER_BOX,
        )
        # Market cap, TTM FCF (body gray, integer format)
        for c in (12, 13):
            apply_cell(
                ws.cell(row, c),
                font=FONT_SUMMARY_DATA, fill=FILL_TABLE_BODY,
                alignment=ALIGN_RIGHT, border=BORDER_BOX,
                number_format=FMT_INT,
            )
        # P FCF — HIGHLIGHT column: primary yellow, bold
        apply_cell(
            ws.cell(row, 14),
            font=FONT_HIGHLIGHT, fill=FILL_HIGHLIGHT,
            alignment=ALIGN_RIGHT, border=BORDER_BOX,
            number_format=FMT_DEC2,
        )
        # 5Y, 10Y (body gray)
        for c in (15, 16):
            apply_cell(
                ws.cell(row, c),
                font=FONT_SUMMARY_DATA, fill=FILL_TABLE_BODY,
                alignment=ALIGN_RIGHT, border=BORDER_BOX,
                number_format=FMT_DEC2,
            )

    return 3 + len(tickers)


def _write_detail_block(ws, td, spine, start_row):
    """Write one ticker's detail block, aligned to the shared date spine.

    Colors per spec:
      Ticker header row → divider gray (#D9D9D9)
      Data cells        → white (#FFFFFF)

    Layout per block:
      Row 0: TICKER | 2016-03-31 | 2016-06-30 | ...   (gray fill, bold)
      Row 1: Net Income TTM | val | val | ...         (white)
      Row 2: FCF TTM | ...
      Row 3: Average market cap | ...
      Row 4: Average PE | ...
      Row 5: Average P FCF | ...
      Row 6: (blank separator)
    """
    r = start_row
    ni, fcf, mcap, pe, pfcf = _align_to_spine(td, spine)

    # ── Ticker header row (divider gray) ──
    apply_cell(
        ws.cell(r, 1), td.ticker,
        FONT_TICKER_DETAIL, FILL_DIVIDER, ALIGN_LEFT,
    )
    for j, q in enumerate(spine):
        apply_cell(
            ws.cell(r, 2 + j), q.strftime('%Y-%m-%d'),
            FONT_DATE, FILL_DIVIDER, ALIGN_CENTER,
        )

    # ── Data rows (white fill, aligned to spine; None → blank) ──
    row_defs = [
        ('Net Income TTM',     ni,   FMT_INT),
        ('FCF TTM',            fcf,  FMT_INT),
        ('Average market cap', mcap, FMT_INT),
        ('Average PE',         pe,   FMT_DEC2),
        ('Average P FCF',      pfcf, FMT_DEC2),
    ]

    for k, (label, values, fmt) in enumerate(row_defs):
        dr = r + 1 + k
        # Row label also gets white fill to match the data area
        apply_cell(
            ws.cell(dr, 1), label,
            FONT_ROW_LABEL, FILL_WHITE, ALIGN_LEFT,
        )
        for j, val in enumerate(values):
            apply_cell(
                ws.cell(dr, 2 + j), val,
                FONT_DATA, FILL_WHITE, ALIGN_RIGHT, number_format=fmt,
            )

    return r + 7  # ticker row + 5 data rows + 1 blank


def build_workbook(tickers_data, output_path):
    """Build the complete valuation workbook and save it."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Valuation'

    # Turn off gridlines — color separation does the work
    ws.sheet_view.showGridLines = False

    # ── Summary tables ──
    last_summary_row = _write_summary_tables(ws, tickers_data)

    # ── Build the shared date spine for all detail blocks ──
    spine = _build_date_spine(tickers_data)

    # ── Detail blocks (start after a blank row) ──
    current_row = last_summary_row + 2

    for td in tickers_data:
        if not td.dates:
            continue
        current_row = _write_detail_block(ws, td, spine, current_row)

    # ── Column widths ──
    ws.column_dimensions['A'].width = COL_WIDTH_LABEL

    # PE Valuation columns (B-G)
    for c in range(2, 8):
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH_SUMMARY
    # Gap columns (H-J)
    for c in range(8, 11):
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH_GAP
    # P FCF Valuation columns (K-P)
    for c in range(11, 17):
        ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH_SUMMARY

    # Ensure detail-block columns span the full spine
    max_data_col = 2 + len(spine)
    for c in range(2, max_data_col + 1):
        cur = ws.column_dimensions[get_column_letter(c)].width or 0
        if cur < COL_WIDTH_DATA:
            ws.column_dimensions[get_column_letter(c)].width = COL_WIDTH_DATA

    ws.column_dimensions[get_column_letter(43)].width = 18
    ws.column_dimensions[get_column_letter(44)].width = 25

    ws.freeze_panes = 'B1'

    wb.save(output_path)
    print(f"\nSaved: {output_path}")