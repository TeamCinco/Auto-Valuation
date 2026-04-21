"""
styles.py
Exact palette spec:
  Primary Yellow    #FFFF00 — section titles, PE/P FCF highlight columns
  Soft Header       #FFF2CC — column headers (Ticker, Price, EPS, ...)
  Input Red         #FF0000 — EPS column (inputs)
  Table Body        #F2F2F2 — valuation table body cells
  White             #FFFFFF — time-series data area (bottom section)
  Divider Gray      #D9D9D9 — company detail headers (META, GOOG, ...)
  Gridline Gray     #BFBFBF — subtle borders when used

Rules:
  - Flat fills only, no gradients
  - No heavy borders — color separation does the work
  - Red = inputs, Yellow = focus metrics, Gray = structure, White = raw data
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ── Palette ──────────────────────────────────────────────────────────
CLR_PRIMARY_YELLOW = 'FFFF00'
CLR_SOFT_HEADER    = 'FFF2CC'
CLR_INPUT_RED      = 'FFFF00'
CLR_TABLE_BODY     = 'F2F2F2'
CLR_WHITE          = 'FFFFFF'
CLR_DIVIDER_GRAY   = 'D9D9D9'
CLR_GRIDLINE_GRAY  = 'BFBFBF'
CLR_BLACK          = '000000'

# ── Fills ────────────────────────────────────────────────────────────
FILL_SECTION_TITLE  = PatternFill('solid', fgColor=CLR_PRIMARY_YELLOW)
FILL_HIGHLIGHT      = PatternFill('solid', fgColor=CLR_PRIMARY_YELLOW)   # PE, P FCF columns
FILL_SUMMARY_HEADER = PatternFill('solid', fgColor=CLR_SOFT_HEADER)
FILL_INPUT          = PatternFill('solid', fgColor=CLR_INPUT_RED)        # EPS column
FILL_TABLE_BODY     = PatternFill('solid', fgColor=CLR_TABLE_BODY)
FILL_WHITE          = PatternFill('solid', fgColor=CLR_WHITE)
FILL_DIVIDER        = PatternFill('solid', fgColor=CLR_DIVIDER_GRAY)     # Company header rows
FILL_NONE           = PatternFill(fill_type=None)

# ── Fonts ────────────────────────────────────────────────────────────
FONT_SECTION_TITLE    = Font(name='Calibri', size=12, bold=True, color=CLR_BLACK)
FONT_HEADER           = Font(name='Calibri', size=11, bold=True, color=CLR_BLACK)
FONT_TICKER_SUMMARY   = Font(name='Calibri', size=11, color=CLR_BLACK)
FONT_SUMMARY_DATA     = Font(name='Calibri', size=11, color=CLR_BLACK)
FONT_HIGHLIGHT        = Font(name='Calibri', size=11, bold=True, color=CLR_BLACK)
FONT_INPUT            = Font(name='Calibri', size=11, bold=True, color=CLR_WHITE)  # white on red
FONT_TICKER_DETAIL    = Font(name='Calibri', size=11, bold=True, color=CLR_BLACK)
FONT_DATE             = Font(name='Calibri', size=11, bold=True, color=CLR_BLACK)
FONT_ROW_LABEL        = Font(name='Calibri', size=11, color=CLR_BLACK)
FONT_DATA             = Font(name='Calibri', size=11, color=CLR_BLACK)
FONT_INSTRUCTION_BOLD = Font(name='Calibri', size=11, bold=True)
FONT_INSTRUCTION      = Font(name='Calibri', size=11)

# ── Alignment ────────────────────────────────────────────────────────
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_RIGHT  = Alignment(horizontal='right',  vertical='center')
ALIGN_LEFT   = Alignment(horizontal='left',   vertical='center')

# ── Borders ──────────────────────────────────────────────────────────
# Subtle gridline-gray borders; color carries structure, borders stay light
THIN_GRAY = Side(style='thin', color=CLR_GRIDLINE_GRAY)

BORDER_BOX = Border(
    left=THIN_GRAY, right=THIN_GRAY,
    top=THIN_GRAY,  bottom=THIN_GRAY,
)
BORDER_NONE = Border()

# ── Number Formats ───────────────────────────────────────────────────
FMT_INT   = '#,##0;(#,##0);"-"'
FMT_DEC2  = '#,##0.00;(#,##0.00);"-"'
FMT_PRICE = '#,##0.00'

# ── Column Widths ────────────────────────────────────────────────────
COL_WIDTH_LABEL   = 20
COL_WIDTH_DATA    = 13
COL_WIDTH_SUMMARY = 12
COL_WIDTH_TICKER  = 8
COL_WIDTH_GAP     = 3

# ── Row Heights ──────────────────────────────────────────────────────
ROW_HEIGHT_DEFAULT = 15
ROW_HEIGHT_TITLE   = 20


def apply_cell(cell, value=None, font=None, fill=None, alignment=None,
               border=None, number_format=None):
    """Apply styling to a cell in one call."""
    if value is not None:
        cell.value = value
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border
    if number_format:
        cell.number_format = number_format
    return cell