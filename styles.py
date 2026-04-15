"""
styles.py
All Excel styling constants matching the original valuation sheet.
Colors, fonts, fills, borders, number formats, column widths.
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ── Fills ─────────────────────────────────────────────────────────────
FILL_HEADER = PatternFill('solid', fgColor='4472C4')       # Blue header row
FILL_TICKER_ROW = PatternFill('solid', fgColor='D6DCE4')   # Light grey ticker row in detail blocks
FILL_DATA_ALT = PatternFill('solid', fgColor='D6DCE4')     # Zebra stripe (light grey)
FILL_NONE = PatternFill(fill_type=None)                     # No fill

# ── Fonts ─────────────────────────────────────────────────────────────
FONT_SECTION_TITLE = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
FONT_HEADER = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
FONT_TICKER_SUMMARY = Font(name='Calibri', size=11, bold=False)
FONT_TICKER_DETAIL = Font(name='Calibri', size=11, bold=True)
FONT_DATE = Font(name='Calibri', size=10)
FONT_ROW_LABEL = Font(name='Calibri', size=11)
FONT_DATA = Font(name='Calibri', size=11)
FONT_INSTRUCTION_BOLD = Font(name='Calibri', size=11, bold=True)
FONT_INSTRUCTION = Font(name='Calibri', size=11)

# ── Alignment ─────────────────────────────────────────────────────────
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')
ALIGN_LEFT = Alignment(horizontal='left', vertical='center')

# ── Borders ───────────────────────────────────────────────────────────
THIN_SIDE = Side(style='thin', color='B4C6E7')
BORDER_ALL = Border(
    left=THIN_SIDE, right=THIN_SIDE,
    top=THIN_SIDE, bottom=THIN_SIDE
)
BORDER_BOTTOM = Border(bottom=Side(style='thin', color='B4C6E7'))

# ── Number Formats ────────────────────────────────────────────────────
FMT_INT = '#,##0'
FMT_DEC2 = '#,##0.00'
FMT_PRICE = '#,##0.00'

# ── Column Widths ─────────────────────────────────────────────────────
COL_WIDTH_LABEL = 22       # Column A (row labels)
COL_WIDTH_DATA = 14        # Data columns
COL_WIDTH_TICKER = 8       # Ticker column in summary
COL_WIDTH_GAP = 3          # Gap columns between tables


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