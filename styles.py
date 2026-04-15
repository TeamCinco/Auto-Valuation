"""
styles.py
All Excel styling constants matching the original valuation sheet.
"""

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ── Fills ─────────────────────────────────────────────────────────────
FILL_TICKER_ROW = PatternFill('solid', fgColor='FFD966')    # Yellow/gold for ticker detail header rows
FILL_NONE = PatternFill(fill_type=None)

# ── Fonts ─────────────────────────────────────────────────────────────
FONT_SECTION_TITLE = Font(name='Calibri', size=11, bold=True)
FONT_HEADER = Font(name='Calibri', size=11, bold=True)
FONT_TICKER_SUMMARY = Font(name='Calibri', size=11)
FONT_TICKER_DETAIL = Font(name='Calibri', size=11, bold=True)
FONT_DATE = Font(name='Calibri', size=11)
FONT_ROW_LABEL = Font(name='Calibri', size=11)
FONT_DATA = Font(name='Calibri', size=11)
FONT_INSTRUCTION_BOLD = Font(name='Calibri', size=11, bold=True)
FONT_INSTRUCTION = Font(name='Calibri', size=11)

# ── Alignment ─────────────────────────────────────────────────────────
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')
ALIGN_LEFT = Alignment(horizontal='left', vertical='center')

# ── Borders ───────────────────────────────────────────────────────────
THIN_SIDE = Side(style='thin', color='000000')
BORDER_BOX = Border(
    left=THIN_SIDE, right=THIN_SIDE,
    top=THIN_SIDE, bottom=THIN_SIDE
)
BORDER_NONE = Border()

# ── Number Formats ────────────────────────────────────────────────────
#    Negatives in parentheses, zeros shown as dash
FMT_INT = '#,##0;(#,##0);"-"'
FMT_DEC2 = '#,##0.00;(#,##0.00);"-"'
FMT_PRICE = '#,##0.00'

# ── Column Widths ─────────────────────────────────────────────────────
COL_WIDTH_LABEL = 20
COL_WIDTH_DATA = 13
COL_WIDTH_SUMMARY = 12
COL_WIDTH_TICKER = 8
COL_WIDTH_GAP = 3


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