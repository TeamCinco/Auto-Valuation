# Valuation Sheet Builder

Automated valuation competition sheet generator for equity research. Pulls market cap and TTM cash flow data from stockanalysis.com, computes quarterly PE and P/FCF ratios across a peer group, and exports a styled Excel workbook matching an exact hedge-fund visual spec.

This README documents the **styling spec and output architecture** — the "what the sheet looks like and why" layer. For usage (folder structure, CLI, input file naming), see the project's main docs.

---

## Output Architecture

Every workbook has two stacked sections:

**Top — Summary tables (side by side)**
Two six-column tables comparing current valuation across the full peer group.

- **PE Valuation** (cols B–G): Ticker · Price · EPS · PE · 5Y · 10Y
- **P FCF Valuation** (cols K–P): Ticker · Market cap · TTM FCF · P FCF · 5Y · 10Y
- Three-column gap (H–J) separates them

**Bottom — Detail blocks (one per ticker)**
Quarterly time series: Net Income TTM, FCF TTM, Average market cap, Average PE, Average P FCF. One block per ticker, vertically stacked, separated by a blank row.

All detail blocks are aligned to a **unified date spine** — every ticker has the same quarter columns, with missing quarters blank. A company that IPO'd later won't break horizontal alignment with its older peers.

---

## Color Spec (exact)

The visual system uses six flat colors. No gradients, no transparency.

| Color | Hex | RGB | Used for |
|---|---|---|---|
| Primary Yellow | `#FFFF00` | (255, 255, 0) | Section titles, PE & P/FCF highlight columns |
| Soft Header Yellow | `#FFF2CC` | (255, 242, 204) | Column label row |
| Input Red | `#FF0000` | (255, 0, 0) | EPS column (manual inputs) |
| Table Body Gray | `#F2F2F2` | (242, 242, 242) | Summary table body cells |
| Divider Gray | `#D9D9D9` | (217, 217, 217) | Company header rows (META, GOOG, ...) |
| White | `#FFFFFF` | (255, 255, 255) | Time-series data cells |
| Gridline Gray | `#BFBFBF` | (191, 191, 191) | Subtle borders only where needed |

**The system in plain English:**

- **Red = inputs.** Cells you type values into.
- **Yellow = focus metrics.** The numbers that matter — the ratios the eye should land on.
- **Gray = structure.** Grouping, separation, header rows. Gray says "this is a label, not a value."
- **White = raw data.** The time-series at the bottom is just numbers. Let them breathe.

---

## Layout Rules

### Rule 1 — No gridlines on the tables

Excel's default gridlines are turned off (`ws.sheet_view.showGridLines = False`). Color separation does the structural work. The only borders are thin gridline-gray (`#BFBFBF`) around summary table cells — subtle, not heavy.

This is the single most important visual rule. Gridlines make a sheet look like a spreadsheet. The absence of gridlines makes it look like a report.

### Rule 2 — Color carries hierarchy, not bold and size

Font size doesn't change by importance. The primary yellow section titles are 12pt bold; everything else is 11pt. Hierarchy comes from fill color, not typography gymnastics.

### Rule 3 — Horizontal alignment of all tickers

Every ticker detail block uses the same date columns. If META has data from 2016-Q1 and SNAP only from 2017-Q2, the spine runs from 2016-Q1 forward and SNAP's pre-IPO columns are blank. This is a data integrity choice: blank ≠ zero. A zero would falsely imply SNAP earned $0; blank correctly says "no data exists."

### Rule 4 — Flat fills only

No gradients, no patterns, no conditional formatting dressed up as style. One color per cell.

### Rule 5 — Negatives in parentheses, missing in blank

Number formats:

- Integers: `#,##0;(#,##0);"-"`
- Two-decimal: `#,##0.00;(#,##0.00);"-"`
- Price: `#,##0.00`

A `None` value leaves the cell truly empty. A `0` value renders as `"-"`. Negatives render in parentheses, accounting-style.

---

## Typography

| Element | Font | Size | Weight |
|---|---|---|---|
| Section titles ("PE Valuation", "P FCF Valuation") | Calibri | 12 | Bold |
| Column headers | Calibri | 11 | Bold |
| Company detail headers (META/GOOG row) | Calibri | 11 | Bold |
| Quarter dates | Calibri | 11 | Bold |
| PE & P/FCF highlight cells | Calibri | 11 | Bold |
| EPS input cells | Calibri | 11 | Bold (white on red) |
| Body data | Calibri | 11 | Regular |
| Row labels | Calibri | 11 | Regular |

Text color is black (`#000000`) everywhere except the EPS input column, where it's white (`#FFFFFF`) for contrast on the red fill.

---

## Alignment

- **Text labels** (Ticker, row labels like "Net Income TTM") → left
- **Numeric data** → right
- **Section titles and column headers** → center
- **Quarter dates in detail block headers** → center

Vertical alignment is centered everywhere.

---

## Column Widths

| Range | Width | Purpose |
|---|---|---|
| Column A | 20 | Row labels ("Net Income TTM", "Average market cap", etc.) |
| B–G | 12 | PE Valuation table |
| H–J | 3 | Gap between the two summary tables |
| K–P | 12 | P FCF Valuation table |
| Detail block columns | 13 (minimum) | Quarter-by-quarter data — auto-widens to cover full spine |
| Column AQ (43) | 18 | Reserved for analyst notes |
| Column AR (44) | 25 | Reserved for analyst notes |

Freeze pane at `B1` — ticker column stays visible when scrolling right through quarters.

---

## Code Structure

The project is split into five modules. This README focuses on the two that shape the visual output.

```
auto_valuation/
├── main.py              # CLI entry point
├── data_loader.py       # Parse market cap + cash flow files
├── calculations.py      # Compute PE, P/FCF; produce TickerResult objects
├── styles.py            # ⬅ All colors, fonts, borders, widths live here
├── excel_writer.py      # ⬅ Builds the workbook using the style constants
└── input_data/          # Drop downloaded files here
```

### `styles.py`

Single source of truth for the visual system. All hex values, fonts, fills, borders, number formats, and column widths are named constants. If you want to change the primary yellow, change `CLR_PRIMARY_YELLOW` in one place — every yellow element in the workbook updates.

The `apply_cell()` helper applies value, font, fill, alignment, border, and number format in one call.

### `excel_writer.py`

Three phases:

1. **`_write_summary_tables(ws, tickers)`** — builds the two six-column tables at the top.
2. **`_build_date_spine(tickers_data)` + `_align_to_spine(td, spine)`** — compute the unified quarter axis and pad each ticker's parallel lists to match it.
3. **`_write_detail_block(ws, td, spine, start_row)`** — writes one ticker's stacked time-series block using the aligned data.

Top-level `build_workbook(tickers_data, output_path)` orchestrates all three.

---

## Color → Cell Map (reference)

If you're reading the code and want to know which constant governs which visual element:

| Cell | Fill | Font |
|---|---|---|
| Row 1 section titles | `FILL_SECTION_TITLE` | `FONT_SECTION_TITLE` |
| Row 2 column headers | `FILL_SUMMARY_HEADER` | `FONT_HEADER` |
| Summary body (Price, 5Y, 10Y, Market cap, TTM FCF) | `FILL_TABLE_BODY` | `FONT_SUMMARY_DATA` |
| Summary ticker column | `FILL_TABLE_BODY` | `FONT_TICKER_SUMMARY` |
| EPS cells | `FILL_INPUT` | `FONT_INPUT` |
| PE cells | `FILL_HIGHLIGHT` | `FONT_HIGHLIGHT` |
| P FCF cells | `FILL_HIGHLIGHT` | `FONT_HIGHLIGHT` |
| Company detail header row (META/GOOG + dates) | `FILL_DIVIDER` | `FONT_TICKER_DETAIL` / `FONT_DATE` |
| Detail block data cells | `FILL_WHITE` | `FONT_DATA` |
| Detail block row labels | `FILL_WHITE` | `FONT_ROW_LABEL` |

---

## Why these choices

**Why yellow for the highlights and red for inputs?**
Red pulls the eye hardest — reserve it for cells that demand attention, which are the ones you have to type into. Yellow is attention-getting but not alarming; perfect for the computed ratios that are the point of the sheet. Following this convention keeps the visual grammar consistent: the eye learns that red = do something, yellow = look here.

**Why gray for company separators instead of yellow?**
Gray says "structural element, not data." Using yellow for both the focus metrics and the company separators would dilute the yellow's meaning. One color, one job.

**Why kill gridlines?**
Default Excel gridlines are visual noise. They compete with your actual borders and make every cell feel equally important. Turning them off lets the color system carry the structure, which is what makes the sheet read as a finished document rather than a working spreadsheet.

**Why align all tickers to a shared date spine?**
Because valuation is a cross-sectional comparison exercise. You want your eye to sweep vertically down 2020-Q1 and see all peers at once. Unaligned blocks force you to hunt for matching quarters, which kills comparison speed.

---

## Dependencies

- `openpyxl` — workbook construction
- `pandas` — data handling upstream
- `numpy` — numeric helpers

Python 3.9+.