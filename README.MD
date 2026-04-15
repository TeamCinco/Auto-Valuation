# Auto-Valuation: Competition Sheet Builder

## What This Tool Does

This tool automates the most tedious part of building a valuation competition sheet. Instead of manually downloading data, copying it into Excel, normalizing numbers, computing quarterly averages, and calculating ratios for every competitor — you drop your raw files into a folder and run one command. The tool does the rest and spits out a formatted Excel workbook ready for analysis.

The output matches the layout of our standard valuation sheet: summary tables at the top (PE Valuation and P/FCF Valuation), followed by detailed quarterly breakdowns for each ticker showing Net Income TTM, FCF TTM, Average Market Cap, Average PE, and Average P/FCF.

---

## The Manual Process (What We Used To Do)

For each competitor in a coverage universe, the old workflow looked like this:

1. Go to stockanalysis.com and download the **daily market cap** history
2. Go to stockanalysis.com and download the **TTM cash flow statement**
3. Open the market cap file and divide every raw number by 1,000,000 to normalize into millions
4. Group the daily market cap data by calendar quarter (Jan–Mar, Apr–Jun, Jul–Sep, Oct–Dec) and compute the **average** for each quarter
5. Open the cash flow statement and pull out just the **Net Income TTM** and **Free Cash Flow TTM** rows
6. For each quarter, compute **Average PE** (market cap ÷ net income) and **Average P/FCF** (market cap ÷ free cash flow)
7. Paste everything into the competition sheet in the right format
8. Repeat for every single ticker

With 7 competitors, this takes 30–60 minutes of copy-paste work and is error-prone. This tool does it in about 5 seconds.

---

## How To Use It

### Step 1: Download the data

For each ticker (e.g., SNAP, META, GOOG), download two files from stockanalysis.com:

- **Market Cap** → the daily historical market cap (download as CSV or Excel)
- **Cash Flow Statement (TTM)** → the trailing twelve months cash flow statement (download as CSV or Excel)

### Step 2: Name the files and put them in one folder

Rename each file following this pattern:

```
TICKER-market-cap.csv              (or .xlsx)
TICKER-cash-flow-statement-ttm.csv (or .xlsx)
```

Drop them all into one folder. For example:

```
input_data/
  SNAP-market-cap.csv
  SNAP-cash-flow-statement-ttm.xlsx
  META-market-cap.csv
  META-cash-flow-statement-ttm.csv
  GOOG-market-cap.xlsx
  GOOG-cash-flow-statement-ttm.xlsx
  PINS-market-cap.csv
  PINS-cash-flow-statement-ttm.csv
```

The ticker is whatever comes before `-market-cap` or `-cash-flow-statement-ttm` in the filename. CSV and XLSX can be mixed freely — use whatever format stockanalysis.com gives you.

### Step 3: Run the tool

Open a terminal, navigate to the project folder, and run:

```
python main.py input_data
```

Or specify a custom output filename:

```
python main.py input_data my_comps.xlsx
```

The tool will print progress as it processes each ticker:

```
Found 3 tickers: META, PINS, SNAP

Processing META...
  Market cap:  META-market-cap.csv
  Cash flow:   META-cash-flow-statement-ttm.xlsx
  Market cap:  40 quarters (2016-03-31 -> 2025-12-31)
  Cash flow:   40 quarters (2016-03-31 -> 2025-12-31)
Processing PINS...
  ...
Processing SNAP...
  ...

Saved: valuation_output.xlsx
Done!
```

### Step 4: Fill in the summary tables

The output Excel file has two summary tables at the top (PE Valuation and P/FCF Valuation) with the ticker names pre-filled but the data columns left blank. This is intentional — those values (current Price, EPS, etc.) change daily and are best entered manually or pulled from a live source.

---

## What Happens Under The Hood

Here is what the tool does internally, step by step:

### 1. File Discovery (`calculations.py`)

The tool scans the input folder for files matching the naming convention. It pairs up each ticker's market cap file with its cash flow file. If a ticker is missing one of the two files, the tool prints a warning and skips that ticker.

### 2. Market Cap Processing (`data_loader.py`)

For each ticker's market cap file:

- Reads the file (CSV or XLSX — auto-detected from the file extension)
- Parses the market cap column, handling multiple formats:
  - Raw numbers: `8,698,100,000`
  - Abbreviated: `11.98B`, `450.5M`, `1.2T`
- Divides every value by 1,000,000 to convert to **millions**
- Assigns each daily data point to its calendar quarter:
  - January 1 – March 31 → Q1
  - April 1 – June 30 → Q2
  - July 1 – September 30 → Q3
  - October 1 – December 31 → Q4
- Computes the **average** market cap for each quarter (mean of all daily values in that quarter)

### 3. Cash Flow Processing (`data_loader.py`)

For each ticker's cash flow file:

- Reads the file (CSV or XLSX)
- Finds the **Year Ending** row to get the quarter dates
- Extracts the **Net Income** row and the **Free Cash Flow** row
- Parses all values, handling stockanalysis.com's formatting:
  - Commas: `" 1,252 "` → `1252`
  - Parentheses for negatives: `"(2,722)"` → `-2722`
  - Spaces, dollar signs, dashes — all handled

### 4. Ratio Calculations (`calculations.py`)

For each quarter where both market cap and cash flow data exist:

- **Average PE** = Average Market Cap ÷ Net Income TTM
- **Average P/FCF** = Average Market Cap ÷ Free Cash Flow TTM

If market cap data is missing for a quarter (e.g., the market cap file doesn't go back far enough), those cells are left blank rather than showing incorrect values.

### 5. Excel Output (`excel_writer.py`, `styles.py`)

The tool builds the output workbook with:

- **Row 1**: Section titles ("PE Valuation" and "P FCF Valuation") with blue headers
- **Row 2**: Column headers (Ticker, Price, EPS, PE, 5Y, 10Y / Ticker, Market cap, TTM FCF, P FCF, 5Y, 10Y)
- **Rows 3–N**: One row per ticker, names pre-filled, data columns blank
- **Below the summary**: One detail block per ticker showing the full quarterly history:
  - Ticker row with quarter-end dates
  - Net Income TTM
  - FCF TTM
  - Average Market Cap (in millions)
  - Average PE
  - Average P/FCF

Styling uses blue headers, grey zebra striping on data rows, Calibri font, and thin borders — matching our standard competition sheet format.

---

## Project Structure

```
auto_valuation/
│
├── main.py              → Entry point. Run this file.
├── data_loader.py       → Reads and parses market cap + cash flow files
├── calculations.py      → Discovers tickers, computes PE and P/FCF ratios
├── styles.py            → Excel colors, fonts, borders (edit here to change styling)
├── excel_writer.py      → Builds the output Excel workbook
├── requirements.txt     → Python dependencies (pandas, numpy, openpyxl)
│
└── input_data/          → Drop your downloaded files here
    ├── SNAP-market-cap.csv
    ├── SNAP-cash-flow-statement-ttm.xlsx
    ├── META-market-cap.csv
    └── META-cash-flow-statement-ttm.csv
```

If you ever need to change the colors or fonts in the output, all of that lives in `styles.py` — one file, clearly labeled constants.

---

## First-Time Setup

You only need to do this once:

```bash
# 1. Install Python (3.9 or newer) if you don't have it

# 2. Create a virtual environment
python -m venv .venv
source .venv/bin/activate    # Mac/Linux
.venv\Scripts\activate       # Windows

# 3. Install dependencies
pip install -r requirements.txt
```

After that, just activate the environment and run `python main.py input_data` each time.

---

## Troubleshooting

**"No valid ticker pairs found"**
→ Check your file naming. Files must follow the pattern `TICKER-market-cap.csv` and `TICKER-cash-flow-statement-ttm.csv`. The ticker part must match between the two files.

**"WARNING: X cash flow quarters have NO market cap data"**
→ Your market cap file doesn't go back far enough. Download the full historical data from stockanalysis.com (it should go back to IPO).

**"TypeError: ufunc 'divide' not supported"**
→ The cash flow file has values that aren't parsing as numbers. This is usually a formatting issue. Update to the latest version of this tool — the current parser handles all known stockanalysis.com formats.

**Blank cells in the detail blocks**
→ This is expected for quarters where market cap data isn't available. The tool only computes ratios where it has both market cap and cash flow data. Missing data = blank cell, not a zero or an error.