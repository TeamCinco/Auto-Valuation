"""
main.py
Entry point for the Valuation Competition Sheet Builder.

Usage:
    python main.py <input_folder> [output_file.xlsx]
    python main.py input_data
    python main.py input_data my_comps.xlsx

File naming convention (all in one flat folder):
    TICKER-market-cap.csv              (or .xlsx)
    TICKER-cash-flow-statement-ttm.csv (or .xlsx)

Example:
    input_data/
      SNAP-market-cap.csv
      SNAP-cash-flow-statement-ttm.xlsx
      META-market-cap.xlsx
      META-cash-flow-statement-ttm.csv
"""

import sys
import os
import logging
import warnings
warnings.filterwarnings('ignore')

from calculations import discover_tickers, process_ticker
from excel_writer import build_workbook

logging.basicConfig(level=logging.INFO, format='%(message)s')


def main():
    input_dir = sys.argv[1] if len(sys.argv) > 1 else 'input_data'
    output_file = sys.argv[2] if len(sys.argv) > 2 else 'valuation_output.xlsx'

    if not os.path.isdir(input_dir):
        print(f"ERROR: Input directory '{input_dir}' not found.")
        print(f"\nCreate the folder and drop your files in:")
        print(f"  {input_dir}/")
        print(f"    SNAP-market-cap.csv")
        print(f"    SNAP-cash-flow-statement-ttm.xlsx")
        print(f"    META-market-cap.csv")
        print(f"    META-cash-flow-statement-ttm.csv")
        print(f"\nFiles can be .csv or .xlsx in any mix.")
        sys.exit(1)

    pairs = discover_tickers(input_dir)

    if not pairs:
        print(f"ERROR: No valid ticker pairs found in '{input_dir}/'")
        print(f"\nExpected naming convention:")
        print(f"  TICKER-market-cap.csv  (or .xlsx)")
        print(f"  TICKER-cash-flow-statement-ttm.csv  (or .xlsx)")
        sys.exit(1)

    tickers = [t for t, _, _ in pairs]
    print(f"Found {len(pairs)} tickers: {', '.join(tickers)}\n")

    all_data = []
    for ticker, mcap_path, cf_path in pairs:
        print(f"Processing {ticker}...")
        result = process_ticker(ticker, mcap_path, cf_path)
        if result.quarters:
            all_data.append(result)

    if not all_data:
        print("ERROR: No valid data processed.")
        sys.exit(1)

    build_workbook(all_data, output_file)
    print("Done!")


if __name__ == '__main__':
    main()