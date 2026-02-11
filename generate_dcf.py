#!/usr/bin/env python3
"""
generate_dcf.py
===============
Main entry point – generates a DCF valuation Excel workbook for any stock.

Usage:
    python generate_dcf.py AAPL                    # Apple Inc.
    python generate_dcf.py MSFT                    # Microsoft
    python generate_dcf.py TSLA -o tesla_dcf.xlsx  # Custom output file
    python generate_dcf.py GOOGL --fred-key YOUR_KEY  # Use FRED API for rates
    python generate_dcf.py AAPL --sample           # Force sample data (offline)

The script will:
  1. Fetch stock price, financials, and balance sheet from Yahoo Finance (free)
  2. Fetch current interest rates from Yahoo Finance ^TNX or FRED API (free)
  3. Run a full DCF model with 5 scenarios:
       - Base Case:     Moderate growth, current rates
       - Bull Case:     Rising sales & profit, falling interest rates
       - Bear Case:     Falling sales & profit, rising interest rates
       - Rising Rates:  Stable sales, aggressive rate hikes (+200bp)
       - Falling Rates: Stable sales, rate cuts (−150bp)
  4. Generate a professionally formatted Excel workbook with 12 sheets
"""

import argparse
import sys
import os

# Add script directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from data_fetcher import fetch_all
from dcf_engine import run_all_scenarios
from excel_builder import build_workbook, save_workbook


def main():
    parser = argparse.ArgumentParser(
        description="Generate a DCF valuation Excel workbook for any stock.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python generate_dcf.py AAPL
  python generate_dcf.py MSFT -o microsoft_dcf.xlsx
  python generate_dcf.py TSLA --projection-years 7
  python generate_dcf.py GOOGL --fred-key YOUR_FRED_API_KEY
  python generate_dcf.py AAPL --sample   # Use sample data (no API calls)
        """
    )
    parser.add_argument("ticker", type=str, help="Stock ticker symbol (e.g., AAPL, MSFT, TSLA)")
    parser.add_argument("-o", "--output", type=str, default=None,
                        help="Output Excel file path (default: <TICKER>_DCF_Model.xlsx)")
    parser.add_argument("--fred-key", type=str, default=None,
                        help="FRED API key for interest rate data (free at https://fred.stlouisfed.org/docs/api/api_key.html)")
    parser.add_argument("--projection-years", type=int, default=5,
                        help="Number of projection years (default: 5)")
    parser.add_argument("--terminal-growth", type=float, default=0.025,
                        help="Terminal growth rate (default: 0.025 = 2.5%%)")
    parser.add_argument("--erp", type=float, default=0.055,
                        help="Equity Risk Premium (default: 0.055 = 5.5%%)")
    parser.add_argument("--sample", action="store_true",
                        help="Force use of sample data (no API calls)")

    args = parser.parse_args()
    ticker = args.ticker.upper().strip()

    output = args.output or f"{ticker}_DCF_Model.xlsx"

    print(f"\n{'='*60}")
    print(f"  DCF Model Generator")
    print(f"  Ticker: {ticker}")
    print(f"{'='*60}\n")

    # Step 1: Fetch data
    print("[1/4] Fetching financial data...")
    data = fetch_all(ticker, fred_api_key=args.fred_key, force_sample=args.sample)
    stock = data["stock"]
    financials = data["financials"]
    rates = data["rates"]
    source = data["source"]

    print(f"       Source: {source.upper()}")
    print(f"       Company: {stock['company_name']}")
    print(f"       Current Price: ${stock['current_price']:.2f}")
    print(f"       Risk-Free Rate: {rates['risk_free_rate']:.2%}")

    # Step 2: Run DCF model
    print("\n[2/4] Running DCF valuation model...")
    model_result = run_all_scenarios(
        stock, financials, rates,
        projection_years=args.projection_years,
        terminal_growth=args.terminal_growth,
        equity_risk_premium=args.erp,
    )

    wacc = model_result["wacc_data"]["wacc"]
    print(f"       WACC: {wacc:.2%}")
    print(f"       Cost of Equity: {model_result['wacc_data']['cost_of_equity']:.2%}")
    print(f"       Cost of Debt: {model_result['wacc_data']['cost_of_debt']:.2%}")

    # Step 3: Build Excel
    print("\n[3/4] Building Excel workbook (12 sheets)...")
    wb = build_workbook(stock, financials, rates, model_result, source)

    # Step 4: Save
    print(f"\n[4/4] Saving to {output}...")
    save_workbook(wb, output)

    # Print summary
    print(f"\n{'='*60}")
    print(f"  VALUATION SUMMARY — {ticker}")
    print(f"{'='*60}")
    print(f"  Current Price: ${stock['current_price']:.2f}")
    print(f"")
    print(f"  {'Scenario':<20} {'Implied Price':>14} {'Upside/Downside':>16}")
    print(f"  {'-'*50}")

    for key in ["bull", "base", "bear", "rate_hike", "rate_cut"]:
        dcf = model_result["scenarios"][key]["dcf"]
        arrow = "▲" if dcf["upside_downside"] > 0 else "▼"
        print(f"  {dcf['scenario']:<20} ${dcf['implied_share_price']:>12,.2f} {arrow} {dcf['upside_downside']:>14.1%}")

    print(f"\n  Output: {os.path.abspath(output)}")
    print(f"  Sheets: Dashboard, Income Statement, Balance Sheet,")
    print(f"          Cash Flow, WACC, 5× DCF Scenarios,")
    print(f"          Scenario Comparison, Sensitivity Analysis")
    print(f"{'='*60}\n")

    if source == "sample":
        print("  NOTE: Using sample data. Run on a machine with internet")
        print("  access and yfinance installed for live market data.\n")


if __name__ == "__main__":
    main()
