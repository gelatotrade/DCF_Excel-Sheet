# DCF Valuation Model - Excel Generator

A Python tool that generates a professional **Discounted Cash Flow (DCF) valuation model** as a formatted Excel workbook for any publicly traded stock.

## Features

- **Live Data**: Pulls real-time stock prices, financial statements, and balance sheet data from Yahoo Finance (free, no API key required)
- **Interest Rates**: Fetches current Treasury yields from Yahoo Finance (or FRED API with optional key)
- **Any Stock**: Just change the ticker symbol — works for any publicly traded company
- **5 Scenario Paths**: Models different interest rate and growth trajectories
- **12 Excel Sheets**: Comprehensive, professionally formatted workbook
- **Sensitivity Analysis**: WACC vs Terminal Growth and Revenue Growth vs Margin sensitivity tables

## Quick Start

```bash
# Install dependencies
pip install -r requirements.txt

# Generate DCF model for Apple
python generate_dcf.py AAPL

# Generate for Microsoft with custom output file
python generate_dcf.py MSFT -o microsoft_valuation.xlsx

# Generate for Tesla with 7-year projections
python generate_dcf.py TSLA --projection-years 7

# Use FRED API for interest rates (optional, free key)
python generate_dcf.py GOOGL --fred-key YOUR_API_KEY
```

## Scenarios Modeled

| Scenario | Revenue Growth | Margin | WACC | Description |
|----------|---------------|--------|------|-------------|
| **Bull Case** | +3% | +2% | -100bp | Rising sales & profit, falling rates |
| **Base Case** | Historical avg | Historical avg | Current | Moderate growth, current rates |
| **Bear Case** | -3% | -2% | +150bp | Falling sales & profit, rising rates |
| **Rising Rates** | Stable | -0.5% | +200bp | Aggressive rate hikes |
| **Falling Rates** | Stable | +0.5% | -150bp | Rate cuts |

## Excel Workbook Contents

1. **Dashboard** — Company overview, key metrics, scenario summary with chart
2. **Income Statement** — 4 years of historical revenue, EBIT, net income
3. **Balance Sheet** — Assets, liabilities, equity, debt, cash
4. **Cash Flow** — Operating CF, CapEx, D&A, Free Cash Flow
5. **WACC** — Full CAPM/WACC breakdown with current interest rates
6. **DCF Base Case** — Detailed base-case projections and valuation bridge
7. **DCF Bull Case** — Rising sales, falling rates scenario
8. **DCF Bear Case** — Falling sales, rising rates scenario
9. **DCF Rising Rates** — Stable growth, rate hike path
10. **DCF Falling Rates** — Stable growth, rate cut path
11. **Scenario Comparison** — Side-by-side comparison with chart
12. **Sensitivity Analysis** — WACC × Terminal Growth and Growth × Margin tables

## Data Sources (All Free)

| Data | Source | API Key Required? |
|------|--------|-------------------|
| Stock price, market cap | Yahoo Finance (yfinance) | No |
| Income statement | Yahoo Finance | No |
| Balance sheet | Yahoo Finance | No |
| Cash flow statement | Yahoo Finance | No |
| 10-Year Treasury yield | Yahoo Finance ^TNX | No |
| Federal Funds Rate | FRED API (optional) | Free key at [fred.stlouisfed.org](https://fred.stlouisfed.org/docs/api/api_key.html) |

## How the DCF Model Works

1. **WACC Calculation**: Uses CAPM (Rf + Beta × ERP) for cost of equity, implied cost of debt from financials, and market-cap-based capital structure weights
2. **FCF Projection**: Projects revenue using historical growth rates (with scenario adjustments), applies operating margins, computes NOPAT, adds back D&A, subtracts CapEx
3. **Terminal Value**: Gordon Growth Model — FCF × (1+g) / (WACC - g)
4. **Valuation Bridge**: Enterprise Value → subtract Net Debt → Equity Value → divide by shares → Implied Share Price

## CLI Options

```
usage: generate_dcf.py [-h] [-o OUTPUT] [--fred-key FRED_KEY]
                       [--projection-years N] [--terminal-growth RATE]
                       [--erp RATE] [--sample] ticker

Arguments:
  ticker                Stock ticker symbol (e.g., AAPL, MSFT, TSLA)
  -o, --output          Output Excel file path (default: <TICKER>_DCF_Model.xlsx)
  --fred-key            FRED API key for interest rate data
  --projection-years    Number of projection years (default: 5)
  --terminal-growth     Terminal growth rate (default: 0.025 = 2.5%)
  --erp                 Equity Risk Premium (default: 0.055 = 5.5%)
  --sample              Force use of sample data (no API calls)
```

## Project Structure

```
DCF_Excel-Sheet/
├── generate_dcf.py     # Main entry point — run this
├── data_fetcher.py     # Yahoo Finance + FRED API data retrieval
├── dcf_engine.py       # DCF calculations, WACC, scenarios
├── excel_builder.py    # Excel workbook generation with formatting
├── requirements.txt    # Python dependencies
└── README.md
```
