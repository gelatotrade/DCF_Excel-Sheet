# DCF Valuation Model - Excel Generator

A Python tool that generates a professional **Discounted Cash Flow (DCF) valuation model** as a formatted Excel workbook for any publicly traded stock — complete with **35 charts** across **12 sheets**.

## Features

- **Live Data**: Pulls real-time stock prices, financial statements, and balance sheet data from Yahoo Finance (free, no API key required)
- **Interest Rates**: Fetches current Treasury yields from Yahoo Finance (or FRED API with optional key)
- **Any Stock**: Just change the ticker symbol — works for any publicly traded company
- **5 Scenario Paths**: Models different interest rate and growth trajectories
- **12 Excel Sheets**: Comprehensive, professionally formatted workbook
- **35 Charts**: Bar charts, line charts, pie charts, and sensitivity curves
- **Sensitivity Analysis**: WACC vs Terminal Growth and Revenue Growth vs Margin sensitivity tables with color-coded heatmaps

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

---

## Sheet Previews & Visualizations

### Sheet 1: Dashboard

The main overview page with company info, market data, interest rates, and scenario summary.

```
+===========================================================================+
|  DCF VALUATION MODEL  —  AAPL                                            |
|  Apple Inc.  |  Data Source: LIVE  |  Generated: 2025-01-15               |
+===========================================================================+
|                                                                           |
| COMPANY INFORMATION    | MARKET DATA           | INTEREST RATES           |
|------------------------+-----------------------+--------------------------|
| Ticker     AAPL        | Current Price $232.47 | 10-Year Treasury  4.35%  |
| Company    Apple Inc.  | Market Cap   $3,490B  | 2-Year Treasury   4.20%  |
| Sector     Technology  | Shares Out.   15.1B   | Fed Funds Rate    4.33%  |
| Industry   Consumer El.| Beta           1.24   | WACC             10.94%  |
| Country    United States| Trailing P/E  37.5x  | Cost of Equity   11.17%  |
+===========================================================================+
|                                                                           |
| SCENARIO VALUATION SUMMARY                                                |
|--------+------------------+-------+------+---------+-------+-------------|
|Scenario| Description      | WACC  | TGR  |Impl.Prc| Curr. | Upside     |
|--------+------------------+-------+------+---------+-------+-------------|
| Bull   | Rising sales...  | 9.94% |3.00% | $91.82 |$232.47|  -60.5%    |
| Base   | Moderate growth  |10.94% |2.50% | $65.37 |$232.47|  -71.9%    |
| Bear   | Falling sales... |12.44% |2.00% | $45.18 |$232.47|  -80.6%    |
| Rates+ | Rate hikes +200bp|12.94% |2.50% | $51.20 |$232.47|  -78.0%    |
| Rates- | Rate cuts -150bp | 9.44% |2.50% | $81.69 |$232.47|  -64.9%    |
+========+==================+=======+======+=========+=======+=============+
```

**Charts on this sheet (4):**

```
 +---------------------------------+   +-------------------------------+
 | Revenue vs FCF vs Net Income    |   | Profitability Margins         |
 | ($B)                            |   | Over Time                     |
 |                                 |   |                               |
 |  ████  ████  ████  ████        |   |  ______ Gross ~46%            |
 |  ████  ████  ████  ████        |   | /                             |
 |  ████  ████  ████  ████        |   |/______ Operating ~31%         |
 |  ▓▓▓▓  ▓▓▓▓  ▓▓▓▓  ▓▓▓▓      |   |  ‾‾‾‾‾ Net ~24%              |
 |  ░░░░  ░░░░  ░░░░  ░░░░      |   |  ..... FCF ~27%               |
 |  2021   2022  2023  2024       |   |  2021  2022  2023  2024       |
 +---------------------------------+   +-------------------------------+

 +---------------------------------+   +-------------------------------+
 | Implied Share Price by Scenario |   | WACC vs Terminal Growth       |
 |                                 |   | by Scenario                   |
 |  ████                           |   |  ████                         |
 |  ████ ████                      |   |  ████ ████                    |
 |  ████ ████ ████ ████ ████      |   |  ████ ████ ████ ████ ████    |
 |  ████ ████ ████ ████ ████      |   |  ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓  |
 |  ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓    |   |  Bull  Base  Bear  R+   R-   |
 |  Bull  Base  Bear  R+   R-     |   |  ████ WACC  ▓▓▓▓ Term.Growth |
 +---------------------------------+   +-------------------------------+
```

---

### Sheet 2: Income Statement

```
+===========================================================================+
|  INCOME STATEMENT                                                         |
+===========================================================================+
| Line Item                  |    2024    |    2023    |    2022    |  2021  |
|----------------------------+-----------+-----------+-----------+---------|
| Revenue                    |   $391.0B |   $383.3B |   $394.3B | $365.8B|
| Cost of Revenue            |   $210.4B |   $214.1B |   $223.5B | $213.0B|
| Gross Profit               |   $180.7B |   $169.1B |   $170.8B | $152.8B|
| Operating Income (EBIT)    |   $123.2B |   $114.3B |   $119.4B | $108.9B|
| Net Income                 |    $93.7B |    $97.0B |    $99.8B |  $94.7B|
+----------------------------+-----------+-----------+-----------+---------+

 Charts: [Revenue/GP/OI/NI Bar Chart]  +  [Margin Trends Line Chart]
```

---

### Sheet 3: Balance Sheet

```
+===========================================================================+
|  BALANCE SHEET                                                            |
+===========================================================================+
| Line Item                  |    2024    |    2023    |    2022    |  2021  |
|----------------------------+-----------+-----------+-----------+---------|
| ASSETS                                                                    |
| Current Assets             |   $153.0B |   $143.6B |   $135.4B | $134.8B|
| Cash & Equivalents         |    $29.9B |    $30.0B |    $23.6B |  $34.9B|
| Total Assets               |   $365.0B |   $352.6B |   $352.8B | $351.0B|
| LIABILITIES & EQUITY                                                      |
| Total Debt                 |    $96.8B |   $111.1B |   $120.1B | $124.7B|
| Total Liabilities          |   $308.0B |   $290.4B |   $302.1B | $287.9B|
| Total Equity               |    $57.0B |    $62.1B |    $50.7B |  $63.1B|
+----------------------------+-----------+-----------+-----------+---------+

 Charts: [Assets/Liabilities/Equity Bar Chart]  +  [D/E & Current Ratio Lines]
```

---

### Sheet 4: Cash Flow

```
+===========================================================================+
|  CASH FLOW STATEMENT                                                      |
+===========================================================================+
| Line Item                  |    2024    |    2023    |    2022    |  2021  |
|----------------------------+-----------+-----------+-----------+---------|
| Operating Cash Flow        |   $118.3B |   $110.5B |   $122.2B | $104.0B|
| Capital Expenditure        |    -$10.0B|   -$11.1B |   -$10.7B | -$11.1B|
| Free Cash Flow             |   $108.3B |    $99.5B |   $111.4B |  $93.0B|
+----------------------------+-----------+-----------+-----------+---------+

 Charts: [Operating CF vs FCF Bar Chart]  +  [FCF Yield & CapEx Intensity Lines]
```

---

### Sheet 5: WACC (Weighted Average Cost of Capital)

```
+===========================================================================+
|  WEIGHTED AVERAGE COST OF CAPITAL (WACC)                                  |
+===========================================================================+
| WACC = (E/V) x Re  +  (D/V) x Rd x (1 - T)                              |
|                                                                           |
| COST OF EQUITY (CAPM)       |        | CAPITAL STRUCTURE PIE CHART       |
|-----------------------------+--------|    +------------------+            |
| Risk-Free Rate (Rf)    4.35%|        |    |    ██████████    |            |
| Beta                    1.24 |        |    |  ██  Equity ██  |            |
| Equity Risk Premium    5.50% |        |    |  ██  97.3%  ██  |            |
| Cost of Equity (Re)   11.17%|        |    |  ██████████████  |            |
|                              |        |    |  ░░ Debt 2.7% ░ |            |
| COST OF DEBT                 |        |    +------------------+            |
| Interest Expense      $3.6B  |        |                                    |
| Total Debt           $96.8B  |        | WACC COMPONENT RATES BAR CHART   |
| Cost of Debt (Rd)     3.72%  |        |    Rf    Re    Rd   ATRd  WACC   |
| Tax Rate             24.10%  |        |   ████  ████  ████  ████  ████   |
| After-Tax CoD         2.82%  |        |   4.4% 11.2%  3.7%  2.8% 10.9%  |
|                              |        |                                    |
| RESULT:  WACC = 10.94%       |        | RATE ENVIRONMENT BAR CHART        |
+==============================+========+====================================+
```

---

### Sheets 6-10: DCF Scenario Models

Each of the 5 scenarios gets its own detailed sheet with **3 charts**.

```
+===========================================================================+
|  DCF MODEL  —  BULL CASE                                                  |
|  Rising sales & profit, falling interest rates (-100bp)                   |
+===========================================================================+
|                                     |                                     |
| SCENARIO ASSUMPTIONS                | REVENUE PROJECTION CHART            |
| Revenue Growth Adj.   +3.0%         |  ████ ████                          |
| Margin Adjustment     +2.0%         |  ████ ████ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓   |
| WACC Adjustment       -1.0%         |  ████ ████ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓   |
| Scenario WACC          9.94%        |  2023 2024 2025 2026 2027 2028     |
| Terminal Growth        3.00%        |  ████ Historical  ▓▓▓▓ Projected    |
|                                     |                                     |
+=============================================================================+
| HISTORICAL  <-->  PROJECTED                                               |
|           | 2024  | 2023  | 2022  | 2021  | -> |  2025 |  2026 |  2027  | |
|-----------+-------+-------+-------+-------+----+-------+-------+--------|
| Revenue   | 391.0B| 383.3B| 394.3B| 365.8B|    | 411.2B| 427.4B| 440.1B| |
| Growth %  |  2.0% | -2.8% |  7.8% |       |    |  5.3% |  4.7% |  4.1%| |
| EBIT      | 123.2B| 114.3B| 119.4B| 108.9B|    | 139.2B| 144.6B| 148.9B| |
| NOPAT     |  93.5B|  86.8B|  90.7B|  82.7B|    | 105.7B| 109.8B| 113.1B| |
| (+) D&A   |  11.4B|  11.5B|  11.1B|  11.3B|    |  12.0B|  12.5B|  12.8B| |
| (-) CapEx |  10.0B|  11.1B|  10.7B|  11.1B|    |  10.5B|  10.9B|  11.2B| |
| = UFCF    | 108.3B|  99.5B| 111.4B|  93.0B|    | 107.3B| 111.4B| 114.7B| |
+=============================================================================+
| VALUATION BRIDGE                |  FCF vs PRESENT VALUE CHART             |
| Sum of PV(FCFs)       $420.1B   |   ████ ████ ████ ████ ████             |
| PV of Terminal Value  $968.4B   |   ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓           |
| = Enterprise Value   $1,388.5B  |   ████ Projected FCF                    |
| (-) Net Debt           $66.9B   |   ▓▓▓▓ PV of FCF                       |
| = Equity Value       $1,321.7B  |   2025  2026  2027  2028  2029         |
| Shares Outstanding      15.1B   +----------------------------------------+
| = Implied Price        $91.82   |  VALUATION BRIDGE BAR CHART            |
| Current Price         $232.47   |   PV FCFs | PV Term | EV | -Debt | Eq |
| Upside/Downside       -60.5%   |   ████████████████████████████████████  |
+=================================+=========================================+
```

**Scenarios available:**

| Sheet | Scenario | Growth | Margin | WACC | Color |
|-------|----------|--------|--------|------|-------|
| 6 | Base Case | +0% | +0% | +0bp | Blue |
| 7 | Bull Case | +3% | +2% | -100bp | Green |
| 8 | Bear Case | -3% | -2% | +150bp | Red |
| 9 | Rising Rates | +0% | -0.5% | +200bp | Orange |
| 10 | Falling Rates | +0% | +0.5% | -150bp | Teal |

---

### Sheet 11: Scenario Comparison

Side-by-side comparison of all 5 scenarios with **5 charts**.

```
+===========================================================================+
|  SCENARIO COMPARISON  —  ALL PATHS                                        |
+===========================================================================+
| Metric              | Bull Case | Base Case | Bear Case | Rates+  | Rates-|
|---------------------+-----------+-----------+-----------+---------+-------|
| WACC                |    9.94%  |   10.94%  |   12.44%  |  12.94% |  9.44%|
| Terminal Growth     |    3.00%  |    2.50%  |    2.00%  |   2.50% |  2.50%|
| Enterprise Value    | $1,388.5B |  $921.1B  |  $616.1B  | $707.6B|$1,168B|
| Implied Price       |   $91.82  |   $65.37  |   $45.18  |  $51.20 | $81.69|
| Upside/Downside     |   -60.5%  |   -71.9%  |   -80.6%  |  -78.0% | -64.9%|
+=====================+===========+===========+===========+=========+=======+

 +-----------------------------------+   +-----------------------------------+
 | Implied Price vs Current Price    |   | Enterprise Value Breakdown        |
 | by Scenario                       |   | (Stacked: PV FCFs + PV Terminal)  |
 |  ▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓ |   |  ████                             |
 |  ████                             |   |  ████ ████                        |
 |  ████ ████                        |   |  ████ ████ ████ ████ ████        |
 |  ████ ████ ████ ████ ████        |   |  ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓      |
 |  Bull  Base  Bear  R+   R-       |   |  Bull  Base  Bear  R+   R-       |
 +-----------------------------------+   +-----------------------------------+

 +-----------------------------------+   +-----------------------------------+
 | WACC vs Terminal Growth           |   | Projected FCF All Scenarios       |
 |  ████       ████                  |   |  _____ Bull                       |
 |  ████ ████  ████ ████            |   | / ____ Base                       |
 |  ████ ████  ████ ████ ████      |   |/ / ___ Rates-                     |
 |  ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓ ▓▓▓▓    |   |  / ___ Rates+                     |
 |  Bull  Base  Bear  R+   R-      |   |   /    Bear                       |
 |  ████ WACC  ▓▓▓▓ Term.Growth   |   |  2025  2026  2027  2028  2029     |
 +-----------------------------------+   +-----------------------------------+

                 +-----------------------------------+
                 | Projected Revenue All Scenarios   |
                 |   _____ Bull                      |
                 |  / ____ Base                      |
                 | / /                               |
                 |  / ____ Rates+/-                  |
                 |  /      Bear                      |
                 |  2025  2026  2027  2028  2029     |
                 +-----------------------------------+
```

---

### Sheet 12: Sensitivity Analysis

Two sensitivity tables + two line charts showing how WACC and terminal growth affect the implied price.

```
+===========================================================================+
|  SENSITIVITY ANALYSIS  —  IMPLIED SHARE PRICE                             |
+===========================================================================+
| WACC vs TERMINAL GROWTH RATE                                              |
|                                                                           |
| WACC \ TGR  | 1.00% | 1.50% | 2.00% | 2.50% | 3.00% | 3.50% | 4.00%  |
|-------------+-------+-------+-------+-------+-------+-------+---------|
|  7.94%      |$84.65 |$91.45 |$99.73 |$110.1 |$123.5 |$141.6 |$167.8  |
|  8.94%      |$73.44 |$78.31 |$84.12 |$91.21 |$100.1 |$111.4 |$126.5  |
|  9.94%      |$64.74 |$68.39 |$72.69 |$77.83 |$84.12 |$91.96 |$102.1  |
| >10.94%<    |$57.78 |$60.60 |$63.86 |>65.37<|$72.65 |$78.23 |$85.31  |
| 11.94%      |$52.10 |$54.31 |$56.83 |$59.75 |$63.15 |$67.22 |$72.14  |
| 12.94%      |$47.37 |$49.11 |$51.09 |$53.37 |$56.01 |$59.09 |$62.74  |
| 13.94%      |$43.37 |$44.76 |$46.33 |$48.10 |$50.15 |$52.52 |$55.28  |
|             |       |       |       |       |       |       |         |
| Green = >10% above current price                                         |
| Red   = >10% below current price                                         |
+==========================================================================+

| REVENUE GROWTH vs OPERATING MARGIN SENSITIVITY                            |
|                                                                           |
| Growth\Marg | 27.4% | 29.4% | 31.5% | 33.5% | 35.5%                    |
|-------------+-------+-------+-------+-------+-------                    |
| -1.7%       |$50.05 |$55.80 |$61.55 |$67.30 |$73.05                    |
|  0.3%       |$53.08 |$59.13 |$65.19 |$71.25 |$77.30                    |
|  2.3%       |$56.24 |$62.62 |$69.01 |$75.39 |$81.78                    |
+==========================================================================+

 +-----------------------------------+   +-----------------------------------+
 | Impact of WACC on Implied Price   |   | Impact of Terminal Growth on      |
 |                                   |   | Implied Price                     |
 |  \                                |   |                           /       |
 |   \                               |   |                         /         |
 |    \                              |   |                       /           |
 |     \______                       |   |                _____/             |
 |             ‾‾‾‾‾‾‾‾‾‾          |   |  ____________/                    |
 |  7.9% 8.9% 9.9% 10.9% ... 13.9% |   |  1.0% 1.5% 2.0% 2.5% ... 4.0%  |
 +-----------------------------------+   +-----------------------------------+
```

---

## Complete Chart Inventory (35 Charts)

| Sheet | Charts | Chart Types |
|-------|--------|-------------|
| Dashboard | 4 | Revenue/FCF/NI bar, Margins line, Scenario prices bar, WACC/TG bar |
| Income Statement | 2 | Revenue waterfall bar, Profit margins line |
| Balance Sheet | 2 | Structure bar, D/E & Current ratio line |
| Cash Flow | 2 | OCF vs FCF bar, FCF yield & CapEx line |
| WACC | 3 | Capital structure pie, Component rates bar, Rate environment bar |
| DCF Base Case | 3 | Revenue projection bar, FCF vs PV bar, Valuation bridge bar |
| DCF Bull Case | 3 | Revenue projection bar, FCF vs PV bar, Valuation bridge bar |
| DCF Bear Case | 3 | Revenue projection bar, FCF vs PV bar, Valuation bridge bar |
| DCF Rising Rates | 3 | Revenue projection bar, FCF vs PV bar, Valuation bridge bar |
| DCF Falling Rates | 3 | Revenue projection bar, FCF vs PV bar, Valuation bridge bar |
| Scenario Comparison | 5 | Price comparison bar, EV stacked bar, WACC/TG bar, FCF line, Revenue line |
| Sensitivity Analysis | 2 | WACC impact line, Terminal growth impact line |
| **Total** | **35** | |

---

## Scenarios Modeled

| Scenario | Revenue Growth | Margin | WACC | Description |
|----------|---------------|--------|------|-------------|
| **Bull Case** | +3% | +2% | -100bp | Rising sales & profit, falling rates |
| **Base Case** | Historical avg | Historical avg | Current | Moderate growth, current rates |
| **Bear Case** | -3% | -2% | +150bp | Falling sales & profit, rising rates |
| **Rising Rates** | Stable | -0.5% | +200bp | Aggressive rate hikes |
| **Falling Rates** | Stable | +0.5% | -150bp | Rate cuts |

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

```
                          HISTORICAL DATA                    PROJECTIONS
                     ┌──────────────────────┐          ┌──────────────────────┐
                     │  Revenue             │          │  Revenue (grown by   │
  Yahoo Finance ────>│  Operating Income    │────>─────│  historical avg +    │
  (Free API)         │  Net Income          │          │  scenario adj.)      │
                     │  Free Cash Flow      │          │                      │
                     │  Balance Sheet       │          │  EBIT = Rev x Margin │
                     │  Debt / Cash         │          │  NOPAT = EBIT(1-T)   │
                     └──────────────────────┘          │  FCF = NOPAT+D&A-CapEx│
                                                       └──────────┬───────────┘
                                                                  │
  ┌─────────────────────────────────────────────────────────────── │
  │                                                                │
  │  WACC CALCULATION                                              v
  │  ┌──────────────────────┐                      ┌──────────────────────┐
  │  │  Cost of Equity (Re) │                      │  DISCOUNT FCFs       │
  │  │  = Rf + Beta x ERP   │                      │                      │
  │  │                      │                      │  PV = FCF/(1+WACC)^t │
  │  │  Cost of Debt (Rd)   │────────>─────────────│                      │
  │  │  = Int.Exp / Debt    │         WACC         │  Terminal Value =    │
  │  │                      │                      │  FCF*(1+g)/(WACC-g)  │
  │  │  WACC = (E/V)Re +   │                      │                      │
  │  │    (D/V)Rd(1-T)     │                      │  Enterprise Value =  │
  │  └──────────────────────┘                      │  Sum(PV) + PV(TV)   │
  │                                                └──────────┬───────────┘
  │                                                           │
  │  Yahoo Finance ^TNX ──> Risk-Free Rate                    v
  │  (or FRED API)                                 ┌──────────────────────┐
  │                                                │  EQUITY VALUE        │
  │  5 SCENARIOS apply adjustments:                │  = EV - Net Debt     │
  │  ┌─────────────────────────┐                   │                      │
  │  │ Bull:  Growth+3%, WACC-1%│                  │  IMPLIED PRICE       │
  │  │ Base:  No adjustments    │──────>───────────│  = Equity / Shares   │
  │  │ Bear:  Growth-3%, WACC+1.5%│               │                      │
  │  │ Rates+: WACC+2%         │                   │  UPSIDE/DOWNSIDE     │
  │  │ Rates-: WACC-1.5%      │                   │  vs Current Price    │
  │  └─────────────────────────┘                   └──────────────────────┘
  └────────────────────────────────────────────────────────────────────────
```

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
├── excel_builder.py    # Excel workbook generation (35 charts, 12 sheets)
├── requirements.txt    # Python dependencies
└── README.md
```
