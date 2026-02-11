"""
data_fetcher.py
===============
Fetches financial data from free APIs for the DCF model.

Sources:
  - Yahoo Finance (via yfinance): Stock price, financials, balance sheet, cash flow
  - FRED (Federal Reserve Economic Data): Risk-free rate (10-Year Treasury)
  - Fallback: Built-in sample data when APIs are unreachable

Usage:
  from data_fetcher import fetch_all
  data = fetch_all("AAPL")
"""

import sys
from datetime import datetime

try:
    import yfinance as yf
    HAS_YFINANCE = True
except ImportError:
    HAS_YFINANCE = False

try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False


# ============================================================================
# SAMPLE DATA  –  used when APIs are unreachable (e.g. sandbox / offline)
# ============================================================================

SAMPLE_DATA = {
    "AAPL": {
        "stock": {
            "ticker": "AAPL",
            "company_name": "Apple Inc.",
            "sector": "Technology",
            "industry": "Consumer Electronics",
            "current_price": 232.47,
            "market_cap": 3_490_000_000_000,
            "shares_outstanding": 15_115_800_000,
            "beta": 1.24,
            "currency": "USD",
            "country": "United States",
            "trailing_pe": 37.5,
            "forward_pe": 32.1,
            "dividend_yield": 0.0044,
            "fifty_two_week_high": 260.10,
            "fifty_two_week_low": 164.08,
        },
        "financials": {
            "years": ["2024", "2023", "2022", "2021"],
            "revenue":              [391_035_000_000, 383_285_000_000, 394_328_000_000, 365_817_000_000],
            "cost_of_revenue":      [210_352_000_000, 214_137_000_000, 223_546_000_000, 212_981_000_000],
            "gross_profit":         [180_683_000_000, 169_148_000_000, 170_782_000_000, 152_836_000_000],
            "operating_income":     [123_216_000_000, 114_301_000_000, 119_437_000_000, 108_949_000_000],
            "ebitda":               [134_661_000_000, 125_820_000_000, 130_541_000_000, 120_233_000_000],
            "net_income":           [93_736_000_000,  96_995_000_000,  99_803_000_000,  94_680_000_000],
            "tax_provision":        [29_749_000_000,  16_741_000_000,  19_300_000_000,  14_527_000_000],
            "interest_expense":     [3_600_000_000,   3_933_000_000,   2_931_000_000,   2_645_000_000],
            "depreciation":         [11_445_000_000,  11_519_000_000,  11_104_000_000,  11_284_000_000],
            "total_assets":         [364_980_000_000, 352_583_000_000, 352_755_000_000, 351_002_000_000],
            "total_liabilities":    [308_030_000_000, 290_437_000_000, 302_083_000_000, 287_912_000_000],
            "total_equity":         [56_950_000_000,  62_146_000_000,  50_672_000_000,  63_090_000_000],
            "total_debt":           [96_796_000_000,  111_088_000_000, 120_069_000_000, 124_719_000_000],
            "cash":                 [29_943_000_000,  29_965_000_000,  23_646_000_000,  34_940_000_000],
            "current_assets":       [152_987_000_000, 143_566_000_000, 135_405_000_000, 134_836_000_000],
            "current_liabilities":  [176_392_000_000, 145_308_000_000, 153_982_000_000, 125_481_000_000],
            "operating_cash_flow":  [118_254_000_000, 110_543_000_000, 122_151_000_000, 104_038_000_000],
            "capex":                [-9_959_000_000,  -11_059_000_000, -10_708_000_000, -11_085_000_000],
            "depreciation_amortization": [11_445_000_000, 11_519_000_000, 11_104_000_000, 11_284_000_000],
            "change_in_working_capital": [-5_195_000_000, -6_577_000_000, 1_200_000_000, -4_911_000_000],
            "free_cash_flow":       [108_295_000_000, 99_484_000_000,  111_443_000_000, 92_953_000_000],
        },
        "rates": {
            "risk_free_rate": 0.0435,
            "treasury_10y": 0.0435,
            "treasury_2y": 0.0420,
            "fed_funds_rate": 0.0433,
            "date_fetched": "2025-01-15 10:00",
        },
    },
}

# Generic fallback when ticker not in SAMPLE_DATA
GENERIC_SAMPLE = {
    "stock": {
        "ticker": "SAMPLE",
        "company_name": "Sample Corp",
        "sector": "Technology",
        "industry": "Software",
        "current_price": 100.00,
        "market_cap": 50_000_000_000,
        "shares_outstanding": 500_000_000,
        "beta": 1.1,
        "currency": "USD",
        "country": "United States",
        "trailing_pe": 25.0,
        "forward_pe": 22.0,
        "dividend_yield": 0.01,
        "fifty_two_week_high": 120.00,
        "fifty_two_week_low": 80.00,
    },
    "financials": {
        "years": ["2024", "2023", "2022", "2021"],
        "revenue":              [10_000_000_000, 9_200_000_000, 8_500_000_000, 7_800_000_000],
        "cost_of_revenue":      [6_000_000_000, 5_600_000_000, 5_200_000_000, 4_800_000_000],
        "gross_profit":         [4_000_000_000, 3_600_000_000, 3_300_000_000, 3_000_000_000],
        "operating_income":     [2_500_000_000, 2_200_000_000, 2_000_000_000, 1_800_000_000],
        "ebitda":               [3_000_000_000, 2_700_000_000, 2_500_000_000, 2_200_000_000],
        "net_income":           [2_000_000_000, 1_800_000_000, 1_600_000_000, 1_400_000_000],
        "tax_provision":        [500_000_000,   450_000_000,   400_000_000,   350_000_000],
        "interest_expense":     [200_000_000,   220_000_000,   250_000_000,   280_000_000],
        "depreciation":         [500_000_000,   500_000_000,   500_000_000,   400_000_000],
        "total_assets":         [20_000_000_000, 18_000_000_000, 16_500_000_000, 15_000_000_000],
        "total_liabilities":    [10_000_000_000, 9_500_000_000,  9_000_000_000,  8_500_000_000],
        "total_equity":         [10_000_000_000, 8_500_000_000,  7_500_000_000,  6_500_000_000],
        "total_debt":           [5_000_000_000,  5_500_000_000,  6_000_000_000,  6_200_000_000],
        "cash":                 [2_000_000_000,  1_800_000_000,  1_500_000_000,  1_200_000_000],
        "current_assets":       [6_000_000_000,  5_500_000_000,  5_000_000_000,  4_500_000_000],
        "current_liabilities":  [4_000_000_000,  3_800_000_000,  3_500_000_000,  3_200_000_000],
        "operating_cash_flow":  [2_800_000_000,  2_500_000_000,  2_200_000_000,  2_000_000_000],
        "capex":                [-500_000_000,   -450_000_000,   -400_000_000,   -380_000_000],
        "depreciation_amortization": [500_000_000, 500_000_000, 500_000_000, 400_000_000],
        "change_in_working_capital": [-200_000_000, -150_000_000, -100_000_000, -120_000_000],
        "free_cash_flow":       [2_300_000_000,  2_050_000_000,  1_800_000_000,  1_620_000_000],
    },
    "rates": {
        "risk_free_rate": 0.0435,
        "treasury_10y": 0.0435,
        "treasury_2y": 0.042,
        "fed_funds_rate": 0.0433,
        "date_fetched": "2025-01-15 10:00",
    },
}


# ============================================================================
# LIVE DATA FETCHERS
# ============================================================================

def _safe_row(df, labels: list, default=0) -> list:
    """Try each label in *labels* until one is found in df.index."""
    for label in labels:
        if label in df.index:
            return df.loc[label].tolist()
    return [default] * (len(df.columns) if hasattr(df, "columns") else 4)


def fetch_stock_data_live(ticker: str) -> dict:
    tk = yf.Ticker(ticker)
    info = tk.info
    fast = tk.fast_info
    current_price = getattr(fast, "last_price", None) or info.get("currentPrice") or info.get("regularMarketPrice", 0)
    market_cap = getattr(fast, "market_cap", None) or info.get("marketCap", 0)
    shares = getattr(fast, "shares", None) or info.get("sharesOutstanding", 0)
    return {
        "ticker": ticker.upper(),
        "company_name": info.get("shortName", ticker.upper()),
        "sector": info.get("sector", "N/A"),
        "industry": info.get("industry", "N/A"),
        "current_price": current_price,
        "market_cap": market_cap,
        "shares_outstanding": shares,
        "beta": info.get("beta", 1.0),
        "currency": info.get("currency", "USD"),
        "country": info.get("country", "N/A"),
        "trailing_pe": info.get("trailingPE", None),
        "forward_pe": info.get("forwardPE", None),
        "dividend_yield": info.get("dividendYield", 0) or 0,
        "fifty_two_week_high": info.get("fiftyTwoWeekHigh", None),
        "fifty_two_week_low": info.get("fiftyTwoWeekLow", None),
    }


def fetch_financials_live(ticker: str) -> dict:
    tk = yf.Ticker(ticker)
    inc = tk.income_stmt or __import__("pandas").DataFrame()
    bs = tk.balance_sheet or __import__("pandas").DataFrame()
    cf = tk.cashflow or __import__("pandas").DataFrame()

    def _years(df):
        if df.empty:
            return []
        return [str(c.year) if hasattr(c, "year") else str(c) for c in df.columns]

    years = _years(inc) or _years(bs) or _years(cf) or ["2024", "2023", "2022", "2021"]

    revenue = _safe_row(inc, ["Total Revenue", "Revenue", "Operating Revenue"])
    cost_of_revenue = _safe_row(inc, ["Cost Of Revenue", "Cost of Revenue"])
    gross_profit = _safe_row(inc, ["Gross Profit", "GrossProfit"])
    operating_income = _safe_row(inc, ["Operating Income", "OperatingIncome", "EBIT"])
    ebitda = _safe_row(inc, ["EBITDA", "Ebitda", "Normalized EBITDA"])
    net_income = _safe_row(inc, ["Net Income", "NetIncome", "Net Income Common Stockholders"])
    tax_provision = _safe_row(inc, ["Tax Provision", "TaxProvision", "Income Tax Expense"])
    interest_expense = _safe_row(inc, ["Interest Expense", "InterestExpense", "Interest Expense Non Operating"])
    depreciation_in_inc = _safe_row(inc, ["Reconciled Depreciation", "Depreciation And Amortization In Income Statement"])

    total_assets = _safe_row(bs, ["Total Assets", "TotalAssets"])
    total_liabilities = _safe_row(bs, ["Total Liabilities Net Minority Interest", "Total Liabilities"])
    total_equity = _safe_row(bs, ["Stockholders Equity", "Total Equity Gross Minority Interest"])
    total_debt = _safe_row(bs, ["Total Debt", "TotalDebt", "Long Term Debt And Capital Lease Obligation"])
    cash = _safe_row(bs, ["Cash And Cash Equivalents", "Cash Cash Equivalents And Short Term Investments"])
    current_assets = _safe_row(bs, ["Current Assets", "CurrentAssets", "Total Current Assets"])
    current_liabilities = _safe_row(bs, ["Current Liabilities", "CurrentLiabilities", "Total Current Liabilities"])

    operating_cf = _safe_row(cf, ["Operating Cash Flow", "OperatingCashFlow", "Cash Flow From Continuing Operating Activities"])
    capex = _safe_row(cf, ["Capital Expenditure", "CapitalExpenditure"])
    depreciation_cf = _safe_row(cf, ["Depreciation And Amortization", "DepreciationAndAmortization"])
    change_in_wc = _safe_row(cf, ["Change In Working Capital", "ChangeInWorkingCapital"])

    fcf = [float(o or 0) + float(c or 0) for o, c in zip(operating_cf, capex)]

    dep = depreciation_cf if any(v for v in depreciation_cf) else depreciation_in_inc

    return {
        "years": years,
        "revenue": [float(v or 0) for v in revenue],
        "cost_of_revenue": [float(v or 0) for v in cost_of_revenue],
        "gross_profit": [float(v or 0) for v in gross_profit],
        "operating_income": [float(v or 0) for v in operating_income],
        "ebitda": [float(v or 0) for v in ebitda],
        "net_income": [float(v or 0) for v in net_income],
        "tax_provision": [float(v or 0) for v in tax_provision],
        "interest_expense": [float(v or 0) for v in interest_expense],
        "depreciation": [float(v or 0) for v in dep],
        "total_assets": [float(v or 0) for v in total_assets],
        "total_liabilities": [float(v or 0) for v in total_liabilities],
        "total_equity": [float(v or 0) for v in total_equity],
        "total_debt": [float(v or 0) for v in total_debt],
        "cash": [float(v or 0) for v in cash],
        "current_assets": [float(v or 0) for v in current_assets],
        "current_liabilities": [float(v or 0) for v in current_liabilities],
        "operating_cash_flow": [float(v or 0) for v in operating_cf],
        "capex": [float(v or 0) for v in capex],
        "depreciation_amortization": [float(v or 0) for v in depreciation_cf],
        "change_in_working_capital": [float(v or 0) for v in change_in_wc],
        "free_cash_flow": fcf,
    }


def fetch_rates_live(fred_api_key: str = None) -> dict:
    risk_free = None
    treasury_2y = None
    fed_funds = None

    # Try FRED first
    if fred_api_key and HAS_REQUESTS:
        base = "https://api.stlouisfed.org/fred/series/observations"
        for series, target in [("DGS10", "rf"), ("DGS2", "t2"), ("FEDFUNDS", "ff")]:
            try:
                resp = requests.get(base, params={
                    "series_id": series, "api_key": fred_api_key,
                    "file_type": "json", "sort_order": "desc", "limit": 5,
                }, timeout=10)
                for obs in resp.json().get("observations", []):
                    if obs["value"] != ".":
                        val = float(obs["value"]) / 100.0
                        if target == "rf":
                            risk_free = val
                        elif target == "t2":
                            treasury_2y = val
                        else:
                            fed_funds = val
                        break
            except Exception:
                pass

    # Fallback: Yahoo Finance ^TNX
    if risk_free is None and HAS_YFINANCE:
        try:
            hist = yf.Ticker("^TNX").history(period="5d")
            if not hist.empty:
                risk_free = float(hist["Close"].iloc[-1]) / 100.0
        except Exception:
            pass

    if treasury_2y is None and HAS_YFINANCE:
        try:
            hist = yf.Ticker("^IRX").history(period="5d")
            if not hist.empty:
                treasury_2y = float(hist["Close"].iloc[-1]) / 100.0
        except Exception:
            pass

    return {
        "risk_free_rate": risk_free or 0.043,
        "treasury_10y": risk_free or 0.043,
        "treasury_2y": treasury_2y or 0.042,
        "fed_funds_rate": fed_funds or (risk_free - 0.005 if risk_free else 0.04),
        "date_fetched": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


# ============================================================================
# PUBLIC INTERFACE
# ============================================================================

def fetch_all(ticker: str, fred_api_key: str = None, force_sample: bool = False) -> dict:
    """
    Fetch all data needed for the DCF model.  Returns:
      { "stock": {...}, "financials": {...}, "rates": {...}, "source": "live"|"sample" }
    """
    ticker = ticker.upper().strip()

    if not force_sample and HAS_YFINANCE:
        try:
            stock = fetch_stock_data_live(ticker)
            financials = fetch_financials_live(ticker)
            rates = fetch_rates_live(fred_api_key)
            return {"stock": stock, "financials": financials, "rates": rates, "source": "live"}
        except Exception as e:
            print(f"[data_fetcher] Live fetch failed ({e}), falling back to sample data.")

    # Fallback to sample data
    if ticker in SAMPLE_DATA:
        data = SAMPLE_DATA[ticker]
    else:
        data = GENERIC_SAMPLE.copy()
        data["stock"] = {**GENERIC_SAMPLE["stock"], "ticker": ticker, "company_name": f"{ticker} (Sample)"}

    return {"stock": data["stock"], "financials": data["financials"],
            "rates": data["rates"], "source": "sample"}


# ============================================================================
if __name__ == "__main__":
    t = sys.argv[1] if len(sys.argv) > 1 else "AAPL"
    result = fetch_all(t)
    print(f"\n{'='*60}")
    print(f"  Data for {t}  (source: {result['source']})")
    print(f"{'='*60}")
    for section in ["stock", "rates"]:
        print(f"\n--- {section.upper()} ---")
        for k, v in result[section].items():
            print(f"  {k}: {v}")
    print(f"\n--- FINANCIALS ---")
    fin = result["financials"]
    print(f"  years: {fin['years']}")
    for k in ["revenue", "net_income", "free_cash_flow", "total_debt", "cash"]:
        vals = [f"${v/1e9:.1f}B" for v in fin[k]]
        print(f"  {k}: {vals}")
