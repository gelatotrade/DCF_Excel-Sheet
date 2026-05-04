"""
data_fetcher.py
===============
Fetches financial data from free APIs for the DCF model.

Sources (tried in order):
  1. Yahoo Finance (via yfinance) — Stock price, financials, balance sheet, cash flow
  2. Alpha Vantage (free API key) — Fallback for stock data and financials
  3. FRED (Federal Reserve Economic Data) — Risk-free rate (10-Year Treasury)
  4. Yahoo Finance ^TNX / ^IRX — Fallback for interest rates
  5. Built-in sample data — When all APIs are unreachable

Usage:
  from data_fetcher import fetch_all
  data = fetch_all("AAPL")
  data = fetch_all("MSFT", alpha_vantage_key="YOUR_KEY")
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
    "MSFT": {
        "stock": {
            "ticker": "MSFT",
            "company_name": "Microsoft Corporation",
            "sector": "Technology",
            "industry": "Software—Infrastructure",
            "current_price": 420.72,
            "market_cap": 3_130_000_000_000,
            "shares_outstanding": 7_433_000_000,
            "beta": 0.89,
            "currency": "USD",
            "country": "United States",
            "trailing_pe": 34.8,
            "forward_pe": 30.2,
            "dividend_yield": 0.0072,
            "fifty_two_week_high": 468.35,
            "fifty_two_week_low": 362.90,
        },
        "financials": {
            "years": ["2024", "2023", "2022", "2021"],
            "revenue":              [245_122_000_000, 211_915_000_000, 198_270_000_000, 168_088_000_000],
            "cost_of_revenue":      [74_073_000_000,  65_863_000_000,  62_650_000_000,  52_232_000_000],
            "gross_profit":         [171_049_000_000, 146_052_000_000, 135_620_000_000, 115_856_000_000],
            "operating_income":     [109_433_000_000, 88_523_000_000,  83_383_000_000,  69_916_000_000],
            "ebitda":               [128_768_000_000, 108_254_000_000, 100_061_000_000, 85_206_000_000],
            "net_income":           [88_136_000_000,  72_361_000_000,  72_738_000_000,  61_271_000_000],
            "tax_provision":        [19_651_000_000,  16_950_000_000,  10_978_000_000,  9_831_000_000],
            "interest_expense":     [2_495_000_000,   2_935_000_000,   2_063_000_000,   2_346_000_000],
            "depreciation":         [19_335_000_000,  19_731_000_000,  16_678_000_000,  15_290_000_000],
            "total_assets":         [512_163_000_000, 411_976_000_000, 364_840_000_000, 333_779_000_000],
            "total_liabilities":    [243_686_000_000, 205_753_000_000, 198_298_000_000, 191_791_000_000],
            "total_equity":         [268_477_000_000, 206_223_000_000, 166_542_000_000, 141_988_000_000],
            "total_debt":           [42_688_000_000,  47_032_000_000,  48_177_000_000,  50_074_000_000],
            "cash":                 [18_315_000_000,  34_704_000_000,  13_931_000_000,  14_224_000_000],
            "current_assets":       [159_734_000_000, 184_257_000_000, 169_684_000_000, 174_149_000_000],
            "current_liabilities":  [109_629_000_000, 104_149_000_000, 95_082_000_000,  88_657_000_000],
            "operating_cash_flow":  [118_548_000_000, 87_582_000_000,  89_035_000_000,  76_740_000_000],
            "capex":                [-44_477_000_000, -28_107_000_000, -23_886_000_000, -20_622_000_000],
            "depreciation_amortization": [19_335_000_000, 19_731_000_000, 16_678_000_000, 15_290_000_000],
            "change_in_working_capital": [-4_564_000_000, -1_346_000_000, -953_000_000,   -936_000_000],
            "free_cash_flow":       [74_071_000_000,  59_475_000_000,  65_149_000_000,  56_118_000_000],
        },
        "rates": {
            "risk_free_rate": 0.0435,
            "treasury_10y": 0.0435,
            "treasury_2y": 0.0420,
            "fed_funds_rate": 0.0433,
            "date_fetched": "2025-01-15 10:00",
        },
    },
    "GOOGL": {
        "stock": {
            "ticker": "GOOGL",
            "company_name": "Alphabet Inc.",
            "sector": "Communication Services",
            "industry": "Internet Content & Information",
            "current_price": 192.53,
            "market_cap": 2_370_000_000_000,
            "shares_outstanding": 12_310_000_000,
            "beta": 1.06,
            "currency": "USD",
            "country": "United States",
            "trailing_pe": 23.1,
            "forward_pe": 20.8,
            "dividend_yield": 0.0045,
            "fifty_two_week_high": 207.05,
            "fifty_two_week_low": 150.22,
        },
        "financials": {
            "years": ["2024", "2023", "2022", "2021"],
            "revenue":              [350_018_000_000, 307_394_000_000, 282_836_000_000, 257_637_000_000],
            "cost_of_revenue":      [148_019_000_000, 133_332_000_000, 126_203_000_000, 110_939_000_000],
            "gross_profit":         [201_999_000_000, 174_062_000_000, 156_633_000_000, 146_698_000_000],
            "operating_income":     [112_387_000_000, 84_293_000_000,  74_842_000_000,  78_714_000_000],
            "ebitda":               [127_762_000_000, 100_533_000_000, 90_776_000_000,  91_155_000_000],
            "net_income":           [100_681_000_000, 73_795_000_000,  59_972_000_000,  76_033_000_000],
            "tax_provision":        [15_569_000_000,  11_922_000_000,  11_356_000_000,  14_701_000_000],
            "interest_expense":     [300_000_000,     308_000_000,     357_000_000,     346_000_000],
            "depreciation":         [15_375_000_000,  16_240_000_000,  15_934_000_000,  12_441_000_000],
            "total_assets":         [432_205_000_000, 402_392_000_000, 365_264_000_000, 359_268_000_000],
            "total_liabilities":    [119_013_000_000, 109_829_000_000, 107_633_000_000, 107_633_000_000],
            "total_equity":         [313_192_000_000, 292_563_000_000, 256_144_000_000, 251_635_000_000],
            "total_debt":           [12_297_000_000,  13_228_000_000,  14_701_000_000,  14_817_000_000],
            "cash":                 [24_480_000_000,  30_691_000_000,  21_879_000_000,  20_945_000_000],
            "current_assets":       [163_085_000_000, 163_922_000_000, 164_795_000_000, 188_143_000_000],
            "current_liabilities":  [81_814_000_000,  81_814_000_000,  69_300_000_000,  64_254_000_000],
            "operating_cash_flow":  [112_765_000_000, 101_746_000_000, 91_495_000_000,  91_652_000_000],
            "capex":                [-52_549_000_000, -32_251_000_000, -31_485_000_000, -24_640_000_000],
            "depreciation_amortization": [15_375_000_000, 16_240_000_000, 15_934_000_000, 12_441_000_000],
            "change_in_working_capital": [-2_834_000_000, 3_766_000_000,  -4_761_000_000, -2_225_000_000],
            "free_cash_flow":       [60_216_000_000,  69_495_000_000,  60_010_000_000,  67_012_000_000],
        },
        "rates": {
            "risk_free_rate": 0.0435,
            "treasury_10y": 0.0435,
            "treasury_2y": 0.0420,
            "fed_funds_rate": 0.0433,
            "date_fetched": "2025-01-15 10:00",
        },
    },
    "TSLA": {
        "stock": {
            "ticker": "TSLA",
            "company_name": "Tesla, Inc.",
            "sector": "Consumer Cyclical",
            "industry": "Auto Manufacturers",
            "current_price": 273.13,
            "market_cap": 877_000_000_000,
            "shares_outstanding": 3_211_000_000,
            "beta": 2.31,
            "currency": "USD",
            "country": "United States",
            "trailing_pe": 130.5,
            "forward_pe": 95.2,
            "dividend_yield": 0.0,
            "fifty_two_week_high": 488.54,
            "fifty_two_week_low": 138.80,
        },
        "financials": {
            "years": ["2024", "2023", "2022", "2021"],
            "revenue":              [97_690_000_000, 96_773_000_000, 81_462_000_000, 53_823_000_000],
            "cost_of_revenue":      [79_688_000_000, 79_113_000_000, 60_609_000_000, 40_217_000_000],
            "gross_profit":         [18_002_000_000, 17_660_000_000, 20_853_000_000, 13_606_000_000],
            "operating_income":     [7_085_000_000,  8_891_000_000,  13_656_000_000, 6_523_000_000],
            "ebitda":               [12_532_000_000, 13_533_000_000, 17_814_000_000, 9_564_000_000],
            "net_income":           [7_091_000_000,  14_997_000_000, 12_556_000_000, 5_519_000_000],
            "tax_provision":        [3_235_000_000,  8_717_000_000,  1_132_000_000,  699_000_000],
            "interest_expense":     [83_000_000,     156_000_000,    191_000_000,    371_000_000],
            "depreciation":         [5_447_000_000,  4_642_000_000,  4_158_000_000,  3_041_000_000],
            "total_assets":         [122_070_000_000, 106_618_000_000, 82_338_000_000, 62_131_000_000],
            "total_liabilities":    [48_015_000_000,  43_009_000_000, 36_440_000_000, 30_548_000_000],
            "total_equity":         [74_055_000_000,  62_634_000_000, 44_704_000_000, 30_189_000_000],
            "total_debt":           [7_526_000_000,   5_748_000_000,  3_380_000_000,  5_245_000_000],
            "cash":                 [36_563_000_000,  29_094_000_000, 22_185_000_000, 17_707_000_000],
            "current_assets":       [57_839_000_000,  49_650_000_000, 40_917_000_000, 27_100_000_000],
            "current_liabilities":  [28_073_000_000,  28_748_000_000, 26_709_000_000, 19_705_000_000],
            "operating_cash_flow":  [11_527_000_000,  13_256_000_000, 14_724_000_000, 11_497_000_000],
            "capex":                [-11_339_000_000, -8_877_000_000, -7_158_000_000, -6_514_000_000],
            "depreciation_amortization": [5_447_000_000, 4_642_000_000, 4_158_000_000, 3_041_000_000],
            "change_in_working_capital": [-1_567_000_000, -2_348_000_000, -1_692_000_000, -518_000_000],
            "free_cash_flow":       [188_000_000,     4_379_000_000,  7_566_000_000,  4_983_000_000],
        },
        "rates": {
            "risk_free_rate": 0.0435,
            "treasury_10y": 0.0435,
            "treasury_2y": 0.0420,
            "fed_funds_rate": 0.0433,
            "date_fetched": "2025-01-15 10:00",
        },
    },
    "AMZN": {
        "stock": {
            "ticker": "AMZN",
            "company_name": "Amazon.com, Inc.",
            "sector": "Consumer Cyclical",
            "industry": "Internet Retail",
            "current_price": 197.12,
            "market_cap": 2_080_000_000_000,
            "shares_outstanding": 10_550_000_000,
            "beta": 1.15,
            "currency": "USD",
            "country": "United States",
            "trailing_pe": 36.4,
            "forward_pe": 28.9,
            "dividend_yield": 0.0,
            "fifty_two_week_high": 242.52,
            "fifty_two_week_low": 161.02,
        },
        "financials": {
            "years": ["2024", "2023", "2022", "2021"],
            "revenue":              [638_000_000_000, 574_785_000_000, 513_983_000_000, 469_822_000_000],
            "cost_of_revenue":      [399_204_000_000, 363_712_000_000, 333_821_000_000, 302_256_000_000],
            "gross_profit":         [238_796_000_000, 211_073_000_000, 180_162_000_000, 167_566_000_000],
            "operating_income":     [68_594_000_000,  36_852_000_000,  12_248_000_000,  24_879_000_000],
            "ebitda":               [115_345_000_000, 85_470_000_000,  55_027_000_000,  59_752_000_000],
            "net_income":           [59_248_000_000,  30_425_000_000,  -2_722_000_000,  33_364_000_000],
            "tax_provision":        [12_785_000_000,  7_120_000_000,   -3_217_000_000,  4_791_000_000],
            "interest_expense":     [3_182_000_000,   3_182_000_000,   2_367_000_000,   1_809_000_000],
            "depreciation":         [46_751_000_000,  48_618_000_000,  42_779_000_000,  34_873_000_000],
            "total_assets":         [624_894_000_000, 527_854_000_000, 462_675_000_000, 420_549_000_000],
            "total_liabilities":    [355_287_000_000, 325_014_000_000, 316_632_000_000, 282_304_000_000],
            "total_equity":         [269_607_000_000, 201_875_000_000, 146_043_000_000, 138_245_000_000],
            "total_debt":           [58_165_000_000,  67_150_000_000,  70_041_000_000,  57_792_000_000],
            "cash":                 [78_837_000_000,  73_387_000_000,  53_888_000_000,  36_220_000_000],
            "current_assets":       [170_188_000_000, 152_733_000_000, 146_791_000_000, 161_580_000_000],
            "current_liabilities":  [179_431_000_000, 164_917_000_000, 155_393_000_000, 142_266_000_000],
            "operating_cash_flow":  [115_877_000_000, 84_946_000_000,  46_752_000_000,  46_327_000_000],
            "capex":                [-83_017_000_000, -48_133_000_000, -58_321_000_000, -61_053_000_000],
            "depreciation_amortization": [46_751_000_000, 48_618_000_000, 42_779_000_000, 34_873_000_000],
            "change_in_working_capital": [-5_745_000_000, -6_242_000_000, -7_894_000_000, -19_308_000_000],
            "free_cash_flow":       [32_860_000_000,  36_813_000_000,  -11_569_000_000, -14_726_000_000],
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


def fetch_stock_data_alphavantage(ticker: str, api_key: str) -> dict:
    base = "https://www.alphavantage.co/query"
    resp = requests.get(base, params={
        "function": "GLOBAL_QUOTE", "symbol": ticker, "apikey": api_key,
    }, timeout=15)
    gq = resp.json().get("Global Quote", {})
    price = float(gq.get("05. price", 0))

    resp2 = requests.get(base, params={
        "function": "OVERVIEW", "symbol": ticker, "apikey": api_key,
    }, timeout=15)
    ov = resp2.json()

    if not price and not ov.get("Symbol"):
        raise ValueError("Alpha Vantage returned no data")

    return {
        "ticker": ticker.upper(),
        "company_name": ov.get("Name", ticker.upper()),
        "sector": ov.get("Sector", "N/A"),
        "industry": ov.get("Industry", "N/A"),
        "current_price": price or float(ov.get("AnalystTargetPrice", 0)),
        "market_cap": float(ov.get("MarketCapitalization", 0)),
        "shares_outstanding": float(ov.get("SharesOutstanding", 0)),
        "beta": float(ov.get("Beta", 1.0)),
        "currency": ov.get("Currency", "USD"),
        "country": ov.get("Country", "N/A"),
        "trailing_pe": float(ov.get("TrailingPE", 0)) or None,
        "forward_pe": float(ov.get("ForwardPE", 0)) or None,
        "dividend_yield": float(ov.get("DividendYield", 0)) or 0,
        "fifty_two_week_high": float(ov.get("52WeekHigh", 0)) or None,
        "fifty_two_week_low": float(ov.get("52WeekLow", 0)) or None,
    }


def fetch_financials_alphavantage(ticker: str, api_key: str) -> dict:
    base = "https://www.alphavantage.co/query"
    inc_r = requests.get(base, params={
        "function": "INCOME_STATEMENT", "symbol": ticker, "apikey": api_key,
    }, timeout=15).json()
    bs_r = requests.get(base, params={
        "function": "BALANCE_SHEET", "symbol": ticker, "apikey": api_key,
    }, timeout=15).json()
    cf_r = requests.get(base, params={
        "function": "CASH_FLOW", "symbol": ticker, "apikey": api_key,
    }, timeout=15).json()

    inc_list = inc_r.get("annualReports", [])[:4]
    bs_list = bs_r.get("annualReports", [])[:4]
    cf_list = cf_r.get("annualReports", [])[:4]

    if not inc_list:
        raise ValueError("Alpha Vantage returned no income statement data")

    def _g(reports, key, default=0):
        return [float(r.get(key, default) or default) for r in reports]

    years = [r.get("fiscalDateEnding", "")[:4] for r in inc_list]

    revenue = _g(inc_list, "totalRevenue")
    cost_of_revenue = _g(inc_list, "costOfRevenue")
    gross_profit = _g(inc_list, "grossProfit")
    operating_income = _g(inc_list, "operatingIncome")
    ebitda = _g(inc_list, "ebitda")
    net_income = _g(inc_list, "netIncome")
    tax_provision = _g(inc_list, "incomeTaxExpense")
    interest_expense = _g(inc_list, "interestExpense")
    depreciation = _g(inc_list, "depreciationAndAmortization")

    total_assets = _g(bs_list, "totalAssets")
    total_liabilities = _g(bs_list, "totalLiabilities")
    total_equity = _g(bs_list, "totalShareholderEquity")
    total_debt = [float(r.get("shortLongTermDebtTotal", 0) or 0) +
                  float(r.get("longTermDebt", 0) or 0) for r in bs_list]
    cash = _g(bs_list, "cashAndCashEquivalentsAtCarryingValue")
    current_assets = _g(bs_list, "totalCurrentAssets")
    current_liabilities = _g(bs_list, "totalCurrentLiabilities")

    operating_cf = _g(cf_list, "operatingCashflow")
    capex = [-abs(float(r.get("capitalExpenditures", 0) or 0)) for r in cf_list]
    depreciation_cf = _g(cf_list, "depreciationDepletionAndAmortization")
    change_in_wc = _g(cf_list, "changeInOperatingLiabilities")

    fcf = [o + c for o, c in zip(operating_cf, capex)]
    dep = depreciation_cf if any(v for v in depreciation_cf) else depreciation

    return {
        "years": years,
        "revenue": revenue,
        "cost_of_revenue": cost_of_revenue,
        "gross_profit": gross_profit,
        "operating_income": operating_income,
        "ebitda": ebitda,
        "net_income": net_income,
        "tax_provision": tax_provision,
        "interest_expense": interest_expense,
        "depreciation": [float(v or 0) for v in dep],
        "total_assets": total_assets,
        "total_liabilities": total_liabilities,
        "total_equity": total_equity,
        "total_debt": total_debt,
        "cash": cash,
        "current_assets": current_assets,
        "current_liabilities": current_liabilities,
        "operating_cash_flow": operating_cf,
        "capex": capex,
        "depreciation_amortization": depreciation_cf,
        "change_in_working_capital": change_in_wc,
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

def fetch_all(ticker: str, fred_api_key: str = None,
              alpha_vantage_key: str = None,
              force_sample: bool = False) -> dict:
    """
    Fetch all data needed for the DCF model.  Returns:
      { "stock": {...}, "financials": {...}, "rates": {...}, "source": "live"|"alphavantage"|"sample" }

    Data source priority:
      1. Yahoo Finance (yfinance) — no API key needed
      2. Alpha Vantage — requires free API key (get at alphavantage.co)
      3. Sample data — built-in fallback for offline use
    """
    ticker = ticker.upper().strip()

    if not force_sample and HAS_YFINANCE:
        try:
            stock = fetch_stock_data_live(ticker)
            financials = fetch_financials_live(ticker)
            rates = fetch_rates_live(fred_api_key)
            return {"stock": stock, "financials": financials, "rates": rates, "source": "live"}
        except Exception as e:
            print(f"[data_fetcher] Yahoo Finance fetch failed ({e}), trying fallbacks...")

    if not force_sample and alpha_vantage_key and HAS_REQUESTS:
        try:
            stock = fetch_stock_data_alphavantage(ticker, alpha_vantage_key)
            financials = fetch_financials_alphavantage(ticker, alpha_vantage_key)
            rates = fetch_rates_live(fred_api_key)
            return {"stock": stock, "financials": financials, "rates": rates, "source": "alphavantage"}
        except Exception as e:
            print(f"[data_fetcher] Alpha Vantage fetch failed ({e}), falling back to sample data.")

    if not force_sample and not HAS_YFINANCE:
        print("[data_fetcher] yfinance not installed. Install with: pip install yfinance")

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
