"""
dcf_engine.py
=============
Core DCF valuation calculations with multi-scenario analysis.

Scenarios modeled:
  1. Base Case         – moderate growth, current rates
  2. Bull Case         – rising sales & profit, falling rates
  3. Bear Case         – falling sales & profit, rising rates
  4. Rate Hike Case    – stable sales, rising rates
  5. Rate Cut Case     – stable sales, falling rates
"""

from dataclasses import dataclass, field


# ============================================================================
# SCENARIO DEFINITIONS
# ============================================================================

@dataclass
class Scenario:
    name: str
    description: str
    # Revenue growth rate adjustments (annual)
    revenue_growth_adj: float        # added to base growth (e.g. +0.02 = 2% higher growth)
    # Operating margin adjustment
    margin_adj: float                # added to base margin (e.g. +0.01 = 1% wider)
    # WACC / discount rate adjustment
    wacc_adj: float                  # added to computed WACC (e.g. +0.01 = 100bp higher)
    # Terminal growth rate adjustment
    terminal_growth_adj: float       # added to terminal growth


SCENARIOS = {
    "base": Scenario(
        name="Base Case",
        description="Moderate growth, current interest rates maintained",
        revenue_growth_adj=0.0,
        margin_adj=0.0,
        wacc_adj=0.0,
        terminal_growth_adj=0.0,
    ),
    "bull": Scenario(
        name="Bull Case",
        description="Rising sales & profit, falling interest rates (-100bp)",
        revenue_growth_adj=0.03,       # +3% revenue growth
        margin_adj=0.02,               # +2% margin expansion
        wacc_adj=-0.01,                # -100bp discount rate (lower rates)
        terminal_growth_adj=0.005,     # +50bp terminal growth
    ),
    "bear": Scenario(
        name="Bear Case",
        description="Falling sales & profit, rising interest rates (+150bp)",
        revenue_growth_adj=-0.03,      # -3% revenue growth
        margin_adj=-0.02,              # -2% margin compression
        wacc_adj=0.015,                # +150bp discount rate (higher rates)
        terminal_growth_adj=-0.005,    # -50bp terminal growth
    ),
    "rate_hike": Scenario(
        name="Rising Rates",
        description="Stable sales, aggressive rate hikes (+200bp)",
        revenue_growth_adj=0.0,
        margin_adj=-0.005,             # slight margin pressure from higher costs
        wacc_adj=0.02,                 # +200bp
        terminal_growth_adj=0.0,
    ),
    "rate_cut": Scenario(
        name="Falling Rates",
        description="Stable sales, rate cuts (-150bp)",
        revenue_growth_adj=0.0,
        margin_adj=0.005,              # slight margin benefit
        wacc_adj=-0.015,               # -150bp
        terminal_growth_adj=0.0,
    ),
}


# ============================================================================
# WACC CALCULATION
# ============================================================================

def compute_wacc(stock: dict, financials: dict, rates: dict,
                 equity_risk_premium: float = 0.055) -> dict:
    """
    Compute Weighted Average Cost of Capital.

    WACC = (E/V) * Re + (D/V) * Rd * (1 - Tax)

    Where:
      Re = Rf + Beta * ERP           (CAPM)
      Rd = Interest Expense / Debt   (implied cost of debt)
    """
    rf = rates["risk_free_rate"]
    beta = stock.get("beta", 1.0) or 1.0

    # Cost of equity (CAPM)
    cost_of_equity = rf + beta * equity_risk_premium

    # Cost of debt (implied)
    debt = financials["total_debt"][0] if financials["total_debt"][0] else 1
    interest = abs(financials["interest_expense"][0]) if financials["interest_expense"][0] else 0
    cost_of_debt = interest / debt if debt > 0 else rf

    # Tax rate (effective)
    tax_prov = abs(financials["tax_provision"][0]) if financials["tax_provision"][0] else 0
    pretax_income = (abs(financials["net_income"][0]) + tax_prov) if financials["net_income"][0] else 1
    tax_rate = tax_prov / pretax_income if pretax_income > 0 else 0.21

    # Capital structure
    equity_value = stock.get("market_cap", 0) or (stock["current_price"] * stock["shares_outstanding"])
    debt_value = financials["total_debt"][0]
    total_value = equity_value + debt_value

    weight_equity = equity_value / total_value if total_value > 0 else 0.7
    weight_debt = debt_value / total_value if total_value > 0 else 0.3

    wacc = weight_equity * cost_of_equity + weight_debt * cost_of_debt * (1 - tax_rate)

    return {
        "wacc": wacc,
        "cost_of_equity": cost_of_equity,
        "cost_of_debt": cost_of_debt,
        "tax_rate": tax_rate,
        "weight_equity": weight_equity,
        "weight_debt": weight_debt,
        "risk_free_rate": rf,
        "beta": beta,
        "equity_risk_premium": equity_risk_premium,
        "equity_value": equity_value,
        "debt_value": debt_value,
    }


# ============================================================================
# FCF PROJECTION
# ============================================================================

def compute_historical_metrics(financials: dict) -> dict:
    """Derive growth rates, margins, and ratios from historical data."""
    rev = financials["revenue"]
    fcf = financials["free_cash_flow"]
    oi = financials["operating_income"]
    ni = financials["net_income"]
    capex = financials["capex"]
    da = financials["depreciation"]

    # Revenue growth rates (YoY, most recent first so [0] vs [1] is latest)
    rev_growths = []
    for i in range(len(rev) - 1):
        if rev[i + 1] and rev[i + 1] != 0:
            rev_growths.append((rev[i] - rev[i + 1]) / abs(rev[i + 1]))
        else:
            rev_growths.append(0)

    # Margins
    op_margins = [oi[i] / rev[i] if rev[i] else 0 for i in range(len(rev))]
    net_margins = [ni[i] / rev[i] if rev[i] else 0 for i in range(len(rev))]
    fcf_margins = [fcf[i] / rev[i] if rev[i] else 0 for i in range(len(rev))]

    # Capex as % of revenue
    capex_pcts = [abs(capex[i]) / rev[i] if rev[i] else 0 for i in range(len(rev))]

    # D&A as % of revenue
    da_pcts = [abs(da[i]) / rev[i] if rev[i] else 0 for i in range(len(rev))]

    avg_rev_growth = sum(rev_growths) / len(rev_growths) if rev_growths else 0.05
    avg_op_margin = sum(op_margins) / len(op_margins) if op_margins else 0.15
    avg_fcf_margin = sum(fcf_margins) / len(fcf_margins) if fcf_margins else 0.10
    avg_capex_pct = sum(capex_pcts) / len(capex_pcts) if capex_pcts else 0.03
    avg_da_pct = sum(da_pcts) / len(da_pcts) if da_pcts else 0.03

    return {
        "revenue_growths": rev_growths,
        "operating_margins": op_margins,
        "net_margins": net_margins,
        "fcf_margins": fcf_margins,
        "capex_pcts": capex_pcts,
        "da_pcts": da_pcts,
        "avg_revenue_growth": avg_rev_growth,
        "avg_operating_margin": avg_op_margin,
        "avg_fcf_margin": avg_fcf_margin,
        "avg_capex_pct": avg_capex_pct,
        "avg_da_pct": avg_da_pct,
    }


def project_fcf(financials: dict, metrics: dict, scenario: Scenario,
                projection_years: int = 5) -> dict:
    """
    Project Free Cash Flow for the next N years under a given scenario.

    Approach:
      1. Project revenue using historical growth + scenario adjustment
      2. Apply operating margin (historical avg + scenario adjustment) to get EBIT
      3. Compute NOPAT = EBIT * (1 - tax_rate)
      4. Add back D&A, subtract CapEx, subtract change in working capital
      5. Result = Unlevered Free Cash Flow
    """
    base_rev = financials["revenue"][0]  # most recent year
    base_growth = metrics["avg_revenue_growth"]
    base_op_margin = metrics["avg_operating_margin"]
    base_capex_pct = metrics["avg_capex_pct"]
    base_da_pct = metrics["avg_da_pct"]

    # Tax rate from most recent year
    tax_prov = abs(financials["tax_provision"][0]) if financials["tax_provision"][0] else 0
    pretax = abs(financials["net_income"][0]) + tax_prov
    tax_rate = tax_prov / pretax if pretax > 0 else 0.21

    # Adjusted rates for this scenario
    adj_growth = max(base_growth + scenario.revenue_growth_adj, -0.15)  # floor at -15%
    adj_margin = max(base_op_margin + scenario.margin_adj, 0.01)       # floor at 1%

    projected_revenue = []
    projected_ebit = []
    projected_nopat = []
    projected_da = []
    projected_capex = []
    projected_fcf = []
    growth_rates = []
    margins = []

    for yr in range(1, projection_years + 1):
        # Growth fades toward terminal rate over time
        fade_factor = 1 - (yr - 1) / (projection_years * 2)  # gradual fade
        growth = adj_growth * fade_factor
        growth = max(growth, 0.005) if adj_growth > 0 else growth  # floor positive growth

        rev = base_rev * (1 + growth) ** yr
        ebit = rev * adj_margin
        nopat = ebit * (1 - tax_rate)
        da = rev * base_da_pct
        capex = rev * base_capex_pct
        fcf = nopat + da - capex

        projected_revenue.append(rev)
        projected_ebit.append(ebit)
        projected_nopat.append(nopat)
        projected_da.append(da)
        projected_capex.append(capex)
        projected_fcf.append(fcf)
        growth_rates.append(growth)
        margins.append(adj_margin)

    return {
        "scenario": scenario.name,
        "projection_years": projection_years,
        "growth_rates": growth_rates,
        "margins": margins,
        "tax_rate": tax_rate,
        "projected_revenue": projected_revenue,
        "projected_ebit": projected_ebit,
        "projected_nopat": projected_nopat,
        "projected_da": projected_da,
        "projected_capex": projected_capex,
        "projected_fcf": projected_fcf,
    }


# ============================================================================
# DCF VALUATION
# ============================================================================

def compute_dcf(projection: dict, wacc_data: dict, scenario: Scenario,
                stock: dict, financials: dict,
                terminal_growth: float = 0.025) -> dict:
    """
    Compute enterprise value via DCF, then derive equity value per share.

    Terminal Value = FCF_n * (1 + g) / (WACC - g)   [Gordon Growth Model]
    """
    wacc = wacc_data["wacc"] + scenario.wacc_adj
    wacc = max(wacc, 0.04)  # floor at 4%

    tg = terminal_growth + scenario.terminal_growth_adj
    tg = min(tg, wacc - 0.01)  # terminal growth must be < WACC
    tg = max(tg, 0.005)        # floor at 0.5%

    fcfs = projection["projected_fcf"]
    n = len(fcfs)

    # Present value of projected FCFs
    pv_fcfs = []
    for i, fcf in enumerate(fcfs):
        pv = fcf / (1 + wacc) ** (i + 1)
        pv_fcfs.append(pv)

    pv_fcf_total = sum(pv_fcfs)

    # Terminal value
    terminal_fcf = fcfs[-1] * (1 + tg)
    terminal_value = terminal_fcf / (wacc - tg)
    pv_terminal = terminal_value / (1 + wacc) ** n

    # Enterprise value
    enterprise_value = pv_fcf_total + pv_terminal

    # Equity value
    net_debt = financials["total_debt"][0] - financials["cash"][0]
    equity_value = enterprise_value - net_debt

    # Per share
    shares = stock["shares_outstanding"]
    value_per_share = equity_value / shares if shares > 0 else 0

    # Upside / downside
    current_price = stock["current_price"]
    upside = (value_per_share - current_price) / current_price if current_price > 0 else 0

    return {
        "scenario": scenario.name,
        "scenario_description": scenario.description,
        "wacc": wacc,
        "terminal_growth": tg,
        "pv_fcfs": pv_fcfs,
        "pv_fcf_total": pv_fcf_total,
        "terminal_value": terminal_value,
        "pv_terminal_value": pv_terminal,
        "enterprise_value": enterprise_value,
        "net_debt": net_debt,
        "equity_value": equity_value,
        "shares_outstanding": shares,
        "implied_share_price": value_per_share,
        "current_price": current_price,
        "upside_downside": upside,
    }


# ============================================================================
# RUN ALL SCENARIOS
# ============================================================================

def run_all_scenarios(stock: dict, financials: dict, rates: dict,
                      projection_years: int = 5,
                      terminal_growth: float = 0.025,
                      equity_risk_premium: float = 0.055) -> dict:
    """
    Run the full DCF model for all scenarios.
    Returns a dict with all intermediate and final results.
    """
    wacc_data = compute_wacc(stock, financials, rates, equity_risk_premium)
    metrics = compute_historical_metrics(financials)

    results = {}
    for key, scenario in SCENARIOS.items():
        projection = project_fcf(financials, metrics, scenario, projection_years)
        dcf = compute_dcf(projection, wacc_data, scenario, stock, financials, terminal_growth)
        results[key] = {
            "scenario": scenario,
            "projection": projection,
            "dcf": dcf,
        }

    return {
        "wacc_data": wacc_data,
        "metrics": metrics,
        "scenarios": results,
        "projection_years": projection_years,
        "terminal_growth": terminal_growth,
    }


# ============================================================================
if __name__ == "__main__":
    from data_fetcher import fetch_all
    import sys

    ticker = sys.argv[1] if len(sys.argv) > 1 else "AAPL"
    data = fetch_all(ticker)

    result = run_all_scenarios(data["stock"], data["financials"], data["rates"])

    print(f"\n{'='*70}")
    print(f"  DCF VALUATION — {ticker}  (data source: {data['source']})")
    print(f"{'='*70}")
    print(f"\n  WACC: {result['wacc_data']['wacc']:.2%}")
    print(f"  Cost of Equity: {result['wacc_data']['cost_of_equity']:.2%}")
    print(f"  Cost of Debt:   {result['wacc_data']['cost_of_debt']:.2%}")
    print(f"  Tax Rate:       {result['wacc_data']['tax_rate']:.2%}")

    print(f"\n  Current Price: ${data['stock']['current_price']:.2f}")
    print(f"\n  {'Scenario':<20} {'Implied Price':>15} {'Upside/Downside':>18}")
    print(f"  {'-'*53}")
    for key in ["bull", "base", "bear", "rate_hike", "rate_cut"]:
        r = result["scenarios"][key]["dcf"]
        print(f"  {r['scenario']:<20} ${r['implied_share_price']:>13,.2f} {r['upside_downside']:>17.1%}")
