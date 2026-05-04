"""
dcf_engine.py
=============
Core DCF valuation calculations with multi-scenario analysis.

Scenarios modeled:
  1. Base Case         – moderate growth, current rates maintained
  2. Bull Case         – rising sales & profit, falling rates (easing cycle)
  3. Bear Case         – falling sales & profit, rising rates (tightening cycle)
  4. Rate Hike Case    – stable sales, aggressive rate hikes (+200bp over 5 years)
  5. Rate Cut Case     – stable sales, rate cuts (-150bp over 5 years)

Each scenario models year-by-year interest rate trajectories so the WACC
changes dynamically rather than using a single static adjustment.
"""

from dataclasses import dataclass, field
from typing import List


# ============================================================================
# SCENARIO DEFINITIONS
# ============================================================================

@dataclass
class Scenario:
    name: str
    description: str
    # Revenue growth rate adjustments (annual)
    revenue_growth_adj: float
    # Operating margin adjustment
    margin_adj: float
    # WACC / discount rate adjustment (total, spread across projection)
    wacc_adj: float
    # Terminal growth rate adjustment
    terminal_growth_adj: float
    # Year-by-year rate trajectory (bp change per year, e.g. [+50, +50, +25, +25, +25])
    rate_path_bp: list = field(default_factory=list)
    # Revenue growth trajectory modifiers per year (multiplicative)
    growth_trajectory: list = field(default_factory=list)
    # Margin trajectory modifiers per year (additive, applied on top of margin_adj)
    margin_trajectory: list = field(default_factory=list)


SCENARIOS = {
    "base": Scenario(
        name="Base Case",
        description="Moderate growth, current interest rates maintained",
        revenue_growth_adj=0.0,
        margin_adj=0.0,
        wacc_adj=0.0,
        terminal_growth_adj=0.0,
        rate_path_bp=[0, 0, 0, 0, 0],
        growth_trajectory=[1.0, 0.95, 0.90, 0.85, 0.80],
        margin_trajectory=[0.0, 0.0, 0.0, 0.0, 0.0],
    ),
    "bull": Scenario(
        name="Bull Case",
        description="Rising sales & profit margins, falling interest rates (-100bp easing cycle)",
        revenue_growth_adj=0.03,
        margin_adj=0.02,
        wacc_adj=-0.01,
        terminal_growth_adj=0.005,
        rate_path_bp=[-25, -25, -25, -15, -10],
        growth_trajectory=[1.10, 1.08, 1.05, 1.02, 1.00],
        margin_trajectory=[0.005, 0.005, 0.004, 0.003, 0.003],
    ),
    "bear": Scenario(
        name="Bear Case",
        description="Falling sales & profit margins, rising interest rates (+150bp tightening)",
        revenue_growth_adj=-0.03,
        margin_adj=-0.02,
        wacc_adj=0.015,
        terminal_growth_adj=-0.005,
        rate_path_bp=[+50, +40, +30, +20, +10],
        growth_trajectory=[0.90, 0.85, 0.82, 0.80, 0.80],
        margin_trajectory=[-0.005, -0.005, -0.004, -0.003, -0.003],
    ),
    "rate_hike": Scenario(
        name="Rising Rates",
        description="Stable sales, aggressive rate hikes (+200bp): Fed fights inflation",
        revenue_growth_adj=0.0,
        margin_adj=-0.005,
        wacc_adj=0.02,
        terminal_growth_adj=0.0,
        rate_path_bp=[+75, +50, +50, +25, 0],
        growth_trajectory=[1.0, 0.98, 0.95, 0.93, 0.90],
        margin_trajectory=[0.0, -0.003, -0.005, -0.003, 0.0],
    ),
    "rate_cut": Scenario(
        name="Falling Rates",
        description="Stable sales, rate cuts (-150bp): Fed easing to support growth",
        revenue_growth_adj=0.0,
        margin_adj=0.005,
        wacc_adj=-0.015,
        terminal_growth_adj=0.0,
        rate_path_bp=[-50, -50, -25, -15, -10],
        growth_trajectory=[1.0, 1.02, 1.03, 1.02, 1.00],
        margin_trajectory=[0.0, 0.002, 0.003, 0.002, 0.0],
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

    Uses year-by-year growth and margin trajectories from the scenario
    definition for more realistic modeling of how economic conditions
    evolve over time (e.g., gradual rate hikes, margin expansion).

    Approach:
      1. Project revenue using historical growth + scenario trajectory
      2. Apply operating margin (with year-by-year adjustments) to get EBIT
      3. Compute NOPAT = EBIT * (1 - tax_rate)
      4. Add back D&A, subtract CapEx
      5. Result = Unlevered Free Cash Flow
    """
    base_rev = financials["revenue"][0]
    base_growth = metrics["avg_revenue_growth"]
    base_op_margin = metrics["avg_operating_margin"]
    base_capex_pct = metrics["avg_capex_pct"]
    base_da_pct = metrics["avg_da_pct"]

    tax_prov = abs(financials["tax_provision"][0]) if financials["tax_provision"][0] else 0
    pretax = abs(financials["net_income"][0]) + tax_prov
    tax_rate = tax_prov / pretax if pretax > 0 else 0.21

    adj_growth = max(base_growth + scenario.revenue_growth_adj, -0.15)
    adj_margin = max(base_op_margin + scenario.margin_adj, 0.01)

    growth_traj = scenario.growth_trajectory or [1.0] * projection_years
    margin_traj = scenario.margin_trajectory or [0.0] * projection_years
    while len(growth_traj) < projection_years:
        growth_traj.append(growth_traj[-1] if growth_traj else 1.0)
    while len(margin_traj) < projection_years:
        margin_traj.append(margin_traj[-1] if margin_traj else 0.0)

    projected_revenue = []
    projected_ebit = []
    projected_nopat = []
    projected_da = []
    projected_capex = []
    projected_fcf = []
    growth_rates = []
    margins = []

    cumulative_rev = base_rev
    for yr in range(projection_years):
        growth = adj_growth * growth_traj[yr]
        growth = max(growth, -0.15)
        if adj_growth > 0:
            growth = max(growth, 0.005)

        year_margin = max(adj_margin + margin_traj[yr], 0.01)

        cumulative_rev = cumulative_rev * (1 + growth)
        ebit = cumulative_rev * year_margin
        nopat = ebit * (1 - tax_rate)
        da = cumulative_rev * base_da_pct
        capex = cumulative_rev * base_capex_pct
        fcf = nopat + da - capex

        projected_revenue.append(cumulative_rev)
        projected_ebit.append(ebit)
        projected_nopat.append(nopat)
        projected_da.append(da)
        projected_capex.append(capex)
        projected_fcf.append(fcf)
        growth_rates.append(growth)
        margins.append(year_margin)

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

    Uses year-by-year WACC adjustments from the scenario's rate_path_bp
    to model how changing interest rates affect the discount rate over time.

    Terminal Value = FCF_n * (1 + g) / (WACC_terminal - g)   [Gordon Growth Model]
    """
    base_wacc = wacc_data["wacc"]
    rate_path = scenario.rate_path_bp or [0] * len(projection["projected_fcf"])
    while len(rate_path) < len(projection["projected_fcf"]):
        rate_path.append(rate_path[-1] if rate_path else 0)

    fcfs = projection["projected_fcf"]
    n = len(fcfs)

    # Year-by-year WACC with cumulative rate path
    yearly_waccs = []
    cumulative_bp = 0
    for i in range(n):
        cumulative_bp += rate_path[i]
        yr_wacc = max(base_wacc + cumulative_bp / 10000.0, 0.04)
        yearly_waccs.append(yr_wacc)

    # Terminal WACC = base + total adjustment
    terminal_wacc = max(base_wacc + scenario.wacc_adj, 0.04)

    # Average WACC for summary display
    avg_wacc = sum(yearly_waccs) / len(yearly_waccs) if yearly_waccs else terminal_wacc

    tg = terminal_growth + scenario.terminal_growth_adj
    tg = min(tg, terminal_wacc - 0.01)
    tg = max(tg, 0.005)

    # Present value of projected FCFs using cumulative discount
    pv_fcfs = []
    cumulative_discount = 1.0
    for i, fcf in enumerate(fcfs):
        cumulative_discount *= (1 + yearly_waccs[i])
        pv = fcf / cumulative_discount
        pv_fcfs.append(pv)

    pv_fcf_total = sum(pv_fcfs)

    # Terminal value (discounted at terminal WACC)
    terminal_fcf = fcfs[-1] * (1 + tg)
    terminal_value = terminal_fcf / (terminal_wacc - tg)
    pv_terminal = terminal_value / cumulative_discount

    enterprise_value = pv_fcf_total + pv_terminal

    net_debt = financials["total_debt"][0] - financials["cash"][0]
    equity_value = enterprise_value - net_debt

    shares = stock["shares_outstanding"]
    value_per_share = equity_value / shares if shares > 0 else 0

    current_price = stock["current_price"]
    upside = (value_per_share - current_price) / current_price if current_price > 0 else 0

    return {
        "scenario": scenario.name,
        "scenario_description": scenario.description,
        "wacc": avg_wacc,
        "terminal_wacc": terminal_wacc,
        "yearly_waccs": yearly_waccs,
        "rate_path_bp": rate_path[:n],
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
