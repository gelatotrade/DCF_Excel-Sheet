#!/usr/bin/env python3
"""
generate_screenshots.py
=======================
Generates PNG screenshots of each sheet + chart previews from the DCF model.
Uses matplotlib + Pillow to render professional-looking images from the actual data.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.patches import FancyBboxPatch
import numpy as np

from data_fetcher import fetch_all
from dcf_engine import run_all_scenarios, SCENARIOS

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs", "screenshots")
os.makedirs(OUT_DIR, exist_ok=True)

# ── Style constants ──────────────────────────────────────────────────────────
DARK_BG = "#1a1a2e"
CARD_BG = "#16213e"
HEADER_BG = "#0f3460"
ACCENT = "#e94560"
ACCENT2 = "#00b4d8"
ACCENT3 = "#2dc653"
ACCENT4 = "#ffd60a"
ACCENT5 = "#9d4edd"
TEXT_WHITE = "#f0f0f0"
TEXT_MUTED = "#8899aa"
GRID_COLOR = "#2a3a5e"

SCENARIO_COLORS = {
    "base": "#4cc9f0",
    "bull": "#2dc653",
    "bear": "#e94560",
    "rate_hike": "#ffd60a",
    "rate_cut": "#9d4edd",
}
SCENARIO_ORDER = ["bull", "base", "bear", "rate_hike", "rate_cut"]


def setup_style():
    plt.rcParams.update({
        "figure.facecolor": DARK_BG,
        "axes.facecolor": CARD_BG,
        "axes.edgecolor": GRID_COLOR,
        "axes.labelcolor": TEXT_WHITE,
        "axes.grid": True,
        "grid.color": GRID_COLOR,
        "grid.alpha": 0.4,
        "text.color": TEXT_WHITE,
        "xtick.color": TEXT_MUTED,
        "ytick.color": TEXT_MUTED,
        "font.family": "monospace",
        "font.size": 11,
        "legend.facecolor": CARD_BG,
        "legend.edgecolor": GRID_COLOR,
    })


def fmt_billions(x):
    return f"${x/1e9:.1f}B" if abs(x) >= 1e9 else f"${x/1e6:.0f}M"


def fmt_pct(x):
    return f"{x:.1%}"


def save(fig, name):
    path = os.path.join(OUT_DIR, name)
    fig.savefig(path, dpi=150, bbox_inches="tight", pad_inches=0.3)
    plt.close(fig)
    print(f"  ✓ {name}")


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 1: DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════════
def render_dashboard(stock, financials, model_result):
    fig = plt.figure(figsize=(20, 14))
    fig.suptitle(f"DCF VALUATION DASHBOARD — {stock['company_name']} ({stock['ticker']})",
                 fontsize=20, fontweight="bold", color=TEXT_WHITE, y=0.98)

    # Layout: 2x2 grid
    gs = fig.add_gridspec(2, 2, hspace=0.35, wspace=0.3, left=0.06, right=0.96, top=0.92, bottom=0.06)

    # ── Top-left: Revenue & FCF bars ──
    ax1 = fig.add_subplot(gs[0, 0])
    years = financials["years"][:4]
    rev = [r / 1e9 for r in financials["revenue"][:4]]
    fcf = [f / 1e9 for f in financials["free_cash_flow"][:4]]
    x = np.arange(len(years))
    w = 0.35
    ax1.bar(x - w/2, rev, w, color=ACCENT2, label="Revenue", edgecolor="none", alpha=0.9)
    ax1.bar(x + w/2, fcf, w, color=ACCENT3, label="Free Cash Flow", edgecolor="none", alpha=0.9)
    ax1.set_xticks(x)
    ax1.set_xticklabels(years)
    ax1.set_ylabel("$ Billions")
    ax1.set_title("Revenue & Free Cash Flow (Historical)", fontsize=13, fontweight="bold", pad=10)
    ax1.legend(loc="upper left", fontsize=9)
    ax1.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.0fB"))

    # ── Top-right: Scenario implied prices ──
    ax2 = fig.add_subplot(gs[0, 1])
    names = []
    prices = []
    colors = []
    for key in SCENARIO_ORDER:
        dcf = model_result["scenarios"][key]["dcf"]
        names.append(dcf["scenario"])
        prices.append(dcf["implied_share_price"])
        colors.append(SCENARIO_COLORS[key])
    bars = ax2.barh(names[::-1], prices[::-1], color=colors[::-1], edgecolor="none", height=0.6)
    ax2.axvline(stock["current_price"], color=ACCENT, linestyle="--", linewidth=2, label=f"Current: ${stock['current_price']:.0f}")
    ax2.set_xlabel("Implied Share Price ($)")
    ax2.set_title("Scenario Valuation Comparison", fontsize=13, fontweight="bold", pad=10)
    ax2.legend(loc="lower right", fontsize=9)
    for bar, price in zip(bars, prices[::-1]):
        ax2.text(bar.get_width() + 1.5, bar.get_y() + bar.get_height()/2,
                 f"${price:.2f}", va="center", fontsize=9, color=TEXT_WHITE)

    # ── Bottom-left: Margin trends ──
    ax3 = fig.add_subplot(gs[1, 0])
    metrics = model_result["metrics"]
    yrs = financials["years"][:len(metrics["operating_margins"])]
    ax3.plot(yrs, [m*100 for m in metrics["operating_margins"]], "-o",
             color=ACCENT2, label="Operating Margin", linewidth=2, markersize=6)
    ax3.plot(yrs, [m*100 for m in metrics["net_margins"]], "-s",
             color=ACCENT3, label="Net Margin", linewidth=2, markersize=6)
    ax3.plot(yrs, [m*100 for m in metrics["fcf_margins"]], "-^",
             color=ACCENT4, label="FCF Margin", linewidth=2, markersize=6)
    ax3.set_ylabel("Margin %")
    ax3.set_title("Profitability Margins (Historical)", fontsize=13, fontweight="bold", pad=10)
    ax3.legend(fontsize=9)
    ax3.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f%%"))

    # ── Bottom-right: WACC components ──
    ax4 = fig.add_subplot(gs[1, 1])
    wacc = model_result["wacc_data"]
    components = ["WACC", "Cost of\nEquity", "Cost of\nDebt", "Risk-Free\nRate"]
    values = [wacc["wacc"]*100, wacc["cost_of_equity"]*100,
              wacc["cost_of_debt"]*100, wacc["risk_free_rate"]*100]
    bar_colors = [ACCENT, ACCENT2, ACCENT3, ACCENT4]
    bars = ax4.bar(components, values, color=bar_colors, edgecolor="none", width=0.55)
    ax4.set_ylabel("Rate %")
    ax4.set_title("WACC & Rate Components", fontsize=13, fontweight="bold", pad=10)
    ax4.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))
    for bar, val in zip(bars, values):
        ax4.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.2,
                 f"{val:.2f}%", ha="center", fontsize=10, color=TEXT_WHITE, fontweight="bold")

    save(fig, "01_dashboard.png")


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 2: INCOME STATEMENT
# ═══════════════════════════════════════════════════════════════════════════════
def render_income_statement(financials):
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))
    fig.suptitle("INCOME STATEMENT", fontsize=18, fontweight="bold", color=TEXT_WHITE, y=0.98)

    years = financials["years"][:4]
    rev = [r / 1e9 for r in financials["revenue"][:4]]
    ni = [n / 1e9 for n in financials["net_income"][:4]]
    oi = [o / 1e9 for o in financials["operating_income"][:4]]

    # Revenue & income bars
    x = np.arange(len(years))
    w = 0.25
    ax1.bar(x - w, rev, w, color=ACCENT2, label="Revenue", alpha=0.9)
    ax1.bar(x, oi, w, color=ACCENT4, label="Operating Income", alpha=0.9)
    ax1.bar(x + w, ni, w, color=ACCENT3, label="Net Income", alpha=0.9)
    ax1.set_xticks(x)
    ax1.set_xticklabels(years)
    ax1.set_ylabel("$ Billions")
    ax1.set_title("Revenue, Operating & Net Income", fontsize=13, fontweight="bold", pad=10)
    ax1.legend(fontsize=9)
    ax1.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.0fB"))

    # Margin trends
    op_margins = [oi[i] / rev[i] * 100 if rev[i] else 0 for i in range(len(years))]
    net_margins = [ni[i] / rev[i] * 100 if rev[i] else 0 for i in range(len(years))]
    ax2.plot(years, op_margins, "-o", color=ACCENT2, label="Operating Margin", linewidth=2.5, markersize=8)
    ax2.plot(years, net_margins, "-s", color=ACCENT3, label="Net Margin", linewidth=2.5, markersize=8)
    ax2.fill_between(years, op_margins, alpha=0.1, color=ACCENT2)
    ax2.fill_between(years, net_margins, alpha=0.1, color=ACCENT3)
    ax2.set_ylabel("Margin %")
    ax2.set_title("Margin Trends", fontsize=13, fontweight="bold", pad=10)
    ax2.legend(fontsize=9)
    ax2.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f%%"))

    save(fig, "02_income_statement.png")


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 3: BALANCE SHEET
# ═══════════════════════════════════════════════════════════════════════════════
def render_balance_sheet(financials):
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))
    fig.suptitle("BALANCE SHEET", fontsize=18, fontweight="bold", color=TEXT_WHITE, y=0.98)

    years = financials["years"][:4]
    ta = [a / 1e9 for a in financials["total_assets"][:4]]
    tl = [l / 1e9 for l in financials["total_liabilities"][:4]]
    eq = [e / 1e9 for e in financials["total_equity"][:4]]
    cash = [c / 1e9 for c in financials["cash"][:4]]
    debt = [d / 1e9 for d in financials["total_debt"][:4]]

    # Stacked structure
    x = np.arange(len(years))
    w = 0.35
    ax1.bar(x - w/2, ta, w, color=ACCENT2, label="Total Assets", alpha=0.9)
    ax1.bar(x + w/2, tl, w, color=ACCENT, label="Total Liabilities", alpha=0.9)
    ax1.bar(x + w/2, eq, w, bottom=tl, color=ACCENT3, label="Equity", alpha=0.9)
    ax1.set_xticks(x)
    ax1.set_xticklabels(years)
    ax1.set_ylabel("$ Billions")
    ax1.set_title("Assets vs Liabilities + Equity", fontsize=13, fontweight="bold", pad=10)
    ax1.legend(fontsize=8)
    ax1.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.0fB"))

    # D/E Ratio & Cash/Debt
    de_ratio = [debt[i] / eq[i] if eq[i] else 0 for i in range(len(years))]
    ax2.plot(years, de_ratio, "-o", color=ACCENT, label="Debt / Equity", linewidth=2.5, markersize=8)
    ax2b = ax2.twinx()
    ax2b.bar(years, cash, width=0.3, color=ACCENT3, alpha=0.5, label="Cash")
    ax2b.bar(years, [-d for d in debt], width=0.3, color=ACCENT, alpha=0.3, label="Debt (neg)")
    ax2.set_title("Debt/Equity Ratio & Cash vs Debt", fontsize=13, fontweight="bold", pad=10)
    ax2.set_ylabel("D/E Ratio")
    ax2b.set_ylabel("$ Billions")
    ax2.legend(loc="upper left", fontsize=9)
    ax2b.legend(loc="upper right", fontsize=9)
    ax2b.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.0fB"))

    save(fig, "03_balance_sheet.png")


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 4: CASH FLOW
# ═══════════════════════════════════════════════════════════════════════════════
def render_cash_flow(financials):
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 8))
    fig.suptitle("CASH FLOW STATEMENT", fontsize=18, fontweight="bold", color=TEXT_WHITE, y=0.98)

    years = financials["years"][:4]
    ocf = [o / 1e9 for o in financials["operating_cash_flow"][:4]]
    fcf = [f / 1e9 for f in financials["free_cash_flow"][:4]]
    capex = [abs(c) / 1e9 for c in financials["capex"][:4]]

    x = np.arange(len(years))
    w = 0.25
    ax1.bar(x - w, ocf, w, color=ACCENT2, label="Operating CF", alpha=0.9)
    ax1.bar(x, fcf, w, color=ACCENT3, label="Free CF", alpha=0.9)
    ax1.bar(x + w, capex, w, color=ACCENT, label="CapEx", alpha=0.9)
    ax1.set_xticks(x)
    ax1.set_xticklabels(years)
    ax1.set_ylabel("$ Billions")
    ax1.set_title("Operating CF, Free CF & CapEx", fontsize=13, fontweight="bold", pad=10)
    ax1.legend(fontsize=9)
    ax1.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.0fB"))

    # FCF Yield
    rev = [r / 1e9 for r in financials["revenue"][:4]]
    fcf_yield = [fcf[i] / rev[i] * 100 if rev[i] else 0 for i in range(len(years))]
    capex_pct = [capex[i] / rev[i] * 100 if rev[i] else 0 for i in range(len(years))]
    ax2.plot(years, fcf_yield, "-o", color=ACCENT3, label="FCF Yield (% of Rev)", linewidth=2.5, markersize=8)
    ax2.plot(years, capex_pct, "-s", color=ACCENT, label="CapEx (% of Rev)", linewidth=2.5, markersize=8)
    ax2.fill_between(years, fcf_yield, alpha=0.15, color=ACCENT3)
    ax2.fill_between(years, capex_pct, alpha=0.15, color=ACCENT)
    ax2.set_ylabel("% of Revenue")
    ax2.set_title("FCF Yield & CapEx Intensity", fontsize=13, fontweight="bold", pad=10)
    ax2.legend(fontsize=9)
    ax2.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.0f%%"))

    save(fig, "04_cash_flow.png")


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 5: WACC
# ═══════════════════════════════════════════════════════════════════════════════
def render_wacc(model_result, stock, financials):
    fig = plt.figure(figsize=(20, 8))
    fig.suptitle("WACC — WEIGHTED AVERAGE COST OF CAPITAL",
                 fontsize=18, fontweight="bold", color=TEXT_WHITE, y=0.98)

    gs = fig.add_gridspec(1, 3, wspace=0.35, left=0.06, right=0.96, top=0.88, bottom=0.10)
    wacc = model_result["wacc_data"]

    # ── Pie: Capital Structure ──
    ax1 = fig.add_subplot(gs[0, 0])
    we = wacc["weight_equity"] * 100
    wd = wacc["weight_debt"] * 100
    wedges, texts, autotexts = ax1.pie(
        [we, wd], labels=["Equity", "Debt"],
        colors=[ACCENT2, ACCENT],
        autopct="%1.1f%%", startangle=90,
        textprops={"color": TEXT_WHITE, "fontsize": 12},
        wedgeprops={"edgecolor": DARK_BG, "linewidth": 2}
    )
    for t in autotexts:
        t.set_fontweight("bold")
    ax1.set_title("Capital Structure", fontsize=13, fontweight="bold", pad=15)

    # ── Bar: Component Rates ──
    ax2 = fig.add_subplot(gs[0, 1])
    labels = ["WACC", "Cost of Equity\n(CAPM)", "Cost of Debt\n(after-tax)", "Risk-Free\nRate"]
    vals = [wacc["wacc"]*100, wacc["cost_of_equity"]*100,
            wacc["cost_of_debt"]*(1-wacc["tax_rate"])*100, wacc["risk_free_rate"]*100]
    bar_colors = [ACCENT, ACCENT2, ACCENT3, ACCENT4]
    bars = ax2.bar(labels, vals, color=bar_colors, edgecolor="none", width=0.55)
    ax2.set_ylabel("Rate (%)")
    ax2.set_title("WACC Component Rates", fontsize=13, fontweight="bold", pad=10)
    ax2.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))
    for bar, val in zip(bars, vals):
        ax2.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.15,
                 f"{val:.2f}%", ha="center", fontsize=10, color=TEXT_WHITE, fontweight="bold")

    # ── Bar: CAPM Breakdown ──
    ax3 = fig.add_subplot(gs[0, 2])
    rf = wacc["risk_free_rate"] * 100
    beta_erp = wacc["beta"] * wacc["equity_risk_premium"] * 100
    ax3.bar(["Risk-Free Rate", f"Beta × ERP\n({wacc['beta']:.2f} × {wacc['equity_risk_premium']:.1%})"],
            [rf, beta_erp], color=[ACCENT4, ACCENT5], edgecolor="none", width=0.45)
    ax3.axhline(wacc["cost_of_equity"]*100, color=ACCENT2, linestyle="--", linewidth=1.5,
                label=f"Cost of Equity = {wacc['cost_of_equity']:.2%}")
    ax3.set_ylabel("Rate (%)")
    ax3.set_title("CAPM Breakdown", fontsize=13, fontweight="bold", pad=10)
    ax3.legend(fontsize=9)
    ax3.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))

    save(fig, "05_wacc.png")


# ═══════════════════════════════════════════════════════════════════════════════
# SHEETS 6-10: DCF SCENARIOS
# ═══════════════════════════════════════════════════════════════════════════════
def render_dcf_scenario(key, idx, model_result, stock, financials):
    sc = model_result["scenarios"][key]
    scenario = sc["scenario"]
    proj = sc["projection"]
    dcf = sc["dcf"]
    color = SCENARIO_COLORS[key]

    fig = plt.figure(figsize=(20, 10))
    fig.suptitle(f"DCF SCENARIO — {scenario.name.upper()}: {scenario.description}",
                 fontsize=16, fontweight="bold", color=color, y=0.98)

    gs = fig.add_gridspec(2, 2, hspace=0.35, wspace=0.3, left=0.06, right=0.96, top=0.90, bottom=0.08)

    proj_years = [f"Year {i+1}" for i in range(proj["projection_years"])]

    # ── Revenue Projection ──
    ax1 = fig.add_subplot(gs[0, 0])
    rev_proj = [r / 1e9 for r in proj["projected_revenue"]]
    hist_rev = financials["revenue"][0] / 1e9
    all_rev = [hist_rev] + rev_proj
    all_labels = ["Current"] + proj_years
    clrs = [TEXT_MUTED] + [color] * len(proj_years)
    ax1.bar(all_labels, all_rev, color=clrs, edgecolor="none", alpha=0.85)
    ax1.set_ylabel("$ Billions")
    ax1.set_title("Revenue Projection", fontsize=13, fontweight="bold", pad=10)
    ax1.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.0fB"))
    ax1.tick_params(axis="x", rotation=30)

    # ── FCF vs PV of FCF ──
    ax2 = fig.add_subplot(gs[0, 1])
    fcf_proj = [f / 1e9 for f in proj["projected_fcf"]]
    pv_fcfs = [p / 1e9 for p in dcf["pv_fcfs"]]
    x = np.arange(len(proj_years))
    w = 0.35
    ax2.bar(x - w/2, fcf_proj, w, color=color, label="Projected FCF", alpha=0.9)
    ax2.bar(x + w/2, pv_fcfs, w, color=ACCENT2, label="PV of FCF", alpha=0.9)
    ax2.set_xticks(x)
    ax2.set_xticklabels(proj_years)
    ax2.set_ylabel("$ Billions")
    ax2.set_title("FCF vs Present Value", fontsize=13, fontweight="bold", pad=10)
    ax2.legend(fontsize=9)
    ax2.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.0fB"))

    # ── Valuation Bridge ──
    ax3 = fig.add_subplot(gs[1, 0])
    bridge_labels = ["PV of FCFs", "PV Terminal\nValue", "Enterprise\nValue", "- Net Debt", "Equity\nValue"]
    bridge_vals = [
        dcf["pv_fcf_total"] / 1e9,
        dcf["pv_terminal_value"] / 1e9,
        dcf["enterprise_value"] / 1e9,
        -dcf["net_debt"] / 1e9,
        dcf["equity_value"] / 1e9,
    ]
    bridge_colors = [ACCENT2, ACCENT5, color, ACCENT, ACCENT3]
    bars = ax3.bar(bridge_labels, bridge_vals, color=bridge_colors, edgecolor="none", width=0.55)
    ax3.set_ylabel("$ Billions")
    ax3.set_title("Valuation Bridge", fontsize=13, fontweight="bold", pad=10)
    ax3.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.0fB"))
    for bar, val in zip(bars, bridge_vals):
        y_pos = bar.get_height() if val >= 0 else bar.get_height() - 0.5
        ax3.text(bar.get_x() + bar.get_width()/2, y_pos + 0.2,
                 f"${val:.1f}B", ha="center", fontsize=9, color=TEXT_WHITE, fontweight="bold")

    # ── Key Metrics Summary ──
    ax4 = fig.add_subplot(gs[1, 1])
    ax4.axis("off")
    metrics_text = [
        ("WACC", f"{dcf['wacc']:.2%}"),
        ("Terminal Growth", f"{dcf['terminal_growth']:.2%}"),
        ("Enterprise Value", fmt_billions(dcf["enterprise_value"])),
        ("Net Debt", fmt_billions(dcf["net_debt"])),
        ("Equity Value", fmt_billions(dcf["equity_value"])),
        ("Shares Outstanding", f"{dcf['shares_outstanding']/1e9:.2f}B"),
        ("", ""),
        ("Implied Price", f"${dcf['implied_share_price']:.2f}"),
        ("Current Price", f"${dcf['current_price']:.2f}"),
        ("Upside / Downside", f"{dcf['upside_downside']:.1%}"),
    ]
    y_start = 0.95
    for i, (label, value) in enumerate(metrics_text):
        if not label:
            continue
        y = y_start - i * 0.09
        is_result = label in ("Implied Price", "Current Price", "Upside / Downside")
        fsize = 14 if is_result else 12
        fweight = "bold" if is_result else "normal"
        lcolor = color if is_result else TEXT_MUTED
        ax4.text(0.05, y, label, fontsize=fsize, color=lcolor, fontweight=fweight,
                 transform=ax4.transAxes, va="top")
        ax4.text(0.95, y, value, fontsize=fsize, color=TEXT_WHITE, fontweight=fweight,
                 transform=ax4.transAxes, va="top", ha="right")

    ax4.set_title("Key Metrics", fontsize=13, fontweight="bold", pad=10, color=TEXT_WHITE)
    # Border
    ax4.add_patch(FancyBboxPatch((0.01, 0.01), 0.98, 0.98, boxstyle="round,pad=0.02",
                                  facecolor=CARD_BG, edgecolor=color, linewidth=2,
                                  transform=ax4.transAxes))

    save(fig, f"{idx:02d}_dcf_{key}.png")


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 11: SCENARIO COMPARISON
# ═══════════════════════════════════════════════════════════════════════════════
def render_scenario_comparison(model_result, stock, financials):
    fig = plt.figure(figsize=(22, 14))
    fig.suptitle("SCENARIO COMPARISON — ALL 5 CASES",
                 fontsize=18, fontweight="bold", color=TEXT_WHITE, y=0.98)

    gs = fig.add_gridspec(2, 3, hspace=0.35, wspace=0.3, left=0.05, right=0.97, top=0.92, bottom=0.06)

    # ── Implied Price Comparison ──
    ax1 = fig.add_subplot(gs[0, 0])
    names, prices, colors = [], [], []
    for key in SCENARIO_ORDER:
        dcf = model_result["scenarios"][key]["dcf"]
        names.append(dcf["scenario"])
        prices.append(dcf["implied_share_price"])
        colors.append(SCENARIO_COLORS[key])
    bars = ax1.bar(names, prices, color=colors, edgecolor="none", width=0.6)
    ax1.axhline(stock["current_price"], color=ACCENT, linestyle="--", linewidth=2,
                label=f"Current: ${stock['current_price']:.0f}")
    ax1.set_title("Implied Share Price by Scenario", fontsize=12, fontweight="bold", pad=10)
    ax1.set_ylabel("Price ($)")
    ax1.legend(fontsize=9)
    ax1.tick_params(axis="x", rotation=25, labelsize=9)
    for bar, p in zip(bars, prices):
        ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.8,
                 f"${p:.0f}", ha="center", fontsize=9, color=TEXT_WHITE, fontweight="bold")

    # ── EV Breakdown ──
    ax2 = fig.add_subplot(gs[0, 1])
    pv_fcf_vals, pv_tv_vals = [], []
    for key in SCENARIO_ORDER:
        dcf = model_result["scenarios"][key]["dcf"]
        pv_fcf_vals.append(dcf["pv_fcf_total"] / 1e9)
        pv_tv_vals.append(dcf["pv_terminal_value"] / 1e9)
    x = np.arange(len(SCENARIO_ORDER))
    ax2.bar(x, pv_fcf_vals, 0.6, color=ACCENT2, label="PV of FCFs", alpha=0.9)
    ax2.bar(x, pv_tv_vals, 0.6, bottom=pv_fcf_vals, color=ACCENT5, label="PV of Terminal Value", alpha=0.9)
    ax2.set_xticks(x)
    ax2.set_xticklabels([model_result["scenarios"][k]["dcf"]["scenario"] for k in SCENARIO_ORDER],
                        rotation=25, fontsize=9)
    ax2.set_title("Enterprise Value Breakdown ($B)", fontsize=12, fontweight="bold", pad=10)
    ax2.set_ylabel("$ Billions")
    ax2.legend(fontsize=9)

    # ── WACC vs Terminal Growth ──
    ax3 = fig.add_subplot(gs[0, 2])
    waccs, tgs = [], []
    for key in SCENARIO_ORDER:
        dcf = model_result["scenarios"][key]["dcf"]
        waccs.append(dcf["wacc"] * 100)
        tgs.append(dcf["terminal_growth"] * 100)
    x = np.arange(len(SCENARIO_ORDER))
    w = 0.3
    ax3.bar(x - w/2, waccs, w, color=ACCENT, label="WACC", alpha=0.9)
    ax3.bar(x + w/2, tgs, w, color=ACCENT3, label="Terminal Growth", alpha=0.9)
    ax3.set_xticks(x)
    ax3.set_xticklabels([model_result["scenarios"][k]["dcf"]["scenario"] for k in SCENARIO_ORDER],
                        rotation=25, fontsize=9)
    ax3.set_title("WACC vs Terminal Growth Rate", fontsize=12, fontweight="bold", pad=10)
    ax3.set_ylabel("Rate (%)")
    ax3.legend(fontsize=9)
    ax3.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))

    # ── FCF Projection Comparison (Line) ──
    ax4 = fig.add_subplot(gs[1, 0:2])
    for key in SCENARIO_ORDER:
        proj = model_result["scenarios"][key]["projection"]
        yrs = [f"Year {i+1}" for i in range(proj["projection_years"])]
        fcfs = [f / 1e9 for f in proj["projected_fcf"]]
        ax4.plot(yrs, fcfs, "-o", color=SCENARIO_COLORS[key],
                 label=model_result["scenarios"][key]["dcf"]["scenario"],
                 linewidth=2.5, markersize=7)
    ax4.set_title("Projected FCF by Scenario ($B)", fontsize=12, fontweight="bold", pad=10)
    ax4.set_ylabel("$ Billions")
    ax4.legend(fontsize=9, ncol=3)
    ax4.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.0fB"))

    # ── Revenue Projection Comparison ──
    ax5 = fig.add_subplot(gs[1, 2])
    for key in SCENARIO_ORDER:
        proj = model_result["scenarios"][key]["projection"]
        yrs = [f"Y{i+1}" for i in range(proj["projection_years"])]
        revs = [r / 1e9 for r in proj["projected_revenue"]]
        ax5.plot(yrs, revs, "-s", color=SCENARIO_COLORS[key],
                 label=model_result["scenarios"][key]["dcf"]["scenario"],
                 linewidth=2, markersize=6)
    ax5.set_title("Projected Revenue ($B)", fontsize=12, fontweight="bold", pad=10)
    ax5.set_ylabel("$ Billions")
    ax5.legend(fontsize=8)
    ax5.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.0fB"))

    save(fig, "11_scenario_comparison.png")


# ═══════════════════════════════════════════════════════════════════════════════
# SHEET 12: SENSITIVITY ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════
def render_sensitivity(model_result, stock, financials):
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(20, 9))
    fig.suptitle("SENSITIVITY ANALYSIS", fontsize=18, fontweight="bold", color=TEXT_WHITE, y=0.98)

    base_dcf = model_result["scenarios"]["base"]["dcf"]
    base_wacc = base_dcf["wacc"]
    base_tg = base_dcf["terminal_growth"]
    base_fcf = model_result["scenarios"]["base"]["projection"]["projected_fcf"]
    shares = stock["shares_outstanding"]
    net_debt = financials["total_debt"][0] - financials["cash"][0]
    n = len(base_fcf)

    # ── WACC Sensitivity ──
    wacc_range = np.arange(max(base_wacc - 0.04, 0.04), base_wacc + 0.05, 0.005)
    wacc_prices = []
    for w in wacc_range:
        tg = min(base_tg, w - 0.01)
        pv = sum(fcf / (1 + w) ** (i + 1) for i, fcf in enumerate(base_fcf))
        tv = base_fcf[-1] * (1 + tg) / (w - tg) / (1 + w) ** n
        eq = pv + tv - net_debt
        wacc_prices.append(eq / shares if shares else 0)

    ax1.plot([w*100 for w in wacc_range], wacc_prices, "-o", color=ACCENT2, linewidth=2.5, markersize=5)
    ax1.axvline(base_wacc*100, color=ACCENT, linestyle="--", linewidth=1.5,
                label=f"Base WACC ({base_wacc:.1%})")
    ax1.axhline(stock["current_price"], color=ACCENT4, linestyle=":", linewidth=1.5,
                label=f"Current Price (${stock['current_price']:.0f})")
    ax1.fill_between([w*100 for w in wacc_range], wacc_prices, alpha=0.1, color=ACCENT2)
    ax1.set_xlabel("WACC (%)")
    ax1.set_ylabel("Implied Share Price ($)")
    ax1.set_title("Impact of WACC on Valuation", fontsize=13, fontweight="bold", pad=10)
    ax1.legend(fontsize=9)
    ax1.xaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))

    # ── Terminal Growth Sensitivity ──
    tg_range = np.arange(0.005, min(base_wacc - 0.01, 0.055), 0.0025)
    tg_prices = []
    for tg in tg_range:
        pv = sum(fcf / (1 + base_wacc) ** (i + 1) for i, fcf in enumerate(base_fcf))
        tv = base_fcf[-1] * (1 + tg) / (base_wacc - tg) / (1 + base_wacc) ** n
        eq = pv + tv - net_debt
        tg_prices.append(eq / shares if shares else 0)

    ax2.plot([t*100 for t in tg_range], tg_prices, "-s", color=ACCENT3, linewidth=2.5, markersize=5)
    ax2.axvline(base_tg*100, color=ACCENT, linestyle="--", linewidth=1.5,
                label=f"Base TG ({base_tg:.1%})")
    ax2.axhline(stock["current_price"], color=ACCENT4, linestyle=":", linewidth=1.5,
                label=f"Current Price (${stock['current_price']:.0f})")
    ax2.fill_between([t*100 for t in tg_range], tg_prices, alpha=0.1, color=ACCENT3)
    ax2.set_xlabel("Terminal Growth Rate (%)")
    ax2.set_ylabel("Implied Share Price ($)")
    ax2.set_title("Impact of Terminal Growth on Valuation", fontsize=13, fontweight="bold", pad=10)
    ax2.legend(fontsize=9)
    ax2.xaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))

    save(fig, "12_sensitivity_analysis.png")


# ═══════════════════════════════════════════════════════════════════════════════
# CHART PREVIEWS (standalone chart images)
# ═══════════════════════════════════════════════════════════════════════════════
def render_chart_revenue_waterfall(financials, model_result):
    """Revenue waterfall: historical → projected across all scenarios."""
    fig, ax = plt.subplots(figsize=(16, 7))
    fig.suptitle("REVENUE WATERFALL — Historical to Projected",
                 fontsize=16, fontweight="bold", color=TEXT_WHITE, y=0.98)

    hist_years = financials["years"][:4][::-1]
    hist_rev = [r / 1e9 for r in financials["revenue"][:4][::-1]]

    # Plot historical
    x_all = list(hist_years)
    for key in ["bear", "base", "bull"]:
        proj = model_result["scenarios"][key]["projection"]
        proj_years = [f"P{i+1}" for i in range(proj["projection_years"])]
        proj_rev = [r / 1e9 for r in proj["projected_revenue"]]
        full_x = hist_years + proj_years
        full_y = hist_rev + proj_rev
        ax.plot(full_x, full_y, "-o", color=SCENARIO_COLORS[key],
                label=model_result["scenarios"][key]["dcf"]["scenario"],
                linewidth=2.5, markersize=7)
        if not x_all or len(full_x) > len(x_all):
            x_all = full_x

    # Shade projection zone
    ax.axvspan(len(hist_years) - 0.5, len(x_all) - 0.5, alpha=0.08, color=TEXT_WHITE)
    ax.axvline(len(hist_years) - 0.5, color=TEXT_MUTED, linestyle=":", linewidth=1)
    ax.text(len(hist_years) + 0.5, ax.get_ylim()[1] * 0.95, "PROJECTED →",
            fontsize=11, color=TEXT_MUTED, fontstyle="italic")

    ax.set_ylabel("Revenue ($ Billions)", fontsize=12)
    ax.set_title("Historical & Projected Revenue Across Scenarios", fontsize=14, fontweight="bold", pad=10)
    ax.legend(fontsize=10, loc="upper left")
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("$%.0fB"))

    save(fig, "chart_revenue_waterfall.png")


def render_chart_valuation_range(model_result, stock):
    """Valuation range chart showing min/max/current."""
    fig, ax = plt.subplots(figsize=(14, 6))
    fig.suptitle("VALUATION RANGE ACROSS SCENARIOS",
                 fontsize=16, fontweight="bold", color=TEXT_WHITE, y=0.98)

    prices = {}
    for key in SCENARIO_ORDER:
        dcf = model_result["scenarios"][key]["dcf"]
        prices[key] = dcf["implied_share_price"]

    min_price = min(prices.values())
    max_price = max(prices.values())
    current = stock["current_price"]

    # Range bar
    ax.barh(["Valuation\nRange"], [max_price - min_price], left=[min_price],
            height=0.4, color=ACCENT2, alpha=0.3, edgecolor=ACCENT2, linewidth=2)

    # Individual scenario markers
    for key in SCENARIO_ORDER:
        ax.plot(prices[key], 0, "D", color=SCENARIO_COLORS[key], markersize=14, zorder=5)
        offset = 0.25 if key in ("bull", "base", "rate_cut") else -0.25
        ax.annotate(f"{model_result['scenarios'][key]['dcf']['scenario']}\n${prices[key]:.0f}",
                    xy=(prices[key], 0), xytext=(prices[key], offset),
                    fontsize=9, color=SCENARIO_COLORS[key], fontweight="bold",
                    ha="center", va="center")

    # Current price line
    ax.axvline(current, color=ACCENT, linestyle="--", linewidth=2.5, zorder=10)
    ax.text(current, 0.42, f"Current: ${current:.0f}", fontsize=11, color=ACCENT,
            fontweight="bold", ha="center")

    ax.set_xlabel("Share Price ($)", fontsize=12)
    ax.set_xlim(min_price * 0.8, max(max_price, current) * 1.15)
    ax.set_ylim(-0.6, 0.6)
    ax.get_yaxis().set_visible(False)
    ax.set_title("Where Does the Current Price Sit vs DCF Implied Values?",
                 fontsize=13, fontweight="bold", pad=15)

    save(fig, "chart_valuation_range.png")


def render_chart_interest_rate_impact(model_result):
    """Show how interest rate changes impact valuation."""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(18, 7))
    fig.suptitle("INTEREST RATE IMPACT ON VALUATION",
                 fontsize=16, fontweight="bold", color=TEXT_WHITE, y=0.98)

    # WACC by scenario
    scenarios_to_show = ["rate_cut", "bull", "base", "bear", "rate_hike"]
    names = [model_result["scenarios"][k]["dcf"]["scenario"] for k in scenarios_to_show]
    waccs = [model_result["scenarios"][k]["dcf"]["wacc"] * 100 for k in scenarios_to_show]
    prices = [model_result["scenarios"][k]["dcf"]["implied_share_price"] for k in scenarios_to_show]
    clrs = [SCENARIO_COLORS[k] for k in scenarios_to_show]

    ax1.bar(names, waccs, color=clrs, edgecolor="none", width=0.6)
    ax1.set_title("WACC Across Rate Scenarios", fontsize=13, fontweight="bold", pad=10)
    ax1.set_ylabel("WACC (%)")
    ax1.tick_params(axis="x", rotation=25, labelsize=9)
    ax1.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.1f%%"))
    for i, (name, wacc) in enumerate(zip(names, waccs)):
        ax1.text(i, wacc + 0.1, f"{wacc:.2f}%", ha="center", fontsize=10,
                 color=TEXT_WHITE, fontweight="bold")

    # Price impact
    ax2.bar(names, prices, color=clrs, edgecolor="none", width=0.6)
    ax2.set_title("Implied Price Across Rate Scenarios", fontsize=13, fontweight="bold", pad=10)
    ax2.set_ylabel("Implied Share Price ($)")
    ax2.tick_params(axis="x", rotation=25, labelsize=9)
    for i, (name, price) in enumerate(zip(names, prices)):
        ax2.text(i, price + 0.5, f"${price:.0f}", ha="center", fontsize=10,
                 color=TEXT_WHITE, fontweight="bold")

    # Arrow annotation showing inverse relationship
    ax2.annotate("Lower rates → Higher value",
                 xy=(0, prices[0]), xytext=(1.5, prices[0] * 1.15),
                 fontsize=10, color=ACCENT3, fontweight="bold",
                 arrowprops=dict(arrowstyle="->", color=ACCENT3, lw=2))

    save(fig, "chart_interest_rate_impact.png")


# ═══════════════════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════════════════
def main():
    print("Generating DCF model screenshots...\n")

    data = fetch_all("AAPL", force_sample=True)
    stock = data["stock"]
    financials = data["financials"]
    rates = data["rates"]

    model_result = run_all_scenarios(stock, financials, rates)

    print("Sheet screenshots:")
    render_dashboard(stock, financials, model_result)
    render_income_statement(financials)
    render_balance_sheet(financials)
    render_cash_flow(financials)
    render_wacc(model_result, stock, financials)

    for i, key in enumerate(["base", "bull", "bear", "rate_hike", "rate_cut"], start=6):
        render_dcf_scenario(key, i, model_result, stock, financials)

    render_scenario_comparison(model_result, stock, financials)
    render_sensitivity(model_result, stock, financials)

    print("\nChart previews:")
    render_chart_revenue_waterfall(financials, model_result)
    render_chart_valuation_range(model_result, stock)
    render_chart_interest_rate_impact(model_result)

    print(f"\nAll images saved to: {OUT_DIR}/")
    print(f"Total: {len(os.listdir(OUT_DIR))} images")


if __name__ == "__main__":
    main()
