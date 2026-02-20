"""
Configuration for Stock Valuation Analyzer — Enhanced Professional Edition.

All assumptions are user-customizable. Set environment variables or edit defaults below.
"""
import os

# ---------------------------------------------------------------------------
# API Keys  –  set via environment variables or replace the defaults
# ---------------------------------------------------------------------------
ALPHA_VANTAGE_KEY = os.getenv("ALPHA_VANTAGE_KEY", "demo")
FMP_API_KEY = os.getenv("FMP_API_KEY", "demo")
FRED_API_KEY = os.getenv("FRED_API_KEY", "demo")
NEWS_API_KEY = os.getenv("NEWS_API_KEY", "")  # optional – newsapi.org

# ============================================================================
# VALUATION MODEL WEIGHTS (must sum to 1.0)
# ============================================================================
# Seasonal model removed from fair-value calculation (used for timing only).
# Replaced with Historical P/E valuation at 10% weight.
VALUATION_WEIGHTS = {
    "dcf": 0.50,
    "comps": 0.40,
    "historical": 0.10,
}

# ============================================================================
# WACC / CAPM INPUTS
# ============================================================================
RISK_FREE_RATE = None          # Auto-fetch 10-yr Treasury from FRED if None
MARKET_RISK_PREMIUM = 0.07     # 7% historical equity risk premium
BETA_PERIOD_YEARS = 5          # Years of data for beta calculation

# ============================================================================
# DCF ASSUMPTIONS
# ============================================================================
DCF_DEFAULTS = {
    "projection_years": 5,
    "terminal_growth_rate": 0.03,        # 3% perpetual growth (roughly GDP)
    "risk_free_rate": None,              # auto-fetched from FRED
    "equity_risk_premium": 0.07,         # 7% equity risk premium
    "tax_rate": 0.21,                    # US corporate
    "margin_of_safety": 0.20,            # 20% discount to fair value for "Buy"
}

# DCF Sensitivity ranges for the WACC vs Terminal Growth matrix
DCF_SENSITIVITY = {
    "wacc_steps": [-0.02, -0.01, 0.0, 0.01, 0.02],
    "tgr_steps": [-0.01, -0.005, 0.0, 0.005, 0.01],
}

# Revenue Growth vs EBIT Margin sensitivity
GROWTH_MARGIN_SENSITIVITY = {
    "growth_steps": [0.02, 0.04, 0.06, 0.08, 0.10],
    "margin_steps": [0.26, 0.28, 0.30, 0.32, 0.34],
}

# ============================================================================
# COMPARABLE COMPANIES
# ============================================================================
COMPS_COUNT = 8
COMPS_MULTIPLES = ["P/E", "EV/EBITDA", "P/S", "PEG", "P/B"]

# Default peer tickers for well-known stocks (fallback when API unavailable)
PEER_TICKERS = {
    "AAPL": ["MSFT", "GOOGL", "META", "NVDA", "AMZN"],
    "MSFT": ["AAPL", "GOOGL", "META", "AMZN", "CRM"],
    "GOOGL": ["META", "MSFT", "AMZN", "AAPL", "NFLX"],
    "META": ["GOOGL", "SNAP", "PINS", "MSFT", "AMZN"],
    "AMZN": ["MSFT", "GOOGL", "AAPL", "WMT", "SHOP"],
    "NVDA": ["AMD", "INTC", "AVGO", "QCOM", "TSM"],
    "TSLA": ["F", "GM", "RIVN", "NIO", "LI"],
    "JNJ": ["PFE", "MRK", "ABBV", "LLY", "UNH"],
    "NFLX": ["DIS", "CMCSA", "WBD", "PARA", "ROKU"],
}

# ============================================================================
# SEASONAL ANALYSIS
# ============================================================================
SEASONAL_YEARS_HISTORY = 10    # Years of historical data for seasonal patterns

# ============================================================================
# MARGIN OF SAFETY
# ============================================================================
MARGIN_OF_SAFETY = 0.20        # 20% discount to fair value required for "Buy"

# ============================================================================
# SCENARIO ANALYSIS
# ============================================================================
SCENARIO_WEIGHTS = {
    "bear": 0.15,
    "base": 0.65,
    "bull": 0.20,
}

# ============================================================================
# MACRO INDICATORS (FRED series IDs)
# ============================================================================
FRED_SERIES = {
    "Fed Funds Rate": "FEDFUNDS",
    "10Y Treasury": "DGS10",
    "2Y Treasury": "DGS2",
    "Unemployment Rate": "UNRATE",
    "CPI YoY": "CPIAUCSL",
    "Real GDP Growth": "A191RL1Q225SBEA",
    "ISM Manufacturing PMI": "MANEMP",
    "Consumer Confidence": "UMCSENT",
    "VIX": "VIXCLS",
}

# ============================================================================
# FINANCIAL RATIO BENCHMARKS
# ============================================================================
RATIO_BENCHMARKS = {
    # Liquidity
    "Current Ratio":      {"good": 1.5, "warn": 1.0},
    "Quick Ratio":        {"good": 1.0, "warn": 0.5},
    "Cash Ratio":         {"good": 0.5, "warn": 0.2},
    # Profitability
    "Gross Margin":       {"good": 0.40, "warn": 0.20},
    "Operating Margin":   {"good": 0.15, "warn": 0.05},
    "Net Margin":         {"good": 0.10, "warn": 0.0},
    "ROE":                {"good": 0.15, "warn": 0.08},
    "ROA":                {"good": 0.08, "warn": 0.03},
    "ROIC":               {"good": 0.12, "warn": 0.06},
    # Leverage
    "Debt / Equity":      {"good_below": 1.0, "warn_above": 2.0},
    "Debt / Assets":      {"good_below": 0.5, "warn_above": 0.7},
    "Interest Coverage":  {"good": 5.0, "warn": 2.0},
    # Efficiency
    "Asset Turnover":     {"good": 0.8, "warn": 0.3},
    # Valuation
    "PEG Ratio":          {"good_below": 1.0, "warn_above": 2.0},
    "FCF Yield":          {"good": 0.05, "warn": 0.02},
}

# Explanation text for key ratios
RATIO_EXPLANATIONS = {
    "Current Ratio": "Measures ability to pay short-term obligations. Some companies (like Apple) run low ratios due to strong cash generation.",
    "Quick Ratio": "Like current ratio but excludes inventory. More conservative liquidity measure.",
    "Gross Margin": "Revenue minus cost of goods sold, divided by revenue. Higher = better pricing power.",
    "Operating Margin": "EBIT / Revenue. Shows operational efficiency before interest and taxes.",
    "Net Margin": "Net Income / Revenue. Bottom-line profitability after all expenses.",
    "ROE": "Net Income / Shareholders' Equity. Measures return on shareholder investment. Can be inflated by leverage.",
    "ROA": "Net Income / Total Assets. Measures efficiency of asset utilization.",
    "Debt / Equity": "Total Debt / Total Equity. Lower is generally safer, but depends on industry.",
    "PEG Ratio": "P/E divided by earnings growth rate. PEG < 1.0 suggests undervalued relative to growth.",
}

# ============================================================================
# QUALITY SCORE WEIGHTS
# ============================================================================
QUALITY_WEIGHTS = {
    "profitability": 0.25,
    "growth": 0.20,
    "moat": 0.20,
    "financial_health": 0.15,
    "management": 0.10,
    "innovation": 0.10,
}

# Quality score thresholds
QUALITY_THRESHOLDS = {
    "gross_margin_target": 0.40,
    "revenue_growth_target": 0.10,
    "roe_target": 0.20,
    "debt_equity_target": 0.50,
    "rd_revenue_target": 0.10,
}

# ============================================================================
# FORMATTING / STYLE CONSTANTS
# ============================================================================
COLOR_SCHEME = {
    "header_bg": "1F4E79",
    "header_font": "FFFFFF",
    "section_bg": "2E75B6",
    "input_bg": "FFF2CC",
    "input_font": "0000FF",
    "link_font": "008000",
    "formula_font": "000000",
    "positive": "00B050",
    "negative": "FF0000",
    "neutral": "808080",
    "warning": "FFC000",
    "light_border": "D9E2F3",
    "alt_row": "F2F7FB",
    "good_bg": "E2EFDA",
    "warn_bg": "FFF2CC",
    "bad_bg": "FCE4EC",
    "strong_buy_bg": "006100",
    "buy_bg": "00B050",
    "hold_bg": "FFC000",
    "sell_bg": "FF0000",
    "strong_sell_bg": "8B0000",
}

# ============================================================================
# OUTPUT PREFERENCES
# ============================================================================
VERBOSE_MODE = True            # Add extra explanation columns
INCLUDE_CHARTS = True          # Excel charts
DECIMAL_PLACES = 2             # Decimal precision for percentages

# ============================================================================
# CACHE
# ============================================================================
CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".cache")
CACHE_EXPIRY_HOURS = 12

# ============================================================================
# OUTPUT
# ============================================================================
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
