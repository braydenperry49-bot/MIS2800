"""
Configuration for Stock Valuation Analyzer — Professional Edition.

API keys and settings. Set environment variables or edit defaults below.
"""
import os

# ---------------------------------------------------------------------------
# API Keys  –  set via environment variables or replace the defaults
# ---------------------------------------------------------------------------
ALPHA_VANTAGE_KEY = os.getenv("ALPHA_VANTAGE_KEY", "demo")
FMP_API_KEY = os.getenv("FMP_API_KEY", "demo")
FRED_API_KEY = os.getenv("FRED_API_KEY", "demo")
NEWS_API_KEY = os.getenv("NEWS_API_KEY", "")  # optional – newsapi.org

# ---------------------------------------------------------------------------
# Valuation Weights (must sum to 1.0)
# ---------------------------------------------------------------------------
VALUATION_WEIGHTS = {
    "dcf": 0.30,
    "comps": 0.20,
    "analyst": 0.20,
    "technical": 0.10,
    "sentiment": 0.10,
    "seasonal": 0.10,
}

# ---------------------------------------------------------------------------
# DCF Assumptions
# ---------------------------------------------------------------------------
DCF_DEFAULTS = {
    "projection_years": 5,
    "terminal_growth_rate": 0.025,       # 2.5 %
    "risk_free_rate": None,              # auto-fetched from FRED
    "equity_risk_premium": 0.055,        # 5.5 %
    "tax_rate": 0.21,                    # US corporate
    "margin_of_safety": 0.15,            # 15 %
}

# DCF Sensitivity ranges for the WACC vs Terminal Growth matrix
DCF_SENSITIVITY = {
    "wacc_steps": [-0.02, -0.01, 0.0, 0.01, 0.02],      # offsets from base WACC
    "tgr_steps": [-0.01, -0.005, 0.0, 0.005, 0.01],      # offsets from base TGR
}

# ---------------------------------------------------------------------------
# Comparable Companies
# ---------------------------------------------------------------------------
COMPS_COUNT = 8
COMPS_MULTIPLES = ["P/E", "EV/EBITDA", "P/S", "PEG", "P/B"]

# ---------------------------------------------------------------------------
# Macro Indicators (FRED series IDs)
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Financial Ratio Benchmarks  (value = "good if above")
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# Formatting / Style Constants
# ---------------------------------------------------------------------------
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
}

# ---------------------------------------------------------------------------
# Cache
# ---------------------------------------------------------------------------
CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".cache")
CACHE_EXPIRY_HOURS = 12

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
