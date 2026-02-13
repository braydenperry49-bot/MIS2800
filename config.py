"""
Configuration for Stock Valuation Analyzer.

API keys and settings. Set environment variables or edit defaults below.
"""
import os

# ---------------------------------------------------------------------------
# API Keys  –  set via environment variables or replace the defaults
# ---------------------------------------------------------------------------
# Free tier keys are sufficient for all of these.
ALPHA_VANTAGE_KEY = os.getenv("ALPHA_VANTAGE_KEY", "demo")
FMP_API_KEY = os.getenv("FMP_API_KEY", "demo")
FRED_API_KEY = os.getenv("FRED_API_KEY", "demo")
NEWS_API_KEY = os.getenv("NEWS_API_KEY", "")  # optional – newsapi.org

# ---------------------------------------------------------------------------
# Valuation Weights (must sum to 1.0)
# ---------------------------------------------------------------------------
VALUATION_WEIGHTS = {
    "dcf": 0.35,
    "comps": 0.25,
    "analyst": 0.20,
    "technical": 0.10,
    "sentiment": 0.10,
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

# ---------------------------------------------------------------------------
# Comparable Companies
# ---------------------------------------------------------------------------
COMPS_COUNT = 8          # number of peers to fetch
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
# Formatting / Style Constants
# ---------------------------------------------------------------------------
COLOR_SCHEME = {
    "header_bg": "1F4E79",       # dark blue
    "header_font": "FFFFFF",     # white
    "input_bg": "FFF2CC",        # light yellow
    "input_font": "0000FF",      # blue
    "link_font": "008000",       # green
    "formula_font": "000000",    # black
    "positive": "00B050",        # green
    "negative": "FF0000",        # red
    "neutral": "808080",         # gray
    "light_border": "D9E2F3",    # light blue
    "alt_row": "F2F7FB",         # very light blue
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
