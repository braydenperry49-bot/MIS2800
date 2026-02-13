#!/usr/bin/env python3
"""
Stock Valuation Analyzer
========================
A comprehensive stock analysis tool that combines multiple valuation methods
(DCF, Comparable Companies, Analyst Targets, Technical, and Sentiment) into
a single weighted fair-value estimate, then exports a formatted Excel report.

Usage:
    python stock_valuation_analyzer.py              # interactive prompt
    python stock_valuation_analyzer.py AAPL         # single ticker
    python stock_valuation_analyzer.py AAPL MSFT    # multiple tickers
"""

import argparse
import datetime
import hashlib
import json
import os
import sys
import time
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import requests
import yfinance as yf
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import config as cfg

warnings.filterwarnings("ignore", category=FutureWarning)

# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 1 — Caching Utilities
# ═══════════════════════════════════════════════════════════════════════════════

def _cache_path(key: str) -> str:
    os.makedirs(cfg.CACHE_DIR, exist_ok=True)
    hashed = hashlib.md5(key.encode()).hexdigest()
    return os.path.join(cfg.CACHE_DIR, f"{hashed}.json")


def cache_get(key: str):
    path = _cache_path(key)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r") as f:
            data = json.load(f)
        ts = data.get("_ts", 0)
        if (time.time() - ts) > cfg.CACHE_EXPIRY_HOURS * 3600:
            os.remove(path)
            return None
        return data.get("payload")
    except (json.JSONDecodeError, OSError):
        return None


def cache_set(key: str, payload):
    path = _cache_path(key)
    with open(path, "w") as f:
        json.dump({"_ts": time.time(), "payload": payload}, f)


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 2 — Data Fetching Helpers
# ═══════════════════════════════════════════════════════════════════════════════

def fetch_yfinance_data(ticker: str) -> dict:
    """Fetch comprehensive data from yfinance for a given ticker."""
    cache_key = f"yf_{ticker}"
    cached = cache_get(cache_key)
    if cached:
        return cached

    print(f"  [yfinance] Fetching data for {ticker}...")
    stock = yf.Ticker(ticker)

    info = {}
    try:
        info = stock.info or {}
    except Exception:
        pass

    # Historical prices — 5 years
    hist = pd.DataFrame()
    try:
        hist = stock.history(period="5y")
    except Exception:
        pass

    # Financial statements
    financials = {}
    for attr in ("income_stmt", "balance_sheet", "cashflow"):
        try:
            df = getattr(stock, attr, pd.DataFrame())
            if df is not None and not df.empty:
                financials[attr] = df.to_dict()
            else:
                financials[attr] = {}
        except Exception:
            financials[attr] = {}

    # Analyst recommendations
    recommendations = []
    try:
        rec = stock.recommendations
        if rec is not None and not rec.empty:
            recommendations = rec.tail(20).reset_index().to_dict(orient="records")
    except Exception:
        pass

    result = _serialize({
        "info": info,
        "history": hist.reset_index().to_dict(orient="list") if not hist.empty else {},
        "financials": financials,
        "recommendations": recommendations,
    })
    cache_set(cache_key, result)
    return result


def _serialize(obj):
    """Make an object JSON-serializable."""
    if isinstance(obj, dict):
        return {str(k): _serialize(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_serialize(v) for v in obj]
    if isinstance(obj, (np.integer,)):
        return int(obj)
    if isinstance(obj, (np.floating, np.float64)):
        return float(obj)
    if isinstance(obj, (np.bool_,)):
        return bool(obj)
    if isinstance(obj, (np.ndarray,)):
        return _serialize(obj.tolist())
    if isinstance(obj, (pd.Timestamp, datetime.datetime, datetime.date)):
        return obj.isoformat()
    if isinstance(obj, pd.Timedelta):
        return str(obj)
    try:
        if pd.isna(obj):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(obj, (str, int, float, bool, type(None))):
        return obj
    return str(obj)


def fetch_fred_series(series_id: str):
    """Fetch latest value of a FRED series via the FRED API."""
    cache_key = f"fred_{series_id}"
    cached = cache_get(cache_key)
    if cached is not None:
        return cached

    api_key = cfg.FRED_API_KEY
    if not api_key or api_key == "demo":
        return None

    url = "https://api.stlouisfed.org/fred/series/observations"
    params = {
        "series_id": series_id,
        "api_key": api_key,
        "file_type": "json",
        "sort_order": "desc",
        "limit": 5,
    }
    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        observations = resp.json().get("observations", [])
        for obs in observations:
            val = obs.get("value", ".")
            if val != ".":
                result = float(val)
                cache_set(cache_key, result)
                return result
    except Exception:
        pass
    return None


def fetch_fmp_peers(ticker: str) -> list:
    """Get peer/comparable company tickers from Financial Modeling Prep."""
    cache_key = f"fmp_peers_{ticker}"
    cached = cache_get(cache_key)
    if cached is not None:
        return cached

    api_key = cfg.FMP_API_KEY
    if not api_key or api_key == "demo":
        return []

    url = f"https://financialmodelingprep.com/api/v4/stock_peers"
    params = {"symbol": ticker, "apikey": api_key}
    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        data = resp.json()
        if data and isinstance(data, list) and "peersList" in data[0]:
            peers = data[0]["peersList"][: cfg.COMPS_COUNT]
            cache_set(cache_key, peers)
            return peers
    except Exception:
        pass
    return []


def fetch_news_sentiment(ticker: str) -> list:
    """Fetch recent news headlines for sentiment analysis."""
    cache_key = f"news_{ticker}"
    cached = cache_get(cache_key)
    if cached is not None:
        return cached

    api_key = cfg.NEWS_API_KEY
    if not api_key:
        return []

    url = "https://newsapi.org/v2/everything"
    params = {
        "q": ticker,
        "language": "en",
        "sortBy": "publishedAt",
        "pageSize": 30,
        "apiKey": api_key,
    }
    try:
        resp = requests.get(url, params=params, timeout=10)
        resp.raise_for_status()
        articles = resp.json().get("articles", [])
        headlines = [
            {"title": a.get("title", ""), "description": a.get("description", "")}
            for a in articles
            if a.get("title")
        ]
        cache_set(cache_key, headlines)
        return headlines
    except Exception:
        return []


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 3 — Valuation Models
# ═══════════════════════════════════════════════════════════════════════════════

class ValuationResult:
    """Container for a single valuation method's output."""
    def __init__(self, method: str, fair_value: float, confidence: float,
                 details: dict = None):
        self.method = method
        self.fair_value = fair_value          # per-share value
        self.confidence = confidence          # 0-1 scale
        self.details = details or {}


# ── 3a. DCF Valuation ────────────────────────────────────────────────────────

def dcf_valuation(data: dict) -> ValuationResult:
    """Discounted Cash Flow valuation using free cash flow projections."""
    info = data.get("info", {})
    financials = data.get("financials", {})

    shares = info.get("sharesOutstanding") or info.get("impliedSharesOutstanding", 0)
    if not shares:
        return ValuationResult("DCF", 0, 0, {"error": "No shares outstanding data"})

    # Gather historical free cash flow
    cf_data = financials.get("cashflow", {})
    if not cf_data:
        return ValuationResult("DCF", 0, 0, {"error": "No cash flow data"})

    cf_df = pd.DataFrame(cf_data)
    fcf_row = None
    for label in ("Free Cash Flow", "FreeCashFlow"):
        if label in cf_df.index:
            fcf_row = cf_df.loc[label]
            break

    if fcf_row is None:
        # Try computing: Operating Cash Flow - CapEx
        op_cf = None
        capex = None
        for label in ("Operating Cash Flow", "OperatingCashFlow",
                       "Total Cash From Operating Activities"):
            if label in cf_df.index:
                op_cf = cf_df.loc[label]
                break
        for label in ("Capital Expenditure", "CapitalExpenditure",
                       "Capital Expenditures"):
            if label in cf_df.index:
                capex = cf_df.loc[label]
                break
        if op_cf is not None and capex is not None:
            fcf_row = op_cf.astype(float) + capex.astype(float)  # capex is negative
        else:
            return ValuationResult("DCF", 0, 0, {"error": "Cannot compute FCF"})

    fcf_values = [v for v in fcf_row.values if v is not None and not np.isnan(float(v))]
    if not fcf_values:
        return ValuationResult("DCF", 0, 0, {"error": "No valid FCF values"})

    latest_fcf = float(fcf_values[0])

    # Growth rate estimate
    if len(fcf_values) >= 2:
        growth_rates = []
        for i in range(len(fcf_values) - 1):
            prev = float(fcf_values[i + 1])
            curr = float(fcf_values[i])
            if prev > 0:
                growth_rates.append((curr - prev) / prev)
        avg_growth = np.mean(growth_rates) if growth_rates else 0.05
        avg_growth = np.clip(avg_growth, -0.05, 0.30)  # cap growth rate
    else:
        avg_growth = 0.05

    # Discount rate (WACC approximation)
    beta = info.get("beta", 1.0) or 1.0
    risk_free = cfg.DCF_DEFAULTS["risk_free_rate"]
    if risk_free is None:
        risk_free_fetched = fetch_fred_series("DGS10")
        risk_free = (risk_free_fetched / 100.0) if risk_free_fetched else 0.04
    erp = cfg.DCF_DEFAULTS["equity_risk_premium"]
    cost_of_equity = risk_free + beta * erp

    # Debt cost approximation
    total_debt = info.get("totalDebt", 0) or 0
    market_cap = info.get("marketCap", 0) or 0
    interest_expense = abs(info.get("interestExpense", 0) or 0)

    if total_debt > 0 and interest_expense > 0:
        cost_of_debt = interest_expense / total_debt
    else:
        cost_of_debt = risk_free + 0.02

    tax_rate = cfg.DCF_DEFAULTS["tax_rate"]
    total_value = market_cap + total_debt
    if total_value > 0:
        equity_weight = market_cap / total_value
        debt_weight = total_debt / total_value
    else:
        equity_weight, debt_weight = 1.0, 0.0

    wacc = equity_weight * cost_of_equity + debt_weight * cost_of_debt * (1 - tax_rate)
    wacc = max(wacc, 0.06)  # floor

    # Project FCFs
    proj_years = cfg.DCF_DEFAULTS["projection_years"]
    terminal_growth = cfg.DCF_DEFAULTS["terminal_growth_rate"]
    projected_fcfs = []
    for yr in range(1, proj_years + 1):
        projected = latest_fcf * ((1 + avg_growth) ** yr)
        projected_fcfs.append(projected)

    # Terminal value (Gordon Growth)
    terminal_fcf = projected_fcfs[-1] * (1 + terminal_growth)
    terminal_value = terminal_fcf / (wacc - terminal_growth) if wacc > terminal_growth else 0

    # Discount everything back
    pv_fcfs = sum(fcf / ((1 + wacc) ** yr) for yr, fcf in enumerate(projected_fcfs, 1))
    pv_terminal = terminal_value / ((1 + wacc) ** proj_years)
    enterprise_value = pv_fcfs + pv_terminal

    # Equity value
    cash = info.get("totalCash", 0) or 0
    equity_value = enterprise_value + cash - total_debt
    fair_value_per_share = equity_value / shares if shares > 0 else 0

    # Apply margin of safety
    mos = cfg.DCF_DEFAULTS["margin_of_safety"]
    buy_price = fair_value_per_share * (1 - mos)

    # Confidence based on data availability
    confidence = 0.7
    if len(fcf_values) >= 3:
        confidence += 0.1
    if total_debt > 0 and interest_expense > 0:
        confidence += 0.1
    confidence = min(confidence, 1.0)

    details = {
        "latest_fcf": latest_fcf,
        "fcf_growth_rate": avg_growth,
        "wacc": wacc,
        "cost_of_equity": cost_of_equity,
        "cost_of_debt": cost_of_debt,
        "beta": beta,
        "risk_free_rate": risk_free,
        "projected_fcfs": projected_fcfs,
        "terminal_value": terminal_value,
        "enterprise_value": enterprise_value,
        "equity_value": equity_value,
        "shares_outstanding": shares,
        "margin_of_safety": mos,
        "buy_price": buy_price,
    }

    return ValuationResult("DCF", fair_value_per_share, confidence, details)


# ── 3b. Comparable Companies ─────────────────────────────────────────────────

def comps_valuation(data: dict, ticker: str) -> ValuationResult:
    """Relative valuation using peer-company multiples."""
    info = data.get("info", {})
    peers = fetch_fmp_peers(ticker)

    # Fallback: use sector/industry from yfinance to pick known peers
    if not peers:
        industry_peers = info.get("recommendedSymbols", [])
        if industry_peers:
            peers = [p.get("symbol") for p in industry_peers if p.get("symbol")]
        if not peers:
            # Last resort — try yfinance sector peers
            try:
                stock = yf.Ticker(ticker)
                rec = stock.recommendations
                if rec is not None and not rec.empty:
                    pass  # recommendations don't give peers directly
            except Exception:
                pass

    if not peers:
        return ValuationResult("Comps", 0, 0, {"error": "No peer companies found"})

    peers = peers[: cfg.COMPS_COUNT]

    # Collect multiples for the target and peers
    target_multiples = _extract_multiples(info)
    peer_data = []

    for peer_ticker in peers:
        try:
            peer_stock = yf.Ticker(peer_ticker)
            peer_info = peer_stock.info or {}
            pm = _extract_multiples(peer_info)
            pm["ticker"] = peer_ticker
            pm["name"] = peer_info.get("shortName", peer_ticker)
            pm["marketCap"] = peer_info.get("marketCap", 0)
            peer_data.append(pm)
        except Exception:
            continue

    if not peer_data:
        return ValuationResult("Comps", 0, 0, {"error": "Could not fetch peer data"})

    # Calculate implied fair values from each multiple
    current_price = info.get("currentPrice") or info.get("regularMarketPrice", 0)
    implied_values = []
    multiple_details = {}

    for mult_name in cfg.COMPS_MULTIPLES:
        key = _multiple_key(mult_name)
        target_val = target_multiples.get(key)
        peer_vals = [p.get(key) for p in peer_data if p.get(key) and p[key] > 0]

        if not peer_vals or not target_val or target_val <= 0:
            continue

        median_peer = np.median(peer_vals)

        # Implied value = current_price * (median_peer / target_val)
        if target_val != 0:
            implied = current_price * (median_peer / target_val)
            implied_values.append(implied)
            multiple_details[mult_name] = {
                "target": round(target_val, 2),
                "peer_median": round(median_peer, 2),
                "implied_value": round(implied, 2),
            }

    if not implied_values:
        return ValuationResult("Comps", 0, 0, {"error": "No valid multiples"})

    fair_value = np.mean(implied_values)
    confidence = min(0.5 + 0.1 * len(implied_values), 0.9)

    details = {
        "peers": [p.get("ticker", "?") for p in peer_data],
        "multiples": multiple_details,
        "current_price": current_price,
    }

    return ValuationResult("Comps", fair_value, confidence, details)


def _extract_multiples(info: dict) -> dict:
    return {
        "pe": info.get("trailingPE") or info.get("forwardPE"),
        "ev_ebitda": info.get("enterpriseToEbitda"),
        "ps": info.get("priceToSalesTrailing12Months"),
        "peg": info.get("pegRatio"),
        "pb": info.get("priceToBook"),
    }


def _multiple_key(name: str) -> str:
    mapping = {"P/E": "pe", "EV/EBITDA": "ev_ebitda", "P/S": "ps",
               "PEG": "peg", "P/B": "pb"}
    return mapping.get(name, name.lower())


# ── 3c. Analyst Price Targets ────────────────────────────────────────────────

def analyst_valuation(data: dict) -> ValuationResult:
    """Use analyst consensus price targets."""
    info = data.get("info", {})

    target_mean = info.get("targetMeanPrice")
    target_median = info.get("targetMedianPrice")
    target_high = info.get("targetHighPrice")
    target_low = info.get("targetLowPrice")
    num_analysts = info.get("numberOfAnalystOpinions", 0)

    if not target_mean and not target_median:
        return ValuationResult("Analyst Targets", 0, 0,
                               {"error": "No analyst target data"})

    fair_value = target_median or target_mean

    # Confidence scales with number of analysts
    if num_analysts >= 20:
        confidence = 0.85
    elif num_analysts >= 10:
        confidence = 0.70
    elif num_analysts >= 5:
        confidence = 0.55
    else:
        confidence = 0.35

    # Recommendation trend
    recommendations = data.get("recommendations", [])
    rec_summary = {}
    for rec in recommendations[-10:]:
        grade = str(rec.get("toGrade", rec.get("To Grade", ""))).lower()
        if not grade:
            continue
        if any(w in grade for w in ("buy", "outperform", "overweight")):
            rec_summary["buy"] = rec_summary.get("buy", 0) + 1
        elif any(w in grade for w in ("sell", "underperform", "underweight")):
            rec_summary["sell"] = rec_summary.get("sell", 0) + 1
        else:
            rec_summary["hold"] = rec_summary.get("hold", 0) + 1

    details = {
        "target_mean": target_mean,
        "target_median": target_median,
        "target_high": target_high,
        "target_low": target_low,
        "num_analysts": num_analysts,
        "recommendation_summary": rec_summary,
    }

    return ValuationResult("Analyst Targets", fair_value, confidence, details)


# ── 3d. Technical Analysis ───────────────────────────────────────────────────

def technical_valuation(data: dict) -> ValuationResult:
    """Simple technical analysis based on moving averages and RSI."""
    hist_raw = data.get("history", {})
    if not hist_raw or "Close" not in hist_raw:
        return ValuationResult("Technical", 0, 0, {"error": "No price history"})

    close_prices = [v for v in hist_raw["Close"] if v is not None]
    if len(close_prices) < 200:
        return ValuationResult("Technical", 0, 0, {"error": "Insufficient history"})

    prices = pd.Series(close_prices)
    current_price = prices.iloc[-1]

    # Moving averages
    sma_50 = prices.tail(50).mean()
    sma_200 = prices.tail(200).mean()
    ema_20 = prices.ewm(span=20, adjust=False).mean().iloc[-1]

    # RSI (14-day)
    delta = prices.diff().tail(15)
    gain = delta.where(delta > 0, 0.0).mean()
    loss = (-delta.where(delta < 0, 0.0)).mean()
    rs = gain / loss if loss != 0 else 100
    rsi = 100 - (100 / (1 + rs))

    # MACD
    ema_12 = prices.ewm(span=12, adjust=False).mean().iloc[-1]
    ema_26 = prices.ewm(span=26, adjust=False).mean().iloc[-1]
    macd = ema_12 - ema_26

    # Bollinger Bands (20-day)
    bb_sma = prices.tail(20).mean()
    bb_std = prices.tail(20).std()
    bb_upper = bb_sma + 2 * bb_std
    bb_lower = bb_sma - 2 * bb_std

    # Score signals
    signals = []
    if current_price > sma_50 > sma_200:
        signals.append(("Golden Cross / Uptrend", 1))
    elif current_price < sma_50 < sma_200:
        signals.append(("Death Cross / Downtrend", -1))
    else:
        signals.append(("Mixed trend", 0))

    if rsi < 30:
        signals.append(("RSI oversold", 1))
    elif rsi > 70:
        signals.append(("RSI overbought", -1))
    else:
        signals.append(("RSI neutral", 0))

    if macd > 0:
        signals.append(("MACD bullish", 1))
    else:
        signals.append(("MACD bearish", -1))

    if current_price < bb_lower:
        signals.append(("Below Bollinger lower band", 1))
    elif current_price > bb_upper:
        signals.append(("Above Bollinger upper band", -1))
    else:
        signals.append(("Within Bollinger bands", 0))

    # Compute technical fair value as an adjusted price
    score = sum(s[1] for s in signals) / len(signals)  # -1 to 1
    adjustment = 1 + score * 0.10  # +/- 10% max adjustment
    fair_value = current_price * adjustment

    # Confidence is moderate for pure technicals
    confidence = 0.50

    details = {
        "current_price": round(current_price, 2),
        "sma_50": round(sma_50, 2),
        "sma_200": round(sma_200, 2),
        "ema_20": round(ema_20, 2),
        "rsi": round(rsi, 2),
        "macd": round(macd, 2),
        "bollinger_upper": round(bb_upper, 2),
        "bollinger_lower": round(bb_lower, 2),
        "signals": signals,
        "score": round(score, 3),
    }

    return ValuationResult("Technical", fair_value, confidence, details)


# ── 3e. Sentiment Analysis ───────────────────────────────────────────────────

def sentiment_valuation(data: dict, ticker: str) -> ValuationResult:
    """Sentiment analysis from news headlines and key metrics."""
    info = data.get("info", {})
    current_price = info.get("currentPrice") or info.get("regularMarketPrice", 0)
    if not current_price:
        return ValuationResult("Sentiment", 0, 0, {"error": "No current price"})

    scores = []

    # 1) News headline sentiment (simple keyword-based)
    headlines = fetch_news_sentiment(ticker)
    positive_words = {"beat", "surge", "jump", "gain", "rise", "profit", "growth",
                      "upgrade", "outperform", "rally", "record", "strong", "bullish",
                      "buy", "exceed", "boom", "soar", "positive", "optimistic"}
    negative_words = {"miss", "fall", "drop", "loss", "decline", "downgrade",
                      "underperform", "crash", "weak", "bearish", "sell", "cut",
                      "warn", "negative", "pessimistic", "layoff", "recall", "lawsuit"}

    if headlines:
        pos_count = 0
        neg_count = 0
        for h in headlines:
            text = (h.get("title", "") + " " + h.get("description", "")).lower()
            pos_count += sum(1 for w in positive_words if w in text)
            neg_count += sum(1 for w in negative_words if w in text)
        total = pos_count + neg_count
        if total > 0:
            news_score = (pos_count - neg_count) / total  # -1 to 1
            scores.append(news_score)

    # 2) Fundamental sentiment signals
    profit_margin = info.get("profitMargins")
    if profit_margin is not None:
        if profit_margin > 0.15:
            scores.append(0.5)
        elif profit_margin > 0:
            scores.append(0.2)
        else:
            scores.append(-0.5)

    revenue_growth = info.get("revenueGrowth")
    if revenue_growth is not None:
        if revenue_growth > 0.15:
            scores.append(0.5)
        elif revenue_growth > 0:
            scores.append(0.2)
        else:
            scores.append(-0.3)

    earnings_growth = info.get("earningsGrowth")
    if earnings_growth is not None:
        if earnings_growth > 0.15:
            scores.append(0.5)
        elif earnings_growth > 0:
            scores.append(0.2)
        else:
            scores.append(-0.3)

    # 3) Insider/institutional sentiment
    insider_pct = info.get("heldPercentInsiders", 0) or 0
    inst_pct = info.get("heldPercentInstitutions", 0) or 0
    if inst_pct > 0.7:
        scores.append(0.3)
    if insider_pct > 0.10:
        scores.append(0.2)

    if not scores:
        return ValuationResult("Sentiment", 0, 0, {"error": "No sentiment data"})

    avg_score = np.mean(scores)  # -1 to 1
    adjustment = 1 + avg_score * 0.08  # +/- 8% max
    fair_value = current_price * adjustment

    confidence = min(0.3 + 0.1 * len(scores), 0.65)

    details = {
        "current_price": current_price,
        "headline_count": len(headlines),
        "sentiment_score": round(avg_score, 3),
        "individual_scores": [round(s, 3) for s in scores],
        "adjustment_factor": round(adjustment, 4),
    }

    return ValuationResult("Sentiment", fair_value, confidence, details)


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 4 — Composite Valuation
# ═══════════════════════════════════════════════════════════════════════════════

def composite_valuation(results: list, weights: dict) -> dict:
    """Combine multiple valuation results into a single weighted estimate."""
    method_map = {
        "DCF": "dcf",
        "Comps": "comps",
        "Analyst Targets": "analyst",
        "Technical": "technical",
        "Sentiment": "sentiment",
    }

    weighted_sum = 0.0
    total_weight = 0.0
    breakdown = {}

    for r in results:
        key = method_map.get(r.method, r.method.lower())
        w = weights.get(key, 0)

        if r.fair_value > 0 and r.confidence > 0:
            effective_weight = w * r.confidence
            weighted_sum += r.fair_value * effective_weight
            total_weight += effective_weight
            breakdown[r.method] = {
                "fair_value": round(r.fair_value, 2),
                "confidence": round(r.confidence, 2),
                "weight": w,
                "effective_weight": round(effective_weight, 4),
            }
        else:
            breakdown[r.method] = {
                "fair_value": 0,
                "confidence": 0,
                "weight": w,
                "effective_weight": 0,
                "note": r.details.get("error", "No data"),
            }

    composite_value = weighted_sum / total_weight if total_weight > 0 else 0

    return {
        "composite_fair_value": round(composite_value, 2),
        "total_effective_weight": round(total_weight, 4),
        "breakdown": breakdown,
    }


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 5 — Macro Environment
# ═══════════════════════════════════════════════════════════════════════════════

def fetch_macro_environment() -> dict:
    """Fetch key macroeconomic indicators from FRED."""
    print("  [FRED] Fetching macro indicators...")
    macro = {}
    for name, series_id in cfg.FRED_SERIES.items():
        val = fetch_fred_series(series_id)
        macro[name] = val
    return macro


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 6 — Excel Report Generation
# ═══════════════════════════════════════════════════════════════════════════════

def _style_header(ws, row, max_col):
    """Apply header styling to a row."""
    hdr_fill = PatternFill(start_color=cfg.COLOR_SCHEME["header_bg"],
                           end_color=cfg.COLOR_SCHEME["header_bg"],
                           fill_type="solid")
    hdr_font = Font(color=cfg.COLOR_SCHEME["header_font"], bold=True, size=11)
    thin_border = Border(
        bottom=Side(style="thin", color=cfg.COLOR_SCHEME["light_border"])
    )
    for col in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border


def _fmt_number(value, fmt_type="number"):
    """Format a number for display."""
    if value is None:
        return "N/A"
    try:
        value = float(value)
    except (TypeError, ValueError):
        return str(value)
    if fmt_type == "currency":
        return f"${value:,.2f}"
    if fmt_type == "percent":
        return f"{value:.2%}"
    if fmt_type == "large_currency":
        if abs(value) >= 1e12:
            return f"${value / 1e12:,.2f}T"
        if abs(value) >= 1e9:
            return f"${value / 1e9:,.2f}B"
        if abs(value) >= 1e6:
            return f"${value / 1e6:,.2f}M"
        return f"${value:,.0f}"
    if fmt_type == "ratio":
        return f"{value:.2f}x"
    return f"{value:,.2f}"


def generate_excel_report(ticker: str, data: dict, results: list,
                          composite: dict, macro: dict) -> str:
    """Generate a comprehensive Excel report."""
    os.makedirs(cfg.OUTPUT_DIR, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{ticker}_valuation_{timestamp}.xlsx"
    filepath = os.path.join(cfg.OUTPUT_DIR, filename)

    wb = Workbook()
    info = data.get("info", {})
    current_price = info.get("currentPrice") or info.get("regularMarketPrice", 0)

    # ── Sheet 1: Executive Summary ──
    ws = wb.active
    ws.title = "Executive Summary"
    ws.sheet_properties.tabColor = cfg.COLOR_SCHEME["header_bg"]

    # Title
    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = f"Stock Valuation Report — {ticker}"
    title_cell.font = Font(size=18, bold=True, color=cfg.COLOR_SCHEME["header_bg"])
    title_cell.alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:G2")
    ws["A2"].value = (f"{info.get('shortName', ticker)}  |  "
                      f"{info.get('sector', 'N/A')}  |  "
                      f"{info.get('industry', 'N/A')}")
    ws["A2"].font = Font(size=11, italic=True)
    ws["A2"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A3:G3")
    ws["A3"].value = f"Report generated: {datetime.datetime.now():%Y-%m-%d %H:%M}"
    ws["A3"].font = Font(size=9, color=cfg.COLOR_SCHEME["neutral"])
    ws["A3"].alignment = Alignment(horizontal="center")

    # Key metrics box
    row = 5
    headers = ["Metric", "Value"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=row, column=c, value=h)
    _style_header(ws, row, len(headers))

    key_metrics = [
        ("Current Price", _fmt_number(current_price, "currency")),
        ("Composite Fair Value", _fmt_number(composite["composite_fair_value"], "currency")),
        ("Upside / Downside",
         _fmt_number((composite["composite_fair_value"] - current_price) / current_price
                     if current_price else 0, "percent")),
        ("Market Cap", _fmt_number(info.get("marketCap"), "large_currency")),
        ("52-Week High", _fmt_number(info.get("fiftyTwoWeekHigh"), "currency")),
        ("52-Week Low", _fmt_number(info.get("fiftyTwoWeekLow"), "currency")),
        ("Beta", _fmt_number(info.get("beta"), "ratio")),
        ("Dividend Yield", _fmt_number(info.get("dividendYield"), "percent")),
        ("P/E (Trailing)", _fmt_number(info.get("trailingPE"), "ratio")),
        ("EPS (TTM)", _fmt_number(info.get("trailingEps"), "currency")),
    ]

    pos_font = Font(color=cfg.COLOR_SCHEME["positive"], bold=True)
    neg_font = Font(color=cfg.COLOR_SCHEME["negative"], bold=True)
    alt_fill = PatternFill(start_color=cfg.COLOR_SCHEME["alt_row"],
                           end_color=cfg.COLOR_SCHEME["alt_row"], fill_type="solid")

    for i, (metric, value) in enumerate(key_metrics):
        r = row + 1 + i
        ws.cell(row=r, column=1, value=metric)
        cell = ws.cell(row=r, column=2, value=value)
        if i % 2 == 0:
            ws.cell(row=r, column=1).fill = alt_fill
            cell.fill = alt_fill
        # Highlight upside/downside
        if metric == "Upside / Downside" and current_price > 0:
            ratio = (composite["composite_fair_value"] - current_price) / current_price
            cell.font = pos_font if ratio >= 0 else neg_font

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 18

    # Signal / verdict
    r = row + len(key_metrics) + 2
    ws.merge_cells(f"A{r}:B{r}")
    if current_price > 0:
        pct_diff = (composite["composite_fair_value"] - current_price) / current_price
        if pct_diff > 0.15:
            verdict = "UNDERVALUED — Consider Buying"
            v_color = cfg.COLOR_SCHEME["positive"]
        elif pct_diff < -0.15:
            verdict = "OVERVALUED — Consider Selling"
            v_color = cfg.COLOR_SCHEME["negative"]
        else:
            verdict = "FAIRLY VALUED — Hold"
            v_color = cfg.COLOR_SCHEME["neutral"]
    else:
        verdict = "INSUFFICIENT DATA"
        v_color = cfg.COLOR_SCHEME["neutral"]

    vc = ws.cell(row=r, column=1, value=verdict)
    vc.font = Font(size=14, bold=True, color=v_color)
    vc.alignment = Alignment(horizontal="center")

    # ── Sheet 2: Valuation Breakdown ──
    ws2 = wb.create_sheet("Valuation Breakdown")
    ws2.sheet_properties.tabColor = "4472C4"

    ws2.merge_cells("A1:F1")
    ws2["A1"].value = "Valuation Method Comparison"
    ws2["A1"].font = Font(size=14, bold=True, color=cfg.COLOR_SCHEME["header_bg"])

    row = 3
    headers = ["Method", "Fair Value", "Confidence", "Weight",
               "Effective Weight", "Notes"]
    for c, h in enumerate(headers, 1):
        ws2.cell(row=row, column=c, value=h)
    _style_header(ws2, row, len(headers))

    for i, (method, bd) in enumerate(composite["breakdown"].items()):
        r = row + 1 + i
        ws2.cell(row=r, column=1, value=method)
        ws2.cell(row=r, column=2, value=_fmt_number(bd["fair_value"], "currency"))
        ws2.cell(row=r, column=3, value=_fmt_number(bd["confidence"], "percent"))
        ws2.cell(row=r, column=4, value=_fmt_number(bd["weight"], "percent"))
        ws2.cell(row=r, column=5, value=_fmt_number(bd["effective_weight"], "percent"))
        ws2.cell(row=r, column=6, value=bd.get("note", ""))
        if i % 2 == 0:
            for c in range(1, 7):
                ws2.cell(row=r, column=c).fill = alt_fill

    # Summary row
    r = row + 1 + len(composite["breakdown"])
    ws2.cell(row=r + 1, column=1, value="COMPOSITE FAIR VALUE").font = Font(bold=True)
    ws2.cell(row=r + 1, column=2,
             value=_fmt_number(composite["composite_fair_value"], "currency"))
    ws2.cell(row=r + 1, column=2).font = Font(bold=True, size=12)

    for c in range(1, 7):
        ws2.column_dimensions[get_column_letter(c)].width = 18

    # ── Sheet 3: DCF Detail ──
    ws3 = wb.create_sheet("DCF Analysis")
    ws3.sheet_properties.tabColor = "00B050"

    dcf_result = next((r for r in results if r.method == "DCF"), None)
    if dcf_result and dcf_result.fair_value > 0:
        d = dcf_result.details
        ws3["A1"].value = "Discounted Cash Flow Analysis"
        ws3["A1"].font = Font(size=14, bold=True, color=cfg.COLOR_SCHEME["header_bg"])

        row = 3
        assumptions = [
            ("Latest Free Cash Flow", _fmt_number(d.get("latest_fcf"), "large_currency")),
            ("FCF Growth Rate", _fmt_number(d.get("fcf_growth_rate"), "percent")),
            ("WACC (Discount Rate)", _fmt_number(d.get("wacc"), "percent")),
            ("Cost of Equity", _fmt_number(d.get("cost_of_equity"), "percent")),
            ("Cost of Debt", _fmt_number(d.get("cost_of_debt"), "percent")),
            ("Beta", _fmt_number(d.get("beta"))),
            ("Risk-Free Rate", _fmt_number(d.get("risk_free_rate"), "percent")),
            ("Terminal Value", _fmt_number(d.get("terminal_value"), "large_currency")),
            ("Enterprise Value", _fmt_number(d.get("enterprise_value"), "large_currency")),
            ("Equity Value", _fmt_number(d.get("equity_value"), "large_currency")),
            ("Shares Outstanding", _fmt_number(d.get("shares_outstanding"), "large_currency")),
            ("DCF Fair Value / Share", _fmt_number(dcf_result.fair_value, "currency")),
            ("Margin of Safety", _fmt_number(d.get("margin_of_safety"), "percent")),
            ("Buy Price (w/ MoS)", _fmt_number(d.get("buy_price"), "currency")),
        ]

        headers = ["Assumption", "Value"]
        for c, h in enumerate(headers, 1):
            ws3.cell(row=row, column=c, value=h)
        _style_header(ws3, row, 2)

        for i, (label, val) in enumerate(assumptions):
            r = row + 1 + i
            ws3.cell(row=r, column=1, value=label)
            ws3.cell(row=r, column=2, value=val)
            if i % 2 == 0:
                ws3.cell(row=r, column=1).fill = alt_fill
                ws3.cell(row=r, column=2).fill = alt_fill

        # Projected FCFs
        proj_row = row + len(assumptions) + 2
        ws3.cell(row=proj_row, column=1, value="Projected Free Cash Flows").font = Font(
            bold=True)
        proj_fcfs = d.get("projected_fcfs", [])
        for i, fcf in enumerate(proj_fcfs):
            ws3.cell(row=proj_row + 1 + i, column=1, value=f"Year {i + 1}")
            ws3.cell(row=proj_row + 1 + i, column=2,
                     value=_fmt_number(fcf, "large_currency"))

        ws3.column_dimensions["A"].width = 28
        ws3.column_dimensions["B"].width = 20
    else:
        ws3["A1"].value = "DCF Analysis — Insufficient Data"
        ws3["A1"].font = Font(size=14, bold=True, color=cfg.COLOR_SCHEME["negative"])
        if dcf_result:
            ws3["A3"].value = dcf_result.details.get("error", "")

    # ── Sheet 4: Technical Analysis ──
    ws4 = wb.create_sheet("Technical Analysis")
    ws4.sheet_properties.tabColor = "FFC000"

    tech_result = next((r for r in results if r.method == "Technical"), None)
    if tech_result and tech_result.fair_value > 0:
        d = tech_result.details
        ws4["A1"].value = "Technical Analysis"
        ws4["A1"].font = Font(size=14, bold=True, color=cfg.COLOR_SCHEME["header_bg"])

        row = 3
        indicators = [
            ("Current Price", _fmt_number(d.get("current_price"), "currency")),
            ("50-Day SMA", _fmt_number(d.get("sma_50"), "currency")),
            ("200-Day SMA", _fmt_number(d.get("sma_200"), "currency")),
            ("20-Day EMA", _fmt_number(d.get("ema_20"), "currency")),
            ("RSI (14)", _fmt_number(d.get("rsi"))),
            ("MACD", _fmt_number(d.get("macd"), "currency")),
            ("Bollinger Upper", _fmt_number(d.get("bollinger_upper"), "currency")),
            ("Bollinger Lower", _fmt_number(d.get("bollinger_lower"), "currency")),
            ("Technical Score", _fmt_number(d.get("score"))),
            ("Technical Fair Value", _fmt_number(tech_result.fair_value, "currency")),
        ]

        headers = ["Indicator", "Value"]
        for c, h in enumerate(headers, 1):
            ws4.cell(row=row, column=c, value=h)
        _style_header(ws4, row, 2)

        for i, (label, val) in enumerate(indicators):
            r = row + 1 + i
            ws4.cell(row=r, column=1, value=label)
            ws4.cell(row=r, column=2, value=val)
            if i % 2 == 0:
                ws4.cell(row=r, column=1).fill = alt_fill
                ws4.cell(row=r, column=2).fill = alt_fill

        # Signals
        sig_row = row + len(indicators) + 2
        ws4.cell(row=sig_row, column=1, value="Signals").font = Font(bold=True)
        for i, (signal, direction) in enumerate(d.get("signals", [])):
            r = sig_row + 1 + i
            ws4.cell(row=r, column=1, value=signal)
            label = "Bullish" if direction > 0 else ("Bearish" if direction < 0 else "Neutral")
            cell = ws4.cell(row=r, column=2, value=label)
            if direction > 0:
                cell.font = pos_font
            elif direction < 0:
                cell.font = neg_font

        ws4.column_dimensions["A"].width = 24
        ws4.column_dimensions["B"].width = 16
    else:
        ws4["A1"].value = "Technical Analysis — Insufficient Data"
        ws4["A1"].font = Font(size=14, bold=True, color=cfg.COLOR_SCHEME["negative"])

    # ── Sheet 5: Macro Environment ──
    ws5 = wb.create_sheet("Macro Environment")
    ws5.sheet_properties.tabColor = "7030A0"

    ws5["A1"].value = "Macroeconomic Environment"
    ws5["A1"].font = Font(size=14, bold=True, color=cfg.COLOR_SCHEME["header_bg"])

    row = 3
    headers = ["Indicator", "Latest Value"]
    for c, h in enumerate(headers, 1):
        ws5.cell(row=row, column=c, value=h)
    _style_header(ws5, row, 2)

    for i, (name, value) in enumerate(macro.items()):
        r = row + 1 + i
        ws5.cell(row=r, column=1, value=name)
        if value is not None:
            ws5.cell(row=r, column=2, value=round(value, 2))
        else:
            ws5.cell(row=r, column=2, value="N/A")
            ws5.cell(row=r, column=2).font = Font(color=cfg.COLOR_SCHEME["neutral"])
        if i % 2 == 0:
            ws5.cell(row=r, column=1).fill = alt_fill
            ws5.cell(row=r, column=2).fill = alt_fill

    ws5.column_dimensions["A"].width = 28
    ws5.column_dimensions["B"].width = 16

    # ── Sheet 6: Company Profile ──
    ws6 = wb.create_sheet("Company Profile")
    ws6.sheet_properties.tabColor = "ED7D31"

    ws6["A1"].value = f"Company Profile — {info.get('shortName', ticker)}"
    ws6["A1"].font = Font(size=14, bold=True, color=cfg.COLOR_SCHEME["header_bg"])

    profile_fields = [
        ("Ticker", ticker),
        ("Company Name", info.get("shortName", "N/A")),
        ("Sector", info.get("sector", "N/A")),
        ("Industry", info.get("industry", "N/A")),
        ("Country", info.get("country", "N/A")),
        ("Exchange", info.get("exchange", "N/A")),
        ("Currency", info.get("currency", "N/A")),
        ("Website", info.get("website", "N/A")),
        ("Full-Time Employees", _fmt_number(info.get("fullTimeEmployees"))),
        ("Market Cap", _fmt_number(info.get("marketCap"), "large_currency")),
        ("Enterprise Value", _fmt_number(info.get("enterpriseValue"), "large_currency")),
        ("Revenue (TTM)", _fmt_number(info.get("totalRevenue"), "large_currency")),
        ("Net Income (TTM)", _fmt_number(info.get("netIncomeToCommon"), "large_currency")),
        ("Profit Margin", _fmt_number(info.get("profitMargins"), "percent")),
        ("Operating Margin", _fmt_number(info.get("operatingMargins"), "percent")),
        ("ROE", _fmt_number(info.get("returnOnEquity"), "percent")),
        ("ROA", _fmt_number(info.get("returnOnAssets"), "percent")),
        ("Debt to Equity", _fmt_number(info.get("debtToEquity"))),
        ("Current Ratio", _fmt_number(info.get("currentRatio"), "ratio")),
    ]

    row = 3
    headers = ["Field", "Value"]
    for c, h in enumerate(headers, 1):
        ws6.cell(row=row, column=c, value=h)
    _style_header(ws6, row, 2)

    for i, (field, value) in enumerate(profile_fields):
        r = row + 1 + i
        ws6.cell(row=r, column=1, value=field)
        ws6.cell(row=r, column=2, value=value)
        if i % 2 == 0:
            ws6.cell(row=r, column=1).fill = alt_fill
            ws6.cell(row=r, column=2).fill = alt_fill

    # Business summary
    summary_row = row + len(profile_fields) + 2
    ws6.cell(row=summary_row, column=1, value="Business Summary").font = Font(bold=True)
    summary = info.get("longBusinessSummary", "N/A")
    ws6.merge_cells(f"A{summary_row + 1}:D{summary_row + 5}")
    cell = ws6.cell(row=summary_row + 1, column=1, value=summary)
    cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws6.column_dimensions["A"].width = 24
    ws6.column_dimensions["B"].width = 20

    # ── Sheet 7: Price History Chart Data ──
    ws7 = wb.create_sheet("Price History")
    ws7.sheet_properties.tabColor = "5B9BD5"

    hist_raw = data.get("history", {})
    if hist_raw and "Close" in hist_raw:
        ws7["A1"].value = "Historical Price Data (Last 252 Trading Days)"
        ws7["A1"].font = Font(size=14, bold=True, color=cfg.COLOR_SCHEME["header_bg"])

        dates = hist_raw.get("Date", [])
        closes = hist_raw.get("Close", [])
        volumes = hist_raw.get("Volume", [])

        # Last ~1 year
        n = min(252, len(closes))
        start_idx = len(closes) - n

        row = 3
        headers = ["Date", "Close", "Volume"]
        for c, h in enumerate(headers, 1):
            ws7.cell(row=row, column=c, value=h)
        _style_header(ws7, row, 3)

        for i in range(n):
            idx = start_idx + i
            r = row + 1 + i
            date_val = dates[idx] if idx < len(dates) else ""
            if isinstance(date_val, str) and "T" in date_val:
                date_val = date_val.split("T")[0]
            ws7.cell(row=r, column=1, value=str(date_val))
            ws7.cell(row=r, column=2, value=round(closes[idx], 2) if idx < len(closes) else 0)
            ws7.cell(row=r, column=3, value=int(volumes[idx]) if idx < len(volumes) and volumes[idx] else 0)

        # Add line chart for close prices
        if n > 10:
            chart = LineChart()
            chart.title = f"{ticker} — Closing Price (1Y)"
            chart.style = 10
            chart.y_axis.title = "Price ($)"
            chart.x_axis.title = "Date"
            chart.width = 30
            chart.height = 15

            data_ref = Reference(ws7, min_col=2, min_row=row,
                                 max_row=row + n)
            chart.add_data(data_ref, titles_from_data=True)
            chart.series[0].graphicalProperties.line.width = 15000

            ws7.add_chart(chart, f"E{row}")

        ws7.column_dimensions["A"].width = 14
        ws7.column_dimensions["B"].width = 12
        ws7.column_dimensions["C"].width = 14
    else:
        ws7["A1"].value = "Price History — No Data Available"

    # Save
    wb.save(filepath)
    return filepath


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 7 — Console Report
# ═══════════════════════════════════════════════════════════════════════════════

def print_console_summary(ticker: str, data: dict, results: list,
                          composite: dict):
    """Print a concise summary to the console."""
    info = data.get("info", {})
    current_price = info.get("currentPrice") or info.get("regularMarketPrice", 0)
    fair_value = composite["composite_fair_value"]

    print("\n" + "=" * 60)
    print(f"  {ticker} — {info.get('shortName', ticker)}")
    print(f"  {info.get('sector', '')} | {info.get('industry', '')}")
    print("=" * 60)
    print(f"  Current Price:        ${current_price:>10,.2f}")
    print(f"  Composite Fair Value: ${fair_value:>10,.2f}")

    if current_price > 0:
        pct = (fair_value - current_price) / current_price
        direction = "Upside" if pct >= 0 else "Downside"
        print(f"  {direction}:               {pct:>10.1%}")

    print("-" * 60)
    print("  Method Breakdown:")
    for method, bd in composite["breakdown"].items():
        fv = bd["fair_value"]
        conf = bd["confidence"]
        note = bd.get("note", "")
        if fv > 0:
            print(f"    {method:<20s}  ${fv:>10,.2f}  (conf: {conf:.0%})")
        else:
            print(f"    {method:<20s}  {'N/A':>11s}  — {note}")

    print("-" * 60)
    if current_price > 0:
        pct = (fair_value - current_price) / current_price
        if pct > 0.15:
            print("  >> UNDERVALUED — Consider Buying")
        elif pct < -0.15:
            print("  >> OVERVALUED — Consider Selling")
        else:
            print("  >> FAIRLY VALUED — Hold")
    print("=" * 60 + "\n")


# ═══════════════════════════════════════════════════════════════════════════════
# SECTION 8 — Main Orchestrator
# ═══════════════════════════════════════════════════════════════════════════════

def analyze_stock(ticker: str) -> str:
    """Run full analysis for a single ticker. Returns path to Excel report."""
    ticker = ticker.upper().strip()
    print(f"\n{'━' * 60}")
    print(f"  Analyzing {ticker}...")
    print(f"{'━' * 60}")

    # Step 1: Fetch data
    print("\n[1/6] Fetching market data...")
    data = fetch_yfinance_data(ticker)

    if not data.get("info"):
        print(f"  ERROR: Could not retrieve data for {ticker}. Skipping.")
        return ""

    # Step 2: Run valuation models
    print("[2/6] Running DCF valuation...")
    dcf = dcf_valuation(data)

    print("[3/6] Running comparable companies analysis...")
    comps = comps_valuation(data, ticker)

    print("[4/6] Gathering analyst price targets...")
    analyst = analyst_valuation(data)

    print("[5/6] Running technical analysis...")
    tech = technical_valuation(data)

    print("[6/6] Analyzing sentiment...")
    sent = sentiment_valuation(data, ticker)

    results = [dcf, comps, analyst, tech, sent]

    # Composite
    composite = composite_valuation(results, cfg.VALUATION_WEIGHTS)

    # Macro
    macro = fetch_macro_environment()

    # Console output
    print_console_summary(ticker, data, results, composite)

    # Excel report
    print("Generating Excel report...")
    filepath = generate_excel_report(ticker, data, results, composite, macro)
    print(f"  Report saved to: {filepath}")

    return filepath


def main():
    parser = argparse.ArgumentParser(
        description="Stock Valuation Analyzer — multi-method fair value estimation"
    )
    parser.add_argument("tickers", nargs="*",
                        help="One or more stock tickers to analyze (e.g. AAPL MSFT)")
    args = parser.parse_args()

    tickers = args.tickers
    if not tickers:
        raw = input("Enter ticker(s) separated by spaces: ").strip()
        tickers = raw.upper().split()

    if not tickers:
        print("No tickers provided. Exiting.")
        sys.exit(1)

    print(f"\nStock Valuation Analyzer")
    print(f"Tickers: {', '.join(tickers)}")
    print(f"Date: {datetime.datetime.now():%Y-%m-%d %H:%M}\n")

    reports = []
    for t in tickers:
        path = analyze_stock(t)
        if path:
            reports.append(path)

    if reports:
        print("\n" + "=" * 60)
        print("  Generated Reports:")
        for rp in reports:
            print(f"    {rp}")
        print("=" * 60)
    else:
        print("\nNo reports were generated.")


if __name__ == "__main__":
    main()
