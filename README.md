# MIS2800
MIS 2800 GitHub

## Stock Valuation Analyzer — Enhanced Professional Edition

A comprehensive stock analysis tool that combines 7 valuation methods, quality scoring, scenario analysis, and risk assessment into a 20-sheet Excel report.

### Valuation Models (Fair Value Weights)

| Method | Weight | Description |
|---|---|---|
| DCF (Discounted Cash Flow) | 50% | Projects free cash flows and discounts to present value using WACC |
| Comparable Companies | 40% | Relative valuation using peer-company multiples (P/E, EV/EBITDA, P/S, PEG, P/B) |
| Historical P/E | 10% | Fair value based on historical average P/E applied to current/forward earnings |

### Additional Analyses (Informational)

| Analysis | Description |
|---|---|
| Analyst Price Targets | Consensus analyst target prices from Wall Street coverage |
| Technical Analysis | Moving averages, RSI, MACD, Bollinger Bands |
| Sentiment Analysis | News headline sentiment and fundamental quality signals |
| Seasonal Analysis | Historical monthly return patterns and forward 3-month outlook |
| Quality Score | 0-100 composite score across profitability, growth, moat, health, management, innovation |
| Scenario Analysis | Bear / Base / Bull weighted fair value estimates |
| Risk Assessment | Valuation, volatility, leverage, liquidity, ownership, and growth risks |

### Setup

1. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. (Optional) Set API keys as environment variables for enhanced data:
   ```bash
   export ALPHA_VANTAGE_KEY="your_key"    # alphavantage.co (free tier)
   export FMP_API_KEY="your_key"           # financialmodelingprep.com (free tier)
   export FRED_API_KEY="your_key"          # fred.stlouisfed.org (free tier)
   export NEWS_API_KEY="your_key"          # newsapi.org (free tier)
   ```
   Without API keys the tool still works using yfinance data. API keys add peer-company lookup, macro indicators, and news sentiment.

3. Edit defaults in `config.py` if desired (valuation weights, DCF assumptions, quality thresholds, etc.).

### Usage

```bash
# Interactive prompt
python stock_valuation_analyzer.py

# Single ticker
python stock_valuation_analyzer.py AAPL

# Multiple tickers
python stock_valuation_analyzer.py AAPL MSFT GOOGL
```

### Output

- **Console**: Prints a detailed summary with composite fair value, method breakdown, quality score, scenario analysis, and a 5-tier verdict (Strong Buy / Buy / Hold / Sell / Strong Sell).
- **Excel** (`output/` folder): A 20-sheet workbook containing:
  1. Executive Summary — key metrics, verdict, upside/downside
  2. Valuation Breakdown — all methods compared side-by-side
  3. DCF Analysis — assumptions, projected FCFs, WACC vs TGR sensitivity matrix
  4. Technical Analysis — indicators, signals, bullish/bearish scoring
  5. Macro Environment — Fed Funds, 10Y Treasury, CPI, GDP, VIX, etc.
  6. Financial Ratios — color-coded benchmarks with explanations
  7. Seasonal Analysis — monthly return patterns and win rates
  8. Company Profile — financials, business summary, key stats
  9. Price History — 1Y closing prices with chart
  10. Comps Detail — peer multiples breakdown and implied values
  11. Analyst Targets — mean/median/high/low targets, recommendation summary
  12. Historical P/E — trailing/forward EPS, historical average P/E valuation
  13. Sentiment Analysis — headline count, sentiment score, adjustment factor
  14. Quality Score — 6-dimension scoring with composite grade
  15. Scenario Analysis — bear/base/bull fair values with probability weights
  16. Growth-Margin Sensitivity — revenue growth vs EBIT margin matrix
  17. Income Statement — annual income statement data
  18. Balance Sheet — annual balance sheet data
  19. Cash Flow — annual cash flow statement data
  20. Peer Comparison — side-by-side metrics vs peer companies
  21. Risk Assessment — color-coded risk factors (valuation, volatility, leverage, etc.)

### Project Files

| File | Purpose |
|---|---|
| `stock_valuation_analyzer.py` | Main script — 7 valuation models, quality scoring, scenario analysis, 20-sheet Excel report (~2,660 lines) |
| `config.py` | API keys, valuation weights, DCF/WACC assumptions, quality thresholds, peer tickers, ratio benchmarks, scenario weights |
| `requirements.txt` | Python package dependencies |
