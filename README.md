# MIS2800
MIS 2800 GitHub

## Stock Valuation Analyzer

A comprehensive stock analysis tool that combines six valuation methods into a single weighted fair-value estimate and exports a formatted Excel report.

### Valuation Methods

| Method | Weight | Description |
|---|---|---|
| DCF (Discounted Cash Flow) | 30% | Projects free cash flows and discounts to present value using WACC |
| Comparable Companies | 20% | Relative valuation using peer-company multiples (P/E, EV/EBITDA, P/S, PEG, P/B) |
| Analyst Price Targets | 20% | Consensus analyst target prices from Wall Street coverage |
| Technical Analysis | 10% | Moving averages, RSI, MACD, Bollinger Bands |
| Sentiment Analysis | 10% | News headline sentiment and fundamental quality signals |
| Seasonal Analysis | 10% | Historical monthly return patterns and forward 3-month outlook |

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

3. Edit defaults in `config.py` if desired (valuation weights, DCF assumptions, etc.).

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

- **Console**: Prints a summary with the composite fair value, method breakdown, and a buy/hold/sell verdict.
- **Excel** (`output/` folder): A multi-sheet workbook containing:
  - Executive Summary with key metrics and verdict
  - Valuation Breakdown comparing all six methods
  - DCF Analysis detail with sensitivity matrix (WACC vs Terminal Growth Rate)
  - Technical Analysis indicators and signals
  - Macro Environment (Fed Funds, 10Y Treasury, CPI, etc.)
  - Financial Ratios with color-coded benchmarks (liquidity, profitability, leverage)
  - Seasonal Analysis with monthly return patterns and win rates
  - Company Profile with financials and business summary
  - Price History with a chart of the last year of closing prices

### Project Files

| File | Purpose |
|---|---|
| `stock_valuation_analyzer.py` | Main script — data fetching, valuation models, Excel report |
| `config.py` | API keys, valuation weights, DCF defaults, ratio benchmarks, formatting constants |
| `requirements.txt` | Python package dependencies |
