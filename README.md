# Finance Dashboard

A personal investment tracking dashboard built with Python and Flask, featuring real-time B3 (Brazilian Stock Exchange) quotes, dividend history, and automated insights.

> Built to track my own investment portfolio — covers variable income (stocks + FIIs), fixed income (Selic), and crypto, with data going back to June 2024.

---

## Features

- **Real-time B3 quotes** — stocks, FIIs, and ETFs fetched from Yahoo Finance via `yfinance`, with 15-minute in-memory cache and parallel fetching (ThreadPoolExecutor)
- **Live Selic rate** — pulled from the Banco Central do Brasil public API and annualized via `(1 + daily)^252 − 1`
- **Monthly performance charts** — portfolio value evolution, META vs REAL yield, monthly income breakdown (Selic + FII dividends + stock dividends)
- **Asset allocation** — doughnut chart and table showing current portfolio composition
- **Automated insights engine** — 8 rule-based insights covering emergency reserve status, opportunity reserve, crypto allocation target (10%), META vs REAL trend over 6 months, best FII by accumulated yield, and contribution/reinvestment suggestions
- **52-week range bars** — visual price position within the yearly range for each asset
- **Dual-format Excel parser** — auto-detects new organized format (`finance_data.xlsx`) or falls back to legacy format (`finance.xlsx`)

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3.13 + Flask 3.x |
| Data parsing | openpyxl |
| Market data | yfinance + BCB REST API |
| Frontend | HTML/CSS/JS — Chart.js 4.x + Bootstrap 5 |
| Concurrency | ThreadPoolExecutor (parallel ticker fetches) |
| Caching | In-memory dict with TTL (15 min quotes, 1 h Selic) |

## Architecture

```
finance_data.xlsx  ←──── you fill this in Excel
       │
       ▼
 data_parser.py   ←──── reads & normalizes data
       │
       ▼
    app.py        ←──── Flask API + insight engine
  /api/data            returns parsed portfolio data
  /api/market          returns real-time quotes + Selic
  /api/insights        returns rule-based recommendations
       │
       ▼
 templates/index.html  ←──── Chart.js dark dashboard
```

The frontend fetches all three endpoints asynchronously on load. Market data is loaded separately (after the main dashboard renders) to avoid blocking the UI while Yahoo Finance responds.

## Excel Structure

The `finance_data.xlsx` workbook has four sheets:

- **Mensal** — one row per month: portfolio value, invested amount, gain/loss, Selic, FII dividends, stock dividends, META/REAL yield
- **Dividendos** — one row per income event: asset, type (FII/Ação/SELIC/ETF/Cripto), amount, payment date
- **Carteira_Atual** — current positions with Yahoo Finance tickers and portfolio % 
- **Resumo** — aggregated totals: variable/fixed/crypto allocation, total profit, contributions, reserves

> The Excel files are excluded from this repo via `.gitignore` since they contain personal financial data. Run `python create_template.py` to generate a blank template with the correct structure.

## Getting Started

```bash
# Install dependencies
pip install -r requirements.txt

# Generate the Excel template (or bring your own finance_data.xlsx)
python create_template.py

# Start the server
python app.py
# → http://localhost:5000
```

The app auto-reloads data from Excel on every browser refresh — no server restart needed after updating the spreadsheet.

## Skills Demonstrated

- **Backend:** RESTful API design with Flask, data normalization with Unicode handling (`unicodedata.normalize`), multi-format file parsing with auto-detection
- **Data engineering:** ETL pipeline from Excel → Python dicts → JSON API, accent-stripping for robust label matching, concurrent I/O with ThreadPoolExecutor
- **Frontend:** Vanilla JS async/await with Promise.all, Chart.js 4.x (line, bar, doughnut, stacked bar), responsive dark UI with CSS custom properties
- **External integrations:** Yahoo Finance (yfinance), Banco Central do Brasil REST API
- **Software design:** In-memory caching with TTL, fallback strategies, separation of data parsing / business logic / presentation
