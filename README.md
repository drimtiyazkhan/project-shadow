# HFT Bulk Deal Tracker

A free, private web tool for NSE traders to track HFT bulk deal activity on TradingView.

## What it does

- Upload NSE bulk deal CSV files
- Automatically detects activity from 12 HFT traders
- Generates a Pine Script indicator ready to paste into TradingView
- Arrows appear on every NSE chart where HFT traders were active
- Arrow size scales with number of traders (1=tiny → 12=enormous)
- **Remembers your previous data** — just upload new week's file each time

## Live Site

👉 **https://YOUR_USERNAME.github.io/hft-bulk-deal-tracker**

## HFT Traders Tracked

| Code | Trader |
|------|--------|
| J | Jump Trading Financial India Private Limited |
| Q | QE Securities LLP |
| N | Junomoneta Finsol Private Limited |
| K | NK Securities Research Private Limited |
| H | HRTI Private Limited |
| G | Graviton Research Capital LLP |
| S | Share India Securities Limited |
| E | Elixir Wealth Management Private Limited |
| D | D3 Stock Vision LLP |
| M | Goldmine Stocks Private Limited |
| C | Microcurves Trading Private Limited |
| MS | Musigma Securities |

## How to use

1. Go to the live site
2. Upload your NSE bulk deal CSV
3. Click **Generate Pine Script**
4. Click **Copy to Clipboard**
5. Open TradingView Pine Editor → paste → Save → Add to chart

## Privacy

Everything runs in your browser. No data is sent to any server.

## How to host on GitHub Pages (free)

1. Create a GitHub account at github.com
2. Create a new repository named `hft-bulk-deal-tracker`
3. Upload `index.html` and `generator.html`
4. Go to Settings → Pages → Source: main branch → Save
5. Your site is live at `https://YOUR_USERNAME.github.io/hft-bulk-deal-tracker`
