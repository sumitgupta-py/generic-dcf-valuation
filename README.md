# Generic DCF Valuation Model

A Python-based tool that creates a Discounted Cash Flow (DCF) valuation model in Excel.

## Features
- Historical financial data using yfinance
- CSV and Excel file input
- Financial forecasting and free cash flow calculations
- DCF valuation and sensitivity analysis
- Excel charts and dashboard

## Requirements
- Python
- pandas
- openpyxl
- yfinance

## How to Run

Install the required libraries:

```bash
python -m pip install -r requirements.txt
```

Run the model using a stock ticker:

```bash
python generic_dcf_from_scratch_millions.py --ticker AAPL --output "AAPL_DCF.xlsx"
```

Run the model using a CSV file:

```bash
python generic_dcf_from_scratch_millions.py --file "sample_historical_financials.csv" --output "Company_DCF.xlsx"
```

## Disclaimer
For educational and analytical purposes only. Verify financial data and assumptions before relying on the valuation. This project is not investment advice.