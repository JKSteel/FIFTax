# NZ FIF Tax Calculator

Calculates Foreign Investment Fund (FIF) income for a NZ tax return from
Sharesies annual reports, under both the FDR and CV methods, and produces an
Excel workbook of workings for record keeping.

## Inputs (not committed to git)

- `incomeReports/Sharesies/investment-holdings-report_<yyyy>-04-01_<yyyy>-03-31.csv`
- `incomeReports/Sharesies/transaction-report_<yyyy>-04-01_<yyyy>-03-31.csv`
- `data/hb1-daily.xlsx` — RBNZ B1 Daily exchange rates
  ([download here](https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/exchange-rates-and-the-trade-weighted-index)).
  The script errors if this file doesn't cover the tax year being calculated.

## Usage

```powershell
.\venv\Scripts\python.exe fif_tax.py                # interactive prompts
.\venv\Scripts\python.exe fif_tax.py --year 2026 --other-income 10000
```

Output workbook is written to `tax/FIF_calculation_<year>.xlsx` with Summary,
Valuations, Transactions Ledger and Quick Sales sheets.

## Setup (first time)

```powershell
python -m venv venv
.\venv\Scripts\python.exe -m pip install -r requirements.txt
```
