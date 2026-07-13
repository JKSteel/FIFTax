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

## Method notes

See the docstring at the top of `fif_tax.py` for the conversion conventions
and which withholding tax is claimed as a foreign tax credit. This is not
tax advice; confirm positions with an accountant, particularly the $50,000
de minimis threshold.

## Reconciling with fif.nz

`fifnz_replica.py` reproduces the fif.nz calculator's results to the cent
(same year/files/other income as `fif_tax.py`). fif.nz's CV differs from the
statutory calculation because its transaction import creates zero-value
"phantom" holdings for non-FIF (NZX/ASX-exempt) tickers, silently drops
holdings whose share balance goes negative (including genuine FIF sales),
grosses up already-gross dividends by the withholding tax, and caps the
foreign tax credit per holding. The FDR figure and box 17B generally agree;
CV and box 17A may not. See the replica's docstring for details.
