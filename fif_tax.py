"""
FIF (Foreign Investment Fund) income calculator for New Zealand tax returns.

Calculates FIF income under both the Fair Dividend Rate (FDR) and Comparative
Value (CV) methods from Sharesies annual reports, using RBNZ B1 daily exchange
rates for all currency conversion. Individuals may use whichever method gives
the lower income, provided the chosen method is applied to the entire FIF
portfolio for the year.

Inputs
------
1. Sharesies Investment Holdings Report  (incomeReports/Sharesies/)
2. Sharesies Transaction Report          (incomeReports/Sharesies/)
3. RBNZ B1 daily exchange rate workbook  (data/hb1-daily.xlsx)

Output
------
- Terminal summary: FIF income under both methods, recommended method,
  residual NZ tax after foreign tax credits, and IR3 box values.
- Excel workbook in tax/ with summary, valuation, transaction-ledger and
  quick-sale sheets for record keeping.

Key conventions
---------------
- Opening values are converted at the RBNZ rate on 1 April, closing values at
  31 March, and each transaction at the rate on its NZ-calendar trade date.
  Weekend/holiday dates take the preceding business day's rate (forward fill).
- FDR is computed per holding (5% of opening NZD value, rounded to the cent)
  and summed, plus any quick sale adjustment.
- Dividends and withholding tax are reported by Sharesies only as annual
  totals with no payment dates, so they are converted at the tax-year average
  of RBNZ daily rates. CV uses dividends on a cash basis (net of withholding
  tax deducted at source).
- Creditable overseas tax is the US and AU withholding tax columns only.
  The 'Foreign withholding tax' column is typically tax withheld inside a
  fund's underlying holdings, which is not directly creditable to the NZ
  investor - review any direct ADR withholding in that column manually.
- Instruments in NON_FIF_OVERRIDES are excluded from the FIF calculation
  even if Sharesies flags them as FIF (see the constant for rationale).

Method references: Income Tax Act 2007 ss EX 44-EX 61, IRD guide IR461.
"""

import argparse
import re
import sys
import warnings
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent
SHARESIES_DIR = BASE_DIR / "incomeReports" / "Sharesies"
RBNZ_FILE = BASE_DIR / "data" / "hb1-daily.xlsx"
OUTPUT_DIR = BASE_DIR / "tax"

FDR_RATE = 0.05          # Fair Dividend Rate: 5% of opening market value
DE_MINIMIS_NZD = 50_000  # FIF rules optional if total COST of FIF shares <= $50k

NZ_TZ = ZoneInfo("Pacific/Auckland")

# Instruments excluded from the FIF regime regardless of the Sharesies flag.
# Bullion-backed "structured" ETPs (e.g. Global X GOLD, Perth Mint PMGOLD and
# the GXLD product held here) legally confer a direct entitlement to physical
# metal rather than shares in a foreign company, so they fall outside the FIF
# definition. Profits on such holdings are instead usually taxable on revenue
# account when sold (gold is acquired for the dominant purpose of disposal).
NON_FIF_OVERRIDES = {
    "GXLD": "Direct entitlement to physical gold (structured product), not a FIF interest",
}

# Personal income tax brackets by tax year (year ended 31 March).
# 2025 uses IRD's composite rates for the mid-year threshold change on
# 31 July 2024. Each entry: (upper limit of bracket, marginal rate).
TAX_BRACKETS = {
    2022: [(14_000, 0.105), (48_000, 0.175), (70_000, 0.30), (180_000, 0.33), (float("inf"), 0.39)],
    2023: [(14_000, 0.105), (48_000, 0.175), (70_000, 0.30), (180_000, 0.33), (float("inf"), 0.39)],
    2024: [(14_000, 0.105), (48_000, 0.175), (70_000, 0.30), (180_000, 0.33), (float("inf"), 0.39)],
    2025: [(14_000, 0.105), (15_600, 0.1282), (48_000, 0.175), (53_500, 0.2164),
           (70_000, 0.30), (78_100, 0.3099), (180_000, 0.33), (float("inf"), 0.39)],
    2026: [(15_600, 0.105), (53_500, 0.175), (78_100, 0.30), (180_000, 0.33), (float("inf"), 0.39)],
    2027: [(15_600, 0.105), (53_500, 0.175), (78_100, 0.30), (180_000, 0.33), (float("inf"), 0.39)],
}

# Required columns in each input, used for early validation with clear errors.
HOLDINGS_REQUIRED = [
    "Investment ticker symbol", "Investment name", "Currency",
    "Starting investment dollar value", "Ending investment dollar value",
    "Starting shareholding", "Ending shareholding",
    "Dollar value of shares purchased (including the value of transferred shares)",
    "Dollar value of shares sold (including the value of transferred shares)",
    "Greatest number of shares held", "Number of shares purchased", "Number of shares sold",
    "Dividends and distributions", "Is FIF",
    "US withholding tax (USD)", "AU withholding tax (AUD)", "Foreign withholding tax (USD)",
]
TRANSACTIONS_REQUIRED = [
    "Trade date", "Instrument code", "Instrument name", "Quantity", "Price",
    "Transaction type", "Currency", "Amount", "Transaction fee",
]


def fail(msg: str) -> None:
    """Print a clear error and exit."""
    print(f"\nERROR: {msg}", file=sys.stderr)
    sys.exit(1)


# ---------------------------------------------------------------------------
# RBNZ exchange rates
# ---------------------------------------------------------------------------

class RbnzRates:
    """Daily NZD cross rates from the RBNZ B1 workbook.

    The B1 'Data' sheet has ~5 metadata rows (description, currency name,
    notes, unit, series id) before the daily observations. Rates are quoted
    as NZD/XXX (1 NZD buys X units of foreign currency), so
    NZD value = foreign amount / rate.
    """

    def __init__(self, path: Path, required_through: pd.Timestamp):
        if not path.exists():
            fail(f"RBNZ rate file not found: {path}")

        with warnings.catch_warnings():
            warnings.simplefilter("ignore")  # openpyxl style warning on RBNZ files
            try:
                raw = pd.read_excel(path, sheet_name="Data", header=None)
            except ValueError:
                fail(f"'Data' sheet not found in {path.name} - is this the B1 Daily workbook?")

        # Locate the 'Unit' row, which labels each column NZD/USD, NZD/AUD, etc.
        unit_rows = raw.index[raw.iloc[:, 0].astype(str).str.strip() == "Unit"]
        if len(unit_rows) == 0:
            fail(f"Could not find the 'Unit' metadata row in {path.name}.")
        units = raw.iloc[unit_rows[0]].astype(str).str.strip()

        # Data starts after the 'Series Id' row.
        series_rows = raw.index[raw.iloc[:, 0].astype(str).str.strip() == "Series Id"]
        data_start = (series_rows[0] if len(series_rows) else unit_rows[0]) + 1

        # Extract every NZD/XXX column into a tidy frame indexed by date.
        rate_cols = {m.group(1): col for col, u in units.items()
                     if (m := re.fullmatch(r"NZD/(\w{3})", u))}
        if not rate_cols:
            fail(f"No NZD/XXX rate columns found in {path.name}.")

        data = raw.iloc[data_start:]
        dates = pd.to_datetime(data.iloc[:, 0], errors="coerce")
        rates = data[list(rate_cols.values())].apply(pd.to_numeric, errors="coerce")
        rates.columns = list(rate_cols.keys())
        rates.index = pd.DatetimeIndex(dates)
        rates = rates[dates.notna().to_numpy()].sort_index()

        self.last_published = rates.index.max()
        if self.last_published < required_through:
            fail(
                "RBNZ exchange rate data requires updating: latest rate is "
                f"{self.last_published:%d %b %Y} but rates through "
                f"{required_through:%d %b %Y} are needed.\n"
                "Download the current 'B1 Daily' file from "
                "https://www.rbnz.govt.nz/statistics/series/exchange-and-interest-rates/exchange-rates-and-the-trade-weighted-index"
            )

        # Reindex to every calendar day and forward-fill so weekends and
        # public holidays take the preceding business day's rate.
        full_index = pd.date_range(rates.index.min(), rates.index.max(), freq="D")
        self.daily = rates.reindex(full_index).ffill()

    def rate(self, date: pd.Timestamp, currency: str) -> float:
        """NZD/XXX rate on a calendar date (forward-filled). NZD returns 1."""
        currency = currency.upper()
        if currency == "NZD":
            return 1.0
        if currency not in self.daily.columns:
            fail(f"Currency '{currency}' not present in the RBNZ B1 file.")
        if not (self.daily.index[0] <= date <= self.daily.index[-1]):
            fail(f"No RBNZ rate available for {date:%d %b %Y} (file covers "
                 f"{self.daily.index[0]:%d %b %Y} to {self.daily.index[-1]:%d %b %Y}).")
        value = self.daily.at[date, currency]
        if pd.isna(value):
            fail(f"RBNZ rate for {currency} on {date:%d %b %Y} is missing.")
        return float(value)

    def average_rate(self, start: pd.Timestamp, end: pd.Timestamp, currency: str) -> float:
        """Mean of daily (ffilled) rates over a period.

        Used for amounts Sharesies reports only as annual totals (dividends,
        withholding tax) where no payment dates are available. IRD accepts a
        reasonable consistent conversion basis for such amounts.
        """
        currency = currency.upper()
        if currency == "NZD":
            return 1.0
        if currency not in self.daily.columns:
            fail(f"Currency '{currency}' not present in the RBNZ B1 file.")
        return float(self.daily.loc[start:end, currency].mean())


# ---------------------------------------------------------------------------
# Sharesies report loading
# ---------------------------------------------------------------------------

def find_report(kind: str, year_end: int) -> Path:
    """Locate a Sharesies report for the tax year ending 31 March `year_end`."""
    path = SHARESIES_DIR / f"{kind}_{year_end - 1}-04-01_{year_end}-03-31.csv"
    if not path.exists():
        fail(f"Sharesies report not found: {path}")
    return path


def check_columns(df: pd.DataFrame, required: list[str], name: str) -> None:
    missing = [c for c in required if c not in df.columns]
    if missing:
        fail(f"{name} is missing expected column(s): {', '.join(missing)}.\n"
             f"Columns found: {', '.join(df.columns)}")


def load_holdings(year_end: int) -> pd.DataFrame:
    path = find_report("investment-holdings-report", year_end)
    df = pd.read_csv(path)
    check_columns(df, HOLDINGS_REQUIRED, path.name)

    numeric_cols = [c for c in HOLDINGS_REQUIRED
                    if c not in ("Investment ticker symbol", "Investment name", "Currency", "Is FIF")]
    df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors="coerce").fillna(0.0)
    df["Currency"] = df["Currency"].astype(str).str.upper()

    # Split the Sharesies flag from the set actually treated as FIF here,
    # so overridden instruments remain visible in the output workbook.
    df["Sharesies FIF flag"] = df["Is FIF"].astype(str).str.strip().str.lower() == "true"
    df["Is FIF"] = df["Sharesies FIF flag"] & ~df["Investment ticker symbol"].isin(NON_FIF_OVERRIDES)
    return df


def load_transactions(year_end: int) -> pd.DataFrame:
    path = find_report("transaction-report", year_end)
    df = pd.read_csv(path)
    check_columns(df, TRANSACTIONS_REQUIRED, path.name)

    # Trade dates are UTC timestamps like '2025-04-01 23:00:00 (UTC)'.
    # Convert to the NZ calendar date, which governs both the tax year a
    # trade falls in and the RBNZ rate date to use.
    ts_utc = pd.to_datetime(
        df["Trade date"].astype(str).str.replace(r"\s*\(UTC\)\s*$", "", regex=True),
        errors="coerce", utc=True, format="mixed",
    )
    if ts_utc.isna().any():
        bad = df.loc[ts_utc.isna(), "Trade date"].head().tolist()
        fail(f"Unparseable trade date(s) in {path.name}: {bad}")
    df["Trade date (NZ)"] = ts_utc.dt.tz_convert(NZ_TZ).dt.normalize().dt.tz_localize(None)

    for col in ("Quantity", "Price", "Amount", "Transaction fee"):
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df["Currency"] = df["Currency"].astype(str).str.upper()
    df["Transaction type"] = df["Transaction type"].astype(str).str.upper()
    return df


# ---------------------------------------------------------------------------
# FIF calculations
# ---------------------------------------------------------------------------

def build_ledger(transactions: pd.DataFrame, holdings: pd.DataFrame,
                 rates: RbnzRates) -> pd.DataFrame:
    """Convert every transaction to NZD at the RBNZ rate on its NZ trade date.

    Net NZD is the actual cash flow: buys cost Amount + fee, sells return
    Amount - fee. Purchase costs therefore include brokerage and sale
    proceeds are net of brokerage, consistent with the CV method's
    treatment of expenditure and gains.
    """
    fif_tickers = set(holdings.loc[holdings["Is FIF"], "Investment ticker symbol"])

    ledger = transactions.copy()
    ledger["RBNZ rate (NZD/ccy)"] = [
        rates.rate(d, c) for d, c in zip(ledger["Trade date (NZ)"], ledger["Currency"])
    ]
    sign = ledger["Transaction type"].map({"BUY": 1, "SELL": -1})
    if sign.isna().any():
        bad = ledger.loc[sign.isna(), "Transaction type"].unique().tolist()
        fail(f"Unrecognised transaction type(s): {bad} (expected BUY or SELL).")

    ledger["Net amount (ccy)"] = ledger["Amount"] + sign * ledger["Transaction fee"]
    ledger["Net NZD"] = ledger["Net amount (ccy)"] / ledger["RBNZ rate (NZD/ccy)"]
    ledger["Counted in FIF"] = ledger["Instrument code"].isin(fif_tickers)
    return ledger


def build_fif_table(holdings: pd.DataFrame, ledger: pd.DataFrame, rates: RbnzRates,
                    opening_date: pd.Timestamp, closing_date: pd.Timestamp) -> pd.DataFrame:
    """Per-holding NZD workings for both FIF methods.

    Includes every holding Sharesies flags as FIF; rows overridden by
    NON_FIF_OVERRIDES are shown but contribute nothing to either method.
    """
    tbl = holdings[holdings["Sharesies FIF flag"]].copy()
    tbl["FIF treatment"] = tbl["Investment ticker symbol"].map(
        lambda t: f"Excluded: {NON_FIF_OVERRIDES[t]}" if t in NON_FIF_OVERRIDES else "FIF"
    )

    tbl["Opening rate"] = tbl["Currency"].map(lambda c: rates.rate(opening_date, c))
    tbl["Closing rate"] = tbl["Currency"].map(lambda c: rates.rate(closing_date, c))
    tbl["Opening value (NZD)"] = tbl["Starting investment dollar value"] / tbl["Opening rate"]
    tbl["Closing value (NZD)"] = tbl["Ending investment dollar value"] / tbl["Closing rate"]

    # FDR income is computed per holding and rounded to the cent, matching
    # how broker/IRD schedules present it.
    tbl["FDR income (NZD)"] = (FDR_RATE * tbl["Opening value (NZD)"]).round(2)

    # Purchases (incl. fees) and sale proceeds (net of fees) from the ledger.
    fif_txns = ledger[ledger["Counted in FIF"]]
    buys = fif_txns[fif_txns["Transaction type"] == "BUY"].groupby("Instrument code")["Net NZD"].sum()
    sells = fif_txns[fif_txns["Transaction type"] == "SELL"].groupby("Instrument code")["Net NZD"].sum()
    tbl["Purchases (NZD)"] = tbl["Investment ticker symbol"].map(buys).fillna(0.0)
    tbl["Sales (NZD)"] = tbl["Investment ticker symbol"].map(sells).fillna(0.0)

    # Dividends and withholding tax are annual totals with no payment dates,
    # converted at the tax-year average rate. Withholding amounts are in the
    # currency named in each column header, not the holding's own currency.
    avg = {c: rates.average_rate(opening_date, closing_date, c)
           for c in set(tbl["Currency"]) | {"USD", "AUD"}}
    tbl["Dividends gross (NZD)"] = tbl["Dividends and distributions"] / tbl["Currency"].map(avg)
    tbl["Withholding tax (NZD)"] = (
        (tbl["US withholding tax (USD)"] + tbl["Foreign withholding tax (USD)"]) / avg["USD"]
        + tbl["AU withholding tax (AUD)"] / avg["AUD"]
    )
    tbl["Dividends net (NZD)"] = tbl["Dividends gross (NZD)"] - tbl["Withholding tax (NZD)"]

    # Only US and AU withholding tax is treated as a creditable foreign tax;
    # the 'Foreign withholding tax' column is generally fund-level tax.
    tbl["Creditable overseas tax (NZD)"] = (
        tbl["US withholding tax (USD)"] / avg["USD"]
        + tbl["AU withholding tax (AUD)"] / avg["AUD"]
    )

    # CV per holding: (closing value + sale proceeds + net dividends)
    # less (opening value + purchase costs). The statutory floor at zero
    # applies to the portfolio total, not per holding.
    tbl["CV contribution (NZD)"] = (
        tbl["Closing value (NZD)"] + tbl["Sales (NZD)"] + tbl["Dividends net (NZD)"]
        - tbl["Opening value (NZD)"] - tbl["Purchases (NZD)"]
    )

    # Overridden instruments carry no FIF income and no credit.
    excluded = ~tbl["Is FIF"]
    zero_cols = ["FDR income (NZD)", "CV contribution (NZD)", "Creditable overseas tax (NZD)",
                 "Dividends gross (NZD)", "Dividends net (NZD)", "Withholding tax (NZD)",
                 "Purchases (NZD)", "Sales (NZD)", "Opening value (NZD)", "Closing value (NZD)"]
    tbl.loc[excluded, zero_cols] = 0.0
    tbl[zero_cols] = tbl[zero_cols].astype(float)  # keeps sums numeric when empty
    return tbl


def quick_sale_adjustment(holdings: pd.DataFrame, ledger: pd.DataFrame) -> tuple[float, pd.DataFrame]:
    """FDR quick sale adjustment (s EX 52) using the peak holding method.

    Applies to FIF shares both purchased and sold within the same tax year.
    Quick-sale share count = peak holding less the greater of the opening
    and closing holdings. The adjustment per holding is the lesser of:
      (a) 5% of the average NZD cost of the quick-sold shares, and
      (b) the actual NZD gain realised on them (floored at zero),
    and is added on top of the base FDR amount.
    """
    rows = []
    for _, h in holdings[holdings["Is FIF"]].iterrows():
        ticker = h["Investment ticker symbol"]
        if h["Number of shares purchased"] <= 0 or h["Number of shares sold"] <= 0:
            continue  # quick sales need both a purchase and a sale in the year

        peak = h["Greatest number of shares held"]
        qs_shares = peak - max(h["Starting shareholding"], h["Ending shareholding"])
        if qs_shares <= 1e-9:
            continue

        txns = ledger[(ledger["Instrument code"] == ticker) & ledger["Counted in FIF"]]
        buys = txns[txns["Transaction type"] == "BUY"]
        sells = txns[txns["Transaction type"] == "SELL"]
        if buys["Quantity"].sum() <= 0 or sells["Quantity"].sum() <= 0:
            continue

        avg_cost = buys["Net NZD"].sum() / buys["Quantity"].sum()        # incl. fees
        avg_proceeds = sells["Net NZD"].sum() / sells["Quantity"].sum()  # net of fees
        qs_cost = qs_shares * avg_cost
        actual_gain = max(0.0, qs_shares * (avg_proceeds - avg_cost))
        adjustment = min(FDR_RATE * qs_cost, actual_gain)

        rows.append({
            "Ticker": ticker,
            "Investment name": h["Investment name"],
            "Quick sale shares": qs_shares,
            "Average cost/share (NZD)": avg_cost,
            "Average proceeds/share (NZD)": avg_proceeds,
            "Cost of quick sale shares (NZD)": qs_cost,
            "5% of average cost (NZD)": FDR_RATE * qs_cost,
            "Actual gain (NZD)": actual_gain,
            "Adjustment (lesser of the two, NZD)": adjustment,
        })

    detail = pd.DataFrame(rows)
    total = detail["Adjustment (lesser of the two, NZD)"].sum() if not detail.empty else 0.0
    return round(total, 2), detail


def progressive_tax(income: float, brackets: list[tuple[float, float]]) -> float:
    """NZ personal income tax on `income` under the given bracket table."""
    tax, lower = 0.0, 0.0
    for upper, rate in brackets:
        if income <= lower:
            break
        tax += (min(income, upper) - lower) * rate
        lower = upper
    return tax


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def available_years() -> list[int]:
    """Tax years (ending 31 March) that have both Sharesies reports present."""
    years = []
    for f in sorted(SHARESIES_DIR.glob("investment-holdings-report_*_*.csv")):
        m = re.search(r"_(\d{4})-03-31\.csv$", f.name)
        if m and (SHARESIES_DIR / f.name.replace("investment-holdings-report",
                                                 "transaction-report")).exists():
            years.append(int(m.group(1)))
    return years


def prompt_year(years: list[int]) -> int:
    print("\nAvailable tax years:")
    for i, y in enumerate(years, 1):
        print(f"  {i}. 1 April {y - 1} - 31 March {y}")
    choice = input(f"Select tax year [1-{len(years)}, default {len(years)}]: ").strip()
    if not choice:
        return years[-1]
    try:
        return years[int(choice) - 1]
    except (ValueError, IndexError):
        fail(f"Invalid selection '{choice}'.")


def prompt_other_income() -> float:
    raw = input("\nTotal other taxable income for the year (NZD, e.g. salary/wages): $").strip()
    try:
        return float(raw.replace(",", "").replace("$", ""))
    except ValueError:
        fail(f"Could not parse income amount '{raw}'.")


def main() -> None:
    parser = argparse.ArgumentParser(description="NZ FIF income calculator (FDR and CV methods).")
    parser.add_argument("--year", type=int, help="Tax year ending 31 March of this year, e.g. 2026.")
    parser.add_argument("--other-income", type=float, help="Total other taxable income (NZD).")
    args = parser.parse_args()

    years = available_years()
    if not years:
        fail(f"No Sharesies report pairs found in {SHARESIES_DIR}.")
    year_end = args.year if args.year else prompt_year(years)
    if year_end not in years:
        fail(f"No reports found for the year ended 31 March {year_end}. Available: {years}")
    if year_end not in TAX_BRACKETS:
        fail(f"No tax bracket table defined for the {year_end} tax year - add it to TAX_BRACKETS.")

    other_income = args.other_income if args.other_income is not None else prompt_other_income()
    if other_income < 0:
        fail("Other income cannot be negative.")

    opening_date = pd.Timestamp(year_end - 1, 4, 1)
    closing_date = pd.Timestamp(year_end, 3, 31)
    year_label = f"{year_end - 1}-{year_end}"

    if closing_date > pd.Timestamp.now(tz=NZ_TZ).tz_localize(None):
        fail(f"The {year_label} tax year has not ended yet - closing values are not available.")

    # --- Load inputs --------------------------------------------------------
    rates = RbnzRates(RBNZ_FILE, required_through=closing_date)
    holdings = load_holdings(year_end)
    transactions = load_transactions(year_end)

    if not holdings["Is FIF"].any():
        print(f"\nNote: no holdings are treated as FIF in {year_label} - "
              "FIF income is $0 for this year.")

    # --- Per-holding workings -----------------------------------------------
    ledger = build_ledger(transactions, holdings, rates)
    fif_tbl = build_fif_table(holdings, ledger, rates, opening_date, closing_date)

    opening_nzd = fif_tbl["Opening value (NZD)"].sum()
    closing_nzd = fif_tbl["Closing value (NZD)"].sum()
    purchases_nzd = fif_tbl["Purchases (NZD)"].sum()
    sales_nzd = fif_tbl["Sales (NZD)"].sum()
    dividends_net_nzd = fif_tbl["Dividends net (NZD)"].sum()
    overseas_tax_nzd = fif_tbl["Creditable overseas tax (NZD)"].sum()

    # --- FIF income under each method ---------------------------------------
    fdr_base = fif_tbl["FDR income (NZD)"].sum()
    qs_total, qs_detail = quick_sale_adjustment(holdings, ledger)
    fdr_income = round(fdr_base + qs_total, 2)

    cv_raw = fif_tbl["CV contribution (NZD)"].sum()
    cv_income = round(max(0.0, cv_raw), 2)  # CV cannot be negative for individuals

    use_fdr = fdr_income <= cv_income
    fif_income = min(fdr_income, cv_income)
    method = "FDR (Fair Dividend Rate)" if use_fdr else "CV (Comparative Value)"

    # --- Tax on the FIF income ----------------------------------------------
    brackets = TAX_BRACKETS[year_end]
    nz_tax_on_fif = (progressive_tax(other_income + fif_income, brackets)
                     - progressive_tax(other_income, brackets))

    # Foreign tax credit is limited to the NZ tax payable on the FIF income.
    credit = round(min(overseas_tax_nzd, nz_tax_on_fif), 2)
    net_tax = round(nz_tax_on_fif, 2) - credit

    excluded_rows = fif_tbl[~fif_tbl["Is FIF"]]

    # --- Terminal summary ----------------------------------------------------
    w = 64
    print("\n" + "=" * w)
    print(f"  NZ FIF INCOME - TAX YEAR 1 APRIL {year_end - 1} TO 31 MARCH {year_end}")
    print("=" * w)
    print(f"  FIF holdings:                        {int(fif_tbl['Is FIF'].sum())} investments")
    print(f"  Opening market value (1 Apr):        ${opening_nzd:>12,.2f} NZD")
    print(f"  Closing market value (31 Mar):       ${closing_nzd:>12,.2f} NZD")
    print(f"  Purchases during year (incl. fees):  ${purchases_nzd:>12,.2f} NZD")
    print(f"  Sales during year (net of fees):     ${sales_nzd:>12,.2f} NZD")
    print(f"  Dividends received (net of WHT):     ${dividends_net_nzd:>12,.2f} NZD")
    print("-" * w)
    print(f"  FDR income (5% x opening value):     ${fdr_base:>12,.2f}")
    print(f"  FDR quick sale adjustment:           ${qs_total:>12,.2f}")
    print(f"  FDR method total:                    ${fdr_income:>12,.2f}")
    print(f"  CV method total:                     ${cv_income:>12,.2f}"
          + (f"  (raw {cv_raw:,.2f}, floored at 0)" if cv_raw < 0 else ""))
    print("-" * w)
    print(f"  RECOMMENDED METHOD: {method} (lower income)")
    print(f"  FIF income to declare:               ${fif_income:>12,.2f}")
    print("-" * w)
    print(f"  NZ tax attributable to FIF income:   ${nz_tax_on_fif:>12,.2f}")
    print(f"  Overseas tax already paid (credit):  ${-credit:>12,.2f}")
    print(f"  Net NZ tax to pay on FIF income:     ${net_tax:>12,.2f}")
    print("-" * w)
    print("  IR3 return:")
    print(f"    Box 17A (overseas tax paid):       ${credit:>12,.2f}")
    print(f"    Box 17B (overseas/FIF income):     ${fif_income:>12,.2f}")
    print("=" * w)
    print("  Notes:")
    for _, r in excluded_rows.iterrows():
        print(f"  - {r['Investment ticker symbol']} ({r['Investment name']}) is excluded from FIF:")
        print(f"    {NON_FIF_OVERRIDES[r['Investment ticker symbol']]}.")
        print("    Any profit when it is sold is likely taxable separately as income")
        print("    from gold held for disposal - not calculated here.")
    print(f"  - If the total COST of your FIF investments never exceeded ${DE_MINIMIS_NZD:,},")
    print("    the FIF rules are optional (de minimis, s CQ 5). Cost history is not")
    print("    in these reports, so confirm this yourself.")
    print("  - The chosen method must be applied to ALL FIF holdings for the year.")
    print("  - Dividends and withholding tax are annual totals in the Sharesies")
    print("    report, converted at the tax-year average RBNZ rate; converting at")
    print("    actual payment dates may shift results by a few dollars.")
    print("  - 'Foreign withholding tax' (fund-level) is not claimed as a credit;")
    print("    any direct ADR withholding in that column may be claimable - review.")
    if overseas_tax_nzd - credit > 0.005:
        print(f"  - Overseas tax of ${overseas_tax_nzd - credit:,.2f} exceeded the NZ tax on this")
        print("    income and cannot be credited (credit capped at NZ tax payable).")

    # --- Excel export ---------------------------------------------------------
    OUTPUT_DIR.mkdir(exist_ok=True)
    out_path = OUTPUT_DIR / f"FIF_calculation_{year_label}.xlsx"

    summary = pd.DataFrame(
        [
            ("Tax year", f"1 April {year_end - 1} - 31 March {year_end}"),
            ("Opening FIF market value (NZD)", round(opening_nzd, 2)),
            ("Closing FIF market value (NZD)", round(closing_nzd, 2)),
            ("Purchases (NZD, incl. fees)", round(purchases_nzd, 2)),
            ("Sales (NZD, net of fees)", round(sales_nzd, 2)),
            ("Dividends (NZD, net of withholding tax)", round(dividends_net_nzd, 2)),
            ("FDR base income (5% x opening, per holding)", round(fdr_base, 2)),
            ("FDR quick sale adjustment", round(qs_total, 2)),
            ("FDR income", fdr_income),
            ("CV income (floored at 0)", cv_income),
            ("Recommended method", method),
            ("FIF income to declare (IR3 Box 17B)", fif_income),
            ("Other taxable income entered", round(other_income, 2)),
            ("NZ tax attributable to FIF income", round(nz_tax_on_fif, 2)),
            ("Overseas tax credit (IR3 Box 17A)", credit),
            ("Net NZ tax to pay on FIF income", round(net_tax, 2)),
            ("", ""),
            ("Assumptions", ""),
            ("FX source", "RBNZ B1 daily (weekends/holidays forward-filled)"),
            ("Transactions", "Converted at RBNZ rate on NZ-calendar trade date"),
            ("Dividends / withholding tax", "Annual totals converted at tax-year average rate"),
            ("Creditable overseas tax", "US + AU withholding tax columns only"),
        ] + [
            (f"Excluded from FIF: {r['Investment ticker symbol']}",
             NON_FIF_OVERRIDES[r["Investment ticker symbol"]])
            for _, r in excluded_rows.iterrows()
        ],
        columns=["Item", "Value"],
    )

    val_cols = [
        "Investment ticker symbol", "Investment name", "Currency", "FIF treatment",
        "Starting investment dollar value", "Opening rate", "Opening value (NZD)",
        "Ending investment dollar value", "Closing rate", "Closing value (NZD)",
        "FDR income (NZD)", "Purchases (NZD)", "Sales (NZD)",
        "Dividends gross (NZD)", "Withholding tax (NZD)", "Dividends net (NZD)",
        "CV contribution (NZD)", "Creditable overseas tax (NZD)",
    ]
    val_out = fif_tbl[val_cols].rename(columns={
        "Starting investment dollar value": f"Opening value ({opening_date:%d %b %Y}, local ccy)",
        "Ending investment dollar value": f"Closing value ({closing_date:%d %b %Y}, local ccy)",
    })

    ledger_cols = ["Trade date (NZ)", "Instrument code", "Instrument name", "Transaction type",
                   "Quantity", "Price", "Currency", "Amount", "Transaction fee",
                   "RBNZ rate (NZD/ccy)", "Net NZD", "Counted in FIF"]
    ledger_out = ledger[ledger_cols].sort_values("Trade date (NZ)").rename(
        columns={"Amount": "Amount (local ccy)",
                 "Net NZD": "Net NZD (buys incl. fee, sells net of fee)"}
    )

    if qs_detail.empty:
        qs_out = pd.DataFrame(
            {"Quick sale calculation": ["No FIF holdings were both purchased and sold "
                                        "within this tax year - no adjustment applies."]}
        )
    else:
        qs_out = qs_detail

    with pd.ExcelWriter(out_path, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        for sheet, frame in (("Summary", summary), ("Valuations", val_out),
                             ("Transactions Ledger", ledger_out), ("Quick Sales", qs_out)):
            frame.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            for i, col in enumerate(frame.columns):  # readable column widths
                longest = frame[col].astype(str).str.len().max()
                width = max(len(str(col)), 0 if pd.isna(longest) else int(longest)) + 2
                ws.set_column(i, i, min(width, 48))

    print(f"\n  Workbook saved: {out_path.relative_to(BASE_DIR)}\n")


if __name__ == "__main__":
    main()
