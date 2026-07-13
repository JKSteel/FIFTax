"""
Replica of the fif.nz calculator engine, for reconciling its output against
fif_tax.py. Reverse-engineered from fif.nz's client-side JavaScript bundle
(the calculator runs entirely in the browser, with RBNZ HB1 daily rates
embedded - the same series as data/hb1-daily.xlsx).

Run:  python fifnz_replica.py --year 2026 --other-income 10000

Reproduces fif.nz's numbers to the cent, including behaviour that differs
from the statutory FIF calculation in fif_tax.py:

1. The transaction-report import creates a holding for EVERY ticker with a
   buy/sell row - including NZX and ASX-exempt shares that the holdings
   report correctly skips as non-FIF. These "phantom" holdings have zero
   opening/closing value, so their purchases depress the CV method total
   and their sales inflate it.
2. Any holding whose running share balance goes negative (sells with no
   recorded opening, e.g. non-FIF shares sold during the year, or FIF
   shares received via a corporate action) is silently dropped as a "short
   position" - removing genuine FIF sale proceeds from CV.
3. Dividends are converted at the midpoint date of the report period and
   withholding tax at 1 October; dividends in CV are grossed up by adding
   withholding tax to the (already gross) Sharesies dividends column.
4. The foreign tax credit is capped per holding at the average tax rate
   times that holding's FIF income, so withholding on a holding with no
   opening value (zero FDR income) earns no credit at all.
"""

import argparse
import csv
from datetime import date
from decimal import Decimal, ROUND_HALF_EVEN

from fif_tax import RbnzRates, TAX_BRACKETS, RBNZ_FILE, SHARESIES_DIR

import pandas as pd

# fif.nz maps withholding-tax columns to holdings by the holding's currency.
WHT_COLUMNS = {"USD": ["US withholding tax (USD)", "Foreign withholding tax (USD)"],
               "AUD": ["AU withholding tax (AUD)"]}


def s0(x: float) -> float:
    """fif.nz rounds each converted amount to 2 dp, banker's rounding."""
    return float(Decimal(repr(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_EVEN))


class Engine:
    def __init__(self, year_end: int, rates: RbnzRates):
        self.y0 = year_end - 1
        self.rates = rates

    def rate(self, ccy: str, dstr: str) -> float:
        # fif.nz walks back day-by-day to the last published daily rate;
        # equivalent to the forward-filled RBNZ series.
        return self.rates.rate(pd.Timestamp(dstr), ccy)

    def convert(self, amount: float, ccy: str, dstr: str, round_: bool = True) -> float:
        v = amount if ccy == "NZD" else amount / self.rate(ccy, dstr)
        return s0(v) if round_ else v

    def load(self) -> list[dict]:
        y0 = self.y0
        span = f"{y0}-04-01_{y0 + 1}-03-31"
        # Synthesised dividend transactions are dated at the midpoint of the
        # report period, read from the file name.
        div_date = (date(y0, 4, 1) + (date(y0 + 1, 3, 31) - date(y0, 4, 1)) / 2).isoformat()

        holdings: dict[tuple, dict] = {}
        with open(SHARESIES_DIR / f"investment-holdings-report_{span}.csv",
                  encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                t = row["Investment ticker symbol"].strip().upper()
                ccy = row["Currency"].strip().upper()
                if not t or row.get("Is FIF", "").strip().upper() == "FALSE":
                    continue  # holdings report import skips non-FIF rows
                wht = sum(float(row.get(c) or 0) for c in WHT_COLUMNS.get(ccy, []))
                g = float(row["Dividends and distributions"] or 0)
                txns = ([{"date": div_date, "type": "dividend", "shares": 1.0,
                          "price": g, "comm": 0.0}] if g > 0 else [])
                holdings[(t, ccy)] = {
                    "ticker": t, "ccy": ccy,
                    "openShares": float(row["Starting shareholding"] or 0),
                    "openMV": float(row["Starting investment dollar value"] or 0),
                    "closeShares": float(row["Ending shareholding"] or 0),
                    "closeMV": float(row["Ending investment dollar value"] or 0),
                    "wht": wht, "txns": txns,
                }

        # The transaction import keeps every BUY/SELL row regardless of FIF
        # status; unmatched tickers become zero-value "phantom" holdings.
        with open(SHARESIES_DIR / f"transaction-report_{span}.csv", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                typ = row["Transaction type"].strip().upper()
                if typ not in ("BUY", "SELL"):
                    continue
                t = row["Instrument code"].strip().upper()
                ccy = row["Currency"].strip().upper()
                fee = abs(float(row["Transaction fee"] or 0))
                h = holdings.setdefault((t, ccy), {
                    "ticker": t, "ccy": ccy, "openShares": 0.0, "openMV": 0.0,
                    "closeShares": 0.0, "closeMV": 0.0, "wht": 0.0, "txns": [],
                })
                h["txns"].append({
                    "date": row["Trade date"].strip()[:10],  # raw UTC date
                    "type": typ.lower(),
                    "shares": float(row["Quantity"]),
                    "price": float(row["Price"]),
                    "comm": fee,
                })
        return list(holdings.values())

    def split_shorts(self, holdings: list[dict]) -> tuple[list[dict], list[dict]]:
        """Holdings whose running balance goes negative are dropped."""
        kept, shorts = [], []
        for h in holdings:
            trades = sorted((x for x in h["txns"] if x["type"] in ("buy", "sell")),
                            key=lambda x: x["date"])
            bal, neg = h["openShares"], h["openShares"] < 0 or h["closeShares"] < 0
            for x in trades:
                bal += x["shares"] if x["type"] == "buy" else -x["shares"]
                if bal < -1e-12:
                    neg = True
                    break
            (shorts if neg and trades else kept).append(h)
        return kept, shorts

    def to_nzd(self, h: dict) -> None:
        y0, c = self.y0, h["ccy"]
        h["openMVn"] = self.convert(h["openMV"], c, f"{y0}-04-01")
        h["closeMVn"] = self.convert(h["closeMV"], c, f"{y0 + 1}-03-31")
        h["whtN"] = self.convert(h["wht"], c, f"{y0}-10-01")  # fixed 1 Oct date
        for x in h["txns"]:
            x["priceN"] = self.convert(x["price"], c, x["date"], round_=False)
            x["commN"] = self.convert(x["comm"], c, x["date"]) if x["comm"] else 0.0

    @staticmethod
    def fdr(h: dict) -> float:
        return s0(h["openMVn"] * 0.05)  # no quick-sale txns in these years

    @staticmethod
    def cv(h: dict) -> float:
        divs = sum(x["shares"] * x["priceN"] for x in h["txns"] if x["type"] == "dividend")
        sells = [x for x in h["txns"] if x["type"] == "sell"]
        buys = [x for x in h["txns"] if x["type"] == "buy"]
        proceeds = sum(x["shares"] * x["priceN"] for x in sells) - sum(x["commN"] for x in sells)
        costs = sum(x["shares"] * x["priceN"] for x in buys) + sum(x["commN"] for x in buys)
        # note: dividends column is already gross, so adding whtN double-counts
        return s0(h["closeMVn"] + divs + h["whtN"] + proceeds - h["openMVn"] - costs)


def progressive_tax(income: float, brackets) -> float:
    tax, lower = 0.0, 0.0
    for upper, r in brackets:
        if income <= lower:
            break
        tax += (min(income, upper) - lower) * r
        lower = upper
    return tax


def run(year_end: int, other_income: float) -> None:
    rates = RbnzRates(RBNZ_FILE, required_through=pd.Timestamp(year_end, 3, 31))
    eng = Engine(year_end, rates)
    kept, shorts = eng.split_shorts(eng.load())
    for h in kept:
        eng.to_nzd(h)

    fdr_per = {h["ticker"]: eng.fdr(h) for h in kept}
    cv_per = {h["ticker"]: eng.cv(h) for h in kept}
    fdr = s0(sum(fdr_per.values()))
    cv = s0(max(0.0, sum(cv_per.values())))
    elected = "CV" if cv <= fdr else "FDR"
    fif_income = min(cv, fdr)

    brackets = TAX_BRACKETS[year_end]
    total = other_income + fif_income
    nz_tax = progressive_tax(total, brackets) - progressive_tax(other_income, brackets)
    avg_rate = progressive_tax(total, brackets) / total if total > 0 else 0.0
    per = cv_per if elected == "CV" else fdr_per
    # credit capped per holding at average rate x that holding's FIF income
    credit = s0(sum(s0(min(max(0.0, per[h["ticker"]]) * avg_rate, h["whtN"])) for h in kept))
    net = s0(max(0.0, nz_tax - credit))

    print(f"\n=== fif.nz replica: year ended 31 Mar {year_end}, other income ${other_income:,.2f} ===")
    print(f"  FDR {fdr:>12,.2f}   CV {cv:>12,.2f}   elected {elected}")
    print(f"  FIF income (17B) {fif_income:,.2f} | NZ tax {s0(nz_tax):,.2f} | "
          f"credit (17A) {credit:,.2f} | net {net:,.2f}")
    if shorts:
        print(f"  dropped as 'short positions': {', '.join(h['ticker'] for h in shorts)}")
    phantoms = [h["ticker"] for h in kept
                if h["openMV"] == 0 and h["closeMV"] == 0
                and any(x["type"] in ("buy", "sell") for x in h["txns"])]
    if phantoms:
        print(f"  phantom transaction-only holdings in CV: {', '.join(phantoms)}")


if __name__ == "__main__":
    p = argparse.ArgumentParser(description="Replicate fif.nz results for reconciliation.")
    p.add_argument("--year", type=int, required=True)
    p.add_argument("--other-income", type=float, default=0.0)
    a = p.parse_args()
    run(a.year, a.other_income)
