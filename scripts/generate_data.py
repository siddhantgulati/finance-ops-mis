"""
Generates the sample data set used in the MIS workbook.

Real billing data can't be shared publicly, so this script builds a realistic
stand-in for one financial year (Apr 2025 - Mar 2026) of a professional
services firm: a client master, invoices and the payments received against them.

Things I wanted the data to have, because they show up in real AR data:
  - a few large clients that make up a big share of billing
  - most clients pay close to terms, but some are habitually late
  - partial payments, plus some invoices still open at year end
  - a handful of data issues (a duplicate payment, a payment with no invoice)
    so the reconciliation checks in the workbook have something to catch

Usage:  python scripts/generate_data.py
Output: data/clients.csv, data/invoices.csv, data/payments.csv
"""

from pathlib import Path

import numpy as np
import pandas as pd

SEED = 2026
FY_START = pd.Timestamp("2025-04-01")
FY_END = pd.Timestamp("2026-03-31")
N_CLIENTS = 160
GST_RATE = 0.18

OUT_DIR = Path(__file__).resolve().parents[1] / "data"
rng = np.random.default_rng(SEED)

PREFIXES = ["Arora", "Bansal", "Chawla", "Dhillon", "Garg", "Kapoor", "Khanna", "Malhotra",
            "Mehra", "Nanda", "Oberoi", "Puri", "Sethi", "Sood", "Tandon", "Verma", "Agarwal",
            "Bhatia", "Iyer", "Joshi", "Kulkarni", "Menon", "Nair", "Rao", "Reddy", "Shetty",
            "Sinha", "Thakur", "Desai", "Mistry", "Patel", "Shah", "Banerjee", "Ghosh", "Mukherjee",
            "Saxena", "Trivedi", "Pillai", "Chopra", "Grover"]
SUFFIXES = {
    "Manufacturing": ["Industries", "Auto Components", "Steels", "Polymers", "Engineering Works"],
    "Retail & FMCG": ["Retail", "Foods", "Consumer Products", "Mart", "Beverages"],
    "BFSI": ["Finserv", "Capital", "Securities", "Housing Finance", "Insurance Brokers"],
    "Technology": ["Infotech", "Software", "Digital", "Systems", "Technologies"],
    "Healthcare": ["Pharma", "Healthcare", "Life Sciences", "Diagnostics", "Hospitals"],
    "Energy & Infra": ["Infra Projects", "Power", "Renewables", "Logistics", "Constructions"],
}
LEGAL = ["Pvt Ltd", "Ltd", "LLP", "Pvt Ltd", "Ltd"]
REGIONS = ["North", "West", "South", "East"]
SERVICE_LINES = {  # share of invoices, typical fee in INR
    "Audit & Assurance": (0.30, 420_000),
    "Tax": (0.25, 180_000),
    "Consulting": (0.20, 650_000),
    "Risk Advisory": (0.15, 350_000),
    "Financial Advisory": (0.10, 500_000),
}


def build_clients() -> pd.DataFrame:
    names = set()
    rows = []
    industries = list(SUFFIXES)
    for i in range(N_CLIENTS):
        industry = industries[rng.integers(len(industries))]
        while True:
            name = f"{rng.choice(PREFIXES)} {rng.choice(SUFFIXES[industry])} {rng.choice(LEGAL)}"
            if name not in names:
                names.add(name)
                break
        rows.append({
            "ClientID": f"C{1001 + i}",
            "ClientName": name,
            "Industry": industry,
            "Region": rng.choice(REGIONS, p=[0.38, 0.30, 0.22, 0.10]),
            "PaymentTermsDays": int(rng.choice([30, 45, 60], p=[0.5, 0.3, 0.2])),
        })
    df = pd.DataFrame(rows)

    # Client size follows a heavy tail: a few big accounts, many small ones.
    df["SizeWeight"] = rng.pareto(1.3, N_CLIENTS) + 0.2

    # Payment behaviour. Most pay within a couple of weeks of the due date;
    # about 12% are habitual late payers, and these skew towards bigger accounts
    # (large clients often push suppliers on terms).
    rank = df["SizeWeight"].rank(ascending=False)
    p_late = np.where(rank <= 20, 0.45, 0.07)
    df["Behaviour"] = np.where(rng.random(N_CLIENTS) < p_late, "Late", "Normal")
    df.loc[rng.choice(N_CLIENTS, 6, replace=False), "Behaviour"] = "Early"
    df["CreditRating"] = np.select(
        [df["Behaviour"] == "Late", df["Behaviour"] == "Early"],
        [rng.choice(["B", "BB"], N_CLIENTS), rng.choice(["AA", "AAA"], N_CLIENTS)],
        rng.choice(["A", "BBB", "AA"], N_CLIENTS, p=[0.45, 0.35, 0.20]),
    )
    return df


def build_invoices(clients: pd.DataFrame) -> pd.DataFrame:
    p = clients["SizeWeight"] / clients["SizeWeight"].sum()
    n = 11_400
    client_idx = rng.choice(len(clients), n, p=p)

    # Billing is heavier at quarter ends, and March is the busiest month.
    days = pd.date_range(FY_START, FY_END, freq="D")
    w = np.ones(len(days))
    w[days.month.isin([6, 9, 12, 3])] *= 1.35
    w[days.month == 3] *= 1.25
    w[days.dayofweek >= 5] *= 0.15  # few invoices raised on weekends
    inv_dates = rng.choice(days, n, p=w / w.sum())

    lines = list(SERVICE_LINES)
    line_p = [SERVICE_LINES[s][0] for s in lines]
    service = rng.choice(lines, n, p=line_p)
    typical = np.array([SERVICE_LINES[s][1] for s in service])
    amount = np.round(typical * rng.lognormal(0, 0.55, n) / 100) * 100
    amount = np.clip(amount, 15_000, 4_500_000)

    df = pd.DataFrame({
        "InvoiceDate": pd.to_datetime(inv_dates),
        "ClientID": clients["ClientID"].values[client_idx],
        "ServiceLine": service,
        "NetAmount": amount,
    }).sort_values("InvoiceDate", kind="stable").reset_index(drop=True)

    df.insert(0, "InvoiceID", [f"INV-{d:%y%m}-{i:05d}" for i, d in enumerate(df["InvoiceDate"], 1)])
    terms = df["ClientID"].map(clients.set_index("ClientID")["PaymentTermsDays"])
    df["DueDate"] = df["InvoiceDate"] + pd.to_timedelta(terms, unit="D")
    df["GST"] = np.round(df["NetAmount"] * GST_RATE, 2)
    df["GrossAmount"] = df["NetAmount"] + df["GST"]
    return df[["InvoiceID", "InvoiceDate", "DueDate", "ClientID", "ServiceLine",
               "NetAmount", "GST", "GrossAmount"]]


def build_payments(invoices: pd.DataFrame, clients: pd.DataFrame) -> pd.DataFrame:
    behaviour = invoices["ClientID"].map(clients.set_index("ClientID")["Behaviour"]).values
    n = len(invoices)

    # Days paid after the due date (negative = before due date).
    delay = np.where(
        behaviour == "Late", rng.gamma(2.2, 38, n) + 10,
        np.where(behaviour == "Early", rng.normal(-12, 5, n), rng.normal(4, 11, n)),
    ).round()
    # A small share of invoices from every client get stuck (disputes, missing POs).
    stuck = rng.random(n) < np.where(behaviour == "Late", 0.10, 0.015)
    delay[stuck] = 400

    pay_date = invoices["DueDate"] + pd.to_timedelta(delay, unit="D")
    pay_date = pay_date.where(pay_date > invoices["InvoiceDate"],
                              invoices["InvoiceDate"] + pd.Timedelta(days=3))

    rows = []
    pid = 1
    for inv, inv_date, gross, d, b in zip(invoices["InvoiceID"], invoices["InvoiceDate"],
                                          invoices["GrossAmount"], pay_date, behaviour):
        # Roughly 9% of invoices are paid in two instalments.
        if rng.random() < (0.18 if b == "Late" else 0.07):
            first = round(gross * rng.uniform(0.4, 0.7), 2)
            first_date = max(d - pd.Timedelta(days=int(rng.integers(15, 45))), inv_date + pd.Timedelta(days=3))
            parts = [(first_date, first),
                     (d + pd.Timedelta(days=int(rng.integers(10, 50))), round(gross - first, 2))]
        else:
            parts = [(d, gross)]
        for when, amt in parts:
            if when <= FY_END:  # anything after 31 Mar is still open at year end
                rows.append((f"PAY-{pid:06d}", inv, when, amt,
                             rng.choice(["NEFT", "RTGS", "Cheque"], p=[0.62, 0.30, 0.08])))
                pid += 1

    pay = pd.DataFrame(rows, columns=["PaymentID", "InvoiceID", "PaymentDate", "AmountReceived", "Mode"])

    # Deliberate data issues for the reconciliation checks to find.
    dup = pay.sample(2, random_state=SEED).copy()
    dup["PaymentID"] = [f"PAY-{pid:06d}", f"PAY-{pid + 1:06d}"]
    orphan = pd.DataFrame([[f"PAY-{pid + 2:06d}", "INV-2602-99999", pd.Timestamp("2026-02-17"), 84_960.0, "NEFT"]],
                          columns=pay.columns)
    pay = pd.concat([pay, dup, orphan], ignore_index=True).sort_values("PaymentDate", kind="stable")
    return pay.reset_index(drop=True)


def main():
    OUT_DIR.mkdir(exist_ok=True)
    clients = build_clients()
    invoices = build_invoices(clients)
    payments = build_payments(invoices, clients)

    clients.drop(columns=["SizeWeight", "Behaviour"]).to_csv(OUT_DIR / "clients.csv", index=False)
    invoices.to_csv(OUT_DIR / "invoices.csv", index=False, date_format="%Y-%m-%d")
    payments.to_csv(OUT_DIR / "payments.csv", index=False, date_format="%Y-%m-%d")

    print(f"clients:  {len(clients):>6,}")
    print(f"invoices: {len(invoices):>6,}   gross billed  Rs {invoices['GrossAmount'].sum() / 1e7:,.1f} Cr")
    print(f"payments: {len(payments):>6,}   collected     Rs {payments['AmountReceived'].sum() / 1e7:,.1f} Cr")


if __name__ == "__main__":
    main()
