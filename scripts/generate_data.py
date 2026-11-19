"""Generate the synthetic dataset for the Budget vs Actual (project cost control) dashboard.

Everything here is fictional: Ridgeline Construction Group, a made-up general contractor
with 10 projects and 4 project managers. Data runs from Jul 2024 to 30 Jun 2026 (the as-of date).
The random seed is fixed, so every run produces exactly the same files.

Output (default ../data):
    projects.csv, project_managers.csv, cost_codes.csv      lookups
    budget_phased.csv                                        estimating system: monthly budget by cost code
    change_orders.csv                                        owner change orders (approved / pending / rejected)
    commitments.csv                                          purchase orders and subcontracts
    erp_job_cost_export.csv                                  ERP job-cost transactions (deliberately messy)
    billing.csv                                              monthly pay applications to the client
    pm_updates/PM_Update_<name>.xlsx                         each PM's weekly % complete and monthly forecast
    pm_updates/PM_Update_TEMPLATE.xlsx                       blank template for a new PM

Usage:
    python generate_data.py [output_dir]
Needs openpyxl for the Excel files (pip install openpyxl).
"""
import csv
import math
import random
import sys
from datetime import date, timedelta
from pathlib import Path

random.seed(20260630)

OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "data"
OUT.mkdir(parents=True, exist_ok=True)

AS_OF = date(2026, 6, 30)


# ----------------------------------------------------------------------------- date helpers
def add_months(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    return date(d.year + y, m + 1, 1)


def month_end(d):
    return add_months(d, 1) - timedelta(days=1)


def months_between(a, b):
    return (b.year - a.year) * 12 + b.month - a.month


def rand_day(month_start, lo=1, hi=None):
    last = month_end(month_start).day
    return date(month_start.year, month_start.month, random.randint(lo, min(hi or last, last)))


def write_csv(name, header, rows):
    with open(OUT / name, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows([[f"{v:.2f}" if isinstance(v, float) else v for v in r] for r in rows])
    print(f"{name:<28}{len(rows):>8,} rows")


# ----------------------------------------------------------------------------- lookups
PMS = [
    ("PM1", "Sarah Mitchell", "sarah.mitchell@ridgeline.example"),
    ("PM2", "David Chen", "david.chen@ridgeline.example"),
    ("PM3", "Maria Lopez", "maria.lopez@ridgeline.example"),
    ("PM4", "James Okafor", "james.okafor@ridgeline.example"),
]

# code, name, category, window start, window end (share of duration), procurement
COST_CODES = [
    ("01-100", "General Conditions", "General Conditions", 0.00, 1.00, "self"),
    ("01-500", "Equipment Rental", "Equipment", 0.00, 0.85, "po"),
    ("02-200", "Sitework & Earthwork", "Subcontract", 0.00, 0.30, "po"),
    ("03-300", "Concrete", "Subcontract", 0.08, 0.50, "po"),
    ("05-100", "Structural Steel", "Subcontract", 0.20, 0.55, "po"),
    ("06-100", "Rough Carpentry (self-perform)", "Labor", 0.30, 0.80, "self"),
    ("07-500", "Roofing & Waterproofing", "Subcontract", 0.45, 0.75, "po"),
    ("08-100", "Doors, Frames & Glazing", "Materials", 0.50, 0.85, "po"),
    ("09-900", "Interior Finishes", "Subcontract", 0.60, 1.00, "po"),
    ("22-000", "Plumbing", "Subcontract", 0.25, 0.95, "po"),
    ("23-000", "HVAC", "Subcontract", 0.30, 0.95, "po"),
    ("26-000", "Electrical", "Subcontract", 0.25, 1.00, "po"),
    ("32-100", "Paving & Traffic Control", "Subcontract", 0.60, 1.00, "po"),
    ("33-000", "Process Piping & Utilities", "Subcontract", 0.20, 0.90, "po"),
]
CC = {c[0]: c for c in COST_CODES}
UNALLOCATED = ("99-999", "Unallocated (needs recoding)", "Unallocated")

VENDORS = {
    "01-500": "Midstate Equipment Rentals", "02-200": "Terra Earthworks", "03-300": "BlueLine Concrete",
    "05-100": "Apex Steel Erectors", "07-500": "Northern Roofing Co.", "08-100": "Keystone Doors & Glass",
    "09-900": "Precision Interiors", "22-000": "Clearwater Plumbing", "23-000": "Summit Mechanical",
    "26-000": "Brightway Electric", "32-100": "Granite Paving & Traffic", "33-000": "Riverbend Piping",
}

MIX = {
    "building": {"01-100": 9, "01-500": 3, "02-200": 6, "03-300": 12, "05-100": 9, "06-100": 6, "07-500": 4,
                 "08-100": 5, "09-900": 14, "22-000": 8, "23-000": 11, "26-000": 11},
    "residential": {"01-100": 9, "01-500": 3, "02-200": 6, "03-300": 11, "05-100": 4, "06-100": 14, "07-500": 5,
                    "08-100": 6, "09-900": 14, "22-000": 9, "23-000": 9, "26-000": 10},
    "industrial": {"01-100": 8, "01-500": 4, "02-200": 7, "03-300": 13, "05-100": 16, "06-100": 1, "07-500": 5,
                   "08-100": 3, "09-900": 4, "22-000": 5, "23-000": 16, "26-000": 18},
    "datacenter": {"01-100": 8, "01-500": 3, "02-200": 2, "03-300": 6, "05-100": 6, "06-100": 2, "08-100": 2,
                   "09-900": 5, "22-000": 4, "23-000": 28, "26-000": 34},
    "water": {"01-100": 10, "01-500": 8, "02-200": 18, "03-300": 22, "05-100": 6, "26-000": 12, "33-000": 24},
    "bridge": {"01-100": 11, "01-500": 15, "02-200": 8, "03-300": 26, "05-100": 22, "26-000": 3, "32-100": 15},
}

# Each project: who runs it, how big it is, and its story.
#   spi      schedule performance (0.90 = progress 10% slower than plan)
#   base     typical cost factor (1.02 = costs run 2% over the earned value)
#   special  {code: (factor, first_month, last_month)}: month indexes from project start, inclusive
#   recog    how much of the overrun the PM puts in the forecast (1.0 = all of it)
#   bill     billing vs earned revenue (1.05 = bills 5% ahead of progress)
PROJECTS = [
    dict(id="P01", name="Riverside Medical Office Building", client="Riverside Health Partners", type="Commercial",
         mix="building", pm="PM1", start=date(2024, 7, 1), dur=14, contract=8_400_000, margin=0.11,
         spi=0.97, base=1.00, special={}, recog=0.9, bill=1.02),
    dict(id="P02", name="Northgate Distribution Center", client="Northgate Logistics LLC", type="Industrial",
         mix="industrial", pm="PM2", start=date(2024, 10, 1), dur=24, contract=16_500_000, margin=0.10,
         spi=0.95, base=1.01, special={"05-100": (1.25, 0, 99), "03-300": (1.06, 0, 99)}, recog=0.30, bill=1.06),
    dict(id="P03", name="Cedar Hills School Expansion", client="Cedar Hills School District", type="Institutional",
         mix="building", pm="PM3", start=date(2025, 1, 1), dur=16, contract=6_200_000, margin=0.12,
         spi=0.98, base=0.99, special={}, recog=0.9, bill=1.00),
    dict(id="P04", name="Harbor Street Mixed-Use", client="Harbor Street Development Co.", type="Commercial",
         mix="building", pm="PM1", start=date(2025, 3, 1), dur=22, contract=14_800_000, margin=0.10,
         spi=0.90, base=1.01, special={"09-900": (1.07, 0, 99)}, recog=0.6, sched_recog=0.35, bill=1.01),
    dict(id="P05", name="Westfield Water Treatment Upgrade", client="City of Westfield", type="Infrastructure",
         mix="water", pm="PM4", start=date(2024, 9, 1), dur=20, contract=11_200_000, margin=0.13,
         spi=0.96, base=1.00, special={"02-200": (1.42, 1, 6)}, recog=0.9, bill=1.00),
    dict(id="P06", name="Lakeview Apartments Phase 2", client="Lakeview Residential Partners", type="Residential",
         mix="residential", pm="PM2", start=date(2025, 6, 1), dur=18, contract=9_600_000, margin=0.09,
         spi=0.96, base=1.01, special={"06-100": (1.30, 6, 8)}, recog=0.40, bill=1.03),
    dict(id="P07", name="Summit Data Center Fit-out", client="Summit Cloud Services", type="Industrial",
         mix="datacenter", pm="PM3", start=date(2025, 8, 1), dur=12, contract=12_400_000, margin=0.11,
         spi=0.97, base=1.00, special={"23-000": (1.19, 5, 99), "26-000": (1.17, 5, 99)}, recog=0.25, bill=0.93),
    dict(id="P08", name="Route 9 Bridge Rehabilitation", client="State DOT District 4", type="Infrastructure",
         mix="bridge", pm="PM4", start=date(2025, 4, 1), dur=16, contract=7_900_000, margin=0.12,
         spi=1.02, base=0.94, special={}, recog=0.9, bill=1.00),
    dict(id="P09", name="Oakmont Retail Plaza", client="Oakmont Commercial REIT", type="Commercial",
         mix="building", pm="PM1", start=date(2025, 11, 1), dur=14, contract=5_600_000, margin=0.11,
         spi=1.00, base=1.00, special={}, recog=0.9, bill=1.00),
    dict(id="P10", name="Pinecrest Senior Living", client="Pinecrest Living Communities", type="Residential",
         mix="residential", pm="PM2", start=date(2026, 3, 1), dur=20, contract=13_100_000, margin=0.10,
         spi=0.99, base=1.01, special={"02-200": (1.08, 0, 99)}, recog=0.6, bill=1.02),
]

# Owner change orders with a story. Everything else is generated at random below.
#   recovery=True: the CO pays for work already done (the cost is in actuals already)
STORY_COS = {
    "P05": [dict(code="02-200", desc="Unforeseen rock excavation (differing site condition)", submit=3, approve=7,
                 cost=380_000, markup=0.10, status="Approved", recovery=True)],
    "P07": [dict(code="23-000", desc="Added cooling capacity for hall B (client request)", submit=7, approve=None,
                 cost=420_000, markup=0.12, status="Pending", recovery=False)],
    "P02": [dict(code="05-100", desc="Steel price escalation claim", submit=9, approve=None,
                 cost=260_000, markup=0.0, status="Rejected", recovery=False)],
}
# Commitment changes (subcontract change orders) issued when a problem was found
STORY_CCOS = {
    "P05": [("02-200", 2, 0.40, "Change: rock excavation and disposal")],
    "P07": [("23-000", 5, 0.17, "Change: added scope and overtime"), ("26-000", 5, 0.15, "Change: added scope and overtime")],
}
STORY_BUYOUT = {("P02", "05-100"): 1.21}   # steel bought out 21% over budget (price spike)


# ----------------------------------------------------------------------------- simulation helpers
def phase_weights(n, power=1.2):
    w = [math.sin(math.pi * (i + 0.5) / n) ** power for i in range(n)]
    s = sum(w)
    return [x / s for x in w]


def spread(amount, months, weights):
    """Split `amount` over months by weights, rounded to cents, exact total."""
    parts = [round(amount * w, 2) for w in weights]
    parts[-1] = round(amount - sum(parts[:-1]), 2)
    return list(zip(months, parts))


def cum_at(series, x):
    """Planned cumulative after x months (x may be fractional). series = list of monthly amounts."""
    if x <= 0:
        return 0.0
    k = int(math.floor(x))
    total = sum(series[:min(k, len(series))])
    if k < len(series):
        total += (x - k) * series[k]
    return total


def code_window(code, dur):
    _, _, _, a, b, _ = CC[code]
    i0 = int(math.floor(a * dur))
    i1 = max(i0, int(math.ceil(b * dur)) - 1)
    return i0, min(i1, dur - 1)


budget_rows, co_rows, commit_rows, txn_rows, billing_rows = [], [], [], [], []
progress_rows, forecast_rows, project_rows = [], [], []
txn_seq = [0]


def next_txn():
    txn_seq[0] += 1
    return f"JC{txn_seq[0]:06d}"


def add_txn(d, pid, code, source, vendor, po, desc, amount):
    if abs(amount) < 0.005:
        return
    txn_rows.append([next_txn(), d.isoformat(), pid, code, source, vendor, po, desc, round(amount, 2)])


for p in PROJECTS:
    pid, dur, start, spi = p["id"], p["dur"], p["start"], p["spi"]
    orig_budget = round(p["contract"] * (1 - p["margin"]), -2)
    mix = MIX[p["mix"]]
    wsum = sum(mix.values())
    codes = list(mix)
    code_budget = {c: round(orig_budget * mix[c] / wsum, -2) for c in codes}
    code_budget[codes[-1]] += orig_budget - sum(code_budget.values())
    code_factor = {c: random.uniform(0.97, 1.04) * p["base"] for c in codes}
    k_asof = months_between(start, AS_OF)          # month index of the as-of month

    # ---- streams: (code, monthly planned list over a horizon, ac_on, label)
    H = int(dur / min(spi, 1.0)) + 8                # simulation horizon in months
    streams = []
    for c in codes:
        series = [0.0] * H
        if c == "01-100":                           # general conditions: flat over the planned duration
            months, weights = list(range(dur)), [1 / dur] * dur
        else:
            i0, i1 = code_window(c, dur)
            months = list(range(i0, i1 + 1))
            weights = phase_weights(len(months), 0.6 if c == "01-500" else 1.2)
        for i, amt in spread(code_budget[c], months, weights):
            series[i] += amt
            budget_rows.append([pid, c, add_months(start, i).isoformat(), "Original", "", amt])
        streams.append(dict(code=c, series=series, ac_on=True, co=None))

    # ---- owner change orders
    cos = [dict(x) for x in STORY_COS.get(pid, [])]
    n_random = random.randint(1, 3)
    elapsed = max(1, min(k_asof, dur) - 4)
    for _ in range(n_random):
        c = random.choice([x for x in codes if x != "01-100"])
        sub = random.randint(1, elapsed)
        lag = random.randint(1, 3)
        if sub + lag <= k_asof:
            status = "Approved" if random.random() < 0.88 else "Rejected"
        else:
            status = "Pending"
        cos.append(dict(code=c, desc=random.choice([
            "Owner-requested scope change", "Upgraded materials (owner selection)", "Added work per RFI response",
            "Revised drawings (design development)", "Owner allowance adjustment", "Added scope per site instruction"]),
            submit=sub, approve=sub + lag if status == "Approved" else None,
            cost=round(code_budget[c] * random.uniform(0.02, 0.08), -2), markup=random.uniform(0.10, 0.15),
            status=status, recovery=False))
    cos = [x for x in cos if x["submit"] <= k_asof]
    cos.sort(key=lambda x: x["submit"])
    for n, co in enumerate(cos, 1):
        co["id"] = f"CO-{pid}-{n:02d}"
        sub_d = rand_day(add_months(start, co["submit"]))
        app_d = rand_day(add_months(start, co["approve"])) if co["approve"] is not None else None
        price = round(co["cost"] * (1 + co["markup"]), -2)
        co_rows.append([co["id"], pid, co["code"], co["desc"], sub_d.isoformat(), app_d.isoformat() if app_d else "",
                        co["status"], float(co["cost"]), float(price)])
        co["approved_date"], co["price"] = app_d, price
        if co["status"] == "Approved":
            a = co["approve"]
            i0, i1 = code_window(co["code"], dur)
            months = list(range(a, max(a, i1) + 1))[:4] if a <= i1 else [a]
            series = [0.0] * H
            for i, amt in spread(co["cost"], months, phase_weights(len(months))):
                series[i] += amt
                budget_rows.append([pid, co["code"], add_months(start, i).isoformat(), "Change order", co["id"], amt])
            streams.append(dict(code=co["code"], series=series, ac_on=not co["recovery"], co=co))

    def bac_at(k, code=None):
        """Revised budget known at the end of month k (original + COs approved by then)."""
        tot = 0.0
        for s in streams:
            if code and s["code"] != code:
                continue
            if s["co"] is None or s["co"]["approve"] <= k:
                tot += sum(s["series"])
        return tot

    def cf(code, k):
        f = code_factor[code]
        if code in p["special"]:
            fac, a, b = p["special"][code]
            if a <= k <= b:
                f = fac
        return f

    # ---- monthly simulation: earned value and actual cost per stream
    total_budget = sum(sum(s["series"]) for s in streams)
    ev_cum = {i: 0.0 for i in range(len(streams))}
    ac_cum_code = {c: 0.0 for c in codes}
    ev_cum_code = {c: 0.0 for c in codes}
    gc_rate = code_budget["01-100"] / dur
    k_done = None
    month_actuals = []                               # (k, code, amount)
    pct_month_end = {}
    for k in range(H):
        if k > k_asof:
            break
        x = (k + 1) * spi
        for i, s in enumerate(streams):
            if s["co"] is not None and s["co"]["approve"] > k:
                continue
            target = cum_at(s["series"], x)
            ev_m = max(0.0, target - ev_cum[i])
            ev_cum[i] = target
            ev_cum_code[s["code"]] += ev_m
            if s["code"] == "01-100" or not s["ac_on"] or ev_m <= 0:
                continue
            amt = ev_m * cf(s["code"], k) * random.gauss(1.0, 0.03)
            month_actuals.append((k, s["code"], amt))
            ac_cum_code[s["code"]] += amt
        earned_total = sum(ev_cum.values())
        known_bac = bac_at(k)
        finished = earned_total >= known_bac - 1 and all(
            (s["co"] is None or s["co"]["approve"] <= k) for s in streams) and x >= dur
        # general conditions are time-based: they run every month until the job is finished
        gc = gc_rate * cf("01-100", k) * random.gauss(1.0, 0.02)
        month_actuals.append((k, "01-100", gc))
        ac_cum_code["01-100"] += gc
        pct_month_end[k] = min(1.0, earned_total / known_bac)
        # ---- PM forecast snapshot at month end: forecast to complete by cost code
        snap = month_end(add_months(start, k))
        recog = p["recog"]
        for c in codes:
            if finished:
                ftc = 0.0
            elif c == "01-100":
                believed_dur = dur + (dur / spi - dur) * p.get("sched_recog", 0.8)
                ftc = max(0.0, believed_dur - (k + 1)) * gc_rate * code_factor["01-100"]
            else:
                remaining = max(0.0, bac_at(k, c) - ev_cum_code[c])
                ratio = ac_cum_code[c] / ev_cum_code[c] if ev_cum_code[c] > 0 else 1.0
                ftc = remaining * (1 + recog * (ratio - 1)) if ratio > 1 else remaining
            forecast_rows.append((p["pm"], pid, c, snap, round(ftc, -2)))
        if finished:
            k_done = k
            break

    # ---- actual cost transactions (one month at a time)
    po_of = {}
    for k, c, amt in month_actuals:
        m0 = add_months(start, k)
        if c == "01-100":
            payroll = amt * 0.62
            for d in (14, 28):
                add_txn(date(m0.year, m0.month, d), pid, c, "Payroll", "Payroll", "", "Site staff payroll", payroll / 2)
            add_txn(rand_day(m0, 3, 25), pid, c, "AP Invoice", "Site Services Co.", "",
                    "Site office, utilities, temporary facilities", amt - payroll)
        elif c == "06-100":
            for d in (14, 28):
                add_txn(date(m0.year, m0.month, d), pid, c, "Payroll", "Payroll", "", "Carpentry crew payroll", amt / 2)
        else:
            n = 1 if amt < 40_000 else random.randint(2, 3)
            parts = [random.uniform(0.6, 1.4) for _ in range(n)]
            ps = sum(parts)
            po_id = f"PO-{pid}-{c.replace('-', '')}"
            po_of[c] = po_id
            for part in parts:
                add_txn(rand_day(m0, 2, 27), pid, c, "AP Invoice", VENDORS[c], po_id,
                        f"{VENDORS[c]} progress invoice", amt * part / ps)

    # ---- commitments: subcontracts / POs issued a month before the work starts
    for c in codes:
        if CC[c][5] != "po":
            continue
        i0, _ = code_window(c, dur)
        k_po = max(0, int(i0 / spi) - 1)
        if k_po > k_asof:
            continue
        d = rand_day(add_months(start, k_po), 3, 25)
        buyout = STORY_BUYOUT.get((pid, c), random.uniform(0.94, 1.00) if p["base"] >= 0.97 else random.uniform(0.90, 0.95))
        commit_rows.append([f"PO-{pid}-{c.replace('-', '')}", 1, pid, c, VENDORS[c], d.isoformat(),
                            f"{CC[c][1]}: original subcontract / PO", round(code_budget[c] * buyout, -2)])
        line = 2
        for code2, k_cco, share, desc in STORY_CCOS.get(pid, []):
            if code2 == c and k_cco <= k_asof:
                commit_rows.append([f"PO-{pid}-{c.replace('-', '')}", line, pid, c, VENDORS[c],
                                    rand_day(add_months(start, k_cco), 3, 25).isoformat(), desc,
                                    round(code_budget[c] * share, -2)])
                line += 1
        for co in cos:
            if co["code"] == c and co["status"] == "Approved" and not co["recovery"]:
                commit_rows.append([f"PO-{pid}-{c.replace('-', '')}", line, pid, c, VENDORS[c],
                                    co["approved_date"].isoformat(), f"Change: {co['desc']} ({co['id']})",
                                    round(co["cost"] * random.uniform(0.92, 0.98), -2)])
                line += 1

    # ---- weekly % complete reported by the PM (Fridays), interpolated inside each month
    first_friday = start + timedelta(days=(4 - start.weekday()) % 7)
    last_day = month_end(add_months(start, k_done)) if k_done is not None else AS_OF
    wk, last_pct = first_friday, 0.0
    while wk <= min(last_day, AS_OF):
        k = months_between(start, wk)
        prev = pct_month_end.get(k - 1, 0.0)
        cur = pct_month_end.get(k, prev)
        frac = wk.day / month_end(wk).day
        pct = prev + (cur - prev) * frac + random.gauss(0, 0.004)
        pct = 1.0 if (k_done is not None and wk >= month_end(add_months(start, k_done)) - timedelta(days=6)) else pct
        pct = round(min(1.0, max(last_pct, pct)), 3)
        last_pct = pct
        progress_rows.append((p["pm"], pid, wk, pct))
        wk += timedelta(days=7)

    # ---- billing: monthly pay applications, 10% retention
    def contract_at(k):
        return p["contract"] + sum(co["price"] for co in cos if co["status"] == "Approved" and co["approve"] <= k)

    billed, retention, app = 0.0, 0.0, 0
    for k in sorted(pct_month_end):
        target = contract_at(k) if k == k_done else min(contract_at(k), pct_month_end[k] * contract_at(k) * p["bill"])
        amt = round(max(0.0, target - billed), 2)
        if amt > 0:
            app += 1
            ret = round(amt * 0.10, 2)
            billed += amt
            retention += ret
            billing_rows.append([pid, month_end(add_months(start, k)).isoformat(), app, amt, ret])
    if k_done is not None and k_done + 1 <= k_asof:
        app += 1
        billing_rows.append([pid, month_end(add_months(start, k_done + 1)).isoformat(), app, 0.0, -round(retention, 2)])

    status = "Completed" if k_done is not None else ("Starting" if pct_month_end[k_asof] < 0.15 else "In progress")
    project_rows.append([pid, p["name"], p["client"], p["type"], p["pm"], start.isoformat(),
                         (add_months(start, dur) - timedelta(days=1)).isoformat(), p["contract"], status])
    print(f"{pid} {p['name']:<36} BAC {orig_budget:>12,.0f}  AC {sum(a for _, _, a in month_actuals):>12,.0f}"
          f"  done {status:<11} pct {pct_month_end[max(pct_month_end)]:.1%}")

# ----------------------------------------------------------------------------- messy ERP export
# 1. some invoices were posted to the wrong cost code, then reversed and reposted
clean = list(txn_rows)
txn_rows.clear()
for r in clean:
    if r[4] == "AP Invoice" and random.random() < 0.02:
        wrong = random.choice([c for c in CC if c != r[3]])
        txn_rows.append([next_txn(), r[1], r[2], wrong, r[4], r[5], r[6], r[7], r[8]])
        rev_d = min(AS_OF, date.fromisoformat(r[1]) + timedelta(days=random.randint(3, 20)))
        txn_rows.append([next_txn(), rev_d.isoformat(), r[2], wrong, "Journal", r[5], r[6],
                         "Reversal: posted to wrong cost code", -r[8]])
        txn_rows.append([next_txn(), rev_d.isoformat(), r[2], r[3], "Journal", r[5], r[6],
                         "Repost to correct cost code", r[8]])
    else:
        txn_rows.append(r)
# 2. small costs coded to an unallocated code that nobody fixed yet
for pid in ("P02", "P04", "P06", "P07", "P09"):
    for _ in range(random.randint(2, 4)):
        m0 = add_months(date(2026, 1, 1), random.randint(0, 5))
        txn_rows.append([next_txn(), rand_day(m0).isoformat(), pid, "99-999", "AP Invoice",
                         random.choice(["City Permits Office", "QuickFix Supplies", "Metro Testing Lab"]), "",
                         random.choice(["Permit fee", "Miscellaneous supplies", "Material testing"]),
                         round(random.uniform(800, 9_000), 2)])
txn_rows.sort(key=lambda r: (r[1], r[0]))
# 3. formatting noise from the export: cost codes without the dash or with spaces, lower-case project ids
for r in txn_rows:
    u = random.random()
    if u < 0.05:
        r[3] = r[3].replace("-", "")
    elif u < 0.08:
        r[3] = " " + r[3] + " "
    if random.random() < 0.02:
        r[2] = r[2].lower()
# 4. one export batch was loaded twice: exact duplicate lines
dups = random.sample(range(len(txn_rows)), k=int(len(txn_rows) * 0.008))
for i in sorted(dups, reverse=True):
    txn_rows.insert(i + 1, list(txn_rows[i]))

# ----------------------------------------------------------------------------- write CSV files
write_csv("project_managers.csv", ["pm_id", "pm_name", "pm_email"], PMS)
write_csv("cost_codes.csv", ["cost_code", "cost_code_name", "cost_category", "sort_order"],
          [[c[0], c[1], c[2], i + 1] for i, c in enumerate(COST_CODES)] + [[*UNALLOCATED, 99]])
write_csv("projects.csv", ["project_id", "project_name", "client", "project_type", "pm_id", "start_date",
                           "planned_finish", "original_contract_value", "status"], project_rows)
write_csv("budget_phased.csv", ["project_id", "cost_code", "month_start", "budget_type", "co_id", "amount"],
          budget_rows)
write_csv("change_orders.csv", ["co_id", "project_id", "cost_code", "description", "submitted_date",
                                "approved_date", "status", "cost_amount", "price_amount"], co_rows)
write_csv("commitments.csv", ["po_id", "po_line", "project_id", "cost_code", "vendor", "commitment_date",
                              "description", "amount"], commit_rows)
write_csv("erp_job_cost_export.csv", ["txn_id", "posting_date", "project_id", "cost_code", "source", "vendor",
                                      "po_id", "description", "amount"], txn_rows)
write_csv("billing.csv", ["project_id", "billing_period_end", "application_no", "billed_amount",
                          "retention_held"], billing_rows)

# ----------------------------------------------------------------------------- PM Excel workbooks
try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError:
    print("openpyxl not installed: PM workbooks skipped (pip install openpyxl)")
    sys.exit(0)

PM_DIR = OUT / "pm_updates"
PM_DIR.mkdir(exist_ok=True)
HEAD = Font(bold=True, color="FFFFFF")
FILL = PatternFill("solid", fgColor="1E3A5F")


def table(ws, name, header, rows, widths, formats):
    ws.append(header)
    for r in rows:
        ws.append(list(r))
    n = max(2, len(rows) + 1)
    ref = f"A1:{chr(64 + len(header))}{n}"
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(t)
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(65 + i)].width = w
    for col, fmt in formats.items():
        for cell in ws[col][1:]:
            cell.number_format = fmt
    ws.freeze_panes = "A2"
    return n


def build_workbook(path, pm, projects, progress, forecast):
    wb = Workbook()
    ws = wb.active
    ws.title = "Read me"
    lines = [
        ("PM monthly update: Ridgeline Construction Group (fictional)", Font(bold=True, size=14, color="1E3A5F")),
        (f"Project manager: {pm[1]}" if pm else "Project manager: <your name>", Font(bold=True)),
        ("", None),
        ("1. Progress sheet: every Friday, add one row per project with the physical % complete.", None),
        ("   Type the % as a percentage (45%) or a decimal (0.45). One row per project per week.", None),
        ("2. Forecast sheet: at each month end, enter the cost still to come (forecast to complete)", None),
        ("   for every cost code of every project. Enter 0 when a cost code is finished.", None),
        ("3. Save the file in the shared pm_updates folder. Power BI picks it up at the next refresh.", None),
        ("", None),
        ("Do not rename the sheets, the tables or the column headers.", Font(italic=True, color="B91C1C")),
    ]
    for text, font in lines:
        ws.append([text])
        if font:
            ws.cell(ws.max_row, 1).font = font
    ws.column_dimensions["A"].width = 100

    wl = wb.create_sheet("Lists")
    wl.append(["project_id", "project_name"])
    for r in projects:
        wl.append([r[0], r[1]])
    wl["D1"] = "cost_code"
    wl["E1"] = "cost_code_name"
    for i, c in enumerate(COST_CODES, 2):
        wl.cell(i, 4, c[0])
        wl.cell(i, 5, c[1])
    for col, w in zip("ABDE", (12, 38, 12, 34)):
        wl.column_dimensions[col].width = w
    for cell in (wl["A1"], wl["B1"], wl["D1"], wl["E1"]):
        cell.font, cell.fill = HEAD, FILL
    n_proj = max(2, len(projects) + 1)

    wp = wb.create_sheet("Progress", 1)
    n = table(wp, "tblProgress", ["project_id", "week_ending", "pct_complete", "comment"], progress,
              [12, 14, 14, 60], {"B": "yyyy-mm-dd", "C": "0.0%"})
    dv = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${n_proj}", allow_blank=True)
    dv.error, dv.errorTitle = "Pick one of your projects from the list.", "Unknown project"
    wp.add_data_validation(dv)
    dv.add(f"A2:A{n + 200}")
    dv2 = DataValidation(type="decimal", operator="between", formula1="0", formula2="1", allow_blank=True,
                         errorStyle="warning")
    dv2.error, dv2.errorTitle = "Enter a % between 0% and 100%.", "Check the % complete"
    wp.add_data_validation(dv2)
    dv2.add(f"C2:C{n + 200}")

    wf = wb.create_sheet("Forecast", 2)
    n = table(wf, "tblForecast", ["project_id", "cost_code", "forecast_month", "forecast_to_complete", "note"],
              forecast, [12, 12, 16, 22, 50], {"C": "yyyy-mm-dd", "D": "#,##0"})
    dv3 = DataValidation(type="list", formula1=f"=Lists!$D$2:$D${len(COST_CODES) + 1}", allow_blank=True)
    wf.add_data_validation(dv3)
    dv3.add(f"B2:B{n + 500}")
    dv4 = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${n_proj}", allow_blank=True)
    wf.add_data_validation(dv4)
    dv4.add(f"A2:A{n + 500}")
    for sheet in (wp, wf):
        for cell in sheet[1]:
            cell.alignment = Alignment(horizontal="left")
    wb.save(path)


COMMENTS = {
    "P02": ["Steel erection on site", "Waiting on steel deliveries", "Roof deck started", "Slab on grade poured"],
    "P04": ["Permit for level 4 still pending", "Crew short this week", "Facade mock-up approved"],
    "P05": ["Rock found in the clarifier excavation", "Blasting crew on site", "Process piping testing"],
    "P06": ["Snow day, framing stopped", "Framing crew overtime", "Drywall started in building A"],
    "P07": ["Owner asked for extra cooling in hall B", "Generators delivered", "Commissioning planned"],
}
for pm in PMS:
    projs = [r for r in project_rows if r[4] == pm[0]]
    prog, last_week = [], {}
    for pm_id, pid, wk, pct in progress_rows:
        if pm_id != pm[0]:
            continue
        val, pid_txt = pct, pid
        if pm[0] == "PM1" and random.random() < 0.15:      # Sarah sometimes types 45 instead of 45%
            val = round(pct * 100, 1)
        if pm[0] == "PM2" and random.random() < 0.06:      # David's ids sometimes typed by hand
            pid_txt = random.choice([pid.lower(), pid + " "])
        comment = random.choice(COMMENTS[pid]) if pid in COMMENTS and random.random() < 0.15 else None
        prog.append((pid_txt, wk, val, comment))
        if pm[0] == "PM2" and random.random() < 0.03:      # entered twice: the second row is the correction
            prog.append((pid, wk, round(min(1.0, pct + 0.004), 3), "Corrected"))
        if random.random() < 0.01:
            prog.append((None, None, None, None))           # blank row left in the table
    fc = [(pid, c, snap, ftc, None) for pm_id, pid, c, snap, ftc in forecast_rows if pm_id == pm[0]]
    path = PM_DIR / f"PM_Update_{pm[1].replace(' ', '_')}.xlsx"
    build_workbook(path, pm, projs, prog, fc)
    print(f"{path.name:<36}{len(prog):>6} progress rows {len(fc):>6} forecast rows")
build_workbook(PM_DIR / "PM_Update_TEMPLATE.xlsx", None, [], [], [])
print("PM_Update_TEMPLATE.xlsx")
