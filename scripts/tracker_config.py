"""
Tracker Config — SINGLE SOURCE OF TRUTH for cohorts and monthly targets.
=====================================================================
Edit targets HERE and nowhere else. build_tracker.py and dashboard.py
both import from this file, so they can never drift apart again.

Month logic
-----------
Month 1 = classroom training (no KPI targets). Agents are first assessed
in Month 2, which covers their first 4 weeks on contacts. So:

    weeks 1-4  after cohort start -> Month 2 targets
    weeks 5-8  after cohort start -> Month 3 targets
    weeks 9-12 after cohort start -> Month 4 targets
    week 13+                      -> Month 4 targets (BAU / final gate)

Sanity check: Cohort 1 started 2026-05-11, so W/C 2026-06-22 (week 7)
correctly resolves to Month 3.
"""

from datetime import date, datetime

# ── Cohorts ───────────────────────────────────────────────────────────────────
COHORTS = {
    1: {
        "name": "Cohort 1",
        "start": date(2026, 5, 11),   # W/C of first week on contacts
        "teams": ["Alex Taylor", "Jordan Lee"],
        "first_month": 2,             # first assessed month (Month 1 = training)
    },
    2: {
        "name": "Cohort 2",
        "start": date(2026, 6, 22),
        "teams": [],                  # 1 team — add TL name when known
        "first_month": 2,
    },
}

MAX_MONTH = 4  # weeks beyond Month 4 keep Month 4 targets

# ── Targets by month ──────────────────────────────────────────────────────────
# NOTE: All values below are illustrative examples — set your own organisation's
# cohorts, start dates and targets here.
# Fractions (0-1) for % metrics measured as rates; QA & Compliance are 0-100
# scores as exported by the QA dashboard. None = tracked only, no target.
TARGETS_BY_MONTH = {
    2: {
        "RPH":         1.50,
        "FCR":         0.65,
        "QA Score":    65.0,
        "QA Autofails": 2,      # max auto-fails allowed
        "Compliance":  80.0,
        "CSAT":        0.65,
        "Adherence":   0.80,    # WFM tool — reference only, not in tracker data
        "D1":          0.50,    # complaints resolved within 1 working day
        "D28":         None,    # N/A in Month 2
        "D56":         None,    # N/A until Month 4
    },
    3: {
        "RPH":         2.00,
        "FCR":         0.70,
        "QA Score":    70.0,
        "QA Autofails": 2,
        "Compliance":  85.0,
        "CSAT":        0.70,
        "Adherence":   0.85,
        "D1":          0.55,
        "D28":         0.75,
        "D56":         None,
    },
    4: {
        "RPH":         2.50,
        "FCR":         0.75,
        "QA Score":    75.0,
        "QA Autofails": 1,
        "Compliance":  90.0,
        "CSAT":        0.75,
        "Adherence":   0.85,
        "D1":          0.60,
        "D28":         0.80,
        "D56":         0.95,
    },
}

# QA pass-rate threshold (% of evaluations that must pass) — same every month
QA_PASS_RATE_THRESHOLD = 80.0


# ── Helpers ───────────────────────────────────────────────────────────────────

def parse_wc(wc) -> date:
    """Accept date or 'YYYY-MM-DD' string."""
    if isinstance(wc, date):
        return wc
    return datetime.strptime(str(wc), "%Y-%m-%d").date()


def resolve_month(cohort_id: int, wc) -> int:
    """Which programme month applies for this cohort in the given week."""
    cohort = COHORTS[int(cohort_id)]
    wc = parse_wc(wc)
    days = (wc - cohort["start"]).days
    if days < 0:
        raise ValueError(
            f"W/C {wc} is before {cohort['name']} start ({cohort['start']})")
    month = cohort["first_month"] + days // 28
    return min(month, MAX_MONTH)


def get_targets(cohort_id: int, wc) -> tuple[int, dict]:
    """Return (month, targets dict) for a cohort + week commencing date."""
    month = resolve_month(cohort_id, wc)
    return month, TARGETS_BY_MONTH[month]


def week_label(cohort_id: int, wc) -> str:
    """Provenance tag for titles/filenames, e.g. 'Cohort 1 · W/C 22 Jun 2026'.
    Every output carries this so files stay identifiable when handed around."""
    wc = parse_wc(wc)
    return f"{COHORTS[int(cohort_id)]['name']} · W/C {wc.day} {wc.strftime('%b %Y')}"


def targets_summary(month: int) -> str:
    """One-line human-readable summary, used in tracker headers/prints."""
    t = TARGETS_BY_MONTH[month]

    def pct(v):
        return f"{v:.0%}" if v is not None else "tracked (no target)"

    def score(v):
        return f"{v:.0f}%" if v is not None else "tracked (no target)"

    return (f"Month {month} Targets:  RPH ≥ {t['RPH']}  |  FCR ≥ {pct(t['FCR'])}  |  "
            f"QA ≥ {score(t['QA Score'])} ({t['QA Autofails']} autofail"
            f"{'s' if t['QA Autofails'] != 1 else ''}, pass rate {QA_PASS_RATE_THRESHOLD:.0f}%)  |  "
            f"Compliance ≥ {score(t['Compliance'])}  |  CSAT ≥ {pct(t['CSAT'])}  |  "
            f"D1 ≥ {pct(t['D1'])}  |  D28 {('≥ ' + pct(t['D28'])) if t['D28'] is not None else 'tracked'}  |  "
            f"D56 {('≥ ' + pct(t['D56'])) if t['D56'] is not None else 'tracked'}  |  "
            f"Adherence ≥ {pct(t['Adherence'])}")


if __name__ == "__main__":
    # Quick self-test: print target month for each cohort for the next 16 weeks
    from datetime import timedelta
    for cid, c in COHORTS.items():
        print(f"\n{c['name']} (started {c['start']}):")
        for w in range(0, 16, 2):
            wc = c["start"] + timedelta(weeks=w)
            m = resolve_month(cid, wc)
            print(f"  W/C {wc}  (week {w+1:>2})  ->  Month {m}")
