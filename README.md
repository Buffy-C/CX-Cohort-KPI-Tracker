# CX Cohort KPI Tracker

A Python reporting pipeline for tracking new-starter cohorts through a contact-centre onboarding programme. Turns raw weekly CSV exports into a formatted Excel tracker and a set of dashboard charts — stamped with cohort and week so files stay identifiable when handed between managers.

Built to replace a manual weekly reporting burden with a **two-command routine**.

---

## What it does

- **Weekly tracker** (`build_tracker.py`) — reads the week's raw CSV exports and builds a multi-sheet Excel workbook: agent performance vs targets, team overview, and a data-sources control sheet.
- **Dashboard charts** (`dashboard.py`) — 7 charts from the tracker: KPI heatmap, performance quadrant (with agent roster boxes), team-leader comparisons and more.
- **Monthly roll-up** (`merge_weeks.py`) — builds a full month's report from the week folders you already have. No separate monthly exports needed.
- **Single source of truth** (`tracker_config.py`) — cohorts, start dates and monthly targets live in one config file. The right targets are picked automatically from the cohort's start date; weekly reporting never requires touching Python.

## Why the monthly merge is non-trivial

Naively averaging four weekly CSVs gives wrong answers, so the merge is done per file type:

- **Row-level files** (complaints, CSAT responses) are stacked and de-duplicated by Case/Contact ID
- **Rates** (resolves per hour, FCR, handle times) are recomputed from summed totals — a month's rate is total resolves ÷ total hours, not the average of four weekly rates
- **Accumulating scores** (QA, compliance) are taken from the latest week's export, since it already reflects the month

Verified two ways: a single-week month reproduces the weekly tracker to the digit, and duplicate-week input de-duplicates cleanly.

## Month logic

Month 1 is classroom training (no KPI targets). Agents are first assessed in Month 2, covering their first 4 weeks on contacts:

```
weeks 1–4  after cohort start → Month 2 targets
weeks 5–8                     → Month 3 targets
weeks 9–12                    → Month 4 targets
week 13+                      → Month 4 targets (BAU / final gate)
```

Test the mapping any time: `python scripts/tracker_config.py` prints a week → month table per cohort.

## Setup

Requires **Python 3.10+** and:

```
pip install pandas openpyxl matplotlib
```

Folder layout:

```
project/
├── scripts/
│   ├── tracker_config.py
│   ├── build_tracker.py
│   ├── dashboard.py
│   └── merge_weeks.py
├── Cohort_1/
└── Cohort_2/
```

Edit `scripts/tracker_config.py` to set your cohorts, start dates, team leaders and monthly targets. **All values shipped in this repo are illustrative examples.**

## Weekly routine

1. Create the week folder and drop the CSV exports in:

   ```
   Cohort_1/WC_2026-07-06/raw_data/
   ```

   Required: `agent_productivity`, `agent_performance`, `Number of Closed Complaints`, `Number of Resolution Rating Responses`, `QA Dashboard-Quality Scores`, `QA Dashboard-Compliance scorecard` (bracketed download suffixes like `(3)` are handled — no renaming needed).

2. Run:

   ```
   python scripts/build_tracker.py --cohort 1 --wc 2026-07-06
   python scripts/dashboard.py     --cohort 1 --wc 2026-07-06
   ```

Output: Excel tracker + 7 charts in the week folder, all stamped with cohort, W/C date and programme month in filenames and titles.

## Monthly review

```
python scripts/merge_weeks.py    --cohort 1 --month 3
python scripts/build_tracker.py  --cohort 1 --month 3
python scripts/dashboard.py      --cohort 1 --month 3
```

`merge_weeks.py` auto-finds the week folders in that month's 4-week window (warning if any are missing) and builds `Cohort_1/Month_3/` with merged raw data, tracker and charts labelled "(full month)".

## Design notes

- **Provenance stamping** — cohort + W/C date in every filename, Excel header and chart title, so a stray file in an inbox still identifies itself
- **Config-driven** — new cohort = one config entry; revised target = one number change
- **No sample data included** — the pipeline runs on your own exports; column expectations are documented in the tracker's Data Sources & Control sheet
