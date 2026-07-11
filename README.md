# CX Cohort KPI Tracker

A Python reporting pipeline for tracking new-starter cohorts through a contact-centre onboarding programme. Turns raw weekly CSV exports into a formatted Excel tracker and a set of dashboard charts — stamped with cohort and week so files stay identifiable when handed between managers.

Built to replace a manual weekly reporting burden with a **two-command routine**.

---

## What it does

- **Weekly tracker** (`build_tracker.py`) — reads the week's raw CSV exports and builds a multi-sheet Excel workbook: agent performance vs targets, team overview, and a data-sources control sheet.
- **Dashboard charts** (`dashboard.py`) — 7 charts from the tracker: KPI heatmap, performance quadrant (with agent roster boxes), team-leader comparisons and more.
- **Monthly roll-up** (`merge_weeks.py`) — builds a full month's report from the week folders you already have. No separate monthly exports needed.
- **Cross-cohort comparison** (`tl_comparison.py`) — compares every Team Leader across every cohort at the same *relative* point in the programme (e.g. everyone's 2nd week on the floor), even though cohorts start on different calendar dates. Also plots every KPI's trend over real calendar time, so improvement or decline is visible at a glance.
- **Batch re-run** (`rerun_all.py`) — re-runs the tracker + dashboard across every existing week/month folder in one command. Useful after a calculation fix, so historical weeks get corrected too instead of just the next new one.
- **Single source of truth** (`tracker_config.py`) — cohorts, start dates and monthly targets live in one config file. The right targets are picked automatically from the cohort's start date; weekly reporting never requires touching Python.

## Why the monthly merge is non-trivial

Naively averaging four weekly CSVs gives wrong answers, so the merge is done per file type:

- **Row-level files** (complaints, CSAT responses) are stacked and de-duplicated by Case/Contact ID
- **Rates** (resolves per hour, FCR, handle times) are recomputed from summed totals — a month's rate is total resolves ÷ total hours, not the average of four weekly rates
- **Accumulating scores** (QA, compliance) are taken from the latest week's export, since it already reflects the month

Verified two ways: a single-week month reproduces the weekly tracker to the digit, and duplicate-week input de-duplicates cleanly.

The same "don't average an average" principle applies inside `build_tracker.py` and `dashboard.py` too — cohort-level KPI averages are pooled by the underlying volume (surveys, evaluations, hours, complaints closed) rather than a plain per-agent mean, so a low-volume agent doesn't get the same pull on the team average as a high-volume one.

## Month logic

Month 1 is classroom training (no KPI targets). Agents are first assessed in Month 2, covering their first 4 weeks on contacts:

```
weeks 1–4 after cohort start → Month 2 targets
weeks 5–8 → Month 3 targets
weeks 9–12 → Month 4 targets
week 13+ → Month 4 targets (BAU / final gate)
```

Test the mapping any time: `python scripts/tracker_config.py` prints a week → month table per cohort.

## Setup

Requires **Python 3.10+** and:

```
pip install pandas openpyxl matplotlib numpy
```

Folder layout:

```
project/
├── scripts/
│   ├── tracker_config.py
│   ├── build_tracker.py
│   ├── dashboard.py
│   ├── merge_weeks.py
│   ├── tl_comparison.py
│   └── rerun_all.py
├── Cohort_1/
└── Cohort_2/
```

Edit `scripts/tracker_config.py` to set your cohorts, start dates, team leaders and monthly targets. **All values shipped in this repo are illustrative examples.**

## Weekly routine

1. Create the week folder and drop the CSV exports in:

```
Cohort_1/WC_2026-07-06/raw_data/
```

Required: `agent_productivity`, `agent_performance`, `Number of Closed Complaints`, `Blended_Customer_Satisfaction__`, `QA Dashboard-Quality Scores`, `QA Dashboard-Compliance scorecard` (bracketed download suffixes like `(3)` are handled — no renaming needed).

2. Run:

```
python scripts/build_tracker.py --cohort 1 --wc 2026-07-06
python scripts/dashboard.py --cohort 1 --wc 2026-07-06
```

Output: Excel tracker + 7 charts in the week folder, all stamped with cohort, W/C date and programme month in filenames and titles.

## Monthly review

```
python scripts/merge_weeks.py --cohort 1 --month 3
python scripts/build_tracker.py --cohort 1 --month 3
python scripts/dashboard.py --cohort 1 --month 3
```

`merge_weeks.py` auto-finds the week folders in that month's 4-week window (warning if any are missing) and builds `Cohort_1/Month_3/` with merged raw data, tracker and charts labelled "(full month)".

## Cross-cohort comparison

Compare every Team Leader across every cohort at the same relative point in the programme — each cohort's own start date is used to resolve which calendar week corresponds to "week N on the floor" for that cohort, so a fair comparison doesn't depend on cohorts starting on the same date:

```
python scripts/tl_comparison.py --floor-week 2
```

Requires a tracker already built for that floor-week in each cohort. Produces `tl_comparison_floorweek{N}.png` (grouped KPI bars, one group per TL) and `tl_trends.png` (every KPI plotted over real calendar time, one line per TL — discovers every week that already has a built tracker automatically, no date range needed). Optional flags: `--cohorts 1 2` to restrict which cohorts are included, `--data` to point at a different root folder.

## Batch re-run

Re-run the tracker and dashboard across every existing week/month folder in one command — useful after a calculation fix, to backfill every past week rather than only new ones going forward:

```
python scripts/rerun_all.py
```

Discovers every `Cohort_N/WC_.../raw_data` and `Cohort_N/Month_.../raw_data` folder automatically. Prints a summary table at the end so a partial run is easy to spot. Useful flags: `--dry-run` (see what would run without running it), `--cohorts 1 2` (restrict to specific cohorts), `--skip-dashboard` (just refresh the Excel trackers), `--stop-on-error` (abort on first failure instead of logging it and continuing).

## Design notes

- **Provenance stamping** — cohort + W/C date in every filename, Excel header and chart title, so a stray file in an inbox still identifies itself
- **Config-driven** — new cohort = one config entry; revised target = one number change
- **No sample data included** — the pipeline runs on your own exports; column expectations are documented in the tracker's Data Sources & Control sheet
