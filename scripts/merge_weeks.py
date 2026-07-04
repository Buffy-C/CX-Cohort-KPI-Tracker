"""
Week Merger — build a full month's raw_data from existing week folders
===========================================================================
No re-exporting needed: this combines the weekly CSVs you already have into
Cohort_N/Month_N/raw_data/, ready for build_tracker.py / dashboard.py --month.

Usage (run from the project root folder):
    python scripts/merge_weeks.py --cohort 1 --month 3
        -> finds every Cohort_1/WC_*/ folder whose date falls in Month 3
        -> merges their raw_data into Cohort_1/Month_3/raw_data/

    python scripts/merge_weeks.py --cohort 1 --weeks 2026-06-08 2026-06-15
        -> merge exactly these week folders instead of auto-selecting

How each file is merged (this matters — averaging averages is wrong):
    Closed Complaints / CSAT responses  row-level: stacked, de-duplicated by
                                        Case ID / Contact ID (latest week wins)
    agent_productivity                  counts summed; RPH/CPH recomputed from
                                        totals; FCR/CSAT weighted by contacts;
                                        Open Cases = latest snapshot
    agent_performance                   counts summed; handle times weighted
                                        by volume
    QA / Compliance scorecards          copied from the LATEST week — these
                                        exports are accumulating scores, so
                                        the newest file already covers the month
    team-level single-value files       copied from the latest week
"""

import argparse
import re
import shutil
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import tracker_config as cfg


# ── Helpers ───────────────────────────────────────────────────────────────────

def pct_to_float(v):
    if pd.isna(v):
        return np.nan
    try:
        return float(str(v).replace("%", "").strip())
    except Exception:
        return np.nan

def time_to_sec(t):
    if pd.isna(t) or str(t).strip() == "":
        return np.nan
    try:
        h, m, s = (int(x) for x in str(t).strip().split(":"))
        return h * 3600 + m * 60 + s
    except Exception:
        return np.nan

def sec_to_time(sec):
    if pd.isna(sec):
        return ""
    sec = int(round(sec))
    return f"{sec // 3600:02d}:{(sec % 3600) // 60:02d}:{sec % 60:02d}"

def wavg(values, weights):
    v = pd.to_numeric(values, errors="coerce")
    w = pd.to_numeric(weights, errors="coerce").fillna(0)
    mask = v.notna() & (w > 0)
    if not mask.any():
        return v.dropna().mean() if v.notna().any() else np.nan
    return np.average(v[mask], weights=w[mask])

def find_file(folder: Path, name: str):
    """Find a file tolerating spaces/underscores and bracketed numbers."""
    if (folder / name).exists():
        return folder / name
    stem = Path(name).stem.lower().replace("_", " ").replace("-", " ")
    stem_clean = re.sub(r"[\s_]*[\(\[]?\d+[\)\]]?[\s_]*", " ", stem).strip()
    for f in folder.iterdir():
        if f.suffix.lower() != Path(name).suffix.lower():
            continue
        cand = re.sub(r"[\s_]*[\(\[]?\d+[\)\]]?[\s_]*", " ",
                      f.stem.lower().replace("_", " ").replace("-", " ")).strip()
        if cand == stem_clean:
            return f
    return None

def read_weeks(week_dirs, fname):
    """Read one CSV from each week folder; returns list of (wc_date, df)."""
    out = []
    for wc, d in week_dirs:
        f = find_file(d / "raw_data", fname)
        if f is None:
            print(f"  ! {fname} missing in WC_{wc} — skipped for that week")
            continue
        df = pd.read_csv(f, index_col=0)
        df.columns = df.columns.str.strip()
        out.append((wc, df))
    return out


# ── Args & week selection ─────────────────────────────────────────────────────

parser = argparse.ArgumentParser()
parser.add_argument("--cohort", type=int, required=True)
parser.add_argument("--month", type=int, default=None,
                    help="Auto-select week folders belonging to this programme month")
parser.add_argument("--weeks", nargs="+", default=None,
                    help="Explicit W/C dates to merge, e.g. 2026-06-08 2026-06-15")
args = parser.parse_args()

if bool(args.month) == bool(args.weeks):
    sys.exit("Pass exactly one of --month (auto-select) or --weeks (explicit list).")

root = Path(__file__).resolve().parent.parent
cohort_dir = root / f"Cohort_{args.cohort}"
cohort = cfg.COHORTS[args.cohort]

# All available week folders
available = {}
for d in sorted(cohort_dir.glob("WC_*")):
    try:
        available[datetime.strptime(d.name[3:], "%Y-%m-%d").date()] = d
    except ValueError:
        continue

if args.month:
    MONTH = args.month
    win_start = cohort["start"] + timedelta(days=(MONTH - cohort["first_month"]) * 28)
    win_end = win_start + timedelta(days=28)
    week_dirs = [(wc, d) for wc, d in available.items() if win_start <= wc < win_end]
    print(f"Month {MONTH} window: {win_start} to {win_end - timedelta(days=1)}")
else:
    wanted = [datetime.strptime(w, "%Y-%m-%d").date() for w in args.weeks]
    missing = [w for w in wanted if w not in available]
    if missing:
        sys.exit(f"Week folder(s) not found: {missing}\nAvailable: {sorted(available)}")
    week_dirs = [(w, available[w]) for w in sorted(wanted)]
    MONTH = cfg.resolve_month(args.cohort, max(wanted))
    print(f"Explicit weeks -> labelled Month {MONTH} (from latest week)")

if not week_dirs:
    sys.exit(f"No WC_* folders found in that range.\nAvailable: {sorted(available)}")
week_dirs.sort(key=lambda x: x[0])
latest_dir = week_dirs[-1][1] / "raw_data"
print(f"Merging {len(week_dirs)} week(s): " + ", ".join(str(w) for w, _ in week_dirs))
if len(week_dirs) < 4:
    print(f"  (note: a full month is usually 4 weeks — check nothing is missing)")

out_dir = cohort_dir / f"Month_{MONTH}" / "raw_data"
out_dir.mkdir(parents=True, exist_ok=True)


# ── 1. Row-level files: stack + de-duplicate ─────────────────────────────────

for fname, key in [("Number of Closed Complaints.csv", "Case ID"),
                   ("Number of Resolution Rating Responses.csv", "Contact ID")]:
    frames = read_weeks(week_dirs, fname)
    if not frames:
        print(f"  ! {fname}: no data in any week")
        continue
    merged = pd.concat([df for _, df in frames], ignore_index=True)
    before = len(merged)
    merged = merged.drop_duplicates(subset=key, keep="last").reset_index(drop=True)
    merged.to_csv(out_dir / fname)
    print(f"  ✓ {fname}: {before} rows -> {len(merged)} unique {key}s")


# ── 2. agent_productivity: totals + weighted rates ────────────────────────────

frames = read_weeks(week_dirs, "agent_productivity.csv")
if frames:
    allp = pd.concat([df.assign(_wc=str(wc)) for wc, df in frames], ignore_index=True)
    for c in ["Productive Hours", "Number of Contacts Handled",
              "Number of Contacts Disconnected", "Number of Closed Cases",
              "Number of Open Cases", "Resolves per Hour", "Contacts per Hour"]:
        allp[c] = pd.to_numeric(allp[c], errors="coerce")
    allp["_fcr"]  = allp["First Contact Resolution %"].apply(pct_to_float)
    allp["_csat"] = allp["CSAT Results - Blended Customer Satisfaction %"].apply(pct_to_float)
    allp["_resolves"] = allp["Resolves per Hour"] * allp["Productive Hours"]

    rows = []
    for agent, g in allp.groupby("Agent Name"):
        g = g.sort_values("_wc")
        hours    = g["Productive Hours"].sum()
        contacts = g["Number of Contacts Handled"].sum()
        rows.append({
            "Agent Name": agent,
            "Team Leader": g["Team Leader"].dropna().iloc[-1] if g["Team Leader"].notna().any() else np.nan,
            "Productive Hours": round(hours, 2),
            "Number of Contacts Handled": contacts,
            "Number of Contacts Disconnected": g["Number of Contacts Disconnected"].sum(),
            "Contacts per Hour": round(contacts / hours, 2) if hours else np.nan,
            "CSAT Results - Blended Customer Satisfaction %":
                (f"{wavg(g['_csat'], g['Number of Contacts Handled']):.1f}%"
                 if g["_csat"].notna().any() else ""),
            "Number of Closed Cases": g["Number of Closed Cases"].sum(),
            "Resolves per Hour": round(g["_resolves"].sum() / hours, 2) if hours else np.nan,
            "First Contact Resolution %":
                (f"{wavg(g['_fcr'], g['Number of Contacts Handled']):.1f}%"
                 if g["_fcr"].notna().any() else ""),
            "Number of Open Cases": g["Number of Open Cases"].dropna().iloc[-1]
                if g["Number of Open Cases"].notna().any() else np.nan,   # snapshot
            "QA Score": g["QA Score"].dropna().iloc[-1] if g["QA Score"].notna().any() else np.nan,
        })
    pd.DataFrame(rows).to_csv(out_dir / "agent_productivity.csv")
    print(f"  ✓ agent_productivity.csv: {len(rows)} agents "
          f"(counts summed, RPH/CPH from totals, FCR/CSAT contact-weighted)")


# ── 3. agent_performance: totals + volume-weighted handle times ───────────────

frames = read_weeks(week_dirs, "agent_performance.csv")
if frames:
    allf = pd.concat([df.assign(_wc=str(wc)) for wc, df in frames], ignore_index=True)
    allf = allf[allf["Agent Name"].notna()].copy()
    COUNT_COLS = [c for c in allf.columns if c.startswith("Number of")]
    TIME_COLS = ["Average Speed of Answer  - Live Channels",
                 "Average Handle Time - All Channels",
                 "Average ACW Time - All Channels",
                 "Average Hold Time - All Channels"]
    for c in COUNT_COLS:
        allf[c] = pd.to_numeric(allf[c], errors="coerce")
    allf["_vol"] = allf[[c for c in COUNT_COLS]].sum(axis=1)

    rows = []
    for agent, g in allf.groupby("Agent Name"):
        g = g.sort_values("_wc")
        row = {"Agent Name": agent,
               "Team Leader": g["Team Leader"].dropna().iloc[-1]
                   if g["Team Leader"].notna().any() else np.nan}
        for c in COUNT_COLS:
            row[c] = g[c].sum()
        for c in TIME_COLS:
            if c in g.columns:
                row[c] = sec_to_time(wavg(g[c].apply(time_to_sec), g["_vol"]))
        rows.append(row)
    pd.DataFrame(rows).to_csv(out_dir / "agent_performance.csv")
    print(f"  ✓ agent_performance.csv: {len(rows)} agents "
          f"(counts summed, handle times volume-weighted)")


# ── 4. QA / Compliance: latest week's file (accumulating scores) ─────────────

for fname in ["QA Dashboard-Quality Scores.csv",
              "QA Dashboard-Compliance scorecard.csv"]:
    src = find_file(latest_dir, fname)
    if src:
        shutil.copy(src, out_dir / fname)
        print(f"  ✓ {fname}: copied from latest week "
              f"(accumulating export — already covers the month)")
    else:
        print(f"  ! {fname}: not found in latest week")


# ── 5. Fallback + team-level files: latest week's copy ────────────────────────

for fname in ["complaints_closed_by_agent.csv",
              "blended_customer_satisfaction__.csv",
              "agent_rating_customer_satisfaction__.csv",
              "average_days_to_resolve_complaint.csv",
              "average_complaint_age_of_open_complaints__calendar_days_.csv",
              "complaint_closure_volumes.csv"]:
    src = find_file(latest_dir, fname)
    if src:
        shutil.copy(src, out_dir / fname)
        print(f"  ✓ {fname}: copied from latest week")

print(f"\nDone -> {out_dir.resolve()}")
print(f"Next:  python scripts/build_tracker.py --cohort {args.cohort} --month {MONTH}")
print(f"       python scripts/dashboard.py     --cohort {args.cohort} --month {MONTH}")
