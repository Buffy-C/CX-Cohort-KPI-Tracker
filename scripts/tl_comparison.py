"""
Team Leader Comparison — cross-cohort, floor-week aligned
================================================================
Compares Team Leaders across cohorts at the same *relative* point in the
programme — e.g. "everyone's 2nd full week on the floor" — rather than the
same calendar week, since cohorts start on different dates. Each cohort's
own start date (from tracker_config.COHORTS) is used to resolve which W/C
folder corresponds to "week N on the floor" for that cohort.

This reads tracker files already built by build_tracker.py — run that first
for every cohort/week this script needs to compare.

Usage (run from the project root folder):
    python scripts/tl_comparison.py --floor-week 2
        -> for every cohort in tracker_config.COHORTS, resolves that cohort's
           own "week 2 on the floor" W/C date from its start date, loads
           Cohort_N/WC_<date>/KPI_Tracker_*.xlsx, and compares every
           Team Leader in every cohort on the same KPI set.

    python scripts/tl_comparison.py --floor-week 2 --cohorts 1 2
        -> restrict the comparison to specific cohorts (default: all
           cohorts defined in tracker_config.COHORTS)

    python scripts/tl_comparison.py --floor-week 2 --data ./some/root
        -> override the root folder that Cohort_N/WC_.../ sits under

Outputs (written to ./comparisons/ under the project root, unless --out is set):
    tl_comparison_floorweek{N}.png    Grouped KPI bars, one group of bars per
                                       TL, matching build_tracker.py's own
                                       "TL Detailed KPI Performance" style
    tl_trends.png                     Every KPI over calendar time, one line
                                       per TL, across every week that already
                                       has a built tracker — shows genuine
                                       improvement/decline, not floor-week
                                       aligned (unlike the comparison chart)
"""

import argparse
import re
import subprocess
import sys

# Force UTF-8 output — see build_tracker.py for why this is needed on Windows.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")
from datetime import date, datetime, timedelta
from pathlib import Path

pkgs = ["pandas", "matplotlib", "openpyxl", "numpy"]
for pkg in pkgs:
    try:
        __import__(pkg.replace("-", "_"))
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install",
                               "--break-system-packages", pkg])

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import warnings
warnings.filterwarnings("ignore")
plt.rcParams["figure.dpi"] = 130

sys.path.insert(0, str(Path(__file__).resolve().parent))
import tracker_config as cfg

# ── Brand palette (matches build_tracker.py / dashboard.py) ─────────────────
BG, INK, MUTE, GRID = "#FFFFFF", "#131409", "#81877C", "#E8E7E1"
BRAND, TEAL, RED = "#CBFB00", "#00D9AF", "#FF3F46"
TL_COLORS = ["#9B59B6", "#00D9AF", "#00BBFF", "#E67E22", "#F5C400", "#FF3F46"]

KPI_LABELS = {"RPH": "RPH", "FCR": "FCR", "QA": "QA", "Compliance": "Compliance",
              "CSAT_score": "CSAT", "D1": "D1", "D28": "D28", "D56": "D56"}
KPI_LONG = {"RPH": "Resolves per Hour", "FCR": "First Contact Resolution",
            "QA": "QA Score", "Compliance": "Compliance (PR)",
            "CSAT_score": "CSAT", "D1": "D1 Complaints (24h)",
            "D28": "D28 Complaints (28d)", "D56": "D56 Complaints (56d)"}
ORDERED = ["RPH", "FCR", "QA", "Compliance", "CSAT_score", "D1", "D28", "D56"]

RENAME = {
    "Agent":            "Agent Name",
    "FCR %":            "FCR",
    "QA Score":         "QA",
    "QA\nEvals":        "QA_Evals",
    "Compliance %":     "Compliance",
    "Comp\nEvals":      "Comp_Evals",
    "CSAT %":           "CSAT_score",
    "D1 %":             "D1",
    "D28 %":            "D28",
    "D56 %":            "D56",
    "D56 %\n(tracked)": "D56",
    "CSAT\nSurveys":    "Surveys",
    "Prod\nHours":      "Prod_Hours",
    "Closed":           "Closed_Cases",
    "Complaints\nClosed": "Complaints_Closed",
}
NUMERIC_COLS = ["RPH", "FCR", "QA", "Compliance", "CSAT_score", "D1", "D28", "D56",
                "Surveys", "QA_Evals", "Comp_Evals", "Prod_Hours", "Closed_Cases",
                "Complaints_Closed"]
RAW_0_100_COLS = ["QA", "Compliance"]   # build_tracker.py stores these on a 0-100 scale

# Same weighting logic as dashboard.py — see that script for the full
# rationale. Where a true volume denominator exists (Surveys/Evals/Prod
# Hours/Complaints Closed), this reconstructs the platform's own pooled
# team-level rate exactly rather than an unweighted mean across agents/TLs.
# FCR has no exported raw success/fail counts, so it is a best-effort
# approximation weighted by Closed Cases, not an exact reconstruction.
WEIGHT_COL = {
    "CSAT_score": "Surveys",
    "QA":         "QA_Evals",
    "Compliance": "Comp_Evals",
    "RPH":        "Prod_Hours",
    "D1":         "Complaints_Closed",
    "D28":        "Complaints_Closed",
    "D56":        "Complaints_Closed",
    "FCR":        "Closed_Cases",
}

def weighted_mean(frame, k):
    wcol = WEIGHT_COL.get(k)
    if wcol and wcol in frame.columns:
        sub = frame[[k, wcol]].dropna()
        if len(sub) and sub[wcol].sum() > 0:
            return (sub[k] * sub[wcol]).sum() / sub[wcol].sum()
        return np.nan
    series = frame[k].dropna()
    return series.mean() if len(series) else np.nan

def fmt_v(v, k):
    if pd.isna(v):
        return "—"
    return f"{v:.2f}" if k == "RPH" else f"{v:.0%}"


# ── Args ──────────────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
parser.add_argument("--floor-week", type=int, required=True,
                    help="Which week on the floor to compare (1 = each cohort's "
                         "own first week, 2 = each cohort's own second week, etc.)")
parser.add_argument("--cohorts", type=int, nargs="+", default=None,
                    help=f"Cohorts to include (default: all of {sorted(cfg.COHORTS)})")
parser.add_argument("--data", default=None,
                    help="Override the project root folder (default: parent of scripts/)")
parser.add_argument("--out", default=None, help="Override output folder")
args = parser.parse_args()

root = Path(args.data) if args.data else Path(__file__).resolve().parent.parent
cohorts = args.cohorts if args.cohorts else sorted(cfg.COHORTS)
OUT = Path(args.out) if args.out else root / "comparisons"
OUT.mkdir(parents=True, exist_ok=True)

print(f"Comparing floor-week {args.floor_week} across cohorts: {cohorts}")


# ── Resolve each cohort's own "week N on the floor" and load its tracker ────
def _parse_cohort_start(value):
    """cfg.COHORTS[...]['start'] format isn't guaranteed to be strict ISO
    (date.fromisoformat only accepts exactly YYYY-MM-DD). Handle a date/
    datetime object directly, then fall back through common string formats,
    then a general-purpose parse, before giving up with a clear message."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y",
                    "%d-%m-%Y", "%d %B %Y", "%d %b %Y", "%B %d, %Y"):
            try:
                return datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                continue
        try:
            return pd.to_datetime(value).date()
        except Exception:
            pass
    raise ValueError(
        f"Could not parse cohort start date {value!r} (type {type(value).__name__}) "
        f"from tracker_config.COHORTS. Expected a date/datetime, or a string like "
        f"'2026-05-11', '11/05/2026', or '11 May 2026'. Update COHORTS[...]['start'] "
        f"to one of these formats, or extend _parse_cohort_start() in this script.")

def resolve_week(cohort, floor_week):
    start = _parse_cohort_start(cfg.COHORTS[cohort]["start"])
    wc_date = start + timedelta(weeks=floor_week - 1)
    return wc_date.isoformat()

def remap_targets(raw_targets):
    """tracker_config stores targets under keys like 'QA Score' on a 0-100 scale;
    the Agent Tracker's short internal keys are 'QA' on a 0-1 scale (same
    remapping dashboard.py applies). Without this, QA/Compliance/CSAT
    targets silently fail to match and get dropped from the comparison."""
    return {
        "RPH":        raw_targets.get("RPH"),
        "FCR":        raw_targets.get("FCR"),
        "QA":         (raw_targets["QA Score"] / 100) if raw_targets.get("QA Score") is not None else None,
        "Compliance": (raw_targets["Compliance"] / 100) if raw_targets.get("Compliance") is not None else None,
        "CSAT_score": raw_targets.get("CSAT"),
        "D1":         raw_targets.get("D1"),
        "D28":        raw_targets.get("D28"),
        "D56":        raw_targets.get("D56"),
    }

records = []          # one row per (cohort, agent) — feeds weighted_mean per TL
meta = []              # one entry per cohort actually loaded, for the subtitle
skipped = []

for cohort in cohorts:
    if cohort not in cfg.COHORTS:
        print(f"  ! Cohort {cohort} not found in tracker_config.COHORTS — skipped")
        skipped.append(cohort)
        continue

    wc_str = resolve_week(cohort, args.floor_week)
    week_dir = root / f"Cohort_{cohort}" / f"WC_{wc_str}"
    matches = sorted(week_dir.glob("KPI_Tracker_*.xlsx"),
                     key=lambda p: p.stat().st_mtime) if week_dir.exists() else []
    if not matches:
        print(f"  ! Cohort {cohort}: no tracker found for its week {args.floor_week} "
              f"(expected W/C {wc_str} in {week_dir}) — skipped")
        skipped.append(cohort)
        continue
    FILE = matches[-1]

    month, targets = cfg.get_targets(cohort, wc_str)
    cohort_name = cfg.COHORTS[cohort]["name"]

    raw = pd.read_excel(FILE, sheet_name="Agent Tracker", header=2)
    raw = raw[raw["Agent"].notna()].copy()
    raw = raw[raw["Agent"] != "▶ Targets"].reset_index(drop=True)
    d = raw.rename(columns=RENAME)
    for col in NUMERIC_COLS:
        if col in d.columns:
            d[col] = pd.to_numeric(d[col], errors="coerce")
    for col in RAW_0_100_COLS:
        if col in d.columns:
            d[col] = d[col] / 100

    if "Team Leader" not in d.columns or not d["Team Leader"].notna().any():
        print(f"  ! Cohort {cohort}: tracker has no Team Leader column — skipped")
        skipped.append(cohort)
        continue

    d["__cohort"] = cohort
    d["__cohort_name"] = cohort_name
    d["__wc"] = wc_str
    d["__month"] = month
    d["__targets"] = [remap_targets(targets)] * len(d)
    records.append(d)
    meta.append(dict(cohort=cohort, cohort_name=cohort_name, wc=wc_str,
                      month=month, file=FILE.name, n_agents=len(d),
                      tls=sorted(d["Team Leader"].dropna().unique())))
    print(f"  ✓ Cohort {cohort} ({cohort_name}): W/C {wc_str} · Month {month} "
          f"targets · {FILE.name} · {len(d)} agents · "
          f"TLs: {', '.join(sorted(d['Team Leader'].dropna().unique()))}")

if not records:
    sys.exit("No cohorts had a tracker for that floor-week — nothing to compare.")

months_seen = sorted({m["month"] for m in meta})
if len(months_seen) > 1:
    print(f"  ! Note: cohorts resolved to different programme months "
          f"({months_seen}) for floor-week {args.floor_week} — each TL is "
          f"still scored against its own cohort's correct target, so the "
          f"%-of-target comparison remains valid, but it's worth checking "
          f"this is intentional.")

df_all = pd.concat(records, ignore_index=True)


# ── Per-TL, per-KPI weighted average + % of that TL's own target ────────────
tl_labels = []       # display label, e.g. "Lauren Van Wyk (Cohort 2)"
tl_stats = {}         # label -> {kpi: {avg, pct}}
tl_colour = {}

tl_rows = df_all[["__cohort", "__cohort_name", "Team Leader"]].drop_duplicates()
tl_rows = tl_rows.sort_values(["__cohort", "Team Leader"])

for i, (_, row) in enumerate(tl_rows.iterrows()):
    cohort, cohort_name, tl = row["__cohort"], row["__cohort_name"], row["Team Leader"]
    label = f"{tl} ({cohort_name})"
    sub = df_all[(df_all["__cohort"] == cohort) & (df_all["Team Leader"] == tl)]
    targets = sub["__targets"].iloc[0]

    tl_labels.append(label)
    tl_colour[label] = TL_COLORS[i % len(TL_COLORS)]
    tl_stats[label] = {}
    for k in ORDERED:
        if k not in sub.columns or targets.get(k) is None:
            continue
        avg = weighted_mean(sub, k)
        tl_stats[label][k] = dict(avg=avg,
                                   pct=(avg / targets[k]) if not pd.isna(avg) else np.nan)

# Only compare KPIs that have a real target for every TL being shown —
# otherwise a %-of-target bar isn't meaningful for that TL.
ACTIVE = [k for k in ORDERED
          if all(k in tl_stats[label] and not np.isnan(tl_stats[label][k]["pct"])
                 for label in tl_labels)]
if not ACTIVE:
    sys.exit("No KPI has a valid target across every included TL for this "
             "floor-week — nothing comparable to chart.")

M_KPI = len(ACTIVE)
print(f"Comparable KPIs this floor-week: {', '.join(KPI_LABELS[k] for k in ACTIVE)}")


# ═════════════════════════════════════════════════════════════════════════════
# CHART 1 — Grouped KPI comparison bars (same visual language as
# build_tracker.py's own tl_kpi_breakdown chart)
# ═════════════════════════════════════════════════════════════════════════════
ks = list(reversed(ACTIVE))
n_tl = len(tl_labels)
bar_h = 0.8 / n_tl

# Two-column layout: a narrow dedicated legend column on the left, chart on
# the right. A legend drawn *inside* the chart axes eats into plot space as
# more TL names are added; a separate column keeps the chart full-size no
# matter how long the TL list gets.
fig = plt.figure(figsize=(17, max(8, M_KPI * 1.3)))
gs = fig.add_gridspec(1, 2, width_ratios=[1, 5], wspace=0.02)
ax_leg = fig.add_subplot(gs[0, 0])
ax = fig.add_subplot(gs[0, 1])
fig.patch.set_facecolor(BG); ax.set_facecolor(BG); ax_leg.set_facecolor(BG)
ax_leg.axis("off")

for ki, k in enumerate(ks):
    for ti, label in enumerate(tl_labels):
        s = tl_stats[label].get(k)
        if s is None or np.isnan(s["pct"]):
            continue
        ypos = ki + (ti - (n_tl - 1) / 2) * (bar_h + 0.04)
        below = s["pct"] < 1.0
        ax.barh(ypos, s["pct"] * 100, height=bar_h,
                color=RED if below else tl_colour[label],
                edgecolor=BG, linewidth=0.5, zorder=3)
        # Accent strip at the bar's start, in this TL's own colour rather
        # than a fixed brand colour — otherwise once a bar turns red for
        # being below target, there's no way to tell whose bar it is.
        ax.barh(ypos, 1.2, height=bar_h, color=tl_colour[label], zorder=4,
                edgecolor="none")
        ax.text(s["pct"] * 100 - 1.5, ypos, fmt_v(s["avg"], k),
                va="center", ha="right", fontsize=9, color="white",
                fontweight="bold", zorder=5)

ax.axvline(100, color=BRAND, lw=1.8, ls="--", zorder=4, alpha=0.9)
ax.set_yticks(range(M_KPI))
ax.set_yticklabels([KPI_LABELS[k] for k in ks], fontsize=12,
                   fontweight="bold", color=INK)
ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.0f}%"))
ax.set_xlabel(f"Score as % of that TL's own Month target  (100% = target met)",
              fontsize=10.5, color=MUTE, labelpad=8)
ax.tick_params(colors=MUTE, labelsize=9)
ax.xaxis.grid(True, color=GRID, linewidth=0.6, zorder=0); ax.set_axisbelow(True)
for spine in ax.spines.values():
    spine.set_visible(False)

# Legend order matches the bars' own top-to-bottom order within each KPI
# group (highest ti = plotted highest), not the arbitrary cohort/name sort
# order used to build tl_labels.
legend_order = list(reversed(tl_labels))
handles = [mpatches.Patch(color=tl_colour[label], label=label) for label in legend_order]
handles.append(mpatches.Patch(color=RED, label="Below target"))
ax_leg.legend(handles=handles, fontsize=9.5, framealpha=0.95, loc="upper left",
              edgecolor=GRID, borderaxespad=0)

cohort_bits = ", ".join(f"Cohort {m['cohort']}: {', '.join(m['tls'])} "
                        f"(W/C {m['wc']})" for m in meta)
month_bit = f"Month {months_seen[0]} targets" if len(months_seen) == 1 \
            else f"Months {'/'.join(map(str, months_seen))} targets (mixed)"
fig.text(0.02, 0.985, f"TL Detailed KPI Performance — Floor-Week "
         f"{args.floor_week}, All Cohorts", fontsize=15, fontweight="bold",
         color=INK, ha="left", va="top")
fig.text(0.02, 0.955, cohort_bits, fontsize=9, color=MUTE, ha="left", va="top")
fig.text(0.02, 0.935, f"Each cohort's own floor-week {args.floor_week}  ·  {month_bit}",
         fontsize=9, color=MUTE, ha="left", va="top")
plt.tight_layout(rect=[0, 0, 1, 0.92])
fname1 = OUT / f"tl_comparison_floorweek{args.floor_week}.png"
plt.savefig(fname1, dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(fig)
print(f"Saved -> {fname1}")


# ═════════════════════════════════════════════════════════════════════════════
# CHART 2 — KPI trends over calendar time (improvements / declines per TL)
# ═════════════════════════════════════════════════════════════════════════════
# Unlike Chart 1, this is NOT floor-week aligned — it plots every week that
# already has a built tracker, on the real calendar, so a TL's own trajectory
# over time is visible. Cohorts starting on different dates just show up at
# different points along the same x-axis, which is the point.

def discover_built_weeks(root, cohort):
    """Every WC_<date> folder for this cohort that already has a built
    tracker, tolerant of case/spacing (WC_ / wc_ / WC ...) — same tolerant
    matching as rerun_all.py's discovery, so this doesn't silently miss
    folders named slightly differently than expected."""
    cohort_dir = root / f"Cohort_{cohort}"
    if not cohort_dir.exists():
        return []
    found = []
    for sub in cohort_dir.iterdir():
        if not sub.is_dir():
            continue
        m = re.match(r"wc[\s_]*(.+)$", sub.name, re.IGNORECASE)
        if not m:
            continue
        matches = sorted(sub.glob("KPI_Tracker_*.xlsx"), key=lambda p: p.stat().st_mtime)
        if not matches:
            continue
        wc_str = m.group(1).strip()
        try:
            wc_date = _parse_cohort_start(wc_str)
        except ValueError:
            continue
        found.append(dict(wc=wc_str, date=wc_date, file=matches[-1]))
    found.sort(key=lambda w: w["date"])
    return found

trend_rows = []
trend_meta = []
for cohort in cohorts:
    if cohort not in cfg.COHORTS:
        continue
    cohort_name = cfg.COHORTS[cohort]["name"]
    weeks = discover_built_weeks(root, cohort)
    if not weeks:
        continue
    trend_meta.append(f"Cohort {cohort}: {len(weeks)} week(s)")
    for w in weeks:
        try:
            raw = pd.read_excel(w["file"], sheet_name="Agent Tracker", header=2)
        except Exception as e:
            print(f"  ! Cohort {cohort} W/C {w['wc']}: couldn't read tracker for "
                  f"trends ({e}) — skipped")
            continue
        raw = raw[raw["Agent"].notna()].copy()
        raw = raw[raw["Agent"] != "▶ Targets"].reset_index(drop=True)
        d = raw.rename(columns=RENAME)
        for col in NUMERIC_COLS:
            if col in d.columns:
                d[col] = pd.to_numeric(d[col], errors="coerce")
        for col in RAW_0_100_COLS:
            if col in d.columns:
                d[col] = d[col] / 100
        if "Team Leader" not in d.columns or not d["Team Leader"].notna().any():
            continue

        _, raw_targets = cfg.get_targets(cohort, w["wc"])
        wk_targets = remap_targets(raw_targets)

        for tl in sorted(d["Team Leader"].dropna().unique()):
            sub = d[d["Team Leader"] == tl]
            label = f"{tl} ({cohort_name})"
            for k in ORDERED:
                if k not in sub.columns:
                    continue
                avg = weighted_mean(sub, k)
                if pd.isna(avg):
                    continue
                trend_rows.append(dict(date=w["date"], wc=w["wc"], cohort=cohort,
                                       label=label, kpi=k, value=avg,
                                       target=wk_targets.get(k)))

if not trend_rows:
    print("\n  (No trend data available — need at least one built tracker "
          "per cohort. Skipping tl_trends.png.)")
else:
    trend_df = pd.DataFrame(trend_rows)

    # Consistent TL colouring across both charts where the same TL appears
    # in both; new colours assigned for any TL only present in trend data.
    trend_labels = sorted(trend_df["label"].unique(),
                          key=lambda l: (l not in tl_colour, l))
    for i, label in enumerate(trend_labels):
        if label not in tl_colour:
            used = set(tl_colour.values())
            avail = [c for c in TL_COLORS if c not in used] or TL_COLORS
            tl_colour[label] = avail[i % len(avail)]

    trend_kpis = [k for k in ORDERED if k in trend_df["kpi"].unique()]
    n_panels = len(trend_kpis)
    n_cols = 2
    n_rows = (n_panels + n_cols - 1) // n_cols

    fig, axes = plt.subplots(n_rows, n_cols, figsize=(15, 4.2 * n_rows), squeeze=False)
    fig.patch.set_facecolor(BG)

    for i, k in enumerate(trend_kpis):
        r, c = divmod(i, n_cols)
        ax = axes[r][c]
        ax.set_facecolor(BG)
        kdf = trend_df[trend_df["kpi"] == k]

        for label in trend_labels:
            ldf = kdf[kdf["label"] == label].sort_values("date")
            if ldf.empty:
                continue
            ax.plot(ldf["date"], ldf["value"], marker="o", markersize=5,
                    lw=2, color=tl_colour[label], label=label, zorder=3,
                    markeredgecolor=BG, markeredgewidth=0.8)

        latest_target = kdf.sort_values("date")["target"].dropna()
        if len(latest_target):
            t = latest_target.iloc[-1]
            ax.axhline(t, color=BRAND, lw=1.5, ls="--", zorder=2, alpha=0.85)
            ax.text(0.01, t, f" Target {fmt_v(t, k)}", transform=ax.get_yaxis_transform(),
                    fontsize=7.5, color=INK, fontstyle="italic", va="bottom")

        if k != "RPH":
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.0%}"))
        ax.set_title(KPI_LONG[k], fontsize=11, fontweight="bold", color=INK, pad=8)
        ax.tick_params(axis="x", colors=MUTE, labelsize=8, rotation=30)
        ax.tick_params(axis="y", colors=MUTE, labelsize=8)
        ax.yaxis.grid(True, color=GRID, linewidth=0.6, zorder=0); ax.set_axisbelow(True)
        for spine in ax.spines.values():
            spine.set_visible(False)

    for j in range(n_panels, n_rows * n_cols):
        r, c = divmod(j, n_cols)
        axes[r][c].axis("off")

    handles = [mpatches.Patch(color=tl_colour[label], label=label) for label in trend_labels]
    fig.legend(handles=handles, loc="upper left", bbox_to_anchor=(0.01, 0.99),
              fontsize=9.5, framealpha=0.95, edgecolor=GRID, ncol=min(len(trend_labels), 3))

    fig.text(0.5, 1.01, "KPI Trends Over Time — All Cohorts, All Team Leaders",
             ha="center", fontsize=15, fontweight="bold", color=INK)
    fig.text(0.5, 0.985, "Every week with a built tracker  ·  "  + "  ·  ".join(trend_meta) +
             "  ·  dashed line = most recent target for that KPI",
             ha="center", fontsize=9, color=MUTE)
    plt.tight_layout(rect=[0, 0, 1, 0.96])
    fname2 = OUT / "tl_trends.png"
    plt.savefig(fname2, dpi=150, bbox_inches="tight", facecolor=BG)
    plt.close(fig)
    print(f"Saved -> {fname2}")

print(f"\nAll outputs saved to: {OUT.resolve()}")
if skipped:
    print(f"Cohorts skipped: {skipped}")
