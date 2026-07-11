"""
KPI Dashboard — cohort & month aware
=====================================
Generates all review charts as PNGs from the Excel tracker built by
build_tracker.py. Targets come from tracker_config.py — never edit them here.

Usage (run from the project root folder):
    python scripts/dashboard.py --cohort 1 --wc 2026-06-29
        -> reads  Cohort_1/WC_2026-06-29/KPI_Tracker_WC_*.xlsx
        -> writes Cohort_1/WC_2026-06-29/charts/*.png

Charts produced:
    1. agent_quadrant.png      Speed vs quality scatter (RPH x FCR)
    2. kpi_bars.png            All KPIs vs targets, per agent
    3. kpi_heatmap.png         Full KPI table with RAG colours
    4. kpi_pareto.png          KPI fails per agent + cumulative line
    5. kpi_scorecard.png       Cohort-level KPI tiles
    6. tl_kpi_breakdown.png    Per-TL detailed KPI performance
    7. tl_team_average.png     Per-TL average vs targets
"""

import argparse
import subprocess
import sys

# Force UTF-8 output — see build_tracker.py for why this is needed on Windows.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")
from pathlib import Path

pkgs = ["pandas", "matplotlib", "openpyxl"]
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
import textwrap
import warnings
warnings.filterwarnings("ignore")
plt.rcParams["figure.dpi"] = 130

sys.path.insert(0, str(Path(__file__).resolve().parent))
import tracker_config as cfg

# ── Brand palette ─────────────────────────────────────────────────────────────
BG, INK, MUTE, GRID = "#FFFFFF", "#131409", "#81877C", "#E8E7E1"
BRAND, TEAL, AMBER  = "#CBFB00", "#00D9AF", "#9B9500"
RED, GREY, NAVY, BLUE = "#FF3F46", "#81877C", "#242C28", "#00BBFF"

QUAD_COLORS = {"High Performance": TEAL, "Experience Focus": BLUE,
               "Speed Risk": AMBER, "Underperforming": RED}
QUAD_BG = {"High Performance": "#E0FBF5", "Experience Focus": "#E0F8FF",
           "Speed Risk": "#F5FCB0", "Underperforming": "#FFE5E6"}
QUAD_LBL = {"High Performance": dict(fg=TEAL, border=TEAL),
            "Experience Focus": dict(fg=BLUE, border=BLUE),
            "Speed Risk":       dict(fg="#4A4600", border=AMBER),
            "Underperforming":  dict(fg=RED, border=RED)}

TL_COLORS = [BLUE, TEAL, "#9B59B6", "#E67E22"]  # per team leader


# ── Args & data ───────────────────────────────────────────────────────────────

parser = argparse.ArgumentParser()
parser.add_argument("--cohort", type=int, required=True,
                    help=f"Cohort number: {sorted(cfg.COHORTS)}")
parser.add_argument("--wc", default=None, help="Weekly mode: W/C date")
parser.add_argument("--month", type=int, default=None, help="Full-month mode")
parser.add_argument("--data", default=None,
                    help="Override week folder containing the tracker xlsx")
args = parser.parse_args()

if bool(args.wc) == bool(args.month):
    sys.exit("Pass exactly one of --wc or --month.")
if args.month:
    MONTH = args.month
    T = cfg.TARGETS_BY_MONTH[MONTH]
    SUBTITLE_TAG = (f"{cfg.COHORTS[args.cohort]['name']} · Month {MONTH} "
                    f"(full month) · Programme Month {MONTH}")
    FSUF = f"_C{args.cohort}_Month_{MONTH}"
    period_sub = f"Month_{MONTH}"
else:
    MONTH, T = cfg.get_targets(args.cohort, args.wc)
    SUBTITLE_TAG = f"{cfg.week_label(args.cohort, args.wc)} · Programme Month {MONTH}"
    FSUF = f"_C{args.cohort}_WC_{args.wc}"
    period_sub = f"WC_{args.wc}"

root = Path(__file__).resolve().parent.parent
week_dir = Path(args.data) if args.data else root / f"Cohort_{args.cohort}" / period_sub
matches = sorted(week_dir.glob("KPI_Tracker_*.xlsx"),
                 key=lambda p: p.stat().st_mtime)
if not matches:
    sys.exit(f"No KPI_Tracker_*.xlsx found in {week_dir.resolve()} — "
             f"run build_tracker.py first.")
FILE = matches[-1]
OUT = week_dir / "charts"
OUT.mkdir(exist_ok=True)

print(f"Cohort {args.cohort}  |  {period_sub.replace(chr(95), chr(32))}  ->  Month {MONTH} targets")
print(f"Tracker: {FILE.name}")

raw = pd.read_excel(FILE, sheet_name="Agent Tracker", header=2)
raw = raw[raw["Agent"].notna()].copy()
raw = raw[raw["Agent"] != "▶ Targets"].reset_index(drop=True)

df = raw.rename(columns={
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
})

for col in ["RPH", "FCR", "QA", "Compliance", "CSAT_score", "D1", "D28", "D56",
            "Surveys", "QA_Evals", "Comp_Evals", "Prod_Hours", "Closed_Cases",
            "Complaints_Closed"]:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# Convert QA/Compliance from build_tracker.py's raw 0-100 scale to a 0-1
# fraction. These two are ALWAYS raw 0-100 (see build_tracker.py's load_all():
# "Compliance"/"QA Score" come straight from pd.to_numeric(...), never divided).
# FCR/CSAT_score/D1/D28/D56 are ALWAYS already a 0-1 fraction by the time they
# reach the Agent Tracker sheet (built via pct_to_float() / boolean .mean()).
#
# NOTE: this used to be guessed from each column's median (">1 means raw
# 0-100"), which breaks whenever a majority of agents score exactly 0% —
# median becomes 0, the /100 gets skipped, and a genuine 45.45% average
# renders as "4545%" (and, via off-canvas text in kpi_bars, blows the chart
# canvas out to thousands of pixels wide). Since we control both scripts,
# hard-code the scale instead of inferring it.
RAW_0_100_COLS = ["QA", "Compliance"]
for col in RAW_0_100_COLS:
    if col in df.columns:
        df[col] = df[col] / 100
# FCR, CSAT_score, D1, D28, D56: no conversion needed — already 0-1 fractions.

# Sanity-check: no percentage metric can realistically exceed 100%. A value
# this far outside range means a source-data/export error (e.g. the Kishay
# Davids AHT case) rather than a real score. Flag it and null it out instead
# of letting one bad row wreck cohort averages and chart geometry (off-canvas
# text stretching bbox_inches="tight" the same way the scale bug above did).
PCT_COLS = ["FCR", "QA", "Compliance", "CSAT_score", "D1", "D28", "D56"]
for col in PCT_COLS:
    if col in df.columns:
        bad = df[col] > 1.5   # >150% — not a plausible score
        if bad.any():
            for agent, val in df.loc[bad, ["Agent Name", col]].itertuples(index=False):
                print(f"  ! {agent}: {col} = {val:.1%} looks like a data error "
                      f"— excluded from charts, check the source export")
            df.loc[bad, col] = np.nan

# Targets as fractions where the data is a fraction
TARGETS = {
    "RPH":        T["RPH"],
    "FCR":        T["FCR"],
    "QA":         (T["QA Score"] / 100) if T["QA Score"] is not None else None,
    "Compliance": (T["Compliance"] / 100) if T["Compliance"] is not None else None,
    "CSAT_score": T["CSAT"],
    "D1":         T["D1"],
    "D28":        T["D28"],
    "D56":        T["D56"],
}

KPI_LABELS = {"RPH": "RPH", "FCR": "FCR", "QA": "QA", "Compliance": "Compliance",
              "CSAT_score": "CSAT", "D1": "D1", "D28": "D28", "D56": "D56"}
KPI_LONG = {"RPH": "Resolves per Hour", "FCR": "First Contact Resolution",
            "QA": "QA Score", "Compliance": "Compliance (PR)",
            "CSAT_score": "CSAT", "D1": "D1 Complaints (24h)",
            "D28": "D28 Complaints (28d)", "D56": "D56 Complaints (56d)"}

# KPIs that carry a target this month (order fixed for charts)
ORDERED = ["RPH", "FCR", "QA", "Compliance", "CSAT_score", "D1", "D28", "D56"]
ACTIVE = [k for k in ORDERED if TARGETS[k] is not None and k in df.columns]
TRACKED_ONLY = [k for k in ORDERED if TARGETS[k] is None and k in df.columns
                and df[k].notna().any()]
M_KPI = len(ACTIVE)

# Volume-weighted cohort averaging. A plain per-agent mean gives a 1-survey
# agent the same pull on the cohort average as a 6-survey agent — this is
# what caused the scorecard's CSAT tile (51.9%) to disagree with the
# platform's own pooled team-level CSAT (62.1%, see build_tracker.py's
# write_team_sheet). Weighting each agent's % by their underlying volume
# (Surveys for CSAT, Evals for QA/Compliance) reproduces the pooled result
# exactly — mathematically identical to summing raw counts and dividing.
WEIGHT_COL = {
    "CSAT_score": "Surveys",       # exact: Surveys is CSAT's own denominator
    "QA":         "QA_Evals",      # exact
    "Compliance": "Comp_Evals",    # exact
    "RPH":        "Prod_Hours",    # exact: RPH = Closed Cases ÷ Prod Hours
    "D1":         "Complaints_Closed",  # exact
    "D28":        "Complaints_Closed",  # exact
    "D56":        "Complaints_Closed",  # exact
    "FCR":        "Closed_Cases",  # approximation — the platform doesn't
                                    # export per-agent FCR success/fail counts,
                                    # so Closed Cases is the closest available
                                    # volume proxy, not an exact reconstruction
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

def is_pct(k):
    return k != "RPH"

def fmt_v(v, k, dp0=True):
    if pd.isna(v):
        return "—"
    if k == "RPH":
        return f"{v:.2f}"
    return f"{v:.0%}" if dp0 else f"{v:.1%}"

def tgt_str(k):
    t = TARGETS[k]
    return f"{t}" if k == "RPH" else f"{t:.0%}"

def classify(row):
    fast    = pd.notna(row["RPH"]) and row["RPH"] >= TARGETS["RPH"]
    quality = pd.notna(row["FCR"]) and row["FCR"] >= TARGETS["FCR"]
    if fast and quality:      return "High Performance"
    if not fast and quality:  return "Experience Focus"
    if fast and not quality:  return "Speed Risk"
    return "Underperforming"

df["quadrant"]   = df.apply(classify, axis=1)
df["first_name"] = df["Agent Name"].apply(lambda n: str(n).split()[0])
N = len(df)
print(f"Loaded {N} agents  |  {M_KPI} KPIs with targets: "
      f"{', '.join(KPI_LABELS[k] for k in ACTIVE)}"
      + (f"  |  tracked only: {', '.join(KPI_LABELS[k] for k in TRACKED_ONLY)}"
         if TRACKED_ONLY else ""))


# ═════════════════════════════════════════════════════════════════════════════
# CHART 1 — Performance Quadrant
# ═════════════════════════════════════════════════════════════════════════════
RPH_T, FCR_T = TARGETS["RPH"], TARGETS["FCR"]
x_max = max(df["RPH"].max() * 1.15, 3.8)
y_max = 1.05

fig, ax = plt.subplots(figsize=(16, 10))
fig.patch.set_facecolor(BG); ax.set_facecolor(BG)

for q, c in QUAD_BG.items():
    xlo = 0     if q in ("Experience Focus", "Underperforming") else RPH_T
    xhi = RPH_T if q in ("Experience Focus", "Underperforming") else x_max
    ylo = FCR_T if q in ("Experience Focus", "High Performance") else 0
    yhi = y_max if q in ("Experience Focus", "High Performance") else FCR_T
    ax.fill_between([xlo, xhi], ylo, yhi, color=c, zorder=0)

ax.axvline(RPH_T, color=MUTE, lw=1.4, ls="--", zorder=1)
ax.axhline(FCR_T, color=MUTE, lw=1.4, ls="--", zorder=1)
ax.text(RPH_T + 0.04, 0.01, f"RPH target\n{RPH_T}", fontsize=7.5, color=MUTE,
        fontstyle="italic", va="bottom")
ax.text(0.03, FCR_T + 0.005, f"FCR target {FCR_T:.0%}", fontsize=7.5, color=MUTE,
        fontstyle="italic", va="bottom")

quad_pos = {"Experience Focus": (0.02, 0.98, "left", "top"),
            "High Performance": (0.98, 0.98, "right", "top"),
            "Underperforming":  (0.02, 0.02, "left", "bottom"),
            "Speed Risk":       (0.98, 0.02, "right", "bottom")}
for q, (x, y, ha, va) in quad_pos.items():
    s = QUAD_LBL[q]
    ax.text(x, y, f"  {q}  ", transform=ax.transAxes, fontsize=12, fontweight="bold",
            color=s["fg"], ha=ha, va=va,
            bbox=dict(boxstyle="round,pad=0.3", facecolor="white",
                      edgecolor=s["border"], linewidth=1.6, alpha=0.95), zorder=5)

# Dots only — no per-point labels. Names are listed per quadrant instead.
for _, row in df.iterrows():
    c = QUAD_COLORS[row["quadrant"]]
    ax.scatter(row["RPH"], row["FCR"], s=150, color=c,
               edgecolors="white", linewidths=1.5, zorder=4, alpha=0.93)

# Name roster in each quadrant's corner (far clearer than point labels)
ROSTER_WRAP = {"High Performance": 46, "Experience Focus": 30,
               "Underperforming": 30, "Speed Risk": 30}
for q, (x, y, ha, va) in quad_pos.items():
    names = sorted(df.loc[df["quadrant"] == q, "first_name"].tolist())
    if not names:
        continue
    yoff = -0.055 if va == "top" else 0.055
    ax.text(x, y + yoff, textwrap.fill(", ".join(names), width=ROSTER_WRAP[q]),
            transform=ax.transAxes, ha=ha, va=va, fontsize=7.3, color=INK,
            linespacing=1.5, zorder=5,
            bbox=dict(boxstyle="round,pad=0.35", facecolor="white",
                      edgecolor=QUAD_LBL[q]["border"], linewidth=0.8, alpha=0.85))

ax.set_xlim(0, x_max); ax.set_ylim(0, y_max)
ax.set_xlabel("Resolves per Hour (RPH)", fontsize=12, fontweight="bold", labelpad=10, color=INK)
ax.set_ylabel("First Contact Resolution (FCR)", fontsize=12, fontweight="bold", labelpad=10, color=INK)
ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.0%}"))
ax.tick_params(colors=MUTE, labelsize=9)
for spine in ax.spines.values():
    spine.set_edgecolor(GRID)

counts = df["quadrant"].value_counts()
handles = [mpatches.Patch(color=QUAD_COLORS[q], label=f"{q}  ({counts.get(q, 0)} agents)")
           for q in ["High Performance", "Experience Focus", "Speed Risk", "Underperforming"]]
ax.legend(handles=handles, title="Quadrant Summary", title_fontsize=9, fontsize=8.5,
          loc="lower center", framealpha=0.95, edgecolor=GRID)

fig.text(0.5, 0.97, "Agent Performance Quadrant — Speed vs Quality",
         ha="center", fontsize=16, fontweight="bold", color=INK)
fig.text(0.5, 0.94, f"{SUBTITLE_TAG}  ·  RPH target {RPH_T}  ·  FCR target {FCR_T:.0%}",
         ha="center", fontsize=9, color=MUTE)
plt.tight_layout(rect=[0, 0, 1, 0.93])
plt.savefig(OUT / f"agent_quadrant{FSUF}.png", dpi=160, bbox_inches="tight", facecolor=BG)
plt.close(fig)
print(f"Saved -> agent_quadrant{FSUF}.png")


# ═════════════════════════════════════════════════════════════════════════════
# CHART 2 — All KPIs vs targets, per agent (small multiples)
# ═════════════════════════════════════════════════════════════════════════════
df_h = df.sort_values("RPH", ascending=True).reset_index(drop=True)
agents_h = df_h["first_name"].tolist()
y = np.arange(N)
BAR_H = 0.62
FIG_H = max(11, N * 0.42 + 3.5)

fig, axes = plt.subplots(1, M_KPI, figsize=(M_KPI * 3.8, FIG_H), sharey=True)
if M_KPI == 1:
    axes = [axes]
fig.patch.set_facecolor(BG)
fig.subplots_adjust(wspace=0.06, left=0.10, right=0.99, top=0.84, bottom=0.05)

for i, k in enumerate(ACTIVE):
    ax = axes[i]; ax.set_facecolor(BG)
    target = TARGETS[k]
    xmax = 5.5 if k == "RPH" else 1.0
    heights, colours = [], []
    for j in range(N):
        v = df_h[k].iloc[j]
        if pd.isna(v):
            heights.append(0.0); colours.append(GREY)
        else:
            heights.append(float(v))
            colours.append(TEAL if v >= target else RED)

    ax.barh(y, heights, height=BAR_H, color=colours, edgecolor=BG, linewidth=0.5, zorder=3)
    for j in range(N):
        ax.barh(j, xmax * 0.008, height=BAR_H, color=BRAND, zorder=4, edgecolor="none")
        v = df_h[k].iloc[j]
        if pd.isna(v):
            ax.barh(j, target * 0.08, height=BAR_H, color=GREY, edgecolor=BG,
                    hatch="///", alpha=0.45, zorder=3)
            ax.text(target * 0.10, j, "N/A", va="center", ha="left",
                    fontsize=6.5, color="#999", zorder=4)
            continue
        v = float(v)
        disp = fmt_v(v, k)
        if v > xmax * 0.18:
            ax.text(v - xmax * 0.01, j, disp, va="center", ha="right",
                    fontsize=7, color="white", fontweight="bold", zorder=5)
        else:
            ax.text(v + xmax * 0.01, j, disp, va="center", ha="left",
                    fontsize=7, color=INK, zorder=5)

    ax.axvline(target, color=BRAND, lw=1.8, ls="--", zorder=4, alpha=0.9)
    ax.text(target + xmax * 0.015, N - 0.5, f"Target {tgt_str(k)}",
            va="bottom", ha="left", fontsize=7.5, color=INK, fontstyle="italic", zorder=5)

    ax.set_xlim(0, xmax * 1.18); ax.set_ylim(-0.6, N - 0.4)
    if is_pct(k):
        ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.0%}"))
    ax.tick_params(axis="x", colors=MUTE, labelsize=8)
    ax.xaxis.grid(True, color=GRID, linewidth=0.6, zorder=0); ax.set_axisbelow(True)
    for spine in ax.spines.values():
        spine.set_visible(False)

    n_on = int((df_h[k].dropna() >= target).sum())
    ax.set_title(f"{KPI_LONG[k]}\n{n_on / N:.0%} on target",
                 fontsize=10, fontweight="bold", color=INK, pad=14)

    if i == M_KPI - 1:
        handles = [mpatches.Patch(color=TEAL, label="On target"),
                   mpatches.Patch(color=RED,  label="Below target"),
                   mpatches.Patch(color=GREY, label="No data")]
        ax.legend(handles=handles, fontsize=7.5, framealpha=0.92,
                  loc="lower right", edgecolor=GRID)

axes[0].set_yticks(y)
axes[0].set_yticklabels(agents_h, fontsize=8.5, color=INK)
axes[0].tick_params(axis="y", length=0)
tracked_note = (f"  ·  {'/'.join(KPI_LABELS[k] for k in TRACKED_ONLY)} tracked only "
                f"(no target — not shown)" if TRACKED_ONLY else "")
fig.text(0.5, 0.965, f"{SUBTITLE_TAG} — All {M_KPI} KPIs vs Targets · Agent Level",
         ha="center", fontsize=15, fontweight="bold", color=INK)
fig.text(0.5, 0.905, "Sorted by RPH (best at top)  ·  Red = below target  ·  "
         f"Grey = not yet assessed{tracked_note}",
         ha="center", fontsize=9, color=MUTE)
plt.savefig(OUT / f"kpi_bars{FSUF}.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(fig)
print(f"Saved -> kpi_bars{FSUF}.png")


# ═════════════════════════════════════════════════════════════════════════════
# CHART 3 — KPI Heatmap (RAG table incl. tracked-only columns)
# ═════════════════════════════════════════════════════════════════════════════
df_hm = df.sort_values("RPH", ascending=False).reset_index(drop=True)
HM_KPIS = ACTIVE + TRACKED_ONLY
M_HM = len(HM_KPIS)

CELL_W, CELL_H, NAME_W, PAD = 1.65, 0.70, 2.7, 0.07
total_w = NAME_W + M_HM * (CELL_W + PAD)
total_h = (N + 3.5) * (CELL_H + PAD)

fig, ax = plt.subplots(figsize=(max(14, M_HM * 1.9 + 3), max(11, N * 0.43 + 3.5)))
fig.patch.set_facecolor(BG); ax.set_facecolor(BG); ax.axis("off")
ax.set_xlim(0, total_w); ax.set_ylim(0, total_h)

def cell_colour(v, k):
    t = TARGETS[k]
    if pd.isna(v):
        return "#EFEFEF", "#888"
    if t is None:
        return "#F5F5F5", "#666"                       # tracked only
    if v >= t:
        return "#CCF7EE", "#007A62"
    if v >= t * 0.85:
        return "#F5FCB0", "#4A4600"
    return "#FFE5E6", "#CC0008"

def sub_label(row, k):
    """Second line inside a heatmap cell — context counts."""
    if k == "QA" and pd.notna(row.get("QA_Evals")):
        n = int(row["QA_Evals"]); return f"{n} eval{'s' if n != 1 else ''}"
    if k == "Compliance" and pd.notna(row.get("Comp_Evals")):
        n = int(row["Comp_Evals"])
        ok = pd.notna(row["Compliance"]) and row["Compliance"] >= TARGETS["Compliance"]
        return f"{n} eval{'s' if n != 1 else ''} · {'pass' if ok else 'risk'}"
    if k == "CSAT_score" and pd.notna(row.get("Surveys")):
        n = int(row["Surveys"]); return f"{n} survey{'s' if n != 1 else ''}"
    if TARGETS[k] is None and pd.notna(row.get(k)):
        return "tracked"
    return None

def draw_cell(x, y, w, h, bg, text_main, fg, text_sub=None, sub_col=None):
    ax.add_patch(plt.Rectangle((x, y), w, h, facecolor=bg, edgecolor="white",
                               linewidth=0.8, zorder=2))
    ty = y + h * (0.62 if text_sub else 0.5)
    ax.text(x + w / 2, ty, text_main, ha="center", va="center", fontsize=8.5,
            fontweight="bold", color=fg, zorder=3)
    if text_sub:
        ax.text(x + w / 2, y + h * 0.22, text_sub, ha="center", va="center",
                fontsize=6.6, color=sub_col or fg, zorder=3)

# Header row
hy = total_h - CELL_H - PAD
ax.add_patch(plt.Rectangle((0, hy), NAME_W, CELL_H, facecolor="#131409",
                           edgecolor="none", zorder=2))
ax.text(NAME_W / 2, hy + CELL_H / 2, "Agent", ha="center", va="center",
        fontsize=9.5, fontweight="bold", color="white", zorder=3)
for mi, k in enumerate(HM_KPIS):
    x = NAME_W + mi * (CELL_W + PAD)
    ax.add_patch(plt.Rectangle((x, hy), CELL_W, CELL_H, facecolor="#131409",
                               edgecolor="none", zorder=2))
    ax.text(x + CELL_W / 2, hy + CELL_H * 0.65, KPI_LABELS[k], ha="center",
            va="center", fontsize=9, fontweight="bold", color="white", zorder=3)
    sub = f"Target ≥ {tgt_str(k)}" if TARGETS[k] is not None else "Tracked — no target"
    ax.text(x + CELL_W / 2, hy + CELL_H * 0.25, sub, ha="center", va="center",
            fontsize=6.6, color=BRAND, zorder=3)

# Cohort average row
avg_y = hy - CELL_H * 1.1
ax.add_patch(plt.Rectangle((0, avg_y), NAME_W, CELL_H, facecolor="#E8E7E1",
                           edgecolor="white", linewidth=0.5, zorder=2))
ax.text(NAME_W / 2, avg_y + CELL_H / 2, "Cohort average", ha="center", va="center",
        fontsize=8.2, fontweight="bold", color=INK, zorder=3)
for mi, k in enumerate(HM_KPIS):
    x = NAME_W + mi * (CELL_W + PAD)
    v = weighted_mean(df_hm, k)
    bg_c, fg_c = cell_colour(v, k)
    if TARGETS[k] is not None and not pd.isna(v):
        d = v - TARGETS[k]
        dlt = (f"+{d:.2f}" if d >= 0 else f"{d:.2f}") if k == "RPH" else \
              (f"+{d:.0%}" if d >= 0 else f"{d:.0%}")
        dcol = "#007A62" if d >= 0 else "#CC0008"
    else:
        dlt, dcol = ("no target", "#888") if TARGETS[k] is None else (None, None)
    draw_cell(x, avg_y, CELL_W, CELL_H, bg_c, fmt_v(v, k), fg_c, dlt, dcol)

# Agent rows
row_start = avg_y - PAD
for ri, (_, row) in enumerate(df_hm.iterrows()):
    yy = row_start - (ri + 1) * (CELL_H + PAD)
    row_bg = "#FFFFFF" if ri % 2 == 0 else "#F5F5F5"
    ax.add_patch(plt.Rectangle((0, yy), NAME_W, CELL_H, facecolor=row_bg,
                               edgecolor="white", linewidth=0.5, zorder=2))
    ax.text(0.15, yy + CELL_H / 2, row["first_name"], ha="left", va="center",
            fontsize=8.2, color=INK, zorder=3)
    for mi, k in enumerate(HM_KPIS):
        x = NAME_W + mi * (CELL_W + PAD)
        v = row[k]
        bg_c, fg_c = cell_colour(v, k)
        draw_cell(x, yy, CELL_W, CELL_H, bg_c, fmt_v(v, k), fg_c,
                  sub_label(row, k), fg_c)

legend_patches = [
    mpatches.Patch(facecolor="#CCF7EE", edgecolor="#007A62", label="On / above target"),
    mpatches.Patch(facecolor="#F5FCB0", edgecolor="#4A4600", label="Within 15% of target"),
    mpatches.Patch(facecolor="#FFE5E6", edgecolor="#CC0008", label="Below target"),
    mpatches.Patch(facecolor="#F5F5F5", edgecolor="#999",    label="Tracked — no target"),
    mpatches.Patch(facecolor="#EFEFEF", edgecolor="#888",    label="Not yet assessed"),
]
fig.legend(handles=legend_patches, loc="lower center", ncol=5, fontsize=8.5,
           framealpha=0.95, edgecolor=GRID, bbox_to_anchor=(0.5, 0.001))
fig.text(0.5, 0.995, f"{SUBTITLE_TAG} — KPI Heatmap: All Metrics",
         ha="center", fontsize=15, fontweight="bold", color=INK)
fig.text(0.5, 0.977, "Sorted by RPH  ·  QA & CSAT sub-labels show eval/survey count"
         + ("  ·  " + "/".join(KPI_LABELS[k] for k in TRACKED_ONLY) +
            " tracked only (no target this period)" if TRACKED_ONLY else ""),
         ha="center", fontsize=8.5, color=MUTE)
plt.tight_layout(rect=[0, 0.03, 1, 0.975])
plt.savefig(OUT / f"kpi_heatmap{FSUF}.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(fig)
print(f"Saved -> kpi_heatmap{FSUF}.png")


# ═════════════════════════════════════════════════════════════════════════════
# CHART 4 — KPI Pareto (fails per agent + cumulative line)
# ═════════════════════════════════════════════════════════════════════════════
records = []
for _, row in df.iterrows():
    on, off, na = [], [], []
    for k in ACTIVE:
        v = row[k]
        if pd.isna(v):
            na.append(KPI_LABELS[k])
        elif float(v) >= TARGETS[k]:
            on.append(KPI_LABELS[k])
        else:
            off.append(KPI_LABELS[k])
    records.append({"agent": row["first_name"], "on": len(on), "off": len(off),
                    "na": len(na), "fail_list": off})

rdf = pd.DataFrame(records).sort_values(["off", "on"], ascending=[True, False]).reset_index(drop=True)
agents_p = rdf["agent"].tolist()
on_v, off_v, na_v = rdf["on"].tolist(), rdf["off"].tolist(), rdf["na"].tolist()
total_fails = sum(off_v)
cumulative = (np.cumsum(off_v) / total_fails * 100) if total_fails > 0 else np.zeros(N)
assessed_slots = N * M_KPI - sum(na_v)
on_slots = sum(on_v)

fig, ax1 = plt.subplots(figsize=(20, 9))
fig.patch.set_facecolor(BG); ax1.set_facecolor(BG)
ax2 = ax1.twinx()
x = np.arange(N); bar_w = 0.60

ax1.bar(x, na_v,  width=bar_w, color=GREY, zorder=3, edgecolor=BG, linewidth=0.8,
        bottom=[o + f for o, f in zip(on_v, off_v)], label="Not yet assessed")
ax1.bar(x, off_v, width=bar_w, color=RED, zorder=3, edgecolor=BG, linewidth=0.8,
        bottom=on_v, label="Below target")
ax1.bar(x, on_v,  width=bar_w, color=TEAL, zorder=3, edgecolor=BG, linewidth=0.8,
        label="On target")
for i in range(N):
    ax1.bar(i, 0.06, width=bar_w, color=BRAND, zorder=4, edgecolor="none")
    if on_v[i] >= 1:
        ax1.text(i, on_v[i] / 2, str(on_v[i]), ha="center", va="center",
                 fontsize=9, color="white", fontweight="bold", zorder=6)
    if off_v[i] >= 1:
        ax1.text(i, on_v[i] + off_v[i] / 2, str(off_v[i]), ha="center", va="center",
                 fontsize=9, color="white", fontweight="bold", zorder=6)
    if na_v[i] >= 1:
        ax1.text(i, on_v[i] + off_v[i] + na_v[i] / 2, str(na_v[i]), ha="center",
                 va="center", fontsize=8, color=INK, fontweight="bold", zorder=6)
    fails = rdf["fail_list"].iloc[i]
    if fails:
        ax1.text(i, on_v[i] + off_v[i] + na_v[i] + 0.12, "\n".join(fails),
                 ha="center", va="bottom", fontsize=7, color="#4A4600",
                 fontstyle="italic", zorder=6)

ax2.plot(x, cumulative, color=NAVY, lw=2.0, marker="o", markersize=5.5,
         markerfacecolor="white", markeredgecolor=NAVY, markeredgewidth=1.6, zorder=7)

ax1.set_xlim(-0.7, N - 0.3); ax1.set_ylim(0, M_KPI + 2.2)
ax1.set_yticks(range(M_KPI + 1))
ax1.set_ylabel(f"Number of KPIs (out of {M_KPI})", fontsize=10.5, color=INK, labelpad=10)
ax1.set_xticks(x)
ax1.set_xticklabels(agents_p, rotation=35, ha="right", fontsize=9.5, color=INK)
ax1.tick_params(axis="both", length=0)
ax1.yaxis.grid(True, color=GRID, linewidth=0.7, zorder=0); ax1.set_axisbelow(True)
for spine in ax1.spines.values():
    spine.set_visible(False)

ax2.set_ylim(0, 118); ax2.set_yticks([0, 20, 40, 60, 80, 100])
ax2.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.0f}%"))
ax2.set_ylabel("Cumulative % of all KPI fails", fontsize=10.5, color=NAVY, labelpad=10)
ax2.tick_params(axis="y", colors=NAVY, labelsize=9, length=0)
for spine in ax2.spines.values():
    spine.set_visible(False)

n_clean = (rdf["off"] == 0).sum(); n_one = (rdf["off"] == 1).sum(); n_two = (rdf["off"] == 2).sum()
fig.text(0.5, 0.915,
         f"{on_slots / assessed_slots:.0%} of assessed KPI slots on target "
         f"({on_slots}/{assessed_slots})     |     0 fails: {n_clean} agents     "
         f"1 fail: {n_one} agents     2 fails: {n_two} agents",
         ha="center", va="top", fontsize=9.5, color=INK,
         bbox=dict(boxstyle="round,pad=0.5", facecolor="white", edgecolor=GRID,
                   linewidth=1, alpha=0.95))

handles = [mpatches.Patch(color=TEAL, label="On target"),
           mpatches.Patch(color=RED,  label="Below target"),
           mpatches.Patch(color=GREY, label="Not yet assessed"),
           plt.Line2D([0], [0], color=NAVY, lw=2, marker="o", markerfacecolor="white",
                      markeredgecolor=NAVY, markersize=5, label="Cumulative % of fails")]
ax1.legend(handles=handles, fontsize=9, framealpha=0.95, loc="upper left",
           edgecolor=GRID, handlelength=1.4, borderpad=0.7, labelspacing=0.5)

fig.text(0.5, 0.975, f"{SUBTITLE_TAG} — KPI Pareto: Agent Target Compliance",
         ha="center", fontsize=15, fontweight="bold", color=INK)
fig.text(0.5, 0.955, "Sorted by KPI fails · italic = KPIs missed · line = cumulative "
         f"share of all fails  ·  {M_KPI} KPIs tracked"
         + (f" ({'/'.join(KPI_LABELS[k] for k in TRACKED_ONLY)} excluded — no target)"
            if TRACKED_ONLY else ""),
         ha="center", fontsize=9, color=MUTE)
plt.tight_layout(rect=[0, 0, 1, 0.91])
plt.savefig(OUT / f"kpi_pareto{FSUF}.png", dpi=150, bbox_inches="tight", facecolor=BG)
plt.close(fig)
print(f"Saved -> kpi_pareto{FSUF}.png")


# ═════════════════════════════════════════════════════════════════════════════
# CHART 5 — KPI Scorecard tiles
# ═════════════════════════════════════════════════════════════════════════════
stats = []
for k in ACTIVE:
    series = df[k].dropna()
    avg = weighted_mean(df, k)
    t = TARGETS[k]
    if np.isnan(avg):
        status, colour = "NO DATA", GREY
    elif avg >= t:
        status, colour = "ON TRACK", TEAL
    elif avg >= t * 0.85:
        status, colour = "AT RISK", AMBER
    else:
        status, colour = "BELOW TARGET", RED
    n_meet = int((series >= t).sum())
    delta = avg - t if not np.isnan(avg) else np.nan
    stats.append(dict(k=k, avg=avg, t=t, status=status, colour=colour,
                      n_valid=len(series), n_meet=n_meet, delta=delta))

ORDER = {"ON TRACK": 0, "AT RISK": 1, "BELOW TARGET": 2, "NO DATA": 3}
stats.sort(key=lambda s: ORDER[s["status"]])

M_SC = len(stats)
fig, axes = plt.subplots(1, M_SC, figsize=(M_SC * 2.6 + 0.4, 5.5))
if M_SC == 1:
    axes = [axes]
fig.patch.set_facecolor(BG)
fig.subplots_adjust(left=0.02, right=0.98, top=0.86, bottom=0.08, wspace=0.04)

for ax, s in zip(axes, stats):
    ax.set_facecolor("white"); ax.set_xlim(0, 1); ax.set_ylim(0, 1); ax.axis("off")
    for spine in ax.spines.values():
        spine.set_visible(True); spine.set_edgecolor("#DDD"); spine.set_linewidth(0.8)
    ax.set_frame_on(True)
    ax.add_patch(plt.Rectangle((0, 0.88), 1, 0.12, facecolor=s["colour"],
                               transform=ax.transAxes, clip_on=False, zorder=3))
    ax.text(0.5, 0.78, KPI_LONG[s["k"]], ha="center", va="center", fontsize=9,
            color=MUTE, transform=ax.transAxes)
    val = fmt_v(s["avg"], s["k"], dp0=False) if not np.isnan(s["avg"]) else "N/A"
    ax.text(0.5, 0.54, val, ha="center", va="center", fontsize=23,
            fontweight="bold", color=INK, transform=ax.transAxes)
    ax.text(0.5, 0.35, f"Target: ≥ {tgt_str(s['k'])}", ha="center", va="center",
            fontsize=8.5, color=MUTE, transform=ax.transAxes)
    if not np.isnan(s["delta"]):
        sign = "+" if s["delta"] >= 0 else ""
        d_str = (f"{sign}{s['delta']:.2f}" if s["k"] == "RPH"
                 else f"{sign}{s['delta']:.0%}") + " vs target"
        ax.text(0.5, 0.27, d_str, ha="center", va="center", fontsize=7.5,
                color=TEAL if s["delta"] >= 0 else RED, transform=ax.transAxes)
    ax.text(0.5, 0.14, s["status"], ha="center", va="center", fontsize=9,
            fontweight="bold", color=s["colour"], transform=ax.transAxes)
    sub = (f"{s['n_meet']}/{s['n_valid']} on target" if s["n_valid"] == N
           else f"{s['n_valid']}/{N} assessed")
    ax.text(0.5, 0.05, sub, ha="center", va="center", fontsize=7.5, color=MUTE,
            transform=ax.transAxes)

fig.text(0.5, 0.96, f"{SUBTITLE_TAG} — KPI Scorecard vs Programme Targets",
         ha="center", fontsize=14, fontweight="bold", color=INK)
fig.text(0.5, 0.90, f"Ordered: Green → Amber → Red  ·  Month {MONTH} targets"
         + (f"  ·  {'/'.join(KPI_LABELS[k] for k in TRACKED_ONLY)} tracked only "
            f"(not shown here)" if TRACKED_ONLY else ""),
         ha="center", fontsize=8.5, color=MUTE)
plt.savefig(OUT / f"kpi_scorecard{FSUF}.png", dpi=160, bbox_inches="tight", facecolor=BG)
plt.close(fig)
print(f"Saved -> kpi_scorecard{FSUF}.png")


# ═════════════════════════════════════════════════════════════════════════════
# CHARTS 6 & 7 — Team Leader breakdowns
# ═════════════════════════════════════════════════════════════════════════════
if "Team Leader" in df.columns and df["Team Leader"].notna().any():
    tls = sorted(df["Team Leader"].dropna().unique())
    tl_colour = {tl: TL_COLORS[i % len(TL_COLORS)] for i, tl in enumerate(tls)}

    # Per-TL average for each active KPI, as % of target
    tl_stats = {}
    for tl in tls:
        sub = df[df["Team Leader"] == tl]
        tl_stats[tl] = {}
        for k in ACTIVE:
            avg = weighted_mean(sub, k)
            tl_stats[tl][k] = dict(
                avg=avg,
                pct=(avg / TARGETS[k]) if not pd.isna(avg) else np.nan)

    # ── Chart 6: detailed KPI breakdown ──────────────────────────────────────
    ks = list(reversed(ACTIVE))          # RPH at bottom, like the original
    n_tl = len(tls)
    bar_h = 0.8 / n_tl
    fig, ax = plt.subplots(figsize=(16, max(8, M_KPI * 1.3)))
    fig.patch.set_facecolor(BG); ax.set_facecolor(BG)

    for ki, k in enumerate(ks):
        for ti, tl in enumerate(tls):
            s = tl_stats[tl][k]
            if np.isnan(s["pct"]):
                continue
            ypos = ki + (ti - (n_tl - 1) / 2) * (bar_h + 0.04)
            below = s["pct"] < 1.0
            ax.barh(ypos, s["pct"] * 100, height=bar_h,
                    color=RED if below else tl_colour[tl],
                    edgecolor=BG, linewidth=0.5, zorder=3)
            ax.barh(ypos, 1.2, height=bar_h, color=BRAND, zorder=4, edgecolor="none")
            ax.text(s["pct"] * 100 - 1.5, ypos, fmt_v(s["avg"], k),
                    va="center", ha="right", fontsize=9, color="white",
                    fontweight="bold", zorder=5)

    ax.axvline(100, color=BRAND, lw=1.8, ls="--", zorder=4, alpha=0.9)
    ax.set_yticks(range(M_KPI))
    ax.set_yticklabels([KPI_LABELS[k] for k in ks], fontsize=12,
                       fontweight="bold", color=INK)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.0f}%"))
    ax.set_xlabel(f"Score as % of Month {MONTH} target  (100% = target met)",
                  fontsize=10.5, color=MUTE, labelpad=8)
    ax.tick_params(colors=MUTE, labelsize=9)
    ax.xaxis.grid(True, color=GRID, linewidth=0.6, zorder=0); ax.set_axisbelow(True)
    for spine in ax.spines.values():
        spine.set_visible(False)

    handles = [mpatches.Patch(color=tl_colour[tl], label=tl) for tl in tls]
    handles.append(mpatches.Patch(color=RED, label="Below target"))
    ax.legend(handles=handles, fontsize=9.5, framealpha=0.95,
              loc="upper right", edgecolor=GRID)
    ax.set_title(f"TL Detailed KPI Performance Breakdown — Month {MONTH} Targets\n",
                 fontsize=15, fontweight="bold", color=INK, pad=18)
    ax.text(0.5, 1.015, SUBTITLE_TAG, transform=ax.transAxes,
            ha="center", fontsize=9, color=MUTE)
    plt.tight_layout()
    plt.savefig(OUT / f"tl_kpi_breakdown{FSUF}.png", dpi=150, bbox_inches="tight", facecolor=BG)
    plt.close(fig)
    print(f"Saved -> tl_kpi_breakdown{FSUF}.png")

    # ── Chart 7: average team performance ────────────────────────────────────
    fig, ax = plt.subplots(figsize=(12, 2 + n_tl * 1.6))
    fig.patch.set_facecolor(BG); ax.set_facecolor(BG)
    yp = np.arange(n_tl)
    avgs = []
    for tl in tls:
        pcts = [tl_stats[tl][k]["pct"] for k in ACTIVE
                if not np.isnan(tl_stats[tl][k]["pct"])]
        avgs.append(np.mean(pcts) * 100 if pcts else np.nan)

    for i, (tl, v) in enumerate(zip(tls, avgs)):
        if np.isnan(v):
            continue
        ax.barh(i, v, height=0.55, color=tl_colour[tl] if v >= 100 else RED,
                edgecolor=BG, zorder=3)
        ax.barh(i, 1.2, height=0.55, color=BRAND, zorder=4, edgecolor="none")
        ax.text(v - 2, i, f"{v:.0f}%", va="center", ha="right", fontsize=15,
                color="white", fontweight="bold", zorder=5)

    ax.axvline(100, color=BRAND, lw=1.8, ls="--", zorder=4, alpha=0.9)
    ax.set_yticks(yp); ax.set_yticklabels(tls, fontsize=12, fontweight="bold", color=MUTE)
    ax.xaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"{v:.0f}%"))
    ax.set_xlabel(f"Average KPI performance vs Month {MONTH} target",
                  fontsize=10.5, color=MUTE, labelpad=8)
    ax.set_xlim(0, max(120, max(v for v in avgs if not np.isnan(v)) * 1.1))
    ax.tick_params(colors=MUTE, labelsize=9)
    ax.xaxis.grid(True, color=GRID, linewidth=0.6, zorder=0); ax.set_axisbelow(True)
    for spine in ax.spines.values():
        spine.set_visible(False)
    all_met = all(not np.isnan(v) and v >= 100 for v in avgs)
    ax.set_title(f"Average Team Performance vs Month {MONTH} Targets\n",
                 fontsize=15, fontweight="bold", color=INK, pad=16)
    ax.text(0.5, 1.03, SUBTITLE_TAG, transform=ax.transAxes,
            ha="center", fontsize=9, color=MUTE)
    if all_met:
        ax.text(1.0, 1.13, "All targets met", transform=ax.transAxes,
                ha="right", fontsize=10, fontstyle="italic", color=INK)
    plt.tight_layout()
    plt.savefig(OUT / f"tl_team_average{FSUF}.png", dpi=150, bbox_inches="tight", facecolor=BG)
    plt.close(fig)
    print(f"Saved -> tl_team_average{FSUF}.png")
else:
    print("No Team Leader data — skipped TL charts (6 & 7)")

print(f"\nAll charts saved to: {OUT.resolve()}")