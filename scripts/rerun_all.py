"""
Batch Rerun — re-run build_tracker.py + dashboard.py across every week
============================================================================
Discovers every Cohort_N/WC_<date>/raw_data and Cohort_N/Month_<M>/raw_data
folder under the project root and re-runs the tracker + dashboard pipeline for
each one, in order. Useful for backfilling every past week after a fix to
build_tracker.py or dashboard.py, so every existing tracker/chart set gets
regenerated with the corrected logic rather than just the next new week.

Usage (run from the project root folder):
    python scripts/rerun_all.py
        -> finds every Cohort_N/WC_.../raw_data and Cohort_N/Month_.../raw_data
           folder, and for each one runs:
               build_tracker.py --cohort N --wc <date>   (or --month M)
               dashboard.py     --cohort N --wc <date>   (or --month M)

    python scripts/rerun_all.py --cohorts 1 2
        -> restrict to specific cohorts (default: every cohort found)

    python scripts/rerun_all.py --skip-dashboard
        -> only rebuild the Excel trackers, skip chart regeneration (faster
           if you just changed a calculation and want the xlsx refreshed)

    python scripts/rerun_all.py --dry-run
        -> list what would be run, without actually running anything

    python scripts/rerun_all.py --stop-on-error
        -> abort the whole batch on the first failure, instead of the
           default behaviour (log it, continue with the rest)

A summary table is printed at the end showing which weeks succeeded, which
failed, and which were skipped, so a partial run is easy to spot and re-run.
"""

import argparse
import re
import subprocess
import sys

# Force UTF-8 output — see build_tracker.py for why this is needed on Windows.
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")
import time
from pathlib import Path

# ── Args ──────────────────────────────────────────────────────────────────────
parser = argparse.ArgumentParser(description=__doc__,
                                  formatter_class=argparse.RawDescriptionHelpFormatter)
parser.add_argument("--cohorts", type=int, nargs="+", default=None,
                    help="Restrict to specific cohort numbers (default: all found)")
parser.add_argument("--data", default=None,
                    help="Override the project root folder (default: parent of scripts/)")
parser.add_argument("--skip-dashboard", action="store_true",
                    help="Only rebuild trackers; skip dashboard.py")
parser.add_argument("--skip-tracker", action="store_true",
                    help="Only regenerate dashboards; skip build_tracker.py "
                         "(the tracker xlsx must already exist for each week)")
parser.add_argument("--dry-run", action="store_true",
                    help="List what would be run without running anything")
parser.add_argument("--stop-on-error", action="store_true",
                    help="Abort the whole batch on the first failure "
                         "(default: log it and keep going)")
args = parser.parse_args()

if args.skip_dashboard and args.skip_tracker:
    sys.exit("--skip-dashboard and --skip-tracker together would run nothing.")

root = Path(args.data) if args.data else Path(__file__).resolve().parent.parent
scripts_dir = Path(__file__).resolve().parent
build_tracker_py = scripts_dir / "build_tracker.py"
dashboard_py = scripts_dir / "dashboard.py"


# ── Discover every WC/Month folder that has raw_data ─────────────────────────
def discover_weeks(root, cohorts_filter, verbose=True):
    weeks = []  # list of dicts: cohort, mode ('wc'|'month'), value, raw_dir
    cohort_dirs = sorted([p for p in root.iterdir() if p.is_dir()
                          and re.match(r"cohort[\s_]*\d+$", p.name, re.IGNORECASE)])

    if verbose and not cohort_dirs:
        print(f"  ! No 'Cohort_N' folders found directly under {root.resolve()} — "
              f"is this the project root? (the folder that itself contains "
              f"Cohort_1, Cohort_2, etc.)")

    for cohort_dir in cohort_dirs:
        m = re.match(r"cohort[\s_]*(\d+)$", cohort_dir.name, re.IGNORECASE)
        cohort = int(m.group(1))
        if cohorts_filter and cohort not in cohorts_filter:
            continue

        sub_dirs = sorted([p for p in cohort_dir.iterdir() if p.is_dir()])
        if verbose:
            print(f"  {cohort_dir.name}/ — {len(sub_dirs)} subfolder(s):")

        for sub in sub_dirs:
            wc_m = re.match(r"wc[\s_]*(.+)$", sub.name, re.IGNORECASE)
            month_m = re.match(r"month[\s_]*(\d+)$", sub.name, re.IGNORECASE)
            has_raw = (sub / "raw_data").exists()

            if wc_m:
                mode, value = "wc", wc_m.group(1).strip()
            elif month_m:
                mode, value = "month", int(month_m.group(1))
            else:
                if verbose:
                    print(f"    - {sub.name}/  (doesn't match 'WC_<date>' or "
                          f"'Month_<N>' — skipped)")
                continue

            if not has_raw:
                if verbose:
                    print(f"    - {sub.name}/  matched, but no 'raw_data' subfolder "
                          f"inside it — skipped. Drop this week's CSVs into "
                          f"{sub.name}/raw_data/ for it to be picked up.")
                continue

            if verbose:
                print(f"    \u2713 {sub.name}/raw_data  -> will run "
                      f"{'--wc ' + value if mode == 'wc' else '--month ' + str(value)}")
            weeks.append(dict(cohort=cohort, mode=mode, value=value,
                              raw_dir=sub / "raw_data", week_dir=sub))

    weeks.sort(key=lambda w: (w["cohort"], w["week_dir"].name))
    return weeks


cohorts_filter = set(args.cohorts) if args.cohorts else None
print(f"Scanning {root.resolve()} ...")
weeks = discover_weeks(root, cohorts_filter)
print()

if not weeks:
    sys.exit(
        "No runnable week/month folders found (see the scan above for why "
        "each folder was or wasn't picked up). Most common causes:\n"
        "  - CSVs are sitting directly in the WC_.../ or Month_.../ folder "
        "instead of a 'raw_data' subfolder inside it\n"
        "  - this isn't the project root (the folder that directly contains "
        "Cohort_1, Cohort_2, ...) — check with --data <path> if needed")

print(f"Will run {len(weeks)} week(s)/month(s):")
for w in weeks:
    label = f"W/C {w['value']}" if w["mode"] == "wc" else f"Month {w['value']}"
    print(f"  Cohort {w['cohort']}  ·  {label}")
print()

if args.dry_run:
    print("--dry-run: nothing was actually run.")
    sys.exit(0)


# ── Run each one ──────────────────────────────────────────────────────────────
def run_step(script, cohort, mode, value):
    cmd = [sys.executable, str(script), "--cohort", str(cohort)]
    cmd += ["--wc", value] if mode == "wc" else ["--month", str(value)]
    result = subprocess.run(cmd, cwd=scripts_dir, capture_output=True,
                            text=True, encoding="utf-8", errors="replace")
    return result.returncode == 0, result.stdout, result.stderr


def wait_for_tracker(week_dir, timeout=20, poll=0.5):
    """build_tracker.py exiting successfully doesn't guarantee the .xlsx is
    immediately readable elsewhere — on a OneDrive-synced folder the file
    can briefly be locked or still mid-upload right after being written.
    Poll until the newest KPI_Tracker_*.xlsx can actually be opened as a
    valid zip before handing off to dashboard.py, rather than assuming a
    completed subprocess means the file is ready. Returns (ready, detail)."""
    import zipfile
    deadline = time.time() + timeout
    last_err = "no tracker file appeared"
    while time.time() < deadline:
        matches = sorted(week_dir.glob("KPI_Tracker_*.xlsx"),
                         key=lambda p: p.stat().st_mtime) if week_dir.exists() else []
        if matches:
            candidate = matches[-1]
            try:
                with zipfile.ZipFile(candidate) as zf:
                    zf.testzip()
                return True, candidate
            except Exception as e:
                last_err = f"{candidate.name} not yet readable ({e})"
        time.sleep(poll)
    return False, last_err


results = []
t0 = time.time()

for i, w in enumerate(weeks, 1):
    label = f"Cohort {w['cohort']} · " + (f"W/C {w['value']}" if w["mode"] == "wc"
                                          else f"Month {w['value']}")
    print(f"[{i}/{len(weeks)}] {label}")

    row = dict(label=label, tracker="skipped", dashboard="skipped")

    if not args.skip_tracker:
        ok, out, err = run_step(build_tracker_py, w["cohort"], w["mode"], w["value"])
        row["tracker"] = "OK" if ok else "FAILED"
        print(f"    build_tracker.py: {'OK' if ok else 'FAILED'}")
        if not ok:
            print(f"      {err.strip().splitlines()[-1] if err.strip() else '(no stderr)'}")
            results.append(row)
            if args.stop_on_error:
                sys.exit(f"Stopping on first error ({label}). Use without "
                         f"--stop-on-error to continue past failures.")
            continue

        if not args.skip_dashboard:
            ready, detail = wait_for_tracker(w["week_dir"])
            if not ready:
                print(f"    ! Tracker file wasn't confirmed readable after "
                      f"waiting ({detail}) — proceeding anyway, but if "
                      f"dashboard.py fails or looks stale for this week, "
                      f"re-run just this one.")
            elif detail != "no tracker file appeared":
                pass  # ready quickly, nothing to report

    if not args.skip_dashboard:
        ok, out, err = run_step(dashboard_py, w["cohort"], w["mode"], w["value"])
        row["dashboard"] = "OK" if ok else "FAILED"
        print(f"    dashboard.py:     {'OK' if ok else 'FAILED'}")
        if not ok:
            print(f"      {err.strip().splitlines()[-1] if err.strip() else '(no stderr)'}")
            if args.stop_on_error:
                results.append(row)
                sys.exit(f"Stopping on first error ({label}). Use without "
                         f"--stop-on-error to continue past failures.")

    results.append(row)

elapsed = time.time() - t0


# ── Summary ───────────────────────────────────────────────────────────────────
print(f"\n{'=' * 70}")
print(f"Done in {elapsed:.0f}s — {len(results)} week(s)/month(s) processed")
print(f"{'=' * 70}")
w_label = max(len(r["label"]) for r in results) + 2
print(f"{'Week'.ljust(w_label)}{'Tracker'.ljust(10)}{'Dashboard'.ljust(10)}")
for r in results:
    print(f"{r['label'].ljust(w_label)}{r['tracker'].ljust(10)}{r['dashboard'].ljust(10)}")

failed = [r for r in results if "FAILED" in (r["tracker"], r["dashboard"])]
if failed:
    print(f"\n{len(failed)} failure(s) — re-run just those with, e.g.:")
    for r in failed:
        print(f"  (see log above for {r['label']})")
    sys.exit(1)
else:
    print("\nAll weeks completed successfully.")
