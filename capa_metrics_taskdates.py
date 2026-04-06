"""
CAPA KPI Metrics Calculator — Task Date Priority Version
=========================================================

ALTERNATIVE LOGIC (for comparison against official method)

HOW THIS VERSION DIFFERS FROM THE OFFICIAL VERSION:

  Official version  → Closed/open status and avg days both use the
                       "Date closed" field on the main Capas tab only.

  This version      → Closed/open status and avg days use the latest
                       completed task date from the Taken sheet if
                       available, falling back to the Capas Date closed,
                       and finally open if neither exists.

HOW WE DETERMINE IF A CAPA IS CLOSED HERE:

  Priority 1 — Taken (actions) sheet:
    If any tasks are marked Completed = Yes with a Date of completion,
    the CAPA is closed. Closed date = latest of those task dates.
    Incomplete tasks on the same CAPA are ignored.

  Priority 2 — Capas sheet fallback:
    If no completed tasks with dates exist, use Date closed from Capas.

  Priority 3 — Open:
    If neither source provides a date, the CAPA is open.

  The Status column is NOT used to determine open/closed.

SOURCE FILES:
  Reads all files matching: export_CAPA *.xls in the same folder.
  Sheets used: "Capas" (main record) + "Taken" (action tasks).
"""

import pandas as pd
import glob
import os
import warnings
from datetime import date

warnings.filterwarnings("ignore", category=UserWarning)

# ── Config ───────────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
FILES_PATTERN = os.path.join(SCRIPT_DIR, "export_CAPA *.xls")
TODAY = date.today()
OPEN_THRESHOLD_DAYS = 90

# ── Load all files ────────────────────────────────────────────────────────────
files = sorted(glob.glob(FILES_PATTERN))
if not files:
    print(f"No files found matching: {FILES_PATTERN}")
    exit(1)

capas_frames = []
for filepath in files:
    location = os.path.basename(filepath).replace("export_CAPA ", "").replace(".xls", "")

    capas = pd.read_excel(filepath, sheet_name="Capas")
    capas["Location"] = location
    capas["Date of notification"] = pd.to_datetime(
        capas["Date of notification"], dayfirst=True, errors="coerce"
    )
    capas["Date closed"] = pd.to_datetime(
        capas["Date closed"], dayfirst=True, errors="coerce"
    )

    taken = pd.read_excel(filepath, sheet_name="Taken")
    taken["Date of completion"] = pd.to_datetime(
        taken["Date of completion"], dayfirst=True, errors="coerce"
    )

    task_groups = taken.groupby("Number")

    def resolve_closed_date(row):
        num = row["Number"]
        capas_date = row["Date closed"]
        if num in task_groups.groups:
            group = task_groups.get_group(num)
            completed = group[group["Completed"].str.strip().str.lower() == "yes"]
            max_date = completed["Date of completion"].dropna().max()
            if pd.notna(max_date):
                return max_date   # latest completed task date
        return capas_date         # fall back to Capas sheet

    capas["Effective closed date"] = capas.apply(resolve_closed_date, axis=1)
    capas_frames.append(capas)

all_capas = pd.concat(capas_frames, ignore_index=True)

# ── Derived flags — based on Effective closed date ───────────────────────────
is_closed = all_capas["Effective closed date"].notna()
is_open   = ~is_closed

closed_2025 = is_closed & (all_capas["Effective closed date"].dt.year == 2025)
closed_2026 = is_closed & (all_capas["Effective closed date"].dt.year == 2026)

days_open = (pd.Timestamp(TODAY) - all_capas["Date of notification"]).dt.days
open_gt90 = is_open & (days_open > OPEN_THRESHOLD_DAYS)

# ── Avg days to close — uses Effective closed date ───────────────────────────
def avg_days(mask):
    df = all_capas[mask].copy()
    df["days_to_close"] = (
        df["Effective closed date"] - df["Date of notification"]
    ).dt.days
    vals = df["days_to_close"].dropna()
    return vals.mean() if not vals.empty else float("nan")

avg_days_to_close_2025 = avg_days(closed_2025)
avg_days_to_close_2026 = avg_days(closed_2026)

total_closed_2025 = closed_2025.sum()
total_closed_2026 = closed_2026.sum()
total_open        = is_open.sum()
total_open_gt90   = open_gt90.sum()

# ── Per-location breakdown ────────────────────────────────────────────────────
location_rows = []
for loc in sorted(all_capas["Location"].unique()):
    mask = all_capas["Location"] == loc

    loc_avg_2025 = avg_days(mask & closed_2025)
    loc_avg_2026 = avg_days(mask & closed_2026)

    location_rows.append({
        "Location":                           loc,
        "Avg Days to Close (2025)":           round(loc_avg_2025, 1) if pd.notna(loc_avg_2025) else "N/A",
        "Closed 2025":                        int((mask & closed_2025).sum()),
        "Avg Days to Close (2026)":           round(loc_avg_2026, 1) if pd.notna(loc_avg_2026) else "N/A",
        "Closed 2026":                        int((mask & closed_2026).sum()),
        "Open":                               int((mask & is_open).sum()),
        f"Open > {OPEN_THRESHOLD_DAYS} days": int((mask & open_gt90).sum()),
    })

df_location = pd.DataFrame(location_rows)

# ── Print summary ─────────────────────────────────────────────────────────────
print("=" * 62)
print("  CAPA KPI METRICS SUMMARY (Task Date Priority)")
print(f"  Run date: {TODAY}")
print("=" * 62)
print(f"  Avg days to close (2025):   {avg_days_to_close_2025:.1f}")
print(f"  Total closed in 2025:       {total_closed_2025}")
print(f"  Avg days to close (2026):   {avg_days_to_close_2026:.1f}")
print(f"  Total closed in 2026:       {total_closed_2026}")
print(f"  Currently open:             {total_open}")
print(f"  Open > {OPEN_THRESHOLD_DAYS} days:            {total_open_gt90}")
print("=" * 62)
print()
print("  BREAKDOWN BY LOCATION")
print("-" * 62)
print(df_location.to_string(index=False))
print()

# ── Export to Excel ───────────────────────────────────────────────────────────
output_path = os.path.join(SCRIPT_DIR, "CAPA_Metrics_Report_TaskDates.xlsx")
exec(open(os.path.join(SCRIPT_DIR, "_excel_export_taskdates.py"), encoding="utf-8").read())
