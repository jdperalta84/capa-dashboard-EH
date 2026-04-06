"""
CAPA KPI Metrics Calculator
============================

LOGIC NOTES (for explanation to others)
-----------------------------------------
HOW WE DETERMINE IF A CAPA IS CLOSED AND WHEN:

  A CAPA is CLOSED if and only if "Date closed" on the main Capas sheet
  has a date. No exceptions. The Status column is not used.

  A CAPA is OPEN if "Date closed" on the main Capas sheet is blank.

  "Closed in 2025" / "Closed in 2026" is determined by the year of
  "Date closed" on the main Capas sheet.

HOW THE AVERAGE DAYS TO CLOSE IS CALCULATED:

  All CAPAs with a Date closed in 2025 are included, regardless of
  when they were originally opened.

  Formula: Date closed (Capas main tab) − Date of notification,
  averaged across all qualifying CAPAs.

HOW EACH METRIC IS CALCULATED:

  - Avg days to close (2025):
      Only includes CAPAs where BOTH Date of notification AND resolved
      closed date fall in calendar year 2025.
      Formula: closed date − date of notification, averaged across those CAPAs.

  - Closed in 2025 / Closed in 2026:
      Resolved closed date falls in that calendar year.

  - Currently open:
      No resolved closed date (per logic above).

  - Open > 90 days:
      Open CAPAs where today − Date of notification > 90 days.
      "Date of notification" is when the CAPA was originally created.

SOURCE FILES:
  Reads all files matching: export_CAPA *.xls in the same folder as this script.
  Sheet used from each file: "Capas" (main record) + "Taken" (action tasks).
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

    # ── Resolve closed date per CAPA ─────────────────────────────────────────
    # Group tasks by CAPA number
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
            # no completed tasks with dates → fall back to Capas sheet
        return capas_date

    capas["Effective closed date"] = capas.apply(resolve_closed_date, axis=1)
    capas_frames.append(capas)

all_capas = pd.concat(capas_frames, ignore_index=True)

# ── Derived flags ─────────────────────────────────────────────────────────────
# Open/closed determined solely by Date closed on the Capas main tab
is_closed = all_capas["Date closed"].notna()
is_open   = ~is_closed

closed_2025 = is_closed & (all_capas["Date closed"].dt.year == 2025)
closed_2026 = is_closed & (all_capas["Date closed"].dt.year == 2026)

days_open = (pd.Timestamp(TODAY) - all_capas["Date of notification"]).dt.days
open_gt90 = is_open & (days_open > OPEN_THRESHOLD_DAYS)

# ── Avg days to close ─────────────────────────────────────────────────────────
def avg_days(mask):
    df = all_capas[mask].copy()
    df["days_to_close"] = (df["Date closed"] - df["Date of notification"]).dt.days
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
        "Location":                         loc,
        "Avg Days to Close (2025)":         round(loc_avg_2025, 1) if pd.notna(loc_avg_2025) else "N/A",
        "Closed 2025":                      int((mask & closed_2025).sum()),
        "Avg Days to Close (2026)":         round(loc_avg_2026, 1) if pd.notna(loc_avg_2026) else "N/A",
        "Closed 2026":                      int((mask & closed_2026).sum()),
        "Open":                             int((mask & is_open).sum()),
        f"Open > {OPEN_THRESHOLD_DAYS} days": int((mask & open_gt90).sum()),
    })

df_location = pd.DataFrame(location_rows)

# ── Print summary ─────────────────────────────────────────────────────────────
print("=" * 62)
print("  CAPA KPI METRICS SUMMARY")
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
exec(open(os.path.join(SCRIPT_DIR, "_excel_export.py"), encoding="utf-8").read())
