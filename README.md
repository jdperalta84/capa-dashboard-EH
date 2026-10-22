# CAPA KPI Dashboard

A Streamlit dashboard for reviewing corrective and preventive action (CAPA) metrics across locations. Upload Excel exports to view closure times, open records, location summaries, and downloadable Excel reports.

## Run locally

From the repository folder, install the dependencies and start the app:

```bash
python -m pip install -r requirements.txt
python -m streamlit run streamlit_app.py
```

Open the local address printed by Streamlit in your browser. The web app entry point is `streamlit_app.py`.

## Use the dashboard

1. Upload one or more CAPA exports with an `.xls` or `.xlsx` extension.
2. Choose **Official (Date Closed)** or **Task Date Priority** in the sidebar.
3. Select the locations to include. All loaded locations are selected initially.
4. Review KPI cards, the 2026 initiation cohort, the monthly trend, and record details.
5. Download the **Official Report**, **Task Dates Report**, or both. Each download uses the selected locations and its named calculation method.

## Supported export formats

| Format | Expected structure | Location |
| --- | --- | --- |
| Legacy Excel exports | `Capas` and `Taken` sheets | Derived from the filename |
| CAPA/Tasks exports | `CAPA` and `Tasks` sheets, matched without regard to capitalization | Derived from the filename |
| Newer sheet-based exports | Sheet names containing `CAR` or `PTO`; English headers on row 1, translation row on row 2, records from row 3 | Column A in each record |

For filename-based locations, use names such as `export_CAPA Thailand.xls` or `capas UK.xlsx`.

The archive formats use fields such as `Number`, `Type`, `Date of notification`, and `Date closed`. CAPA/Tasks exports use `Closure date`, which the app renames internally. Task records use `Date of completion` and either `Completed` (legacy) or `Status` (CAPA/Tasks).

Newer CAR/PTO imports depend on fixed column positions: CAR notification/closure dates are in K/M; PTO notification/closure dates are in I/L. PAR and CAF sheets are not loaded by this importer. Keep the source export layout intact.

## Included CAPA types

For the legacy and CAPA/Tasks formats, the app trims surrounding whitespace and matches Type values without regard to capitalization. The inclusion list accepts:

- `Client Complaint`, `Customer Complaint`
- `Incident`
- `ISO 17025 Audit`
- `Major Non-Conformity`, `Major Nonconformity`
- `Minor Non Conformity`, `Minor Nonconformity`
- `Non-Conformity`, `Nonconformity`
- `Non-compliance`, `Non-conformance`
- `PT Outlier`, `PT Outlier (insignificant risk)`
- `Proficiency Testing Outlier`
- `PTP Outlier full`, `PTP outlier short (insignificant risk)`

Other Type values are excluded, including observations, opportunities for improvement, recommendations, points of attention/improvement, risks, and site complaints. Punctuation and wording variants need explicit entries in `CAPA_TYPE_INCLUDE` in `streamlit_app.py`.

If an archive sheet has no Type column, this filter is bypassed. Newer sheet-based imports assign CAR/PTO types and use their separate sheet-selection rules.

## Calculation methods

| Method | Closure date used |
| --- | --- |
| **Official (Date Closed)** — default | Main record's Date closed |
| **Task Date Priority** | Latest task completion date when all associated tasks are completed and at least one completion date exists; otherwise the main record's Date closed |

Legacy tasks count as completed when `Completed` is `Yes`. CAPA/Tasks records use `Status = Completed`. Both comparisons ignore capitalization and surrounding whitespace.

A valid closure date under the selected method means closed. A missing or invalid closure date means open. For newer CAR/PTO imports, both methods use the imported closure date.

## Metrics

The dashboard currently reports calendar years **2025** and **2026**, plus a cohort of CAPAs initiated in **2026**. These years are fixed in the current code.

- **Total closed in a year:** records whose selected closure date falls in that year, regardless of notification year.
- **Average days to close:** closure date minus notification date, averaged over records closed in the relevant year.
- **Currently open:** records with no valid selected closure date, regardless of notification year.
- **Open > 90 days:** open records where the report date minus notification date is strictly greater than 90 days.
- **Initiated in 2026:** records whose notification date falls in 2026, with completed/open counts, aged-open counts, and average completion time for closed records.

Missing notification dates are omitted from duration averages and cannot qualify for the aged-open count. They do not remove a record from closed/open totals. Negative durations are currently retained, so source date errors can affect averages.

The report date comes from the application's runtime date.

### Trend chart

The monthly close-time chart always uses the main record's Date closed, including when Task Date Priority is selected. Its averages can therefore differ from the KPI cards under that method.

## Excel reports

Each report contains Dashboard, By Location, Summary, Logic Notes, and KPI detail sheets. Downloads use the same selected location population as the dashboard.

## Repository structure

| File | Purpose |
| --- | --- |
| `streamlit_app.py` | Web app, imports, KPI calculations, charts, and Excel report generation |
| `requirements.txt` | Runtime dependencies |
| `capa_metrics.py`, `capa_metrics_taskdates.py` | Separate command-line calculators for local exports |
| `_excel_export.py`, `_excel_export_taskdates.py` | Export helpers for the command-line calculators |
| `CLAUDE.md` | Development notes |

The command-line calculators have separate logic and are not the web app entry point. Dashboard type rules described above apply to `streamlit_app.py`.

## Troubleshooting

- **Workbook fails to load:** check the expected sheets, headers, and export layout. The app tries both legacy XLS and XLSX readers.
- **A type is missing:** compare its spelling and punctuation with the inclusion list and check the selected locations.
- **Counts differ between methods:** compare the main closure date with task completion dates.
- **An average looks unexpected:** inspect missing notification dates, negative durations, and the selected closure method.

There is currently no configured automated test suite or CI workflow in this repository. For calculation or import changes, verify representative exports in both methods and check the downloaded reports.
