# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Common development commands

- **Install dependencies**
  ```bash
  pip install -r requirements.txt
  ```
  Installs Streamlit, pandas, openpyxl, Altair and other runtime packages.

- **Run the dashboard (development mode)**
  ```bash
  streamlit run streamlit_app.py
  ```
  The app will start a local web server (by default at `http://localhost:8501`).

- **Run lint / static analysis** (if you install the optional dev extras)
  ```bash
  pip install flake8  # optional
  flake8 .
  ```
  There is no pre‑configured lint configuration; the default rules are sufficient.

- **Run a single test** – the repository currently has no test suite. When tests are added, use the conventional pattern:
  ```bash
  pytest path/to/test_file.py
  ```

- **Re‑install after changes to requirements**
  ```bash
  pip install -r requirements.txt --upgrade
  ```

## High‑level architecture & structure

- **`streamlit_app.py`** – the sole entry point for the Streamlit web‑app.
  - **UI layout** – top‑level title, sidebar for settings (calculation method, location filter), KPI cards, location table, detail tabs, and download sections.
  - **Data ingestion** – `load_data(uploaded_files)` reads the user‑uploaded `export_CAPA *.xls` files, extracts the location from the filename, parses the `Capas` and `Taken` sheets, and resolves an *effective closed date* using task completion dates when available. **If any task in the Taken sheet is not marked `Completed = Yes`, the CAPA is considered open; only when all tasks are completed does the CAPA close on the latest task completion date.** A progress bar (`st.progress`) provides feedback during multi‑file uploads.
  - **Metric calculation** – `compute_metrics(all_capas, method)` computes KPI aggregates (average days to close, totals per year, open counts, open > threshold) and assembles per‑location and detailed DataFrames.
  - **Excel report generation** – `build_excel_report(metrics, method)` creates a fully styled workbook (Dashboard, By Location, Summary, Logic Notes, and a detail sheet per KPI) using `openpyxl`.
  - **Trend visualization** – a small Altair line chart (average days to close per month) is rendered after the KPI cards to give a quick visual trend.

- **Data flow**
  1. User uploads one or more `export_CAPA *.xls` files.
  2. `load_data` parses each file, builds a unified DataFrame (`all_capas`).
  3. Sidebar selections (`method`, `selected_locations`) filter the DataFrame.
  4. `compute_metrics` produces the KPI dictionary.
  5. UI components consume this dictionary for cards, tables, tabs, and the download buttons.
  6. The Excel builder receives the same metric dict to produce downloadable reports.

- **Extensibility points**
  - **Threshold** – `OPEN_THRESHOLD_DAYS` is defined near the top; it can be exposed as a sidebar numeric input if needed.
  - **Additional calculation methods** – the `method` parameter can be extended; the UI already supports a radio selector.
  - **Custom visualizations** – Altair is imported; new charts can be added after the KPI cards.
  - **Testing** – the pure‑function nature of `load_data`, `compute_metrics`, and `build_excel_report` makes them straightforward to unit‑test with synthetic DataFrames.

## Useful references in this repo

- `requirements.txt` – pinpoints the exact third‑party libraries required to run the app.
- No existing test suite, lint configuration, or CI scripts are present; they can be added later as needed.

---
*Generated for Claude Code to accelerate onboarding and future automated assistance.*