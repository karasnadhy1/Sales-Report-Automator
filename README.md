# 📊 Sales Report Automator

Turn a messy sales export into a clean, client-ready Excel report — in a browser, no spreadsheet wrangling required.

Upload any sales CSV → the app detects your columns → cleans the data → shows revenue, top categories, and monthly trends instantly → download a polished, formula-driven Excel report with charts.

## Why this exists

Small businesses export sales data full of the same problems every month: mixed date formats, inconsistent capitalization, stray currency symbols, duplicate rows from double-exports. Cleaning it by hand wastes hours. This tool automates that entire process and hands back a report that's actually presentable.

## Features

- **Works with any CSV** — auto-detects which column is the date, price, quantity, etc., and lets you confirm or fix the guess before processing
- **Transparent cleaning** — shows exactly how many rows were removed and why (duplicates, bad dates, bad prices), never silently drops data
- **Live Excel formulas** — the report's Summary sheet uses real `SUMIF`/`SUMIFS` formulas referencing the raw data, so it recalculates itself if you edit the data directly in Excel
- **Handles edge cases** — empty files, all-garbage data, all-return datasets, mismatched column mappings, and more
- **Try before you upload** — a one-click sample dataset for first-time visitors

## Quick start

```bash
pip install -r requirements.txt
streamlit run app.py
```
This opens the app in your browser at `http://localhost:8501`. No command-line skills needed after that point — everything else is clicking.

## Tech stack

- **Python** — pandas for data cleaning and analysis
- **openpyxl** — builds the Excel report with live formulas, conditional formatting, and charts
- **Streamlit** — the web interface

## How it works

```
app.py                  → the web UI (Streamlit)
scripts/pipeline.py     → all the actual logic: column detection, cleaning, report building
                           (kept separate from the UI so it can be tested without a browser)
```

1. **Column detection** — keyword-matches your file's column names against what's needed (date, price, quantity, etc.), then lets you confirm the guess
2. **Cleaning** — strips whitespace, standardizes casing, parses mixed date formats, converts currency-formatted prices to numbers, flags returns, removes duplicates
3. **Report generation** — builds a two-sheet Excel workbook (raw data + formula-driven summary with charts) entirely in memory

## Testing

The cleaning and report-building logic (`scripts/pipeline.py`) was tested independently of the UI: against the original column layout, against a synthetic file with completely different column names, and against edge cases including empty files, unparseable data, and datasets where every order is a return. All resulting Excel reports were verified with LibreOffice recalculation (0 formula errors).

---

## Development journal

This project was built incrementally over several sessions, each with its own written notes on what was learned. Kept here for anyone curious about the process, or for my own reference.

<details>
<summary><b>Day 1 — Data cleaning</b></summary>

`scripts/clean_data.py` takes a raw CSV and:
- Strips stray whitespace from text fields
- Standardizes inconsistent casing (`electronics` / `ELECTRONICS` → `Electronics`)
- Parses mixed date formats (`9/7/2026`, `2026-09-16`, `1-6-2026`) into one consistent type
- Strips currency symbols and converts prices to proper numbers
- Flags negative quantities as returns instead of treating them as errors
- Removes exact duplicate rows
- Reports exactly how many rows were dropped and why

**Result on the demo dataset**: 186 raw rows → 170 clean rows (6 duplicates, 10 rows with unrecoverable missing data).
</details>

<details>
<summary><b>Day 2 — Analysis</b></summary>

`scripts/analyze_data.py` computes total revenue (returns excluded), revenue by category/region/sales rep, top products, and month-over-month growth, saving everything to `summary_metrics.csv`.

**Result on the demo dataset**: $75,848.04 total revenue, Electronics ~91% of category revenue, West region leading, 8 returns worth $3,052.93.
</details>

<details>
<summary><b>Day 3 — Excel report</b></summary>

`scripts/build_report.py` builds a two-sheet Excel workbook: a Data sheet with the full clean dataset, and a Summary sheet with KPI cards, category/region breakdowns with color-scale conditional formatting, a monthly trend table, and two charts — all built from live `SUMIF`/`SUMIFS` formulas, not hardcoded numbers.

Verified with LibreOffice recalculation: 194 formulas, 0 errors.
</details>

<details>
<summary><b>Day 4 — The app</b></summary>

Wrapped the full pipeline in a Streamlit app: upload any CSV, confirm the auto-detected column mapping, see metrics instantly, download the report. Logic was split into `scripts/pipeline.py` so it could be tested independently of the UI — tested against both the original column layout and a synthetic file with completely different column names, both handled correctly with zero formula errors.

Later revised to add: a random (not fixed-seed) sample dataset generator, full edge-case handling (empty files, all-garbage data, duplicate column mappings, all-return datasets), and a custom "ledger/receipt" visual design with entrance animations and hover interactions.
</details>
