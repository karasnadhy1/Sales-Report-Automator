"""
Day 4 (redesigned): Same tested pipeline, new visual identity.

Design direction: a "ledger / receipt" theme — serif display type,
monospace tabular numbers (like a receipt printer), a dashed
perforation line on the KPI strip, and a rotated stamp badge as the
one deliberate flourish. Chosen because the subject IS a sales
report — leaning into that felt more honest than a generic dashboard
look.

Run with:  streamlit run app.py
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "scripts"))

import streamlit as st
import pandas as pd
from pipeline import (
    suggest_column_mapping, clean_dataframe, build_excel_report,
    generate_sample_csv, REQUIRED_FIELDS, OPTIONAL_FIELDS
)

st.set_page_config(page_title="Sales Report Automator", layout="wide")

# ============================================================
# Design system — CSS injected once, used throughout
# ============================================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,wght@0,600;1,600&family=IBM+Plex+Mono:wght@400;600;700&family=IBM+Plex+Sans:wght@400;500&display=swap');

:root {
  --paper: #F5F1E8;
  --card: #FFFFFF;
  --ink: #1B1F23;
  --muted: #6B6459;
  --green: #1F5C4F;
  --gold: #C98A2B;
  --red: #B23A3A;
  --line: #D8D2C2;
}

.stApp { background: var(--paper); font-family: 'IBM Plex Sans', sans-serif; }

/* Entrance animation for major sections */
@keyframes fadeSlideIn {
  from { opacity: 0; transform: translateY(10px); }
  to { opacity: 1; transform: translateY(0); }
}
.app-title, .app-subtitle, .eyebrow, .receipt, .section-title { animation: fadeSlideIn 0.6s ease-out; }
.receipt { animation-delay: 0.1s; animation-fill-mode: backwards; }

/* Bars grow in on load instead of appearing static */
@keyframes growBar { from { width: 0; } to { width: var(--w); } }
.bar-fill { animation: growBar 900ms cubic-bezier(0.22, 1, 0.36, 1) forwards; }

/* Header block */
.eyebrow {
  font-family: 'IBM Plex Mono', monospace;
  letter-spacing: 0.12em;
  font-size: 12px;
  color: var(--green);
  text-transform: uppercase;
  margin-bottom: 4px;
}
.app-title {
  font-family: 'Fraunces', Georgia, serif;
  font-weight: 600;
  font-style: italic;
  font-size: 40px;
  color: var(--ink);
  margin: 0 0 4px 0;
}
.app-subtitle {
  color: var(--muted);
  font-size: 15px;
  margin-bottom: 8px;
}
.stamp {
  border: 2px solid var(--green);
  color: var(--green);
  border-radius: 50%;
  width: 76px;
  height: 76px;
  display: flex;
  align-items: center;
  justify-content: center;
  text-align: center;
  font-family: 'IBM Plex Mono', monospace;
  font-size: 10px;
  font-weight: 700;
  letter-spacing: 0.04em;
  transform: rotate(-10deg);
  opacity: 0.85;
  float: right;
  margin-top: -8px;
  transition: transform 0.3s ease, opacity 0.3s ease;
}
.stamp:hover { transform: rotate(0deg) scale(1.06); opacity: 1; }

/* Receipt strip (KPI row) */
.receipt {
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 3px;
  padding: 26px 30px 30px 30px;
  margin: 18px 0 32px 0;
  transition: box-shadow 0.3s ease;
}
.receipt:hover { box-shadow: 0 6px 24px rgba(31, 92, 79, 0.08); }
.perf {
  height: 1px;
  background-image: repeating-linear-gradient(to right, var(--line) 0, var(--line) 6px, transparent 6px, transparent 12px);
}
.perf.top { margin: -26px -30px 22px -30px; }
.kpi-row { display: flex; justify-content: space-between; flex-wrap: wrap; gap: 16px; }
.kpi-label {
  font-family: 'IBM Plex Mono', monospace;
  font-size: 11px;
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: var(--muted);
  margin-bottom: 6px;
}
.kpi-value {
  font-family: 'IBM Plex Mono', monospace;
  font-size: 25px;
  font-weight: 700;
  color: var(--ink);
  transition: transform 0.25s ease;
  display: inline-block;
}
.kpi-value:hover { transform: scale(1.05); }
.kpi-value.positive { color: var(--green); }
.kpi-value.negative { color: var(--red); }

/* Section headers */
.section-title {
  font-family: 'Fraunces', Georgia, serif;
  font-weight: 600;
  font-size: 18px;
  color: var(--ink);
  margin: 6px 0 14px 0;
  padding-bottom: 8px;
  border-bottom: 1px solid var(--line);
}

/* Ledger table with inline bars */
.ledger-table { width: 100%; border-collapse: collapse; font-family: 'IBM Plex Mono', monospace; font-size: 13.5px; }
.ledger-table th {
  text-align: left;
  font-family: 'IBM Plex Sans', sans-serif;
  font-size: 11px;
  letter-spacing: 0.05em;
  text-transform: uppercase;
  color: var(--muted);
  border-bottom: 1px solid var(--ink);
  padding: 6px 4px;
}
.ledger-table tr { transition: background 0.2s ease; }
.ledger-table tbody tr:hover { background: #FBF8F2; }
.ledger-table td { padding: 7px 4px; border-bottom: 1px solid var(--line); color: var(--ink); }
.bar-cell { display: flex; align-items: center; gap: 8px; }
.bar-bg { flex: 1; background: #EFEAE0; height: 12px; border-radius: 2px; overflow: hidden; min-width: 60px; }
.bar-fill { height: 100%; background: var(--green); }
.bar-fill.low { background: var(--red); }

/* Buttons */
div[data-testid="stButton"] > button, div[data-testid="stDownloadButton"] > button {
  background: var(--green);
  color: white;
  border: none;
  border-radius: 3px;
  font-family: 'IBM Plex Sans', sans-serif;
  font-weight: 500;
  transition: background 0.2s ease, transform 0.15s ease, box-shadow 0.2s ease;
}
div[data-testid="stButton"] > button:hover, div[data-testid="stDownloadButton"] > button:hover {
  background: #17473D;
  color: white;
  transform: translateY(-1px);
  box-shadow: 0 4px 12px rgba(31, 92, 79, 0.25);
}
div[data-testid="stButton"] > button:active, div[data-testid="stDownloadButton"] > button:active {
  transform: translateY(0);
}

/* File uploader dropzone accent */
div[data-testid="stFileUploaderDropzone"] {
  border: 1.5px dashed var(--green) !important;
  background: #FBFAF6 !important;
  transition: border-color 0.2s ease, background 0.2s ease;
}
div[data-testid="stFileUploaderDropzone"]:hover {
  border-color: #17473D !important;
  background: #F5F1E8 !important;
}
</style>
""", unsafe_allow_html=True)


def ledger_table_html(df_grouped: pd.Series, label: str) -> str:
    """Renders a Series (index=label, values=revenue) as a ledger-style
    HTML table with inline animated bars, matching the receipt theme."""
    max_val = df_grouped.max() if len(df_grouped) else 1
    rows = ""
    for name, value in df_grouped.sort_values(ascending=False).items():
        pct = max(int((value / max_val) * 100), 2) if max_val > 0 else 2
        bar_class = "bar-fill" if pct > 25 else "bar-fill low"
        rows += (f"<tr><td>{name}</td><td>${value:,.2f}</td>"
                 f"<td class='bar-cell'><div class='bar-bg'>"
                 f"<div class='{bar_class}' style='--w:{pct}%'></div></div></td></tr>")
    return (f"<table class='ledger-table'><tr><th>{label}</th><th>Revenue</th><th></th></tr>"
            f"{rows}</table>")


# ============================================================
# Header
# ============================================================
st.markdown("""
<div class="stamp">AUTO<br>MATED</div>
<div class="eyebrow">Sales &middot; Ledger &middot; Automation</div>
<div class="app-title">Sales Report Automator</div>
<div class="app-subtitle">Upload a messy sales export. Get a clean, formatted Excel report back.</div>
""", unsafe_allow_html=True)

# --- Entry point: upload a file, or try the sample data ---
col_upload, col_sample = st.columns([3, 1])
with col_upload:
    uploaded_file = st.file_uploader("Upload your sales CSV", type=["csv"])
with col_sample:
    st.write("")
    st.write("")
    use_sample = st.button("🎲 Try with sample data")

if use_sample:
    st.session_state["raw_df"] = generate_sample_csv()
    st.session_state["source"] = "sample"
    for key in ("clean_df", "diagnostics"):
        st.session_state.pop(key, None)
elif uploaded_file is not None:
    try:
        new_raw_df = pd.read_csv(uploaded_file, dtype=str)
    except pd.errors.EmptyDataError:
        st.error("This file is empty. Please upload a CSV with at least a header row and one data row.")
        st.stop()
    except Exception as e:
        st.error(f"Couldn't read this file as a CSV: {e}")
        st.stop()

    if new_raw_df.shape[0] == 0:
        st.error("This CSV has headers but no data rows. Please upload a file with actual sales data.")
        st.stop()
    if new_raw_df.shape[1] < 3:
        st.warning("This file only has a few columns — double check it's the right file.")

    st.session_state["raw_df"] = new_raw_df
    st.session_state["source"] = "upload"
    for key in ("clean_df", "diagnostics"):
        st.session_state.pop(key, None)

if "raw_df" not in st.session_state:
    st.info("Upload a CSV, or click **Try with sample data** to see how it works first.")
    st.stop()

raw_df = st.session_state["raw_df"]
if st.session_state.get("source") == "sample":
    st.success("Using sample data — upload your own file above anytime to replace it.")

st.markdown("<div class='section-title'>Raw data preview</div>", unsafe_allow_html=True)
st.dataframe(raw_df.head(10), use_container_width=True)

# --- Column mapping step ---
st.markdown("<div class='section-title'>Confirm what each column means</div>", unsafe_allow_html=True)
st.caption("We guessed based on your column names — fix anything that looks wrong.")

suggested = suggest_column_mapping(raw_df)
column_options = ["-- none --"] + list(raw_df.columns)
mapping = {}

col1, col2 = st.columns(2)
with col1:
    st.markdown("**Required**")
    for field in REQUIRED_FIELDS:
        default = suggested.get(field)
        default_idx = column_options.index(default) if default in column_options else 0
        choice = st.selectbox(field.replace("_", " ").title(), column_options, index=default_idx, key=f"map_{field}")
        mapping[field] = None if choice == "-- none --" else choice

with col2:
    st.markdown("**Optional** (skip any that don't apply)")
    for field in OPTIONAL_FIELDS:
        default = suggested.get(field)
        default_idx = column_options.index(default) if default in column_options else 0
        choice = st.selectbox(field.replace("_", " ").title(), column_options, index=default_idx, key=f"map_{field}")
        mapping[field] = None if choice == "-- none --" else choice

missing_required = [f for f in REQUIRED_FIELDS if mapping.get(f) is None]
if missing_required:
    st.warning(f"Please map required fields: {', '.join(missing_required)}")
    st.stop()

chosen_columns = [v for v in mapping.values() if v is not None]
duplicates = {c for c in chosen_columns if chosen_columns.count(c) > 1}
if duplicates:
    st.error(f"You've mapped the same column to more than one field: {', '.join(duplicates)}. "
              "Each column should represent only one thing — please fix before continuing.")
    st.stop()

if st.button("Clean & Analyze", type="primary"):
    with st.spinner("Cleaning and analyzing your data..."):
        try:
            clean_df, diagnostics = clean_dataframe(raw_df, mapping)
        except Exception as e:
            st.error(f"Something went wrong while cleaning: {e}")
            st.stop()
    st.session_state["clean_df"] = clean_df
    st.session_state["diagnostics"] = diagnostics
    st.session_state["just_analyzed"] = True

if "clean_df" in st.session_state:
    clean_df = st.session_state["clean_df"]
    diag = st.session_state["diagnostics"]

    removed_total = diag["starting_rows"] - diag["ending_rows"]
    with st.expander(f"Cleaning summary — {diag['ending_rows']} of {diag['starting_rows']} rows kept", expanded=(removed_total > 0)):
        st.write(f"- {diag['duplicates_removed']} exact duplicate rows removed")
        st.write(f"- {diag['missing_date']} rows had an unreadable/missing date")
        st.write(f"- {diag['missing_price']} rows had a missing/invalid price")
        st.write(f"- {diag['missing_quantity']} rows had a missing/invalid quantity")
        st.write(f"- {diag['unrecoverable_removed']} rows dropped total (missing date, price, or quantity)")
        if diag["starting_rows"] > 0 and removed_total / diag["starting_rows"] > 0.3:
            st.warning("More than 30% of rows were removed. This often means a column was mapped "
                       "to the wrong field — double check your mapping above.")

    if len(clean_df) == 0:
        st.error("No rows survived cleaning. This usually means the date, price, or quantity "
                  "column is mapped incorrectly. Check the mapping above and try again.")
        st.stop()

    if st.session_state.get("just_analyzed"):
        st.markdown(
            "<div style=\"font-family:'IBM Plex Mono',monospace; color:#1F5C4F; "
            "font-size:13px; letter-spacing:0.04em; margin:-4px 0 18px 0; "
            "animation: fadeSlideIn 0.5s ease-out both;\">"
            "&#10003; Report ready</div>",
            unsafe_allow_html=True
        )
        st.session_state["just_analyzed"] = False

    sales = clean_df[~clean_df["is_return"]]
    avg_value_html = f"${sales['revenue'].mean():,.2f}" if len(sales) > 0 else "N/A"

    # --- Receipt-style KPI strip ---
    st.markdown(f"""
    <div class="receipt">
      <div class="perf top"></div>
      <div class="kpi-row">
        <div><div class="kpi-label">Total Revenue</div><div class="kpi-value positive">${sales['revenue'].sum():,.2f}</div></div>
        <div><div class="kpi-label">Total Orders</div><div class="kpi-value">{len(clean_df):,}</div></div>
        <div><div class="kpi-label">Avg Order Value</div><div class="kpi-value">{avg_value_html}</div></div>
        <div><div class="kpi-label">Returns</div><div class="kpi-value negative">${clean_df[clean_df['is_return']]['revenue'].sum():,.2f}</div></div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    if len(sales) == 0:
        st.warning("Every remaining order is a return — revenue-based charts and the "
                   "Excel report will be empty until there's at least one non-return sale.")
        st.stop()

    chart_col1, chart_col2 = st.columns(2)
    if "category" in clean_df.columns:
        with chart_col1:
            st.markdown("<div class='section-title'>Revenue by Category</div>", unsafe_allow_html=True)
            st.markdown(ledger_table_html(sales.groupby("category")["revenue"].sum(), "Category"), unsafe_allow_html=True)
    if "region" in clean_df.columns:
        with chart_col2:
            st.markdown("<div class='section-title'>Revenue by Region</div>", unsafe_allow_html=True)
            st.markdown(ledger_table_html(sales.groupby("region")["revenue"].sum(), "Region"), unsafe_allow_html=True)

    st.markdown("<div class='section-title'>Monthly Revenue Trend</div>", unsafe_allow_html=True)
    st.line_chart(sales.groupby("month")["revenue"].sum(), color="#1F5C4F")

    st.markdown("<div class='section-title'>Download your report</div>", unsafe_allow_html=True)
    excel_bytes = build_excel_report(clean_df)
    st.download_button(
        label="⬇️ Download Excel Report",
        data=excel_bytes,
        file_name="sales_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
