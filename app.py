"""
Excel Data Cleaner — Streamlit UI
Run with: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime
from clean_excel import clean_dataframe, write_cleaned_excel

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Excel Cleaner",
    page_icon="🧹",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Styles (Light / Rich Theme) ───────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        color: #1e293b;
    }

    /* Hide default Streamlit chrome */
    #MainMenu, footer, header { visibility: hidden; }
    .block-container { padding: 2rem 3rem 4rem 3rem; max-width: 1100px; }

    /* Hero */
    .hero {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 2.5rem 3rem;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05), 0 8px 10px -6px rgba(0, 0, 0, 0.01);
    }
    .hero h1 {
        color: #0f172a;
        font-size: 2rem;
        font-weight: 700;
        margin: 0 0 0.4rem 0;
        letter-spacing: -0.02em;
    }
    .hero p {
        color: #64748b;
        font-size: 0.95rem;
        margin: 0;
    }
    .hero .badge {
        display: inline-block;
        background: #eff6ff;
        color: #2563eb;
        font-size: 0.7rem;
        font-weight: 600;
        padding: 2px 8px;
        border-radius: 999px;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        margin-bottom: 0.8rem;
    }

    /* Upload zone */
    [data-testid="stFileUploader"] {
        border: 2px dashed #cbd5e1 !important;
        border-radius: 12px !important;
        background: #f8fafc !important;
        padding: 1rem !important;
        transition: border-color 0.2s;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #2563eb !important;
    }

    /* Metric cards */
    .metric-row { display: flex; gap: 1rem; margin: 1.5rem 0; flex-wrap: wrap; }
    .metric-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.1rem 1.4rem;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
        flex: 1;
        min-width: 130px;
    }
    .metric-card .label {
        color: #64748b;
        font-size: 0.72rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-bottom: 0.3rem;
    }
    .metric-card .value {
        color: #0f172a;
        font-size: 1.6rem;
        font-weight: 700;
        font-family: 'JetBrains Mono', monospace;
        line-height: 1;
    }
    .metric-card .delta {
        font-size: 0.75rem;
        margin-top: 0.25rem;
    }
    .delta-good { color: #16a34a; }
    .delta-bad  { color: #dc2626; }
    .delta-neu  { color: #64748b; }

    /* Section headers */
    .section-label {
        color: #64748b;
        font-size: 0.72rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.1em;
        margin: 2rem 0 0.75rem 0;
        padding-bottom: 0.4rem;
        border-bottom: 1px solid #e2e8f0;
    }

    /* Tag pills */
    .pill {
        display: inline-block;
        background: #f1f5f9;
        color: #475569;
        font-size: 0.72rem;
        font-weight: 500;
        padding: 3px 10px;
        border-radius: 999px;
        margin: 2px;
        font-family: 'JetBrains Mono', monospace;
    }
    .pill-warn {
        background: #fff7ed;
        color: #ea580c;
    }
    .pill-ok {
        background: #f0fdf4;
        color: #16a34a;
    }

    /* Download button */
    .stDownloadButton > button {
        background: #2563eb !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        padding: 0.6rem 1.4rem !important;
        font-size: 0.9rem !important;
        width: 100%;
        transition: background 0.15s;
        box-shadow: 0 4px 6px -1px rgba(37, 99, 235, 0.2);
    }
    .stDownloadButton > button:hover {
        background: #1d4ed8 !important;
    }

    /* Dataframe */
    [data-testid="stDataFrame"] {
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        overflow: hidden;
    }

    /* Rename table */
    .rename-row {
        display: flex;
        align-items: center;
        gap: 0.5rem;
        padding: 0.3rem 0;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.8rem;
        color: #64748b;
    }
    .rename-row .arrow { color: #2563eb; }
    .rename-row .new-name { color: #16a34a; }
</style>
""", unsafe_allow_html=True)

# ── Hero ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="badge">Data Quality Tool</div>
    <h1>🧹 Excel Cleaner</h1>
    <p>Upload a messy .xlsx or .csv file — duplicates, missing values, bad headers, and outliers handled automatically.</p>
</div>
""", unsafe_allow_html=True)

# ── Upload ────────────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Drop your file here or click to browse",
    type=["xlsx", "xls", "csv"],
    label_visibility="collapsed",
)

if not uploaded:
    st.markdown("""
    <div style="text-align:center; color:#64748b; padding: 1.5rem 0; font-size:0.85rem;">
        Supports <strong style="color:#475569">.xlsx</strong>, <strong style="color:#475569">.xls</strong>, and <strong style="color:#475569">.csv</strong>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ── Load ──────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_file(file_bytes, name):
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes), dtype=str)
    else:
        return pd.read_excel(io.BytesIO(file_bytes), dtype=str)

with st.spinner("Reading file…"):
    file_bytes = uploaded.read()
    try:
        df_raw = load_file(file_bytes, uploaded.name)
    except Exception as e:
        st.error(f"Could not read file: {e}")
        st.stop()

# ── Preview raw ───────────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Raw Preview</div>', unsafe_allow_html=True)
st.dataframe(df_raw.head(50), use_container_width=True, height=220)

# ── Clean ─────────────────────────────────────────────────────────────────────
with st.spinner("Cleaning…"):
    df_clean, report, outlier_mask = clean_dataframe(df_raw.copy())

orig = report["original_shape"]
final = report["final_shape"]

# ── Metrics ───────────────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Cleaning Summary</div>', unsafe_allow_html=True)

rows_removed = orig[0] - final[0]
cols_removed  = orig[1] - final[1]
total_outliers = sum(report.get("outliers_flagged", {}).values())

st.markdown(f"""
<div class="metric-row">
    <div class="metric-card">
        <div class="label">Rows Before</div>
        <div class="value">{orig[0]:,}</div>
    </div>
    <div class="metric-card">
        <div class="label">Rows After</div>
        <div class="value">{final[0]:,}</div>
        <div class="delta {'delta-good' if rows_removed > 0 else 'delta-neu'}">
            {"−" + str(rows_removed) + " removed" if rows_removed > 0 else "No rows removed"}
        </div>
    </div>
    <div class="metric-card">
        <div class="label">Duplicates</div>
        <div class="value">{report["duplicates_removed"]}</div>
        <div class="delta delta-good">removed</div>
    </div>
    <div class="metric-card">
        <div class="label">Empty Rows</div>
        <div class="value">{report["empty_rows_removed"]}</div>
        <div class="delta delta-good">removed</div>
    </div>
    <div class="metric-card">
        <div class="label">Empty Cols</div>
        <div class="value">{report["empty_cols_removed"]}</div>
        <div class="delta delta-good">removed</div>
    </div>
    <div class="metric-card">
        <div class="label">Outliers</div>
        <div class="value">{total_outliers}</div>
        <div class="delta {'delta-bad' if total_outliers > 0 else 'delta-neu'}">
            {"flagged" if total_outliers > 0 else "none"}
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Detail pills ──────────────────────────────────────────────────────────────
col1, col2 = st.columns(2)

with col1:
    if report.get("date_cols_detected"):
        st.markdown("**📅 Date columns auto-detected**")
        pills = " ".join(f'<span class="pill pill-ok">{c}</span>' for c in report["date_cols_detected"])
        st.markdown(pills, unsafe_allow_html=True)

    if report.get("numeric_cols_converted"):
        st.markdown("**🔢 Numeric columns converted from text**")
        pills = " ".join(f'<span class="pill">{c}</span>' for c in report["numeric_cols_converted"])
        st.markdown(pills, unsafe_allow_html=True)

    if report.get("fixed_counts"):
        st.markdown("**🛠️ Cells fixed per column**")
        pills = " ".join(
            f'<span class="pill pill-ok">{c}: {n}</span>'
            for c, n in report["fixed_counts"].items() if n > 0
        )
        st.markdown(pills, unsafe_allow_html=True)

with col2:
    if report.get("outliers_flagged"):
        st.markdown("**⚠️ Outlier counts by column**")
        pills = " ".join(
            f'<span class="pill pill-warn">{c}: {n}</span>'
            for c, n in report["outliers_flagged"].items()
        )
        st.markdown(pills, unsafe_allow_html=True)

    if report.get("columns_renamed"):
        st.markdown("**✏️ Column headers normalized**")
        for old, new in report["columns_renamed"].items():
            st.markdown(
                f'<div class="rename-row">'
                f'<span class="pill">{old}</span>'
                f'<span class="arrow">→</span>'
                f'<span class="new-name pill pill-ok">{new}</span>'
                f'</div>',
                unsafe_allow_html=True
            )

# ── Cleaned preview ───────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Cleaned Data Preview</div>', unsafe_allow_html=True)
st.dataframe(df_clean.head(50), use_container_width=True, height=260)

# ── Outliers tab (if any) ─────────────────────────────────────────────────────
if outlier_mask.any(axis=None):
    with st.expander(f"🔎 View {int(outlier_mask.any(axis=1).sum())} outlier rows"):
        flagged = df_clean[outlier_mask.any(axis=1)]
        st.dataframe(flagged, use_container_width=True, height=200)

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown('<div class="section-label">Download</div>', unsafe_allow_html=True)

output_buf = io.BytesIO()

# write_cleaned_excel needs a file path; use a temp file approach
import tempfile, os
with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
    tmp_path = tmp.name

write_cleaned_excel(df_clean, report, outlier_mask, tmp_path)
with open(tmp_path, "rb") as f:
    cleaned_bytes = f.read()
os.unlink(tmp_path)

base_name = os.path.splitext(uploaded.name)[0]
out_name  = f"{base_name}_cleaned.xlsx"

dl_col, info_col = st.columns([1, 2])
with dl_col:
    st.download_button(
        label="⬇ Download Cleaned File",
        data=cleaned_bytes,
        file_name=out_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with info_col:
    st.markdown(f"""
    <div style="color:#64748b; font-size:0.82rem; padding-top:0.6rem;">
        Includes <strong style="color:#475569">3 sheets</strong>: 
        Cleaned Data · Outliers · Cleaning Report<br>
        <span style="font-family:'JetBrains Mono',monospace;">{out_name}</span>
    </div>
    """, unsafe_allow_html=True)

st.markdown(f"""
<div style="color:#94a3b8; font-size:0.72rem; text-align:center; margin-top:3rem;">
    Cleaned at {datetime.now().strftime("%H:%M · %d %b %Y")}
</div>
""", unsafe_allow_html=True)
