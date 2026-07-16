"""
AI Excel Data Quality Assistant — Premium Streamlit UI
"""

import streamlit as st
import pandas as pd
import numpy as np
import io
import os
import tempfile
import plotly.graph_objects as go
import plotly.express as px
import base64
from datetime import datetime
from clean_excel import clean_dataframe, generate_analytics, write_cleaned_excel

# ── Page Config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="AI Data Quality Assistant",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Premium CSS Styling ───────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
        color: #0f172a;
    }

    /* Hide Streamlit Branding */
    #MainMenu, footer, header { visibility: hidden; }
    .block-container { padding: 3rem 4rem 4rem 4rem; max-width: 1200px; margin: 0 auto; }

    /* Hero Section */
    .hero {
        text-align: center;
        padding: 3rem 0 2rem 0;
    }
    .hero h1 {
        font-size: 3rem;
        font-weight: 800;
        letter-spacing: -0.03em;
        background: linear-gradient(to right, #2563EB, #7C3AED);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 1rem;
        line-height: 1.1;
    }
    .hero p {
        font-size: 1.25rem;
        color: #64748b;
        font-weight: 400;
        max-width: 700px;
        margin: 0 auto;
    }

    /* Cards */
    .card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 24px;
        box-shadow: 0 4px 20px -2px rgba(0,0,0,0.04);
        transition: all 0.3s ease;
        height: 100%;
    }
    .card:hover {
        box-shadow: 0 10px 30px -5px rgba(0,0,0,0.08);
        transform: translateY(-2px);
        border-color: #cbd5e1;
    }
    
    .kpi-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 20px;
        text-align: center;
        box-shadow: 0 4px 12px rgba(0,0,0,0.03);
    }
    .kpi-card .label {
        font-size: 0.8rem;
        font-weight: 600;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        margin-bottom: 8px;
    }
    .kpi-card .value {
        font-size: 2.5rem;
        font-weight: 700;
        font-family: 'JetBrains Mono', monospace;
        line-height: 1;
    }
    .kpi-blue .value { color: #2563EB; }
    .kpi-red .value { color: #DC2626; }
    .kpi-green .value { color: #16A34A; }
    .kpi-purple .value { color: #7C3AED; }

    /* Section Headers */
    .section-header {
        font-size: 1.5rem;
        font-weight: 700;
        color: #0f172a;
        margin-top: 3rem;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: center;
    }
    .section-header::before {
        content: '';
        display: inline-block;
        width: 6px;
        height: 24px;
        background: #2563EB;
        border-radius: 4px;
        margin-right: 12px;
    }

    /* Upload Area */
    [data-testid="stFileUploader"] {
        background: #f8fafc;
        border: 2px dashed #cbd5e1;
        border-radius: 12px;
        padding: 20px;
    }
    [data-testid="stFileUploader"]:hover {
        border-color: #2563EB;
        background: #eff6ff;
    }

    /* Buttons */
    .stButton > button {
        background: #2563EB;
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.5rem;
        font-weight: 600;
        width: 100%;
        transition: all 0.2s;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.2);
    }
    .stButton > button:hover {
        background: #1D4ED8;
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.3);
    }
    .stDownloadButton > button {
        background: #ffffff;
        color: #2563EB;
        border: 1px solid #2563EB;
        box-shadow: none;
    }
    .stDownloadButton > button:hover {
        background: #eff6ff;
        border-color: #1D4ED8;
    }

    /* Chat */
    .chat-msg {
        padding: 12px 16px;
        border-radius: 12px;
        margin-bottom: 10px;
        max-width: 80%;
        font-size: 0.9rem;
    }
    .chat-user {
        background: #2563EB;
        color: white;
        margin-left: auto;
        border-bottom-right-radius: 2px;
    }
    .chat-ai {
        background: #f1f5f9;
        color: #0f172a;
        margin-right: auto;
        border-bottom-left-radius: 2px;
    }
    
    /* Tables */
    .dataframe {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.85rem;
    }
    .dataframe th {
        background: #f8fafc;
        padding: 12px;
        text-align: left;
        font-weight: 600;
        border-bottom: 2px solid #e2e8f0;
        color: #475569;
    }
    .dataframe td {
        padding: 10px 12px;
        border-bottom: 1px solid #f1f5f9;
        color: #334155;
    }
    
    /* Badges */
    .badge {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
    }
    .badge-red { background: #FEE2E2; color: #DC2626; }
    .badge-green { background: #D1FAE5; color: #059669; }
    .badge-blue { background: #DBEAFE; color: #2563EB; }
</style>
""", unsafe_allow_html=True)

# ── Helper Functions ──────────────────────────────────────────────────────────

def create_gauge_chart(score):
    fig = go.Figure(go.Indicator(
        mode = "gauge+number",
        value = score,
        domain = {'x': [0, 1], 'y': [0, 1]},
        number = {'font': {'size': 40, 'color': '#2563EB'}},
        gauge = {
            'axis': {'range': [0, 100], 'tickwidth': 1, 'tickcolor': "#e2e8f0"},
            'bar': {'color': "#2563EB"},
            'bgcolor': "white",
            'borderwidth': 2,
            'bordercolor': "#e2e8f0",
            'steps': [
                {'range': [0, 50], 'color': '#FEE2E2'},
                {'range': [50, 80], 'color': '#FEF3C7'},
                {'range': [80, 100], 'color': '#D1FAE5'}
            ],
            'threshold': {
                'line': {'color': "#0f172a", 'width': 4},
                'thickness': 0.75,
                'value': score
            }
        }
    ))
    fig.update_layout(height=250, margin=dict(l=20, r=20, t=0, b=20))
    return fig

def create_missing_heatmap(df):
    missing_df = df.isna().sum().to_frame(name='Missing Count')
    missing_df['Percentage'] = (missing_df['Missing Count'] / len(df)) * 100
    missing_df = missing_df[missing_df['Missing Count'] > 0].sort_values('Percentage', ascending=True)
    
    if missing_df.empty:
        return None
    
    fig = px.bar(
        missing_df, 
        x='Percentage', 
        y=missing_df.index,
        orientation='h',
        labels={'Percentage': '% Missing', 'index': 'Column'},
        color='Percentage',
        color_continuous_scale=['#10B981', '#F59E0B', '#EF4444']
    )
    fig.update_layout(height=300, margin=dict(l=20, r=20, t=20, b=20), plot_bgcolor='white', paper_bgcolor='white')
    return fig

def ai_copilot_response(question, analytics):
    q = question.lower()
    if "unhealthy" in q or "why" in q:
        return f"Your dataset health score is {analytics['health_score']}/100. The main issues dragging down the score are: {analytics['total_missing']} missing values, {analytics['total_duplicates']} duplicate rows, and {analytics['total_outliers']} outliers. I recommend addressing missing values first by either imputing or dropping them."
    elif "fix first" in q or "columns" in q:
        missing_cols = sorted(analytics['missing_by_col'].items(), key=lambda x: x[1], reverse=True)[:3]
        cols = ", ".join([f"{c[0]} ({c[1]} missing)" for c in missing_cols if c[1] > 0])
        return f"You should fix these columns first: {cols}. Prioritizing these will yield the highest improvement in data quality."
    elif "cleaning" in q or "recommendations" in q:
        return "Here are my AI recommendations: \n1. Remove exact duplicate rows. \n2. Standardize all email formats to lowercase and remove invalid entries. \n3. Impute missing numeric values (like Unit Price) with 0 or the column median. \n4. Convert all dates to YYYY-MM-DD format."
    else:
        return "I can help you understand your dataset's health. Try asking: 'Why is my dataset unhealthy?', 'Which columns should I fix first?', or 'Explain the cleaning recommendations'."

# ── UI Layout ─────────────────────────────────────────────────────────────────

# Hero Section
st.markdown("""
<div class="hero">
    <h1>AI Excel Data Quality Assistant</h1>
    <p>Automatically detect, clean, validate, and transform messy datasets using AI-powered data quality intelligence.</p>
</div>
""", unsafe_allow_html=True)

# Upload Section
uploaded = st.file_uploader("Drop your Excel or CSV file here", type=["xlsx", "csv"], label_visibility="collapsed")

if not uploaded:
    st.markdown("""
    <div style="text-align:center; color:#64748b; padding: 2rem 0; font-size:0.9rem;">
        <b>Supported formats:</b> .xlsx, .xls, .csv &nbsp;|&nbsp; <b>Secure:</b> Files are processed in-memory and never stored.
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Load Data
@st.cache_data(show_spinner=False)
def load_file(file_bytes, name):
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes), dtype=str)
    else:
        return pd.read_excel(io.BytesIO(file_bytes), dtype=str)

try:
    file_bytes = uploaded.read()
    df_raw = load_file(file_bytes, uploaded.name)
except Exception as e:
    st.error(f"Error reading file: {e}")
    st.stop()

# File Info Card
st.markdown("<div class='card' style='margin-bottom: 2rem; padding: 15px 25px;'>", unsafe_allow_html=True)
col1, col2, col3, col4 = st.columns([2, 1, 1, 1])
with col1:
    st.markdown(f"**📄 File:** {uploaded.name}")
with col2:
    st.markdown(f"**📏 Size:** {uploaded.size / 1024:.1f} KB")
with col3:
    st.markdown(f"**🔢 Rows:** {len(df_raw):,}")
with col4:
    st.markdown(f"**📊 Cols:** {len(df_raw.columns)}")
st.markdown("</div>", unsafe_allow_html=True)

# Generate Analytics
analytics = generate_analytics(df_raw)

# ── KPI Dashboard ─────────────────────────────────────────────────────────────
st.markdown("<div class='section-header'>Dataset Health Overview</div>", unsafe_allow_html=True)

kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
with kpi1:
    st.markdown(f"<div class='kpi-card kpi-blue'><div class='label'>Health Score</div><div class='value'>{analytics['health_score']}</div></div>", unsafe_allow_html=True)
with kpi2:
    st.markdown(f"<div class='kpi-card kpi-red'><div class='label'>Missing Values</div><div class='value'>{analytics['total_missing']:,}</div></div>", unsafe_allow_html=True)
with kpi3:
    st.markdown(f"<div class='kpi-card kpi-purple'><div class='label'>Duplicates</div><div class='value'>{analytics['total_duplicates']:,}</div></div>", unsafe_allow_html=True)
with kpi4:
    st.markdown(f"<div class='kpi-card kpi-red'><div class='label'>Outliers</div><div class='value'>{analytics['total_outliers']:,}</div></div>", unsafe_allow_html=True)
with kpi5:
    st.markdown(f"<div class='kpi-card kpi-green'><div class='label'>Columns</div><div class='value'>{analytics['total_columns']}</div></div>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Visualizations ────────────────────────────────────────────────────────────
viz1, viz2 = st.columns([1, 2])

with viz1:
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.markdown("<h3 style='margin-top:0; color:#1e293b;'>Data Quality Score</h3>", unsafe_allow_html=True)
    st.plotly_chart(create_gauge_chart(analytics['health_score']), use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

with viz2:
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.markdown("<h3 style='margin-top:0; color:#1e293b;'>Missing Value Analysis</h3>", unsafe_allow_html=True)
    heatmap = create_missing_heatmap(df_raw)
    if heatmap:
        st.plotly_chart(heatmap, use_container_width=True)
    else:
        st.markdown("<div style='text-align:center; padding:3rem; color:#16A34A;'>✅ No missing values detected!</div>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ── AI Data Copilot ───────────────────────────────────────────────────────────
st.markdown("<div class='section-header'>AI Data Quality Copilot</div>", unsafe_allow_html=True)

st.markdown("<div class='card' style='background:#f8fafc; border-color:#e2e8f0;'>", unsafe_allow_html=True)
st.markdown("<h3 style='margin-top:0; color:#1e293b;'>💬 Ask our AI about your dataset</h3>", unsafe_allow_html=True)

# Chat History
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = [
        ("ai", "Hello! I've analyzed your dataset. Ask me anything about its quality or how to fix it.")
    ]

for sender, msg in st.session_state.chat_history:
    css_class = "chat-ai" if sender == "ai" else "chat-user"
    st.markdown(f"<div class='chat-msg {css_class}'>{msg}</div>", unsafe_allow_html=True)

# Quick buttons
btn_col1, btn_col2, btn_col3 = st.columns(3)
if btn_col1.button("Why is my dataset unhealthy?"):
    st.session_state.chat_history.append(("user", "Why is my dataset unhealthy?"))
    st.session_state.chat_history.append(("ai", ai_copilot_response("unhealthy", analytics)))
    st.rerun()
if btn_col2.button("Which columns to fix first?"):
    st.session_state.chat_history.append(("user", "Which columns should I fix first?"))
    st.session_state.chat_history.append(("ai", ai_copilot_response("fix first", analytics)))
    st.rerun()
if btn_col3.button("Explain recommendations"):
    st.session_state.chat_history.append(("user", "Explain the cleaning recommendations."))
    st.session_state.chat_history.append(("ai", ai_copilot_response("recommendations", analytics)))
    st.rerun()

user_input = st.text_input("Type your question here...", key="chat_input", label_visibility="collapsed")
if st.button("Send") and user_input:
    st.session_state.chat_history.append(("user", user_input))
    st.session_state.chat_history.append(("ai", ai_copilot_response(user_input, analytics)))
    st.rerun()

st.markdown("</div>", unsafe_allow_html=True)

# ── Data Preview with Highlighting ────────────────────────────────────────────
st.markdown("<div class='section-header'>Raw Data Explorer</div>", unsafe_allow_html=True)

st.markdown("<div class='card'>", unsafe_allow_html=True)

# Pagination & Search
search_col, page_col = st.columns([3, 1])
with search_col:
    search_term = st.text_input("🔍 Search in dataset...", "")
with page_col:
    page_size = st.selectbox("Rows per page", [10, 25, 50], index=0)

# Filter data
df_display = df_raw.copy()
if search_term:
    mask = df_display.apply(lambda row: row.astype(str).str.contains(search_term, case=False).any(), axis=1)
    df_display = df_display[mask]

# Pagination logic
total_pages = max(1, len(df_display) // page_size + (1 if len(df_display) % page_size > 0 else 0))
page = st.number_input("Page", 1, total_pages, 1)
start_idx = (page - 1) * page_size
end_idx = start_idx + page_size
df_page = df_display.iloc[start_idx:end_idx]

# Highlight missing values
def highlight_missing(val):
    if pd.isna(val) or (isinstance(val, str) and val.strip().lower() in ["", "n/a", "null", "none", "-", "tbd"]):
        return 'background-color: #FEE2E2; color: #DC2626;'
    return ''

st.dataframe(
    df_page.style.map(highlight_missing),
    use_container_width=True,
    height=400
)
st.markdown(f"<p style='color:#64748b; font-size:0.8rem; margin-top:10px;'>Showing rows {start_idx+1} to {min(end_idx, len(df_display))} of {len(df_display)} (Filtered from {len(df_raw)} total)</p>", unsafe_allow_html=True)
st.markdown("</div>", unsafe_allow_html=True)

# ── Process & Download Section ────────────────────────────────────────────────
st.markdown("<div class='section-header'>Automated Cleaning Engine</div>", unsafe_allow_html=True)

st.markdown("<div class='card' style='text-align:center; padding: 40px;'>", unsafe_allow_html=True)
st.markdown("<h3 style='margin-top:0; color:#1e293b;'>Your Clean Dataset is Ready</h3>", unsafe_allow_html=True)
st.markdown("<p style='color:#64748b;'>Click the button below to run the AI-powered cleaning engine. This will standardize formats, fix missing values, recalculate totals, and remove duplicates.</p>", unsafe_allow_html=True)

if st.button("✨ Run AI Cleaning Engine"):
    with st.spinner('Processing your data...'):
        df_clean, report = clean_dataframe(df_raw.copy())
        
        # Store in session
        st.session_state['df_clean'] = df_clean
        st.session_state['report'] = report

if 'df_clean' in st.session_state:
    st.markdown("<div style='margin-top: 30px;'>", unsafe_allow_html=True)
    st.markdown("<div class='badge badge-green' style='font-size: 1rem; padding: 8px 16px;'>✅ Cleaning Complete!</div>", unsafe_allow_html=True)
    
    df_clean = st.session_state['df_clean']
    report = st.session_state['report']
    
    # Before/After metrics
    bef_col, aft_col = st.columns(2)
    with bef_col:
        st.metric("Rows Before", report['original_shape'][0])
    with aft_col:
        st.metric("Rows After (Cleaned)", report['final_shape'][0], delta=f"-{report['duplicates_removed']} duplicates")
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Export options
    dl_col1, dl_col2 = st.columns(2)
    
    with dl_col1:
        # Excel Export
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
        write_cleaned_excel(df_clean, report, tmp_path)
        with open(tmp_path, "rb") as f:
            excel_bytes = f.read()
        os.unlink(tmp_path)
        
        st.download_button(
            label="📊 Download Clean Excel",
            data=excel_bytes,
            file_name=f"cleaned_data_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with dl_col2:
        # CSV Export
        csv_data = df_clean.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📄 Download Clean CSV",
            data=csv_data,
            file_name=f"cleaned_data_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("</div>", unsafe_allow_html=True)
