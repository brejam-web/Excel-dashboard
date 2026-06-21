"""
Excel Insight Dashboard
------------------------
Upload an Excel file and get an automatic, interactive dashboard with
descriptive statistics, distributions, correlations, and filtering.

Run with:
    streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go

# -------------------------------------------------------------------
# Page config & style
# -------------------------------------------------------------------
st.set_page_config(
    page_title="Ledger — Spreadsheet Dashboards",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

PRIMARY = "#2E7D5B"      # ledger green — accent, derived from spreadsheet grid lines, desaturated
ACCENT = "#C97A2B"       # amber — secondary accent for charts
INK = "#1B2430"          # near-black ink for headings/text
PAPER = "#F7F5F0"        # warm paper background
LINE = "#D8D4C9"         # hairline grid color

st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');

        html, body, [class*="css"] {{
            font-family: 'Space Grotesk', -apple-system, sans-serif;
        }}
        .main {{ background-color: {PAPER}; }}
        .stMetric {{
            background-color: white;
            border: 1px solid {LINE};
            border-radius: 4px;
            padding: 10px 15px;
            box-shadow: none;
        }}
        h1, h2, h3 {{ color: {INK}; font-family: 'Space Grotesk', sans-serif; font-weight: 600; }}
        .block-container {{ padding-top: 2rem; }}
        div[data-testid="stMetricValue"] {{ color: {PRIMARY}; font-family: 'JetBrains Mono', monospace; }}
        code, .mono {{ font-family: 'JetBrains Mono', monospace; }}

        /* ---------------- Landing page ---------------- */
        .hero-wrap {{
            border: 1px solid {LINE};
            border-radius: 6px;
            background: white;
            padding: 0;
            margin-bottom: 2rem;
            overflow: hidden;
        }}
        .hero-top {{
            display: flex;
            border-bottom: 1px solid {LINE};
        }}
        .hero-eyebrow {{
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.72rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            color: {PRIMARY};
            padding: 14px 28px 0 28px;
        }}
        .hero-headline {{
            font-size: 2.6rem;
            font-weight: 700;
            line-height: 1.12;
            color: {INK};
            padding: 6px 28px 18px 28px;
            margin: 0;
            letter-spacing: -0.01em;
        }}
        .hero-headline span {{ color: {PRIMARY}; }}
        .hero-sub {{
            font-size: 1.05rem;
            color: #5b5f66;
            padding: 0 28px 24px 28px;
            max-width: 640px;
            line-height: 1.55;
        }}
        .ledger-demo {{
            display: flex;
            align-items: stretch;
            border-top: 1px dashed {LINE};
        }}
        .ledger-grid {{
            flex: 1.1;
            font-family: 'JetBrains Mono', monospace;
            font-size: 0.78rem;
            padding: 18px 0 18px 28px;
        }}
        .ledger-row {{
            display: grid;
            grid-template-columns: 1.3fr 0.9fr 0.9fr;
            border-bottom: 1px solid {LINE};
        }}
        .ledger-row.header {{
            color: {INK};
            font-weight: 600;
            border-bottom: 1px solid {INK};
        }}
        .ledger-row div {{
            padding: 7px 10px;
            color: #444;
            border-right: 1px solid {LINE};
        }}
        .ledger-row div:last-child {{ border-right: none; }}
        .ledger-arrow {{
            flex: 0.25;
            display: flex;
            align-items: center;
            justify-content: center;
            color: {LINE};
            font-size: 1.6rem;
        }}
        .ledger-chart {{
            flex: 0.9;
            padding: 18px 28px 18px 0;
            display: flex;
            align-items: flex-end;
            gap: 8px;
            height: 100%;
        }}
        .bar {{
            flex: 1;
            background: {PRIMARY};
            border-radius: 2px 2px 0 0;
            opacity: 0.85;
        }}
        .bar:nth-child(2) {{ background: {ACCENT}; }}
        .bar:nth-child(4) {{ background: {ACCENT}; }}

        /* Feature row, ledger-cell style not card-shadow style */
        .feat-row {{ display: flex; gap: 0; border: 1px solid {LINE}; border-radius: 6px; overflow: hidden; margin-bottom: 1.6rem; }}
        .feat-cell {{ flex: 1; padding: 20px 22px; border-right: 1px solid {LINE}; background: white; }}
        .feat-cell:last-child {{ border-right: none; }}
        .feat-num {{ font-family: 'JetBrains Mono', monospace; font-size: 0.72rem; color: {PRIMARY}; letter-spacing: 0.06em; }}
        .feat-title {{ font-weight: 600; color: {INK}; font-size: 1.0rem; margin: 6px 0 4px 0; }}
        .feat-desc {{ font-size: 0.88rem; color: #6b6f76; line-height: 1.45; }}

        .footnote {{ font-family: 'JetBrains Mono', monospace; font-size: 0.75rem; color: #9a978c; text-align: center; padding-top: 0.5rem; }}

        @media (max-width: 900px) {{
            .ledger-demo {{ flex-direction: column; }}
            .hero-headline {{ font-size: 1.9rem; }}
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

px.defaults.template = "plotly_white"
px.defaults.color_continuous_scale = "Viridis"
COLOR_SEQ = px.colors.qualitative.Set2


# -------------------------------------------------------------------
# Helpers
# -------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_excel(file):
    """Load all sheets from an uploaded Excel file."""
    xls = pd.ExcelFile(file)
    sheets = {name: xls.parse(name) for name in xls.sheet_names}
    return sheets


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Light auto-cleaning: drop empty cols/rows, try to parse dates."""
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df.columns = [str(c).strip() for c in df.columns]

    for col in df.columns:
        if df[col].dtype == object:
            # Try numeric conversion
            converted = pd.to_numeric(df[col], errors="coerce")
            if converted.notna().sum() >= 0.8 * df[col].notna().sum() and df[col].notna().sum() > 0:
                df[col] = converted
                continue
            # Try date conversion
            try:
                converted_dates = pd.to_datetime(df[col], errors="coerce")
                if converted_dates.notna().sum() >= 0.8 * df[col].notna().sum() and df[col].notna().sum() > 0:
                    df[col] = converted_dates
            except Exception:
                pass
    return df


def get_column_types(df: pd.DataFrame):
    numeric_cols = df.select_dtypes(include=np.number).columns.tolist()
    datetime_cols = df.select_dtypes(include="datetime64[ns]").columns.tolist()
    categorical_cols = [
        c for c in df.columns
        if c not in numeric_cols and c not in datetime_cols
    ]
    # Treat low-cardinality numeric columns as potentially categorical too
    return numeric_cols, categorical_cols, datetime_cols


# -------------------------------------------------------------------
# Sidebar — upload & filters
# -------------------------------------------------------------------
st.sidebar.markdown(
    "<div style='font-family:Space Grotesk,sans-serif;font-weight:700;font-size:1.3rem;color:#1B2430;'>📊 Ledger</div>",
    unsafe_allow_html=True,
)
st.sidebar.caption("Spreadsheet → insight, automatically.")

uploaded_file = st.sidebar.file_uploader(
    "Upload Excel file", type=["xlsx", "xls", "xlsm"]
)

if uploaded_file is None:
    st.markdown(
        f"""
        <div class="hero-wrap">
            <div class="hero-eyebrow">SPREADSHEET → INSIGHT, NO FORMULAS</div>
            <div class="hero-headline">Your Excel file<br>has a dashboard <span>hiding inside it.</span></div>
            <div class="hero-sub">
                Upload a workbook and this app reads every column, figures out what's
                numeric, what's a date, what's a category — then builds the charts,
                stats, and filters for you. No pivot tables. No setup.
            </div>
            <div class="ledger-demo">
                <div class="ledger-grid">
                    <div class="ledger-row header"><div>region</div><div>units</div><div>revenue</div></div>
                    <div class="ledger-row"><div>north</div><div>184</div><div>9,210</div></div>
                    <div class="ledger-row"><div>south</div><div>261</div><div>13,040</div></div>
                    <div class="ledger-row"><div>east</div><div>97</div><div>4,870</div></div>
                    <div class="ledger-row"><div>west</div><div>305</div><div>15,690</div></div>
                </div>
                <div class="ledger-arrow">→</div>
                <div class="ledger-chart">
                    <div class="bar" style="height:46%"></div>
                    <div class="bar" style="height:65%"></div>
                    <div class="bar" style="height:24%"></div>
                    <div class="bar" style="height:78%"></div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="feat-row">
            <div class="feat-cell">
                <div class="feat-num">READS</div>
                <div class="feat-title">Every sheet, as-is</div>
                <div class="feat-desc">Multi-sheet workbooks, messy headers, mixed text and numbers — it cleans and types each column automatically.</div>
            </div>
            <div class="feat-cell">
                <div class="feat-num">BUILDS</div>
                <div class="feat-title">Stats and charts</div>
                <div class="feat-desc">Descriptive statistics, distributions, correlations, and a full auto-generated chart grid, all in one pass.</div>
            </div>
            <div class="feat-cell">
                <div class="feat-num">RESPONDS</div>
                <div class="feat-title">To your filters, live</div>
                <div class="feat-desc">Slice by category, numeric range, or date in the sidebar — every chart and stat updates together.</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("##### Start here")
    st.caption("Drop a workbook in the sidebar uploader, or use the field below.")
    uploaded_file_main = st.file_uploader(
        "Upload Excel file", type=["xlsx", "xls", "xlsm"], key="main_uploader",
        label_visibility="collapsed",
    )
    if uploaded_file_main is not None:
        uploaded_file = uploaded_file_main

    st.markdown(
        '<div class="footnote">.xlsx · .xls · .xlsm — processed in your session, never stored</div>',
        unsafe_allow_html=True,
    )

if uploaded_file is None:
    st.stop()

# Load workbook
try:
    sheets = load_excel(uploaded_file)
except Exception as e:
    st.error(f"Couldn't read this file. Error: {e}")
    st.stop()

sheet_names = list(sheets.keys())
if len(sheet_names) > 1:
    selected_sheet = st.sidebar.selectbox("Select sheet", sheet_names)
else:
    selected_sheet = sheet_names[0]

raw_df = sheets[selected_sheet]
df = clean_dataframe(raw_df.copy())

if df.empty:
    st.warning("This sheet appears to be empty after cleaning.")
    st.stop()

numeric_cols, categorical_cols, datetime_cols = get_column_types(df)

# ---- Sidebar filters ----
st.sidebar.markdown("---")
st.sidebar.subheader("🔍 Filters")

filtered_df = df.copy()

with st.sidebar.expander("Categorical filters", expanded=False):
    for col in categorical_cols:
        unique_vals = df[col].dropna().unique().tolist()
        if 1 < len(unique_vals) <= 50:
            selected_vals = st.multiselect(f"{col}", sorted(map(str, unique_vals)), default=[])
            if selected_vals:
                filtered_df = filtered_df[filtered_df[col].astype(str).isin(selected_vals)]

with st.sidebar.expander("Numeric range filters", expanded=False):
    for col in numeric_cols:
        col_min = float(df[col].min())
        col_max = float(df[col].max())
        if col_min < col_max:
            sel_range = st.slider(f"{col}", col_min, col_max, (col_min, col_max))
            filtered_df = filtered_df[
                (filtered_df[col] >= sel_range[0]) & (filtered_df[col] <= sel_range[1])
            ]

if datetime_cols:
    with st.sidebar.expander("Date filters", expanded=False):
        for col in datetime_cols:
            min_date = df[col].min()
            max_date = df[col].max()
            if pd.notna(min_date) and pd.notna(max_date) and min_date < max_date:
                date_range = st.date_input(f"{col}", (min_date.date(), max_date.date()))
                if isinstance(date_range, tuple) and len(date_range) == 2:
                    start, end = date_range
                    filtered_df = filtered_df[
                        (filtered_df[col] >= pd.Timestamp(start))
                        & (filtered_df[col] <= pd.Timestamp(end))
                    ]

st.sidebar.markdown("---")
st.sidebar.caption(f"Showing **{len(filtered_df):,}** of **{len(df):,}** rows")

# -------------------------------------------------------------------
# Main layout
# -------------------------------------------------------------------
st.sidebar.markdown("---")
full_dashboard_mode = st.sidebar.checkbox("🖥️ Full-screen landscape dashboard mode", value=False)

st.title("📊 Excel Insight Dashboard")
st.caption(f"Sheet: **{selected_sheet}** · {df.shape[0]:,} rows × {df.shape[1]} columns")

if full_dashboard_mode:
    st.info("Full dashboard mode — all charts below in one continuous landscape view. Untick the sidebar box to return to tabs.")

    MAX_CHARTS_PER_TYPE = 8
    CHART_HEIGHT = 340
    charts = []

    for col in numeric_cols[:MAX_CHARTS_PER_TYPE]:
        fig = px.histogram(filtered_df, x=col, nbins=30, color_discrete_sequence=[PRIMARY])
        fig.update_layout(height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10))
        charts.append((f"Distribution — {col}", fig))

    for col in categorical_cols[:MAX_CHARTS_PER_TYPE]:
        n_unique = filtered_df[col].nunique()
        if 1 < n_unique <= 100:
            vc = filtered_df[col].astype(str).value_counts().reset_index().head(10)
            vc.columns = [col, "Count"]
            fig = px.bar(vc, x="Count", y=col, orientation="h", color="Count", color_continuous_scale="Viridis")
            fig.update_layout(height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10),
                               yaxis=dict(autorange="reversed"), coloraxis_showscale=False)
            charts.append((f"Top categories — {col}", fig))

    if datetime_cols and numeric_cols:
        date_col = datetime_cols[0]
        line_data = filtered_df.dropna(subset=[date_col]).sort_values(date_col)
        for val_col in numeric_cols[:MAX_CHARTS_PER_TYPE]:
            fig = px.line(line_data, x=date_col, y=val_col, color_discrete_sequence=[ACCENT])
            fig.update_layout(height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10))
            charts.append((f"{val_col} over time", fig))

    if len(numeric_cols) >= 2:
        corr = filtered_df[numeric_cols].corr()
        fig = px.imshow(corr, text_auto=".2f", color_continuous_scale="RdBu", zmin=-1, zmax=1, aspect="auto")
        fig.update_layout(height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10))
        charts.append(("Correlation heatmap", fig))

        corr_abs = filtered_df[numeric_cols].corr().abs()
        pairs = corr_abs.where(~np.eye(len(corr_abs), dtype=bool)).stack().reset_index()
        pairs.columns = ["var1", "var2", "abs_corr"]
        pairs = pairs.drop_duplicates(subset="abs_corr").sort_values("abs_corr", ascending=False)
        for _, row in pairs.head(4).iterrows():
            fig = px.scatter(filtered_df, x=row["var1"], y=row["var2"], color_discrete_sequence=[PRIMARY], opacity=0.7)
            fig.update_layout(height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10))
            charts.append((f"{row['var1']} vs {row['var2']}", fig))

    for col in categorical_cols[:MAX_CHARTS_PER_TYPE]:
        n_unique = filtered_df[col].nunique()
        if 1 < n_unique <= 8:
            vc = filtered_df[col].astype(str).value_counts().reset_index()
            vc.columns = [col, "Count"]
            fig = px.pie(vc, names=col, values="Count", color_discrete_sequence=COLOR_SEQ, hole=0.35)
            fig.update_layout(height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10))
            charts.append((f"Share breakdown — {col}", fig))

    if not charts:
        st.info("No applicable charts could be generated for this dataset.")
    else:
        st.caption(f"{len(charts)} charts · 4 per row")
        cols_per_row = 4
        for i in range(0, len(charts), cols_per_row):
            row_charts = charts[i:i + cols_per_row]
            cols = st.columns(cols_per_row)
            for col_widget, (title, fig) in zip(cols, row_charts):
                with col_widget:
                    st.markdown(f"**{title}**")
                    st.plotly_chart(fig, use_container_width=True)

    st.stop()

tab_overview, tab_stats, tab_dashboard, tab_visuals, tab_correlation, tab_data = st.tabs(
    ["🏠 Overview", "🧮 Descriptive Stats", "🖼️ Auto Dashboard", "📈 Visual Explorer", "🔗 Correlations", "🗂️ Raw Data"]
)

# ---------------- Overview tab ----------------
with tab_overview:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Rows", f"{filtered_df.shape[0]:,}")
    c2.metric("Columns", f"{filtered_df.shape[1]:,}")
    c3.metric("Numeric columns", len(numeric_cols))
    c4.metric("Missing values", f"{int(filtered_df.isna().sum().sum()):,}")

    st.markdown("#### Column overview")
    overview_rows = []
    for col in filtered_df.columns:
        overview_rows.append({
            "Column": col,
            "Type": str(filtered_df[col].dtype),
            "Unique values": filtered_df[col].nunique(),
            "Missing": int(filtered_df[col].isna().sum()),
            "Missing %": round(100 * filtered_df[col].isna().mean(), 1),
        })
    st.dataframe(pd.DataFrame(overview_rows), use_container_width=True, hide_index=True)

    if numeric_cols:
        st.markdown("#### Quick look — first numeric column")
        first_num = numeric_cols[0]
        fig = px.histogram(
            filtered_df, x=first_num, nbins=30,
            color_discrete_sequence=[PRIMARY],
            title=f"Distribution of {first_num}",
        )
        st.plotly_chart(fig, use_container_width=True)

# ---------------- Descriptive stats tab ----------------
with tab_stats:
    st.markdown("#### Numeric summary")
    if numeric_cols:
        desc = filtered_df[numeric_cols].describe().T
        desc["median"] = filtered_df[numeric_cols].median()
        desc["skew"] = filtered_df[numeric_cols].skew()
        desc["variance"] = filtered_df[numeric_cols].var()
        desc = desc[["count", "mean", "median", "std", "variance", "skew", "min", "25%", "50%", "75%", "max"]]
        st.dataframe(desc.round(3), use_container_width=True)
    else:
        st.info("No numeric columns detected in this sheet.")

    if categorical_cols:
        st.markdown("#### Categorical summary")
        cat_col = st.selectbox("Choose a categorical column", categorical_cols, key="cat_summary")
        value_counts = filtered_df[cat_col].astype(str).value_counts().reset_index()
        value_counts.columns = [cat_col, "Count"]
        col1, col2 = st.columns([1, 1])
        with col1:
            st.dataframe(value_counts, use_container_width=True, hide_index=True)
        with col2:
            top_n = value_counts.head(15)
            fig = px.bar(
                top_n, x="Count", y=cat_col, orientation="h",
                color="Count", color_continuous_scale="Viridis",
                title=f"Top categories — {cat_col}",
            )
            fig.update_layout(yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig, use_container_width=True)

# ---------------- Auto Dashboard tab (all charts, one grid) ----------------
with tab_dashboard:
    st.markdown("#### Everything in one view")
    st.caption(
        "Every applicable chart for this dataset, generated automatically and laid out in a grid."
    )

    MAX_CHARTS_PER_TYPE = 8  # safety cap so huge datasets don't render hundreds of charts
    CHART_HEIGHT = 320

    charts = []  # list of (title, plotly_fig)

    # 1. Histograms for numeric columns
    for col in numeric_cols[:MAX_CHARTS_PER_TYPE]:
        fig = px.histogram(
            filtered_df, x=col, nbins=30,
            color_discrete_sequence=[PRIMARY],
        )
        fig.update_layout(height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10))
        charts.append((f"Distribution — {col}", fig))

    # 2. Bar charts for categorical columns (top categories)
    for col in categorical_cols[:MAX_CHARTS_PER_TYPE]:
        n_unique = filtered_df[col].nunique()
        if 1 < n_unique <= 100:
            vc = filtered_df[col].astype(str).value_counts().reset_index().head(10)
            vc.columns = [col, "Count"]
            fig = px.bar(
                vc, x="Count", y=col, orientation="h",
                color="Count", color_continuous_scale="Viridis",
            )
            fig.update_layout(
                height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10),
                yaxis=dict(autorange="reversed"), coloraxis_showscale=False,
            )
            charts.append((f"Top categories — {col}", fig))

    # 3. Time series line charts (date x numeric)
    if datetime_cols and numeric_cols:
        date_col = datetime_cols[0]
        line_data = filtered_df.dropna(subset=[date_col]).sort_values(date_col)
        for val_col in numeric_cols[:MAX_CHARTS_PER_TYPE]:
            fig = px.line(line_data, x=date_col, y=val_col, color_discrete_sequence=[ACCENT])
            fig.update_layout(height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10))
            charts.append((f"{val_col} over time", fig))

    # 4. Correlation heatmap
    if len(numeric_cols) >= 2:
        corr = filtered_df[numeric_cols].corr()
        fig = px.imshow(
            corr, text_auto=".2f", color_continuous_scale="RdBu", zmin=-1, zmax=1,
            aspect="auto",
        )
        fig.update_layout(height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10))
        charts.append(("Correlation heatmap", fig))

    # 5. Top scatter plots — strongest correlated numeric pairs, capped
    if len(numeric_cols) >= 2:
        corr = filtered_df[numeric_cols].corr().abs()
        pairs = (
            corr.where(~np.eye(len(corr), dtype=bool))
            .stack()
            .reset_index()
        )
        pairs.columns = ["var1", "var2", "abs_corr"]
        pairs = pairs.drop_duplicates(subset="abs_corr").sort_values("abs_corr", ascending=False)
        for _, row in pairs.head(4).iterrows():
            fig = px.scatter(
                filtered_df, x=row["var1"], y=row["var2"],
                color_discrete_sequence=[PRIMARY], opacity=0.7,
            )
            fig.update_layout(height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10))
            charts.append((f"{row['var1']} vs {row['var2']}", fig))

    # 6. Pie charts for low-cardinality categoricals
    for col in categorical_cols[:MAX_CHARTS_PER_TYPE]:
        n_unique = filtered_df[col].nunique()
        if 1 < n_unique <= 8:
            vc = filtered_df[col].astype(str).value_counts().reset_index()
            vc.columns = [col, "Count"]
            fig = px.pie(vc, names=col, values="Count", color_discrete_sequence=COLOR_SEQ, hole=0.35)
            fig.update_layout(height=CHART_HEIGHT, margin=dict(t=30, b=10, l=10, r=10))
            charts.append((f"Share breakdown — {col}", fig))

    if not charts:
        st.info("No applicable charts could be generated for this dataset.")
    else:
        st.caption(f"{len(charts)} charts generated · scroll for more · use the sidebar filters to update them all at once")
        cols_per_row = 3
        for i in range(0, len(charts), cols_per_row):
            row_charts = charts[i:i + cols_per_row]
            cols = st.columns(cols_per_row)
            for col_widget, (title, fig) in zip(cols, row_charts):
                with col_widget:
                    st.markdown(f"**{title}**")
                    st.plotly_chart(fig, use_container_width=True)

# ---------------- Visual explorer tab ----------------
with tab_visuals:
    st.markdown("#### Build your own chart")
    chart_type = st.selectbox(
        "Chart type",
        ["Histogram", "Box plot", "Scatter plot", "Bar chart", "Line chart", "Pie chart"],
    )

    if chart_type == "Histogram" and numeric_cols:
        col = st.selectbox("Column", numeric_cols, key="hist_col")
        color_col = st.selectbox("Color by (optional)", ["None"] + categorical_cols, key="hist_color")
        fig = px.histogram(
            filtered_df, x=col,
            color=None if color_col == "None" else color_col,
            nbins=40, color_discrete_sequence=COLOR_SEQ,
            marginal="box",
        )
        st.plotly_chart(fig, use_container_width=True)

    elif chart_type == "Box plot" and numeric_cols:
        col = st.selectbox("Numeric column", numeric_cols, key="box_col")
        group_col = st.selectbox("Group by (optional)", ["None"] + categorical_cols, key="box_group")
        fig = px.box(
            filtered_df, y=col,
            x=None if group_col == "None" else group_col,
            color=None if group_col == "None" else group_col,
            color_discrete_sequence=COLOR_SEQ,
            points="outliers",
        )
        st.plotly_chart(fig, use_container_width=True)

    elif chart_type == "Scatter plot" and len(numeric_cols) >= 2:
        x_col = st.selectbox("X axis", numeric_cols, key="scatter_x")
        y_col = st.selectbox("Y axis", [c for c in numeric_cols if c != x_col], key="scatter_y")
        color_col = st.selectbox("Color by (optional)", ["None"] + categorical_cols, key="scatter_color")
        fig = px.scatter(
            filtered_df, x=x_col, y=y_col,
            color=None if color_col == "None" else color_col,
            color_discrete_sequence=COLOR_SEQ,
            trendline="ols" if st.checkbox("Show trendline") else None,
        )
        st.plotly_chart(fig, use_container_width=True)

    elif chart_type == "Bar chart" and categorical_cols:
        cat_col = st.selectbox("Category column", categorical_cols, key="bar_cat")
        if numeric_cols:
            agg_col = st.selectbox("Value column", ["Count"] + numeric_cols, key="bar_val")
            agg_func = st.selectbox("Aggregation", ["sum", "mean", "median", "max", "min"], key="bar_agg")
        else:
            agg_col, agg_func = "Count", "sum"

        if agg_col == "Count":
            data = filtered_df[cat_col].astype(str).value_counts().reset_index()
            data.columns = [cat_col, "Count"]
            y_field = "Count"
        else:
            data = filtered_df.groupby(filtered_df[cat_col].astype(str))[agg_col].agg(agg_func).reset_index()
            y_field = agg_col

        data = data.sort_values(y_field, ascending=False).head(25)
        fig = px.bar(data, x=cat_col, y=y_field, color=y_field, color_continuous_scale="Viridis")
        st.plotly_chart(fig, use_container_width=True)

    elif chart_type == "Line chart" and datetime_cols and numeric_cols:
        date_col = st.selectbox("Date column", datetime_cols, key="line_date")
        value_col = st.selectbox("Value column", numeric_cols, key="line_val")
        line_data = filtered_df.dropna(subset=[date_col]).sort_values(date_col)
        fig = px.line(line_data, x=date_col, y=value_col, color_discrete_sequence=[ACCENT])
        st.plotly_chart(fig, use_container_width=True)

    elif chart_type == "Pie chart" and categorical_cols:
        cat_col = st.selectbox("Category column", categorical_cols, key="pie_cat")
        data = filtered_df[cat_col].astype(str).value_counts().reset_index().head(10)
        data.columns = [cat_col, "Count"]
        fig = px.pie(data, names=cat_col, values="Count", color_discrete_sequence=COLOR_SEQ, hole=0.35)
        st.plotly_chart(fig, use_container_width=True)

    else:
        st.info("Not enough of the right column types are available for this chart. Try another chart type.")

# ---------------- Correlation tab ----------------
with tab_correlation:
    if len(numeric_cols) >= 2:
        st.markdown("#### Correlation matrix")
        corr = filtered_df[numeric_cols].corr()
        fig = px.imshow(
            corr, text_auto=".2f", color_continuous_scale="RdBu", zmin=-1, zmax=1,
            aspect="auto",
        )
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Strongest relationships")
        corr_pairs = (
            corr.where(~np.eye(len(corr), dtype=bool))
            .stack()
            .reset_index()
        )
        corr_pairs.columns = ["Variable 1", "Variable 2", "Correlation"]
        corr_pairs["abs_corr"] = corr_pairs["Correlation"].abs()
        corr_pairs = corr_pairs.sort_values("abs_corr", ascending=False).drop_duplicates(subset="abs_corr")
        st.dataframe(corr_pairs.head(10)[["Variable 1", "Variable 2", "Correlation"]].round(3), hide_index=True, use_container_width=True)
    else:
        st.info("Need at least two numeric columns to compute correlations.")

# ---------------- Raw data tab ----------------
with tab_data:
    st.markdown("#### Filtered data")
    st.dataframe(filtered_df, use_container_width=True)
    csv = filtered_df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇️ Download filtered data as CSV", csv, "filtered_data.csv", "text/csv")
