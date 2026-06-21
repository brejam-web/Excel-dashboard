"""
Ledger Pro — Premium Analytics Workspace
------------------------------------------
A dark, glassmorphic Streamlit dashboard for exploring Excel data:
upload -> KPI matrix -> descriptive stats -> inferential significance testing
-> correlation -> free-form chart builder -> raw data export.

Run with:
    streamlit run app.py
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from scipy import stats

# ============================================================================
# 1. GLOBAL CANVAS SETUP
# ============================================================================
st.set_page_config(
    page_title="Ledger Pro — Analytics Workspace",
    page_icon="◆",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---- Design tokens ----
BG_DEEP   = "#0B0E14"   # page background, near-black with a blue cast
BG_PANEL  = "#11151D"   # card/panel background
BG_RAISED = "#161B26"   # raised elements (inputs, expanders)
BORDER    = "#222837"   # default hairline border
BORDER_HI = "#2E3648"   # slightly brighter border for hover rest-state
ACCENT    = "#6E8BFF"   # primary accent — cool electric blue/indigo
ACCENT_2  = "#34D8C4"   # secondary accent — teal, for contrast series
WARN      = "#F2A65A"   # amber, used sparingly for inferential/attention items
TEXT_HI   = "#E8EAF0"   # primary text
TEXT_MED  = "#9AA3B5"   # secondary text
TEXT_LOW  = "#5C6578"   # tertiary / caption text

PLOTLY_SERIES = [ACCENT, ACCENT_2, WARN, "#C792EA", "#FF6B81", "#7FD9A0"]

px.defaults.template = "plotly_dark"
px.defaults.color_discrete_sequence = PLOTLY_SERIES


# ============================================================================
# 2. NATIVE CSS STYLE INJECTION
# ============================================================================
st.html(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap');

/* ---------- Page-load fade + slide animation ---------- */
@keyframes fadeSlideIn {{
    from {{ opacity: 0; transform: translateY(14px); }}
    to   {{ opacity: 1; transform: translateY(0); }}
}}
.main .block-container {{
    animation: fadeSlideIn 0.55s ease-out;
}}

/* ---------- Base canvas ---------- */
html, body, [class*="css"] {{
    font-family: 'Inter', -apple-system, sans-serif;
}}
.stApp {{
    background-color: {BG_DEEP};
}}
.main .block-container {{
    padding-top: 1.6rem;
    padding-bottom: 3rem;
    max-width: 1320px;
}}
h1, h2, h3, h4 {{
    color: {TEXT_HI} !important;
    font-family: 'Inter', sans-serif;
    font-weight: 700;
}}
p, span, label, .stMarkdown {{
    color: {TEXT_MED};
}}
code, .mono {{ font-family: 'JetBrains Mono', monospace; }}

/* ---------- Hide native Streamlit chrome (white-label) ---------- */
#MainMenu {{ visibility: hidden; }}
header[data-testid="stHeader"] {{ background: transparent; height: 0; }}
footer {{ visibility: hidden; height: 0; }}
div[data-testid="stToolbar"] {{ visibility: hidden; height: 0; }}
div[data-testid="stDecoration"] {{ display: none; }}
#stDecoration {{ display: none; }}

/* ---------- Sidebar ---------- */
section[data-testid="stSidebar"] {{
    background-color: {BG_PANEL};
    border-right: 1px solid {BORDER};
}}
section[data-testid="stSidebar"] .block-container {{
    padding-top: 1.8rem;
}}

/* ---------- KPI Metric cards: glassmorphism ---------- */
div[data-testid="stMetric"] {{
    background: linear-gradient(160deg, {BG_RAISED} 0%, {BG_PANEL} 100%);
    border: 1px solid {BORDER};
    border-radius: 14px;
    padding: 18px 20px 16px 20px;
    box-shadow: 0 1px 0 rgba(255,255,255,0.02) inset, 0 8px 24px rgba(0,0,0,0.25);
    transition: transform 0.22s ease, border-color 0.22s ease, box-shadow 0.22s ease;
}}
div[data-testid="stMetric"]:hover {{
    transform: translateY(-4px);
    border-color: {ACCENT};
    box-shadow: 0 0 0 1px {ACCENT}33, 0 14px 28px rgba(0,0,0,0.35);
}}
div[data-testid="stMetricLabel"] {{
    color: {TEXT_MED} !important;
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    font-weight: 600;
}}
div[data-testid="stMetricValue"] {{
    color: {TEXT_HI} !important;
    font-family: 'JetBrains Mono', monospace;
    font-weight: 600;
}}
div[data-testid="stMetricDelta"] {{ font-family: 'JetBrains Mono', monospace; }}

/* ---------- Generic panel card (used for chart wrappers, hero, etc.) ---------- */
.panel {{
    background: {BG_PANEL};
    border: 1px solid {BORDER};
    border-radius: 14px;
    padding: 22px 24px;
    margin-bottom: 1.25rem;
    transition: border-color 0.22s ease;
}}
.panel:hover {{ border-color: {BORDER_HI}; }}

/* ---------- Tabs styled as segmented control ---------- */
button[data-baseweb="tab"] {{
    color: {TEXT_MED};
    font-weight: 600;
    font-size: 0.92rem;
}}
button[data-baseweb="tab"][aria-selected="true"] {{
    color: {ACCENT};
}}
div[data-baseweb="tab-highlight"] {{
    background-color: {ACCENT} !important;
}}
div[data-baseweb="tab-border"] {{ background-color: {BORDER}; }}

/* ---------- Buttons ---------- */
.stButton button, .stDownloadButton button {{
    background-color: {ACCENT};
    color: #0B0E14;
    border: none;
    border-radius: 8px;
    font-weight: 600;
    transition: transform 0.15s ease, box-shadow 0.15s ease;
}}
.stButton button:hover, .stDownloadButton button:hover {{
    transform: translateY(-2px);
    box-shadow: 0 6px 16px {ACCENT}40;
}}

/* ---------- File uploader dropzone ---------- */
div[data-testid="stFileUploaderDropzone"] {{
    background: {BG_RAISED};
    border: 1.5px dashed {BORDER_HI};
    border-radius: 14px;
    transition: border-color 0.2s ease, background 0.2s ease;
}}
div[data-testid="stFileUploaderDropzone"]:hover {{
    border-color: {ACCENT};
    background: #131826;
}}

/* ---------- Inputs / selects / sliders / expanders ---------- */
div[data-baseweb="select"] > div, .stTextInput input, .stNumberInput input {{
    background-color: {BG_RAISED} !important;
    border-color: {BORDER} !important;
    color: {TEXT_HI} !important;
}}
div[data-testid="stExpander"] {{
    background-color: {BG_PANEL};
    border: 1px solid {BORDER};
    border-radius: 12px;
}}
.stSlider [data-baseweb="slider"] div {{ background-color: {ACCENT}; }}

/* ---------- Dataframe ---------- */
div[data-testid="stDataFrame"] {{
    border: 1px solid {BORDER};
    border-radius: 10px;
    overflow: hidden;
}}

/* ---------- Eyebrow / pill labels ---------- */
.eyebrow {{
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: {ACCENT};
}}
.pill {{
    display: inline-block;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.7rem;
    padding: 3px 9px;
    border-radius: 999px;
    border: 1px solid {BORDER_HI};
    color: {TEXT_MED};
    margin-right: 6px;
}}
.pill-accent {{ border-color: {ACCENT}66; color: {ACCENT}; }}
.pill-warn {{ border-color: {WARN}66; color: {WARN}; }}

hr {{ border-color: {BORDER}; }}
</style>
""")


# ============================================================================
# Helpers
# ============================================================================
@st.cache_data(show_spinner=False)
def load_excel(file):
    xls = pd.ExcelFile(file)
    return {name: xls.parse(name) for name in xls.sheet_names}


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Light auto-cleaning: drop empty cols/rows, coerce numeric/date text columns."""
    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    df.columns = [str(c).strip() for c in df.columns]

    for col in df.columns:
        if df[col].dtype == object:
            converted = pd.to_numeric(df[col], errors="coerce")
            if converted.notna().sum() >= 0.8 * df[col].notna().sum() and df[col].notna().sum() > 0:
                df[col] = converted
                continue
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
    categorical_cols = [c for c in df.columns if c not in numeric_cols and c not in datetime_cols]
    return numeric_cols, categorical_cols, datetime_cols


def style_fig(fig, height=380):
    """Apply the dark, transparent, glass-compatible look to any Plotly figure."""
    fig.update_layout(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", color=TEXT_MED, size=12),
        height=height,
        margin=dict(t=44, b=36, l=44, r=24),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(bgcolor=BG_RAISED, font_color=TEXT_HI, bordercolor=BORDER_HI),
    )
    fig.update_xaxes(gridcolor=BORDER, zerolinecolor=BORDER)
    fig.update_yaxes(gridcolor=BORDER, zerolinecolor=BORDER)
    return fig


def panel_open(title=None, eyebrow=None):
    html = '<div class="panel">'
    if eyebrow:
        html += f'<div class="eyebrow mb-1">{eyebrow}</div>'
    if title:
        html += f'<div style="font-weight:700;font-size:1.05rem;color:{TEXT_HI};margin-bottom:0.6rem;">{title}</div>'
    st.html(html)


# scipy.stats based significance tests --------------------------------------
def run_ttest(df, numeric_col, group_col):
    groups = df[[numeric_col, group_col]].dropna()
    cats = groups[group_col].astype(str).unique()
    if len(cats) != 2:
        return None
    a = groups[groups[group_col].astype(str) == cats[0]][numeric_col]
    b = groups[groups[group_col].astype(str) == cats[1]][numeric_col]
    if len(a) < 2 or len(b) < 2:
        return None
    t_stat, p_val = stats.ttest_ind(a, b, equal_var=False, nan_policy="omit")
    return {
        "test": "Welch's t-test (independent samples)",
        "groups": f"{cats[0]} vs {cats[1]}",
        "statistic": t_stat,
        "p_value": p_val,
        "n1": len(a), "n2": len(b),
        "mean1": a.mean(), "mean2": b.mean(),
    }


def run_anova(df, numeric_col, group_col):
    groups = df[[numeric_col, group_col]].dropna()
    cats = groups[group_col].astype(str).unique()
    if len(cats) < 3:
        return None
    samples = [groups[groups[group_col].astype(str) == c][numeric_col] for c in cats]
    samples = [s for s in samples if len(s) >= 2]
    if len(samples) < 3:
        return None
    f_stat, p_val = stats.f_oneway(*samples)
    return {
        "test": "One-way ANOVA",
        "groups": f"{len(samples)} groups",
        "statistic": f_stat,
        "p_value": p_val,
        "n_groups": len(samples),
    }


def run_chi_square(df, col_a, col_b):
    table = pd.crosstab(df[col_a].astype(str), df[col_b].astype(str))
    if table.shape[0] < 2 or table.shape[1] < 2:
        return None
    chi2, p_val, dof, _ = stats.chi2_contingency(table)
    return {
        "test": "Chi-square test of independence",
        "groups": f"{col_a} × {col_b}",
        "statistic": chi2,
        "p_value": p_val,
        "dof": dof,
        "table": table,
    }


def run_correlation_test(df, col_a, col_b):
    sub = df[[col_a, col_b]].dropna()
    if len(sub) < 3:
        return None
    r, p_val = stats.pearsonr(sub[col_a], sub[col_b])
    return {
        "test": "Pearson correlation significance",
        "groups": f"{col_a} vs {col_b}",
        "statistic": r,
        "p_value": p_val,
        "n": len(sub),
    }


def verdict_badge(p_val, alpha=0.05):
    if p_val < alpha:
        return f'<span class="pill pill-accent">p = {p_val:.4f} — significant at α={alpha}</span>'
    return f'<span class="pill pill-warn">p = {p_val:.4f} — not significant at α={alpha}</span>'


# ============================================================================
# 3. HEADER
# ============================================================================
st.html(f"""
<div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:0.4rem;">
    <div style="font-weight:800;font-size:1.5rem;color:{TEXT_HI};letter-spacing:-0.01em;">◆ Ledger Pro</div>
    <div class="eyebrow">analytics workspace</div>
</div>
<div style="height:1px;background:{BORDER};margin-bottom:1.6rem;"></div>
""")


# ============================================================================
# SIDEBAR — Controls
# ============================================================================
with st.sidebar:
    st.html(f'<div style="font-weight:700;color:{TEXT_HI};font-size:1.05rem;margin-bottom:0.2rem;">Controls</div>')
    st.caption("Upload a workbook, then filter and configure your analysis here.")
    st.markdown("---")

    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls", "xlsm"])

    selected_sheet = None
    filtered_df = None
    df = None

    if uploaded_file is not None:
        try:
            sheets = load_excel(uploaded_file)
        except Exception as e:
            st.error(f"Couldn't read this file: {e}")
            sheets = None

        if sheets:
            sheet_names = list(sheets.keys())
            selected_sheet = st.selectbox("Sheet", sheet_names) if len(sheet_names) > 1 else sheet_names[0]
            raw_df = sheets[selected_sheet]
            df = clean_dataframe(raw_df.copy())

            if df.empty:
                st.warning("This sheet is empty after cleaning.")
            else:
                numeric_cols, categorical_cols, datetime_cols = get_column_types(df)
                filtered_df = df.copy()

                st.markdown("---")
                st.markdown("**Filters**")

                with st.expander("Categorical", expanded=False):
                    for col in categorical_cols:
                        unique_vals = df[col].dropna().unique().tolist()
                        if 1 < len(unique_vals) <= 50:
                            sel = st.multiselect(col, sorted(map(str, unique_vals)), default=[], key=f"cat_{col}")
                            if sel:
                                filtered_df = filtered_df[filtered_df[col].astype(str).isin(sel)]

                with st.expander("Numeric range", expanded=False):
                    for col in numeric_cols:
                        col_min, col_max = float(df[col].min()), float(df[col].max())
                        if col_min < col_max:
                            sel_range = st.slider(col, col_min, col_max, (col_min, col_max), key=f"num_{col}")
                            filtered_df = filtered_df[(filtered_df[col] >= sel_range[0]) & (filtered_df[col] <= sel_range[1])]

                with st.expander("Date range", expanded=False):
                    if datetime_cols:
                        for col in datetime_cols:
                            min_d, max_d = df[col].min(), df[col].max()
                            if pd.notna(min_d) and pd.notna(max_d) and min_d < max_d:
                                dr = st.date_input(col, (min_d.date(), max_d.date()), key=f"date_{col}")
                                if isinstance(dr, tuple) and len(dr) == 2:
                                    start, end = dr
                                    filtered_df = filtered_df[(filtered_df[col] >= pd.Timestamp(start)) & (filtered_df[col] <= pd.Timestamp(end))]
                    else:
                        st.caption("No date columns detected.")

                st.markdown("---")
                st.caption(f"Showing **{len(filtered_df):,}** of **{len(df):,}** rows")


# ============================================================================
# STATE 1 — Empty state / dropzone hero
# ============================================================================
if uploaded_file is None or df is None or df.empty:
    st.html(f"""
    <div class="panel" style="text-align:center;padding:64px 32px;">
        <div class="eyebrow" style="margin-bottom:10px;">step 1 of 3</div>
        <div style="font-size:1.9rem;font-weight:800;color:{TEXT_HI};margin-bottom:10px;">
            Drop a workbook to begin
        </div>
        <div style="color:{TEXT_MED};max-width:480px;margin:0 auto;line-height:1.6;">
            Use the uploader in the sidebar. Once loaded, this workspace builds
            KPI summaries, descriptive statistics, and significance tests automatically —
            no formulas required.
        </div>
        <div style="margin-top:18px;">
            <span class="pill">.xlsx</span><span class="pill">.xls</span><span class="pill">.xlsm</span>
        </div>
    </div>
    """)
    st.stop()


# ============================================================================
# STATE 2 — KPI Matrix (st.columns(3)+) immediately on upload
# ============================================================================
numeric_cols, categorical_cols, datetime_cols = get_column_types(df)

st.html(f'<div class="eyebrow" style="margin-bottom:6px;">step 2 of 3 — {selected_sheet}</div>')
st.html(f'<div style="font-weight:800;font-size:1.4rem;color:{TEXT_HI};margin-bottom:1rem;">Summary metrics</div>')

k1, k2, k3, k4 = st.columns(4)
k1.metric("Rows", f"{filtered_df.shape[0]:,}", delta=f"{filtered_df.shape[0] - df.shape[0]:,} vs unfiltered" if filtered_df.shape[0] != df.shape[0] else None)
k2.metric("Columns", f"{filtered_df.shape[1]:,}")
k3.metric("Numeric fields", len(numeric_cols))
k4.metric("Missing cells", f"{int(filtered_df.isna().sum().sum()):,}")

if numeric_cols:
    k5, k6, k7, k8 = st.columns(4)
    first = numeric_cols[0]
    k5.metric(f"{first} — mean", f"{filtered_df[first].mean():,.2f}")
    k6.metric(f"{first} — median", f"{filtered_df[first].median():,.2f}")
    k7.metric(f"{first} — std dev", f"{filtered_df[first].std():,.2f}")
    k8.metric(f"{first} — range", f"{filtered_df[first].max() - filtered_df[first].min():,.2f}")

st.write("")


# ============================================================================
# STATE 3 — Analytics workspace: Descriptive vs Inferential toggle, + extras
# ============================================================================
st.html(f'<div class="eyebrow" style="margin-bottom:6px;">step 3 of 3</div>')
st.html(f'<div style="font-weight:800;font-size:1.4rem;color:{TEXT_HI};margin-bottom:1rem;">Analytics workspace</div>')

tab_desc, tab_infer, tab_corr, tab_explore, tab_data = st.tabs([
    "📊 Descriptive Statistics", "🧪 Inferential Analysis", "🔗 Correlations", "🎛️ Chart Builder", "🗂️ Raw Data"
])

# ---------------- Descriptive Statistics ----------------
with tab_desc:
    st.caption("Distribution overviews — shape, spread, and central tendency for every field.")

    if numeric_cols:
        desc = filtered_df[numeric_cols].describe().T
        desc["median"] = filtered_df[numeric_cols].median()
        desc["skew"] = filtered_df[numeric_cols].skew()
        desc["variance"] = filtered_df[numeric_cols].var()
        desc = desc[["count", "mean", "median", "std", "variance", "skew", "min", "25%", "50%", "75%", "max"]]
        st.dataframe(desc.round(3), use_container_width=True)

        st.write("")
        sel_col = st.selectbox("Inspect distribution", numeric_cols, key="desc_dist_col")
        c1, c2 = st.columns(2)
        with c1:
            fig = px.histogram(filtered_df, x=sel_col, nbins=30, marginal="box", color_discrete_sequence=[ACCENT])
            fig.update_layout(title=f"Distribution — {sel_col}")
            st.plotly_chart(style_fig(fig), use_container_width=True)
        with c2:
            fig2 = px.box(filtered_df, y=sel_col, color_discrete_sequence=[ACCENT_2], points="outliers")
            fig2.update_layout(title=f"Spread & outliers — {sel_col}")
            st.plotly_chart(style_fig(fig2), use_container_width=True)
    else:
        st.info("No numeric columns detected in this sheet.")

    if categorical_cols:
        st.markdown("##### Categorical breakdown")
        cat_col = st.selectbox("Column", categorical_cols, key="desc_cat_col")
        vc = filtered_df[cat_col].astype(str).value_counts().reset_index().head(15)
        vc.columns = [cat_col, "Count"]
        cc1, cc2 = st.columns([1, 1.2])
        with cc1:
            st.dataframe(vc, use_container_width=True, hide_index=True)
        with cc2:
            fig3 = px.bar(vc, x="Count", y=cat_col, orientation="h", color="Count", color_continuous_scale=[BG_RAISED, ACCENT])
            fig3.update_layout(yaxis=dict(autorange="reversed"), title=f"Top categories — {cat_col}", coloraxis_showscale=False)
            st.plotly_chart(style_fig(fig3), use_container_width=True)

# ---------------- Inferential Analysis ----------------
with tab_infer:
    st.caption("Significance testing — check whether observed differences or relationships are likely real or due to chance.")

    test_type = st.radio(
        "Test type",
        ["Compare 2 groups (t-test)", "Compare 3+ groups (ANOVA)", "Relationship between categories (chi-square)", "Relationship between numbers (correlation)"],
        horizontal=False,
    )

    alpha = st.select_slider("Significance level (α)", options=[0.01, 0.05, 0.10], value=0.05)

    st.markdown("---")

    if test_type == "Compare 2 groups (t-test)":
        if not numeric_cols or not categorical_cols:
            st.info("Need at least one numeric column and one categorical column.")
        else:
            cnum = st.selectbox("Numeric variable", numeric_cols, key="tt_num")
            ccat = st.selectbox("Group variable (must have exactly 2 groups)", categorical_cols, key="tt_cat")
            result = run_ttest(filtered_df, cnum, ccat)
            if result is None:
                st.warning("Selected group column doesn't have exactly 2 categories with enough data, or sample sizes are too small.")
            else:
                st.html(f"""
                <div class="panel">
                    <div style="font-weight:700;color:{TEXT_HI};margin-bottom:6px;">{result['test']}</div>
                    <div style="color:{TEXT_MED};margin-bottom:10px;">Comparing <b style="color:{TEXT_HI}">{result['groups']}</b></div>
                    {verdict_badge(result['p_value'], alpha)}
                    <div class="mono" style="margin-top:12px;color:{TEXT_MED};font-size:0.85rem;">
                        t = {result['statistic']:.4f} &nbsp;·&nbsp; n₁ = {result['n1']} &nbsp;·&nbsp; n₂ = {result['n2']}<br>
                        mean₁ = {result['mean1']:.3f} &nbsp;·&nbsp; mean₂ = {result['mean2']:.3f}
                    </div>
                </div>
                """)
                fig = px.box(filtered_df, x=ccat, y=cnum, color=ccat, color_discrete_sequence=PLOTLY_SERIES, points="all")
                fig.update_layout(title=f"{cnum} by {ccat}", showlegend=False)
                st.plotly_chart(style_fig(fig), use_container_width=True)

    elif test_type == "Compare 3+ groups (ANOVA)":
        if not numeric_cols or not categorical_cols:
            st.info("Need at least one numeric column and one categorical column.")
        else:
            cnum = st.selectbox("Numeric variable", numeric_cols, key="an_num")
            ccat = st.selectbox("Group variable (3+ groups)", categorical_cols, key="an_cat")
            result = run_anova(filtered_df, cnum, ccat)
            if result is None:
                st.warning("Selected group column doesn't have at least 3 categories with enough data.")
            else:
                st.html(f"""
                <div class="panel">
                    <div style="font-weight:700;color:{TEXT_HI};margin-bottom:6px;">{result['test']}</div>
                    <div style="color:{TEXT_MED};margin-bottom:10px;">Across <b style="color:{TEXT_HI}">{result['groups']}</b></div>
                    {verdict_badge(result['p_value'], alpha)}
                    <div class="mono" style="margin-top:12px;color:{TEXT_MED};font-size:0.85rem;">
                        F = {result['statistic']:.4f} &nbsp;·&nbsp; groups = {result['n_groups']}
                    </div>
                </div>
                """)
                fig = px.box(filtered_df, x=ccat, y=cnum, color=ccat, color_discrete_sequence=PLOTLY_SERIES, points="outliers")
                fig.update_layout(title=f"{cnum} by {ccat}", showlegend=False)
                st.plotly_chart(style_fig(fig), use_container_width=True)

    elif test_type == "Relationship between categories (chi-square)":
        if len(categorical_cols) < 2:
            st.info("Need at least two categorical columns.")
        else:
            ca = st.selectbox("Variable A", categorical_cols, key="chi_a")
            cb = st.selectbox("Variable B", [c for c in categorical_cols if c != ca], key="chi_b")
            result = run_chi_square(filtered_df, ca, cb)
            if result is None:
                st.warning("Both columns need at least 2 categories each with enough data.")
            else:
                st.html(f"""
                <div class="panel">
                    <div style="font-weight:700;color:{TEXT_HI};margin-bottom:6px;">{result['test']}</div>
                    <div style="color:{TEXT_MED};margin-bottom:10px;">Testing <b style="color:{TEXT_HI}">{result['groups']}</b></div>
                    {verdict_badge(result['p_value'], alpha)}
                    <div class="mono" style="margin-top:12px;color:{TEXT_MED};font-size:0.85rem;">
                        χ² = {result['statistic']:.4f} &nbsp;·&nbsp; dof = {result['dof']}
                    </div>
                </div>
                """)
                fig = px.imshow(result["table"], text_auto=True, color_continuous_scale=[BG_RAISED, ACCENT], aspect="auto")
                fig.update_layout(title=f"Contingency table — {ca} × {cb}")
                st.plotly_chart(style_fig(fig), use_container_width=True)

    elif test_type == "Relationship between numbers (correlation)":
        if len(numeric_cols) < 2:
            st.info("Need at least two numeric columns.")
        else:
            ca = st.selectbox("Variable A", numeric_cols, key="corr_a")
            cb = st.selectbox("Variable B", [c for c in numeric_cols if c != ca], key="corr_b")
            result = run_correlation_test(filtered_df, ca, cb)
            if result is None:
                st.warning("Not enough overlapping data between these two columns.")
            else:
                st.html(f"""
                <div class="panel">
                    <div style="font-weight:700;color:{TEXT_HI};margin-bottom:6px;">{result['test']}</div>
                    <div style="color:{TEXT_MED};margin-bottom:10px;">Between <b style="color:{TEXT_HI}">{result['groups']}</b></div>
                    {verdict_badge(result['p_value'], alpha)}
                    <div class="mono" style="margin-top:12px;color:{TEXT_MED};font-size:0.85rem;">
                        r = {result['statistic']:.4f} &nbsp;·&nbsp; n = {result['n']}
                    </div>
                </div>
                """)
                fig = px.scatter(filtered_df, x=ca, y=cb, trendline="ols", color_discrete_sequence=[ACCENT])
                fig.update_layout(title=f"{ca} vs {cb}")
                st.plotly_chart(style_fig(fig), use_container_width=True)

# ---------------- Correlations ----------------
with tab_corr:
    if len(numeric_cols) >= 2:
        corr = filtered_df[numeric_cols].corr()
        fig = px.imshow(corr, text_auto=".2f", color_continuous_scale="RdBu", zmin=-1, zmax=1, aspect="auto")
        fig.update_layout(title="Correlation matrix")
        st.plotly_chart(style_fig(fig, height=440), use_container_width=True)

        pairs = corr.where(~np.eye(len(corr), dtype=bool)).stack().reset_index()
        pairs.columns = ["Variable 1", "Variable 2", "Correlation"]
        pairs["abs_corr"] = pairs["Correlation"].abs()
        pairs = pairs.sort_values("abs_corr", ascending=False).drop_duplicates(subset="abs_corr")
        st.markdown("##### Strongest relationships")
        st.dataframe(pairs.head(10)[["Variable 1", "Variable 2", "Correlation"]].round(3), hide_index=True, use_container_width=True)
    else:
        st.info("Need at least two numeric columns to compute correlations.")

# ---------------- Chart Builder ----------------
with tab_explore:
    chart_type = st.selectbox("Chart type", ["Histogram", "Box plot", "Scatter plot", "Bar chart", "Line chart", "Pie chart"])

    if chart_type == "Histogram" and numeric_cols:
        col = st.selectbox("Column", numeric_cols, key="hist_col")
        color_col = st.selectbox("Color by (optional)", ["None"] + categorical_cols, key="hist_color")
        fig = px.histogram(filtered_df, x=col, color=None if color_col == "None" else color_col, nbins=40, marginal="box")
        st.plotly_chart(style_fig(fig), use_container_width=True)

    elif chart_type == "Box plot" and numeric_cols:
        col = st.selectbox("Numeric column", numeric_cols, key="box_col")
        group_col = st.selectbox("Group by (optional)", ["None"] + categorical_cols, key="box_group")
        fig = px.box(filtered_df, y=col, x=None if group_col == "None" else group_col,
                     color=None if group_col == "None" else group_col, points="outliers")
        st.plotly_chart(style_fig(fig), use_container_width=True)

    elif chart_type == "Scatter plot" and len(numeric_cols) >= 2:
        x_col = st.selectbox("X axis", numeric_cols, key="scatter_x")
        y_col = st.selectbox("Y axis", [c for c in numeric_cols if c != x_col], key="scatter_y")
        color_col = st.selectbox("Color by (optional)", ["None"] + categorical_cols, key="scatter_color")
        trend = st.checkbox("Show trendline")
        fig = px.scatter(filtered_df, x=x_col, y=y_col, color=None if color_col == "None" else color_col,
                          trendline="ols" if trend else None)
        st.plotly_chart(style_fig(fig), use_container_width=True)

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
        fig = px.bar(data, x=cat_col, y=y_field, color=y_field, color_continuous_scale=[BG_RAISED, ACCENT])
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(style_fig(fig), use_container_width=True)

    elif chart_type == "Line chart" and datetime_cols and numeric_cols:
        date_col = st.selectbox("Date column", datetime_cols, key="line_date")
        value_col = st.selectbox("Value column", numeric_cols, key="line_val")
        line_data = filtered_df.dropna(subset=[date_col]).sort_values(date_col)
        fig = px.line(line_data, x=date_col, y=value_col, color_discrete_sequence=[ACCENT_2])
        st.plotly_chart(style_fig(fig), use_container_width=True)

    elif chart_type == "Pie chart" and categorical_cols:
        cat_col = st.selectbox("Category column", categorical_cols, key="pie_cat")
        data = filtered_df[cat_col].astype(str).value_counts().reset_index().head(10)
        data.columns = [cat_col, "Count"]
        fig = px.pie(data, names=cat_col, values="Count", hole=0.45)
        st.plotly_chart(style_fig(fig), use_container_width=True)

    else:
        st.info("Not enough of the right column types are available for this chart. Try another chart type.")

# ---------------- Raw Data ----------------
with tab_data:
    st.dataframe(filtered_df, use_container_width=True)
    csv = filtered_df.to_csv(index=False).encode("utf-8")
    st.download_button("⬇ Download filtered data as CSV", csv, "filtered_data.csv", "text/csv")
