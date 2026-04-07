"""
App Adoption Dashboard — Streamlit Web Interface
Upload your Excel file and view adoption metrics instantly.
Styled to match Q4 Scheme Dashboard.
"""

import io
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ============================================================================
# SECTION 1 — PAGE CONFIG & CONSTANTS
# ============================================================================

st.set_page_config(
    page_title="App Adoption Dashboard",
    page_icon="📊",
    layout="wide",
)

VALID_STATUSES = ["Installed", "Not Installed", "Uninstalled", "Tech Issue", "Not Working"]
REQUIRED_COLUMNS = ["Account Sf ID", "Name of the Dealer", "Status"]

STATUS_CONFIG = [
    {"status": "Installed",     "color": "#10b981", "color_light": "#d1fae5", "icon": "✅"},
    {"status": "Not Installed", "color": "#94a3b8", "color_light": "#f1f5f9", "icon": "⬜"},
    {"status": "Uninstalled",   "color": "#f59e0b", "color_light": "#fef3c7", "icon": "🟡"},
    {"status": "Tech Issue",    "color": "#f97316", "color_light": "#ffedd5", "icon": "🟠"},
    {"status": "Not Working",   "color": "#ef4444", "color_light": "#fee2e2", "icon": "🔴"},
]

STATUS_COLORS = {s["status"]: s["color"] for s in STATUS_CONFIG}
STATUS_COLORS_LIGHT = {s["status"]: s["color_light"] for s in STATUS_CONFIG}


# ============================================================================
# SECTION 2 — CUSTOM CSS (Q4 style)
# ============================================================================

def inject_custom_css():
    """Inject Q4-style CSS into the Streamlit app."""
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        .stApp { background-color: #f8fafc; }
        .block-container { padding-top: 1rem !important; }

        /* ── Header bar ── */
        .header-bar {
            display: flex; justify-content: space-between; align-items: center;
            padding: 1rem 1.5rem;
            background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 50%, #1e40af 100%);
            border-radius: 0.75rem; margin-bottom: 0.75rem;
            color: #ffffff; box-shadow: 0 4px 12px rgba(15,23,42,0.25);
        }
        .header-bar h1 { margin: 0; font-size: 1.5rem; font-weight: 800; color: #ffffff; letter-spacing: -0.01em; }
        .header-bar .header-right { display: flex; align-items: center; gap: 0.75rem; }
        .header-bar .header-subtitle { font-size: 0.85rem; font-weight: 500; color: #94a3b8; }
        .header-badge {
            font-size: 0.85rem; font-weight: 600; color: #ffffff;
            background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.3);
            border-radius: 1rem; padding: 0.3rem 1rem;
        }

        /* ── Section elements ── */
        .section-divider { border: none; border-top: 1px solid #e2e8f0; margin: 0.5rem 0 1rem 0; }
        .section-title {
            font-size: 0.8rem; font-weight: 700; text-transform: uppercase;
            letter-spacing: 0.08em; color: #64748b; margin-bottom: 0.5rem;
        }

        /* ── Filter panel ── */
        .filter-panel {
            background: #ffffff; border: 1px solid #e2e8f0; border-radius: 0.75rem;
            padding: 0.9rem 1.2rem 0.5rem 1.2rem; margin-bottom: 1rem;
            box-shadow: 0 1px 4px rgba(0,0,0,0.04);
        }
        .filter-panel-title {
            font-size: 0.75rem; font-weight: 700; text-transform: uppercase;
            letter-spacing: 0.1em; color: #64748b; margin-bottom: 0.4rem;
            display: flex; align-items: center; gap: 0.4rem;
        }

        /* ── KPI cards ── */
        .kpi-card {
            background: #ffffff; border: 1px solid #e2e8f0;
            border-top: 3px solid #2563eb; border-radius: 0.6rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
            padding: 1rem 0.8rem; text-align: center;
            transition: box-shadow 0.2s, transform 0.2s;
        }
        .kpi-card:hover { box-shadow: 0 4px 16px rgba(37,99,235,0.12); transform: translateY(-1px); }
        .kpi-label {
            font-size: 0.72rem; font-weight: 700; text-transform: uppercase;
            letter-spacing: 0.06em; color: #64748b; margin-bottom: 0.35rem;
        }
        .kpi-value { font-size: 1.45rem; font-weight: 800; color: #0f172a; letter-spacing: -0.02em; }

        /* ── Status cards (like slab cards) ── */
        .status-card {
            background: #ffffff; border-radius: 0.6rem;
            box-shadow: 0 2px 8px rgba(0,0,0,0.07); padding: 0;
            margin-bottom: 0.5rem; overflow: hidden;
            border: 1px solid #e2e8f0; transition: box-shadow 0.2s, transform 0.2s;
        }
        .status-card:hover { box-shadow: 0 4px 16px rgba(0,0,0,0.12); transform: translateY(-1px); }
        .status-color-band { height: 5px; width: 100%; }
        .status-card-body { padding: 0.9rem 1rem; }
        .status-count { font-size: 2rem; font-weight: 800; line-height: 1.1; }
        .status-pct { font-size: 0.75rem; font-weight: 600; color: #64748b; margin-left: 0.3rem; }
        .status-label {
            font-size: 0.8rem; font-weight: 700; text-transform: uppercase;
            letter-spacing: 0.04em; color: #1e293b; margin-top: 0.25rem;
        }

        /* ── Table styling ── */
        .stDataFrame [data-testid="stDataFrameResizable"] {
            border: 1px solid #cbd5e1; border-radius: 0.6rem; overflow: hidden;
        }
        .stDataFrame thead tr th {
            background-color: #1e293b !important; color: #ffffff !important;
            font-weight: 700 !important; font-size: 0.82rem !important;
        }
        [data-testid="stDataFrame"] [role="columnheader"],
        [data-testid="stDataFrame"] [data-testid="glide-cell"] { font-weight: 700 !important; }
        [data-testid="stDataFrame"] .gdg-header { font-weight: 800 !important; }

        /* ── Tabs ── */
        .stTabs [data-baseweb="tab-list"] {
            gap: 0; background: #ffffff; border-radius: 0.6rem 0.6rem 0 0;
            border-bottom: 2px solid #e2e8f0; padding: 0 0.5rem;
        }
        .stTabs [data-baseweb="tab"] {
            padding: 0.7rem 1.5rem; font-size: 0.88rem; font-weight: 600; color: #64748b;
            border-bottom: 3px solid transparent; transition: color 0.2s, border-color 0.2s;
        }
        .stTabs [data-baseweb="tab"]:hover { color: #334155; }
        .stTabs [data-baseweb="tab"][aria-selected="true"] {
            font-weight: 700; color: #1e40af; border-bottom: 3px solid #2563eb;
        }

        /* ── Selectbox ── */
        .stSelectbox label { font-weight: 600 !important; color: #334155 !important; font-size: 0.8rem !important; }
        .stSelectbox [data-baseweb="select"] { border-color: #cbd5e1 !important; border-radius: 0.4rem !important; }

        /* ── Subheaders ── */
        h2, h3 { color: #0f172a !important; font-weight: 700 !important; }
        .stCaption, [data-testid="stCaptionContainer"] { color: #94a3b8 !important; font-size: 0.75rem !important; }

        /* ── Hide Streamlit chrome ── */
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        </style>
        """,
        unsafe_allow_html=True,
    )


# ============================================================================
# SECTION 3 — HELPERS
# ============================================================================

def _opts(series):
    """Build filter options: ['All'] + sorted unique non-blank values."""
    unique_vals = series[series.astype(str).str.strip() != ""].unique()
    return ["All"] + sorted(str(v) for v in unique_vals if str(v).strip())


def render_cascading_filters(df, key, filter_fields):
    """Render cascading dropdown filters and return filtered DataFrame."""
    filter_fields = [f for f in filter_fields if f in df.columns]
    if not filter_fields:
        return df

    cols = st.columns(len(filter_fields))
    filtered = df.copy()

    for i, field in enumerate(filter_fields):
        with cols[i]:
            options = _opts(filtered[field])
            selected = st.selectbox(field, options, key=f"{key}_{field}")
            if selected != "All":
                filtered = filtered[filtered[field] == selected]

    return filtered


def _kpi_card(label, value, color=None):
    """Return HTML for a KPI card."""
    border_color = color or "#2563eb"
    value_color = color or "#0f172a"
    return f"""
    <div class="kpi-card" style="border-top-color: {border_color};">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value" style="color: {value_color};">{value}</div>
    </div>
    """


def _status_card(status, count, total, color):
    """Return HTML for a status distribution card."""
    pct = f"{count / total * 100:.0f}%" if total > 0 else "0%"
    return f"""
    <div class="status-card">
        <div class="status-color-band" style="background: {color};"></div>
        <div class="status-card-body">
            <div class="status-count" style="color: {color};">
                {count}<span class="status-pct">({pct})</span>
            </div>
            <div class="status-label">{status}</div>
        </div>
    </div>
    """


# ============================================================================
# SECTION 4 — DATA LOADING & VALIDATION
# ============================================================================

@st.cache_data
def load_and_clean(uploaded_bytes):
    """Load Excel bytes into a cleaned DataFrame."""
    df = pd.read_excel(io.BytesIO(uploaded_bytes), engine="openpyxl", dtype=str)
    df.columns = [c.strip() for c in df.columns]

    int_cols = ["# orders (total)", "# orders (via. app.)"]
    for col in int_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).astype(int)

    numeric_cols = ["Order quantity (total)", "Quantity (via. app.)"]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(
                df[col].astype(str).str.strip().replace("-", "0"), errors="coerce"
            ).fillna(0)

    float_cols = ["% orders via. app.", "% order quantity (via. app.)"]
    for col in float_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def validate_data(df):
    """Validate the DataFrame. Returns (is_valid, errors)."""
    errors = []
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            errors.append(f"Missing required column: **{col}**")
    if errors:
        return False, errors

    if "Account Sf ID" in df.columns:
        blank_ids = df["Account Sf ID"].isna().sum()
        if blank_ids > 0:
            errors.append(f"**Account Sf ID** has {blank_ids} blank value(s)")

    if "Status" in df.columns:
        blank_status = df["Status"].isna().sum()
        if blank_status > 0:
            errors.append(f"**Status** has {blank_status} blank value(s)")

    return len(errors) == 0, errors


# ============================================================================
# SECTION 5 — EXCEL REPORT GENERATION
# ============================================================================

def generate_excel_report(df):
    """Generate the Excel report and return bytes."""
    total = len(df)
    status_counts = df["Status"].value_counts()
    today = datetime.now().strftime("%Y-%m-%d")

    wb = Workbook()

    def style_header(ws):
        hf = Font(bold=True, color="FFFFFF", size=11)
        hfill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        for cell in ws[1]:
            cell.font = hf
            cell.fill = hfill
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
        ws.freeze_panes = "A2"

    def auto_width(ws):
        for ci, cc in enumerate(ws.columns, 1):
            ml = max((len(str(c.value or "")) for c in cc), default=8)
            ws.column_dimensions[get_column_letter(ci)].width = min(ml + 3, 45)

    # Sheet 1: Summary
    ws = wb.active
    ws.title = "Summary"
    ws.append(["Metric", "Value"])
    bold = Font(bold=True, size=11)
    installed = int(status_counts.get("Installed", 0))
    not_installed = int(status_counts.get("Not Installed", 0))
    uninstalled = int(status_counts.get("Uninstalled", 0))
    tech_issue = int(status_counts.get("Tech Issue", 0))
    not_working = int(status_counts.get("Not Working", 0))
    adoption_rate = round(installed / total * 100, 1) if total > 0 else 0
    ever_adopted = installed + uninstalled + tech_issue + not_working
    ever_adopted_rate = round(ever_adopted / total * 100, 1) if total > 0 else 0
    for label, value in [
        ("Report Date", today), ("", ""),
        ("Total Dealers", total), ("Installed", installed),
        ("Not Installed", not_installed), ("Uninstalled", uninstalled),
        ("Tech Issue", tech_issue), ("Not Working", not_working),
        ("", ""), ("Currently Installed (%)", f"{adoption_rate}%"),
        ("Ever Adopted (%)", f"{ever_adopted_rate}%"),
    ]:
        ws.append([label, value])
        ws.cell(row=ws.max_row, column=1).font = bold
    style_header(ws)
    auto_width(ws)

    # Sheet 2: Zone-wise
    ws2 = wb.create_sheet("Zone-wise Adoption")
    ws2.append(["Zone", "State", "Total", "Installed", "Not Installed", "Uninstalled",
                 "Tech Issue", "Not Working", "Adoption Rate (%)"])
    if "Zone" in df.columns and "State" in df.columns:
        zs = df.groupby(["Zone", "State"]).agg(
            total=("Account Sf ID", "count"),
            inst=("Status", lambda x: (x == "Installed").sum()),
            ni=("Status", lambda x: (x == "Not Installed").sum()),
            un=("Status", lambda x: (x == "Uninstalled").sum()),
            ti=("Status", lambda x: (x == "Tech Issue").sum()),
            nw=("Status", lambda x: (x == "Not Working").sum()),
        ).reset_index()
        zs["rate"] = (zs["inst"] / zs["total"] * 100).round(1)
        for _, r in zs.sort_values(["Zone", "rate"], ascending=[True, False]).iterrows():
            ws2.append([r["Zone"], r["State"], int(r["total"]), int(r["inst"]),
                        int(r["ni"]), int(r["un"]), int(r["ti"]), int(r["nw"]), f"{r['rate']}%"])
    style_header(ws2)
    auto_width(ws2)

    # Sheet 3: Distributor-wise
    ws3 = wb.create_sheet("Distributor-wise Adoption")
    ws3.append(["Distributor Name", "Total", "Installed", "Not Installed", "Uninstalled",
                 "Other", "Adoption Rate (%)"])
    if "Distributor Name" in df.columns:
        dist = df.groupby("Distributor Name").agg(
            total=("Account Sf ID", "count"),
            inst=("Status", lambda x: (x == "Installed").sum()),
            ni=("Status", lambda x: (x == "Not Installed").sum()),
            un=("Status", lambda x: (x == "Uninstalled").sum()),
            oth=("Status", lambda x: x.isin(["Tech Issue", "Not Working"]).sum()),
        ).reset_index()
        dist["rate"] = (dist["inst"] / dist["total"] * 100).round(1)
        for _, r in dist.sort_values("rate", ascending=False).iterrows():
            ws3.append([r["Distributor Name"], int(r["total"]), int(r["inst"]),
                        int(r["ni"]), int(r["un"]), int(r["oth"]), f"{r['rate']}%"])
    style_header(ws3)
    auto_width(ws3)

    # Sheet 4: Dealer Detail
    ws4 = wb.create_sheet("Dealer Detail")
    detail_cols = ["Account Sf ID", "Name of the Dealer", "State", "Zone",
                   "Distributor Name", "Account Owner As per SF", "Mobile No.",
                   "# orders (total)", "# orders (via. app.)", "Status"]
    avail = [c for c in detail_cols if c in df.columns]
    ws4.append(avail)
    status_order = {"Installed": 0, "Uninstalled": 1, "Tech Issue": 2, "Not Working": 3, "Not Installed": 4}
    sdf = df.copy()
    sdf["_s"] = sdf["Status"].map(status_order).fillna(5)
    sdf = sdf.sort_values(["_s", "Name of the Dealer"])
    for _, row in sdf.iterrows():
        ws4.append([row.get(c, "") if pd.notna(row.get(c)) else "" for c in avail])
    style_header(ws4)
    auto_width(ws4)

    # Sheet 5: App Usage
    ws5 = wb.create_sheet("App Usage (Orders)")
    ws5.append(["Account Sf ID", "Name of the Dealer", "Zone", "Distributor Name",
                "# Orders (Total)", "# Orders (via App)", "% Orders via App", "Status"])
    if "# orders (via. app.)" in df.columns:
        app_u = df[df["# orders (via. app.)"] > 0].sort_values("# orders (via. app.)", ascending=False)
        for _, r in app_u.iterrows():
            pct = f"{r['% orders via. app.']:.1f}%" if pd.notna(r.get("% orders via. app.")) else ""
            ws5.append([r.get("Account Sf ID", ""), r.get("Name of the Dealer", ""),
                        r.get("Zone", ""), r.get("Distributor Name", ""),
                        r.get("# orders (total)", 0), r.get("# orders (via. app.)", 0),
                        pct, r.get("Status", "")])
    style_header(ws5)
    auto_width(ws5)

    # Sheet 6: At Risk
    ws6 = wb.create_sheet("At Risk - Action Needed")
    ws6.append(["Account Sf ID", "Name of the Dealer", "State", "Zone",
                "Distributor Name", "Account Owner As per SF", "Mobile No.", "Status", "Action Needed"])
    risk_map = {"Uninstalled": "Re-engage: app was uninstalled",
                "Tech Issue": "Resolve technical issue", "Not Working": "Fix: app not working"}
    risk = df[df["Status"].isin(risk_map.keys())].copy()
    risk["Action Needed"] = risk["Status"].map(risk_map)
    risk = risk.sort_values("Status")
    sfills = {
        "Not Working": PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid"),
        "Tech Issue": PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid"),
        "Uninstalled": PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid"),
    }
    for _, r in risk.iterrows():
        ws6.append([r.get("Account Sf ID", ""), r.get("Name of the Dealer", ""),
                     r.get("State", ""), r.get("Zone", ""), r.get("Distributor Name", ""),
                     r.get("Account Owner As per SF", ""), r.get("Mobile No.", ""),
                     r.get("Status", ""), r.get("Action Needed", "")])
    for ri in range(2, ws6.max_row + 1):
        sv = ws6.cell(row=ri, column=8).value
        if sv in sfills:
            for ci in range(1, 10):
                ws6.cell(row=ri, column=ci).fill = sfills[sv]
    style_header(ws6)
    auto_width(ws6)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================================
# SECTION 6 — STYLED DATAFRAME HELPERS
# ============================================================================

def _highlight_by_status(row):
    """Apply light status-based background color to each row."""
    status = row.get("Status", "")
    bg = STATUS_COLORS_LIGHT.get(status, "")
    if bg:
        return [f"background-color: {bg}"] * len(row)
    return [""] * len(row)


def _highlight_adoption_rate(row):
    """Color rows by adoption rate: green ≥50%, yellow ≥25%, red <25%."""
    rate = row.get("Adoption %", 0)
    if rate >= 50:
        return ["background-color: #d1fae5"] * len(row)
    if rate >= 25:
        return ["background-color: #fef3c7"] * len(row)
    return ["background-color: #fee2e2"] * len(row)


def _highlight_risk(row):
    """Color at-risk rows by severity."""
    status = row.get("Status", "")
    colors = {"Not Working": "#fee2e2", "Tech Issue": "#ffedd5", "Uninstalled": "#fef3c7"}
    bg = colors.get(status, "")
    if bg:
        return [f"background-color: {bg}"] * len(row)
    return [""] * len(row)


def _bold_columns(*col_names):
    """Return a styler function that bolds specified columns."""
    def styler(col):
        if col.name in col_names:
            return ["font-weight: 700"] * len(col)
        return [""] * len(col)
    return styler


# ============================================================================
# SECTION 7 — MAIN UI
# ============================================================================

inject_custom_css()

# ── Header Bar ──
st.markdown(
    """
    <div class="header-bar">
        <h1>App Adoption Dashboard</h1>
        <div class="header-right">
            <span class="header-subtitle">Dealer App Tracking</span>
            <span class="header-badge">FY 2025-26</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Sidebar: File Upload ──
with st.sidebar:
    st.header("Upload Data")
    uploaded_file = st.file_uploader(
        "Upload your Excel file",
        type=["xlsx", "xls"],
        help="Upload the app adoption Excel file with dealer data",
    )
    if uploaded_file:
        st.success(f"Loaded: {uploaded_file.name}")
    st.markdown("---")
    st.markdown("**Expected columns:**")
    st.markdown("- Account Sf ID *(required)*\n- Name of the Dealer *(required)*\n- Status *(required)*\n- State, Zone, Distributor Name\n- Order data columns")

# ── Guard: no file uploaded ──
if uploaded_file is None:
    st.info("👈 Upload your Excel file in the sidebar to get started.")
    st.stop()

# ── Load & Validate ──
raw_bytes = uploaded_file.getvalue()
df = load_and_clean(raw_bytes)

is_valid, errors = validate_data(df)
if not is_valid:
    st.error("**Validation errors:**")
    for e in errors:
        st.write(f"- {e}")
    st.stop()

# ── Data source caption ──
st.caption(f"Data source: `{uploaded_file.name}` — {len(df)} dealers loaded")

# ── Global Cascading Filters ──
st.markdown(
    '<div class="filter-panel"><div class="filter-panel-title">&#9660; Filters</div>',
    unsafe_allow_html=True,
)
global_filters = ["Status", "Zone", "State", "Distributor Name"]
filtered_df = render_cascading_filters(df, key="global", filter_fields=global_filters)
st.markdown('</div>', unsafe_allow_html=True)

# ── KPI Row ──
total = len(filtered_df)
status_counts = filtered_df["Status"].value_counts()
installed = int(status_counts.get("Installed", 0))
not_installed = int(status_counts.get("Not Installed", 0))
uninstalled = int(status_counts.get("Uninstalled", 0))
tech_issue = int(status_counts.get("Tech Issue", 0))
not_working = int(status_counts.get("Not Working", 0))
adoption_rate = round(installed / total * 100, 1) if total > 0 else 0
ever_adopted = installed + uninstalled + tech_issue + not_working
ever_adopted_rate = round(ever_adopted / total * 100, 1) if total > 0 else 0
orders_via_app = int(filtered_df["# orders (via. app.)"].sum()) if "# orders (via. app.)" in filtered_df.columns else 0

st.markdown('<div class="section-title">Key Metrics</div>', unsafe_allow_html=True)
kpi_cols = st.columns(5)
kpis = [
    ("Total Dealers", total, None),
    ("Currently Installed", installed, "#10b981"),
    ("Adoption Rate", f"{adoption_rate}%", "#2563eb"),
    ("Ever Adopted", f"{ever_adopted_rate}%", "#6366f1"),
    ("Orders via App", orders_via_app, "#f59e0b"),
]
for col, (label, value, color) in zip(kpi_cols, kpis):
    with col:
        st.markdown(_kpi_card(label, value, color), unsafe_allow_html=True)

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

# ── Status Distribution Cards ──
st.markdown('<div class="section-title">Status Distribution</div>', unsafe_allow_html=True)
card_cols = st.columns(len(STATUS_CONFIG))
for col, cfg in zip(card_cols, STATUS_CONFIG):
    count = int(status_counts.get(cfg["status"], 0))
    with col:
        st.markdown(_status_card(cfg["status"], count, total, cfg["color"]), unsafe_allow_html=True)

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

# ── Tabs ──
tab_summary, tab_details, tab_orders, tab_risk = st.tabs([
    "📊 Summary", "🔍 Dealer Details", "📱 App Usage (Orders)", "⚠️ At Risk"
])

# ── Tab 1: Summary ──
with tab_summary:
    # Zone-wise
    if "Zone" in filtered_df.columns:
        st.subheader("Zone-wise Adoption")
        zone_df = filtered_df.groupby("Zone").agg(
            Total=("Account Sf ID", "count"),
            Installed=("Status", lambda x: (x == "Installed").sum()),
            Not_Installed=("Status", lambda x: (x == "Not Installed").sum()),
            Uninstalled=("Status", lambda x: (x == "Uninstalled").sum()),
            Other=("Status", lambda x: x.isin(["Tech Issue", "Not Working"]).sum()),
        ).reset_index()
        zone_df["Adoption %"] = (zone_df["Installed"] / zone_df["Total"] * 100).round(1)
        zone_df = zone_df.sort_values("Adoption %", ascending=False)
        zone_df.columns = ["Zone", "Total", "Installed", "Not Installed", "Uninstalled", "Tech Issue / NW", "Adoption %"]
        styled_zone = (
            zone_df.style
            .apply(_highlight_adoption_rate, axis=1)
            .apply(_bold_columns("Zone", "Total", "Adoption %"), axis=0)
        )
        st.dataframe(styled_zone, use_container_width=True, hide_index=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    # Distributor-wise
    if "Distributor Name" in filtered_df.columns:
        st.subheader("Distributor-wise Adoption")
        dist_df = filtered_df.groupby("Distributor Name").agg(
            Total=("Account Sf ID", "count"),
            Installed=("Status", lambda x: (x == "Installed").sum()),
            Not_Installed=("Status", lambda x: (x == "Not Installed").sum()),
            Uninstalled=("Status", lambda x: (x == "Uninstalled").sum()),
            Other=("Status", lambda x: x.isin(["Tech Issue", "Not Working"]).sum()),
        ).reset_index()
        dist_df["Adoption %"] = (dist_df["Installed"] / dist_df["Total"] * 100).round(1)
        dist_df = dist_df.sort_values("Adoption %", ascending=False)
        dist_df.columns = ["Distributor Name", "Total", "Installed", "Not Installed", "Uninstalled", "Tech Issue / NW", "Adoption %"]
        styled_dist = (
            dist_df.style
            .apply(_highlight_adoption_rate, axis=1)
            .apply(_bold_columns("Distributor Name", "Total", "Adoption %"), axis=0)
        )
        st.dataframe(styled_dist, use_container_width=True, hide_index=True, height=400)

# ── Tab 2: Dealer Details ──
with tab_details:
    # Per-tab cascading filters
    detail_filter_fields = ["Status", "Zone", "State", "Distributor Name"]
    detail_filters_available = [f for f in detail_filter_fields if f in filtered_df.columns]

    detail_cols_ui = st.columns(len(detail_filters_available) + 1)
    detail_filtered = filtered_df.copy()

    for i, field in enumerate(detail_filters_available):
        with detail_cols_ui[i]:
            options = _opts(detail_filtered[field])
            selected = st.selectbox(field, options, key=f"detail_{field}")
            if selected != "All":
                detail_filtered = detail_filtered[detail_filtered[field] == selected]

    with detail_cols_ui[-1]:
        search = st.text_input("Search Dealer Name", key="detail_search")
        if search:
            detail_filtered = detail_filtered[
                detail_filtered["Name of the Dealer"].fillna("").str.contains(search, case=False, na=False)
            ]

    st.subheader(f"Dealer Details ({len(detail_filtered)} records)")

    display_cols = ["Account Sf ID", "Name of the Dealer", "State", "Zone",
                    "Distributor Name", "Account Owner As per SF", "Mobile No.",
                    "# orders (total)", "# orders (via. app.)", "Status"]
    avail_cols = [c for c in display_cols if c in detail_filtered.columns]

    styled_detail = (
        detail_filtered[avail_cols].style
        .apply(_highlight_by_status, axis=1)
        .apply(_bold_columns("Name of the Dealer", "Status"), axis=0)
    )
    st.dataframe(styled_detail, hide_index=True, use_container_width=True, height=500)

# ── Tab 3: App Usage (Orders) ──
with tab_orders:
    if "# orders (via. app.)" in filtered_df.columns:
        app_users = filtered_df[filtered_df["# orders (via. app.)"] > 0].sort_values(
            "# orders (via. app.)", ascending=False
        )
        if len(app_users) == 0:
            st.info("No dealers have placed orders via the app yet.")
        else:
            # KPI cards for orders
            total_orders = int(app_users["# orders (total)"].sum())
            total_app_orders = int(app_users["# orders (via. app.)"].sum())
            app_pct = round(total_app_orders / total_orders * 100, 1) if total_orders > 0 else 0

            st.markdown('<div class="section-title">App Order Metrics</div>', unsafe_allow_html=True)
            ocols = st.columns(4)
            for col, (lbl, val, clr) in zip(ocols, [
                ("Dealers Using App", len(app_users), "#10b981"),
                ("Total Orders", total_orders, "#2563eb"),
                ("Orders via App", total_app_orders, "#6366f1"),
                ("% via App", f"{app_pct}%", "#f59e0b"),
            ]):
                with col:
                    st.markdown(_kpi_card(lbl, val, clr), unsafe_allow_html=True)

            st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
            st.subheader(f"Dealers with App Orders ({len(app_users)})")

            order_cols = ["Account Sf ID", "Name of the Dealer", "Zone", "Distributor Name",
                          "# orders (total)", "# orders (via. app.)", "% orders via. app.", "Status"]
            avail_order = [c for c in order_cols if c in app_users.columns]
            styled_orders = (
                app_users[avail_order].style
                .apply(_highlight_by_status, axis=1)
                .apply(_bold_columns("Name of the Dealer", "# orders (via. app.)"), axis=0)
            )
            st.dataframe(styled_orders, hide_index=True, use_container_width=True, height=400)
    else:
        st.info("No order data columns found.")

# ── Tab 4: At Risk ──
with tab_risk:
    risk_map = {
        "Not Working": "Fix: app not working",
        "Tech Issue": "Resolve technical issue",
        "Uninstalled": "Re-engage: app was uninstalled",
    }
    risk_df = filtered_df[filtered_df["Status"].isin(risk_map.keys())].copy()
    risk_df["Action Needed"] = risk_df["Status"].map(risk_map)

    # KPI cards
    st.markdown('<div class="section-title">Action Required</div>', unsafe_allow_html=True)
    rcols = st.columns(4)
    r_not_working = int((risk_df["Status"] == "Not Working").sum())
    r_tech = int((risk_df["Status"] == "Tech Issue").sum())
    r_uninstalled = int((risk_df["Status"] == "Uninstalled").sum())
    for col, (lbl, val, clr) in zip(rcols, [
        ("Total At Risk", len(risk_df), "#ef4444"),
        ("Not Working", r_not_working, "#ef4444"),
        ("Tech Issue", r_tech, "#f97316"),
        ("Uninstalled", r_uninstalled, "#f59e0b"),
    ]):
        with col:
            st.markdown(_kpi_card(lbl, val, clr), unsafe_allow_html=True)

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    if len(risk_df) > 0:
        st.subheader(f"Dealers Needing Action ({len(risk_df)})")
        risk_cols = ["Account Sf ID", "Name of the Dealer", "State", "Zone",
                     "Distributor Name", "Account Owner As per SF", "Mobile No.",
                     "Status", "Action Needed"]
        avail_risk = [c for c in risk_cols if c in risk_df.columns]
        styled_risk = (
            risk_df[avail_risk].sort_values("Status").style
            .apply(_highlight_risk, axis=1)
            .apply(_bold_columns("Name of the Dealer", "Status", "Action Needed"), axis=0)
        )
        st.dataframe(styled_risk, hide_index=True, use_container_width=True, height=500)
    else:
        st.success("No at-risk dealers!")

# ── Download Report ──
st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
st.markdown('<div class="section-title">Export</div>', unsafe_allow_html=True)

report_bytes = generate_excel_report(df)
st.download_button(
    label="📥 Download Excel Report",
    data=report_bytes,
    file_name=f"adoption_report_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
