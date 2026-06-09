"""
VendorLens — Vendor Payment Dashboard
======================================
CEO-grade AP dashboard. Upload your JD Edwards Excel export to refresh.
New in v2: Due-Next-Week table, Supplier Lookup module, clickable supplier
drill-down in the main invoice table, and a redesigned dark UI.
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
from datetime import datetime, timedelta

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="VendorLens · Payment Dashboard",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL CSS — redesigned palette: deep navy base, vivid accent rails
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

/* ── Base ── */
html,body,[class*="css"]{font-family:'Inter',sans-serif}
.main{background:#070b14}
.block-container{padding:1.8rem 2rem 4rem;max-width:1700px}

/* ── Sidebar ── */
[data-testid="stSidebar"]{
  background:linear-gradient(180deg,#0d1117 0%,#0d1117 100%)!important;
  border-right:1px solid #1e2533!important}
[data-testid="stSidebar"] .stMarkdown h3{
  color:#94a3b8;font-size:.65rem;letter-spacing:.14em;
  text-transform:uppercase;font-weight:600}

/* ── KPI cards ── */
.kpi-card{
  background:#0d1117;
  border:1px solid #1e2533;
  border-radius:14px;
  padding:1.25rem 1.5rem;
  position:relative;overflow:hidden;
  transition:transform .2s ease,border-color .2s ease,box-shadow .2s ease}
.kpi-card:hover{transform:translateY(-3px);border-color:#334155;
  box-shadow:0 8px 30px rgba(0,0,0,.4)}
.kpi-card::before{content:"";position:absolute;top:0;left:0;right:0;
  height:3px;border-radius:14px 14px 0 0}
.kpi-card.red::before   {background:linear-gradient(90deg,#f43f5e,#e11d48)}
.kpi-card.amber::before {background:linear-gradient(90deg,#f59e0b,#d97706)}
.kpi-card.blue::before  {background:linear-gradient(90deg,#3b82f6,#2563eb)}
.kpi-card.green::before {background:linear-gradient(90deg,#10b981,#059669)}
.kpi-card.teal::before  {background:linear-gradient(90deg,#06b6d4,#0891b2)}
.kpi-card.purple::before{background:linear-gradient(90deg,#8b5cf6,#7c3aed)}
.kpi-icon{font-size:1.4rem;margin-bottom:.5rem;opacity:.7}
.kpi-label{font-size:.62rem;color:#64748b;text-transform:uppercase;
  letter-spacing:.12em;font-weight:600;margin-bottom:.35rem}
.kpi-value{font-size:1.65rem;font-weight:700;color:#f8fafc;
  font-family:'JetBrains Mono',monospace;line-height:1.1;letter-spacing:-.02em}
.kpi-sub{font-size:.7rem;color:#64748b;margin-top:.3rem}
.kpi-badge{display:inline-block;padding:2px 8px;border-radius:20px;
  font-size:.62rem;font-weight:600;margin-top:.4rem}
.badge-danger{background:#4c0519;color:#fda4af}
.badge-warn  {background:#451a03;color:#fcd34d}
.badge-info  {background:#0c1a3d;color:#93c5fd}
.badge-ok    {background:#022c22;color:#6ee7b7}

/* ── Section headers ── */
.section-hd{
  font-size:.95rem;font-weight:600;color:#e2e8f0;
  padding:.6rem 0 .4rem;border-bottom:1px solid #1e2533;margin-bottom:1rem;
  display:flex;align-items:center;gap:.5rem}
.section-hd .dot{width:8px;height:8px;border-radius:50%;flex-shrink:0}

/* ── Alert banner ── */
.alert-box{
  background:linear-gradient(135deg,#4c0519 0%,#3f0f1e 100%);
  border:1px solid #9f1239;border-radius:12px;
  padding:1rem 1.4rem;color:#fecdd3;font-size:.88rem;margin-bottom:1.4rem;
  display:flex;align-items:flex-start;gap:.75rem}
.alert-box .alert-icon{font-size:1.2rem;flex-shrink:0;margin-top:.05rem}
.alert-box strong{color:#fb7185}

/* ── Due-next-week panel ── */
.nw-panel{
  background:#0d1117;border:1px solid #1e3a5f;
  border-radius:14px;padding:1.2rem 1.4rem;margin-bottom:1.4rem}
.nw-title{font-size:.8rem;font-weight:600;color:#93c5fd;
  text-transform:uppercase;letter-spacing:.1em;margin-bottom:.8rem;
  display:flex;align-items:center;gap:.5rem}
.nw-row{display:flex;justify-content:space-between;align-items:center;
  padding:.55rem .75rem;border-radius:8px;margin-bottom:.3rem;
  background:#131c2e;border:1px solid #1e2d47;font-size:.83rem}
.nw-row:hover{background:#172035;border-color:#2563eb}
.nw-supplier{color:#e2e8f0;font-weight:500;flex:2;
  white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.nw-invoice{color:#64748b;font-size:.75rem;flex:1;text-align:center}
.nw-date{color:#94a3b8;font-size:.75rem;flex:1;text-align:center}
.nw-amount{color:#93c5fd;font-weight:600;font-family:'JetBrains Mono',monospace;
  font-size:.82rem;flex:1;text-align:right}

/* ── Supplier lookup ── */
.lookup-card{
  background:#0d1117;border:1px solid #1e2533;
  border-radius:14px;padding:1.4rem 1.6rem;margin-bottom:1.4rem}
.lookup-header{display:flex;align-items:center;gap:.75rem;margin-bottom:1.2rem}
.lookup-avatar{width:46px;height:46px;border-radius:50%;
  background:#1e3a5f;display:flex;align-items:center;justify-content:center;
  font-size:1.1rem;font-weight:700;color:#93c5fd;flex-shrink:0}
.lookup-name{font-size:1.1rem;font-weight:600;color:#f1f5f9}
.lookup-sub{font-size:.75rem;color:#64748b;margin-top:.1rem}
.lookup-stat-row{display:grid;grid-template-columns:repeat(4,1fr);gap:.8rem;margin-bottom:1rem}
.lookup-stat{background:#131c2e;border:1px solid #1e2d47;border-radius:10px;
  padding:.75rem 1rem}
.lookup-stat-label{font-size:.6rem;color:#64748b;text-transform:uppercase;
  letter-spacing:.1em;margin-bottom:.25rem}
.lookup-stat-value{font-size:1.1rem;font-weight:600;color:#f1f5f9;
  font-family:'JetBrains Mono',monospace}

/* ── Dataframe overrides ── */
[data-testid="stDataFrame"]{border-radius:12px;overflow:hidden}
.stDataFrame [data-testid="stTable"]{background:#0d1117}

/* ── Buttons ── */
.stDownloadButton>button{
  background:#0d1117!important;border:1px solid #3b82f6!important;
  color:#60a5fa!important;border-radius:8px!important;
  font-size:.8rem!important;padding:.4rem 1.1rem!important;font-weight:500!important}
.stDownloadButton>button:hover{background:#0c1a3d!important;border-color:#60a5fa!important}

/* ── File uploader ── */
[data-testid="stFileUploader"]{
  background:#0d1117!important;border:1.5px dashed #1e2533!important;
  border-radius:12px!important}

/* ── Expander ── */
[data-testid="stExpander"]{background:#0d1117;border:1px solid #1e2533;border-radius:12px}

/* ── Divider ── */
hr{border-color:#1e2533!important}

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"]{
  background:#0d1117;padding:4px;
  border-radius:10px;border:1px solid #1e2533;gap:4px}
.stTabs [data-baseweb="tab"]{
  border-radius:8px;color:#64748b;font-weight:500;font-size:.83rem}
.stTabs [aria-selected="true"]{
  background:#1e2533!important;color:#f1f5f9!important}

/* ── Select / input ── */
.stSelectbox>div>div,.stMultiSelect>div>div{
  background:#0d1117;border-color:#1e2533;color:#e2e8f0}
.stTextInput>div>div>input{
  background:#0d1117;border-color:#1e2533;color:#e2e8f0}
.stNumberInput>div>div>input{
  background:#0d1117;border-color:#1e2533;color:#e2e8f0}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
CAT_ORDER  = ["Overdue", "Due This Week", "Due Next Week", "Due Later"]
CAT_COLORS = {
    "Overdue":       "#f43f5e",
    "Due This Week": "#f59e0b",
    "Due Next Week": "#3b82f6",
    "Due Later":     "#10b981",
}

PLOTLY_BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Inter,sans-serif", color="#64748b", size=12),
    margin=dict(l=10, r=10, t=28, b=10),
)

COL_MAP = {
    "supplier":   ["Supplier Number Desc","Vendor Name","Supplier Name","Vendor","Beneficiary"],
    "invoice":    ["Invoice Number","Invoice No","Invoice #","Inv No","Bill Number"],
    "inv_date":   ["Invoice Date","Inv Date","Bill Date","Document Date"],
    "due_date":   ["Due Date","Payment Due","Due By","Maturity Date"],
    "gross_amt":  ["Gross Amount","Invoice Amount"],
    "open_amount":["Open Amount","Outstanding Amount","Balance Due","Amount Due",
                   "Outstanding Balance","Open Amt","Balance"],
    "currency":   ["Curr Code","Currency","Ccy"],
    "doc_type":   ["Document Type","Doc Type","Voucher Type","Type"],
    "pay_status": ["Pay Status Code","Payment Status","Status","Pay Stat"],
}

# ─────────────────────────────────────────────────────────────────────────────
# DATA PROCESSING
# ─────────────────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Processing file...", ttl=3600)
def load_and_process(file_bytes: bytes, ref_date_str: str) -> pd.DataFrame:
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]

    col_lower = {c.lower(): c for c in df.columns}
    rename = {}
    for canonical, variants in COL_MAP.items():
        for v in variants:
            if v.lower() in col_lower:
                rename[col_lower[v.lower()]] = canonical
                break

    df = df.rename(columns=rename)

    for req in ["open_amount", "due_date", "supplier"]:
        if req not in df.columns:
            raise ValueError(
                f"Required column not found: '{req}'\n"
                f"Columns detected: {', '.join(df.columns.tolist())}"
            )

    for col in ["inv_date", "due_date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    df["open_amount"] = pd.to_numeric(
        df["open_amount"].astype(str).str.replace(r"[^\d.\-]", "", regex=True),
        errors="coerce"
    ).fillna(0)

    df = df[df["open_amount"] > 0].copy()
    if df.empty:
        raise ValueError("No rows with Open Amount > 0 found. Check your file.")

    for col, default in [("supplier","Unknown"),("invoice","—"),
                          ("currency","AED"),("doc_type","—"),("pay_status","—")]:
        if col in df.columns:
            df[col] = df[col].fillna(default).astype(str).str.strip()
        else:
            df[col] = default

    today = pd.Timestamp(ref_date_str)
    ws = today - timedelta(days=today.weekday())
    we = ws + timedelta(days=6)
    ns = we + timedelta(days=1)
    ne = ns + timedelta(days=6)

    def cat(d):
        if pd.isna(d):  return "Due Later"
        if d < today:   return "Overdue"
        if d <= we:     return "Due This Week"
        if d <= ne:     return "Due Next Week"
        return "Due Later"

    df["category"]     = df["due_date"].apply(cat)
    df["days_overdue"] = ((today - df["due_date"]).dt.days).clip(lower=0).fillna(0).astype(int)
    return df


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def fmt(v: float) -> str:
    if abs(v) >= 1_000_000: return f"AED {v/1_000_000:.2f}M"
    if abs(v) >= 1_000:     return f"AED {v:,.0f}"
    return f"AED {v:.2f}"

def initials(name: str) -> str:
    parts = name.strip().split()
    if len(parts) >= 2: return (parts[0][0] + parts[1][0]).upper()
    return name[:2].upper()

def kpi_html(label, value, sub="", color="blue", badge="", badge_cls="badge-info") -> str:
    badge_html = f'<span class="kpi-badge {badge_cls}">{badge}</span>' if badge else ""
    return (
        f'<div class="kpi-card {color}">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        + (f'<div class="kpi-sub">{sub}</div>' if sub else "")
        + badge_html + '</div>'
    )

def section_hd(title: str, dot_color: str = "#3b82f6") -> None:
    st.markdown(
        f'<div class="section-hd">'
        f'<span class="dot" style="background:{dot_color}"></span>{title}</div>',
        unsafe_allow_html=True,
    )

def bar_chart(labels, values, colors):
    fig = go.Figure(go.Bar(
        x=labels, y=values, marker_color=colors, marker_line_width=0,
        text=[fmt(v) for v in values], textposition="outside",
        textfont=dict(color="#94a3b8", size=11),
        hovertemplate="<b>%{x}</b><br>AED %{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(
        **PLOTLY_BASE, height=300, showlegend=False,
        xaxis=dict(showgrid=False, tickfont=dict(color="#64748b"), linecolor="#1e2533"),
        yaxis=dict(showgrid=True, gridcolor="#1e2533", zeroline=False,
                   tickformat=",.0f", tickfont=dict(color="#64748b")),
    )
    return fig

def pie_chart(labels, values, colors):
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.6,
        marker=dict(colors=colors, line=dict(color="#070b14", width=3)),
        textinfo="percent", textfont=dict(color="#e2e8f0", size=12),
        hovertemplate="<b>%{label}</b><br>AED %{value:,.0f} (%{percent})<extra></extra>",
    ))
    fig.update_layout(**PLOTLY_BASE, height=300, showlegend=True,
        legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(color="#94a3b8", size=11),
                    orientation="v", x=1, y=.5))
    return fig

def hbar_chart(labels, values, color="#6366f1"):
    h = max(300, len(labels) * 40)
    fig = go.Figure(go.Bar(
        x=values, y=labels, orientation="h",
        marker_color=color, marker_line_width=0,
        text=[fmt(v) for v in values], textposition="outside",
        textfont=dict(color="#94a3b8", size=10),
        hovertemplate="<b>%{y}</b><br>AED %{x:,.0f}<extra></extra>",
    ))
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter,sans-serif", color="#64748b", size=12),
        height=h, showlegend=False,
        xaxis=dict(showgrid=True, gridcolor="#1e2533", zeroline=False,
                   tickformat=",.0f", tickfont=dict(color="#64748b")),
        yaxis=dict(showgrid=False, tickfont=dict(color="#e2e8f0", size=11)),
        margin=dict(l=10, r=100, t=10, b=10),
    )
    return fig

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()

def build_display_df(data: pd.DataFrame) -> pd.DataFrame:
    """Format a dataframe for display — dates, amounts, clean columns."""
    COLS = {
        "supplier":"Supplier","invoice":"Invoice #",
        "inv_date":"Invoice Date","due_date":"Due Date",
        "open_amount":"Open Amount","currency":"Ccy",
        "category":"Category","days_overdue":"Days O/D",
        "doc_type":"Doc Type","pay_status":"Status",
    }
    avail = [c for c in COLS if c in data.columns]
    tbl = data[avail].rename(columns=COLS).copy()
    for dc in ("Invoice Date","Due Date"):
        if dc in tbl.columns:
            tbl[dc] = pd.to_datetime(tbl[dc]).dt.strftime("%d %b %Y")
    if "Open Amount" in tbl.columns:
        tbl["Open Amount"] = tbl["Open Amount"].apply(lambda x: f"{x:,.2f}")
    return tbl


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:1.2rem 0 1.5rem">
      <div style="font-size:1.3rem;font-weight:700;color:#f8fafc;letter-spacing:-.03em">
        VendorLens
      </div>
      <div style="font-size:.68rem;color:#475569;margin-top:.25rem;letter-spacing:.05em;
                  text-transform:uppercase">
        Payment Intelligence
      </div>
    </div>""", unsafe_allow_html=True)

    st.markdown("### 📂 Weekly Upload")
    uploaded = st.file_uploader(
        "Drop your Excel file here",
        type=["xlsx","xls"],
        help="Upload the weekly JD Edwards AP export. Refreshes instantly.",
    )

    st.markdown("---")
    st.markdown("### ⚙️ Settings")
    ref_date = st.date_input(
        "Reference date",
        value=datetime.now().date(),
        help="Override today's date for scenario analysis",
    )

    st.markdown("---")
    st.markdown("""
    <div style="font-size:.66rem;color:#334155;line-height:2.1">
      <div style="color:#475569;font-weight:600;letter-spacing:.08em;
                  text-transform:uppercase;margin-bottom:.3rem">Expected columns</div>
      Supplier Number Desc<br>Invoice Number<br>
      Invoice Date · Due Date<br>
      Open Amount · Curr Code<br>
      Document Type · Pay Status Code
    </div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="margin-bottom:1.6rem;padding-bottom:1rem;border-bottom:1px solid #1e2533">
  <h1 style="font-size:1.7rem;font-weight:700;color:#f8fafc;
             letter-spacing:-.04em;margin:0 0 .25rem">
    Vendor Payment Dashboard
  </h1>
  <div style="color:#475569;font-size:.82rem;letter-spacing:.01em">
    Outstanding balances &nbsp;·&nbsp; Aging analysis &nbsp;·&nbsp; Supplier lookup &nbsp;·&nbsp; Weekly refresh
  </div>
</div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# EMPTY STATE
# ─────────────────────────────────────────────────────────────────────────────
if uploaded is None:
    st.markdown("""
    <div style="text-align:center;padding:6rem 2rem;color:#1e2533">
      <div style="font-size:3rem;margin-bottom:1.2rem;opacity:.4">📊</div>
      <div style="font-size:1rem;font-weight:600;color:#334155;margin-bottom:.7rem">
        Ready for your data
      </div>
      <div style="font-size:.84rem;max-width:400px;margin:0 auto;
                  color:#475569;line-height:1.8">
        Upload your weekly JD Edwards AP export via the
        <strong style="color:#64748b">sidebar uploader</strong>.<br>
        The dashboard updates <em>instantly</em> on every upload.
      </div>
    </div>""", unsafe_allow_html=True)
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────────────────────
try:
    df = load_and_process(uploaded.read(), str(ref_date))
except ValueError as e:
    st.error(f"❌ {e}")
    st.stop()
except Exception as e:
    st.error(f"❌ Could not read file: {e}")
    st.stop()

st.caption(
    f"✓ **{uploaded.name}** — {len(df):,} open lines · "
    f"{df['supplier'].nunique()} suppliers · "
    f"Reference: {ref_date:%d %b %Y}"
)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR FILTERS
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    st.markdown("### 🔍 Filters")

    sel_suppliers = st.multiselect(
        "Supplier", sorted(df["supplier"].unique()), placeholder="All suppliers"
    )
    sel_cats = st.multiselect("Category", CAT_ORDER, placeholder="All categories")
    sel_curr = st.multiselect(
        "Currency", sorted(df["currency"].unique()), placeholder="All"
    )
    sel_doc  = st.multiselect(
        "Document type", sorted(df["doc_type"].unique()), placeholder="All types"
    )
    min_amt  = st.number_input("Min open amount", min_value=0.0, value=0.0, step=1000.0)
    search   = st.text_input("Search supplier", placeholder="Type to filter…")


# ─────────────────────────────────────────────────────────────────────────────
# APPLY FILTERS
# ─────────────────────────────────────────────────────────────────────────────
dff = df.copy()
if sel_suppliers: dff = dff[dff["supplier"].isin(sel_suppliers)]
if sel_cats:      dff = dff[dff["category"].isin(sel_cats)]
if sel_curr:      dff = dff[dff["currency"].isin(sel_curr)]
if sel_doc:       dff = dff[dff["doc_type"].isin(sel_doc)]
if min_amt > 0:   dff = dff[dff["open_amount"] >= min_amt]
if search:        dff = dff[dff["supplier"].str.contains(search, case=False, na=False)]

if dff.empty:
    st.warning("⚠️ No data matches the current filters.")
    st.stop()


# ─────────────────────────────────────────────────────────────────────────────
# KPI CALCULATIONS
# ─────────────────────────────────────────────────────────────────────────────
total     = dff["open_amount"].sum()
overdue   = dff[dff["category"] == "Overdue"]["open_amount"].sum()
this_week = dff[dff["category"] == "Due This Week"]["open_amount"].sum()
next_week = dff[dff["category"] == "Due Next Week"]["open_amount"].sum()
later     = dff[dff["category"] == "Due Later"]["open_amount"].sum()
ov_pct    = overdue / total * 100 if total else 0
ov_cnt    = int((dff["category"] == "Overdue").sum())
nw_cnt    = int((dff["category"] == "Due Next Week").sum())
tw_cnt    = int((dff["category"] == "Due This Week").sum())


# ─────────────────────────────────────────────────────────────────────────────
# ALERT BANNER
# ─────────────────────────────────────────────────────────────────────────────
if ov_pct > 0:
    oldest = int(dff[dff["days_overdue"] > 0]["days_overdue"].max()) if ov_cnt else 0
    st.markdown(
        f'<div class="alert-box">'
        f'<span class="alert-icon">⚠️</span>'
        f'<span><strong>{ov_pct:.1f}%</strong> of total outstanding ({fmt(overdue)}) is '
        f'overdue — <strong>{ov_cnt:,}</strong> invoice lines. '
        f'Oldest invoice is <strong>{oldest} days</strong> past due.</span>'
        f'</div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# KPI CARDS
# ─────────────────────────────────────────────────────────────────────────────
section_hd("Key metrics", "#06b6d4")
c1, c2, c3, c4, c5 = st.columns(5)
for col, label, val, sub, color, badge, bcls in [
    (c1, "Total outstanding", fmt(total),
     f"{len(dff):,} invoices · {dff['supplier'].nunique()} suppliers",
     "teal", "", ""),
    (c2, "Overdue", fmt(overdue),
     f"{ov_pct:.1f}% of total",
     "red", f"{ov_cnt:,} lines", "badge-danger"),
    (c3, "Due this week", fmt(this_week),
     "Action required",
     "amber", f"{tw_cnt} invoices", "badge-warn"),
    (c4, "Due next week", fmt(next_week),
     "Prepare payments",
     "blue", f"{nw_cnt} invoices", "badge-info"),
    (c5, "Due later", fmt(later),
     "Future pipeline",
     "green", f"{int((dff['category']=='Due Later').sum())} invoices", "badge-ok"),
]:
    with col:
        st.markdown(kpi_html(label, val, sub, color, badge, bcls), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# ★ NEW: DUE NEXT WEEK PANEL
# ─────────────────────────────────────────────────────────────────────────────
nw_df = dff[dff["category"] == "Due Next Week"].sort_values("open_amount", ascending=False)

if not nw_df.empty:
    section_hd("Payments due next week", "#3b82f6")

    # Summary strip
    nw_total    = nw_df["open_amount"].sum()
    nw_suppliers = nw_df["supplier"].nunique()
    s1, s2, s3, s4 = st.columns(4)
    with s1:
        st.markdown(kpi_html("Total due next week", fmt(nw_total), "", "blue"), unsafe_allow_html=True)
    with s2:
        st.markdown(kpi_html("Invoice lines", str(len(nw_df)), "", "blue"), unsafe_allow_html=True)
    with s3:
        st.markdown(kpi_html("Suppliers involved", str(nw_suppliers), "", "blue"), unsafe_allow_html=True)
    with s4:
        avg_nw = nw_df["open_amount"].mean()
        st.markdown(kpi_html("Average invoice", fmt(avg_nw), "", "blue"), unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # Row-by-row styled list (up to 25 rows, rest available via table)
    rows_html = ""
    for _, row in nw_df.head(25).iterrows():
        due_str  = row["due_date"].strftime("%d %b %Y") if pd.notna(row.get("due_date")) else "—"
        inv_str  = str(row.get("invoice","—"))[:20]
        rows_html += (
            f'<div class="nw-row">'
            f'<span class="nw-supplier" title="{row["supplier"]}">{row["supplier"]}</span>'
            f'<span class="nw-invoice">{inv_str}</span>'
            f'<span class="nw-date">Due {due_str}</span>'
            f'<span class="nw-amount">{fmt(row["open_amount"])}</span>'
            f'</div>'
        )
    if len(nw_df) > 25:
        rows_html += (
            f'<div style="text-align:center;padding:.5rem;font-size:.75rem;color:#64748b">'
            f'+ {len(nw_df)-25} more rows — use Export below</div>'
        )

    st.markdown(
        f'<div class="nw-panel">'
        f'<div class="nw-title">📅 &nbsp;Next week invoice list '
        f'<span style="color:#475569;font-weight:400;font-size:.72rem">'
        f'({len(nw_df)} invoices)</span></div>'
        f'<div style="display:flex;justify-content:space-between;'
        f'padding:.3rem .75rem .5rem;font-size:.65rem;color:#475569;'
        f'text-transform:uppercase;letter-spacing:.08em">'
        f'<span style="flex:2">Supplier</span>'
        f'<span style="flex:1;text-align:center">Invoice #</span>'
        f'<span style="flex:1;text-align:center">Due date</span>'
        f'<span style="flex:1;text-align:right">Amount</span></div>'
        + rows_html +
        f'</div>',
        unsafe_allow_html=True,
    )

    nw_dl1, nw_dl2, _ = st.columns([1,1,5])
    with nw_dl1:
        st.download_button("⬇ Export CSV", nw_df.to_csv(index=False).encode(),
            f"due_next_week_{datetime.now():%Y%m%d}.csv", "text/csv")
    with nw_dl2:
        st.download_button("⬇ Export Excel", to_excel_bytes(nw_df),
            f"due_next_week_{datetime.now():%Y%m%d}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("<br>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# AGING CHARTS
# ─────────────────────────────────────────────────────────────────────────────
section_hd("Aging breakdown", "#f59e0b")
cb, cp = st.columns([3, 2])

cat_df  = dff.groupby("category")["open_amount"].sum().reindex(CAT_ORDER).fillna(0)
cat_amt = cat_df.tolist()
cat_col = [CAT_COLORS[c] for c in CAT_ORDER]

with cb:
    st.plotly_chart(bar_chart(CAT_ORDER, cat_amt, cat_col),
                    use_container_width=True, config={"displayModeBar":False})
with cp:
    nz = [(c, a) for c, a in zip(CAT_ORDER, cat_amt) if a > 0]
    if nz:
        lbl, val_ = zip(*nz)
        st.plotly_chart(
            pie_chart(list(lbl), list(val_), [CAT_COLORS[c] for c in lbl]),
            use_container_width=True, config={"displayModeBar":False},
        )


# ─────────────────────────────────────────────────────────────────────────────
# EXPOSURE CHARTS
# ─────────────────────────────────────────────────────────────────────────────
section_hd("Exposure analysis", "#8b5cf6")
cs, cc = st.columns([3, 2])

with cs:
    st.markdown("**Top suppliers by open balance**")
    top_sup = (dff.groupby("supplier")["open_amount"].sum()
                  .nlargest(12).sort_values().reset_index())
    st.plotly_chart(
        hbar_chart(top_sup["supplier"].tolist(), top_sup["open_amount"].tolist()),
        use_container_width=True, config={"displayModeBar":False},
    )

with cc:
    st.markdown("**Currency split**")
    curr_df = dff.groupby("currency")["open_amount"].sum().reset_index()
    st.plotly_chart(
        pie_chart(curr_df["currency"].tolist(), curr_df["open_amount"].tolist(),
                  ["#3b82f6","#10b981","#f59e0b","#f43f5e","#8b5cf6"]),
        use_container_width=True, config={"displayModeBar":False},
    )
    st.markdown("**Document type breakdown**")
    doc_df = dff.groupby("doc_type")["open_amount"].sum().nlargest(6).reset_index()
    st.plotly_chart(
        bar_chart(doc_df["doc_type"].str[:22].tolist(), doc_df["open_amount"].tolist(),
                  ["#6366f1"] * len(doc_df)),
        use_container_width=True, config={"displayModeBar":False},
    )


# ─────────────────────────────────────────────────────────────────────────────
# ★ NEW: SUPPLIER LOOKUP MODULE
# ─────────────────────────────────────────────────────────────────────────────
section_hd("Supplier lookup", "#10b981")

all_suppliers = sorted(df["supplier"].unique())  # search across full dataset (no filters)
lookup_name = st.selectbox(
    "Search by supplier name",
    options=[""] + all_suppliers,
    format_func=lambda x: "— select a supplier —" if x == "" else x,
    key="supplier_lookup",
)

if lookup_name:
    sup_data = df[df["supplier"] == lookup_name].sort_values("open_amount", ascending=False)

    # Header card
    sup_total   = sup_data["open_amount"].sum()
    sup_ov      = sup_data[sup_data["category"]=="Overdue"]["open_amount"].sum()
    sup_ov_pct  = sup_ov / sup_total * 100 if sup_total else 0
    sup_inv_cnt = len(sup_data)
    sup_oldest  = int(sup_data["days_overdue"].max()) if sup_ov > 0 else 0
    sup_init    = initials(lookup_name)

    st.markdown(
        f'<div class="lookup-card">'
        f'<div class="lookup-header">'
        f'<div class="lookup-avatar">{sup_init}</div>'
        f'<div><div class="lookup-name">{lookup_name}</div>'
        f'<div class="lookup-sub">{sup_inv_cnt} invoice lines &nbsp;·&nbsp; '
        f'{sup_data["currency"].nunique()} currencies &nbsp;·&nbsp; '
        f'{sup_data["doc_type"].nunique()} document types</div></div>'
        f'</div>'
        f'<div class="lookup-stat-row">'
        f'<div class="lookup-stat"><div class="lookup-stat-label">Total open</div>'
        f'<div class="lookup-stat-value">{fmt(sup_total)}</div></div>'
        f'<div class="lookup-stat"><div class="lookup-stat-label">Overdue</div>'
        f'<div class="lookup-stat-value" style="color:#f43f5e">{fmt(sup_ov)}</div></div>'
        f'<div class="lookup-stat"><div class="lookup-stat-label">Overdue %</div>'
        f'<div class="lookup-stat-value" style="color:#f43f5e">{sup_ov_pct:.1f}%</div></div>'
        f'<div class="lookup-stat"><div class="lookup-stat-label">Oldest O/D</div>'
        f'<div class="lookup-stat-value" style="color:#f59e0b">'
        f'{sup_oldest if sup_oldest>0 else "—"} {"days" if sup_oldest>0 else ""}</div></div>'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

    # Full invoice table for this supplier
    st.markdown(f"**All invoice lines for {lookup_name}**")
    sup_tbl = build_display_df(sup_data)
    st.dataframe(sup_tbl, use_container_width=True, height=350, hide_index=True)

    sd1, sd2, _ = st.columns([1,1,5])
    with sd1:
        st.download_button("⬇ Export CSV", sup_data.to_csv(index=False).encode(),
            f"supplier_{lookup_name[:20]}_{datetime.now():%Y%m%d}.csv", "text/csv")
    with sd2:
        st.download_button("⬇ Export Excel", to_excel_bytes(sup_data),
            f"supplier_{lookup_name[:20]}_{datetime.now():%Y%m%d}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("<br>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# INVOICE DETAIL TABLE  (with drill-down via session state)
# ─────────────────────────────────────────────────────────────────────────────
section_hd("Invoice detail", "#f43f5e")

# Sort control
sort_opt = st.selectbox(
    "Sort",
    ["Open Amount ↓","Open Amount ↑","Days Overdue ↓","Supplier A→Z",
     "Due Date ↑","Due Date ↓"],
    label_visibility="collapsed",
)
_sort = {
    "Open Amount ↓":  ("open_amount",  False),
    "Open Amount ↑":  ("open_amount",  True),
    "Days Overdue ↓": ("days_overdue", False),
    "Supplier A→Z":   ("supplier",     True),
    "Due Date ↑":     ("due_date",     True),
    "Due Date ↓":     ("due_date",     False),
}
sc, sa = _sort[sort_opt]
display = dff.sort_values(sc, ascending=sa)

# ── Drill-down toggle via session state ──────────────────────────────────────
if "drilldown_supplier" not in st.session_state:
    st.session_state.drilldown_supplier = None

# Show drill-down panel if active
if st.session_state.drilldown_supplier:
    drill_name = st.session_state.drilldown_supplier
    drill_data = df[df["supplier"] == drill_name].sort_values("open_amount", ascending=False)

    drill_total = drill_data["open_amount"].sum()
    drill_ov    = drill_data[drill_data["category"]=="Overdue"]["open_amount"].sum()

    st.markdown(
        f'<div style="background:#0d1117;border:2px solid #f43f5e;border-radius:14px;'
        f'padding:1.2rem 1.5rem;margin-bottom:1rem">'
        f'<div style="display:flex;justify-content:space-between;align-items:center;'
        f'margin-bottom:.8rem">'
        f'<div style="font-size:1rem;font-weight:600;color:#f1f5f9">'
        f'📋 &nbsp;{drill_name}</div>'
        f'<div style="font-size:.75rem;color:#64748b">'
        f'Total open: <strong style="color:#f8fafc">{fmt(drill_total)}</strong> &nbsp;|&nbsp; '
        f'Overdue: <strong style="color:#f43f5e">{fmt(drill_ov)}</strong></div>'
        f'</div>',
        unsafe_allow_html=True,
    )
    drill_tbl = build_display_df(drill_data)
    st.dataframe(drill_tbl, use_container_width=True, height=300, hide_index=True)

    dd1, dd2, dd3, _ = st.columns([1,1,1,4])
    with dd1:
        st.download_button("⬇ CSV", drill_data.to_csv(index=False).encode(),
            f"drill_{drill_name[:15]}_{datetime.now():%Y%m%d}.csv", "text/csv",
            key="drill_csv")
    with dd2:
        st.download_button("⬇ Excel", to_excel_bytes(drill_data),
            f"drill_{drill_name[:15]}_{datetime.now():%Y%m%d}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="drill_xlsx")
    with dd3:
        if st.button("✕ Close", key="close_drill"):
            st.session_state.drilldown_supplier = None
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

# ── Main table ───────────────────────────────────────────────────────────────
st.caption(
    f"Showing {len(display):,} rows · "
    "Click a supplier button below the table to drill into their full payment details"
)

main_tbl = build_display_df(display)
st.dataframe(main_tbl, use_container_width=True, height=430, hide_index=True)

# ── Supplier drill-down buttons (grouped, compact) ───────────────────────────
st.markdown(
    '<div style="font-size:.72rem;color:#475569;margin:.5rem 0 .4rem;'
    'text-transform:uppercase;letter-spacing:.08em">'
    '🔍 Click supplier to drill down</div>',
    unsafe_allow_html=True,
)

visible_suppliers = display["supplier"].unique().tolist()[:30]  # cap at 30 buttons
btn_cols = st.columns(min(len(visible_suppliers), 5))
for i, sup in enumerate(visible_suppliers):
    with btn_cols[i % 5]:
        sup_short = sup[:22] + "…" if len(sup) > 22 else sup
        if st.button(sup_short, key=f"drill_btn_{sup}", help=sup,
                     use_container_width=True):
            st.session_state.drilldown_supplier = sup
            st.rerun()

st.markdown("<br>", unsafe_allow_html=True)

# ── Download buttons ─────────────────────────────────────────────────────────
dl1, dl2, _ = st.columns([1,1,5])
with dl1:
    st.download_button("⬇ Export CSV",
        display.to_csv(index=False).encode(),
        f"vendor_outstanding_{datetime.now():%Y%m%d}.csv", "text/csv")
with dl2:
    st.download_button("⬇ Export Excel",
        to_excel_bytes(display),
        f"vendor_outstanding_{datetime.now():%Y%m%d}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ─────────────────────────────────────────────────────────────────────────────
# EXECUTIVE SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
with st.expander("📋 Executive summary", expanded=False):
    top3 = (dff[dff["category"]=="Overdue"]
              .groupby("supplier")["open_amount"].sum().nlargest(3))
    oldest = int(dff[dff["days_overdue"]>0]["days_overdue"].max()) if ov_cnt else 0
    avg_ov = dff[dff["days_overdue"]>0]["days_overdue"].mean() if ov_cnt else 0

    st.markdown(f"""
**As at {ref_date:%d %b %Y}**

| Metric | Value |
|--------|-------|
| Total outstanding | **{fmt(total)}** |
| Overdue | **{fmt(overdue)}** ({ov_pct:.1f}%) |
| Overdue invoice lines | **{ov_cnt:,}** |
| Oldest overdue | **{oldest} days** past due |
| Avg overdue age | **{avg_ov:.0f} days** |
| Due this week | **{fmt(this_week)}** ({tw_cnt} invoices) |
| Due next week | **{fmt(next_week)}** ({nw_cnt} invoices) |
| Unique suppliers | **{dff['supplier'].nunique()}** |

**Top 3 overdue suppliers:**
""")
    for s, a in top3.items():
        st.markdown(f"- **{s}**: {fmt(a)}")


# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    "VendorLens v2 &nbsp;·&nbsp; Upload a new Excel file each week to refresh &nbsp;·&nbsp; "
    "Data processed in-memory, never stored on any server"
)
