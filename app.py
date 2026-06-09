"""
Vendor Payment Dashboard  —  VendorLens
========================================
CEO-grade dashboard built around your JD Edwards vendor AP export format.
Upload any Excel in the same column format to refresh data instantly.

Local:   streamlit run app.py
Deploy:  push to GitHub → connect to Streamlit Cloud (share.streamlit.io)
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
    page_title="Vendor Payment Dashboard",
    page_icon="💳",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');
html,body,[class*="css"]{font-family:'DM Sans',sans-serif}
.main{background:#0f1117}
.block-container{padding:2rem 2rem 4rem;max-width:1600px}
[data-testid="stSidebar"]{background:#16181f!important;border-right:1px solid #2a2d3a}
[data-testid="stSidebar"] h3{color:#e2e8f0;font-size:.75rem;letter-spacing:.12em;text-transform:uppercase;font-weight:600}

.kpi-card{background:#1a1d27;border:1px solid #2a2d3a;border-radius:12px;
  padding:1.2rem 1.4rem;position:relative;overflow:hidden;transition:transform .18s,border-color .18s}
.kpi-card:hover{transform:translateY(-2px);border-color:#3a3d4a}
.kpi-card::before{content:"";position:absolute;top:0;left:0;right:0;height:3px;border-radius:12px 12px 0 0}
.kpi-card.red::before{background:#ef4444}
.kpi-card.amber::before{background:#f59e0b}
.kpi-card.blue::before{background:#3b82f6}
.kpi-card.green::before{background:#10b981}
.kpi-card.teal::before{background:#06b6d4}
.kpi-label{font-size:.68rem;color:#6b7280;text-transform:uppercase;letter-spacing:.1em;font-weight:600;margin-bottom:.4rem}
.kpi-value{font-size:1.75rem;font-weight:700;color:#f1f5f9;font-family:'DM Mono',monospace;line-height:1.1}
.kpi-sub{font-size:.73rem;color:#9ca3af;margin-top:.3rem}

.alert-box{background:#3f1515;border:1px solid #7f1d1d;border-radius:10px;
  padding:.9rem 1.2rem;color:#fca5a5;font-size:.88rem;margin-bottom:1.2rem}
.alert-box strong{color:#f87171}

.section-hd{font-size:1rem;font-weight:600;color:#e2e8f0;padding:.5rem 0 .3rem;
  border-bottom:1px solid #2a2d3a;margin-bottom:1rem}

[data-testid="stDataFrame"]{border-radius:10px;overflow:hidden}
.stDownloadButton>button{background:#1a1d27!important;border:1px solid #3b82f6!important;
  color:#60a5fa!important;border-radius:8px!important;font-size:.82rem!important;padding:.4rem 1rem!important}
.stDownloadButton>button:hover{background:#0f2344!important}

[data-testid="stFileUploader"]{background:#1a1d27!important;
  border:1.5px dashed #3a3d4a!important;border-radius:10px!important}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
CAT_ORDER  = ["Overdue", "Due This Week", "Due Next Week", "Due Later"]
CAT_COLORS = {"Overdue":"#ef4444","Due This Week":"#f59e0b","Due Next Week":"#3b82f6","Due Later":"#10b981"}

PLOTLY_BASE = dict(
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="DM Sans,sans-serif", color="#9ca3af", size=12),
    margin=dict(l=10, r=10, t=28, b=10),
)

# Column name variants — maps canonical key → acceptable raw names
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
    """
    Parse uploaded Excel, normalize column names, compute aging categories.
    ref_date_str is passed so cache invalidates if user changes the date.
    """
    df = pd.read_excel(io.BytesIO(file_bytes))
    df.columns = [str(c).strip() for c in df.columns]

    # Build case-insensitive rename map
    col_lower = {c.lower(): c for c in df.columns}
    rename = {}
    for canonical, variants in COL_MAP.items():
        for v in variants:
            if v.lower() in col_lower:
                rename[col_lower[v.lower()]] = canonical
                break

    df = df.rename(columns=rename)

    # Validate required columns
    for req in ["open_amount", "due_date", "supplier"]:
        if req not in df.columns:
            raise ValueError(
                f"Required column not found: '{req}'\n"
                f"Columns detected: {', '.join(df.columns.tolist())}"
            )

    # Parse dates
    for col in ["inv_date", "due_date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # Parse amounts
    df["open_amount"] = pd.to_numeric(
        df["open_amount"].astype(str).str.replace(r"[^\d.\-]", "", regex=True),
        errors="coerce"
    ).fillna(0)

    # Keep only positive open amounts
    df = df[df["open_amount"] > 0].copy()
    if df.empty:
        raise ValueError("No rows with Open Amount > 0 found. Check your file.")

    # Fill optional text columns
    for col, default in [("supplier","Unknown"),("invoice","—"),
                          ("currency","AED"),("doc_type","—"),("pay_status","—")]:
        if col in df.columns:
            df[col] = df[col].fillna(default).astype(str).str.strip()
        else:
            df[col] = default

    # Aging  (week boundaries based on ref_date)
    today = pd.Timestamp(ref_date_str)
    ws = today - timedelta(days=today.weekday())   # Monday
    we = ws + timedelta(days=6)                     # Sunday
    ns = we + timedelta(days=1)                     # Next Monday
    ne = ns + timedelta(days=6)                     # Next Sunday

    def cat(d):
        if pd.isna(d):     return "Due Later"
        if d < today:      return "Overdue"
        if d <= we:        return "Due This Week"
        if d <= ne:        return "Due Next Week"
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

def kpi_html(label, value, sub="", color="blue") -> str:
    return (f'<div class="kpi-card {color}">'
            f'<div class="kpi-label">{label}</div>'
            f'<div class="kpi-value">{value}</div>'
            + (f'<div class="kpi-sub">{sub}</div>' if sub else "")
            + '</div>')

def bar_chart(labels, values, colors):
    fig = go.Figure(go.Bar(
        x=labels, y=values, marker_color=colors, marker_line_width=0,
        text=[fmt(v) for v in values], textposition="outside",
        textfont=dict(color="#d1d5db", size=11),
        hovertemplate="<b>%{x}</b><br>AED %{y:,.0f}<extra></extra>",
    ))
    fig.update_layout(
        **PLOTLY_BASE, height=300, showlegend=False,
        xaxis=dict(showgrid=False, tickfont=dict(color="#9ca3af")),
        yaxis=dict(showgrid=True, gridcolor="#2a2d3a", zeroline=False,
                   tickformat=",.0f", tickfont=dict(color="#9ca3af")),
    )
    return fig

def pie_chart(labels, values, colors):
    fig = go.Figure(go.Pie(
        labels=labels, values=values, hole=0.55,
        marker=dict(colors=colors, line=dict(color="#0f1117", width=3)),
        textinfo="percent+label", textfont=dict(color="#e2e8f0", size=11),
        hovertemplate="<b>%{label}</b><br>AED %{value:,.0f} (%{percent})<extra></extra>",
    ))
    fig.update_layout(**PLOTLY_BASE, height=300, showlegend=False)
    return fig

def hbar_chart(labels, values, color="#8b5cf6"):
    h = max(280, len(labels) * 38)
    fig = go.Figure(go.Bar(
        x=values, y=labels, orientation="h",
        marker_color=color, marker_line_width=0,
        text=[fmt(v) for v in values], textposition="outside",
        textfont=dict(color="#d1d5db", size=10),
        hovertemplate="<b>%{y}</b><br>AED %{x:,.0f}<extra></extra>",
    ))
    fig.update_layout(
        **PLOTLY_BASE, height=h, showlegend=False,
        xaxis=dict(showgrid=True, gridcolor="#2a2d3a", zeroline=False,
                   tickformat=",.0f", tickfont=dict(color="#9ca3af")),
        yaxis=dict(showgrid=False, tickfont=dict(color="#e2e8f0", size=11)),
        margin=dict(l=10, r=90, t=10, b=10),
    )
    return fig

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="padding:1rem 0 1.4rem">
      <div style="font-size:1.15rem;font-weight:700;color:#f1f5f9;letter-spacing:-.02em">
        💳 VendorLens
      </div>
      <div style="font-size:.7rem;color:#6b7280;margin-top:.2rem">
        CEO Payment Intelligence
      </div>
    </div>""", unsafe_allow_html=True)

    st.markdown("### 📂 Weekly Upload")
    uploaded = st.file_uploader(
        "Drop your Excel file here",
        type=["xlsx", "xls"],
        help="Upload the weekly JD Edwards AP export. Dashboard refreshes instantly.",
    )

    st.markdown("---")
    st.markdown("### ⚙️ Settings")
    ref_date = st.date_input(
        "Reference date (today)",
        value=datetime.now().date(),
        help="Change to run scenario analysis against a different date",
    )

    st.markdown("---")
    st.markdown("""
    <div style="font-size:.67rem;color:#4b5563;line-height:2">
      <div style="color:#6b7280;font-weight:600;margin-bottom:.2rem">EXPECTED COLUMNS</div>
      Supplier Number Desc<br>Invoice Number<br>
      Invoice Date · Due Date<br>
      Open Amount · Curr Code<br>
      Document Type · Pay Status Code
    </div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="margin-bottom:1.5rem">
  <h1 style="font-size:1.6rem;font-weight:700;color:#f1f5f9;letter-spacing:-.03em;margin:0 0 .2rem">
    Vendor Payment Dashboard
  </h1>
  <div style="color:#6b7280;font-size:.83rem">
    Outstanding balances · Aging analysis · Weekly refresh
  </div>
</div>""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# EMPTY STATE
# ─────────────────────────────────────────────────────────────────────────────
if uploaded is None:
    st.markdown("""
    <div style="text-align:center;padding:5rem 2rem;color:#4b5563">
      <div style="font-size:3.5rem;margin-bottom:1.2rem">📂</div>
      <div style="font-size:1.05rem;font-weight:600;color:#6b7280;margin-bottom:.6rem">
        Ready for your data
      </div>
      <div style="font-size:.85rem;max-width:420px;margin:0 auto;line-height:1.7">
        Upload your weekly Excel export using the
        <strong style="color:#9ca3af">sidebar uploader</strong>.<br>
        The dashboard updates <em>instantly</em> — no page reload needed.
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
    f"Reference date: {ref_date:%d %b %Y}"
)


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR FILTERS (after data is loaded so we can populate options)
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    st.markdown("### 🔍 Filters")

    sel_suppliers = st.multiselect(
        "Supplier", sorted(df["supplier"].unique()), placeholder="All suppliers"
    )
    sel_cats = st.multiselect(
        "Category", CAT_ORDER, placeholder="All categories"
    )
    sel_curr = st.multiselect(
        "Currency", sorted(df["currency"].unique()), placeholder="All"
    )
    sel_doc = st.multiselect(
        "Document type", sorted(df["doc_type"].unique()), placeholder="All types"
    )
    min_amt = st.number_input(
        "Min open amount (AED)", min_value=0.0, value=0.0, step=1000.0
    )
    search = st.text_input("Search supplier name", placeholder="Type to filter...")


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
    st.warning("⚠️ No data matches the current filters. Adjust and try again.")
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
ov_cnt    = (dff["category"] == "Overdue").sum()


# ─────────────────────────────────────────────────────────────────────────────
# ALERT BANNER
# ─────────────────────────────────────────────────────────────────────────────
if ov_pct > 0:
    st.markdown(
        f'<div class="alert-box">⚠️ <strong>{ov_pct:.1f}%</strong> of outstanding '
        f'({fmt(overdue)}) is overdue — '
        f'<strong>{ov_cnt:,}</strong> invoice lines require immediate attention.</div>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────────────────────────────────────
# KPI CARDS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-hd">Key metrics</div>', unsafe_allow_html=True)
c1, c2, c3, c4, c5 = st.columns(5)
for col, label, val, sub, color in [
    (c1, "Total outstanding", fmt(total),
     f"{len(dff):,} invoices · {dff['supplier'].nunique()} suppliers", "teal"),
    (c2, "Overdue",           fmt(overdue),
     f"{ov_pct:.1f}% of total · {ov_cnt:,} lines", "red"),
    (c3, "Due this week",     fmt(this_week),
     f"{(dff['category']=='Due This Week').sum()} invoices", "amber"),
    (c4, "Due next week",     fmt(next_week),
     f"{(dff['category']=='Due Next Week').sum()} invoices", "blue"),
    (c5, "Due later",         fmt(later),
     f"{(dff['category']=='Due Later').sum()} invoices", "green"),
]:
    with col:
        st.markdown(kpi_html(label, val, sub, color), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# CHARTS — ROW 1: Aging bar + pie
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-hd">Aging breakdown</div>', unsafe_allow_html=True)
cb, cp = st.columns([3, 2])

cat_df  = dff.groupby("category")["open_amount"].sum().reindex(CAT_ORDER).fillna(0)
cat_amt = cat_df.tolist()
cat_col = [CAT_COLORS[c] for c in CAT_ORDER]

with cb:
    st.plotly_chart(bar_chart(CAT_ORDER, cat_amt, cat_col),
                    use_container_width=True, config={"displayModeBar": False})

with cp:
    nz = [(c, a) for c, a in zip(CAT_ORDER, cat_amt) if a > 0]
    if nz:
        lbl, val_ = zip(*nz)
        st.plotly_chart(
            pie_chart(list(lbl), list(val_), [CAT_COLORS[c] for c in lbl]),
            use_container_width=True, config={"displayModeBar": False},
        )


# ─────────────────────────────────────────────────────────────────────────────
# CHARTS — ROW 2: Top suppliers + currency + doc type
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-hd">Exposure analysis</div>', unsafe_allow_html=True)
cs, cc = st.columns([3, 2])

with cs:
    st.markdown("**Top suppliers by open balance**")
    top_sup = (dff.groupby("supplier")["open_amount"].sum()
                  .nlargest(12).sort_values().reset_index())
    st.plotly_chart(
        hbar_chart(top_sup["supplier"].tolist(), top_sup["open_amount"].tolist()),
        use_container_width=True, config={"displayModeBar": False},
    )

with cc:
    st.markdown("**Currency split**")
    curr_df = dff.groupby("currency")["open_amount"].sum().reset_index()
    st.plotly_chart(
        pie_chart(curr_df["currency"].tolist(), curr_df["open_amount"].tolist(),
                  ["#3b82f6","#10b981","#f59e0b","#ef4444","#8b5cf6"]),
        use_container_width=True, config={"displayModeBar": False},
    )
    st.markdown("**Document type breakdown**")
    doc_df = dff.groupby("doc_type")["open_amount"].sum().nlargest(6).reset_index()
    st.plotly_chart(
        bar_chart(doc_df["doc_type"].str[:22].tolist(),
                  doc_df["open_amount"].tolist(),
                  ["#3b82f6"] * len(doc_df)),
        use_container_width=True, config={"displayModeBar": False},
    )


# ─────────────────────────────────────────────────────────────────────────────
# INVOICE DETAIL TABLE
# ─────────────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-hd">Invoice detail</div>', unsafe_allow_html=True)

sort_opt = st.selectbox(
    "Sort by",
    ["Open Amount ↓","Open Amount ↑","Days Overdue ↓","Supplier A→Z","Due Date ↑","Due Date ↓"],
    label_visibility="collapsed",
)
_sort = {"Open Amount ↓":("open_amount",False),"Open Amount ↑":("open_amount",True),
         "Days Overdue ↓":("days_overdue",False),"Supplier A→Z":("supplier",True),
         "Due Date ↑":("due_date",True),"Due Date ↓":("due_date",False)}
sc, sa = _sort[sort_opt]
display = dff.sort_values(sc, ascending=sa)

COLS = {"supplier":"Supplier","invoice":"Invoice #","inv_date":"Invoice Date",
        "due_date":"Due Date","open_amount":"Open Amount","currency":"Ccy",
        "category":"Category","days_overdue":"Days O/D",
        "doc_type":"Doc Type","pay_status":"Status"}
avail = [c for c in COLS if c in display.columns]
tbl = display[avail].rename(columns=COLS).copy()

for datecol in ("Invoice Date","Due Date"):
    if datecol in tbl.columns:
        tbl[datecol] = pd.to_datetime(tbl[datecol]).dt.strftime("%d %b %Y")
if "Open Amount" in tbl.columns:
    tbl["Open Amount"] = tbl["Open Amount"].apply(lambda x: f"{x:,.2f}")

st.caption(f"Showing {len(tbl):,} rows (filtered from {len(df):,} total)")
st.dataframe(tbl, use_container_width=True, height=430, hide_index=True)

dl1, dl2, _ = st.columns([1,1,5])
with dl1:
    st.download_button(
        "⬇ Export CSV",
        data=display.to_csv(index=False).encode(),
        file_name=f"vendor_outstanding_{datetime.now():%Y%m%d}.csv",
        mime="text/csv",
    )
with dl2:
    st.download_button(
        "⬇ Export Excel",
        data=to_excel_bytes(display),
        file_name=f"vendor_outstanding_{datetime.now():%Y%m%d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ─────────────────────────────────────────────────────────────────────────────
# EXECUTIVE SUMMARY EXPANDER
# ─────────────────────────────────────────────────────────────────────────────
with st.expander("📋 Executive summary (click to expand)", expanded=False):
    top3 = (dff[dff["category"]=="Overdue"]
              .groupby("supplier")["open_amount"].sum().nlargest(3))
    oldest = int(dff[dff["days_overdue"] > 0]["days_overdue"].max()) if ov_cnt else 0
    avg_ov = dff[dff["days_overdue"] > 0]["days_overdue"].mean() if ov_cnt else 0

    st.markdown(f"""
**As at {ref_date:%d %b %Y}**

| Metric | Value |
|--------|-------|
| Total outstanding | **{fmt(total)}** |
| Overdue | **{fmt(overdue)}** ({ov_pct:.1f}%) |
| Overdue invoice lines | **{ov_cnt:,}** |
| Oldest overdue | **{oldest} days** past due |
| Average overdue age | **{avg_ov:.0f} days** |
| Due this week | **{fmt(this_week)}** |
| Due next week | **{fmt(next_week)}** |
| Unique suppliers | **{dff['supplier'].nunique()}** |

**Top 3 overdue suppliers:**
""")
    for s, a in top3.items():
        st.markdown(f"- **{s}**: {fmt(a)}")


# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("VendorLens · Upload a new Excel file each week to refresh · "
           "Data is processed in-browser and never stored on any server")
