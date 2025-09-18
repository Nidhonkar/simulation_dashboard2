
import streamlit as st
import pandas as pd
import numpy as np
import re, json, os
from pathlib import Path

# Optional plotting backends
HAS_PLOTLY = True
try:
    import plotly.express as px
except Exception:
    HAS_PLOTLY = False
try:
    import altair as alt
except Exception:
    alt = None

st.set_page_config(page_title="Ganga Jamuna — VP Dashboard (Fresh Connection)", layout="wide")

st.markdown(
    """
    <style>
    .title {font-size: 38px; font-weight: 800; margin-bottom: 0rem;}
    .subtitle {font-size: 16px; opacity: 0.85; margin-bottom: 1rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

DATA_DIR = Path("data")

EXPECTED_NAMES = [
    "TFC_0_6.xlsx",
    "FinanceReport_6.xlsx",           # safe alias
    "FinanceReport (6).xlsx"          # original
]

def find_local_excels():
    files = []
    if DATA_DIR.exists():
        # First grab expected names if present, then any xlsx
        for name in EXPECTED_NAMES:
            p = DATA_DIR / name
            if p.exists():
                files.append(str(p))
        # Add any other .xlsx files not already included
        for p in DATA_DIR.glob("*.xlsx"):
            if str(p) not in files:
                files.append(str(p))
    return files

@st.cache_data(show_spinner=False)
def load_excels(filepaths):
    frames, meta = [], []
    for f in filepaths:
        p = Path(f)
        try:
            xls = pd.ExcelFile(p)
            for s in xls.sheet_names:
                try:
                    df = xls.parse(s)
                    df["__source_file__"] = p.name
                    df["__sheet__"] = s
                    frames.append(df)
                    meta.append({"file": p.name, "sheet": s, "rows": int(len(df)), "cols": [str(c) for c in df.columns]})
                except Exception as e:
                    meta.append({"file": p.name, "sheet": s, "error": str(e)})
        except Exception as e:
            meta.append({"file": p.name if p.exists() else str(p), "error": f"Failed to read: {e}"})
    return frames, meta

def safe_concat(dfs):
    if len(dfs) == 0:
        return pd.DataFrame()
    all_cols=set()
    for d in dfs: all_cols.update(list(map(str, d.columns)))
    cols=list(all_cols)
    outs=[]
    for d in dfs:
        dd=d.copy()
        dd.columns = list(map(str, dd.columns))
        for c in cols:
            if c not in dd.columns: dd[c]=np.nan
        outs.append(dd[cols])
    return pd.concat(outs, ignore_index=True)

# ---- File selection UI: local + uploader ----
st.sidebar.markdown("### Data Source")
local_files = find_local_excels()
use_files = local_files.copy()

uploaded = st.sidebar.file_uploader("Upload one or more .xlsx files", type=["xlsx"], accept_multiple_files=True)
if uploaded:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    for uf in uploaded:
        dest = DATA_DIR / uf.name
        with open(dest, "wb") as f:
            f.write(uf.getbuffer())
        if str(dest) not in use_files:
            use_files.append(str(dest))

st.sidebar.write("**Using files:**")
if len(use_files)==0:
    st.sidebar.warning("No Excel files detected yet. Upload files above or add to the `data/` folder and refresh.")
else:
    for f in use_files:
        st.sidebar.write(f"- {Path(f).name}")

frames, meta = load_excels(use_files)

if len(frames)==0:
    st.error("Could not load any sheets from the selected files. Please verify the Excel files.")
    st.stop()

data_all = safe_concat(frames)

# --------- Column matching + dashboard (same as v2, trimmed for brevity) ---------
def match_col(candidates, patterns, default=None):
    if isinstance(patterns, str): patterns=[patterns]
    for pat in patterns:
        rx=re.compile(pat, flags=re.I)
        for c in candidates:
            if rx.search(str(c)): return c
    return default

cands = list(map(str, data_all.columns))

defaults = {
    "round": match_col(cands, [r"^round$", r"week", r"period", r"cycle", r"game\s*round"]),
    "date": match_col(cands, [r"\bdate\b"]),
    "product": match_col(cands, [r"^product$", r"sku", r"item"]),
    "customer": match_col(cands, [r"^customer", r"account", r"client"]),
    "component": match_col(cands, [r"^component", r"part"]),
    "supplier": match_col(cands, [r"^supplier", r"vendor"]),
    "ROI": match_col(cands, [r"\bROI\b", r"return\s*on\s*investment"]),
    "Revenue": match_col(cands, [r"realized\s*revenue", r"\brevenue", r"sales\s*revenue"]),
    "COGS": match_col(cands, [r"\bCOGS\b", r"cost\s*of\s*goods"]),
    "Indirect": match_col(cands, [r"indirect\s*cost", r"overhead"]),
    "ShelfLife": match_col(cands, [r"attained\s*shelf\s*life", r"avg.*shelf.*life", r"\bshelf\s*life\b"]),
    "ServiceLevel": match_col(cands, [r"achieved\s*service\s*level", r"service\s*level"]),
    "ForecastError": match_col(cands, [r"forecast(ing)?\s*error", r"MAPE", r"bias"]),
    "ObsolescencePct": match_col(cands, [r"obsolesc(en)?ce\s*%?", r"obsolete\s*%"]),
    "CompAvail": match_col(cands, [r"component\s*availability"]),
    "ProdAvail": match_col(cands, [r"product\s*availability"]),
    "InboundUtil": match_col(cands, [r"inbound\s*warehouse.*cube\s*util", r"inbound.*util"]),
    "OutboundUtil": match_col(cands, [r"outbound\s*warehouse.*cube\s*util", r"outbound.*util"]),
    "PlanAdherence": match_col(cands, [r"production\s*plan\s*adherence", r"\bplan\s*adherence"]),
    "DeliveryReliability": match_col(cands, [r"delivery\s*reliab", r"component\s*delivery\s*reliab"]),
    "RejectionPct": match_col(cands, [r"rejection\s*%|reject\s*%"]),
    "ComponentObsoletePct": match_col(cands, [r"component\s*obsolete\s*%|obsolete\s*component\s*%"]),
    "RMCostPct": match_col(cands, [r"raw\s*material\s*cost\s*%|RM\s*cost\s*%"]),
}

with st.sidebar:
    st.markdown("### Filters & Mapping")
    with st.expander("KPI Mapper (override auto-detect)"):
        options = ["—"] + cands
        mapper = {}
        for k, v in defaults.items():
            mapper[k] = st.selectbox(k, options, index=(options.index(v) if v in options else 0))
            if mapper[k] == "—": mapper[k] = None

    round_col = mapper["round"]; date_col = mapper["date"]
    product_col = mapper["product"]; customer_col = mapper["customer"]
    component_col = mapper["component"]; supplier_col = mapper["supplier"]
    roi_col = mapper["ROI"]; revenue_col = mapper["Revenue"]
    cogs_col = mapper["COGS"]; indirect_col = mapper["Indirect"]
    shelf_life_col = mapper["ShelfLife"]; service_level_col = mapper["ServiceLevel"]
    forecast_error_col = mapper["ForecastError"]; obsolescence_pct_col = mapper["ObsolescencePct"]
    comp_avail_col = mapper["CompAvail"]; prod_avail_col = mapper["ProdAvail"]
    inb_util_col = mapper["InboundUtil"]; outb_util_col = mapper["OutboundUtil"]; plan_adherence_col = mapper["PlanAdherence"]
    deliv_rel_col = mapper["DeliveryReliability"]; rej_pct_col = mapper["RejectionPct"]
    comp_obsol_pct_col = mapper["ComponentObsoletePct"]; rm_cost_pct_col = mapper["RMCostPct"]

    def make_filter(col, label):
        if col and col in data_all.columns:
            vals = sorted(pd.Series(col=data_all[col]).dropna().unique().tolist()) if False else sorted(pd.Series(data_all[col]).dropna().unique().tolist())
            return st.multiselect(label, vals, [])
        return []

    rounds = make_filter(round_col, "Round / Week")
    products = make_filter(product_col, "Product")
    customers = make_filter(customer_col, "Customer")
    components = make_filter(component_col, "Component")
    suppliers = make_filter(supplier_col, "Supplier")

def apply_filters(df):
    d = df.copy()
    def filt(c, vals):
        nonlocal d
        if c and c in d.columns and len(vals)>0:
            d = d[d[c].isin(vals)]
    filt(round_col, rounds); filt(product_col, products); filt(customer_col, customers)
    filt(component_col, components); filt(supplier_col, suppliers)
    return d

filtered = apply_filters(data_all)

st.markdown('<div class="title">Ganga Jamuna — Executive VP Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Fresh Connection • Rounds 0–6 • Dynamic link between Functional and Financial KPIs</div>', unsafe_allow_html=True)

# --- helper chart functions with backend fallback
def chart_line_or_bar(df, x, y, title, kind="line"):
    if not x or not y or x not in df.columns or y not in df.columns: 
        st.info(f"Missing data for: {title}"); return
    d = df[[x,y]].apply(pd.to_numeric, errors="ignore").dropna()
    if len(d)==0: st.info(f"No rows for: {title}"); return
    d = d.sort_values(by=x)
    if HAS_PLOTLY:
        fig = px.line(d, x=x, y=y, markers=True, title=title) if kind=="line" else px.bar(d, x=x, y=y, title=title)
        st.plotly_chart(fig, use_container_width=True)
    elif alt:
        if kind=="line":
            chart = alt.Chart(d).mark_line(point=True).encode(x=x, y=y).properties(title=title).interactive()
        else:
            chart = alt.Chart(d).mark_bar().encode(x=x, y=y).properties(title=title).interactive()
        st.altair_chart(chart, use_container_width=True)
    else:
        st.dataframe(d, use_container_width=True)

def chart_scatter(df, x, y, color=None, title=""):
    if not x or not y or x not in df.columns or y not in df.columns: 
        st.info(f"Missing data for: {title}"); return
    cols = [x,y] + ([color] if color and color in df.columns else [])
    d = df[cols].dropna()
    if len(d)==0: st.info(f"No rows for: {title}"); return
    if HAS_PLOTLY:
        fig = px.scatter(d, x=x, y=y, color=color if color in d.columns else None, title=title)
        st.plotly_chart(fig, use_container_width=True)
    elif alt:
        enc = {'x': x, 'y': y}
        if color and color in d.columns: enc['color'] = color
        chart = alt.Chart(d).mark_point().encode(**enc).properties(title=title).interactive()
        st.altair_chart(chart, use_container_width=True)
    else:
        st.dataframe(d, use_container_width=True)

# --- overview KPI box
with st.container(border=True):
    left, right = st.columns([2,1])
    with left:
        st.markdown("**Functional KPIs**")
        st.markdown("- **Purchase** — Delivery Reliability, Rejection %, Component Obsolete %, Raw Material Cost %")
        st.markdown("- **Sales** — Attained Shelf Life, Achieved Service Level, Forecasting Error, Obsolescence %")
        st.markdown("- **Supply Chain** — Component availability, Product availability")
        st.markdown("- **Operations** — Inbound & Outbound Warehouse Cube Utilization, Production Plan Adherence %")
    with right:
        st.markdown("**Financial KPIs**")
        st.markdown("1. ROI  \n2. Realized Revenues  \n3. Cost of Goods Sold (COGS)  \n4. Indirect Cost")

# --- tabs
tab_fin, tab_sales, tab_scm, tab_ops, tab_purch = st.tabs(
    ["🏦 Financials", "🛒 Sales", "🔗 Supply Chain", "🏭 Operations", "📦 Purchasing"]
)

with tab_fin:
    st.subheader("Financial KPIs")
    xaxis = round_col if round_col else date_col
    chart_line_or_bar(filtered, xaxis, roi_col, "ROI by Round", "line")
    chart_line_or_bar(filtered, xaxis, revenue_col, "Realized Revenues by Round", "bar")
    chart_line_or_bar(filtered, xaxis, cogs_col, "COGS by Round", "bar")
    chart_line_or_bar(filtered, xaxis, indirect_col, "Indirect Cost by Round", "bar")
    chart_scatter(filtered, revenue_col, roi_col, color=product_col or customer_col, title="Revenue vs ROI")
    chart_scatter(filtered, cogs_col, roi_col, color=supplier_col or product_col, title="COGS vs ROI")

with tab_sales:
    st.subheader("VP Sales — KPI to Financial impact")
    chart_scatter(filtered, service_level_col, roi_col, color=customer_col, title="Service Level vs ROI (by Customer)")
    chart_scatter(filtered, shelf_life_col, revenue_col, color=product_col, title="Shelf Life vs Revenue (by Product)")
    chart_scatter(filtered, forecast_error_col, roi_col, color=customer_col, title="Forecast Error vs ROI")
    chart_scatter(filtered, obsolescence_pct_col, revenue_col, color=product_col, title="Obsolescence % vs Revenue")

with tab_scm:
    st.subheader("VP Supply Chain — Availability & Financials")
    xaxis = round_col if round_col else date_col
    if comp_avail_col: chart_line_or_bar(filtered, xaxis, comp_avail_col, "Component Availability by Round")
    if prod_avail_col: chart_line_or_bar(filtered, xaxis, prod_avail_col, "Product Availability by Round")

with tab_ops:
    st.subheader("VP Operations — Warehouses & Production")
    chart_scatter(filtered, inb_util_col, cogs_col, title="Inbound WH Util vs COGS")
    chart_scatter(filtered, outb_util_col, cogs_col, title="Outbound WH Util vs COGS")
    chart_scatter(filtered, plan_adherence_col, roi_col, title="Production Plan Adherence vs ROI")

with tab_purch:
    st.subheader("VP Purchasing — Supplier Performance & Financials")
    chart_scatter(filtered, deliv_rel_col, roi_col, color=supplier_col, title="Delivery Reliability vs ROI (by Supplier)")
    chart_scatter(filtered, rej_pct_col, roi_col, color=supplier_col, title="Rejection % vs ROI (by Supplier)")
    chart_scatter(filtered, rm_cost_pct_col, roi_col, color=supplier_col, title="RM Cost % vs ROI (by Supplier)")

with st.expander("📄 Data sources & detected columns"):
    meta_df = pd.DataFrame(meta)
    for c in meta_df.columns:
        meta_df[c] = meta_df[c].apply(lambda x: json.dumps(x) if isinstance(x, (list, dict)) else x)
    st.dataframe(meta_df, use_container_width=True)
    st.json(defaults)
