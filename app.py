
import streamlit as st
import pandas as pd
import numpy as np
import re
from pathlib import Path

# Try Plotly; if missing, fall back to Altair so the app still runs.
HAS_PLOTLY = True
try:
    import plotly.express as px
    import plotly.graph_objects as go
except Exception:
    HAS_PLOTLY = False
import altair as alt

st.set_page_config(page_title="Ganga Jamuna — VP Dashboard (Fresh Connection)", layout="wide")

# ---- Styles
st.markdown(
    """
    <style>
    .title {font-size: 38px; font-weight: 800; margin-bottom: 0rem;}
    .subtitle {font-size: 16px; opacity: 0.85; margin-bottom: 1rem;}
    </style>
    """,
    unsafe_allow_html=True,
)

DATA_FILES = ["data/TFC_0_6.xlsx","data/FinanceReport (6).xlsx"]

@st.cache_data(show_spinner=False)
def load_excels(files):
    frames, meta = [], []
    for f in files:
        p = Path(f)
        if not p.exists(): 
            continue
        try:
            xls = pd.ExcelFile(p)
            for s in xls.sheet_names:
                try:
                    df = xls.parse(s)
                    df["__source_file__"] = p.name
                    df["__sheet__"] = s
                    frames.append(df)
                    meta.append({"file": p.name, "sheet": s, "rows": len(df), "cols": list(df.columns)})
                except Exception as e:
                    meta.append({"file": p.name, "sheet": s, "error": str(e)})
        except Exception as e:
            meta.append({"file": p.name, "error": str(e)})
    return frames, meta

frames, meta = load_excels(DATA_FILES)
if len(frames) == 0:
    st.error("No data found. Place Excel files under ./data/. Expected: TFC_0_6.xlsx and FinanceReport (6).xlsx")
    st.stop()

def safe_concat(dfs):
    all_cols=set()
    for d in dfs: all_cols.update(list(d.columns))
    cols=list(all_cols)
    outs=[]
    for d in dfs:
        dd=d.copy()
        for c in cols:
            if c not in dd.columns: dd[c]=np.nan
        outs.append(dd[cols])
    return pd.concat(outs, ignore_index=True)

data_all = safe_concat(frames)

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
    # financial
    "ROI": match_col(cands, [r"\bROI\b", r"return\s*on\s*investment"]),
    "Revenue": match_col(cands, [r"realized\s*revenue", r"\brevenue", r"sales\s*revenue"]),
    "COGS": match_col(cands, [r"\bCOGS\b", r"cost\s*of\s*goods"]),
    "Indirect": match_col(cands, [r"indirect\s*cost", r"overhead"]),
    # sales
    "ShelfLife": match_col(cands, [r"attained\s*shelf\s*life", r"avg.*shelf.*life", r"\bshelf\s*life\b"]),
    "ServiceLevel": match_col(cands, [r"achieved\s*service\s*level", r"service\s*level"]),
    "ForecastError": match_col(cands, [r"forecast(ing)?\s*error", r"MAPE", r"bias"]),
    "ObsolescencePct": match_col(cands, [r"obsolesc(en)?ce\s*%?", r"obsolete\s*%"]),
    # SCM
    "CompAvail": match_col(cands, [r"component\s*availability"]),
    "ProdAvail": match_col(cands, [r"product\s*availability"]),
    # Ops
    "InboundUtil": match_col(cands, [r"inbound\s*warehouse.*cube\s*util", r"inbound.*util"]),
    "OutboundUtil": match_col(cands, [r"outbound\s*warehouse.*cube\s*util", r"outbound.*util"]),
    "PlanAdherence": match_col(cands, [r"production\s*plan\s*adherence", r"\bplan\s*adherence"]),
    # Purchasing
    "DeliveryReliability": match_col(cands, [r"delivery\s*reliab", r"component\s*delivery\s*reliab"]),
    "RejectionPct": match_col(cands, [r"rejection\s*%|reject\s*%"]),
    "ComponentObsoletePct": match_col(cands, [r"component\s*obsolete\s*%|obsolete\s*component\s*%"]),
    "RMCostPct": match_col(cands, [r"raw\s*material\s*cost\s*%|RM\s*cost\s*%"]),
}

with st.sidebar:
    st.markdown("### Filters")
    st.caption("Slice the dashboard for presentation")
    # mapper
    with st.expander("KPI Mapper (advanced)"):
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
            vals = sorted(pd.Series(data_all[col]).dropna().unique().tolist())
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

with st.container(border=True):
    c1, c2 = st.columns([2,1])
    with c1:
        st.markdown("**Functional KPIs**")
        st.markdown("- **Purchase** — Delivery Reliability, Rejection %, Component Obsolete %, Raw Material Cost %")
        st.markdown("- **Sales** — Attained Shelf Life, Achieved Service Level, Forecasting Error, Obsolescence %")
        st.markdown("- **Supply Chain** — Component availability, Product availability")
        st.markdown("- **Operations** — Inbound & Outbound Warehouse Cube Utilization, Production Plan Adherence %")
    with c2:
        st.markdown("**Financial KPIs**")
        st.markdown("1. ROI  \n2. Realized Revenues  \n3. Cost of Goods Sold (COGS)  \n4. Indirect Cost")

def kpi_row(metrics):
    cols = st.columns(len(metrics))
    for i, (label, value, hint) in enumerate(metrics):
        with cols[i]:
            with st.container(border=True):
                st.caption(label)
                if value is None or (isinstance(value, float) and np.isnan(value)):
                    st.metric(label="", value="—")
                else:
                    st.metric(label="", value=f"{value:,.2f}" if isinstance(value, (int,float)) else str(value))
                if hint: st.caption(f":gray[{hint}]")

def chart_line_or_bar(df, x, y, title, kind="line"):
    if not x or not y or x not in df.columns or y not in df.columns: 
        st.info(f"Missing data for: {title}"); return
    d = df[[x,y]].apply(pd.to_numeric, errors="ignore").dropna()
    if len(d)==0: st.info(f"No rows for: {title}"); return
    d = d.sort_values(by=x)
    if HAS_PLOTLY:
        import plotly.express as px
        fig = px.line(d, x=x, y=y, markers=True, title=title) if kind=="line" else px.bar(d, x=x, y=y, title=title)
        st.plotly_chart(fig, use_container_width=True)
    else:
        import altair as alt
        if kind=="line":
            chart = alt.Chart(d).mark_line(point=True).encode(x=x, y=y).properties(title=title).interactive()
        else:
            chart = alt.Chart(d).mark_bar().encode(x=x, y=y).properties(title=title).interactive()
        st.altair_chart(chart, use_container_width=True)

def chart_scatter_rel(df, x, y, color=None, size=None, title=""):
    if not x or not y or x not in df.columns or y not in df.columns: 
        st.info(f"Missing data for: {title}"); return
    cols = [x,y] + [c for c in [color,size] if c and c in df.columns]
    d = df[cols].dropna()
    if len(d)==0: st.info(f"No rows for: {title}"); return
    if HAS_PLOTLY:
        import plotly.express as px
        fig = px.scatter(d, x=x, y=y, color=color if color in d.columns else None,
                         size=size if size in d.columns else None, trendline="ols", title=title)
        st.plotly_chart(fig, use_container_width=True)
    else:
        import altair as alt
        enc = {'x': x, 'y': y}
        if color and color in d.columns: enc['color'] = color
        if size and size in d.columns: enc['size'] = size
        chart = alt.Chart(d).mark_point().encode(**enc).properties(title=title).interactive()
        reg = alt.Chart(d).transform_regression(x, y).mark_line()
        st.altair_chart((chart + reg), use_container_width=True)

def impact_heatmap(df, col_map):
    use = {k:v for k,v in col_map.items() if v and v in df.columns}
    if len(use) < 2: 
        st.info("Not enough numeric columns detected to build impact matrix."); return
    num = df[list(use.values())].apply(pd.to_numeric, errors="coerce")
    corr = num.corr().round(2)
    corr.index = [k for k,v in use.items()]; corr.columns = [k for k,v in use.items()]
    if HAS_PLOTLY:
        import plotly.express as px
        fig = px.imshow(corr, text_auto=True, aspect="auto",
                        title="Correlation heatmap (higher absolute values = stronger linear relationship)")
        st.plotly_chart(fig, use_container_width=True)
    else:
        import altair as alt
        corr_reset = corr.reset_index().melt(id_vars='index', var_name='KPI', value_name='corr')
        corr_reset = corr_reset.rename(columns={'index':'Metric'})
        chart = alt.Chart(corr_reset).mark_rect().encode(
            x='Metric:N', y='KPI:N', tooltip=['Metric','KPI','corr:Q'], color='corr:Q'
        ).properties(title="Correlation heatmap (|corr| high ⇒ strong relationship)").interactive()
        text = alt.Chart(corr_reset).mark_text(baseline='middle').encode(x='Metric:N', y='KPI:N', text='corr:Q')
        st.altair_chart(chart + text, use_container_width=True)

col_map = {
    "ROI": defaults["ROI"], "Revenues": defaults["Revenue"], "COGS": defaults["COGS"], "Indirect": defaults["Indirect"],
    "Service Level": defaults["ServiceLevel"], "Shelf Life": defaults["ShelfLife"],
    "Forecast Error": defaults["ForecastError"], "Obsolescence %": defaults["ObsolescencePct"],
    "Component Avail": defaults["CompAvail"], "Product Avail": defaults["ProdAvail"],
    "Inbound Util": defaults["InboundUtil"], "Outbound Util": defaults["OutboundUtil"], "Plan Adherence %": defaults["PlanAdherence"],
    "Delivery Reliability": defaults["DeliveryReliability"], "Rejection %": defaults["RejectionPct"],
    "Component Obsolete %": defaults["ComponentObsoletePct"], "RM Cost %": defaults["RMCostPct"]
}

st.subheader("Impact Matrix — Functional ↔ Financial KPIs")
impact_heatmap(data_all, col_map)

tab_fin, tab_sales, tab_scm, tab_ops, tab_purch = st.tabs(
    ["🏦 Financials", "🛒 Sales", "🔗 Supply Chain", "🏭 Operations", "📦 Purchasing"]
)

# FINANCIALS
with tab_fin:
    st.subheader("Financial KPIs")
    fin = []
    if defaults["ROI"]: fin.append(("ROI (avg)", data_all[defaults["ROI"]].astype(float).dropna().mean(), "Average"))
    if defaults["Revenue"]: fin.append(("Realized Revenues (sum)", data_all[defaults["Revenue"]].astype(float).dropna().sum(), "Total"))
    if defaults["COGS"]: fin.append(("COGS (sum)", data_all[defaults["COGS"]].astype(float).dropna().sum(), "Total"))
    if defaults["Indirect"]: fin.append(("Indirect Cost (sum)", data_all[defaults["Indirect"]].astype(float).dropna().sum(), "Total"))
    if len(fin)>0: kpi_row(fin)
    xaxis = defaults["round"] if defaults["round"] else defaults["date"]
    chart_line_or_bar(data_all, xaxis, defaults["ROI"], "ROI by Round", kind="line")
    chart_line_or_bar(data_all, xaxis, defaults["Revenue"], "Realized Revenues by Round", kind="bar")
    chart_line_or_bar(data_all, xaxis, defaults["COGS"], "COGS by Round", kind="bar")
    chart_line_or_bar(data_all, xaxis, defaults["Indirect"], "Indirect Cost by Round", kind="bar")
    chart_scatter_rel(data_all, defaults["Revenue"], defaults["ROI"], color=defaults["product"] or defaults["customer"], title="Revenue vs ROI")
    chart_scatter_rel(data_all, defaults["COGS"], defaults["ROI"], color=defaults["supplier"] or defaults["product"], title="COGS vs ROI")
    if defaults["Indirect"]:
        chart_scatter_rel(data_all, defaults["Indirect"], defaults["ROI"], color=defaults["customer"], title="Indirect Cost vs ROI")

# SALES
with tab_sales:
    st.subheader("VP Sales — KPI to Financial impact")
    cards=[]
    if defaults["ServiceLevel"]: cards.append(("Service Level (avg)", data_all[defaults["ServiceLevel"]].astype(float).dropna().mean(), None))
    if defaults["ShelfLife"]: cards.append(("Attained Shelf Life (avg)", data_all[defaults["ShelfLife"]].astype(float).dropna().mean(), None))
    if defaults["ForecastError"]: cards.append(("Forecast Error (avg)", data_all[defaults["ForecastError"]].astype(float).dropna().mean(), None))
    if defaults["ObsolescencePct"]: cards.append(("Obsolescence % (avg)", data_all[defaults["ObsolescencePct"]].astype(float).dropna().mean(), None))
    if len(cards)>0: kpi_row(cards)
    chart_scatter_rel(data_all, defaults["ServiceLevel"], defaults["ROI"], color=defaults["customer"], title="Service Level vs ROI (by Customer)")
    chart_scatter_rel(data_all, defaults["ShelfLife"], defaults["Revenue"], color=defaults["product"], title="Shelf Life vs Revenue (by Product)")
    chart_scatter_rel(data_all, defaults["ForecastError"], defaults["ROI"], color=defaults["customer"], title="Forecast Error vs ROI")
    chart_scatter_rel(data_all, defaults["ObsolescencePct"], defaults["Revenue"], color=defaults["product"], title="Obsolescence % vs Revenue")

# SUPPLY CHAIN
with tab_scm:
    st.subheader("VP Supply Chain — Availability & Financials")
    xaxis = defaults["round"] if defaults["round"] else defaults["date"]
    if defaults["CompAvail"]:
        chart_line_or_bar(data_all, xaxis, defaults["CompAvail"], "Component Availability by Round")
    if defaults["ProdAvail"]:
        chart_line_or_bar(data_all, xaxis, defaults["ProdAvail"], "Product Availability by Round")

# OPERATIONS
with tab_ops:
    st.subheader("VP Operations — Warehouses & Production")
    cards=[]
    if defaults["InboundUtil"]: cards.append(("Inbound WH Util (avg)", data_all[defaults["InboundUtil"]].astype(float).dropna().mean(), None))
    if defaults["OutboundUtil"]: cards.append(("Outbound WH Util (avg)", data_all[defaults["OutboundUtil"]].astype(float).dropna().mean(), None))
    if defaults["PlanAdherence"]: cards.append(("Plan Adherence % (avg)", data_all[defaults["PlanAdherence"]].astype(float).dropna().mean(), None))
    if len(cards)>0: kpi_row(cards)
    chart_scatter_rel(data_all, defaults["InboundUtil"], defaults["COGS"], title="Inbound WH Util vs COGS")
    chart_scatter_rel(data_all, defaults["OutboundUtil"], defaults["COGS"], title="Outbound WH Util vs COGS")
    chart_scatter_rel(data_all, defaults["PlanAdherence"], defaults["ROI"], title="Production Plan Adherence vs ROI")

# PURCHASING
with tab_purch:
    st.subheader("VP Purchasing — Supplier Performance & Financials")
    cards=[]
    if defaults["DeliveryReliability"]: cards.append(("Delivery Reliability (avg)", data_all[defaults["DeliveryReliability"]].astype(float).dropna().mean(), None))
    if defaults["RejectionPct"]: cards.append(("Rejection % (avg)", data_all[defaults["RejectionPct"]].astype(float).dropna().mean(), None))
    if defaults["ComponentObsoletePct"]: cards.append(("Component Obsolete % (avg)", data_all[defaults["ComponentObsoletePct"]].astype(float).dropna().mean(), None))
    if defaults["RMCostPct"]: cards.append(("Raw Material Cost % (avg)", data_all[defaults["RMCostPct"]].astype(float).dropna().mean(), None))
    if len(cards)>0: kpi_row(cards)
    chart_scatter_rel(data_all, defaults["DeliveryReliability"], defaults["ROI"], color=defaults["supplier"], title="Delivery Reliability vs ROI (by Supplier)")
    chart_scatter_rel(data_all, defaults["RejectionPct"], defaults["ROI"], color=defaults["supplier"], title="Rejection % vs ROI (by Supplier)")
    chart_scatter_rel(data_all, defaults["RMCostPct"], defaults["ROI"], color=defaults["supplier"], title="RM Cost % vs ROI (by Supplier)")

with st.expander("📄 Data sources & detected columns"):
    st.write(pd.DataFrame(meta))
    st.json(defaults)
