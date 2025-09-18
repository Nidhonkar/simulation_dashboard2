
import streamlit as st
import pandas as pd
import numpy as np
import re, json, io
from pathlib import Path

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

st.markdown("""
<style>
.title {font-size: 38px; font-weight: 800; margin-bottom: 0rem;}
.subtitle {font-size: 16px; opacity: .85; margin-bottom: 1rem;}
.kpi-card {padding: .8rem; border: 1px solid #e5e7eb; border-radius: 12px; background: #fff;}
</style>
""", unsafe_allow_html=True)

DATA_DIR = Path("data")
EXPECTED_NAMES = ["TFC_0_6.xlsx", "FinanceReport_6.xlsx", "FinanceReport (6).xlsx"]

# ---------------- Sidebar: sources ----------------
st.sidebar.markdown("### Data Source")
use_files = []
if DATA_DIR.exists():
    for n in EXPECTED_NAMES:
        p = DATA_DIR / n
        if p.exists(): use_files.append(str(p))
    for p in DATA_DIR.glob("*.xlsx"):
        if str(p) not in use_files: use_files.append(str(p))

uploaded = st.sidebar.file_uploader("Upload .xlsx files", type=["xlsx"], accept_multiple_files=True)
if uploaded:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    for uf in uploaded:
        dest = DATA_DIR / uf.name
        with open(dest, "wb") as f: f.write(uf.getbuffer())
        if str(dest) not in use_files: use_files.append(str(dest))

url = st.sidebar.text_input("Or paste a CSV/XLSX URL (optional)")
if url:
    try:
        if url.lower().endswith(".csv"):
            df_url = pd.read_csv(url)
            df_url["__source_file__"] = "URL.csv"; df_url["__sheet__"] = "csv"
            use_files.append("__FROM_URL__")
        else:
            df_url = pd.read_excel(url, sheet_name=None)
            # concatenate sheets
            frames = []
            for s,df in df_url.items():
                df["__source_file__"] = "URL.xlsx"; df["__sheet__"] = s; frames.append(df)
            df_url = pd.concat(frames, ignore_index=True)
            use_files.append("__FROM_URL__")
    except Exception as e:
        st.sidebar.error(f"URL load failed: {e}")
        df_url = None
else:
    df_url = None

if len(use_files)==0:
    st.sidebar.info("No Excel detected — using demo data (DEMO.xlsx).")
    demo = pd.read_excel(DATA_DIR / "DEMO.xlsx", sheet_name="DemoData")
    frames = [demo]
    meta = [{"file":"DEMO.xlsx","sheet":"DemoData","rows":len(demo),"cols":list(map(str,demo.columns))}]
else:
    st.sidebar.caption("**Using files:**")
    for f in use_files: st.sidebar.write(f"- {Path(f).name if f!='__FROM_URL__' else '(URL)'}")

    @st.cache_data(show_spinner=False)
    def load_excels(filepaths, df_url):
        frames, meta = [], []
        for f in filepaths:
            if f == "__FROM_URL__" and df_url is not None:
                dfu = df_url.copy()
                frames.append(dfu)
                meta.append({"file":"URL","sheet":"(all)","rows":len(dfu),"cols":list(map(str,dfu.columns))})
                continue
            p = Path(f)
            try:
                xls = pd.ExcelFile(p)
                for s in xls.sheet_names:
                    try:
                        df = xls.parse(s)
                        df.columns = [str(c) for c in df.columns]
                        df["__source_file__"] = p.name
                        df["__sheet__"] = s
                        frames.append(df)
                        meta.append({"file": p.name, "sheet": s, "rows": int(len(df)), "cols": [str(c) for c in df.columns]})
                    except Exception as e:
                        meta.append({"file": p.name, "sheet": s, "error": str(e)})
            except Exception as e:
                meta.append({"file": p.name if p.exists() else str(p), "error": f"Failed to read: {e}"})
        return frames, meta

    frames, meta = load_excels(use_files, df_url)

def safe_concat(dfs):
    if not dfs: return pd.DataFrame()
    all_cols=set()
    for d in dfs: all_cols.update(map(str, d.columns))
    cols=list(all_cols)
    outs=[]
    for d in dfs:
        dd=d.copy(); dd.columns=list(map(str, dd.columns))
        for c in cols:
            if c not in dd.columns: dd[c]=np.nan
        outs.append(dd[cols])
    return pd.concat(outs, ignore_index=True)

raw = safe_concat(frames)

# ---------------- Helpers ----------------
def match_col(candidates, patterns, default=None):
    if isinstance(patterns, str): patterns=[patterns]
    for pat in patterns:
        rx=re.compile(pat, flags=re.I)
        for c in candidates:
            if rx.search(str(c)): return c
    return default

def parse_number(x):
    if pd.isna(x): return np.nan
    if isinstance(x,(int,float)): return float(x)
    s = str(x).strip()
    if s == "": return np.nan
    neg = s.startswith("(") and s.endswith(")")
    if neg: s = s[1:-1]
    s = re.sub(r"[^\d\.\,\-\+%]", "", s)
    pct = s.endswith("%"); s = s.replace("%","")
    if s.count(",")>0 and s.count(".")<=1: s = s.replace(",","")
    try:
        v = float(s); v = -v if neg else v
        return v if not pct else v
    except: return np.nan

def to_num(series): return series.apply(parse_number)

cands = list(map(str, raw.columns))
defaults = {
    "round": match_col(cands, [r"^round$", r"week", r"period", r"cycle", r"game\s*round"]),
    "date": match_col(cands, [r"\bdate\b"]),
    "product": match_col(cands, [r"^product\b", r"sku", r"item"]),
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
    st.markdown("### KPI Mapper (override auto-detect)")
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

# Round key
work = raw.copy()
if round_col and round_col in work.columns:
    work["__ROUND__"] = work[round_col]
elif date_col and date_col in work.columns:
    work["__ROUND__"] = pd.to_datetime(work[date_col], errors="coerce").dt.to_period("W").astype(str)
else:
    work["__ROUND__"] = 1

# Derive ROI if missing
if not roi_col and (revenue_col or cogs_col or indirect_col):
    if revenue_col and revenue_col in work.columns and (cogs_col in work.columns or indirect_col in work.columns):
        rev = to_num(work[revenue_col]) if revenue_col in work.columns else 0.0
        cgs = to_num(work[cogs_col]) if cogs_col in work.columns else 0.0
        ind = to_num(work[indirect_col]) if indirect_col in work.columns else 0.0
        denom = (cgs.fillna(0) + ind.fillna(0)).replace({0: np.nan})
        work["__ROI_DERIVED__"] = (rev - cgs - ind) / denom * 100.0
        roi_col = "__ROI_DERIVED__"

def agg_by_round(df):
    g = df.groupby("__ROUND__")
    res = pd.DataFrame(index=g.size().index)
    if roi_col and roi_col in df.columns:       res["ROI"] = to_num(g[roi_col].mean())
    if revenue_col and revenue_col in df.columns:   res["Revenue"] = to_num(g[revenue_col].sum())
    if cogs_col and cogs_col in df.columns:      res["COGS"] = to_num(g[cogs_col].sum())
    if indirect_col and indirect_col in df.columns: res["Indirect"] = to_num(g[indirect_col].sum())
    res = res.reset_index().rename(columns={"__ROUND__":"Round"})
    return res

def attach_finance_by_round(df_sub, how="left"):
    base = agg_by_round(work)
    if "Round" not in df_sub.columns and "__ROUND__" in df_sub.columns:
        df_sub = df_sub.rename(columns={"__ROUND__":"Round"})
    if "Round" not in df_sub.columns:
        df_sub["Round"] = work["__ROUND__"]
    return df_sub.merge(base, on="Round", how=how)

def group_metric(df, dim_col, val_col, agg="mean"):
    if not dim_col or not val_col or dim_col not in df.columns or val_col not in df.columns:
        return pd.DataFrame()
    d = df[[dim_col, "__ROUND__", val_col]].copy()
    d[val_col] = to_num(d[val_col]); d = d.dropna(subset=[val_col])
    d = d.rename(columns={dim_col:"Dim"})
    if agg=="mean":
        out = d.groupby(["__ROUND__","Dim"], as_index=False)[val_col].mean()
    else:
        out = d.groupby(["__ROUND__","Dim"], as_index=False)[val_col].sum()
    out = out.rename(columns={"__ROUND__":"Round", val_col:"Value"})
    return out

def mk_scatter(data, x, y, color=None, title=""):
    if len(data)==0:
        st.info(f"No rows for: {title}"); return
    if HAS_PLOTLY:
        fig = px.scatter(data, x=x, y=y, color=color if (color and color in data.columns) else None, title=title)
        st.plotly_chart(fig, use_container_width=True)
    elif alt:
        enc={'x':x,'y':y}; 
        if color and color in data.columns: enc['color']=color
        chart = alt.Chart(data).mark_point().encode(**enc).properties(title=title).interactive()
        st.altair_chart(chart, use_container_width=True)
    else:
        st.dataframe(data, use_container_width=True)

def mk_linebar(data, x, y, title, kind="line"):
    if len(data)==0 or x not in data.columns or y not in data.columns:
        st.info(f"Missing data for: {title}"); return
    if HAS_PLOTLY:
        fig = px.line(data, x=x, y=y, markers=True, title=title) if kind=="line" else px.bar(data, x=x, y=y, title=title)
        st.plotly_chart(fig, use_container_width=True)
    elif alt:
        if kind=="line":
            chart = alt.Chart(data).mark_line(point=True).encode(x=x, y=y).properties(title=title).interactive()
        else:
            chart = alt.Chart(data).mark_bar().encode(x=x, y=y).properties(title=title).interactive()
        st.altair_chart(chart, use_container_width=True)
    else:
        st.dataframe(data[[x,y]], use_container_width=True)

# ======== Header and KPIs ========
st.markdown('<div class="title">Ganga Jamuna — Executive VP Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Functional ↔ Financial KPIs • Rounds 0–6</div>', unsafe_allow_html=True)

fin_by_round = agg_by_round(work)
cols = st.columns(4)
def metric_box(c, title, val, fmt="{:,.2f}", suffix=""):
    with c:
        st.markdown('<div class="kpi-card">', unsafe_allow_html=True)
        st.caption(title)
        if val is None or pd.isna(val): st.metric("", "—")
        else: st.metric("", (fmt.format(val)+suffix))
        st.markdown('</div>', unsafe_allow_html=True)
metric_box(cols[0], "Avg ROI", fin_by_round["ROI"].mean() if "ROI" in fin_by_round else None, "{:,.1f}", "%")
metric_box(cols[1], "Total Revenue", fin_by_round["Revenue"].sum() if "Revenue" in fin_by_round else None)
metric_box(cols[2], "Total COGS", fin_by_round["COGS"].sum() if "COGS" in fin_by_round else None)
metric_box(cols[3], "Total Indirect", fin_by_round["Indirect"].sum() if "Indirect" in fin_by_round else None)

tab_fin, tab_sales, tab_scm, tab_ops, tab_purch = st.tabs(
    ["🏦 Financials", "🛒 Sales", "🔗 Supply Chain", "🏭 Operations", "📦 Purchasing"]
)

with tab_fin:
    st.subheader("Financial KPIs")
    mk_linebar(fin_by_round.dropna(subset=["ROI"]), "Round", "ROI", "ROI by Round", "line")
    mk_linebar(fin_by_round.dropna(subset=["Revenue"]), "Round", "Revenue", "Realized Revenues by Round", "bar")
    mk_linebar(fin_by_round.dropna(subset=["COGS"]), "Round", "COGS", "COGS by Round", "bar")
    mk_linebar(fin_by_round.dropna(subset=["Indirect"]), "Round", "Indirect", "Indirect Cost by Round", "bar")
    if {"Revenue","ROI"}.issubset(fin_by_round.columns):
        mk_scatter(fin_by_round.dropna(subset=["Revenue","ROI"]), "Revenue", "ROI", title="Revenue vs ROI (by Round)")
    if {"COGS","ROI"}.issubset(fin_by_round.columns):
        mk_scatter(fin_by_round.dropna(subset=["COGS","ROI"]), "COGS", "ROI", title="COGS vs ROI (by Round)")

with tab_sales:
    st.subheader("VP Sales — KPI to Financial impact")
    sl = group_metric(work, "Customer", "Product Average Achieved Service Level", "mean")
    sl = attach_finance_by_round(sl)
    mk_scatter(sl.dropna(subset=["Value","ROI"]).rename(columns={"Value":"Service Level"}),
               "Service Level", "ROI", color="Dim", title="Service Level vs ROI (by Customer / Round)")
    sh = group_metric(work, "Product", "Product Average Attained Shelf Life", "mean")
    sh = attach_finance_by_round(sh)
    mk_scatter(sh.dropna(subset=["Value","Revenue"]).rename(columns={"Value":"Shelf Life"}),
               "Shelf Life", "Revenue", color="Dim", title="Shelf Life vs Revenue (by Product / Round)")
    fe = group_metric(work, "Customer", "Product Average Forecasting Error", "mean")
    fe = attach_finance_by_round(fe)
    mk_scatter(fe.dropna(subset=["Value","ROI"]).rename(columns={"Value":"Forecast Error"}),
               "Forecast Error", "ROI", color="Dim", title="Forecast Error vs ROI (by Customer / Round)")
    ob = group_metric(work, "Product", "Product Obsolescence percent", "mean")
    ob = attach_finance_by_round(ob)
    mk_scatter(ob.dropna(subset=["Value","Revenue"]).rename(columns={"Value":"Obsolescence %"}),
               "Obsolescence %", "Revenue", color="Dim", title="Obsolescence % vs Revenue (by Product / Round)")

with tab_scm:
    st.subheader("VP Supply Chain — Availability & Financials")
    comp = group_metric(work, "Supplier", "Component availability", "mean")
    comp = attach_finance_by_round(comp)
    if len(comp)>0:
        comp_agg = comp.groupby("Dim", as_index=False).agg({"Value":"mean","Revenue":"sum","ROI":"mean"}).rename(columns={"Value":"Component Avail"})
        mk_scatter(comp_agg, "Revenue", "ROI", color="Dim", title="Components — Revenue vs ROI (avg by component)")
    prod = group_metric(work, "Product", "product availability", "mean")
    prod = attach_finance_by_round(prod)
    if len(prod)>0:
        prod_agg = prod.groupby("Dim", as_index=False).agg({"Value":"mean","Revenue":"sum","ROI":"mean"}).rename(columns={"Value":"Product Avail"})
        mk_scatter(prod_agg, "Revenue", "ROI", color="Dim", title="Products — Revenue vs ROI (avg by product)")

with tab_ops:
    st.subheader("VP Operations — Warehouses & Production")
    ib = group_metric(work, "__ROUND__", "Inbound warehouse cube utilization", "mean")
    ib = attach_finance_by_round(ib.rename(columns={"Dim":"Round"}))
    mk_scatter(ib.dropna(subset=["Value","COGS"]).rename(columns={"Value":"Inbound Util"}),
               "Inbound Util", "COGS", title="Inbound WH Util vs COGS (by Round)")
    ob = group_metric(work, "__ROUND__", "Outbound warehouse cube utilization", "mean")
    ob = attach_finance_by_round(ob.rename(columns={"Dim":"Round"}))
    mk_scatter(ob.dropna(subset=["Value","COGS"]).rename(columns={"Value":"Outbound Util"}),
               "Outbound Util", "COGS", title="Outbound WH Util vs COGS (by Round)")
    pa = group_metric(work, "__ROUND__", "Production Plan Adherence Percentage", "mean")
    pa = attach_finance_by_round(pa.rename(columns={"Dim":"Round"}))
    mk_scatter(pa.dropna(subset=["Value","ROI"]).rename(columns={"Value":"Plan Adherence %"}),
               "Plan Adherence %", "ROI", title="Production Plan Adherence vs ROI (by Round)")

with tab_purch:
    st.subheader("VP Purchasing — Supplier Performance & Financials")
    dr = group_metric(work, "Supplier", "Component Delivery Reliability", "mean")
    dr = attach_finance_by_round(dr)
    mk_scatter(dr.groupby("Dim", as_index=False).agg({"Value":"mean","ROI":"mean"}).rename(columns={"Value":"Delivery Reliability"}),
               "Delivery Reliability", "ROI", color="Dim", title="Delivery Reliability vs ROI (avg by supplier)")
    rj = group_metric(work, "Supplier", "Component Rejection percentage", "mean")
    rj = attach_finance_by_round(rj)
    mk_scatter(rj.groupby("Dim", as_index=False).agg({"Value":"mean","ROI":"mean"}).rename(columns={"Value":"Rejection %"}),
               "Rejection %", "ROI", color="Dim", title="Rejection % vs ROI (avg by supplier)")
    rm = group_metric(work, "Supplier", "Raw Material Cost %", "mean")
    rm = attach_finance_by_round(rm)
    mk_scatter(rm.groupby("Dim", as_index=False).agg({"Value":"mean","ROI":"mean"}).rename(columns={"Value":"RM Cost %"}),
               "RM Cost %", "ROI", color="Dim", title="RM Cost % vs ROI (avg by supplier)")
