# Build a v4 package that fixes "Missing data" by aggregating by Round and joining metrics
# across sheets. Includes robust numeric parsing and a data-binding wizard.
# Output: /mnt/data/GangaJamuna_VP_Dashboard_v4.zip

import os, shutil, zipfile
from pathlib import Path

root = Path("/mnt/data/GangaJamuna_VP_Dashboard_v4")
data_dir = root / "data"
streamlit_dir = root / ".streamlit"
root.mkdir(parents=True, exist_ok=True)
data_dir.mkdir(parents=True, exist_ok=True)
streamlit_dir.mkdir(parents=True, exist_ok=True)

# Copy datasets if present
for src in ["/mnt/data/TFC_0_6.xlsx", "/mnt/data/FinanceReport (6).xlsx"]:
    if Path(src).exists():
        shutil.copy2(src, data_dir / Path(src).name)

app_py = r'''
import streamlit as st
import pandas as pd
import numpy as np
import re, json
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

# ---- Styles
st.markdown(
    """
    <style>
    .title {font-size: 38px; font-weight: 800; margin-bottom: 0rem;}
    .subtitle {font-size: 16px; opacity: 0.85; margin-bottom: 1rem;}
    .note {font-size: 13px; opacity: 0.8;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ========================= LOAD DATA =========================
DATA_DIR = Path("data")

@st.cache_data(show_spinner=False)
def load_all_xlsx(data_dir: Path):
    files = []
    if data_dir.exists():
        for p in sorted(list(data_dir.glob("*.xlsx"))):
            files.append(p)
    frames, meta = [], []
    for p in files:
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
            meta.append({"file": p.name, "error": f"Failed to read: {e}"})
    return frames, meta

frames, meta = load_all_xlsx(DATA_DIR)
if len(frames)==0:
    st.error("No Excel files found. Place .xlsx files under ./data or upload via sidebar.")
    st.stop()

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

raw = safe_concat(frames)

# ========================= COLUMN & TYPE HELPERS =========================
def match_col(candidates, patterns, default=None):
    if isinstance(patterns, str): patterns=[patterns]
    for pat in patterns:
        rx=re.compile(pat, flags=re.I)
        for c in candidates:
            if rx.search(str(c)): return c
    return default

def parse_number(x):
    """Parse numbers like '1,234', '12.3%', '€5,400.90', or '(1,234)'."""
    if pd.isna(x): return np.nan
    if isinstance(x, (int,float)): return x
    s = str(x).strip()
    if s == "": return np.nan
    # handle parentheses negatives
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    # remove currency and spaces
    s = re.sub(r"[^\d\.\,\-\+%]", "", s)
    # percent?
    is_pct = s.endswith("%")
    s = s.replace("%","")
    # remove thousand separators
    if s.count(",")>0 and s.count(".")<=1:
        s = s.replace(",", "")
    try:
        val = float(s)
        if neg: val = -val
        if is_pct: val = val  # keep percentage in 0–100 scale
        return val
    except:
        return np.nan

def to_num(series):
    return series.apply(parse_number)

cands = list(map(str, raw.columns))

defaults = {
    "round": match_col(cands, [r"^round$", r"week", r"period", r"cycle", r"game\s*round"]),
    "date": match_col(cands, [r"\bdate\b"]),
    "product": match_col(cands, [r"^product\b", r"sku", r"item"]),
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
    st.markdown("### Data & Mapping")
    st.caption("Override any auto-detected columns below.")
    options = ["—"] + cands
    mapper = {}
    for k, v in defaults.items():
        mapper[k] = st.selectbox(k, options, index=(options.index(v) if v in options else 0))
        if mapper[k] == "—": mapper[k] = None

# Unpack mapping
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

# Create a Round key even if only date exists
work = raw.copy()
if round_col and round_col in work.columns:
    work["__ROUND__"] = work[round_col]
elif date_col and date_col in work.columns:
    work["__ROUND__"] = pd.to_datetime(work[date_col], errors="coerce").dt.to_period("W").astype(str)
else:
    work["__ROUND__"] = 1  # single bucket

# ========================= AGGREGATION LAYER =========================
def agg_by_round(df):
    out = pd.DataFrame({"Round": df["__ROUND__"].astype(str)})
    out["Round"] = df["__ROUND__"].astype(str)
    g = df.groupby("__ROUND__")
    res = pd.DataFrame(index=g.size().index)
    if roi_col:       res["ROI"] = to_num(g[roi_col].mean() if roi_col in g.obj else pd.Series(dtype=float))
    if revenue_col:   res["Revenue"] = to_num(g[revenue_col].sum() if revenue_col in g.obj else pd.Series(dtype=float))
    if cogs_col:      res["COGS"] = to_num(g[cogs_col].sum() if cogs_col in g.obj else pd.Series(dtype=float))
    if indirect_col:  res["Indirect"] = to_num(g[indirect_col].sum() if indirect_col in g.obj else pd.Series(dtype=float))
    res = res.reset_index().rename(columns={"__ROUND__":"Round"})
    return res

def attach_roi_by_round(df_sub, how="left"):
    """Attach ROI by round to any dataset which has '__ROUND__'."""
    base = agg_by_round(work)
    if "Round" not in df_sub.columns and "__ROUND__" in df_sub.columns:
        df_sub = df_sub.rename(columns={"__ROUND__":"Round"})
    if "Round" not in df_sub.columns:
        df_sub["Round"] = work["__ROUND__"]
    return df_sub.merge(base[["Round","ROI","Revenue","COGS","Indirect"]], on="Round", how=how)

def group_metric(df, dim_col, val_col, agg="mean"):
    if not dim_col or not val_col or dim_col not in df.columns or val_col not in df.columns:
        return pd.DataFrame()
    d = df[[dim_col, "__ROUND__", val_col]].copy()
    d[val_col] = to_num(d[val_col])
    d = d.dropna(subset=[val_col])
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
        fig = px.scatter(data, x=x, y=y, color=color if color in data.columns else None, title=title)
        st.plotly_chart(fig, use_container_width=True)
    elif alt:
        enc={'x':x,'y':y}
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

# ========================= HEADER =========================
st.markdown('<div class="title">Ganga Jamuna — Executive VP Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Functional ↔ Financial KPIs • Fresh Connection (Rounds 0–6)</div>', unsafe_allow_html=True)

# ========================= FINANCIALS TAB =========================
tab_fin, tab_sales, tab_scm, tab_ops, tab_purch = st.tabs(
    ["🏦 Financials", "🛒 Sales", "🔗 Supply Chain", "🏭 Operations", "📦 Purchasing"]
)

with tab_fin:
    st.subheader("Financial KPIs")
    fin_by_round = agg_by_round(work)
    mk_linebar(fin_by_round.dropna(subset=["ROI"]), "Round", "ROI", "ROI by Round", "line")
    mk_linebar(fin_by_round.dropna(subset=["Revenue"]), "Round", "Revenue", "Realized Revenues by Round", "bar")
    mk_linebar(fin_by_round.dropna(subset=["COGS"]), "Round", "COGS", "COGS by Round", "bar")
    mk_linebar(fin_by_round.dropna(subset=["Indirect"]), "Round", "Indirect", "Indirect Cost by Round", "bar")
    # Relationships (by Round)
    if "Revenue" in fin_by_round.columns and "ROI" in fin_by_round.columns:
        mk_scatter(fin_by_round.dropna(subset=["Revenue","ROI"]), "Revenue", "ROI", title="Revenue vs ROI (by Round)")
    if "COGS" in fin_by_round.columns and "ROI" in fin_by_round.columns:
        mk_scatter(fin_by_round.dropna(subset=["COGS","ROI"]), "COGS", "ROI", title="COGS vs ROI (by Round)")

# ========================= SALES TAB =========================
with tab_sales:
    st.subheader("VP Sales — KPI to Financial impact")
    # Service Level vs ROI (by Customer), join on Round
    sl = group_metric(work, customer_col, service_level_col, "mean")
    sl = attach_roi_by_round(sl)
    mk_scatter(sl.dropna(subset=["Value","ROI"]).rename(columns={"Value":"Service Level"}), 
               "Service Level", "ROI", color="Dim", title="Service Level vs ROI (by Customer / Round)")
    sh = group_metric(work, product_col, shelf_life_col, "mean")
    sh = attach_roi_by_round(sh.merge(agg_by_round(work)[["Round","Revenue"]], on="Round", how="left"))
    mk_scatter(sh.dropna(subset=["Value","Revenue"]).rename(columns={"Value":"Shelf Life"}),
               "Shelf Life", "Revenue", color="Dim", title="Shelf Life vs Revenue (by Product / Round)")
    fe = group_metric(work, customer_col, forecast_error_col, "mean")
    fe = attach_roi_by_round(fe)
    mk_scatter(fe.dropna(subset=["Value","ROI"]).rename(columns={"Value":"Forecast Error"}),
               "Forecast Error", "ROI", color="Dim", title="Forecast Error vs ROI (by Customer / Round)")
    ob = group_metric(work, product_col, obsolescence_pct_col, "mean")
    ob = attach_roi_by_round(ob.merge(agg_by_round(work)[["Round","Revenue"]], on="Round", how="left"))
    mk_scatter(ob.dropna(subset=["Value","Revenue"]).rename(columns={"Value":"Obsolescence %"}),
               "Obsolescence %", "Revenue", color="Dim", title="Obsolescence % vs Revenue (by Product / Round)")

# ========================= SUPPLY CHAIN TAB =========================
with tab_scm:
    st.subheader("VP Supply Chain — Availability & Financials")
    comp = group_metric(work, component_col, comp_avail_col, "mean")
    comp = attach_roi_by_round(comp.merge(agg_by_round(work)[["Round","Revenue"]], on="Round", how="left"))
    if len(comp)>0:
        comp_agg = comp.groupby("Dim", as_index=False).agg({"Value":"mean","Revenue":"sum","ROI":"mean"}).rename(columns={"Value":"Component Avail"})
        mk_scatter(comp_agg, "Revenue", "ROI", color="Dim", title="Components — Revenue vs ROI (avg by component)")
    prod = group_metric(work, product_col, prod_avail_col, "mean")
    prod = attach_roi_by_round(prod.merge(agg_by_round(work)[["Round","Revenue"]], on="Round", how="left"))
    if len(prod)>0:
        prod_agg = prod.groupby("Dim", as_index=False).agg({"Value":"mean","Revenue":"sum","ROI":"mean"}).rename(columns={"Value":"Product Avail"})
        mk_scatter(prod_agg, "Revenue", "ROI", color="Dim", title="Products — Revenue vs ROI (avg by product)")

# ========================= OPERATIONS TAB =========================
with tab_ops:
    st.subheader("VP Operations — Warehouses & Production")
    ib = group_metric(work, "__ROUND__", inb_util_col, "mean")
    ib = attach_roi_by_round(ib.rename(columns={"Dim":"Round"}))
    mk_scatter(ib.dropna(subset=["Value","COGS"]).rename(columns={"Value":"Inbound Util"}),
               "Inbound Util", "COGS", title="Inbound WH Util vs COGS (by Round)")
    ob = group_metric(work, "__ROUND__", outb_util_col, "mean")
    ob = attach_roi_by_round(ob.rename(columns={"Dim":"Round"}))
    mk_scatter(ob.dropna(subset=["Value","COGS"]).rename(columns={"Value":"Outbound Util"}),
               "Outbound Util", "COGS", title="Outbound WH Util vs COGS (by Round)")
    pa = group_metric(work, "__ROUND__", plan_adherence_col, "mean")
    pa = attach_roi_by_round(pa.rename(columns={"Dim":"Round"}))
    mk_scatter(pa.dropna(subset=["Value","ROI"]).rename(columns={"Value":"Plan Adherence %"}),
               "Plan Adherence %", "ROI", title="Production Plan Adherence vs ROI (by Round)")

# ========================= PURCHASING TAB =========================
with tab_purch:
    st.subheader("VP Purchasing — Supplier Performance & Financials")
    dr = group_metric(work, supplier_col, deliv_rel_col, "mean")
    dr = attach_roi_by_round(dr.merge(agg_by_round(work)[["Round","COGS"]], on="Round", how="left"))
    mk_scatter(dr.groupby("Dim", as_index=False).agg({"Value":"mean","COGS":"sum","ROI":"mean"}).rename(columns={"Value":"Delivery Reliability"}),
               "Delivery Reliability", "ROI", color="Dim", title="Delivery Reliability vs ROI (avg by supplier)")
    rj = group_metric(work, supplier_col, rej_pct_col, "mean")
    rj = attach_roi_by_round(rj.merge(agg_by_round(work)[["Round","COGS"]], on="Round", how="left"))
    mk_scatter(rj.groupby("Dim", as_index=False).agg({"Value":"mean","COGS":"sum","ROI":"mean"}).rename(columns={"Value":"Rejection %"}),
               "Rejection %", "ROI", color="Dim", title="Rejection % vs ROI (avg by supplier)")
    rm = group_metric(work, supplier_col, rm_cost_pct_col, "mean")
    rm = attach_roi_by_round(rm.merge(agg_by_round(work)[["Round","COGS"]], on="Round", how="left"))
    mk_scatter(rm.groupby("Dim", as_index=False).agg({"Value":"mean","COGS":"sum","ROI":"mean"}).rename(columns={"Value":"RM Cost %"}),
               "RM Cost %", "ROI", color="Dim", title="RM Cost % vs ROI (avg by supplier)")

# ========================= DIAGNOSTICS =========================
with st.expander("📄 Data sources & mapping (diagnostics)"):
    meta_df = pd.DataFrame(meta)
    for c in meta_df.columns:
        meta_df[c] = meta_df[c].apply(lambda x: json.dumps(x) if isinstance(x, (list, dict)) else x)
    st.dataframe(meta_df, use_container_width=True)
    st.json(mapper)
    st.markdown("<div class='note'>If a chart shows little or no data, use the KPI Mapper to select the correct columns. This version aggregates by Round and joins KPIs across sheets so graphs render even when metrics live on different tabs/files.</div>", unsafe_allow_html=True)
'''
