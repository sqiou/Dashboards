"""
Warehouse & Inventory — Stock Intelligence Story
------------------------------------------------
A narrative Streamlit + Plotly dashboard built on Warehouse_Inventory_Model.xlsx
(a star schema: 5 dimensions + 5 fact tables). It walks an inventory buyer / ops
manager from "what is on my shelves and what is it worth" to "where is my cash
trapped and what must I act on today".

RUN:
    pip install streamlit pandas plotly openpyxl
    streamlit run warehouse_inventory_story.py

The workbook must sit in the same folder as this script. There is no upload prompt.

HONESTY NOTE: every figure here is derived from the dataset.
- Snapshot metrics (SKUs, stock value, fill %, stock-outs) are point-in-time as of
  the snapshot date. Transaction metrics (order fill, turns, picks/hr, return rate)
  are quarter-to-date rates over the 90-day window.
- Inventory turns is annualised from quarter COGS over the *current* stock value
  (single snapshot), so it is an approximation, not a trailing-12-month figure.
- Fill % is whole-warehouse capacity utilisation; it does NOT subset by Category/ABC
  (capacity is not modelled per category), and is held at warehouse grain by design.
- There is NO bin-level layout in the data, so no bin heatmap is shown — the
  warehouse x category matrix is value distribution, which is what the data supports.
"""

from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------------
st.set_page_config(page_title="Warehouse & Inventory — Stock Intelligence",
                   layout="wide", page_icon="📦")

# chrome (kept close to the reference) + status semantics for stock health
DARK, TEAL = "#1f2a44", "#1a9e8f"
GREEN, AMBER, RED = "#15803d", "#d97706", "#dc2626"
NAVY, STEEL = "#20324a", "#6b7a90"
STATUS_COLORS = {"Healthy": GREEN, "Below Min": AMBER,
                 "Overstock/Dead": "#7c5cc4", "Stock-out": RED}
ABC_COLORS = {"A": DARK, "B": TEAL, "C": "#9aa7b8"}
HEAT_SCALE = ["#eef4f3", "#7cc4a3", "#1a8f6f", "#0f5f48"]  # low -> high value

st.markdown(f"""
<style>
  .block-container {{ padding-top: 1.1rem; max-width: 1500px; }}
  .hero {{ border-left:6px solid {TEAL}; border-top:3px solid {TEAL};
           background:linear-gradient(90deg,#f1faf8,#fff 60%); padding:18px 24px;
           border-radius:6px; margin-bottom:8px; }}
  .hero h1 {{ color:{DARK}; font-size:2.0rem; margin:0; font-weight:800; }}
  .hero p {{ color:{TEAL}; font-weight:600; margin:3px 0 0 0; }}
  .story {{ background:#fbfcfe; border:1px solid #e3e8f0; border-left:4px solid {DARK};
            padding:12px 16px; border-radius:6px; color:{DARK}; font-size:.92rem;
            line-height:1.55; margin:8px 0 14px 0; }}
  .story b {{ color:{TEAL}; }}
  .sect {{ color:{DARK}; font-weight:800; font-size:1.35rem; margin:14px 0 2px 0; }}
  .kpi-l {{ color:#5e667d; font-size:.78rem; margin-bottom:2px; }}
  .kpi-v {{ color:{DARK}; font-size:1.7rem; font-weight:800; line-height:1.1; }}
  .kpi-s {{ font-size:.72rem; font-weight:600; margin-top:2px; }}
  hr {{ margin:10px 0 12px 0; }}
  .foot {{ text-align:center; color:#8c98a8; font-size:.8rem; padding:18px 0; }}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# Data — loaded directly, no upload. Star schema: facts joined to their dims.
# ----------------------------------------------------------------------------
DATA = Path(__file__).parent / "Warehouse_Inventory_Model.xlsx"
DAYS = 90          # quarter window the transaction facts cover
ANNUALISE = 365 / DAYS


@st.cache_data(show_spinner=False)
def load():
    xls = pd.ExcelFile(DATA)
    dim_date = pd.read_excel(xls, "DimDate")[["DateKey", "Date"]]
    prod = pd.read_excel(xls, "DimProduct")[
        ["SKUKey", "SKU", "ProductName", "Category", "SubCategory",
         "ABCClass", "SupplierKey", "MinThreshold"]]
    wh = pd.read_excel(xls, "DimWarehouse")[
        ["WarehouseKey", "WarehouseCode", "WarehouseName", "City", "TotalCapacity_m3"]]
    sup = pd.read_excel(xls, "DimSupplier")[
        ["SupplierKey", "SupplierName", "AvgLeadTimeDays", "ReliabilityPct"]]

    snap = (pd.read_excel(xls, "FactInventorySnapshot")
            .merge(prod, on="SKUKey").merge(wh, on="WarehouseKey"))
    trend = (pd.read_excel(xls, "FactInventoryTrend")
             .merge(wh, on="WarehouseKey").merge(dim_date, on="DateKey"))
    out = (pd.read_excel(xls, "FactOutbound")
           .merge(prod[["SKUKey", "Category", "ABCClass"]], on="SKUKey")
           .merge(wh[["WarehouseKey", "WarehouseName"]], on="WarehouseKey")
           .merge(dim_date, on="DateKey"))
    inb = (pd.read_excel(xls, "FactInbound")
           .merge(wh[["WarehouseKey", "WarehouseName"]], on="WarehouseKey")
           .merge(sup, on="SupplierKey").merge(dim_date, on="DateKey"))
    ret = (pd.read_excel(xls, "FactReturns")
           .merge(prod[["SKUKey", "Category", "ABCClass"]], on="SKUKey")
           .merge(wh[["WarehouseKey", "WarehouseName"]], on="WarehouseKey"))
    return dict(snap=snap, trend=trend, out=out, inb=inb, ret=ret, sup=sup, wh=wh)


if not DATA.exists():
    st.error(f"Cannot find {DATA.name}. Place the workbook in the same folder as this script and rerun.")
    st.stop()

D = load()


def pkr(v):
    if abs(v) >= 1e6:
        return f"PKR {v/1e6:,.1f}M"
    if abs(v) >= 1e3:
        return f"PKR {v/1e3:,.0f}k"
    return f"PKR {v:,.0f}"


# ----------------------------------------------------------------------------
# HERO + story framing
# ----------------------------------------------------------------------------
st.markdown("""
<div class="hero">
  <h1>Warehouse &amp; Inventory — Stock Intelligence</h1>
  <p>From "what is on my shelves and what is it worth" to "where is my cash trapped, and what do I reorder today"</p>
</div>
""", unsafe_allow_html=True)

st.markdown(
    '<div class="story">For a warehouse operation, the easy number is how many cases are on the floor. '
    'The dangerous one is <b>where the cash is</b>: stock value is only useful if it is moving. This '
    'dashboard follows the inventory in five moves — <b>how healthy the stock is</b>, <b>which categories '
    'and ABC classes hold the value</b>, <b>which lines turn fast vs sit dead</b>, <b>where space and '
    'stock-outs concentrate</b>, and <b>what to reorder or liquidate today</b>. The payoff is concrete: '
    'seeing trapped cash and reorder risk on one screen is what lets an ops team free working capital, '
    'avoid lost sales, and stop paying to store stock that no one is buying.</div>',
    unsafe_allow_html=True)

# ----- global filters -----
f1, f2, f3 = st.columns(3)
wh_opts = ["All"] + D["wh"]["WarehouseName"].tolist()
cat_opts = ["All"] + sorted(D["snap"]["Category"].unique())
abc_opts = ["All", "A", "B", "C"]
sel_wh = f1.selectbox("Warehouse", wh_opts)
sel_cat = f2.selectbox("Category", cat_opts)
sel_abc = f3.selectbox("ABC class", abc_opts)


def flt(df, use_cat=True, use_abc=True):
    out = df
    if sel_wh != "All" and "WarehouseName" in out.columns:
        out = out[out["WarehouseName"] == sel_wh]
    if use_cat and sel_cat != "All" and "Category" in out.columns:
        out = out[out["Category"] == sel_cat]
    if use_abc and sel_abc != "All" and "ABCClass" in out.columns:
        out = out[out["ABCClass"] == sel_abc]
    return out


snap = flt(D["snap"])
out = flt(D["out"])
ret = flt(D["ret"])
inb = flt(D["inb"], use_cat=False, use_abc=False)          # inbound has no product dims joined here
trend = flt(D["trend"], use_cat=False, use_abc=False)      # fill is whole-warehouse only

if snap.empty:
    st.warning("No stock matches those filters.")
    st.stop()

# ----------------------------------------------------------------------------
# KPI BAND  (8 cards, matching the source blueprint)
# ----------------------------------------------------------------------------
total_skus = snap["SKUKey"].nunique()
stock_val = snap["StockValue_PKR"].sum()

net_oh = snap.groupby("SKUKey")["OnHandQty"].sum()
stockouts = int((net_oh == 0).sum())

# fill: whole-warehouse capacity utilisation from the latest trend day (warehouse grain)
last_day = trend["DateKey"].max() if not trend.empty else None
fill = float("nan")
if last_day is not None:
    td = trend[trend["DateKey"] == last_day]
    if td["TotalCapacity_m3"].sum() > 0:
        fill = td["VolumeUsed_m3"].sum() / td["TotalCapacity_m3"].sum() * 100

ofr = out["QtyShipped"].sum() / out["QtyOrdered"].sum() * 100 if out["QtyOrdered"].sum() else float("nan")
cogs = out["LineCOGS_PKR"].sum()
turns = cogs * ANNUALISE / stock_val if stock_val else float("nan")
pph = len(out) / (out["PickTimeMin"].sum() / 60) if out["PickTimeMin"].sum() else float("nan")
return_rate = ret["QtyReturned"].sum() / out["QtyShipped"].sum() * 100 if out["QtyShipped"].sum() else float("nan")

below_min = int((snap["StockStatus"] == "Below Min").sum())
dead_val = snap.loc[snap["StockStatus"] == "Overstock/Dead", "StockValue_PKR"].sum()


def sub(text, color):
    return f'<div class="kpi-s" style="color:{color}">{text}</div>'


st.markdown('<div class="sect">Key Indicators</div><hr>', unsafe_allow_html=True)
row1, row2 = st.columns(4), st.columns(4)
cards = [
    ("Total SKUs", f"{total_skus:,}", sub("active stock-keeping units", STEEL)),
    ("Stock Value", pkr(stock_val), sub("working capital on the floor", STEEL)),
    ("Warehouse Fill", f"{fill:.1f}%" if fill == fill else "n/a",
     sub("whole-warehouse capacity used", AMBER if fill >= 85 else STEEL)),
    ("Order Fill Rate", f"{ofr:.1f}%" if ofr == ofr else "n/a",
     sub("units shipped vs ordered", GREEN if ofr >= 95 else AMBER)),
    ("Inventory Turns", f"{turns:.1f}x" if turns == turns else "n/a",
     sub("annualised, approx (1 snapshot)", STEEL)),
    ("Stock-outs Today", f"{stockouts}", sub("SKUs at zero on-hand", RED if stockouts else GREEN)),
    ("Picks / Hour", f"{pph:.1f}" if pph == pph else "n/a", sub("pick-line productivity", STEEL)),
    ("Return Rate", f"{return_rate:.1f}%" if return_rate == return_rate else "n/a",
     sub("returned vs shipped units", GREEN if return_rate <= 3 else AMBER)),
]
for col, (lab, val, s) in zip(list(row1) + list(row2), cards):
    col.markdown(f'<div class="kpi-l">{lab}</div><div class="kpi-v">{val}</div>{s}',
                 unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 1 — THE STORY IN ONE CHART: stock value by health (where the cash sits)
# ----------------------------------------------------------------------------
st.markdown('<div class="sect">1 · Where the cash sits — stock value by health</div>', unsafe_allow_html=True)
order = ["Healthy", "Below Min", "Overstock/Dead", "Stock-out"]
by_status = (snap.groupby("StockStatus")["StockValue_PKR"].sum()
             .reindex(order).fillna(0))
wf = go.Figure(go.Waterfall(
    orientation="v", measure=["relative"] * len(order) + ["total"],
    x=order + ["Total Stock Value"], y=list(by_status.values) + [0],
    text=[pkr(v) for v in by_status.values] + [pkr(stock_val)], textposition="outside",
    connector={"line": {"color": "#c9d2e0"}},
    increasing={"marker": {"color": TEAL}}, totals={"marker": {"color": DARK}}))
wf.update_layout(template="plotly_white", height=420, margin=dict(l=10, r=10, t=20, b=10),
                 yaxis_title="Stock value (PKR)")
st.plotly_chart(wf, use_container_width=True)
trapped = by_status[["Below Min", "Overstock/Dead", "Stock-out"]].sum()
healthy_pct = by_status["Healthy"] / stock_val * 100 if stock_val else 0
st.markdown(
    f'<div class="story">{healthy_pct:.0f}% of stock value is in <b>healthy</b> stock that is selling at a '
    f'sensible pace. The rest — about {pkr(trapped)} — is the part an ops team can actually act on: '
    f'<b>overstock/dead</b> is cash sitting still and racking up storage cost, <b>below-min</b> is a '
    f'reorder backlog about to bite, and <b>stock-out</b> is sales you are losing right now. Same warehouse, '
    f'three completely different fixes — liquidate, reorder, expedite.</div>',
    unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 2 — Which categories & ABC classes carry the value (Pareto)
# ----------------------------------------------------------------------------
st.markdown('<div class="sect">2 · Which categories &amp; ABC classes carry the value</div>',
            unsafe_allow_html=True)
cA, cB = st.columns([3, 2])
cat_abc = (snap.groupby(["Category", "ABCClass"])["StockValue_PKR"].sum()
           .reset_index())
cat_order = (snap.groupby("Category")["StockValue_PKR"].sum()
             .sort_values(ascending=False).index.tolist())
bar = px.bar(cat_abc, x="StockValue_PKR", y="Category", color="ABCClass", orientation="h",
             color_discrete_map=ABC_COLORS, category_orders={"Category": cat_order[::-1]},
             labels={"StockValue_PKR": "Stock value (PKR)", "Category": ""})
bar.update_layout(template="plotly_white", height=430, margin=dict(l=10, r=10, t=10, b=10),
                  legend_title="ABC")
cA.plotly_chart(bar, use_container_width=True)

# ABC Pareto on SKU count vs value share
abc_sum = snap.groupby("ABCClass").agg(val=("StockValue_PKR", "sum"),
                                       skus=("SKUKey", "nunique")).reindex(["A", "B", "C"]).fillna(0)
abc_sum["val_share"] = abc_sum["val"] / abc_sum["val"].sum() * 100
abc_sum["sku_share"] = abc_sum["skus"] / abc_sum["skus"].sum() * 100
fig_abc = go.Figure()
fig_abc.add_bar(x=abc_sum.index, y=abc_sum["val_share"], name="% of value",
                marker_color=DARK)
fig_abc.add_bar(x=abc_sum.index, y=abc_sum["sku_share"], name="% of SKUs",
                marker_color="#9aa7b8")
fig_abc.update_layout(template="plotly_white", height=430, margin=dict(l=10, r=10, t=10, b=10),
                      barmode="group", yaxis_title="%", legend=dict(orientation="h", y=1.12),
                      xaxis_title="ABC class")
cB.plotly_chart(fig_abc, use_container_width=True)
a_val = abc_sum.loc["A", "val_share"] if "A" in abc_sum.index else 0
a_sku = abc_sum.loc["A", "sku_share"] if "A" in abc_sum.index else 0
st.markdown(
    f'<div class="story">This is the inventory version of the 80/20 rule. <b>Class-A SKUs are {a_sku:.0f}% of '
    f'the catalogue but {a_val:.0f}% of the value</b> — they deserve tight min/max control, frequent counts '
    f'and first call on prime pick locations. Class-C is the long tail: lots of SKUs, little value, and the '
    f'usual home of dead stock. The category bars show where that A-value concentrates, so cycle-count effort '
    f'and safety stock land where the money actually is, not spread evenly across 4,000+ lines.</div>',
    unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 3 — The turns-vs-coverage trade-off by category
# ----------------------------------------------------------------------------
st.markdown('<div class="sect">3 · The turns-vs-coverage trade-off</div>', unsafe_allow_html=True)
cC, cD = st.columns([3, 2])
cat_val = snap.groupby("Category")["StockValue_PKR"].sum()
cat_doh = snap.groupby("Category").apply(
    lambda g: (g["DOH"] * g["StockValue_PKR"]).sum() / g["StockValue_PKR"].sum()
    if g["StockValue_PKR"].sum() else g["DOH"].mean())
cat_cogs = out.groupby("Category")["LineCOGS_PKR"].sum()
tr = pd.DataFrame({"value": cat_val, "doh": cat_doh})
tr["turns"] = (cat_cogs * ANNUALISE / cat_val).reindex(tr.index)
tr = tr.dropna().reset_index()
fig_sc = px.scatter(tr, x="doh", y="turns", size="value", color="Category",
                    hover_name="Category", size_max=55,
                    labels={"doh": "Avg days-on-hand (value-weighted)", "turns": "Inventory turns (annualised)"})
fig_sc.update_layout(template="plotly_white", height=440, margin=dict(l=10, r=10, t=10, b=10),
                     legend_title="")
cC.plotly_chart(fig_sc, use_container_width=True)

turns_bar = tr.sort_values("turns")
fig_tb = px.bar(turns_bar, x="turns", y="Category", orientation="h",
                color="turns", color_continuous_scale=HEAT_SCALE,
                labels={"turns": "Turns (annualised)", "Category": ""})
fig_tb.update_layout(template="plotly_white", height=440, margin=dict(l=10, r=10, t=10, b=10),
                     coloraxis_showscale=False)
cD.plotly_chart(fig_tb, use_container_width=True)
slow_cat = tr.sort_values("turns").iloc[0]["Category"] if not tr.empty else "n/a"
fast_cat = tr.sort_values("turns").iloc[-1]["Category"] if not tr.empty else "n/a"
st.markdown(
    f'<div class="story">Each bubble is a category; size is stock value. The corners are the story: '
    f'<b>high turns + low days-on-hand</b> (top-left) is healthy velocity — cash recycles fast. '
    f'<b>Low turns + high days-on-hand</b> (bottom-right) is where working capital gets stuck on the '
    f'shelf. Here <b>{fast_cat}</b> turns fastest and <b>{slow_cat}</b> slowest — the slow, high-value '
    f'bubbles are exactly where to cut purchase quantities or run promotions before more cash piles up. '
    f'Turns are annualised from quarter COGS over current stock value, so read them as direction, not a '
    f'trailing-12-month audit figure.</div>',
    unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 4 — Warehouse x category stock-value matrix + stock-out / space pressure
# ----------------------------------------------------------------------------
st.markdown('<div class="sect">4 · Where stock &amp; risk concentrate by site</div>', unsafe_allow_html=True)
mat = (snap.groupby(["WarehouseName", "Category"])["StockValue_PKR"].sum()
       .unstack().fillna(0))
mat = mat.loc[mat.sum(axis=1).sort_values().index]
heat = px.imshow(mat / 1e6, color_continuous_scale=HEAT_SCALE, aspect="auto",
                 labels=dict(color="Stock value (PKR M)"), text_auto=".1f")
heat.update_layout(template="plotly_white", height=360, margin=dict(l=10, r=10, t=10, b=10),
                   xaxis_title="", yaxis_title="")
heat.update_xaxes(tickangle=20)
st.plotly_chart(heat, use_container_width=True)

# per-site fill + stock-outs table (whole-warehouse, from snapshot + trend)
site = (D["snap"].groupby("WarehouseName")
        .apply(lambda g: pd.Series({
            "Stock value": g["StockValue_PKR"].sum(),
            "SKUs": g["SKUKey"].nunique(),
            "Stock-outs": int((g.groupby("SKUKey")["OnHandQty"].sum() == 0).sum()),
            "Below-min lines": int((g["StockStatus"] == "Below Min").sum()),
        })).reset_index())
tlast = D["trend"][D["trend"]["DateKey"] == D["trend"]["DateKey"].max()]
fillmap = (tlast.assign(f=tlast["VolumeUsed_m3"] / tlast["TotalCapacity_m3"] * 100)
           .set_index("WarehouseName")["f"])
site["Fill %"] = site["WarehouseName"].map(fillmap).round(1)
site["Stock value"] = site["Stock value"].map(pkr)
st.dataframe(site[["WarehouseName", "Stock value", "Fill %", "SKUs",
                   "Below-min lines", "Stock-outs"]],
             use_container_width=True, hide_index=True)
st.markdown(
    '<div class="story">Read the matrix one row at a time: each site\'s stock is concentrated in a couple '
    'of categories, and the darkest cells are where the working capital — and the space pressure — actually '
    'sits. The table beside it adds the operational read: a site running high <b>fill %</b> with rising '
    '<b>below-min</b> lines is heading for congestion <i>and</i> stock-outs at the same time, which is the '
    'worst combination — no room to receive, yet gaps on the shelf. '
    '<i>Note: the data has no bin-level layout, so this is value distribution by site, not a bin heatmap.</i></div>',
    unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 5 — Stock value trajectory over the quarter (vs period-average reference)
# ----------------------------------------------------------------------------
st.markdown('<div class="sect">5 · Stock value &amp; fill trajectory</div>', unsafe_allow_html=True)
traj = (trend.groupby("Date").agg(val=("TotalStockValue_PKR", "sum"),
                                  used=("VolumeUsed_m3", "sum"),
                                  cap=("TotalCapacity_m3", "sum")).reset_index().sort_values("Date"))
traj["fill"] = traj["used"] / traj["cap"] * 100
avg_val = traj["val"].mean()
fig_tr = go.Figure()
fig_tr.add_trace(go.Scatter(x=traj["Date"], y=traj["val"], mode="lines",
                            name="Stock value", line=dict(color=TEAL, width=3)))
fig_tr.add_trace(go.Scatter(x=traj["Date"], y=[avg_val] * len(traj), mode="lines",
                            name="Quarter average (reference)",
                            line=dict(color="#9aa7b8", width=2, dash="dash")))
fig_tr.add_trace(go.Scatter(x=traj["Date"], y=traj["fill"], mode="lines", name="Fill %",
                            yaxis="y2", line=dict(color=NAVY, width=2)))
fig_tr.update_layout(template="plotly_white", height=420, margin=dict(l=10, r=10, t=10, b=10),
                     yaxis=dict(title="Stock value (PKR)"),
                     yaxis2=dict(title="Fill %", overlaying="y", side="right",
                                 range=[0, 100], showgrid=False),
                     legend=dict(orientation="h", y=1.12))
st.plotly_chart(fig_tr, use_container_width=True)
st.markdown(
    '<div class="story">The dashed line is the quarter\'s own average stock value, used as a reference '
    'because the dataset carries no budget or plan feed — swap it for a real target the moment one exists '
    'and the chart logic is unchanged. Where the solid line runs above the dash, working capital is building '
    'up; sustained climbs with no matching sales are the early signal of overstock forming. The navy <b>fill '
    'line on the right axis</b> is the constraint: once it pushes toward the high-80s the receiving dock '
    'starts to choke, regardless of how healthy the value looks.</div>',
    unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# 6 — What to act on: free trapped cash + reorder list
# ----------------------------------------------------------------------------
st.markdown('<div class="sect">6 · What to act on today</div>', unsafe_allow_html=True)
cE, cF = st.columns([2, 3])

# (a) trapped cash you could free, by category
trap = (snap[snap["StockStatus"] == "Overstock/Dead"]
        .groupby("Category")["StockValue_PKR"].sum().sort_values())
fig_trap = px.bar(trap.reset_index(), x="StockValue_PKR", y="Category", orientation="h",
                  color="StockValue_PKR", color_continuous_scale=["#d9c7f0", "#7c5cc4"],
                  labels={"StockValue_PKR": "Overstock/dead value (PKR)", "Category": ""})
fig_trap.update_layout(template="plotly_white", height=420, margin=dict(l=10, r=10, t=10, b=10),
                       coloraxis_showscale=False)
cE.plotly_chart(fig_trap, use_container_width=True)

# (b) reorder action list — stock-out + below-min, with supplier lead time
prod_sup = D["snap"].merge(D["sup"][["SupplierKey", "SupplierName", "AvgLeadTimeDays"]],
                           on="SupplierKey", how="left")
act = flt(prod_sup)
act = act[act["StockStatus"].isin(["Stock-out", "Below Min"])].copy()
act["gap"] = act["MinThreshold"] - act["OnHandQty"]
act["urg"] = act["StockStatus"].map({"Stock-out": 0, "Below Min": 1})
act = act.sort_values(["urg", "gap"], ascending=[True, False]).head(15)
act_view = act[["SKU", "ProductName", "WarehouseName", "StockStatus",
                "OnHandQty", "MinThreshold", "AvgLeadTimeDays", "SupplierName"]].rename(
    columns={"ProductName": "Product", "WarehouseName": "Site", "StockStatus": "Status",
             "OnHandQty": "On-hand", "MinThreshold": "Min", "AvgLeadTimeDays": "Lead (d)",
             "SupplierName": "Supplier"})
cF.dataframe(act_view, use_container_width=True, hide_index=True, height=420)
total_trapped = trap.sum()
st.markdown(
    f'<div class="story">Two actions, ranked. <b>Left:</b> about {pkr(total_trapped)} is tied up in '
    f'overstock and dead stock — the categories at the top are the first candidates for markdown, bundling '
    f'or supplier return, because every day they sit is storage cost on stock that is not selling. '
    f'<b>Right:</b> the live reorder list puts stock-outs first, then the deepest below-min gaps, with '
    f'supplier lead time attached — long-lead items must be raised first or the gap is still open weeks from '
    f'now. This is the screen a buyer works from at the start of the day: free cash on the left, protect '
    f'sales on the right.</div>',
    unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# Honesty expander + footer
# ----------------------------------------------------------------------------
with st.expander("What's real here, and what would need more data"):
    st.markdown(
        "- **Fully real (from the dataset):** SKU counts, stock value, stock health status, ABC/category "
        "splits, days-on-hand, order fill rate, picks/hour, return rate, per-site fill %, the 90-day value "
        "and fill trajectory, overstock/dead value, and the reorder list with supplier lead times.\n"
        "- **Derived & labelled:** inventory turns is annualised from quarter COGS over the *current* stock "
        "value (one snapshot), so it is an approximation, not trailing-12-month. The dashed trajectory line "
        "is the quarter's own average used as a reference, not a real budget.\n"
        "- **Held at warehouse grain by design:** fill % is whole-warehouse capacity utilisation; it does "
        "not subset by Category/ABC because capacity is not modelled per category — subsetting it would be "
        "misleading.\n"
        "- **Deliberately omitted (not in the data — would be fabrication to show):** bin-level slot/heat "
        "maps, real-time labour rosters, carrying-cost % and obsolescence reserves, demand forecasts and "
        "statistical safety stock, and supplier price history. Each needs a feed this set does not include — "
        "wire those in and the matching sections drop straight onto this same star schema."
    )
st.markdown('<div class="foot">Warehouse &amp; Inventory — Stock Intelligence &bull; '
            'Built with Streamlit + Plotly &bull; figures derived from a 5-fact star schema (snapshot + '
            'trend + inbound + outbound + returns)</div>', unsafe_allow_html=True)