"""
US 3PL — Logistics Operations Command Center (multi-page)
---------------------------------------------------------
Streamlit + Plotly. Eight working pages on US_3PL_Logistics_Dataset.xlsx, built around the
three things that define a third-party-logistics view a single-warehouse dashboard misses:
  1. SLA compliance judged against EACH CLIENT'S OWN contract target
  2. Cost-to-serve and gross margin tracked PER CLIENT (some accounts lose money)
  3. Exception-first surfacing - what is breaching, not a wall of green

RUN:
    pip install streamlit pandas plotly openpyxl
    streamlit run us_3pl_dashboard.py

Workbook must sit beside this script. Data is synthetic (trailing 12 months ending 2026-06-23);
every figure shown is computed from it. Carrier names are real but cost/transit values are modeled.
"""

from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="US 3PL Operations", layout="wide", page_icon="truck")

NAVY, BLUE, TEAL, INK, GREY = "#1b2a4a", "#2f6fed", "#0ea5a4", "#1a2233", "#8a93a6"
GREEN, AMBER, RED = "#1f9d6b", "#e0922f", "#d9485b"
MODE_COLOR = {"Parcel": BLUE, "LTL": TEAL, "FTL": NAVY}

st.markdown(f"""
<style>
  .block-container {{ padding-top: 4.2rem; max-width: 1600px; }}
  section[data-testid="stSidebar"] {{ background: linear-gradient(180deg,#16223c,#1b2a4a); }}
  /* all sidebar text white (labels, radio options, selected values, caption, logo) */
  section[data-testid="stSidebar"] *:not(svg):not(path) {{ color:#eef3ff !important; }}
  .side-logo {{ color:#eef3ff !important; }}
  /* closed selectbox control: dark translucent so white text is readable on it */
  section[data-testid="stSidebar"] div[data-baseweb="select"] > div {{
      background-color: rgba(255,255,255,0.10) !important;
      border-color: rgba(255,255,255,0.28) !important; }}
  /* the dropdown menu is portaled onto a WHITE surface outside the sidebar -> keep its text dark */
  ul[role="listbox"] li, div[data-baseweb="popover"] li,
  div[data-baseweb="popover"] [role="option"] {{ color:#1a2233 !important; }}
  .side-logo {{ font-weight:800; font-size:1.05rem; letter-spacing:.4px; margin:2px 0 10px 0; }}
  .h-title {{ color:{INK}; font-size:1.9rem; font-weight:800; margin:0; }}
  .h-sub {{ color:{GREY}; margin:2px 0 8px 0; }}
  .card {{ background:#fff; border:1px solid #e7ebf3; border-radius:13px; padding:13px 15px;
           box-shadow:0 1px 3px rgba(27,42,74,.05); height:100%; min-height:104px; margin-bottom:10px; }}
  .kpi-l {{ color:{GREY}; font-size:.76rem; font-weight:600; text-transform:uppercase; letter-spacing:.4px; }}
  .kpi-v {{ color:{INK}; font-size:1.62rem; font-weight:800; line-height:1.1; }}
  .kpi-s {{ font-size:.74rem; font-weight:600; margin-top:1px; }}
  .pitch {{ background:linear-gradient(90deg,#eef3ff,#fff 70%); border-left:4px solid {BLUE};
            border-radius:10px; padding:12px 16px; color:{INK}; font-size:.92rem; line-height:1.55; margin:6px 0 14px 0; }}
  .pitch b {{ color:{BLUE}; }}
  .sect {{ color:{INK}; font-weight:800; font-size:1.06rem; margin:10px 0 4px 0; }}
  .panel {{ background:#f7f9fd; border:1px solid #e4eaf4; border-radius:13px; padding:12px 15px; margin-bottom:11px; }}
  .panel h4 {{ color:{NAVY}; margin:0 0 6px 0; font-size:.98rem; }}
  .alert {{ background:#fff; border-radius:9px; padding:8px 11px; margin:6px 0; border-left:4px solid {GREY};
            font-size:.85rem; color:{INK}; }}
  .blk {{ font-weight:800; }}
  hr {{ margin:8px 0 10px 0; }}
</style>
""", unsafe_allow_html=True)

DATA = Path(__file__).parent / "US_3PL_Logistics_Dataset.xlsx"
if not DATA.exists():
    st.error(f"Cannot find {DATA.name}. Place it beside this script and rerun.")
    st.stop()


@st.cache_data(show_spinner="Loading 3PL dataset...")
def load():
    xls = pd.ExcelFile(DATA)
    g = lambda s: pd.read_excel(xls, s)
    return dict(date=g("DimDate"), wh=g("DimWarehouse"), car=g("DimCarrier"), cl=g("DimClient"),
                orders=g("FactOrders"), ships=g("FactShipments"), inb=g("FactInbound"),
                ret=g("FactReturns"), snap=g("FactInventorySnapshot"),
                trend=g("FactInventoryTrend"), fin=g("FactFinancials"))


D = load()
CL, WH, CAR = D["cl"], D["wh"], D["car"]
CLNAME = CL.set_index("ClientKey").ClientName.to_dict()
WHNAME = WH.set_index("WarehouseKey").WarehouseName.to_dict()
CARNAME = CAR.set_index("CarrierKey").CarrierName.to_dict()
MAXKEY = int(D["orders"].DateKey.max())
MAXDATE = pd.to_datetime(str(MAXKEY), format="%Y%m%d")


def money(v):
    a = abs(v)
    if a >= 1e6: return f"${v/1e6:,.2f}M"
    if a >= 1e3: return f"${v/1e3:,.0f}k"
    return f"${v:,.0f}"


def card(col, label, value, sub="", scolor=GREY):
    col.markdown(f'<div class="card"><div class="kpi-l">{label}</div>'
                 f'<div class="kpi-v">{value}</div>'
                 f'<div class="kpi-s" style="color:{scolor}">{sub}</div></div>', unsafe_allow_html=True)


def pitch(html):
    st.markdown(f'<div class="pitch">{html}</div>', unsafe_allow_html=True)


def status_color(ok): return GREEN if ok else RED


with st.sidebar:
    st.markdown('<div class="side-logo">US 3PL - OPS CENTER</div>', unsafe_allow_html=True)
    page = st.radio("Navigate", ["Executive Overview", "Client Scorecards", "Inbound / Receiving",
                                 "Outbound / Fulfillment", "Transportation", "Inventory & Space",
                                 "Cost-to-Serve", "Returns"], label_visibility="collapsed")
    st.markdown("---")
    sel_client = st.selectbox("Client account", ["All clients"] + CL.ClientName.tolist())
    sel_wh = st.selectbox("Warehouse", ["All warehouses"] + WH.WarehouseName.tolist())
    period = st.selectbox("Period", ["Last 12 months", "Last 180 days", "Last 90 days", "Last 30 days"])
    st.caption(f"As of {MAXDATE.date()} - synthetic demo data")

days = {"Last 12 months": 365, "Last 180 days": 180, "Last 90 days": 90, "Last 30 days": 30}[period]
cutoff_key = int((MAXDATE - pd.Timedelta(days=days - 1)).strftime("%Y%m%d"))
cutoff_month = int((MAXDATE - pd.Timedelta(days=days - 1)).strftime("%Y%m"))
ckey = None if sel_client == "All clients" else CL.set_index("ClientName").ClientKey[sel_client]
wkey = None if sel_wh == "All warehouses" else WH.set_index("WarehouseName").WarehouseKey[sel_wh]


def F(df, datecol="DateKey", monthly=False):
    o = df
    if ckey is not None and "ClientKey" in o.columns: o = o[o.ClientKey == ckey]
    if wkey is not None and "WarehouseKey" in o.columns: o = o[o.WarehouseKey == wkey]
    if monthly and "MonthKey" in o.columns:
        o = o[o.MonthKey >= cutoff_month]
    elif datecol in o.columns:
        o = o[o[datecol] >= cutoff_key]
    return o


orders, ships = F(D["orders"]), F(D["ships"])
inb, ret = F(D["inb"]), F(D["ret"])
fin = F(D["fin"], monthly=True)
snap = D["snap"][:]
if ckey is not None: snap = snap[snap.ClientKey == ckey]
if wkey is not None: snap = snap[snap.WarehouseKey == wkey]
trend = D["trend"][D["trend"].DateKey >= cutoff_key]
if wkey is not None: trend = trend[trend.WarehouseKey == wkey]

if orders.empty:
    st.markdown(f'<div class="h-title">{page}</div>', unsafe_allow_html=True)
    st.warning(f"No order activity for **{sel_client}** at **{sel_wh}** in the **{period.lower()}**. "
               "Loosen a filter in the sidebar.")
    st.stop()

T_OTIF = CL.set_index("ClientKey").SLA_OTIF.to_dict()
T_DISP = CL.set_index("ClientKey").SLA_Dispatch.to_dict()
T_CYC = CL.set_index("ClientKey").SLA_CycleHrs.to_dict()
T_PICK = CL.set_index("ClientKey").SLA_PickAcc.to_dict()


def wtarget(tmap, frame):
    w = frame.groupby("ClientKey").size()
    t = pd.Series({k: tmap[k] for k in w.index})
    return float((t * w).sum() / w.sum())


PLT = dict(template="plotly_white", margin=dict(l=10, r=10, t=10, b=10))

# ===================== PAGE 1 - EXECUTIVE OVERVIEW =====================
if page == "Executive Overview":
    st.markdown('<div class="h-title">Executive Overview</div>'
                '<div class="h-sub">Network health, SLA exposure, and profitability - what needs attention first.</div>',
                unsafe_allow_html=True)
    pitch("A 3PL is paid against contracts, not averages. <b>One client quietly slipping below its SLA is a "
          "penalty and a renewal risk; one client billed below cost-to-serve is a loss you keep booking.</b> "
          "This page leads with both exposures, then the headline service and financial numbers behind them.")

    otif = ships.OTIF.mean() * 100
    perfect = orders.PerfectOrder.mean() * 100
    dispatch = orders.OnTimeDispatch.mean() * 100
    invacc = snap.InventoryAccuracyPct.mean() * 100 if len(snap) else np.nan
    prod = orders.Units.sum() / ((orders.PickTimeMin.sum() + orders.PackTimeMin.sum()) / 60)
    cost_order = orders.LaborCostUSD.mean()
    rev = fin.RevenueUSD.sum()
    margin = fin.GrossMarginUSD.sum() / rev * 100 if rev else np.nan
    otif_tgt = wtarget(T_OTIF, ships) * 100
    disp_tgt = wtarget(T_DISP, orders) * 100

    r1 = st.columns(4)
    card(r1[0], "OTIF", f"{otif:.1f}%", f"target {otif_tgt:.1f}%", status_color(otif >= otif_tgt))
    card(r1[1], "Perfect Order", f"{perfect:.1f}%", "on-time+complete+undamaged+docs", GREEN if perfect >= 90 else AMBER)
    card(r1[2], "On-time Dispatch", f"{dispatch:.1f}%", f"target {disp_tgt:.1f}%", status_color(dispatch >= disp_tgt))
    card(r1[3], "Inventory Accuracy", f"{invacc:.1f}%" if invacc == invacc else "n/a", "cycle-count based", GREEN if invacc >= 98 else AMBER)
    r2 = st.columns(4)
    card(r2[0], "Productivity", f"{prod:.0f}", "units / labor-hour", GREY)
    card(r2[1], "Cost / Order", money(cost_order), "labor to pick+pack", GREY)
    card(r2[2], "Revenue", money(rev), period.lower(), GREY)
    card(r2[3], "Gross Margin", f"{margin:.1f}%" if margin == margin else "n/a", "after cost-to-serve",
         GREEN if margin >= 15 else (AMBER if margin >= 5 else RED))
    st.markdown("<hr>", unsafe_allow_html=True)

    left, rail = st.columns([3.3, 1.3])
    with left:
        st.markdown('<div class="sect">SLA compliance by client - OTIF actual vs each contract target</div>', unsafe_allow_html=True)
        co = ships.groupby("ClientKey").OTIF.mean().mul(100)
        comp = pd.DataFrame({"Client": [CLNAME[k] for k in co.index], "OTIF": co.values,
                             "Target": [T_OTIF[k] * 100 for k in co.index]})
        comp["Meets"] = comp.OTIF >= comp.Target
        comp = comp.sort_values("OTIF")
        fig = go.Figure()
        fig.add_bar(x=comp.OTIF, y=comp.Client, orientation="h",
                    marker_color=[GREEN if m else RED for m in comp.Meets], name="OTIF actual")
        fig.add_trace(go.Scatter(x=comp.Target, y=comp.Client, mode="markers", name="Contract target",
                                 marker=dict(symbol="line-ns", color=INK, size=16, line=dict(width=2))))
        fig.update_layout(**PLT, height=360, legend=dict(orientation="h", y=1.12),
                          xaxis_title="OTIF %", xaxis_range=[max(80, comp.OTIF.min() - 4), 100])
        st.plotly_chart(fig, use_container_width=True)

        g1, g2 = st.columns(2)
        with g1:
            st.markdown('<div class="sect">OTIF trend (monthly)</div>', unsafe_allow_html=True)
            mt = ships.groupby("MonthKey").OTIF.mean().mul(100).reset_index()
            mt["Month"] = pd.to_datetime(mt.MonthKey, format="%Y%m")
            lf = px.line(mt, x="Month", y="OTIF", markers=True)
            lf.update_traces(line_color=BLUE)
            lf.add_hline(y=otif_tgt, line_dash="dash", line_color=GREY, annotation_text="wtd target", annotation_position="bottom right")
            lf.update_layout(**PLT, height=290, yaxis_title="OTIF %")
            st.plotly_chart(lf, use_container_width=True)
        with g2:
            st.markdown('<div class="sect">Revenue vs cost-to-serve (monthly)</div>', unsafe_allow_html=True)
            mf = fin.groupby("MonthKey").agg(Rev=("RevenueUSD", "sum"), Cost=("CostToServeUSD", "sum")).reset_index()
            mf["Month"] = pd.to_datetime(mf.MonthKey, format="%Y%m")
            bf = go.Figure()
            bf.add_bar(x=mf.Month, y=mf.Rev, name="Revenue", marker_color=BLUE)
            bf.add_bar(x=mf.Month, y=mf.Cost, name="Cost-to-serve", marker_color=GREY)
            bf.update_layout(**PLT, height=290, barmode="group", yaxis_title="$", legend=dict(orientation="h", y=1.14))
            st.plotly_chart(bf, use_container_width=True)

    with rail:
        breaches = comp[~comp.Meets]
        loss = fin.groupby("ClientKey").apply(lambda g: g.GrossMarginUSD.sum() / g.RevenueUSD.sum() * 100)
        lossers = [CLNAME[k] for k in loss.index if loss[k] < 0]
        worst = comp.iloc[0]
        st.markdown(f'<div class="panel"><h4>SLA exposure</h4>'
                    f'<b class="blk">{len(breaches)}</b> of {len(comp)} clients are below their OTIF contract. '
                    f'Worst: <b>{worst.Client}</b> at {worst.OTIF:.1f}% vs {worst.Target:.0f}% target.</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="panel"><h4>Margin exposure</h4>'
                    f'<b class="blk">{len(lossers)}</b> account(s) billed below cost-to-serve'
                    f'{": " + ", ".join(lossers) if lossers else ""}.</div>', unsafe_allow_html=True)
        st.markdown('<div class="panel"><h4>Alerts</h4></div>', unsafe_allow_html=True)
        rr = ret.Units.sum() / max(1, orders.Units.sum()) * 100
        for c, t in [(RED if len(breaches) else GREEN, f'<b class="blk">{len(breaches)}</b> OTIF breaches'),
                     (RED if lossers else GREEN, f'<b class="blk">{len(lossers)}</b> loss-making accounts'),
                     (AMBER if perfect < 90 else GREEN, f'Perfect order <b class="blk">{perfect:.1f}%</b>'),
                     (BLUE, f'Return rate <b class="blk">{rr:.1f}%</b>')]:
            st.markdown(f'<div class="alert" style="border-left-color:{c}">{t}</div>', unsafe_allow_html=True)

# ===================== PAGE 2 - CLIENT SCORECARDS =====================
elif page == "Client Scorecards":
    st.markdown('<div class="h-title">Client Scorecard</div>'
                '<div class="h-sub">One account against its own contract - the view you take into a QBR.</div>', unsafe_allow_html=True)
    pitch("Every client signed a different contract, so every client is graded on a different curve. "
          "<b>This scorecard measures one account against the targets it actually agreed to</b> - and pairs "
          "service with the number that decides renewals on your side: is this account profitable?")

    who = sel_client if sel_client != "All clients" else st.selectbox("Choose a client", CL.ClientName.tolist())
    k = CL.set_index("ClientName").ClientKey[who]
    od = D["orders"][D["orders"].ClientKey == k]; od = od[od.DateKey >= cutoff_key]
    sd = D["ships"][D["ships"].ClientKey == k]; sd = sd[sd.DateKey >= cutoff_key]
    fd = D["fin"][(D["fin"].ClientKey == k) & (D["fin"].MonthKey >= cutoff_month)]
    if od.empty:
        st.warning("No activity for this client in the selected period."); st.stop()
    row = CL.set_index("ClientKey").loc[k]

    otif = sd.OTIF.mean() * 100; disp = od.OnTimeDispatch.mean() * 100
    cyc = od.OrderCycleHrs.median(); pick = od.PickAccurate.mean() * 100
    rev = fd.RevenueUSD.sum(); mar = fd.GrossMarginUSD.sum() / rev * 100 if rev else np.nan
    st.markdown(f'<div class="sect">{who} - {row.Industry} - {row.Tier} tier</div>', unsafe_allow_html=True)
    k1 = st.columns(4)
    card(k1[0], "OTIF", f"{otif:.1f}%", f"target {row.SLA_OTIF*100:.0f}%", status_color(otif >= row.SLA_OTIF * 100))
    card(k1[1], "On-time Dispatch", f"{disp:.1f}%", f"target {row.SLA_Dispatch*100:.0f}%", status_color(disp >= row.SLA_Dispatch * 100))
    card(k1[2], "Order Cycle (median)", f"{cyc:.1f}h", f"target <={int(row.SLA_CycleHrs)}h", status_color(cyc <= row.SLA_CycleHrs))
    card(k1[3], "Pick Accuracy", f"{pick:.2f}%", f"target {row.SLA_PickAcc*100:.1f}%", status_color(pick >= row.SLA_PickAcc * 100))
    k2 = st.columns(4)
    card(k2[0], "Revenue", money(rev), period.lower(), GREY)
    card(k2[1], "Gross Margin", f"{mar:.1f}%" if mar == mar else "n/a", "after cost-to-serve", GREEN if mar >= 15 else (AMBER if mar >= 0 else RED))
    card(k2[2], "Orders", f"{len(od):,}", f"{od.Units.sum():,} units", GREY)
    card(k2[3], "Perfect Order", f"{od.PerfectOrder.mean()*100:.1f}%", "all-criteria orders", GREY)
    st.markdown("<hr>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sect">SLA attainment vs target</div>', unsafe_allow_html=True)
        att = pd.DataFrame({"Metric": ["OTIF", "On-time dispatch", "Pick accuracy"],
                            "Actual": [otif, disp, pick],
                            "Target": [row.SLA_OTIF * 100, row.SLA_Dispatch * 100, row.SLA_PickAcc * 100]})
        gf = go.Figure()
        gf.add_bar(x=att.Metric, y=att.Actual, marker_color=[status_color(a >= t) for a, t in zip(att.Actual, att.Target)], name="Actual")
        gf.add_trace(go.Scatter(x=att.Metric, y=att.Target, mode="markers", name="Target",
                                marker=dict(symbol="line-ew", color=INK, size=40, line=dict(width=3))))
        gf.update_layout(**PLT, height=300, yaxis_title="%", yaxis_range=[85, 100], legend=dict(orientation="h", y=1.15))
        st.plotly_chart(gf, use_container_width=True)
    with c2:
        st.markdown('<div class="sect">Monthly volume & OTIF</div>', unsafe_allow_html=True)
        mv = od.groupby("MonthKey").size().rename("Orders").reset_index()
        mo = sd.groupby("MonthKey").OTIF.mean().mul(100).rename("OTIF").reset_index()
        mv = mv.merge(mo, on="MonthKey", how="left"); mv["Month"] = pd.to_datetime(mv.MonthKey, format="%Y%m")
        vf = go.Figure()
        vf.add_bar(x=mv.Month, y=mv.Orders, name="Orders", marker_color=BLUE, opacity=.55)
        vf.add_trace(go.Scatter(x=mv.Month, y=mv.OTIF, name="OTIF %", yaxis="y2", line=dict(color=NAVY, width=3)))
        vf.update_layout(**PLT, height=300, yaxis_title="Orders",
                         yaxis2=dict(title="OTIF %", overlaying="y", side="right", range=[80, 100]),
                         legend=dict(orientation="h", y=1.15))
        st.plotly_chart(vf, use_container_width=True)
    verdict_sla = "meeting" if otif >= row.SLA_OTIF * 100 else "below"
    verdict_mar = "profitable" if mar >= 0 else "LOSS-MAKING"
    pitch(f"{who} is <b>{verdict_sla}</b> its OTIF contract ({otif:.1f}% vs {row.SLA_OTIF*100:.0f}%) and is "
          f"currently <b>{verdict_mar}</b> at {mar:.1f}% gross margin. "
          f"{'Strong service but check whether the rate card covers the cost-to-serve.' if mar < 10 else 'Healthy on both service and margin.'}")

# ===================== PAGE 3 - INBOUND / RECEIVING =====================
elif page == "Inbound / Receiving":
    st.markdown('<div class="h-title">Inbound / Receiving</div>'
                '<div class="h-sub">How fast and how cleanly freight moves from the dock into sellable stock.</div>', unsafe_allow_html=True)
    pitch("Nothing ships until it is received and put away. <b>Slow dock-to-stock hides inventory you already "
          "own, and poor ASN compliance turns receiving into guesswork</b> - both quietly throttle the outbound "
          "SLAs everyone watches.")
    if inb.empty:
        st.warning("No inbound activity for this selection."); st.stop()

    d2s = inb.DockToStockHrs.median(); asn = inb.ASNCompliant.mean() * 100
    racc = (inb.UnitsReceived.sum() - inb.UnitsDamaged.sum()) / inb.UnitsExpected.sum() * 100
    putaway = inb.PutawayLines.sum() / (inb.PutawayMin.sum() / 60)
    k = st.columns(4)
    card(k[0], "Dock-to-Stock (median)", f"{d2s:.1f}h", "target <=24h", status_color(d2s <= 24))
    card(k[1], "ASN Compliance", f"{asn:.1f}%", "advance ship notices", GREEN if asn >= 95 else AMBER)
    card(k[2], "Receiving Accuracy", f"{racc:.1f}%", "good units vs expected", GREEN if racc >= 98 else AMBER)
    card(k[3], "Putaway Productivity", f"{putaway:.0f}", "lines / labor-hour", GREY)
    st.markdown("<hr>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sect">Dock-to-stock distribution</div>', unsafe_allow_html=True)
        hf = px.histogram(inb, x="DockToStockHrs", nbins=40, color_discrete_sequence=[BLUE])
        hf.add_vline(x=24, line_dash="dash", line_color=RED, annotation_text="24h target")
        hf.update_layout(**PLT, height=300, yaxis_title="Receipts", xaxis_title="Hours")
        st.plotly_chart(hf, use_container_width=True)
    with c2:
        st.markdown('<div class="sect">ASN compliance by client</div>', unsafe_allow_html=True)
        ac = inb.groupby("ClientKey").ASNCompliant.mean().mul(100).sort_values()
        af = px.bar(x=ac.values, y=[CLNAME[k] for k in ac.index], orientation="h",
                    color=ac.values, color_continuous_scale=["#d9485b", "#e0922f", "#1f9d6b"],
                    labels={"x": "ASN compliance %", "y": ""})
        af.update_layout(**PLT, height=300, coloraxis_showscale=False)
        st.plotly_chart(af, use_container_width=True)

    st.markdown('<div class="sect">Inbound volume trend (units received, weekly)</div>', unsafe_allow_html=True)
    inb2 = inb.assign(Date=pd.to_datetime(inb.DateKey, format="%Y%m%d"))
    wk = inb2.groupby(pd.Grouper(key="Date", freq="W")).UnitsReceived.sum().reset_index()
    tf = px.area(wk, x="Date", y="UnitsReceived", color_discrete_sequence=[TEAL])
    tf.update_layout(**PLT, height=270, yaxis_title="Units received")
    st.plotly_chart(tf, use_container_width=True)

# ===================== PAGE 4 - OUTBOUND / FULFILLMENT =====================
elif page == "Outbound / Fulfillment":
    st.markdown('<div class="h-title">Outbound / Fulfillment</div>'
                '<div class="h-sub">The pick-pack-dispatch engine - speed, accuracy, and what breaks a perfect order.</div>', unsafe_allow_html=True)
    pitch("Perfect-order rate is the honest fulfillment grade because it is unforgiving: miss on time, "
          "completeness, damage, OR paperwork and the whole order fails. <b>Decomposing what breaks it tells "
          "you which lever actually moves the number.</b>")

    perfect = orders.PerfectOrder.mean() * 100; disp = orders.OnTimeDispatch.mean() * 100
    pick = orders.PickAccurate.mean() * 100; cyc = orders.OrderCycleHrs.median()
    k = st.columns(4)
    card(k[0], "Perfect Order", f"{perfect:.1f}%", "all four criteria", GREEN if perfect >= 90 else AMBER)
    card(k[1], "On-time Dispatch", f"{disp:.1f}%", "vs client cycle SLA", GREEN if disp >= 93 else AMBER)
    card(k[2], "Pick Accuracy", f"{pick:.2f}%", "lines picked right", GREEN if pick >= 99 else AMBER)
    card(k[3], "Order Cycle (median)", f"{cyc:.1f}h", "receipt to dispatch", GREY)
    st.markdown("<hr>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sect">What breaks the perfect order</div>', unsafe_allow_html=True)
        n = len(orders)
        fails = {"Late dispatch": (~orders.OnTimeDispatch).sum(), "Pick error": (~orders.PickAccurate).sum(),
                 "Incomplete": (~orders.Complete).sum(), "Damaged": orders.Damaged.sum(),
                 "Doc error": (~orders.DocsCorrect).sum()}
        ff = (pd.Series(fails).sort_values() / n * 100)
        bf = px.bar(x=ff.values, y=ff.index, orientation="h", color=ff.values,
                    color_continuous_scale=["#f0c98f", "#d9485b"], labels={"x": "% of orders failing", "y": ""})
        bf.update_layout(**PLT, height=300, coloraxis_showscale=False)
        st.plotly_chart(bf, use_container_width=True)
    with c2:
        st.markdown('<div class="sect">Productivity by warehouse (units/labor-hr)</div>', unsafe_allow_html=True)
        pw = orders.groupby("WarehouseKey").apply(
            lambda g: g.Units.sum() / ((g.PickTimeMin.sum() + g.PackTimeMin.sum()) / 60)).sort_values()
        pf = px.bar(x=pw.values, y=[WHNAME[k] for k in pw.index], orientation="h",
                    color_discrete_sequence=[BLUE], labels={"x": "Units / labor-hour", "y": ""})
        pf.update_layout(**PLT, height=300)
        st.plotly_chart(pf, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="sect">Pick accuracy by client</div>', unsafe_allow_html=True)
        pa = orders.groupby("ClientKey").PickAccurate.mean().mul(100).sort_values()
        af = px.bar(x=pa.values, y=[CLNAME[k] for k in pa.index], orientation="h",
                    color=pa.values, color_continuous_scale=["#d9485b", "#1f9d6b"],
                    labels={"x": "Pick accuracy %", "y": ""})
        af.update_layout(**PLT, height=320, coloraxis_showscale=False, xaxis_range=[97, 100])
        st.plotly_chart(af, use_container_width=True)
    with c4:
        st.markdown('<div class="sect">Order volume by channel (monthly)</div>', unsafe_allow_html=True)
        ch = orders.assign(Month=pd.to_datetime(orders.MonthKey, format="%Y%m")).groupby(["Month", "Channel"]).size().reset_index(name="Orders")
        cf = px.bar(ch, x="Month", y="Orders", color="Channel", color_discrete_map={"DTC": BLUE, "B2B": NAVY, "Retail": TEAL})
        cf.update_layout(**PLT, height=320, legend=dict(orientation="h", y=1.13))
        st.plotly_chart(cf, use_container_width=True)

# ===================== PAGE 5 - TRANSPORTATION =====================
elif page == "Transportation":
    st.markdown('<div class="h-title">Transportation & Carriers</div>'
                '<div class="h-sub">Delivery performance and freight cost across carriers, modes, and lanes.</div>', unsafe_allow_html=True)
    pitch("Once a parcel leaves the dock, the carrier owns the customer experience - but the 3PL owns the "
          "scorecard. <b>Carrier on-time performance and cost per shipment are the two levers behind OTIF and "
          "freight margin</b>, and they vary enormously by mode and zone.")
    if ships.empty:
        st.warning("No shipments for this selection."); st.stop()

    otd = ships.OnTime.mean() * 100; otif = ships.OTIF.mean() * 100
    cps = ships.ShipCostUSD.mean(); transit = ships.TransitDaysActual.mean()
    k = st.columns(4)
    card(k[0], "On-time Delivery", f"{otd:.1f}%", "actual <= planned transit", GREEN if otd >= 95 else AMBER)
    card(k[1], "OTIF", f"{otif:.1f}%", "on-time AND in-full", GREEN if otif >= 95 else AMBER)
    card(k[2], "Cost / Shipment", money(cps), "avg freight cost", GREY)
    card(k[3], "Avg Transit", f"{transit:.1f}d", "actual days", GREY)
    st.markdown("<hr>", unsafe_allow_html=True)

    st.markdown('<div class="sect">Carrier scorecard</div>', unsafe_allow_html=True)
    cs = ships.groupby("CarrierKey").agg(Shipments=("ShipmentID", "size"), OnTime=("OnTime", "mean"),
                                         OTIF=("OTIF", "mean"), CostPer=("ShipCostUSD", "mean"),
                                         Cost=("ShipCostUSD", "sum")).reset_index()
    cs["Carrier"] = cs.CarrierKey.map(CARNAME); cs["Mode"] = cs.CarrierKey.map(CAR.set_index("CarrierKey").Mode)
    cs = cs.sort_values("OnTime")
    sf = go.Figure()
    sf.add_bar(x=cs.OnTime * 100, y=cs.Carrier, orientation="h", marker_color=[MODE_COLOR[m] for m in cs.Mode],
               text=[f"{v:.1f}%" for v in cs.OnTime * 100], textposition="outside")
    sf.update_layout(**PLT, height=320, xaxis_title="On-time delivery %", xaxis_range=[80, 102])
    st.plotly_chart(sf, use_container_width=True)
    show = cs[["Carrier", "Mode", "Shipments", "OnTime", "OTIF", "CostPer", "Cost"]].sort_values("Shipments", ascending=False)
    st.dataframe(show.style.format({"OnTime": "{:.1%}", "OTIF": "{:.1%}", "CostPer": "${:.2f}", "Cost": "${:,.0f}", "Shipments": "{:,}"}),
                 use_container_width=True, hide_index=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sect">OTIF & cost by shipping zone</div>', unsafe_allow_html=True)
        zk = ships.groupby("Zone").agg(OTIF=("OTIF", "mean"), Cost=("ShipCostUSD", "mean")).reset_index()
        zf = go.Figure()
        zf.add_bar(x=zk.Zone, y=zk.OTIF * 100, name="OTIF %", marker_color=GREEN, opacity=.6)
        zf.add_trace(go.Scatter(x=zk.Zone, y=zk.Cost, name="Cost/shipment $", yaxis="y2", line=dict(color=NAVY, width=3)))
        zf.update_layout(**PLT, height=300, xaxis_title="Zone", yaxis_title="OTIF %",
                         yaxis2=dict(title="$/shipment", overlaying="y", side="right"), legend=dict(orientation="h", y=1.15))
        st.plotly_chart(zf, use_container_width=True)
    with c2:
        st.markdown('<div class="sect">Shipments by destination region</div>', unsafe_allow_html=True)
        dr = ships.groupby("DestRegion").size().sort_values()
        df_ = px.bar(x=dr.values, y=dr.index, orientation="h", color_discrete_sequence=[TEAL], labels={"x": "Shipments", "y": ""})
        df_.update_layout(**PLT, height=300)
        st.plotly_chart(df_, use_container_width=True)

# ===================== PAGE 6 - INVENTORY & SPACE =====================
elif page == "Inventory & Space":
    st.markdown('<div class="h-title">Inventory & Space</div>'
                '<div class="h-sub">Accuracy, capital tied up, and how full the buildings are.</div>', unsafe_allow_html=True)
    pitch("Space is the 3PL's scarcest asset and inventory accuracy is its credibility. <b>A building run too "
          "hot has no room to receive; one run too cold is unbilled capacity</b> - and inaccurate counts break "
          "every downstream promise.")
    if snap.empty:
        st.warning("No inventory for this selection."); st.stop()

    invval = snap.InventoryValueUSD.sum(); acc = snap.InventoryAccuracyPct.mean() * 100
    doh = (snap.OnHandUnits.sum() / max(1, orders.Units.sum() / 365))
    util = D["trend"][D["trend"].DateKey == D["trend"].DateKey.max()]
    if wkey is not None: util = util[util.WarehouseKey == wkey]
    util_pct = util.UtilizationPct.mean() * 100
    k = st.columns(4)
    card(k[0], "Inventory Value", money(invval), "capital on the floor", GREY)
    card(k[1], "Inventory Accuracy", f"{acc:.1f}%", "cycle-count based", GREEN if acc >= 98 else AMBER)
    card(k[2], "Days on Hand", f"{doh:.0f}", "at current throughput", GREY)
    card(k[3], "Slot Utilization", f"{util_pct:.0f}%", "pallet positions used",
         AMBER if util_pct >= 88 else (GREEN if util_pct >= 70 else BLUE))
    st.markdown("<hr>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sect">Slot utilization by warehouse</div>', unsafe_allow_html=True)
        u = util.copy(); u["WH"] = u.WarehouseKey.map(WHNAME)
        uf = px.bar(u.sort_values("UtilizationPct"), x="UtilizationPct", y="WH", orientation="h",
                    color="UtilizationPct", color_continuous_scale=["#2f6fed", "#1f9d6b", "#e0922f"],
                    range_color=[0.5, 1.0], labels={"UtilizationPct": "Utilization", "WH": ""})
        uf.add_vline(x=0.88, line_dash="dash", line_color=RED, annotation_text="congestion")
        uf.update_layout(**PLT, height=300, coloraxis_showscale=False, xaxis_tickformat=".0%")
        st.plotly_chart(uf, use_container_width=True)
    with c2:
        st.markdown('<div class="sect">Utilization trend (weekly)</div>', unsafe_allow_html=True)
        tt = trend.assign(Date=pd.to_datetime(trend.DateKey, format="%Y%m%d"))
        tg = tt.groupby("Date").UtilizationPct.mean().reset_index()
        tf = px.line(tg, x="Date", y="UtilizationPct")
        tf.update_traces(line_color=NAVY)
        tf.update_layout(**PLT, height=300, yaxis_tickformat=".0%", yaxis_title="Utilization")
        st.plotly_chart(tf, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="sect">Inventory value by client</div>', unsafe_allow_html=True)
        iv = snap.groupby("ClientKey").InventoryValueUSD.sum().sort_values()
        vf = px.bar(x=iv.values, y=[CLNAME[k] for k in iv.index], orientation="h",
                    color_discrete_sequence=[BLUE], labels={"x": "Inventory value ($)", "y": ""})
        vf.update_layout(**PLT, height=330)
        st.plotly_chart(vf, use_container_width=True)
    with c4:
        st.markdown('<div class="sect">Inventory accuracy by warehouse</div>', unsafe_allow_html=True)
        ia = snap.groupby("WarehouseKey").InventoryAccuracyPct.mean().mul(100).sort_values()
        iaf = px.bar(x=ia.values, y=[WHNAME[k] for k in ia.index], orientation="h",
                     color=ia.values, color_continuous_scale=["#d9485b", "#1f9d6b"],
                     range_color=[97, 100], labels={"x": "Accuracy %", "y": ""})
        iaf.update_layout(**PLT, height=330, coloraxis_showscale=False, xaxis_range=[97, 100])
        st.plotly_chart(iaf, use_container_width=True)

# ===================== PAGE 7 - COST-TO-SERVE =====================
elif page == "Cost-to-Serve":
    st.markdown('<div class="h-title">Cost-to-Serve & Profitability</div>'
                '<div class="h-sub">The 3PL P&L by client - where the money is made and lost.</div>', unsafe_allow_html=True)
    pitch("Operational excellence and profitability are not the same thing. <b>A 3PL can hit every SLA for a "
          "client it loses money on</b>, because the rate card was negotiated below the true cost-to-serve. "
          "This page puts revenue against fully-loaded cost, per account, so the unprofitable ones stop hiding.")
    if fin.empty:
        st.warning("No financials for this selection."); st.stop()

    rev = fin.RevenueUSD.sum(); cost = fin.CostToServeUSD.sum(); gm = rev - cost
    k = st.columns(4)
    card(k[0], "Revenue", money(rev), period.lower(), GREY)
    card(k[1], "Cost-to-Serve", money(cost), "labor+freight+space+OH", GREY)
    card(k[2], "Gross Margin", money(gm), "revenue - cost", GREEN if gm > 0 else RED)
    card(k[3], "Margin %", f"{gm/rev*100:.1f}%", "blended", GREEN if gm/rev >= .15 else (AMBER if gm > 0 else RED))
    st.markdown("<hr>", unsafe_allow_html=True)

    cm = fin.groupby("ClientKey").agg(Rev=("RevenueUSD", "sum"), Cost=("CostToServeUSD", "sum"),
                                      GM=("GrossMarginUSD", "sum")).reset_index()
    cm["Margin%"] = cm.GM / cm.Rev * 100; cm["Client"] = cm.ClientKey.map(CLNAME)
    c1, c2 = st.columns([3, 2])
    with c1:
        st.markdown('<div class="sect">Gross margin % by client (red = loss-making)</div>', unsafe_allow_html=True)
        cmm = cm.sort_values("Margin%")
        mf = px.bar(cmm, x="Margin%", y="Client", orientation="h",
                    color=cmm["Margin%"] > 0, color_discrete_map={True: GREEN, False: RED},
                    labels={"Margin%": "Gross margin %", "Client": ""})
        mf.add_vline(x=0, line_color=INK)
        mf.update_layout(**PLT, height=380, showlegend=False)
        st.plotly_chart(mf, use_container_width=True)
    with c2:
        st.markdown('<div class="sect">Revenue vs cost-to-serve</div>', unsafe_allow_html=True)
        sc = px.scatter(cm, x="Cost", y="Rev", size="Rev", color=cm["GM"] > 0,
                        color_discrete_map={True: GREEN, False: RED}, hover_name="Client", size_max=42,
                        labels={"Cost": "Cost-to-serve ($)", "Rev": "Revenue ($)"})
        lim = max(cm.Rev.max(), cm.Cost.max()) * 1.05
        sc.add_trace(go.Scatter(x=[0, lim], y=[0, lim], mode="lines", name="break-even", line=dict(color=GREY, dash="dash")))
        sc.update_layout(**PLT, height=380, showlegend=False)
        st.plotly_chart(sc, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="sect">Revenue by stream</div>', unsafe_allow_html=True)
        streams = {"Handling": fin.HandlingRev.sum(), "Freight billed": fin.FreightBilled.sum(),
                   "Storage": fin.StorageRev.sum(), "Receiving": fin.ReceiveRev.sum(), "Base fees": fin.BaseFee.sum()}
        s = pd.Series(streams).sort_values()
        rf = px.bar(x=s.values, y=s.index, orientation="h", color_discrete_sequence=[BLUE], labels={"x": "Revenue ($)", "y": ""})
        rf.update_layout(**PLT, height=290)
        st.plotly_chart(rf, use_container_width=True)
    with c4:
        st.markdown('<div class="sect">Cost-to-serve by component</div>', unsafe_allow_html=True)
        comp = {"Labor": fin.LaborCost.sum(), "Freight": fin.FreightCost.sum(),
                "Space": fin.SpaceCost.sum(), "Overhead": fin.OverheadCost.sum()}
        s = pd.Series(comp).sort_values()
        cf = px.bar(x=s.values, y=s.index, orientation="h", color_discrete_sequence=[GREY], labels={"x": "Cost ($)", "y": ""})
        cf.update_layout(**PLT, height=290)
        st.plotly_chart(cf, use_container_width=True)

    losers = cm[cm.GM < 0].sort_values("Margin%")
    if len(losers):
        names = ", ".join(f"{r.Client} ({r['Margin%']:.0f}%)" for _, r in losers.iterrows())
        pitch(f"<b>{len(losers)} account(s) are billed below cost-to-serve:</b> {names}. These are typically "
              f"large accounts that negotiated aggressive rates - high volume, high service, thin or negative "
              f"margin. The fix is a rate review or a cost-to-serve redesign at renewal, not more operational effort.")
    else:
        pitch("Every account in this selection is profitable. Watch the ones near the break-even line - a peak-season "
              "freight spike or a service upgrade can tip a thin-margin account negative.")

# ===================== PAGE 8 - RETURNS =====================
elif page == "Returns":
    st.markdown('<div class="h-title">Returns</div>'
                '<div class="h-sub">Reverse logistics - rate, reasons, and how much comes back resellable.</div>', unsafe_allow_html=True)
    pitch("Returns are the margin leak most dashboards ignore. <b>A return costs the inbound handling twice and "
          "often writes off the unit</b>, so the restock rate - how much comes back sellable - matters as much as "
          "the return rate itself.")
    if ret.empty:
        st.warning("No returns for this selection."); st.stop()

    rr = ret.Units.sum() / max(1, orders.Units.sum()) * 100
    proc = ret.ProcessingDays.mean(); restock = ret.Restocked.mean() * 100
    k = st.columns(4)
    card(k[0], "Return Rate", f"{rr:.1f}%", "returned vs shipped units", GREEN if rr <= 3 else AMBER)
    card(k[1], "Return Lines", f"{len(ret):,}", f"{ret.Units.sum():,} units", GREY)
    card(k[2], "Processing Time", f"{proc:.1f}d", "receipt to disposition", GREEN if proc <= 5 else AMBER)
    card(k[3], "Restock Rate", f"{restock:.0f}%", "returned to sellable", GREEN if restock >= 60 else AMBER)
    st.markdown("<hr>", unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sect">Return reasons</div>', unsafe_allow_html=True)
        rs = ret.groupby("Reason").Units.sum().sort_values()
        rf = px.bar(x=rs.values, y=rs.index, orientation="h", color_discrete_sequence=[BLUE], labels={"x": "Units", "y": ""})
        rf.update_layout(**PLT, height=300)
        st.plotly_chart(rf, use_container_width=True)
    with c2:
        st.markdown('<div class="sect">Condition on arrival</div>', unsafe_allow_html=True)
        cond = ret.groupby("Condition").Units.sum()
        cf = go.Figure(go.Pie(labels=cond.index, values=cond.values, hole=.6,
                              marker=dict(colors=[GREEN, RED, AMBER])))
        cf.update_layout(**PLT, height=300, annotations=[dict(text="condition", showarrow=False)])
        st.plotly_chart(cf, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="sect">Return rate by client</div>', unsafe_allow_html=True)
        rc = ret.groupby("ClientKey").Units.sum()
        oc = orders.groupby("ClientKey").Units.sum()
        rcl = (rc / oc * 100).dropna().sort_values()
        rcf = px.bar(x=rcl.values, y=[CLNAME[k] for k in rcl.index], orientation="h",
                     color=rcl.values, color_continuous_scale=["#1f9d6b", "#e0922f", "#d9485b"], labels={"x": "Return rate %", "y": ""})
        rcf.update_layout(**PLT, height=320, coloraxis_showscale=False)
        st.plotly_chart(rcf, use_container_width=True)
    with c4:
        st.markdown('<div class="sect">Processing time distribution</div>', unsafe_allow_html=True)
        pf = px.histogram(ret, x="ProcessingDays", nbins=9, color_discrete_sequence=[TEAL])
        pf.update_layout(**PLT, height=320, yaxis_title="Returns", xaxis_title="Processing days")
        st.plotly_chart(pf, use_container_width=True)

# ===================== HONESTY FOOTER (all pages) =====================
with st.expander("What's real here, and what would need a live feed"):
    st.markdown(
        "- **Computed from the dataset:** every KPI on every page - OTIF and on-time/in-full, perfect order and "
        "its failure breakdown, dispatch, pick accuracy, dock-to-stock, ASN compliance, productivity, inventory "
        "accuracy/value/days-on-hand, slot utilization, carrier scorecards, returns, and the full per-client P&L.\n"
        "- **Per-contract SLAs are real to the data:** each client carries its own OTIF/dispatch/cycle/pick targets, "
        "and compliance is judged against those, not a global average.\n"
        "- **Modeled, not quoted:** carrier names are real but their cost/transit values are synthetic; freight cost "
        "is a zone/weight/mode function, not a live rate API.\n"
        "- **Simplifications:** ship date = order date (no carrier pickup lag modeled); space is pallet-position based "
        "(no bin/slot geometry); cost-to-serve uses a flat overhead allocation. Each is a place a real WMS/TMS/ERP "
        "feed would replace an assumption - the schema already has the keys to wire them in.\n"
        "- **This is synthetic demo data** for a fictional 3PL; client names are invented."
    )
st.markdown(f'<div style="text-align:center;color:#9aa3bd;font-size:.8rem;padding:14px 0;">'
            f'US 3PL Operations Command Center - Streamlit + Plotly - 12 clients, 6 US DCs, '
            f'trailing 12 months to {MAXDATE.date()}</div>', unsafe_allow_html=True)