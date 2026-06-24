"""
Supply Chain Control Tower — the CEO's live view
-------------------------------------------------
Streamlit + Plotly. A command-center UI: dark status header, traffic-light KPI tiles
(colour = status against target), funnel, customer scorecard matrix, live exceptions,
OTIF trend, and network performance. Built on Control_Tower_Dataset.xlsx.

RUN:
    pip install streamlit pandas plotly openpyxl
    streamlit run control_tower_dashboard.py

Workbook must sit beside this script. No upload prompt. Synthetic demo data; every
figure is computed from it. Colour legend: green = on track, amber = at risk, red = breach.
"""

from pathlib import Path
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Supply Chain Control Tower", layout="wide", page_icon="🛰️")

NAVY, SLATE, GREY = "#11203a", "#1f2a44", "#6e7891"
GREEN, AMBER, RED, BLUE = "#1f9d6b", "#e0a13a", "#e5556b", "#3a6df0"
TINT = {"green": "#e6f6ee", "amber": "#fdf3e2", "red": "#fdecee"}

st.markdown(f"""
<style>
  .stApp {{ background:#eef1f6; }}
  .block-container {{ padding-top: 2rem; max-width: 1560px; }}
  .cmd {{ background:linear-gradient(100deg,{NAVY},#1d3357); color:#eaf0fb; border-radius:12px;
          padding:18px 26px; margin-bottom:14px; }}
  .cmd h1 {{ margin:0; font-size:1.85rem; font-weight:800; color:#fff; }}
  .cmd p {{ margin:3px 0 10px 0; color:#9fb2d4; font-size:.86rem; }}
  .chip {{ display:inline-block; background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.15);
           border-radius:20px; padding:5px 13px; margin-right:8px; font-size:.82rem; font-weight:600; }}
  .dot {{ display:inline-block; width:9px; height:9px; border-radius:50%; margin-right:6px; vertical-align:middle; }}
  .tile {{ background:#fff; border:1px solid #e5e9f2; border-left:5px solid {GREY}; border-radius:10px;
           padding:12px 16px; position:relative; box-shadow:0 1px 2px rgba(17,32,58,.05); }}
  .t-dot {{ position:absolute; top:14px; right:14px; width:10px; height:10px; border-radius:50%; }}
  .t-l {{ color:{GREY}; font-size:.76rem; }}
  .t-v {{ color:{SLATE}; font-size:1.7rem; font-weight:800; line-height:1.05; }}
  .t-s {{ color:{GREY}; font-size:.74rem; }}
  .sect {{ color:{SLATE}; font-weight:800; font-size:1.15rem; margin:14px 0 4px 0; }}
  .pitch {{ background:#fff; border:1px solid #e5e9f2; border-left:4px solid {BLUE};
            border-radius:10px; padding:12px 16px; color:{SLATE}; font-size:.9rem; line-height:1.55;
            margin:6px 0 12px 0; }}
  .pitch b {{ color:{BLUE}; }}
  .foot {{ text-align:center; color:#9aa3bd; font-size:.8rem; padding:16px 0; }}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
DATA = Path(__file__).parent / "Control_Tower_Dataset.xlsx"
if not DATA.exists():
    st.error(f"Cannot find {DATA.name}. Place it beside this script and rerun.")
    st.stop()


@st.cache_data(show_spinner=False)
def load():
    xls = pd.ExcelFile(DATA)
    return (pd.read_excel(xls, "Dim_Customer"), pd.read_excel(xls, "Fact_Shipments"),
            pd.read_excel(xls, "Fact_Exceptions"), pd.read_excel(xls, "Fact_Funnel"),
            pd.read_excel(xls, "Fact_OTIFTrend", parse_dates=["Date"]), pd.read_excel(xls, "Dim_Hub"))


cust, ship, exc, funnel, trend, hub = load()


def stat(v, good, warn, higher=True):
    if higher:
        return GREEN if v >= good else (AMBER if v >= warn else RED)
    return GREEN if v <= good else (AMBER if v <= warn else RED)


def tile(col, label, value, color, sub=""):
    col.markdown(f'<div class="tile" style="border-left-color:{color}">'
                 f'<div class="t-dot" style="background:{color}"></div>'
                 f'<div class="t-l">{label}</div><div class="t-v">{value}</div>'
                 f'<div class="t-s">{sub}</div></div>', unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# Region filter (applies to live operational panels)
# ----------------------------------------------------------------------------
region = st.selectbox("Region focus", ["All Regions"] + sorted(ship.Region.unique()), index=0)
ship_f = ship if region == "All Regions" else ship[ship.Region == region]
exc_f = exc if region == "All Regions" else exc[exc.Region == region]
hub_f = hub if region == "All Regions" else hub[hub.Region == region]

# ----------------------------------------------------------------------------
# KPIs
# ----------------------------------------------------------------------------
active = int((ship_f.Status != "Delivered").sum())
delivered = ship_f[ship_f.Status == "Delivered"]
ontime = delivered.OnTimeFlag.mean() * 100 if len(delivered) else 0
crit = int(((exc_f.Severity == "Critical") & (exc_f.Resolution == "Open")).sum())
fulfilled = int(len(delivered))
rev_risk = exc_f[(exc_f.Severity == "Critical") & (exc_f.Resolution == "Open")].ValueUSD.sum()
sla = (cust.SLACompliancePct * cust.VolumeOrders).sum() / cust.VolumeOrders.sum()
csat = (cust.CSAT * cust.VolumeOrders).sum() / cust.VolumeOrders.sum()
por = trend.PerfectOrderRate.iloc[-1]

# ----------------------------------------------------------------------------
# Command header with live status chips
# ----------------------------------------------------------------------------
def chip(color, txt):
    return f'<span class="chip"><span class="dot" style="background:{color}"></span>{txt}</span>'


st.markdown(
    f'<div class="cmd"><h1>🛰️ Supply Chain Control Tower</h1>'
    f'<p>One live view of every shipment, order, and exception — {region}. '
    f'Legend: <span class="dot" style="background:{GREEN}"></span>on track '
    f'<span class="dot" style="background:{AMBER}"></span>at risk '
    f'<span class="dot" style="background:{RED}"></span>breach</p>'
    + chip(stat(ontime, 95, 90), f"On-Time {ontime:.1f}%")
    + chip(stat(crit, 5, 15, higher=False), f"{crit} Critical Exceptions")
    + chip(stat(sla, 97, 94), f"SLA {sla:.1f}%")
    + chip(BLUE, f"{active:,} Active Shipments")
    + '</div>', unsafe_allow_html=True)

st.markdown('<div class="pitch">This is the CEO\'s morning screen. Every other dashboard answers one '
            'question; <b>the control tower answers "is anything on fire right now?"</b> — and colour is '
            'the whole language: a wall of green means the network is running itself, a single red tile '
            'is where leadership attention should go before lunch.</div>', unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# Traffic-light KPI tiles
# ----------------------------------------------------------------------------
r1 = st.columns(4)
tile(r1[0], "Active Shipments", f"{active:,}", BLUE, "in the network now")
tile(r1[1], "On-Time Delivery", f"{ontime:.1f}%", stat(ontime, 95, 90), "target ≥ 95%")
tile(r1[2], "Critical Exceptions", f"{crit}", stat(crit, 5, 15, higher=False), "open & unresolved")
tile(r1[3], "Orders Fulfilled Today", f"{fulfilled:,}", BLUE, "delivered in last 24h")
st.write("")
r2 = st.columns(4)
tile(r2[0], "Revenue at Risk", f"${rev_risk/1e3:,.0f}k", stat(rev_risk, 50e3, 150e3, higher=False), "tied to open criticals")
tile(r2[1], "CSAT Score", f"{csat:.1f} / 5", stat(csat, 4.5, 4.0), "volume-weighted")
tile(r2[2], "Perfect Order Rate", f"{por:.1f}%", stat(por, 92, 88), "on-time, complete, undamaged")
tile(r2[3], "SLA Compliance", f"{sla:.1f}%", stat(sla, 97, 94), "target ≥ 97%")
st.write("")

# ----------------------------------------------------------------------------
# Row: Order-to-Delivery funnel  |  OTIF & Perfect Order trend
# ----------------------------------------------------------------------------
c1, c2 = st.columns(2)
with c1:
    st.markdown('<div class="sect">Order-to-Delivery Flow</div>', unsafe_allow_html=True)
    ff = go.Figure(go.Funnel(y=funnel.Stage, x=funnel.Orders,
                             marker=dict(color=[NAVY, "#28406b", "#3a6df0", "#5b86e8", GREEN]),
                             textinfo="value+percent initial"))
    ff.update_layout(template="plotly_white", height=320, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(ff, use_container_width=True)
    drop = funnel.Orders.iloc[0] - funnel.Orders.iloc[-1]
    st.markdown(f'<div class="pitch">{funnel.Orders.iloc[0]:,} orders entered today; '
                f'{funnel.Orders.iloc[-1]:,} reached the customer. <b>The widest step-down is the '
                f'bottleneck</b> — that is where a day of operational focus buys the most throughput.</div>',
                unsafe_allow_html=True)
with c2:
    st.markdown('<div class="sect">Performance Trend (14 days)</div>', unsafe_allow_html=True)
    tf = go.Figure()
    for col, color, name in [("OnTimePct", GREEN, "On-time %"), ("OTIF", BLUE, "OTIF %"),
                             ("PerfectOrderRate", AMBER, "Perfect order %")]:
        tf.add_trace(go.Scatter(x=trend.Date, y=trend[col], mode="lines", name=name,
                                line=dict(color=color, width=3)))
    tf.update_layout(template="plotly_white", height=320, margin=dict(l=10, r=10, t=10, b=10),
                     legend=dict(orientation="h", y=1.13), yaxis_title="%")
    st.plotly_chart(tf, use_container_width=True)
    st.markdown('<div class="pitch">Three lines, one question: <b>is service trending up or sliding?</b> '
                'A control tower exists to catch the slide on day two, not in the quarterly review.</div>',
                unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# Live exceptions
# ----------------------------------------------------------------------------
st.markdown('<div class="sect">Live Exception Management</div>', unsafe_allow_html=True)
e1, e2 = st.columns([2, 3])
with e1:
    bt = exc_f.groupby(["Type", "Severity"]).size().reset_index(name="n")
    fb = px.bar(bt, x="n", y="Type", color="Severity", orientation="h",
                color_discrete_map={"Critical": RED, "High": AMBER, "Medium": "#c9cfdd"},
                labels={"n": "Open + resolved", "Type": ""})
    fb.update_layout(template="plotly_white", height=300, margin=dict(l=10, r=10, t=10, b=10),
                     legend_title="", legend=dict(orientation="h", y=1.15))
    st.plotly_chart(fb, use_container_width=True)
with e2:
    crit_open = (exc_f[(exc_f.Severity == "Critical") & (exc_f.Resolution == "Open")]
                 .sort_values("AgeHours", ascending=False)
                 [["ExceptionID", "Customer", "Region", "Type", "AgeHours", "ValueUSD"]].head(12))
    st.markdown('<div style="font-size:.85rem;color:#6e7891;margin-bottom:4px;">Open critical exceptions — oldest first</div>',
                unsafe_allow_html=True)
    if crit_open.empty:
        st.success("No open critical exceptions in this region. Network is clean.")
    else:
        st.dataframe(crit_open.style
                     .background_gradient(subset=["AgeHours"], cmap="Reds")
                     .format({"ValueUSD": "${:,.0f}", "AgeHours": "{:d}h"}),
                     use_container_width=True, hide_index=True)

# ----------------------------------------------------------------------------
# Customer scorecard matrix (conditional formatting)
# ----------------------------------------------------------------------------
st.markdown('<div class="sect">Customer Scorecard</div>', unsafe_allow_html=True)
st.markdown('<div style="font-size:.85rem;color:#6e7891;margin-bottom:4px;">'
            'SLA status by contract — green on track, amber at risk, red breach</div>', unsafe_allow_html=True)
sc = cust.sort_values("RevenueUSD", ascending=False)[
    ["Customer", "Tier", "OnTimePct", "SLACompliancePct", "CSAT", "DamageRatePct", "VolumeOrders", "RevenueUSD", "SLAStatus"]]


def hi_higher(v, good, warn):
    return f'background-color:{TINT["green"] if v>=good else (TINT["amber"] if v>=warn else TINT["red"])}'


def hi_lower(v, good, warn):
    return f'background-color:{TINT["green"] if v<=good else (TINT["amber"] if v<=warn else TINT["red"])}'


def hi_status(v):
    return f'background-color:{TINT["green"] if v=="On Track" else (TINT["amber"] if v=="At Risk" else TINT["red"])};font-weight:700'


sty = (sc.style
       .map(lambda v: hi_higher(v, 95, 90), subset=["OnTimePct"])
       .map(lambda v: hi_higher(v, 97, 93), subset=["SLACompliancePct"])
       .map(lambda v: hi_higher(v, 4.5, 4.0), subset=["CSAT"])
       .map(lambda v: hi_lower(v, 1.5, 3.0), subset=["DamageRatePct"])
       .map(hi_status, subset=["SLAStatus"])
       .format({"OnTimePct": "{:.1f}%", "SLACompliancePct": "{:.1f}%", "CSAT": "{:.1f}",
                "DamageRatePct": "{:.1f}%", "RevenueUSD": "${:,.0f}", "VolumeOrders": "{:,}"}))
st.dataframe(sty, use_container_width=True, hide_index=True)
breach = (cust.SLAStatus == "Breach").sum()
st.markdown(f'<div class="pitch">Sorted by revenue, so the accounts that matter most sit at the top. '
            f'<b>{breach} customer(s) are in SLA breach</b> — in a control tower that is not a metric, it '
            f'is a phone call to make today before it becomes a lost contract.</div>', unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# Network performance
# ----------------------------------------------------------------------------
st.markdown('<div class="sect">Network Performance by Hub</div>', unsafe_allow_html=True)
n1, n2 = st.columns([3, 2])
with n1:
    hd = hub_f.sort_values("OTIFPct")
    hd["clr"] = hd.Status.map({"On Track": GREEN, "At Risk": AMBER, "Breach": RED})
    fh = go.Figure(go.Bar(x=hd.OTIFPct, y=hd.Hub, orientation="h",
                          marker_color=hd.clr, text=hd.OTIFPct.map(lambda v: f"{v:.0f}%"),
                          textposition="outside", cliponaxis=False))
    fh.update_layout(template="plotly_white", height=320, margin=dict(l=10, r=30, t=10, b=10),
                     xaxis_title="OTIF %")
    st.plotly_chart(fh, use_container_width=True)
with n2:
    st.markdown('<div style="font-size:.85rem;color:#6e7891;margin-bottom:4px;">Capacity utilisation (watch &gt; 90%)</div>',
                unsafe_allow_html=True)
    ht = hub_f[["Hub", "Region", "CapacityUtilPct", "CostPerOrderUSD", "Status"]].sort_values("CapacityUtilPct", ascending=False)
    st.dataframe(ht.style
                 .map(lambda v: hi_lower(v, 85, 92), subset=["CapacityUtilPct"])
                 .map(hi_status, subset=["Status"])
                 .format({"CapacityUtilPct": "{:.0f}%", "CostPerOrderUSD": "${:.1f}"}),
                 use_container_width=True, hide_index=True)

st.markdown('<div class="foot">Supply Chain Control Tower • Built with Streamlit + Plotly • '
            'synthetic live snapshot • colour = status against target</div>', unsafe_allow_html=True)