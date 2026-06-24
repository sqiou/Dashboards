"""
Fleet & Fuel Analytics — light/teal UI (matches the dashboard suite)
--------------------------------------------------------------------
Streamlit + Plotly. Six tabs mirroring the Power BI report pages, built on
Fleet_Fuel_Dataset.xlsx. Currency USD. Synthetic demo data; figures computed live.

RUN:
    pip install streamlit pandas plotly openpyxl
    streamlit run fleet_fuel_dashboard.py

NOTE: the spec's "best visual" is live Azure Maps GPS tracking. That needs real
per-vehicle coordinates + a tile provider, which synthetic data cannot honestly
supply, so geography is shown at depot/region level instead.
"""

from pathlib import Path
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Fleet & Fuel Analytics", layout="wide", page_icon="⛽")

TEAL, DARK, GREY = "#1a9e8f", "#1f2a44", "#5e667d"
GREEN, AMBER, RED = "#00a39c", "#e0a13a", "#e5566b"
VIVID = px.colors.qualitative.Bold + px.colors.qualitative.Vivid
PRICE = {"Diesel": 3.90, "Gasoline": 3.45}   # USD/gallon (assumption, mirrors generator)
TODAY = pd.Timestamp("2026-06-22")

st.markdown(f"""
<style>
  .block-container {{ padding-top: 3rem; max-width: 1500px; }}
  .hero {{ border-left:6px solid {TEAL}; border-top:3px solid {TEAL};
           background:linear-gradient(90deg,#f1faf8,#fff 60%); padding:18px 24px; border-radius:6px; margin-bottom:10px; }}
  .hero h1 {{ color:{DARK}; font-size:2.0rem; margin:0; font-weight:800; }}
  .hero p  {{ color:{TEAL}; font-weight:600; margin:3px 0 0 0; font-size:.85rem; }}
  .sect {{ color:{DARK}; font-weight:800; font-size:1.3rem; margin:10px 0 2px 0; }}
  .ssub {{ color:{GREY}; font-size:.85rem; margin-bottom:6px; }}
  .kpi-l {{ color:{GREY}; font-size:.8rem; margin-bottom:2px; }}
  .kpi-v {{ color:{DARK}; font-size:1.8rem; font-weight:800; line-height:1.05; }}
  .kpi-s {{ color:{GREY}; font-size:.76rem; }}
  .insight {{ background:#f1faf8; border:1px solid #cfeae4; border-left:4px solid {TEAL};
              padding:12px 16px; border-radius:6px; color:{DARK}; font-size:.9rem; line-height:1.5; margin:8px 0 12px 0; }}
  .insight b {{ color:{TEAL}; }}
  .tipbox {{ background:#e8f6fb; border-left:4px solid #3aa6c4; padding:10px 14px; border-radius:4px;
             color:{DARK}; font-size:.86rem; margin:6px 0 12px 0; }}
  hr {{ margin:8px 0 12px 0; }}
  .foot {{ text-align:center; color:#8c98a8; font-size:.8rem; padding:18px 0; }}
  .stTabs [data-baseweb="tab-list"] {{ gap:24px; }}
  .stTabs [data-baseweb="tab"] {{ font-weight:600; padding:6px 4px; }}
</style>
""", unsafe_allow_html=True)

DATA = Path(__file__).parent / "Fleet_Fuel_Dataset.xlsx"
if not DATA.exists():
    st.error(f"Cannot find {DATA.name}. Place it beside this script and rerun.")
    st.stop()


@st.cache_data(show_spinner=False)
def load():
    xls = pd.ExcelFile(DATA)
    veh = pd.read_excel(xls, "Dim_Vehicle", parse_dates=["LastServiceDate", "NextServiceDue"])
    drv = pd.read_excel(xls, "Dim_Driver")
    route = pd.read_excel(xls, "Dim_Route")
    ddate = pd.read_excel(xls, "Dim_Date", parse_dates=["Date"])
    trip = pd.read_excel(xls, "Fact_Trips")
    REQUIRED = {
        "Dim_Vehicle": {"VehicleID","VehicleType","Make","FuelType","AgeBand","Depot","Region","RatedMPG","AnnualInsuranceUSD","Status","NextServiceDue","LastServiceDate"},
        "Dim_Driver": {"DriverID","DriverName"},
        "Dim_Route": {"RouteID","RouteName","PlannedDistanceMiles"},
        "Dim_Date": {"DateKey","Date","MonthYear","MonthSort"},
        "Fact_Trips": {"VehicleID","DriverID","RouteID","DateKey","DistanceMiles","PlannedDistanceMiles","FuelGallons","FuelCostUSD","MaintenanceCostUSD","CO2Lbs","IdleHours","DriveHours","SpeedingEvents","HarshEvents","OnScheduleFlag"},
    }
    _frames = {"Dim_Vehicle": veh, "Dim_Driver": drv, "Dim_Route": route, "Dim_Date": ddate, "Fact_Trips": trip}
    _missing = {s: sorted(REQUIRED[s] - set(f.columns)) for s, f in _frames.items() if REQUIRED[s] - set(f.columns)}
    if _missing:
        st.error("Dataset schema mismatch — workbook does not match this dashboard's expected schema:\n" +
                 "\n".join(f"  - {s}: missing {cols}" for s, cols in _missing.items()))
        st.stop()
    df = (trip
          .merge(veh[["VehicleID", "VehicleType", "Make", "FuelType", "AgeBand", "Depot", "Region",
                      "RatedMPG", "AnnualInsuranceUSD"]], on="VehicleID")
          .merge(drv[["DriverID", "DriverName"]], on="DriverID")
          .merge(route[["RouteID", "RouteName"]], on="RouteID")
          .merge(ddate[["DateKey", "Date", "MonthYear", "MonthSort"]], on="DateKey"))
    df["BudgetGallons"] = df.DistanceMiles / df.RatedMPG
    df["BudgetFuelUSD"] = df.BudgetGallons * df.FuelType.map(PRICE)
    df["TotalCostUSD"] = df.FuelCostUSD + df.MaintenanceCostUSD
    return veh, drv, route, ddate, df


veh, drv, route, ddate, df = load()

st.markdown('<div class="hero"><h1>Fleet &amp; Fuel Analytics</h1>'
            '<p>Every vehicle, every gallon, every driver — cost and waste in one place (USD)</p></div>',
            unsafe_allow_html=True)

# ---- dynamic headline hook (the one number that matters, computed live) ----
_waste = df.FuelCostUSD.sum() - df.BudgetFuelUSD.sum()
_wd = df.groupby("Depot").agg(c=("TotalCostUSD", "sum"), k=("DistanceMiles", "sum"))
_worst_depot = (_wd.c / _wd.k).idxmax()
_idle = df.IdleHours.sum() / (df.IdleHours.sum() + df.DriveHours.sum()) * 100
st.markdown(
    f'<div class="insight">⛽ <b>The headline:</b> the fleet burned <b>${_waste/1e6:,.1f}M</b> in fuel '
    f'above what it would cost if every trip hit its rated MPG — that gap is idling, harsh driving and '
    f'ageing engines, and it is recoverable. {_idle:.0f}% of engine hours are spent idling, and '
    f'<b>{_worst_depot}</b> runs the most expensive mile. The tabs below show where each dollar leaks.</div>',
    unsafe_allow_html=True)

# ---- global filters ----
c1, c2, _ = st.columns([1, 1, 2])
reg = c1.selectbox("Region", ["All"] + sorted(veh.Region.unique()))
vt = c2.selectbox("Vehicle Type", ["All"] + sorted(veh.VehicleType.unique()))
veh_f = veh.copy(); df_f = df.copy()
if reg != "All":
    veh_f = veh_f[veh_f.Region == reg]; df_f = df_f[df_f.Region == reg]
if vt != "All":
    veh_f = veh_f[veh_f.VehicleType == vt]; df_f = df_f[df_f.VehicleType == vt]
if df_f.empty:
    st.warning("No trips match those filters."); st.stop()


def kpi(col, label, value, sub=""):
    col.markdown(f'<div class="kpi-l">{label}</div><div class="kpi-v">{value}</div>'
                 f'<div class="kpi-s">{sub}</div>', unsafe_allow_html=True)


def sect(t, s=""):
    st.markdown(f'<div class="sect">{t}</div>' + (f'<div class="ssub">{s}</div>' if s else ''),
                unsafe_allow_html=True)


def insight(html):
    st.markdown(f'<div class="insight">{html}</div>', unsafe_allow_html=True)


def score_tint(v):
    return f"background-color:{'#e6f6ee' if v >= 66 else ('#fdf3e2' if v >= 33 else '#fdecee')}"


tabs = st.tabs(["Overview", "Fuel Analysis", "Driver Behaviour", "Vehicle Health",
                "Route Efficiency", "Cost Breakdown"])

# ============================================================ OVERVIEW
with tabs[0]:
    insight("Start here every morning. <b>Two numbers decide whether the fleet is healthy: cost per mile "
            "and idle %.</b> Cost/mi is your true running rate — the figure to quote clients and defend in "
            "budget reviews — and idle % is pure waste, fuel burned going nowhere. Everything else on this "
            "page is context for those two.")
    mi, gal = df_f.DistanceMiles.sum(), df_f.FuelGallons.sum()
    cost = df_f.TotalCostUSD.sum()
    mtd = df_f[df_f.MonthSort == df_f.MonthSort.max()]
    idle = df_f.IdleHours.sum() / (df_f.IdleHours.sum() + df_f.DriveHours.sum()) * 100
    overdue = int((veh_f.NextServiceDue < TODAY).sum())
    sect("Fleet KPIs"); st.markdown("<hr>", unsafe_allow_html=True)
    r1 = st.columns(4)
    kpi(r1[0], "Total Fleet Size", f"{len(veh_f)}", "vehicles")
    kpi(r1[1], "Avg Cost / mi", f"${cost/mi:,.2f}", "fuel + maintenance")
    kpi(r1[2], "Fuel Spend MTD", f"${mtd.FuelCostUSD.sum()/1e6:,.2f}M", df_f.MonthYear.iloc[-1] if len(df_f) else "")
    kpi(r1[3], "Idle Hours %", f"{idle:.1f}%", "idle ÷ (idle+drive)")
    st.write("")
    r2 = st.columns(4)
    kpi(r2[0], "Active Vehicles", f"{(veh_f.Status=='Active').sum()} / {len(veh_f)}", "status = active")
    kpi(r2[1], "Avg Fuel Eff.", f"{mi/gal:,.1f} MPG", "weighted (mi÷gal)")
    kpi(r2[2], "Overdue Services", f"{overdue}", "past next-service date")
    kpi(r2[3], "CO₂ This Month", f"{mtd.CO2Lbs.sum()/2204.62:,.1f} t", "standard emission factors")
    st.markdown("<hr>", unsafe_allow_html=True)

    g1, g2 = st.columns(2)
    with g1:
        sect("Fuel spend trend", "Daily fuel cost across the fleet")
        tr = df_f.groupby("Date")["FuelCostUSD"].sum().reset_index()
        ft = px.line(tr, x="Date", y="FuelCostUSD", labels={"FuelCostUSD": "USD", "Date": ""})
        ft.update_traces(line_color=TEAL, line_width=2)
        ft.update_layout(template="plotly_white", height=300, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(ft, width='stretch')
    with g2:
        sect("Idle % vs fuel spend", "Each dot is a vehicle — top-right are the worst offenders")
        vv = df_f.groupby("VehicleID").agg(idle=("IdleHours", "sum"), drive=("DriveHours", "sum"),
                                           fuel=("FuelCostUSD", "sum"), vtype=("VehicleType", "first")).reset_index()
        vv["IdlePct"] = vv.idle / (vv.idle + vv.drive) * 100
        sc = px.scatter(vv, x="IdlePct", y="fuel", color="vtype", hover_name="VehicleID",
                        color_discrete_sequence=VIVID, labels={"IdlePct": "Idle %", "fuel": "Fuel spend (USD)"})
        sc.update_layout(template="plotly_white", height=300, margin=dict(l=10, r=10, t=10, b=10), legend_title="")
        st.plotly_chart(sc, width='stretch')

    sect("Cost per mile by driver", "Worst 15 — the running-cost leaderboard")
    _d = df_f.groupby("DriverName").agg(c=("TotalCostUSD", "sum"), k=("DistanceMiles", "sum"))
    dby = (_d.c / _d.k).sort_values(ascending=False).head(15).reset_index(name="CostKm")
    fb = px.bar(dby.sort_values("CostKm"), x="CostKm", y="DriverName", orientation="h",
                color="CostKm", color_continuous_scale=["#7cc4a3", RED], labels={"CostKm": "$ / mi", "DriverName": ""})
    fb.update_layout(template="plotly_white", height=420, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
    st.plotly_chart(fb, width='stretch')
    st.markdown('<div class="tipbox"><b>Tip:</b> Live GPS map omitted — synthetic data has no real '
                'coordinates. With a telematics feed (Samsara, etc.) this is where the live vehicle map goes.</div>',
                unsafe_allow_html=True)

# ============================================================ FUEL ANALYSIS
with tabs[1]:
    insight("Fuel is usually 30–50% of fleet running cost, so a 1% efficiency gain compounds fast. "
            "<b>The two levers are driving behaviour and vehicle choice</b> — this view separates them.")
    sect("Fuel efficiency trend", "Fleet MPG over time (weighted)")
    _e = df_f.groupby("Date").agg(km=("DistanceMiles", "sum"), l=("FuelGallons", "sum"))
    eff = (_e.km / _e.l).reset_index(name="mpg")
    fe = px.line(eff, x="Date", y="mpg", labels={"mpg": "MPG", "Date": ""})
    fe.update_traces(line_color=TEAL, line_width=2)
    fe.update_layout(template="plotly_white", height=280, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fe, width='stretch')
    g1, g2 = st.columns(2)
    with g1:
        sect("Gallons by vehicle type")
        lt = df_f.groupby("VehicleType")["FuelGallons"].sum().sort_values().reset_index()
        fl = px.bar(lt, x="FuelGallons", y="VehicleType", orientation="h", color_discrete_sequence=[TEAL],
                    labels={"FuelGallons": "Gallons", "VehicleType": ""})
        fl.update_layout(template="plotly_white", height=300, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fl, width='stretch')
    with g2:
        sect("Fuel spend by fuel type")
        fs = df_f.groupby("FuelType")["FuelCostUSD"].sum().reset_index()
        fp = px.pie(fs, names="FuelType", values="FuelCostUSD", hole=.5,
                    color_discrete_sequence=[TEAL, AMBER, "#8aa1b8"])
        fp.update_layout(template="plotly_white", height=300, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fp, width='stretch')
    sect("Actual vs efficiency-budget fuel", "Budget = fuel if every trip hit its rated MPG")
    mb = df_f.groupby("MonthYear").agg(Actual=("FuelCostUSD", "sum"), Budget=("BudgetFuelUSD", "sum"),
                                       s=("MonthSort", "first")).sort_values("s").reset_index()
    mbar = go.Figure()
    mbar.add_trace(go.Bar(x=mb.MonthYear, y=mb.Budget, name="Efficiency budget", marker_color="#bfe3dc"))
    mbar.add_trace(go.Bar(x=mb.MonthYear, y=mb.Actual, name="Actual", marker_color=TEAL))
    mbar.update_layout(template="plotly_white", height=300, margin=dict(l=10, r=10, t=10, b=10),
                       barmode="group", yaxis_title="USD", legend=dict(orientation="h", y=1.15))
    st.plotly_chart(mbar, width='stretch')
    waste = mb.Actual.sum() - mb.Budget.sum()
    insight(f"Actual fuel ran <b>${waste/1e6:,.2f}M</b> above the efficiency budget over the period — "
            f"that gap is the cost of driving below rated MPG (idling, harsh driving, ageing engines). "
            f"It is the single most attackable number on this page. <i>Budget here is a derived efficiency "
            f"baseline, not a finance budget.</i>")

# ============================================================ DRIVER BEHAVIOUR
with tabs[2]:
    insight("Behaviour-driven cost is invisible on a fuel invoice but real: speeding and harsh events "
            "burn fuel and wear brakes/tyres. <b>Scoring drivers objectively turns a vague complaint into "
            "a coaching list.</b>")
    db = df_f.groupby("DriverName").agg(km=("DistanceMiles", "sum"), speeding=("SpeedingEvents", "sum"),
                                        harsh=("HarshEvents", "sum"), idle=("IdleHours", "sum"),
                                        drive=("DriveHours", "sum"), sched=("OnScheduleFlag", "mean")).reset_index()
    db["EventsPer1000mi"] = (db.speeding + db.harsh) / db.km * 1000
    db["IdlePct"] = db.idle / (db.idle + db.drive) * 100
    risk = db.EventsPer1000mi + 0.3 * db.IdlePct
    rng_ = (risk.max() - risk.min()) or 1
    db["SafetyScore"] = (100 * (1 - (risk - risk.min()) / rng_)).round(0)
    sect("Driver summary"); st.markdown("<hr>", unsafe_allow_html=True)
    k = st.columns(3)
    kpi(k[0], "Drivers Tracked", f"{len(db)}")
    kpi(k[1], "Avg Safety Score", f"{db.SafetyScore.mean():.0f} / 100")
    kpi(k[2], "On-Schedule Trips", f"{df_f.OnScheduleFlag.mean()*100:.1f}%")
    st.markdown("<hr>", unsafe_allow_html=True)
    g1, g2 = st.columns(2)
    with g1:
        sect("Speeding vs harsh-braking", "Bubble = distance driven")
        bs = px.scatter(db, x="speeding", y="harsh", size="km", color="SafetyScore", hover_name="DriverName",
                        color_continuous_scale=[RED, AMBER, GREEN], size_max=34,
                        labels={"speeding": "Speeding events", "harsh": "Harsh events"})
        bs.update_layout(template="plotly_white", height=320, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(bs, width='stretch')
    with g2:
        sect("Safety leaderboard", "Lowest scores first — coach these")
        lb = db.sort_values("SafetyScore")[["DriverName", "SafetyScore", "EventsPer1000mi", "IdlePct"]].head(12)
        st.dataframe(lb.style.map(score_tint, subset=["SafetyScore"])
                     .format({"EventsPer1000mi": "{:.1f}", "IdlePct": "{:.1f}%", "SafetyScore": "{:.0f}"}),
                     width='stretch', hide_index=True)
    st.markdown('<div class="tipbox"><b>Tip:</b> SafetyScore is a <i>relative</i> 0–100 ranking — it '
                'min-max normalises each driver\'s risk (events per 1000 mi + 0.3×idle%) across the current '
                'selection, so 100 = safest here, 0 = riskiest. It is comparative, not an absolute rating; '
                're-weight the risk terms to your own policy.</div>', unsafe_allow_html=True)

# ============================================================ VEHICLE HEALTH
with tabs[3]:
    insight("Maintenance is cheaper than breakdown. <b>The job here is to see what is due before it fails</b> "
            "and which makes cost the most to keep running.")
    vh = veh_f.copy()
    vh["DaysToService"] = (vh.NextServiceDue - TODAY).dt.days
    vh["ServiceStatus"] = np.where(vh.DaysToService < 0, "Overdue",
                                   np.where(vh.DaysToService <= 30, "Due ≤30d", "OK"))
    sect("Service status"); st.markdown("<hr>", unsafe_allow_html=True)
    k = st.columns(3)
    kpi(k[0], "Overdue", f"{(vh.ServiceStatus=='Overdue').sum()}", "act now")
    kpi(k[1], "Due in 30 days", f"{(vh.ServiceStatus=='Due ≤30d').sum()}", "schedule")
    kpi(k[2], "Healthy", f"{(vh.ServiceStatus=='OK').sum()}", "no action")
    st.markdown("<hr>", unsafe_allow_html=True)
    g1, g2 = st.columns(2)
    with g1:
        sect("Service status mix")
        ss = vh.ServiceStatus.value_counts().reindex(["Overdue", "Due ≤30d", "OK"]).fillna(0).reset_index()
        ss.columns = ["Status", "n"]
        fp = px.pie(ss, names="Status", values="n", hole=.55,
                    color="Status", color_discrete_map={"Overdue": RED, "Due ≤30d": AMBER, "OK": GREEN})
        fp.update_layout(template="plotly_white", height=300, margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fp, width='stretch')
    with g2:
        sect("Maintenance cost by make", "Where repair money goes")
        mc = df_f.groupby("Make")["MaintenanceCostUSD"].sum().sort_values().reset_index()
        fm = px.bar(mc, x="MaintenanceCostUSD", y="Make", orientation="h", color_discrete_sequence=[TEAL],
                    labels={"MaintenanceCostUSD": "USD", "Make": ""})
        fm.update_layout(template="plotly_white", height=300, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fm, width='stretch')
    sect("Vehicles due within 30 days")
    due = vh[vh.DaysToService.between(-999, 30)].sort_values("DaysToService")[
        ["VehicleID", "VehicleType", "Make", "Depot", "NextServiceDue", "DaysToService", "ServiceStatus"]]
    if due.empty:
        st.info("No services due in the next 30 days for this selection.")
    else:
        st.dataframe(due.style.apply(
            lambda r: [f"background-color:{'#fdecee' if r.ServiceStatus=='Overdue' else '#fdf3e2'}"] * len(r), axis=1)
            .format({"DaysToService": "{:d}"}), width='stretch', hide_index=True)
    st.markdown('<div class="tipbox"><b>Note:</b> The dataset has next-service dates and ad-hoc repair '
                'costs, but no scheduled-vs-completed service history, so that comparison is not shown '
                '(it would need a real maintenance log).</div>', unsafe_allow_html=True)

# ============================================================ ROUTE EFFICIENCY
with tabs[4]:
    insight("Routes are where cost hides in plain sight. <b>The same delivery on a poorly chosen lane can "
            "cost double</b> — ranking cost per mile exposes it.")
    rt = df_f.groupby("RouteName").agg(cost=("TotalCostUSD", "sum"), km=("DistanceMiles", "sum"),
                                       planned=("PlannedDistanceMiles", "sum"), trips=("RouteID", "size")).reset_index()
    rt["CostPerKm"] = rt.cost / rt.km
    rt["Utilisation"] = (rt.km / rt.planned).clip(upper=3)
    g1, g2 = st.columns(2)
    with g1:
        sect("Top 10 routes by cost per mile")
        top = rt.sort_values("CostPerKm", ascending=False).head(10)
        fr = px.bar(top.sort_values("CostPerKm"), x="CostPerKm", y="RouteName", orientation="h",
                    color="CostPerKm", color_continuous_scale=["#7cc4a3", RED],
                    labels={"CostPerKm": "$ / mi", "RouteName": ""})
        fr.update_layout(template="plotly_white", height=360, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
        st.plotly_chart(fr, width='stretch')
    with g2:
        sect("Planned vs actual distance", "Above the line = drivers exceeding planned distance")
        sc = px.scatter(rt, x="planned", y="km", size="trips", hover_name="RouteName",
                        color_discrete_sequence=[TEAL], labels={"planned": "Planned miles (sum)", "km": "Actual miles (sum)"})
        mx = max(rt.planned.max(), rt.km.max())
        sc.add_trace(go.Scatter(x=[0, mx], y=[0, mx], mode="lines", line=dict(dash="dash", color=GREY), showlegend=False))
        sc.update_layout(template="plotly_white", height=360, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(sc, width='stretch')
    sect("Cost per mile by depot", "Geography at depot level (live GPS map needs a telematics feed)")
    _p = df_f.groupby("Depot").agg(c=("TotalCostUSD", "sum"), k=("DistanceMiles", "sum"))
    dp = (_p.c / _p.k).sort_values().reset_index(name="CostKm")
    fd = px.bar(dp, x="CostKm", y="Depot", orientation="h", color_discrete_sequence=[TEAL],
                labels={"CostKm": "$ / mi", "Depot": ""})
    fd.update_layout(template="plotly_white", height=300, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fd, width='stretch')

# ============================================================ COST BREAKDOWN
with tabs[5]:
    insight("Total cost of ownership is fuel + maintenance + insurance. <b>Most teams only watch fuel</b> "
            "— this view restores the other two so decisions (keep vs replace, age policy) are made on full cost.")
    period_days = (df_f.Date.max() - df_f.Date.min()).days + 1
    fuel_t = df_f.FuelCostUSD.sum()
    maint_t = df_f.MaintenanceCostUSD.sum()
    ins_t = veh_f.AnnualInsuranceUSD.sum() * period_days / 365.0   # prorated
    total_t = fuel_t + maint_t + ins_t
    sect("Cost composition"); st.markdown("<hr>", unsafe_allow_html=True)
    k = st.columns(4)
    kpi(k[0], "Fuel", f"${fuel_t/1e6:,.1f}M", f"{fuel_t/total_t*100:.0f}%")
    kpi(k[1], "Maintenance", f"${maint_t/1e6:,.1f}M", f"{maint_t/total_t*100:.0f}%")
    kpi(k[2], "Insurance (prorated)", f"${ins_t/1e6:,.1f}M", f"{ins_t/total_t*100:.0f}%")
    kpi(k[3], "Total Cost", f"${total_t/1e6:,.1f}M", f"{period_days} days")
    st.markdown("<hr>", unsafe_allow_html=True)
    g1, g2 = st.columns(2)
    with g1:
        sect("Cost build-up", "Fuel → +maintenance → +insurance → total")
        wf = go.Figure(go.Waterfall(orientation="v", measure=["relative", "relative", "relative", "total"],
                                    x=["Fuel", "Maintenance", "Insurance", "Total"],
                                    y=[fuel_t, maint_t, ins_t, 0],
                                    text=[f"{v/1e6:,.1f}M" for v in [fuel_t, maint_t, ins_t, total_t]],
                                    textposition="outside", connector={"line": {"color": "#c9d2e0"}},
                                    increasing={"marker": {"color": TEAL}}, totals={"marker": {"color": DARK}}))
        wf.update_layout(template="plotly_white", height=320, margin=dict(l=10, r=10, t=20, b=10), yaxis_title="USD")
        st.plotly_chart(wf, width='stretch')
    with g2:
        sect("Total cost by vehicle age band", "Older bands usually cost more to run")
        ab = df_f.groupby("AgeBand")["TotalCostUSD"].sum().reindex(["0-3", "4-7", "8-11", "12+"]).fillna(0).reset_index()
        fa = px.bar(ab, x="AgeBand", y="TotalCostUSD", color="AgeBand",
                    color_discrete_sequence=[GREEN, "#7cc4a3", AMBER, RED], labels={"TotalCostUSD": "USD", "AgeBand": ""})
        fa.update_layout(template="plotly_white", height=320, margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
        st.plotly_chart(fa, width='stretch')
    insight("Insurance is prorated from annual premiums across the period; fuel and maintenance are actual. "
            "<b>If the 8+ yr band's running cost approaches a newer band's lease/finance cost, that is your "
            "replace signal</b> — the analysis that justifies capex.")

st.markdown('<div class="foot">Fleet &amp; Fuel Analytics • Built with Streamlit + Plotly • '
            'synthetic dataset, USD • prices &amp; insurance are labelled assumptions</div>', unsafe_allow_html=True)