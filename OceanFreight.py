"""
Ocean Freight & Port Operations — light/teal UI (matches the dashboard suite)
-----------------------------------------------------------------------------
Streamlit + Plotly. Six tabs mirroring the report pages, on a 20k-voyage dataset.
Includes a country choropleth (with metric selector) AND a port-level bubble map.

RUN:
    pip install streamlit pandas plotly openpyxl
    streamlit run ocean_freight_dashboard.py

Workbook must sit beside this script. No upload prompt. Synthetic demo data; figures
computed live. Port coords/ISO-3 are real (approximate); transit/demurrage/freight are
labelled assumptions.

VISUAL NOTE: a choropleth shades whole COUNTRIES, so it loses port-level detail — a
busy single-port country looks like a multi-port one. The port BUBBLE map beside it is
the faithful terminal-level view. Use the choropleth for "which markets", the bubbles
for "which ports".
"""

from pathlib import Path
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Ocean Freight & Port Operations", layout="wide", page_icon="🚢")

TEAL, DARK, GREY = "#1a9e8f", "#1f2a44", "#5e667d"
GREEN, AMBER, RED = "#00a39c", "#e0a13a", "#e5566b"
VIVID = px.colors.qualitative.Bold + px.colors.qualitative.Vivid
SEQ = ["#dff3ef", "#7cc4b8", "#1a9e8f", "#0f6f63"]

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
  .kpi-v {{ color:{DARK}; font-size:1.7rem; font-weight:800; line-height:1.05; }}
  .kpi-s {{ color:{GREY}; font-size:.74rem; }}
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

DATA = Path(__file__).parent / "Ocean_Freight_Dataset.xlsx"
if not DATA.exists():
    st.error(f"Cannot find {DATA.name}. Place it beside this script and rerun.")
    st.stop()


@st.cache_data(show_spinner=False)
def load():
    xls = pd.ExcelFile(DATA)
    port = pd.read_excel(xls, "Dim_Port")
    carrier = pd.read_excel(xls, "Dim_Carrier")
    ddate = pd.read_excel(xls, "Dim_Date", parse_dates=["Date"])
    fact = pd.read_excel(xls, "Fact_Shipment")

    # ---- numeric guard: numbers can arrive as str (Excel round-trip / pandas 3.0 str dtype) ----
    def coerce(df, cols, sheet):
        for c in cols:
            if c not in df.columns:
                continue
            if not pd.api.types.is_numeric_dtype(df[c]):
                parsed = pd.to_numeric(df[c], errors="coerce")
                if parsed.notna().mean() >= 0.5:                 # looks numeric -> fix it, and report
                    bad = df[c][parsed.isna() & df[c].notna()].unique()[:5]
                    print(f"[guard] {sheet}.{c}: stored as {df[c].dtype}, coerced to numeric. "
                          f"Bad cells: {list(bad) if len(bad) else 'none'}")
                    df[c] = parsed
        return df

    fact = coerce(fact, ["TEUs", "PlannedTransitDays", "ActualTransitDays", "PortDwellDays",
                         "FreeDaysAllowed", "DemurrageUSD", "FreightUSD", "OnTimeFlag",
                         "DocErrorFlag", "CustomsHoldFlag", "DamageFlag", "DepartDateKey",
                         "ArriveDateKey"], "Fact_Shipment")
    port = coerce(port, ["Lat", "Lon"], "Dim_Port")
    carrier = coerce(carrier, ["RelBase"], "Dim_Carrier")
    ddate = coerce(ddate, ["DateKey", "MonthSort"], "Dim_Date")

    # drop rows whose required numerics failed to parse (avoid silently charting partial garbage)
    before = len(fact)
    fact = fact.dropna(subset=["TEUs", "FreightUSD", "DemurrageUSD", "ArriveDateKey"])
    if before - len(fact):
        print(f"[guard] Fact_Shipment: dropped {before - len(fact)} rows with non-numeric required values")

    o = port.add_prefix("O_"); d = port.add_prefix("D_")
    df = (fact
          .merge(o, left_on="OriginPort", right_on="O_PortCode")
          .merge(d, left_on="DestPort", right_on="D_PortCode")
          .merge(carrier[["CarrierName", "Alliance", "Tier"]], on="CarrierName")
          .merge(ddate[["DateKey", "Date", "MonthYear", "MonthSort"]], left_on="ArriveDateKey", right_on="DateKey"))
    df["Lane"] = df.O_PortName + " → " + df.D_PortName
    df["Delay"] = df.ActualTransitDays - df.PlannedTransitDays
    return port, carrier, ddate, df


port, carrier, ddate, df = load()

st.markdown('<div class="hero"><h1>Ocean Freight &amp; Port Operations</h1>'
            '<p>Shipment volumes, carrier reliability, port dwell &amp; lane performance — 20k voyages</p></div>',
            unsafe_allow_html=True)

# ---- dynamic headline hook ----
_arr = df[df.Status != "In Transit"]
_worst_lane = (df.groupby("Lane").agg(c=("DemurrageUSD", "sum")).c.idxmax())
_dem_total = df.DemurrageUSD.sum()
st.markdown(
    f'<div class="insight">🚢 <b>The headline:</b> {_arr.OnTimeFlag.mean()*100:.0f}% of arrivals hit their '
    f'window, the rest sit in port — and demurrage (the fee for overstaying a berth) has cost '
    f'<b>${_dem_total/1e6:,.1f}M</b> across these voyages, concentrated on lanes like {_worst_lane}. '
    f'Demurrage is the most avoidable cost in ocean freight: it is pure penalty for time, not distance.</div>',
    unsafe_allow_html=True)

# ---- filters ----
c1, c2, _ = st.columns([1, 1, 2])
reg = c1.selectbox("Destination region", ["All"] + sorted(port.Region.unique()))
car = c2.selectbox("Carrier", ["All"] + sorted(carrier.CarrierName.unique()))
df_f = df.copy()
if reg != "All":
    df_f = df_f[df_f.D_Region == reg]
if car != "All":
    df_f = df_f[df_f.CarrierName == car]
if df_f.empty:
    st.warning("No voyages match those filters."); st.stop()


def kpi(col, label, value, sub=""):
    col.markdown(f'<div class="kpi-l">{label}</div><div class="kpi-v">{value}</div>'
                 f'<div class="kpi-s">{sub}</div>', unsafe_allow_html=True)


def sect(t, s=""):
    st.markdown(f'<div class="sect">{t}</div>' + (f'<div class="ssub">{s}</div>' if s else ''), unsafe_allow_html=True)


def insight(html):
    st.markdown(f'<div class="insight">{html}</div>', unsafe_allow_html=True)


tabs = st.tabs(["Shipment Overview", "Lane Performance", "Carrier Scorecard",
                "Port Dwell & Delays", "Container Tracking", "Cost Analysis"])

# ============================================================ SHIPMENT OVERVIEW
with tabs[0]:
    arr = df_f[df_f.Status != "In Transit"]
    latest = df_f.MonthSort.max()
    mtd = df_f[df_f.MonthSort == latest]
    sect("Port & shipment KPIs"); st.markdown("<hr>", unsafe_allow_html=True)
    r1 = st.columns(4)
    kpi(r1[0], "Active Shipments", f"{int(df_f.Status.isin(['In Transit','At Port']).sum()):,}", "in transit + at port")
    kpi(r1[1], "On-Time Arrival", f"{arr.OnTimeFlag.mean()*100:.1f}%", "delay ≤ 1 day")
    kpi(r1[2], "Avg Port Dwell", f"{arr.PortDwellDays.mean():.1f} days", "berth to gate-out")
    kpi(r1[3], "TEUs This Month", f"{int(mtd.TEUs.sum()):,}", df_f.MonthYear.iloc[-1] if len(df_f) else "")
    st.write("")
    r2 = st.columns(4)
    kpi(r2[0], "Delayed Shipments", f"{int((df_f.Delay>2).sum()):,}", "> 2 days late")
    kpi(r2[1], "Avg Transit Time", f"{arr.ActualTransitDays.mean():.1f} days", "actual port-to-port")
    kpi(r2[2], "Demurrage (MTD)", f"${mtd.DemurrageUSD.sum()/1e3:,.0f}k", "berth-overstay fees")
    kpi(r2[3], "Carrier Reliability", f"{df_f.groupby('CarrierName').OnTimeFlag.mean().mean()*100:.1f}%", "avg across carriers")
    st.markdown("<hr>", unsafe_allow_html=True)

    sect("Where the volume lands", "One clear map at a time — switch between terminal and market view")
    mc1, mc2 = st.columns([1.4, 1])
    view = mc1.radio("Map view", ["Port bubbles (terminals)", "Country choropleth (markets)"], horizontal=True)
    mopt = mc2.selectbox("Metric", ["TEUs shipped", "Demurrage cost", "On-time arrival %"])
    mcol, mfn = {"TEUs shipped": ("TEUs", "sum"), "Demurrage cost": ("DemurrageUSD", "sum"),
                 "On-time arrival %": ("OnTimeFlag", "mean")}[mopt]
    cscale = [RED, AMBER, GREEN] if mopt == "On-time arrival %" else SEQ
    if view.startswith("Port"):
        pdf = df_f.groupby(["DestPort", "D_PortName", "D_Country", "D_Lat", "D_Lon"]).agg(
            val=(mcol, mfn), TEUs=("TEUs", "sum")).reset_index()
        if mopt == "On-time arrival %":
            pdf["val"] = pdf["val"] * 100
        m = px.scatter_geo(pdf, lat="D_Lat", lon="D_Lon", size="TEUs", color="val",
                           hover_name="D_PortName", color_continuous_scale=cscale, size_max=44,
                           labels={"val": mopt})
    else:
        cdf = df_f.groupby(["D_ISO3", "D_Country"]).agg(val=(mcol, mfn)).reset_index()
        if mopt == "On-time arrival %":
            cdf["val"] = cdf["val"] * 100
        m = px.choropleth(cdf, locations="D_ISO3", color="val", hover_name="D_Country",
                          color_continuous_scale=cscale, labels={"val": mopt})
    m.update_geos(showland=True, landcolor="#eef2f4", showocean=True, oceancolor="#d9edef",
                  showcountries=True, countrycolor="#ffffff", coastlinecolor="#cdd6da",
                  projection_type="natural earth", showframe=False, bgcolor="rgba(0,0,0,0)",
                  lataxis_range=[-55, 78])
    m.update_layout(template="plotly_white", height=540, margin=dict(l=0, r=0, t=10, b=0),
                    coloraxis_colorbar=dict(title=mopt.split()[0]))
    st.plotly_chart(m, width='stretch')
    insight("Shown one map at a time for clarity. <b>Port bubbles</b> are the real terminals (size = TEUs, "
            "colour = the chosen metric); <b>country choropleth</b> rolls the same metric up to the import "
            "market. The bubbles deliberately expose a single struggling terminal that a country shade would "
            "average away — that is why a choropleth alone is the wrong tool for port operations.")

    g3, g4 = st.columns(2)
    with g3:
        sect("Monthly TEU volume")
        vt = df_f.groupby(["MonthSort", "MonthYear"]).agg(TEUs=("TEUs", "sum")).reset_index().sort_values("MonthSort")
        fv = px.bar(vt, x="MonthYear", y="TEUs", color_discrete_sequence=[TEAL])
        fv.update_layout(template="plotly_white", height=300, margin=dict(l=10, r=10, t=10, b=10), xaxis_title="")
        st.plotly_chart(fv, width='stretch')
    with g4:
        sect("Shipment status")
        sc = df_f.Status.value_counts().reset_index()
        sc.columns = ["Status", "n"]
        fp = px.pie(sc, names="Status", values="n", hole=.55,
                    color="Status", color_discrete_map={"In Transit": TEAL, "At Port": AMBER, "Cleared": "#8aa1b8"})
        fp.update_layout(template="plotly_white", height=300, margin=dict(l=0, r=0, t=10, b=0))
        st.plotly_chart(fp, width='stretch')

# ============================================================ LANE PERFORMANCE
with tabs[1]:
    insight("A lane is an origin–destination pair, and lanes are not equal: <b>the same TEU can cost far "
            "more on one routing than another</b> once transit time and reliability are priced in.")
    lane = df_f.groupby("Lane").agg(TEUs=("TEUs", "sum"), Freight=("FreightUSD", "sum"),
                                    Transit=("ActualTransitDays", "mean"), OnTime=("OnTimeFlag", "mean"),
                                    Voyages=("ShipmentID", "size")).reset_index()
    lane["CostPerTEU"] = lane.Freight / lane.TEUs
    lane["OnTime"] = lane.OnTime * 100
    g1, g2 = st.columns(2)
    with g1:
        sect("Top 12 lanes by volume")
        tv = lane.sort_values("TEUs", ascending=False).head(12)
        fl = px.bar(tv.sort_values("TEUs"), x="TEUs", y="Lane", orientation="h", color_discrete_sequence=[TEAL])
        fl.update_layout(template="plotly_white", height=380, margin=dict(l=10, r=10, t=10, b=10), yaxis_title="")
        st.plotly_chart(fl, width='stretch')
    with g2:
        sect("Cost/TEU vs transit time", "Bubble = volume; top-right lanes are slow and pricey")
        sc = px.scatter(lane, x="Transit", y="CostPerTEU", size="TEUs", color="OnTime",
                        hover_name="Lane", color_continuous_scale=[RED, AMBER, GREEN], size_max=34,
                        labels={"Transit": "Avg transit days", "CostPerTEU": "Cost / TEU ($)", "OnTime": "On-time %"})
        sc.update_layout(template="plotly_white", height=380, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(sc, width='stretch')
    sect("Lane detail")
    st.dataframe(lane.sort_values("TEUs", ascending=False).head(15)[
        ["Lane", "Voyages", "TEUs", "CostPerTEU", "Transit", "OnTime"]].style.format(
        {"TEUs": "{:,.0f}", "CostPerTEU": "${:,.0f}", "Transit": "{:.1f}", "OnTime": "{:.1f}%"}),
        width='stretch', hide_index=True)

# ============================================================ CARRIER SCORECARD
with tabs[2]:
    insight("Carrier choice is the biggest lever on reliability. <b>The trap is paying a premium carrier "
            "that does not deliver the premium service</b> — this view puts price and on-time side by side.")
    cs = df_f.groupby("CarrierName").agg(Voyages=("ShipmentID", "size"), TEUs=("TEUs", "sum"),
                                         OnTime=("OnTimeFlag", "mean"), Freight=("FreightUSD", "sum"),
                                         Damage=("DamageFlag", "mean")).reset_index()
    cs["OnTime"] = cs.OnTime * 100
    cs["CostPerTEU"] = cs.Freight / cs.TEUs
    cs["Damage"] = cs.Damage * 100
    g1, g2 = st.columns(2)
    with g1:
        sect("On-time % by carrier (ranked)")
        fb = px.bar(cs.sort_values("OnTime"), x="OnTime", y="CarrierName", orientation="h",
                    color="OnTime", color_continuous_scale=[RED, AMBER, GREEN], labels={"OnTime": "On-time %", "CarrierName": ""})
        fb.update_layout(template="plotly_white", height=360, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
        st.plotly_chart(fb, width='stretch')
    with g2:
        sect("Rate vs reliability", "Bottom-right = expensive AND unreliable (drop these)")
        sc = px.scatter(cs, x="CostPerTEU", y="OnTime", size="TEUs", color="CarrierName",
                        hover_name="CarrierName", color_discrete_sequence=VIVID, size_max=40,
                        labels={"CostPerTEU": "Cost / TEU ($)", "OnTime": "On-time %"})
        sc.update_layout(template="plotly_white", height=360, margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
        st.plotly_chart(sc, width='stretch')
    sect("Carrier table")
    st.dataframe(cs.sort_values("OnTime", ascending=False)[
        ["CarrierName", "Voyages", "TEUs", "OnTime", "CostPerTEU", "Damage"]].style.format(
        {"TEUs": "{:,.0f}", "OnTime": "{:.1f}%", "CostPerTEU": "${:,.0f}", "Damage": "{:.2f}%"}),
        width='stretch', hide_index=True)

# ============================================================ PORT DWELL & DELAYS
with tabs[3]:
    insight("Dwell time is the hidden cost of sitting at port. <b>Past the free days, every day is "
            "demurrage</b> — and the worst ports quietly drain budget while the freight rate looks fine.")
    arr = df_f[df_f.Status != "In Transit"]
    pv = arr.groupby("D_PortName").agg(Dwell=("PortDwellDays", "mean"), Dem=("DemurrageUSD", "sum"),
                                       Voyages=("ShipmentID", "size")).reset_index()
    g1, g2 = st.columns(2)
    with g1:
        sect("Avg dwell by destination port (worst 12)")
        tw = pv.sort_values("Dwell", ascending=False).head(12)
        fd = px.bar(tw.sort_values("Dwell"), x="Dwell", y="D_PortName", orientation="h",
                    color="Dwell", color_continuous_scale=["#7cc4b8", RED], labels={"Dwell": "Avg dwell (days)", "D_PortName": ""})
        fd.update_layout(template="plotly_white", height=360, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
        st.plotly_chart(fd, width='stretch')
    with g2:
        sect("Demurrage by port (top 12)")
        td = pv.sort_values("Dem", ascending=False).head(12)
        fde = px.bar(td.sort_values("Dem"), x="Dem", y="D_PortName", orientation="h", color_discrete_sequence=[AMBER],
                     labels={"Dem": "Demurrage ($)", "D_PortName": ""})
        fde.update_layout(template="plotly_white", height=360, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(fde, width='stretch')
    sect("Delay trend over time", "Average days late by arrival month")
    dt = df_f.groupby(["MonthSort", "MonthYear"]).agg(Delay=("Delay", "mean")).reset_index().sort_values("MonthSort")
    fdt = px.line(dt, x="MonthYear", y="Delay", labels={"Delay": "Avg days late", "MonthYear": ""})
    fdt.update_traces(line_color=RED, line_width=2)
    fdt.update_layout(template="plotly_white", height=280, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fdt, width='stretch')

# ============================================================ CONTAINER TRACKING
with tabs[4]:
    insight("Tracking is about exceptions, not the happy path. <b>Customs holds and documentation errors "
            "are where containers vanish into limbo</b> — surfacing their rate is the first step to cutting it.")
    sect("Pipeline status"); st.markdown("<hr>", unsafe_allow_html=True)
    k = st.columns(4)
    kpi(k[0], "In Transit", f"{int((df_f.Status=='In Transit').sum()):,}", "on the water")
    kpi(k[1], "At Port", f"{int((df_f.Status=='At Port').sum()):,}", "discharging / dwell")
    kpi(k[2], "Customs Hold Rate", f"{df_f.CustomsHoldFlag.mean()*100:.1f}%", "flagged by customs")
    kpi(k[3], "Doc Error Rate", f"{df_f.DocErrorFlag.mean()*100:.1f}%", "B/L or paperwork")
    st.markdown("<hr>", unsafe_allow_html=True)
    sect("Status mix by arrival month")
    sm = df_f.groupby(["MonthYear", "MonthSort", "Status"]).size().reset_index(name="n").sort_values("MonthSort")
    fs = px.area(sm, x="MonthYear", y="n", color="Status",
                 color_discrete_map={"In Transit": TEAL, "At Port": AMBER, "Cleared": "#8aa1b8"},
                 labels={"n": "Shipments", "MonthYear": ""})
    fs.update_layout(template="plotly_white", height=320, margin=dict(l=10, r=10, t=10, b=10), legend_title="")
    st.plotly_chart(fs, width='stretch')
    sect("Exception rates by destination region")
    ex = df_f.groupby("D_Region").agg(Customs=("CustomsHoldFlag", "mean"), Doc=("DocErrorFlag", "mean"),
                                      Damage=("DamageFlag", "mean")).reset_index()
    for c in ["Customs", "Doc", "Damage"]:
        ex[c] = ex[c] * 100
    fe = px.bar(ex.melt(id_vars="D_Region", var_name="Exception", value_name="Rate"),
                x="D_Region", y="Rate", color="Exception", barmode="group",
                color_discrete_sequence=[RED, AMBER, "#8aa1b8"], labels={"D_Region": "", "Rate": "Rate %"})
    fe.update_layout(template="plotly_white", height=300, margin=dict(l=10, r=10, t=10, b=10), legend_title="")
    st.plotly_chart(fe, width='stretch')

# ============================================================ COST ANALYSIS
with tabs[5]:
    insight("Ocean cost is freight + demurrage. <b>Freight you negotiate once; demurrage you bleed daily</b> "
            "— and it is the line most teams never put on a chart, so it never gets fixed.")
    sect("Cost summary"); st.markdown("<hr>", unsafe_allow_html=True)
    fr, dem = df_f.FreightUSD.sum(), df_f.DemurrageUSD.sum()
    k = st.columns(4)
    kpi(k[0], "Total Freight", f"${fr/1e6:,.1f}M", "negotiated rate")
    kpi(k[1], "Total Demurrage", f"${dem/1e6:,.2f}M", f"{dem/(fr+dem)*100:.1f}% of ocean cost")
    kpi(k[2], "Cost / TEU", f"${(fr+dem)/df_f.TEUs.sum():,.0f}", "all-in")
    kpi(k[3], "Demurrage / Voyage", f"${dem/len(df_f):,.0f}", "avg penalty")
    st.markdown("<hr>", unsafe_allow_html=True)
    g1, g2 = st.columns(2)
    with g1:
        sect("Freight vs demurrage by month")
        cm = df_f.groupby(["MonthSort", "MonthYear"]).agg(Freight=("FreightUSD", "sum"),
                                                          Demurrage=("DemurrageUSD", "sum")).reset_index().sort_values("MonthSort")
        fc = go.Figure()
        fc.add_trace(go.Bar(x=cm.MonthYear, y=cm.Freight, name="Freight", marker_color=TEAL))
        fc.add_trace(go.Bar(x=cm.MonthYear, y=cm.Demurrage, name="Demurrage", marker_color=RED))
        fc.update_layout(template="plotly_white", height=320, margin=dict(l=10, r=10, t=10, b=10),
                         barmode="stack", yaxis_title="USD", legend=dict(orientation="h", y=1.15))
        st.plotly_chart(fc, width='stretch')
    with g2:
        sect("Cost per TEU by lane (worst 12)")
        lc = df_f.groupby("Lane").agg(Freight=("FreightUSD", "sum"), Dem=("DemurrageUSD", "sum"),
                                      TEUs=("TEUs", "sum")).reset_index()
        # NOTE: never name a column "T" — df.T is the transpose attribute, not the column.
        lc["CostPerTEU"] = (lc["Freight"] + lc["Dem"]) / lc["TEUs"]
        tw = lc.sort_values("CostPerTEU", ascending=False).head(12)
        flc = px.bar(tw.sort_values("CostPerTEU"), x="CostPerTEU", y="Lane", orientation="h",
                     color="CostPerTEU", color_continuous_scale=["#7cc4b8", RED], labels={"CostPerTEU": "Cost / TEU ($)", "Lane": ""})
        flc.update_layout(template="plotly_white", height=320, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
        st.plotly_chart(flc, width='stretch')
    insight(f"Demurrage is <b>{dem/(fr+dem)*100:.1f}%</b> of total ocean cost here — small next to freight, "
            f"but unlike freight it is <b>100% avoidable</b> with faster gate-out and better free-day usage. "
            f"That is the cheapest money on this dashboard to go and recover.")

st.markdown('<div class="foot">Ocean Freight &amp; Port Operations • Built with Streamlit + Plotly • '
            '20k-voyage synthetic dataset • port coords/ISO-3 real, rates are labelled assumptions</div>',
            unsafe_allow_html=True)