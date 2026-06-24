"""
Demand & Inventory Forecasting — Operations Center (multi-page)
---------------------------------------------------------------
Streamlit + Plotly. Five working pages driven by the purple sidebar nav, all built
on Demand_Inventory_Dataset.xlsx (the same 50-SKU catalogue as the freight data).

RUN:
    pip install streamlit pandas plotly openpyxl
    streamlit run demand_forecast_dashboard.py

Workbook must sit beside this script. No upload prompt. Dataset is synthetic for demo;
every figure shown is computed from it.
"""

from pathlib import Path
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="Demand & Inventory Forecasting", layout="wide", page_icon="📦")

PURPLE, INK, GREEN, AMBER, RED, GREY = "#6c4ee0", "#1f2440", "#2bb673", "#f0a830", "#e5556b", "#9aa3bd"
STATUS_COLOR = {"Healthy": GREEN, "Reorder Now": AMBER, "Stockout Risk": RED, "Excess": "#8a7bd8"}

st.markdown(f"""
<style>
  .block-container {{ padding-top: 3.5rem; max-width: 1600px; }}
  section[data-testid="stSidebar"] {{ background: linear-gradient(180deg,#5b3df5,#6c4ee0); }}
  section[data-testid="stSidebar"] * {{ color: #f3f1ff !important; }}
  .side-logo {{ font-weight:800; font-size:1.05rem; letter-spacing:.5px; margin:4px 0 10px 0; }}
  .h-title {{ color:{INK}; font-size:1.9rem; font-weight:800; margin:0; }}
  .h-sub {{ color:{GREY}; margin:2px 0 6px 0; }}
  .card {{ background:#fff; border:1px solid #ececf6; border-radius:14px; padding:14px 16px;
           box-shadow:0 1px 3px rgba(31,36,64,.05); }}
  .kpi-l {{ color:{GREY}; font-size:.78rem; }}
  .kpi-v {{ color:{INK}; font-size:1.6rem; font-weight:800; line-height:1.05; }}
  .delta-up {{ color:{GREEN}; font-size:.78rem; font-weight:700; }}
  .delta-dn {{ color:{RED}; font-size:.78rem; font-weight:700; }}
  .delta-n {{ color:{GREY}; font-size:.78rem; }}
  .panel {{ background:#f6f4ff; border:1px solid #e4ddff; border-radius:14px; padding:13px 16px; margin-bottom:12px; }}
  .panel h4 {{ color:{PURPLE}; margin:0 0 6px 0; font-size:1rem; }}
  .pitch {{ background:linear-gradient(90deg,#f1ecff,#fff 70%); border-left:4px solid {PURPLE};
            border-radius:10px; padding:12px 16px; color:{INK}; font-size:.92rem; line-height:1.55;
            margin:6px 0 14px 0; }}
  .pitch b {{ color:{PURPLE}; }}
  .alert {{ background:#fff; border-radius:10px; padding:9px 12px; margin:6px 0;
            border-left:4px solid {GREY}; font-size:.86rem; color:{INK}; }}
  .blk {{ font-weight:800; }}
  .sect {{ color:{INK}; font-weight:800; font-size:1.05rem; margin:8px 0 4px 0; }}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
DATA = Path(__file__).parent / "Demand_Inventory_Dataset.xlsx"
if not DATA.exists():
    st.error(f"Cannot find {DATA.name}. Place it beside this script and rerun.")
    st.stop()


@st.cache_data(show_spinner=False)
def load():
    xls = pd.ExcelFile(DATA)
    fc = pd.read_excel(xls, "Fact_DemandForecast", parse_dates=["WeekStart"])
    inv = pd.read_excel(xls, "Dim_Inventory")
    return fc, inv


fc, inv = load()
WEEKS = sorted(fc.WeekStart.unique())


def mape_of(d):
    return (d.ActualDemand.sub(d.ForecastDemand).abs().sum() / d.ActualDemand.sum() * 100) if len(d) and d.ActualDemand.sum() else np.nan


def card(col, ico, label, value, sub=""):
    col.markdown(f'<div class="card">{ico} <span class="kpi-l">{label}</span><br>'
                 f'<span class="kpi-v">{value}</span><br>{sub}</div>', unsafe_allow_html=True)


def delta(v, good_up=True, suffix="pts"):
    if pd.isna(v):
        return '<span class="delta-n">—</span>'
    up = v >= 0
    good = up if good_up else not up
    return (f'<span class="{"delta-up" if good else "delta-dn"}">{"▲" if up else "▼"} '
            f'{abs(v):.1f} {suffix}</span>')


def pitch(html):
    st.markdown(f'<div class="pitch">{html}</div>', unsafe_allow_html=True)


def band_fig(t, height=340):
    f = go.Figure()
    f.add_trace(go.Scatter(x=t.WeekStart, y=t.ForecastHigh, mode="lines", line=dict(width=0),
                           showlegend=False, hoverinfo="skip"))
    f.add_trace(go.Scatter(x=t.WeekStart, y=t.ForecastLow, mode="lines", line=dict(width=0),
                           fill="tonexty", fillcolor="rgba(108,78,224,.15)", name="Confidence band",
                           hoverinfo="skip"))
    f.add_trace(go.Scatter(x=t.WeekStart, y=t.ForecastDemand, mode="lines",
                           line=dict(color=PURPLE, width=2, dash="dash"), name="Forecast"))
    f.add_trace(go.Scatter(x=t.WeekStart, y=t.ActualDemand, mode="lines",
                           line=dict(color=INK, width=3), name="Actual demand"))
    f.update_layout(template="plotly_white", height=height, margin=dict(l=10, r=10, t=10, b=10),
                    legend=dict(orientation="h", y=1.13), yaxis_title="Units")
    return f


# ----------------------------------------------------------------------------
# Sidebar: real navigation + global filters
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown('<div class="side-logo">📦 DEMAND &amp; INVENTORY</div>', unsafe_allow_html=True)
    page = st.radio("Navigate", ["Forecast Overview", "SKU Forecast", "Replenishment",
                                 "Accuracy Tracker", "Capacity"], label_visibility="collapsed")
    st.markdown("---")
    cat = st.selectbox("Product Category", ["All"] + sorted(inv.Category.unique()))
    abc = st.selectbox("ABC Class", ["All", "A", "B", "C"])
    st.caption(f"As of {pd.Timestamp(WEEKS[-1]).date()} · synthetic demo data")

inv_f = inv.copy()
if cat != "All":
    inv_f = inv_f[inv_f.Category == cat]
if abc != "All":
    inv_f = inv_f[inv_f.ABCClass == abc]
fc_f = fc[fc.SKU.isin(set(inv_f.SKU))]
recent = fc_f[fc_f.WeekStart >= WEEKS[-26]]

# ============================================================================
# PAGE 1 — FORECAST OVERVIEW
# ============================================================================
if page == "Forecast Overview":
    st.markdown('<div class="h-title">Forecast Overview</div>'
                '<div class="h-sub">The whole portfolio at a glance — accuracy, risk, and what needs ordering.</div>',
                unsafe_allow_html=True)
    pitch("Forecasting is not an academic score — it is cash. <b>Every point of accuracy you gain "
          "lets you hold less safety stock without risking a stockout</b>, and every stockout you "
          "prevent is a sale you keep. This page turns the model's performance into the two numbers a "
          "planner acts on daily: what to trust, and what to reorder now.")

    mape = mape_of(recent); acc = 100 - mape
    l4, p4 = fc_f[fc_f.WeekStart >= WEEKS[-4]], fc_f[(fc_f.WeekStart >= WEEKS[-8]) & (fc_f.WeekStart < WEEKS[-4])]
    acc_d = (100 - mape_of(l4)) - (100 - mape_of(p4))
    dvf = (recent.ActualDemand.sum() - recent.ForecastDemand.sum()) / recent.ForecastDemand.sum() * 100
    n_so = int((inv_f.Status == "Stockout Risk").sum())
    excess = inv_f.ExcessStockValueUSD.sum()
    n_re = int((inv_f.ReplenishmentDueUnits > 0).sum())

    r1 = st.columns(4)
    card(r1[0], "🎯", "Forecast Accuracy", f"{acc:.1f}%", delta(acc_d, True))
    card(r1[1], "📉", "MAPE", f"{mape:.1f}%", delta(mape_of(l4) - mape_of(p4), False))
    card(r1[2], "⚠️", "Stockout-Risk SKUs", f"{n_so}", '<span class="delta-n">below safety stock</span>')
    card(r1[3], "💰", "Excess Stock Value", f"${excess/1e3:,.0f}k", '<span class="delta-n">capital tied up</span>')

    left, rail = st.columns([3.4, 1.25])
    with left:
        st.markdown('<div class="sect">Demand vs Forecast (with confidence band)</div>', unsafe_allow_html=True)
        t = fc_f.groupby("WeekStart")[["ActualDemand", "ForecastDemand", "ForecastLow", "ForecastHigh"]].sum().reset_index()
        st.plotly_chart(band_fig(t), use_container_width=True)
        g1, g2 = st.columns(2)
        with g1:
            st.markdown('<div class="sect">Inventory Status</div>', unsafe_allow_html=True)
            sc = inv_f.Status.value_counts().reindex(list(STATUS_COLOR), fill_value=0)
            d = go.Figure(go.Pie(labels=sc.index, values=sc.values, hole=.6,
                                 marker=dict(colors=[STATUS_COLOR[s] for s in sc.index])))
            d.update_layout(template="plotly_white", height=300, margin=dict(l=0, r=0, t=10, b=0),
                            annotations=[dict(text=f"{len(inv_f)}<br>SKUs", showarrow=False, font_size=16)])
            st.plotly_chart(d, use_container_width=True)
        with g2:
            st.markdown('<div class="sect">Forecast Error vs SKU Value</div>', unsafe_allow_html=True)
            sca = px.scatter(inv_f, x="UnitValueUSD", y="MAPE", size="AvgWeeklyDemand", color="Status",
                             color_discrete_map=STATUS_COLOR, hover_name="ProductName", size_max=32,
                             labels={"UnitValueUSD": "Unit value ($)", "MAPE": "MAPE %"})
            sca.update_layout(template="plotly_white", height=300, margin=dict(l=10, r=10, t=10, b=10), legend_title="")
            st.plotly_chart(sca, use_container_width=True)
        st.markdown('<div class="sect">Replenishment Planner (urgency)</div>', unsafe_allow_html=True)
        rep = (inv_f[inv_f.Status.isin(["Stockout Risk", "Reorder Now"])].sort_values("DaysOfCover")
               [["SKU", "ProductName", "OnHandUnits", "ReorderPoint", "DaysOfCover", "LeadTimeDays",
                 "ReplenishmentDueUnits", "Status"]])
        if rep.empty:
            st.info("Nothing at or below reorder point for this selection.")
        else:
            st.dataframe(rep.style.apply(
                lambda r: [f"background-color:{'#fdecee' if r.Status=='Stockout Risk' else '#fdf3e2'}"] * len(r),
                axis=1).format({"DaysOfCover": "{:.1f}"}), use_container_width=True, hide_index=True)
    with rail:
        wc = inv_f.groupby("Category")["MAPE"].mean()
        bias = (recent.ForecastDemand.sum() - recent.ActualDemand.sum()) / recent.ActualDemand.sum() * 100
        urg = inv_f.sort_values("DaysOfCover").head(1)
        utxt = (f"{urg.iloc[0]['ProductName']} has {urg.iloc[0]['DaysOfCover']:.0f} days of cover left"
                if len(urg) else "no SKU is critically low")
        st.markdown(f'<div class="panel"><h4>✨ Forecast Insights</h4><b>{wc.idxmax()}</b> is hardest to '
                    f'forecast ({wc.max():.0f}% MAPE); the plan is '
                    f'{"over" if bias>0 else "under"}-forecasting by {abs(bias):.1f}% overall. {utxt} — '
                    f'release that order first.</div>', unsafe_allow_html=True)
        n_hi = int((inv_f.MAPE > 25).sum())
        st.markdown('<div class="panel"><h4>🔔 Critical Alerts</h4></div>', unsafe_allow_html=True)
        for c, t_ in [(RED, f'<b class="blk">{n_so}</b> SKUs below safety stock'),
                      (AMBER, f'<b class="blk">${excess/1e3:,.0f}k</b> in excess stock'),
                      (RED if n_hi else GREEN, f'<b class="blk">{n_hi}</b> SKUs MAPE &gt; 25%'),
                      (PURPLE, f'<b class="blk">{n_re}</b> replenishment orders due')]:
            st.markdown(f'<div class="alert" style="border-left-color:{c}">{t_}</div>', unsafe_allow_html=True)

# ============================================================================
# PAGE 2 — SKU FORECAST (single-SKU deep dive)
# ============================================================================
elif page == "SKU Forecast":
    st.markdown('<div class="h-title">SKU Forecast</div>'
                '<div class="h-sub">Deep-dive one product: how it sells, how well we predict it, and where it stands.</div>',
                unsafe_allow_html=True)
    pitch("Portfolio averages hide the SKUs that hurt. <b>A single high-value product with a drifting "
          "forecast can swallow more working capital than a hundred well-behaved ones.</b> This view "
          "lets a planner interrogate one SKU end to end — demand shape, forecast bias, and stock "
          "position — before committing a purchase order.")

    skus = inv_f.sort_values("AnnualUsageValue", ascending=False)
    sku = st.selectbox("Choose SKU (sorted by annual usage value)",
                       skus.SKU + " — " + skus.ProductName).split(" — ")[0]
    row = inv[inv.SKU == sku].iloc[0]
    k = st.columns(4)
    card(k[0], "🎯", "Accuracy", f"{row.ForecastAccuracy:.0f}%", f'<span class="delta-n">MAPE {row.MAPE:.0f}%</span>')
    card(k[1], "📦", "On Hand", f"{int(row.OnHandUnits):,}", f'<span class="delta-n">reorder at {int(row.ReorderPoint):,}</span>')
    card(k[2], "🛡️", "Days of Cover", f"{row.DaysOfCover:.0f}", f'<span class="delta-n">lead time {int(row.LeadTimeDays)}d</span>')
    status_c = STATUS_COLOR.get(row.Status, GREY)
    card(k[3], "🚦", "Status", f'<span style="color:{status_c}">{row.Status}</span>',
         f'<span class="delta-n">{row.Category} · class {row.ABCClass}</span>')

    st.markdown('<div class="sect">Actual vs Forecast</div>', unsafe_allow_html=True)
    t = fc[fc.SKU == sku].sort_values("WeekStart")
    st.plotly_chart(band_fig(t, height=360), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sect">Forecast residual (Actual − Forecast)</div>', unsafe_allow_html=True)
        t = t.assign(resid=t.ActualDemand - t.ForecastDemand)
        rf = px.bar(t, x="WeekStart", y="resid", color=t.resid >= 0,
                    color_discrete_map={True: GREEN, False: RED}, labels={"resid": "Units", "WeekStart": ""})
        rf.update_layout(template="plotly_white", height=280, margin=dict(l=10, r=10, t=10, b=10), showlegend=False)
        st.plotly_chart(rf, use_container_width=True)
    with c2:
        st.markdown('<div class="sect">Position vs thresholds</div>', unsafe_allow_html=True)
        bars = pd.DataFrame({"Level": ["Safety stock", "Reorder point", "On hand"],
                             "Units": [row.SafetyStock, row.ReorderPoint, row.OnHandUnits]})
        pf = px.bar(bars, x="Units", y="Level", orientation="h",
                    color="Level", color_discrete_sequence=[RED, AMBER, PURPLE])
        pf.update_layout(template="plotly_white", height=280, margin=dict(l=10, r=10, t=10, b=10),
                         showlegend=False, yaxis_title="")
        st.plotly_chart(pf, use_container_width=True)
    bias = (t.ForecastDemand.sum() - t.ActualDemand.sum()) / t.ActualDemand.sum() * 100
    pitch(f"Over the last 18 months this SKU was {'over' if bias>0 else 'under'}-forecast by "
          f"<b>{abs(bias):.1f}%</b>. At {row.DaysOfCover:.0f} days of cover against a {int(row.LeadTimeDays)}-day "
          f"lead time, the buying decision is { 'urgent — reorder now' if row.Status in ('Stockout Risk','Reorder Now') else 'comfortable for now'}.")

# ============================================================================
# PAGE 3 — REPLENISHMENT (PO planner with EOQ)
# ============================================================================
elif page == "Replenishment":
    st.markdown('<div class="h-title">Replenishment Planner</div>'
                '<div class="h-sub">Turn reorder points into a ranked, costed purchase-order list.</div>',
                unsafe_allow_html=True)
    pitch("A reorder flag is a question; <b>a purchase order is the answer.</b> This page closes that gap "
          "— ranking every SKU by how soon it runs out and sizing the order with classic economic order "
          "quantity so you are not over-ordering cash into a warehouse or under-ordering into a stockout.")

    c1, c2 = st.columns(2)
    order_cost = c1.slider("Ordering cost per PO ($)", 25, 300, 75, 5)
    hold_rate = c2.slider("Annual holding cost (% of unit value)", 5, 50, 25, 1) / 100

    need = inv_f[inv_f.Status.isin(["Stockout Risk", "Reorder Now"])].copy()
    need["AnnualDemand"] = need.AvgWeeklyDemand * 52
    H = (hold_rate * need.UnitValueUSD).replace(0, np.nan)
    need["EOQ"] = np.sqrt(2 * need.AnnualDemand * order_cost / H).round(0)
    need["SuggestedOrderQty"] = np.maximum(need.EOQ, need.ReplenishmentDueUnits).round(0)
    need["OrderValueUSD"] = (need.SuggestedOrderQty * need.UnitValueUSD).round(0)

    k = st.columns(3)
    card(k[0], "🧾", "SKUs to Order", f"{len(need)}")
    card(k[1], "📦", "Units to Order", f"{int(need.SuggestedOrderQty.sum()):,}")
    card(k[2], "💵", "Total Order Value", f"${need.OrderValueUSD.sum()/1e3:,.0f}k")

    cL, cR = st.columns([3, 2])
    with cL:
        st.markdown('<div class="sect">Order queue (most urgent first)</div>', unsafe_allow_html=True)
        if need.empty:
            st.info("No SKUs need replenishing for this selection.")
        else:
            tbl = need.sort_values("DaysOfCover")[
                ["SKU", "ProductName", "DaysOfCover", "LeadTimeDays",
                 "EOQ", "SuggestedOrderQty", "OrderValueUSD", "Status"]]
            sty = tbl.style.apply(
                lambda r: [f"background-color:{'#fdecee' if r.Status=='Stockout Risk' else '#fdf3e2'}"] * len(r),
                axis=1).format({"DaysOfCover": "{:.1f}", "OrderValueUSD": "${:,.0f}",
                                "EOQ": "{:,.0f}", "SuggestedOrderQty": "{:,.0f}",
                                "LeadTimeDays": "{:d}"})
            st.dataframe(sty, use_container_width=True, hide_index=True, column_config={
                "ProductName": st.column_config.Column("Product", width="medium"),
                "DaysOfCover": st.column_config.Column("Days cover"),
                "LeadTimeDays": st.column_config.Column("Lead (d)"),
                "SuggestedOrderQty": st.column_config.Column("Order qty"),
                "OrderValueUSD": st.column_config.Column("Order $"),
                "Status": st.column_config.Column("Status", width="small"),
            })
    with cR:
        st.markdown('<div class="sect">Order value by category</div>', unsafe_allow_html=True)
        if not need.empty:
            byc = need.groupby("Category")["OrderValueUSD"].sum().sort_values().reset_index()
            byc["lbl"] = byc.OrderValueUSD.map(lambda v: f"${v/1e3:,.0f}k")
            bf = px.bar(byc, x="OrderValueUSD", y="Category", orientation="h", text="lbl",
                        color="OrderValueUSD", color_continuous_scale=["#cdbff5", PURPLE],
                        labels={"OrderValueUSD": "Order value ($)", "Category": ""})
            bf.update_traces(textposition="outside", cliponaxis=False)
            bf.update_layout(template="plotly_white", height=420, margin=dict(l=10, r=30, t=10, b=30),
                             coloraxis_showscale=False, xaxis_title="Order value ($)")
            st.plotly_chart(bf, use_container_width=True)
    pitch("EOQ here uses your two assumptions above — change them and the order sizes move. "
          "<b>It assumes steady demand and no quantity discounts</b>, so treat the quantities as a "
          "disciplined starting point for negotiation, not gospel.")

# ============================================================================
# PAGE 4 — ACCURACY TRACKER
# ============================================================================
elif page == "Accuracy Tracker":
    st.markdown('<div class="h-title">Accuracy Tracker</div>'
                '<div class="h-sub">Is the model improving or drifting? Catch it before it costs you.</div>',
                unsafe_allow_html=True)
    pitch("A forecast you do not monitor quietly rots. <b>Demand patterns shift, a model that was sharp "
          "last quarter starts missing</b> — and the first symptom is usually a stockout, not a warning. "
          "Tracking accuracy over time is the cheap insurance that turns a silent failure into a flag.")

    wk = fc_f.groupby("WeekStart").apply(
        lambda g: pd.Series({"MAPE": g.ActualDemand.sub(g.ForecastDemand).abs().sum() / g.ActualDemand.sum() * 100,
                             "Bias": (g.ForecastDemand.sum() - g.ActualDemand.sum()) / g.ActualDemand.sum() * 100})
    ).reset_index()
    wk["Accuracy"] = 100 - wk.MAPE
    wk["Acc_roll"] = wk.Accuracy.rolling(4, min_periods=1).mean()

    k = st.columns(3)
    cur = wk.Accuracy.tail(4).mean(); prev = wk.Accuracy.iloc[-8:-4].mean()
    card(k[0], "🎯", "Accuracy (last 4 wk)", f"{cur:.1f}%", delta(cur - prev, True))
    card(k[1], "📉", "Avg MAPE (26 wk)", f"{wk.MAPE.tail(26).mean():.1f}%")
    card(k[2], "⚖️", "Forecast Bias (26 wk)", f"{wk.Bias.tail(26).mean():+.1f}%",
         '<span class="delta-n">+ = over-forecast</span>')

    st.markdown('<div class="sect">Forecast accuracy over time (4-week rolling)</div>', unsafe_allow_html=True)
    af = go.Figure()
    af.add_trace(go.Scatter(x=wk.WeekStart, y=wk.Accuracy, mode="lines", name="Weekly",
                            line=dict(color="#cdbff5", width=1)))
    af.add_trace(go.Scatter(x=wk.WeekStart, y=wk.Acc_roll, mode="lines", name="4-wk rolling",
                            line=dict(color=PURPLE, width=3)))
    af.update_layout(template="plotly_white", height=320, margin=dict(l=10, r=10, t=10, b=10),
                     yaxis_title="Accuracy %", legend=dict(orientation="h", y=1.13))
    st.plotly_chart(af, use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sect">Accuracy by category</div>', unsafe_allow_html=True)
        bc = inv_f.groupby("Category")["ForecastAccuracy"].mean().sort_values().reset_index()
        cf = px.bar(bc, x="ForecastAccuracy", y="Category", orientation="h", color="ForecastAccuracy",
                    color_continuous_scale=["#e5556b", "#f0a830", "#2bb673"],
                    labels={"ForecastAccuracy": "Accuracy %", "Category": ""})
        cf.update_layout(template="plotly_white", height=320, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
        st.plotly_chart(cf, use_container_width=True)
    with c2:
        st.markdown('<div class="sect">Worst-forecast SKUs (highest MAPE)</div>', unsafe_allow_html=True)
        worst = inv_f.nlargest(10, "MAPE")[["ProductName", "MAPE", "Category", "Status"]]
        st.dataframe(worst.style.format({"MAPE": "{:.1f}%"}), use_container_width=True, hide_index=True)
    pitch(f"The 4-week line is what matters: a steady or rising rolling accuracy means the model is "
          f"trustworthy; a falling one is your cue to retrain or override. Right now the weakest "
          f"category is <b>{inv_f.groupby('Category')['ForecastAccuracy'].mean().idxmin()}</b> — fix that "
          f"and the portfolio number follows.")

# ============================================================================
# PAGE 5 — CAPACITY (inventory investment & space proxy)
# ============================================================================
elif page == "Capacity":
    st.markdown('<div class="h-title">Capacity &amp; Working Capital</div>'
                '<div class="h-sub">How much cash and shelf your stock ties up — and what tightening cover frees.</div>',
                unsafe_allow_html=True)
    pitch("Inventory is frozen cash. <b>Every unit sitting beyond what demand justifies is money that "
          "could be financing growth instead of gathering dust</b> — and warehouse space you are paying "
          "to keep occupied. This page sizes that, then lets you test how much you'd release by holding "
          "a tighter days-of-cover target.")

    inv_c = inv_f.copy()
    inv_c["InvValue"] = inv_c.OnHandUnits * inv_c.UnitValueUSD
    k = st.columns(4)
    card(k[0], "📦", "Units On Hand", f"{int(inv_c.OnHandUnits.sum()):,}")
    card(k[1], "💵", "Inventory Value", f"${inv_c.InvValue.sum()/1e6:,.2f}M")
    card(k[2], "🧊", "Excess Value", f"${inv_c.ExcessStockValueUSD.sum()/1e3:,.0f}k")
    card(k[3], "🔢", "SKUs Held", f"{len(inv_c)}")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="sect">Inventory value by category (space &amp; cash proxy)</div>', unsafe_allow_html=True)
        bc = inv_c.groupby("Category")["InvValue"].sum().sort_values().reset_index()
        vf = px.bar(bc, x="InvValue", y="Category", orientation="h", color="InvValue",
                    color_continuous_scale=["#cdbff5", PURPLE], labels={"InvValue": "Inventory value ($)", "Category": ""})
        vf.update_layout(template="plotly_white", height=340, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
        st.plotly_chart(vf, use_container_width=True)
    with c2:
        st.markdown('<div class="sect">What-if: tighten the days-of-cover target</div>', unsafe_allow_html=True)
        target = st.slider("Target days of cover", 20, 120, 60, 5)
        daily = (inv_c.AvgWeeklyDemand / 7).replace(0, np.nan)
        target_units = daily * target
        inv_c["FreedUnits"] = (inv_c.OnHandUnits - target_units).clip(lower=0)
        inv_c["FreedValue"] = inv_c.FreedUnits * inv_c.UnitValueUSD
        freed = inv_c.FreedValue.sum()
        st.metric(f"Capital freed at {target}-day cover", f"${freed/1e3:,.0f}k",
                  f"{freed / inv_c.InvValue.sum() * 100:.0f}% of inventory value")
        topf = inv_c.nlargest(8, "FreedValue")[["ProductName", "FreedValue"]]
        ff = px.bar(topf.sort_values("FreedValue"), x="FreedValue", y="ProductName", orientation="h",
                    color_discrete_sequence=[GREEN], labels={"FreedValue": "Capital freed ($)", "ProductName": ""})
        ff.update_layout(template="plotly_white", height=240, margin=dict(l=10, r=10, t=10, b=10))
        st.plotly_chart(ff, use_container_width=True)
    pitch("This treats units and dollars as the capacity proxy — the dataset has no warehouse m³ feed. "
          "Wire in real bin/volume data and the same logic reports cubic space freed alongside the cash. "
          "<b>The slider is the pitch: show a CFO the exact capital a tighter policy releases, per SKU.</b>")