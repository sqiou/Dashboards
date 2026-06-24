"""
Freight & Landed Cost Analytics Dashboard
-----------------------------------------
A Streamlit + Plotly app that mirrors the layout of the US Flight Analytics
dashboard, but driven by Freight_Landed_Cost_Dataset.xlsx.

RUN:
    pip install streamlit pandas plotly openpyxl
    streamlit run freight_dashboard.py

Place Freight_Landed_Cost_Dataset.xlsx in the same folder (or upload it when prompted).
"""

import pandas as pd
import plotly.express as px
import streamlit as st
from pathlib import Path

# ----------------------------------------------------------------------------
# Page config + styling (light theme, teal accent — matches the template)
# ----------------------------------------------------------------------------
st.set_page_config(page_title="Freight & Landed Cost Analytics",
                   layout="wide", page_icon="🚢")

TEAL = "#1a9e8f"
DARK = "#1f2a44"
TEAL_SCALE = "Tealgrn"
DELAY_SCALE = ["#0b6e63", "#2bb6a3", "#e8c33f", "#e07b39", "#d6455d"]  # low->high
VIVID = px.colors.qualitative.Bold + px.colors.qualitative.Vivid

st.markdown(f"""
<style>
    .block-container {{ padding-top: 1.2rem; max-width: 1500px; }}
    .header-card {{
        border-left: 6px solid {TEAL};
        background: linear-gradient(90deg, #f1faf8 0%, #ffffff 60%);
        padding: 18px 24px; border-radius: 6px; margin-bottom: 10px;
        border-top: 3px solid {TEAL};
    }}
    .header-card h1 {{ color:{DARK}; font-size: 2.0rem; margin:0; font-weight:800; }}
    .header-card p  {{ color:{TEAL}; font-weight:600; margin:2px 0 0 0; font-size:.85rem; }}
    .section-title {{ color:{DARK}; font-weight:800; font-size:1.4rem; margin:6px 0 2px 0; }}
    .section-sub   {{ color:#5e667d; font-size:.85rem; margin-bottom:8px; }}
    .kpi-label {{ color:#5e667d; font-size:.8rem; margin-bottom:2px; }}
    .kpi-value {{ color:{DARK}; font-size:2.0rem; font-weight:800; line-height:1.1; }}
    .tipbox {{ background:#e8f6fb; border-left:4px solid #3aa6c4; padding:10px 14px;
              border-radius:4px; color:{DARK}; font-size:.9rem; margin:6px 0 14px 0; }}
    .insight {{ background:#f1faf8; border:1px solid #cfeae4; border-left:4px solid {TEAL};
               padding:12px 16px; border-radius:6px; color:{DARK}; font-size:.9rem;
               margin:10px 0 6px 0; line-height:1.5; }}
    .insight b {{ color:{TEAL}; }}
    .foot {{ text-align:center; color:#8c98a8; font-size:.8rem; padding:18px 0; }}
    hr {{ margin: 8px 0 14px 0; }}
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# Data loading
# ----------------------------------------------------------------------------
ISO2_NAME = {
    'CN': 'China', 'KR': 'South Korea', 'VN': 'Vietnam', 'SG': 'Singapore', 'IN': 'India',
    'PK': 'Pakistan', 'AE': 'United Arab Emirates', 'TR': 'Turkey', 'DE': 'Germany',
    'NL': 'Netherlands', 'BE': 'Belgium', 'JP': 'Japan', 'TW': 'Taiwan', 'TH': 'Thailand',
    'PH': 'Philippines', 'LK': 'Sri Lanka', 'MY': 'Malaysia', 'MX': 'Mexico', 'BD': 'Bangladesh',
    'ID': 'Indonesia', 'AU': 'Australia', 'BR': 'Brazil', 'PE': 'Peru', 'CO': 'Colombia',
    'KE': 'Kenya', 'EG': 'Egypt', 'GR': 'Greece', 'ES': 'Spain', 'FR': 'France', 'PL': 'Poland',
    'GB': 'United Kingdom', 'US': 'United States', 'CA': 'Canada', 'ZA': 'South Africa', 'IT': 'Italy'
}


@st.cache_data(show_spinner=False)
def load_data(source):
    xls = pd.ExcelFile(source)
    fact = pd.read_excel(xls, "Fact_Shipment")
    dprod = pd.read_excel(xls, "Dim_Product")
    dcar = pd.read_excel(xls, "Dim_Carrier")
    dsup = pd.read_excel(xls, "Dim_Supplier")
    dlane = pd.read_excel(xls, "Dim_Lane")
    dmode = pd.read_excel(xls, "Dim_Mode")
    dinc = pd.read_excel(xls, "Dim_Incoterm")
    ddate = pd.read_excel(xls, "Dim_Date")

    df = (fact
          .merge(dprod[['ProductKey', 'Category', 'SubCategory', 'ProductName']], on='ProductKey', how='left')
          .merge(dcar[['CarrierKey', 'CarrierName', 'CarrierType', 'CarrierTier']], on='CarrierKey', how='left')
          .merge(dsup[['SupplierKey', 'SupplierName', 'SupplierRegion']], on='SupplierKey', how='left')
          .merge(dlane[['LaneKey', 'LaneName', 'TradeLane', 'OriginCountry', 'OriginRegion',
                        'DestCountry', 'DestRegion', 'DestPort', 'OriginPort']], on='LaneKey', how='left')
          .merge(dmode[['ModeKey', 'Mode', 'ModeCategory']], on='ModeKey', how='left')
          .merge(dinc[['IncotermKey', 'Incoterm']], on='IncotermKey', how='left')
          .merge(ddate[['DateKey', 'Date', 'Year', 'MonthYear', 'MonthSort']],
                 left_on='ShipDateKey', right_on='DateKey', how='left'))

    df['TransitDelay'] = df['ActualTransitDays'] - df['PlannedTransitDays']
    df['DestCountryName'] = df['DestCountry'].map(ISO2_NAME)
    return df


# Find the workbook; allow upload if not present
default_path = Path(__file__).parent / "Freight_Landed_Cost_Dataset.xlsx"
if default_path.exists():
    df = load_data(default_path)
else:
    up = st.file_uploader("Upload Freight_Landed_Cost_Dataset.xlsx", type=["xlsx"])
    if up is None:
        st.info("Upload the dataset to begin.")
        st.stop()
    df = load_data(up)


# ----------------------------------------------------------------------------
# Helper aggregations
# ----------------------------------------------------------------------------
def agg_metric(frame, by, metric):
    """Return a dataframe grouped by `by` with a single 'value' column for the metric."""
    if metric == "Total Landed Cost":
        out = frame.groupby(by)['TotalLandedCostUSD'].sum()
    elif metric == "Total Freight":
        out = frame.groupby(by)['FreightCostUSD'].sum()
    elif metric == "Avg Landed Cost / Unit":
        g = frame.groupby(by).agg(c=('TotalLandedCostUSD', 'sum'), u=('Units', 'sum'))
        out = (g['c'] / g['u'])
    elif metric == "On-Time Delivery %":
        out = frame.groupby(by)['OnTimeFlag'].mean() * 100
    elif metric == "Avg Transit Delay (days)":
        out = frame.groupby(by)['TransitDelay'].mean()
    elif metric == "Shipment Count":
        out = frame.groupby(by)['ShipmentID'].nunique()
    else:
        out = frame.groupby(by)['TotalLandedCostUSD'].sum()
    return out.reset_index(name='value')


def insight(html):
    """Render a teal 'what this means for a buyer' panel."""
    st.markdown(f'<div class="insight">{html}</div>', unsafe_allow_html=True)


def money(v):
    return f"${v/1e6:,.1f}M" if abs(v) >= 1e6 else f"${v:,.0f}"


# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
st.markdown("""
<div class="header-card">
  <h1>Freight &amp; Landed Cost Analytics Dashboard</h1>
  <p>Interactive analysis of global freight, landed cost &amp; carrier performance</p>
</div>
""", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# KPIs
# ----------------------------------------------------------------------------
st.markdown('<div class="section-title">Key Performance Indicators</div>', unsafe_allow_html=True)
st.markdown("<hr>", unsafe_allow_html=True)

total_lc = df['TotalLandedCostUSD'].sum()
lc_unit = total_lc / df['Units'].sum()
ontime = df['OnTimeFlag'].mean() * 100
avg_delay = df['TransitDelay'].mean()

k1, k2, k3, k4 = st.columns(4)
for col, label, val in [
    (k1, "Total Landed Cost", f"${total_lc/1e6:,.1f}M"),
    (k2, "Landed Cost / Unit", f"${lc_unit:,.2f}"),
    (k3, "On-Time Delivery", f"{ontime:.1f}%"),
    (k4, "Avg Transit Delay", f"{avg_delay:.1f} days"),
]:
    col.markdown(f'<div class="kpi-label">{label}</div>'
                 f'<div class="kpi-value">{val}</div>', unsafe_allow_html=True)

insight(
    f"Across {df['ShipmentID'].nunique():,} shipments, {money(total_lc)} is on the line in landed "
    f"cost — roughly ${lc_unit:,.2f} for every unit that touches down. {ontime:.1f}% arrive on time; "
    f"the rest drift about {avg_delay:.1f} days past plan. Landed cost per unit is the real buy "
    f"price once freight, duty and handling are folded in — own that number and you own the margin."
)

st.markdown("<hr>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# MAP — Freight metrics by destination country
# ----------------------------------------------------------------------------
st.markdown('<div class="section-title">Freight Metrics by Destination Country</div>', unsafe_allow_html=True)
st.markdown('<div class="section-sub">Where your goods land, and what they cost to get there</div>',
            unsafe_allow_html=True)

map_metric = st.selectbox("Map Metric",
                          ["Avg Landed Cost / Unit", "Total Landed Cost",
                           "On-Time Delivery %", "Avg Transit Delay (days)"], index=0)

mp = agg_metric(df, 'DestCountryName', map_metric).dropna(subset=['DestCountryName'])
scale = TEAL_SCALE if "On-Time" in map_metric else DELAY_SCALE
fig_map = px.choropleth(mp, locations='DestCountryName', locationmode='country names',
                        color='value', color_continuous_scale=scale,
                        labels={'value': map_metric})
fig_map.update_layout(template='plotly_white', height=470, margin=dict(l=0, r=0, t=10, b=0),
                      geo=dict(showframe=False, projection_type='natural earth', bgcolor='rgba(0,0,0,0)'),
                      coloraxis_colorbar=dict(title=map_metric.split('(')[0].strip()))
st.plotly_chart(fig_map, use_container_width=True)

if len(mp) >= 2:
    hi = mp.loc[mp['value'].idxmax()]
    lo = mp.loc[mp['value'].idxmin()]
    if "On-Time" in map_metric:
        fmt = lambda v: f"{v:.1f}%"
        benefit = ("The destinations dragging on-time down are exactly where buffer stock and "
                   "last-minute expediting quietly eat the budget — tighten the SLAs there and the "
                   "savings surface fast.")
    elif "Delay" in map_metric:
        fmt = lambda v: f"{v:.1f} days"
        benefit = ("Every extra day in transit is cash sitting in a container. Trimming the slowest "
                   "destinations hands working capital straight back.")
    elif "Total" in map_metric:
        fmt = money
        benefit = ("Spend clusters in a handful of destinations, and concentration is leverage — "
                   "these are the lanes where a single renegotiation moves the whole number.")
    else:
        fmt = lambda v: f"${v:,.2f}"
        benefit = ("The priciest destinations per unit are where margin leaks out one shipment at a "
                   "time. Line them up against cheaper lanes into the same market before next "
                   "season's volume is committed.")
    insight(
        f"{hi['DestCountryName']} tops the chart at {fmt(hi['value'])} while "
        f"{lo['DestCountryName']} sits lowest at {fmt(lo['value'])}. {benefit}"
    )

st.markdown("<hr>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# ANIMATED CARRIER PERFORMANCE (over months)
# ----------------------------------------------------------------------------
st.markdown('<div class="section-title">Monthly Carrier Performance Animation</div>', unsafe_allow_html=True)
st.markdown('<div class="section-sub">Press play to watch carriers shift month by month</div>',
            unsafe_allow_html=True)

months = (df[['MonthYear', 'MonthSort']].drop_duplicates()
          .sort_values('MonthSort')['MonthYear'].tolist())
all_carriers = (df.groupby('CarrierName')['ShipmentID'].nunique()
                .sort_values(ascending=False).index.tolist())

c1, c2 = st.columns([3, 2])
sel_carriers = c1.multiselect("Select Carriers", all_carriers, default=all_carriers[:12])
anim_metric = c2.selectbox("Select Metric",
                           ["Shipment Count", "Total Landed Cost", "Total Freight",
                            "Avg Transit Delay (days)", "On-Time Delivery %"], index=0)

if sel_carriers:
    sub = df[df['CarrierName'].isin(sel_carriers)]
    g = agg_metric(sub, ['MonthYear', 'CarrierName'], anim_metric)
    # complete grid so every carrier appears in every frame (stable animation)
    full = pd.MultiIndex.from_product([months, sel_carriers], names=['MonthYear', 'CarrierName'])
    g = g.set_index(['MonthYear', 'CarrierName']).reindex(full, fill_value=0).reset_index()
    order = (g.groupby('CarrierName')['value'].sum().sort_values().index.tolist())  # asc -> largest on top
    xmax = g['value'].max() * 1.1 or 1

    fig_bar = px.bar(g, x='value', y='CarrierName', color='CarrierName',
                     orientation='h', animation_frame='MonthYear',
                     category_orders={'MonthYear': months, 'CarrierName': order},
                     color_discrete_sequence=VIVID, range_x=[0, xmax],
                     labels={'value': anim_metric, 'CarrierName': ''})
    fig_bar.update_layout(template='plotly_white', height=520, showlegend=False,
                          margin=dict(l=10, r=10, t=10, b=10))
    try:  # slow the animation a touch
        fig_bar.layout.updatemenus[0].buttons[0].args[1]['frame']['duration'] = 650
        fig_bar.layout.updatemenus[0].buttons[0].args[1]['transition']['duration'] = 300
    except Exception:
        pass
    st.plotly_chart(fig_bar, use_container_width=True)

    # dynamic insight: volume leader + on-time spread among selected carriers
    tot = agg_metric(sub, 'CarrierName', anim_metric)
    leader = tot.loc[tot['value'].idxmax()]
    ot_by = sub.groupby('CarrierName')['OnTimeFlag'].mean() * 100
    best_ot, worst_ot = ot_by.idxmax(), ot_by.idxmin()
    lead_val = (money(leader['value']) if "Cost" in anim_metric or "Freight" in anim_metric
                else (f"{leader['value']:.1f}%" if "%" in anim_metric
                      else (f"{leader['value']:.1f} days" if "Delay" in anim_metric
                            else f"{leader['value']:,.0f}")))
    insight(
        f"{leader['CarrierName']} carries the window on {anim_metric} ({lead_val}). Reliability "
        f"splits wider: {best_ot} lands {ot_by.max():.1f}% on time while {worst_ot} manages only "
        f"{ot_by.min():.1f}% — a {ot_by.max()-ot_by.min():.1f}-point spread that is pure negotiating "
        f"room. Shift volume to the carrier that delivers, or put the laggard on a contractual "
        f"on-time target and hold them to it."
    )

st.markdown('<div class="tipbox"><b>Tip:</b> Click the play button to see how carriers '
            'compare across the 30-month window.</div>', unsafe_allow_html=True)

st.markdown("<hr>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# SUNBURST — Freight volume hierarchy
# ----------------------------------------------------------------------------
st.markdown('<div class="section-title">Freight Volume Hierarchy</div>', unsafe_allow_html=True)
st.markdown('<div class="section-sub">Drill down: Region &rarr; Country &rarr; Port &rarr; Carrier</div>',
            unsafe_allow_html=True)

regions = ["All Regions"] + sorted(df['DestRegion'].dropna().unique().tolist())
sel_region = st.selectbox("Filter by Destination Region", regions, index=0)

sb = df if sel_region == "All Regions" else df[df['DestRegion'] == sel_region]
path_cols = ['DestRegion', 'DestCountryName', 'DestPort', 'CarrierName']
sbg = (sb.groupby(path_cols)
       .agg(Shipments=('ShipmentID', 'nunique'), AvgDelay=('TransitDelay', 'mean'))
       .reset_index().dropna())

fig_sb = px.sunburst(sbg, path=path_cols, values='Shipments', color='AvgDelay',
                     color_continuous_scale=DELAY_SCALE,
                     labels={'AvgDelay': 'Avg Transit Delay (days)'})
fig_sb.update_layout(template='plotly_white', height=620, margin=dict(l=0, r=0, t=10, b=0),
                     coloraxis_colorbar=dict(title="Avg Delay<br>(days)"))
st.plotly_chart(fig_sb, use_container_width=True)

if len(sbg) >= 1:
    busiest = sbg.loc[sbg['Shipments'].idxmax()]
    country_delay = sbg.groupby('DestCountryName')['AvgDelay'].mean()
    worst_country = country_delay.idxmax()
    insight(
        f"The busiest single route runs {busiest['DestRegion']} &rarr; {busiest['DestCountryName']} "
        f"&rarr; {busiest['DestPort']} &rarr; {busiest['CarrierName']} "
        f"({int(busiest['Shipments'])} shipments), and {worst_country} carries the worst average "
        f"delay at {country_delay.max():.1f} days. Leaning this hard on one country-port-carrier "
        f"branch is a single point of failure — one strike, one congested port, one rate hike and it "
        f"lands squarely on you. This is precisely where qualifying a second carrier or an alternate "
        f"port pays for itself."
    )

st.markdown('<div class="tipbox"><b>Tip:</b> Click any segment to zoom in and explore the hierarchy.</div>',
            unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# FOOTER
# ----------------------------------------------------------------------------
st.markdown('<div class="foot">Freight &amp; Landed Cost Analytics Dashboard '
            '&bull; Built with Streamlit + Plotly</div>', unsafe_allow_html=True)