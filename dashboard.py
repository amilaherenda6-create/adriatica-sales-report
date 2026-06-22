"""Adriatica Sales Dashboard — Sales Director view.
Run: python -m streamlit run dashboard.py
"""
from pathlib import Path
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

DATA = Path(__file__).parent / "data" / "orders.csv"

st.set_page_config(
    page_title="Adriatica Sales Dashboard",
    page_icon="📊",
    layout="wide",
)


@st.cache_data
def load_data() -> pd.DataFrame:
    df = pd.read_csv(DATA, parse_dates=["order_date"])
    df["revenue"] = df["quantity"] * df["unit_price"]
    df["month"] = df["order_date"].dt.to_period("M").astype(str)
    df["month_dt"] = pd.to_datetime(df["month"])
    return df


df_all = load_data()

# ── Sidebar filters ────────────────────────────────────────────────────────────
st.sidebar.title("Filters")

categories = sorted(df_all["category"].unique())
selected_cats = st.sidebar.multiselect("Category", categories, default=categories)

months = sorted(df_all["month"].unique())
from_month = st.sidebar.selectbox("From month", months, index=0)
to_month = st.sidebar.selectbox("To month", months, index=len(months) - 1)
if from_month > to_month:
    st.sidebar.warning("'From' is after 'To' — showing full range.")
    from_month, to_month = months[0], months[-1]

df = df_all[
    df_all["category"].isin(selected_cats)
    & df_all["month"].between(from_month, to_month)
].copy()

# ── KPI row ────────────────────────────────────────────────────────────────────
monthly = df.groupby("month")["revenue"].sum().sort_index()
total_rev = df["revenue"].sum()
best_month = monthly.idxmax() if not monthly.empty else "—"
best_rev = monthly.max() if not monthly.empty else 0
avg_monthly = monthly.mean() if not monthly.empty else 0
unique_customers = df["customer_id"].nunique()
total_orders = df["order_id"].nunique()

mom_series = monthly.pct_change() * 100
last_mom = mom_series.iloc[-1] if len(mom_series) >= 2 else None

st.title("Adriatica Sales Dashboard")
st.caption(f"Data: {from_month} → {to_month}  ·  {len(df):,} orders")

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Total Revenue", f"€{total_rev:,.0f}")
c2.metric(
    "Last MoM Growth",
    f"{last_mom:+.1f}%" if last_mom is not None else "—",
    delta=f"{last_mom:+.1f}%" if last_mom is not None else None,
)
c3.metric("Best Month", best_month, f"€{best_rev:,.0f}")
c4.metric("Avg Monthly Revenue", f"€{avg_monthly:,.0f}")
c5.metric("Unique Customers", f"{unique_customers:,}", f"{total_orders:,} orders")

st.divider()

# ── Monthly revenue + MoM growth ──────────────────────────────────────────────
st.subheader("Monthly Revenue & Month-over-Month Growth")

monthly_df = monthly.reset_index()
monthly_df.columns = ["month", "revenue"]
monthly_df["mom_pct"] = monthly_df["revenue"].pct_change() * 100

fig_trend = make_subplots(specs=[[{"secondary_y": True}]])
fig_trend.add_trace(
    go.Bar(
        x=monthly_df["month"],
        y=monthly_df["revenue"],
        name="Revenue (€)",
        marker_color="#4C78A8",
    ),
    secondary_y=False,
)
fig_trend.add_trace(
    go.Scatter(
        x=monthly_df["month"],
        y=monthly_df["mom_pct"],
        name="MoM Growth %",
        mode="lines+markers",
        line=dict(color="#E45756", width=2),
        marker=dict(size=7),
    ),
    secondary_y=True,
)
fig_trend.update_layout(
    height=380,
    legend=dict(orientation="h", y=1.1),
    hovermode="x unified",
    margin=dict(t=20, b=0),
)
fig_trend.update_yaxes(title_text="Revenue (€)", secondary_y=False, tickformat=",.0f")
fig_trend.update_yaxes(title_text="MoM Growth %", secondary_y=True, ticksuffix="%")
st.plotly_chart(fig_trend, use_container_width=True)

# ── Revenue by Category (bar) ──────────────────────────────────────────────────
st.subheader("Revenue by Category")

cat_total = (
    df.groupby("category")["revenue"].sum().sort_values(ascending=False).reset_index()
)
cat_total.columns = ["category", "revenue"]
fig_cat = px.bar(
    cat_total,
    x="category",
    y="revenue",
    color="category",
    text_auto=".3s",
    labels={"revenue": "Revenue (€)", "category": "Category"},
    height=340,
)
fig_cat.update_layout(showlegend=False, margin=dict(t=10, b=0))
fig_cat.update_yaxes(tickformat=",.0f")
fig_cat.update_traces(textposition="outside")
st.plotly_chart(fig_cat, use_container_width=True)

st.divider()

# ── Top 10 Customers (bar) ─────────────────────────────────────────────────────
st.subheader("Top 10 Customers by Revenue")

top_customers = (
    df.groupby("customer_id")
    .agg(revenue=("revenue", "sum"), orders=("order_id", "nunique"))
    .nlargest(10, "revenue")
    .sort_values("revenue", ascending=False)
    .reset_index()
)
fig_cust = px.bar(
    top_customers,
    x="customer_id",
    y="revenue",
    color="orders",
    text_auto=".3s",
    color_continuous_scale="Greens",
    labels={"revenue": "Revenue (€)", "customer_id": "Customer", "orders": "# Orders"},
    height=340,
    hover_data={"orders": True},
)
fig_cust.update_layout(margin=dict(t=10, b=0))
fig_cust.update_yaxes(tickformat=",.0f")
fig_cust.update_traces(textposition="outside")
st.plotly_chart(fig_cust, use_container_width=True)

st.divider()

# ── Top 10 Products ────────────────────────────────────────────────────────────
st.subheader("Top 10 Products by Revenue")

top_products = (
    df.groupby("product_name")["revenue"]
    .sum()
    .nlargest(10)
    .sort_values(ascending=False)
    .reset_index()
)
top_products.columns = ["product", "revenue"]
fig_prod = px.bar(
    top_products,
    x="product",
    y="revenue",
    color="revenue",
    text_auto=".3s",
    color_continuous_scale="Blues",
    labels={"revenue": "Revenue (€)", "product": "Product"},
    height=380,
)
fig_prod.update_layout(showlegend=False, coloraxis_showscale=False, margin=dict(t=10, b=0))
fig_prod.update_yaxes(tickformat=",.0f")
fig_prod.update_traces(textposition="outside")
fig_prod.update_xaxes(tickangle=-30)
st.plotly_chart(fig_prod, use_container_width=True)

# ── Heatmap (collapsed) ────────────────────────────────────────────────────────
with st.expander("Revenue heatmap — Month × Category"):
    cat_month = df.groupby(["month", "category"])["revenue"].sum().reset_index()
    fig_heat = px.density_heatmap(
        cat_month,
        x="month",
        y="category",
        z="revenue",
        color_continuous_scale="Blues",
        labels={"revenue": "Revenue (€)", "month": "Month", "category": "Category"},
        height=300,
    )
    fig_heat.update_layout(margin=dict(t=10, b=0))
    fig_heat.update_coloraxes(colorbar_tickformat=",.0f")
    st.plotly_chart(fig_heat, use_container_width=True)

# ── Data table (collapsed) ─────────────────────────────────────────────────────
with st.expander("Raw monthly summary table"):
    summary = (
        df.groupby("month")
        .agg(
            revenue=("revenue", "sum"),
            orders=("order_id", "nunique"),
            customers=("customer_id", "nunique"),
        )
        .reset_index()
    )
    summary["mom_growth"] = summary["revenue"].pct_change().map(
        lambda x: f"{x:+.1%}" if pd.notna(x) else "—"
    )
    summary["revenue"] = summary["revenue"].map(lambda x: f"€{x:,.2f}")
    st.dataframe(summary, use_container_width=True, hide_index=True)
