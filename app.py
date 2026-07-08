import streamlit as st
import pandas as pd
import plotly.express as px

# --------------------------------------------------
# PAGE SETUP
# --------------------------------------------------
st.set_page_config(
    page_title="HSD HQ Brief",
    layout="wide"
)

# --------------------------------------------------
# HSD STYLE COLORS - BLUE / NAVY ONLY
# --------------------------------------------------
HSD_NAVY = "#1B2A4A"
HSD_BLUE = "#2D6DB5"
HSD_MEDIUM_BLUE = "#4A90D9"
HSD_LIGHT_BLUE = "#EBF4FF"
HSD_SKY_BLUE = "#BFDFFF"
HSD_BG = "#F8FAFC"
HSD_BG2 = "#F1F5F9"
HSD_TEXT = "#1A202C"
HSD_MUTED = "#6B7280"
HSD_BORDER = "#E2E8F0"
HSD_WHITE = "#FFFFFF"

BLUE_SCALE = [HSD_NAVY, HSD_BLUE, HSD_MEDIUM_BLUE, HSD_SKY_BLUE]

# --------------------------------------------------
# CUSTOM CSS
# --------------------------------------------------
st.markdown(
    f"""
    <style>
    .stApp {{
        background-color: {HSD_BG};
        color: {HSD_TEXT};
    }}

    section[data-testid="stSidebar"] {{
        background-color: {HSD_BG2};
        border-right: 1px solid {HSD_BORDER};
    }}

    h1, h2, h3 {{
        color: {HSD_NAVY};
        font-weight: 700;
    }}

    div[data-testid="stMetric"] {{
        background: {HSD_WHITE};
        border: 1px solid {HSD_BORDER};
        border-left: 6px solid {HSD_BLUE};
        border-radius: 14px;
        padding: 18px;
        box-shadow: 0px 2px 8px rgba(27, 42, 74, 0.08);
    }}

    div[data-testid="stMetricLabel"] {{
        color: {HSD_MUTED};
        font-size: 13px;
    }}

    div[data-testid="stMetricValue"] {{
        color: {HSD_NAVY};
        font-weight: 700;
    }}

    .hsd-hero {{
        background: linear-gradient(135deg, {HSD_NAVY}, {HSD_BLUE});
        padding: 28px;
        border-radius: 18px;
        margin-bottom: 22px;
        color: white;
    }}

    .hsd-hero h1 {{
        color: white;
        margin-bottom: 8px;
        font-size: 34px;
    }}

    .hsd-hero p {{
        color: rgba(255,255,255,0.88);
        font-size: 16px;
        margin-bottom: 0px;
    }}

    .hsd-card {{
        background: {HSD_WHITE};
        border: 1px solid {HSD_BORDER};
        border-radius: 14px;
        padding: 18px;
        box-shadow: 0px 2px 8px rgba(27, 42, 74, 0.08);
        margin-bottom: 16px;
    }}

    .hsd-blue-box {{
        background: {HSD_LIGHT_BLUE};
        border: 1px solid {HSD_SKY_BLUE};
        border-radius: 14px;
        padding: 18px;
        color: {HSD_NAVY};
        margin-top: 12px;
    }}

    .stTabs [data-baseweb="tab-list"] {{
        gap: 18px;
        border-bottom: 1px solid {HSD_BORDER};
    }}

    .stTabs [data-baseweb="tab"] {{
        color: {HSD_MUTED};
        font-weight: 600;
    }}

    .stTabs [aria-selected="true"] {{
        color: {HSD_BLUE};
        border-bottom: 3px solid {HSD_BLUE};
    }}
    </style>
    """,
    unsafe_allow_html=True
)

# --------------------------------------------------
# HELPER FUNCTIONS
# --------------------------------------------------
def money(value):
    return f"${value:,.0f}"

def apply_hsd_theme(fig):
    fig.update_layout(
        template="plotly_white",
        font=dict(family="Arial", color=HSD_TEXT),
        title_font=dict(color=HSD_NAVY, size=18),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=30, r=30, t=60, b=30),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=-0.25,
            xanchor="center",
            x=0.5
        )
    )
    return fig

# --------------------------------------------------
# SIDEBAR INPUTS - NO INDUSTRY ASSUMPTIONS
# --------------------------------------------------
st.sidebar.title("Build Prospect Profile")
st.sidebar.caption("Enter numbers manually. No industry assumptions are used.")

company = st.sidebar.text_input("Company Name", "New Prospect Company")

industry = st.sidebar.selectbox(
    "Industry",
    [
        "Healthcare",
        "Manufacturing",
        "Logistics",
        "Hospitality",
        "Retail",
        "Education",
        "Financial Services",
        "Other"
    ]
)

region = st.sidebar.selectbox(
    "Region",
    ["Midwest", "Southwest", "Southeast", "Northeast", "West", "National"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("Client Inputs")

headcount = st.sidebar.number_input(
    "Approx. Headcount",
    min_value=0,
    max_value=100000,
    value=0,
    step=100
)

frontline_pct = st.sidebar.slider(
    "Frontline Workforce %",
    min_value=0,
    max_value=100,
    value=0
)

turnover_rate = st.sidebar.slider(
    "Estimated Turnover Rate %",
    min_value=0,
    max_value=100,
    value=0
)

avg_salary = st.sidebar.number_input(
    "Average Salary",
    min_value=0,
    max_value=300000,
    value=0,
    step=5000
)

replacement_multiplier_pct = st.sidebar.slider(
    "Replacement Cost Multiplier %",
    min_value=0,
    max_value=300,
    value=0
)

software_cost = st.sidebar.number_input(
    "Annual Software Cost",
    min_value=0,
    value=0,
    step=5000
)

internal_cost = st.sidebar.number_input(
    "Internal HR Effort Cost",
    min_value=0,
    value=0,
    step=5000
)

external_cost = st.sidebar.number_input(
    "External Support Cost",
    min_value=0,
    value=0,
    step=5000
)

hsd_investment = st.sidebar.number_input(
    "Estimated HSD Investment",
    min_value=0,
    value=0,
    step=5000
)

hsd_improvement_pct = st.sidebar.slider(
    "Estimated HSD Improvement %",
    min_value=0,
    max_value=100,
    value=0
)

st.sidebar.markdown("---")
st.sidebar.subheader("Listening Inputs")

maturity_score = st.sidebar.slider(
    "Listening Maturity Score",
    min_value=0,
    max_value=100,
    value=0
)

expected_maturity_lift = st.sidebar.slider(
    "Expected HSD Maturity Lift",
    min_value=0,
    max_value=100,
    value=0
)

days_to_action = st.sidebar.slider(
    "Days to Leadership Action",
    min_value=0,
    max_value=180,
    value=0
)

expected_days_reduced = st.sidebar.slider(
    "Expected Days Reduced with HSD",
    min_value=0,
    max_value=180,
    value=0
)

retention_plan = st.sidebar.selectbox(
    "Retention Action Plan",
    ["Not entered", "No", "In progress", "Yes"]
)

# --------------------------------------------------
# CALCULATIONS - ALL BASED ON USER INPUTS
# --------------------------------------------------
replacement_multiplier = replacement_multiplier_pct / 100

annual_departures = headcount * turnover_rate / 100
cost_per_resignation = avg_salary * replacement_multiplier
turnover_cost = annual_departures * cost_per_resignation

status_quo_cost = turnover_cost + software_cost + internal_cost + external_cost
hsd_opportunity = status_quo_cost * hsd_improvement_pct / 100
with_hsd_cost = status_quo_cost - hsd_opportunity

roi = hsd_opportunity / hsd_investment if hsd_investment > 0 else 0
monthly_savings = hsd_opportunity / 12
payback_months = hsd_investment / monthly_savings if monthly_savings > 0 else 0

hsd_supported_maturity = min(maturity_score + expected_maturity_lift, 100)
hsd_days_to_action = max(days_to_action - expected_days_reduced, 0)

has_required_inputs = (
    headcount > 0
    and turnover_rate > 0
    and avg_salary > 0
    and replacement_multiplier_pct > 0
)

# --------------------------------------------------
# WORKFORCE DATA
# --------------------------------------------------
frontline_employees = int(headcount * frontline_pct / 100)
non_frontline_employees = headcount - frontline_employees

workforce_data = pd.DataFrame({
    "Workforce Group": ["Frontline", "Non-Frontline"],
    "Employees": [frontline_employees, non_frontline_employees]
})

# --------------------------------------------------
# HEADER
# --------------------------------------------------
st.markdown(
    f"""
    <div class="hsd-hero">
        <h1>HSD HQ Brief</h1>
        <p>
        A simple visual pitch tool showing employee listening gaps, turnover cost,
        cost of status quo, and estimated HSD opportunity using only entered inputs.
        </p>
    </div>
    """,
    unsafe_allow_html=True
)

st.caption("Directional pre-sales estimate. No industry assumptions are used. Results depend only on entered values.")

if not has_required_inputs:
    st.info(
        "Start by entering headcount, turnover rate, average salary, and replacement cost multiplier in the sidebar. "
        "Once entered, the KPI cards and graphs will update automatically."
    )

# --------------------------------------------------
# TABS
# --------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs([
    "Prospect Profile",
    "Cost & Opportunity",
    "Forecast & Scenarios",
    "HQ Brief Summary"
])

# --------------------------------------------------
# TAB 1: PROSPECT PROFILE
# --------------------------------------------------
with tab1:
    st.header("Prospect Profile")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Industry", industry)
    col2.metric("Region", region)
    col3.metric("Headcount", f"{headcount:,}")
    col4.metric("Frontline %", f"{frontline_pct}%")

    live1, live2, live3, live4 = st.columns(4)
    live1.metric("Turnover Rate", f"{turnover_rate}%")
    live2.metric("Annual Departures", f"{annual_departures:,.0f}")
    live3.metric("Cost of Status Quo", money(status_quo_cost))
    live4.metric("HSD Opportunity", money(hsd_opportunity))

    st.subheader("Employee Listening Snapshot")

    s1, s2, s3 = st.columns(3)
    s1.metric("Current Maturity Score", maturity_score)
    s2.metric("Current Days to Action", days_to_action)
    s3.metric("Retention Action Plan", retention_plan)

    profile_left, profile_right = st.columns(2)

    with profile_left:
        listening_data = pd.DataFrame({
            "Metric": ["Maturity Score", "Days to Action"],
            "Current": [maturity_score, days_to_action],
            "With HSD": [hsd_supported_maturity, hsd_days_to_action]
        })

        fig_listening = px.bar(
            listening_data,
            x="Metric",
            y=["Current", "With HSD"],
            barmode="group",
            title="Employee Listening: Current vs With HSD",
            color_discrete_sequence=[HSD_SKY_BLUE, HSD_NAVY]
        )
        fig_listening.update_layout(xaxis_title="", yaxis_title="Value")
        fig_listening = apply_hsd_theme(fig_listening)
        st.plotly_chart(fig_listening, use_container_width=True)

    with profile_right:
        fig_workforce = px.pie(
            workforce_data,
            names="Workforce Group",
            values="Employees",
            title="Workforce Mix",
            color_discrete_sequence=[HSD_NAVY, HSD_SKY_BLUE]
        )
        fig_workforce = apply_hsd_theme(fig_workforce)
        st.plotly_chart(fig_workforce, use_container_width=True)

# --------------------------------------------------
# TAB 2: COST & OPPORTUNITY
# --------------------------------------------------
with tab2:
    st.header("Cost & Opportunity")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Annual Departures", f"{annual_departures:,.0f}")
    c2.metric("Turnover Cost", money(turnover_cost))
    c3.metric("Status Quo Cost", money(status_quo_cost))
    c4.metric("Estimated ROI", f"{roi:.1f}x")

    cost_data = pd.DataFrame({
        "Cost Category": [
            "Turnover Cost",
            "Software Cost",
            "Internal HR Effort",
            "External Support"
        ],
        "Amount": [
            turnover_cost,
            software_cost,
            internal_cost,
            external_cost
        ]
    })

    left, right = st.columns(2)

    with left:
        fig_cost = px.pie(
            cost_data,
            names="Cost Category",
            values="Amount",
            title="Cost of Status Quo Breakdown",
            color_discrete_sequence=BLUE_SCALE
        )
        fig_cost = apply_hsd_theme(fig_cost)
        st.plotly_chart(fig_cost, use_container_width=True)

    with right:
        fig_cost_bar = px.bar(
            cost_data,
            x="Cost Category",
            y="Amount",
            text="Amount",
            title="Estimated Cost Components",
            color="Cost Category",
            color_discrete_sequence=BLUE_SCALE
        )
        fig_cost_bar.update_traces(texttemplate="$%{text:,.0f}", textposition="outside")
        fig_cost_bar.update_layout(xaxis_title="", yaxis_title="Amount", showlegend=False)
        fig_cost_bar = apply_hsd_theme(fig_cost_bar)
        st.plotly_chart(fig_cost_bar, use_container_width=True)

# --------------------------------------------------
# TAB 3: FORECAST & SCENARIOS
# --------------------------------------------------
with tab3:
    st.header("Forecast & Scenarios")

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Status Quo Cost", money(status_quo_cost))
    s2.metric("With HSD Cost", money(with_hsd_cost))
    s3.metric("HSD Opportunity", money(hsd_opportunity))
    s4.metric("Payback Period", f"{payback_months:.1f} months")

    scenario_data = pd.DataFrame({
        "Scenario": [
            "Current Status Quo",
            "With HSD",
            "Estimated Opportunity"
        ],
        "Amount": [
            status_quo_cost,
            with_hsd_cost,
            hsd_opportunity
        ]
    })

    col_left, col_right = st.columns(2)

    with col_left:
        fig_scenario = px.bar(
            scenario_data,
            x="Scenario",
            y="Amount",
            text="Amount",
            title="Current State vs HSD Scenario",
            color="Scenario",
            color_discrete_sequence=[HSD_NAVY, HSD_BLUE, HSD_SKY_BLUE]
        )
        fig_scenario.update_traces(texttemplate="$%{text:,.0f}", textposition="outside")
        fig_scenario.update_layout(xaxis_title="", yaxis_title="Estimated Annual Cost", showlegend=False)
        fig_scenario = apply_hsd_theme(fig_scenario)
        st.plotly_chart(fig_scenario, use_container_width=True)

    with col_right:
        months = list(range(1, 13))
        forecast_data = pd.DataFrame({
            "Month": [f"Month {m}" for m in months],
            "Current Status Quo": [(status_quo_cost / 12) * m for m in months],
            "With HSD": [((with_hsd_cost + hsd_investment) / 12) * m for m in months]
        })

        fig_forecast = px.line(
            forecast_data,
            x="Month",
            y=["Current Status Quo", "With HSD"],
            markers=True,
            title="12-Month Cumulative Cost Forecast",
            color_discrete_sequence=[HSD_NAVY, HSD_BLUE]
        )
        fig_forecast.update_layout(xaxis_title="", yaxis_title="Cumulative Cost")
        fig_forecast = apply_hsd_theme(fig_forecast)
        st.plotly_chart(fig_forecast, use_container_width=True)

    st.markdown(
        f"""
        <div class="hsd-blue-box">
            <b>Estimated annual HSD opportunity:</b> {money(hsd_opportunity)}<br>
            <b>Estimated HSD investment:</b> {money(hsd_investment)}<br>
            <b>Estimated ROI:</b> {roi:.1f}x<br>
            <b>Estimated payback period:</b> {payback_months:.1f} months
        </div>
        """,
        unsafe_allow_html=True
    )

# --------------------------------------------------
# TAB 4: HQ BRIEF SUMMARY
# --------------------------------------------------
with tab4:
    st.header("HQ Brief Summary")

    st.markdown(
        f"""
        <div class="hsd-card">
            <h2>{company}</h2>
            <p><b>Industry:</b> {industry}</p>
            <p><b>Region:</b> {region}</p>
            <p><b>Headcount:</b> {headcount:,}</p>
            <p><b>Frontline Workforce:</b> {frontline_pct}%</p>
            <p><b>Estimated Turnover Rate:</b> {turnover_rate}%</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    col_a, col_b, col_c = st.columns(3)
    col_a.metric("Cost of Status Quo", money(status_quo_cost))
    col_b.metric("HSD Opportunity", money(hsd_opportunity))
    col_c.metric("ROI Multiple", f"{roi:.1f}x")

    st.markdown(
        f"""
        <div class="hsd-blue-box">
            <h3>Recommended Sales Message</h3>
            <p>
            Based on the numbers entered for <b>{company}</b>, the estimated cost of status quo is
            <b>{money(status_quo_cost)}</b>.
            </p>
            <p>
            If HSD helps improve the current state by <b>{hsd_improvement_pct}%</b>,
            the estimated annual opportunity is <b>{money(hsd_opportunity)}</b>.
            </p>
            <p>
            HSD can help by improving employee listening coverage, reducing time to leadership action,
            and connecting feedback to targeted retention planning.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.warning(
        "This dashboard uses only the values entered in the sidebar. Final client-facing numbers should be validated with real HR, turnover, compensation, and employee-listening data."
    )