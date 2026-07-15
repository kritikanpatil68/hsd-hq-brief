from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
import base64
import html
import re

import pandas as pd
import plotly.express as px
import streamlit as st
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

# --------------------------------------------------
# PAGE SETUP
# --------------------------------------------------
HSD_WEBSITE = "https://hsdmetrics.com"
HSD_LOGO_URL = "https://mma.prnewswire.com/media/2925136/HSD_Logo_Dark_Blue_Logo.jpg"
HSD_LOGO_PATH = Path(__file__).with_name("hsd_logo.jpg")


def get_logo_src() -> str:
    """Use the local official logo when available, with a remote fallback."""
    if HSD_LOGO_PATH.exists():
        encoded = base64.b64encode(HSD_LOGO_PATH.read_bytes()).decode("ascii")
        return f"data:image/jpeg;base64,{encoded}"
    return HSD_LOGO_URL


HSD_LOGO_SRC = get_logo_src()

st.set_page_config(
    page_title="HSD HQ Brief",
    page_icon="📊",
    layout="wide",
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
        box-shadow: 0 2px 8px rgba(27, 42, 74, 0.08);
    }}

    div[data-testid="stMetricLabel"] {{
        color: {HSD_MUTED};
        font-size: 13px;
    }}

    div[data-testid="stMetricValue"] {{
        color: {HSD_NAVY};
        font-weight: 700;
    }}

    .hsd-header {{
        display: flex;
        align-items: center;
        gap: 26px;
        background: linear-gradient(135deg, {HSD_NAVY}, {HSD_BLUE});
        padding: 24px 28px;
        border-radius: 18px;
        margin-bottom: 18px;
        color: white;
    }}

    .hsd-logo-wrap {{
        background: white;
        padding: 10px 14px;
        border-radius: 12px;
        min-width: 175px;
        text-align: center;
    }}

    .hsd-logo-wrap img {{
        width: 155px;
        max-width: 100%;
        display: block;
    }}

    .hsd-header h1 {{
        color: white;
        margin: 0 0 6px 0;
        font-size: 34px;
    }}

    .hsd-header p {{
        color: rgba(255, 255, 255, 0.90);
        font-size: 16px;
        margin: 0;
    }}

    .hsd-card {{
        background: {HSD_WHITE};
        border: 1px solid {HSD_BORDER};
        border-radius: 14px;
        padding: 18px;
        box-shadow: 0 2px 8px rgba(27, 42, 74, 0.08);
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

    .hsd-note {{
        background: {HSD_WHITE};
        border-left: 5px solid {HSD_MEDIUM_BLUE};
        border-radius: 10px;
        padding: 12px 15px;
        margin: 10px 0 16px 0;
        color: {HSD_TEXT};
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

    @media (max-width: 800px) {{
        .hsd-header {{
            display: block;
        }}
        .hsd-logo-wrap {{
            width: fit-content;
            margin-bottom: 16px;
        }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# --------------------------------------------------
# HELPER FUNCTIONS
# --------------------------------------------------
def money(value: float) -> str:
    return f"${value:,.0f}"


def percent(value: float) -> str:
    return f"{value:,.1f}%"


def multiple(value: float | None) -> str:
    return "N/A" if value is None else f"{value:,.2f}x"


def months(value: float | None) -> str:
    return "N/A" if value is None else f"{value:,.1f} months"


def roi_text(value: float | None) -> str:
    return "N/A" if value is None else f"{value:,.0f}%"


def money_range(low: float, high: float) -> str:
    return money(low) if abs(low - high) < 0.01 else f"{money(low)} - {money(high)}"


def roi_range(low: float | None, high: float | None) -> str:
    if low is None or high is None:
        return "N/A"
    return roi_text(low) if abs(low - high) < 0.01 else f"{roi_text(low)} - {roi_text(high)}"


def multiple_range(low: float | None, high: float | None) -> str:
    if low is None or high is None:
        return "N/A"
    return multiple(low) if abs(low - high) < 0.001 else f"{multiple(low)} - {multiple(high)}"


def months_range(first: float | None, second: float | None) -> str:
    if first is None or second is None:
        return "N/A"
    low, high = sorted([first, second])
    return months(low) if abs(low - high) < 0.01 else f"{months(low)} - {months(high)}"


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
            x=0.5,
        ),
    )
    return fig


def safe_filename(text: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9 _-]+", "", text).strip()
    return cleaned or "Prospect"


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_text(cell, text: str, *, bold: bool = False, color: str | None = None) -> None:
    cell.text = ""
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.name = "Arial"
    run.font.size = Pt(9)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def add_doc_heading(doc: Document, text: str, size: int = 15) -> None:
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(4)
    paragraph.paragraph_format.space_after = Pt(5)
    run = paragraph.add_run(text)
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(27, 42, 74)


def add_doc_bullet(doc: Document, text: str) -> None:
    paragraph = doc.add_paragraph(style="List Bullet")
    paragraph.paragraph_format.space_after = Pt(2)
    run = paragraph.add_run(text)
    run.font.name = "Arial"
    run.font.size = Pt(9)


def create_hq_brief_docx(
    *,
    company: str,
    industry: str,
    turnover_rate: float,
    annual_departures: int,
    cost_per_departure: float,
    current_listening_spend: float,
    hsd_investment: float,
    maturity_score: float,
    retention_plan: str,
    improvement_low: float,
    improvement_high: float,
    status_quo_exposure: float,
    annual_benefit_low: float,
    annual_benefit_base: float,
    annual_benefit_high: float,
    net_benefit_low: float,
    net_benefit_base: float,
    net_benefit_high: float,
    roi_low: float | None,
    roi_base: float | None,
    roi_high: float | None,
    benefit_cost_low: float | None,
    benefit_cost_base: float | None,
    benefit_cost_high: float | None,
    payback_low: float | None,
    payback_base: float | None,
    payback_high: float | None,
) -> bytes:
    """Create a concise, company-specific two-page Word brief."""
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.55)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(9)
    styles["Normal"].paragraph_format.space_after = Pt(3)

    # PAGE 1
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run(f"{company} Employee Listening Enhancement Brief")
    title_run.bold = True
    title_run.font.name = "Arial"
    title_run.font.size = Pt(20)
    title_run.font.color.rgb = RGBColor(27, 42, 74)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle_run = subtitle.add_run("Prepared using the HSD HQ Brief directional pre-sales model")
    subtitle_run.italic = True
    subtitle_run.font.name = "Arial"
    subtitle_run.font.size = Pt(9)
    subtitle_run.font.color.rgb = RGBColor(107, 114, 128)

    date_p = doc.add_paragraph()
    date_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_run = date_p.add_run(date.today().strftime("%B %d, %Y"))
    date_run.font.name = "Arial"
    date_run.font.size = Pt(8)
    date_run.font.color.rgb = RGBColor(107, 114, 128)

    add_doc_heading(doc, "Executive overview")
    overview = doc.add_paragraph()
    overview.paragraph_format.space_after = Pt(5)
    overview_run = overview.add_run(
        f"{company} is evaluating how a more fully supported employee listening program could improve "
        "feedback reach, insight quality, and retention action. The financial estimate below uses only "
        "the prospect information entered in the dashboard and presents a low-to-high scenario rather "
        "than a single guaranteed result."
    )
    overview_run.font.name = "Arial"
    overview_run.font.size = Pt(9)

    add_doc_heading(doc, "Prospect profile")
    profile_table = doc.add_table(rows=4, cols=2)
    profile_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    profile_table.style = "Table Grid"
    profile_rows = [
        ("Industry", industry or "Not entered"),
        ("Estimated turnover rate", percent(turnover_rate)),
        ("Annual employee departures", f"{annual_departures:,}"),
        ("Listening maturity / retention plan", f"{maturity_score:.0f}/100 / {retention_plan}"),
    ]
    for row, (label, value) in zip(profile_table.rows, profile_rows):
        set_cell_text(row.cells[0], label, bold=True, color="1B2A4A")
        shade_cell(row.cells[0], "EBF4FF")
        set_cell_text(row.cells[1], value)

    add_doc_heading(doc, "Current annual cost exposure")
    cost_table = doc.add_table(rows=5, cols=2)
    cost_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cost_table.style = "Table Grid"
    cost_rows = [
        ("Annual employee departures", f"{annual_departures:,}"),
        ("Average cost per departure", money(cost_per_departure)),
        ("Estimated annual turnover cost", money(annual_departures * cost_per_departure)),
        ("Current listening spend expected to be replaced/reduced", money(current_listening_spend)),
        ("Total status-quo exposure", money(status_quo_exposure)),
    ]
    for row, (label, value) in zip(cost_table.rows, cost_rows):
        set_cell_text(row.cells[0], label, bold=True, color="1B2A4A")
        shade_cell(row.cells[0], "EBF4FF")
        set_cell_text(row.cells[1], value, bold=label == "Total status-quo exposure")

    add_doc_heading(doc, "Potential HSD business case")
    business_case = doc.add_table(rows=5, cols=2)
    business_case.alignment = WD_TABLE_ALIGNMENT.CENTER
    business_case.style = "Table Grid"
    case_rows = [
        ("HSD improvement range", f"{improvement_low:.1f}% - {improvement_high:.1f}%"),
        ("Estimated annual benefit range", money_range(annual_benefit_low, annual_benefit_high)),
        ("Estimated annual HSD investment", money(hsd_investment)),
        ("Estimated net benefit range", money_range(net_benefit_low, net_benefit_high)),
        ("Estimated net ROI range", roi_range(roi_low, roi_high)),
    ]
    for row, (label, value) in zip(business_case.rows, case_rows):
        set_cell_text(row.cells[0], label, bold=True, color="1B2A4A")
        shade_cell(row.cells[0], "EBF4FF")
        set_cell_text(row.cells[1], value, bold="range" in label.lower())

    key = doc.add_paragraph()
    key.paragraph_format.space_before = Pt(6)
    key.paragraph_format.space_after = Pt(2)
    key_run = key.add_run(
        f"Sales message: The entered assumptions indicate a possible annual benefit of "
        f"{money_range(annual_benefit_low, annual_benefit_high)}. HSD can help {company} move from "
        "disconnected listening activity to a supported program that reaches employees, interprets "
        "feedback, and helps leadership act."
    )
    key_run.bold = True
    key_run.font.name = "Arial"
    key_run.font.size = Pt(9)
    key_run.font.color.rgb = RGBColor(27, 42, 74)

    # PAGE 2
    doc.add_page_break()
    add_doc_heading(doc, "Scenario comparison", size=16)
    scenario_table = doc.add_table(rows=4, cols=7)
    scenario_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    scenario_table.style = "Table Grid"
    headers = [
        "Scenario",
        "Improvement",
        "Annual benefit",
        "Net benefit",
        "Net ROI",
        "Benefit-cost",
        "Payback",
    ]
    for i, header in enumerate(headers):
        set_cell_text(scenario_table.rows[0].cells[i], header, bold=True, color="FFFFFF")
        shade_cell(scenario_table.rows[0].cells[i], "1B2A4A")

    improvement_base = (improvement_low + improvement_high) / 2
    scenarios = [
        ("Low", improvement_low, annual_benefit_low, net_benefit_low, roi_low, benefit_cost_low, payback_low),
        ("Midpoint", improvement_base, annual_benefit_base, net_benefit_base, roi_base, benefit_cost_base, payback_base),
        ("High", improvement_high, annual_benefit_high, net_benefit_high, roi_high, benefit_cost_high, payback_high),
    ]
    for row, values in zip(scenario_table.rows[1:], scenarios):
        display_values = [
            values[0],
            percent(values[1]),
            money(values[2]),
            money(values[3]),
            roi_text(values[4]),
            multiple(values[5]),
            months(values[6]),
        ]
        for i, value in enumerate(display_values):
            set_cell_text(row.cells[i], value, bold=i == 0)
            if i == 0:
                shade_cell(row.cells[i], "EBF4FF")

    add_doc_heading(doc, "Recommended HSD approach")
    add_doc_bullet(doc, "Reach: use multiple outreach methods to include employees who are often missed by standard email-based surveys.")
    add_doc_bullet(doc, "Listen: gather feedback across key moments of the employee lifecycle, not only through one annual survey.")
    add_doc_bullet(doc, "Analyze: connect feedback themes to turnover and other business outcomes using supported analysis and reporting.")
    add_doc_bullet(doc, "Act: provide leaders with clear priorities, ownership, and follow-through rather than only another dashboard.")

    add_doc_heading(doc, "Calculation methodology")
    methodology = [
        "Annual turnover cost = annual employee departures x average cost per departure.",
        "Status-quo exposure = annual turnover cost + current listening costs entered as replaceable or reducible.",
        "Annual HSD benefit = avoided turnover cost under the selected improvement scenario + replaceable/reducible listening spend.",
        "Net benefit = annual HSD benefit - annual HSD investment.",
        "Net ROI = net benefit / HSD investment. Benefit-cost ratio = annual HSD benefit / HSD investment.",
        "Payback period = HSD investment / estimated monthly benefit.",
    ]
    for item in methodology:
        add_doc_bullet(doc, item)

    add_doc_heading(doc, "Important interpretation")
    disclaimer = doc.add_paragraph()
    disclaimer_run = disclaimer.add_run(
        "This is a directional pre-sales estimate, not a guarantee or audited financial forecast. Final client-facing "
        "numbers should be validated with HR, finance, procurement, and HSD subject-matter experts."
    )
    disclaimer_run.font.name = "Arial"
    disclaimer_run.font.size = Pt(8.5)
    disclaimer_run.font.color.rgb = RGBColor(75, 85, 99)

    footer = doc.add_paragraph()
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.space_before = Pt(7)
    footer_run = footer.add_run(f"HSD Metrics | {HSD_WEBSITE}")
    footer_run.bold = True
    footer_run.font.name = "Arial"
    footer_run.font.size = Pt(9)
    footer_run.font.color.rgb = RGBColor(45, 109, 181)

    output = BytesIO()
    doc.save(output)
    return output.getvalue()


# --------------------------------------------------
# SIDEBAR INPUTS - MANUAL, NO INDUSTRY ASSUMPTIONS
# --------------------------------------------------
st.sidebar.title("Build Prospect Profile")
st.sidebar.caption("Type the prospect's known values. No industry assumptions are inserted automatically.")

company = st.sidebar.text_input("Company Name", "New Prospect Company")
industry = st.sidebar.text_input("Industry", placeholder="e.g., Healthcare")
st.sidebar.markdown("---")
st.sidebar.subheader("Turnover & Cost Inputs")

turnover_rate = st.sidebar.number_input(
    "Estimated Turnover Rate %",
    min_value=0.0,
    max_value=100.0,
    value=0.0,
    step=0.5,
    help="This is displayed for context. Annual departures are entered separately because headcount was removed.",
)

annual_departures = st.sidebar.number_input(
    "Annual Employee Departures",
    min_value=0,
    value=0,
    step=1,
    help="Enter the number of employees who leave in a typical year.",
)

cost_per_departure = st.sidebar.number_input(
    "Average Cost per Employee Departure",
    min_value=0.0,
    value=0.0,
    step=1000.0,
    help="Enter a direct dollar estimate. This replaces average salary and the replacement-cost multiplier.",
)

st.sidebar.caption("Enter only the portion of current listening spend that HSD is expected to replace or reduce.")
software_cost = st.sidebar.number_input(
    "Replaceable Annual Software Cost",
    min_value=0.0,
    value=0.0,
    step=5000.0,
)
internal_cost = st.sidebar.number_input(
    "Reducible Internal HR Effort Cost",
    min_value=0.0,
    value=0.0,
    step=5000.0,
)
external_cost = st.sidebar.number_input(
    "Replaceable External Support Cost",
    min_value=0.0,
    value=0.0,
    step=5000.0,
)
hsd_investment = st.sidebar.number_input(
    "Estimated Annual HSD Investment",
    min_value=0.0,
    value=0.0,
    step=5000.0,
)

st.sidebar.markdown("---")
st.sidebar.subheader("HSD Improvement Range")
st.sidebar.caption("Use temporary low and high values until the approved HSD ranges are provided.")
raw_improvement_low = st.sidebar.number_input(
    "Low Improvement %",
    min_value=0.0,
    max_value=100.0,
    value=0.0,
    step=0.5,
)
raw_improvement_high = st.sidebar.number_input(
    "High Improvement %",
    min_value=0.0,
    max_value=100.0,
    value=0.0,
    step=0.5,
)

improvement_low = min(raw_improvement_low, raw_improvement_high)
improvement_high = max(raw_improvement_low, raw_improvement_high)
if raw_improvement_high < raw_improvement_low:
    st.sidebar.warning("The app has reordered the improvement values from low to high.")

st.sidebar.markdown("---")
st.sidebar.subheader("Listening Inputs")
maturity_score = st.sidebar.number_input(
    "Current Listening Maturity Score",
    min_value=0.0,
    max_value=100.0,
    value=0.0,
    step=1.0,
    help="Use only if the score comes from an agreed assessment or questionnaire.",
)
retention_plan = st.sidebar.selectbox(
    "Retention Action Plan",
    ["Not entered", "No", "In progress", "Yes"],
)

# --------------------------------------------------
# CALCULATIONS - TRANSPARENT AND RANGE-BASED
# --------------------------------------------------
annual_turnover_cost = annual_departures * cost_per_departure
current_listening_spend = software_cost + internal_cost + external_cost
status_quo_exposure = annual_turnover_cost + current_listening_spend

improvement_base = (improvement_low + improvement_high) / 2


def scenario_values(improvement_pct: float) -> dict[str, float | None]:
    avoided_turnover_cost = annual_turnover_cost * improvement_pct / 100
    annual_benefit = avoided_turnover_cost + current_listening_spend
    net_benefit = annual_benefit - hsd_investment
    annual_cost_with_hsd = status_quo_exposure - annual_benefit + hsd_investment

    if hsd_investment > 0:
        net_roi_pct = (net_benefit / hsd_investment) * 100
        benefit_cost_ratio = annual_benefit / hsd_investment
    else:
        net_roi_pct = None
        benefit_cost_ratio = None

    payback = hsd_investment / (annual_benefit / 12) if annual_benefit > 0 and hsd_investment > 0 else None

    return {
        "improvement_pct": improvement_pct,
        "avoided_turnover_cost": avoided_turnover_cost,
        "annual_benefit": annual_benefit,
        "net_benefit": net_benefit,
        "annual_cost_with_hsd": max(annual_cost_with_hsd, 0),
        "net_roi_pct": net_roi_pct,
        "benefit_cost_ratio": benefit_cost_ratio,
        "payback_months": payback,
    }


low = scenario_values(improvement_low)
base = scenario_values(improvement_base)
high = scenario_values(improvement_high)

has_cost_inputs = annual_departures > 0 and cost_per_departure > 0
has_improvement_range = improvement_high > 0

# Escape user-entered text before placing it inside HTML blocks.
company_html = html.escape(company or "New Prospect Company")
industry_html = html.escape(industry or "Not entered")

# --------------------------------------------------
# HEADER
# --------------------------------------------------
st.markdown(
    f"""
    <div class="hsd-header">
        <div class="hsd-logo-wrap">
            <a href="{HSD_WEBSITE}" target="_blank" rel="noopener noreferrer">
                <img src="{HSD_LOGO_SRC}" alt="HSD Metrics logo" title="Open HSD Metrics website">
            </a>
        </div>
        <div>
            <h1>HSD HQ Brief</h1>
            <p>A directional pre-sales tool for showing current cost exposure and a low-to-high HSD opportunity range.</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.caption(
    "Click the HSD logo to open the HSD Metrics website. Results use only the values entered in the sidebar."
)

if not has_cost_inputs:
    st.info(
        "Start by entering Annual Employee Departures and Average Cost per Employee Departure. "
        "These two manual inputs replace headcount, average salary, and the replacement-cost multiplier."
    )

if not has_improvement_range:
    st.info("Enter the low and high HSD improvement percentages to create the scenario range.")

# --------------------------------------------------
# TABS
# --------------------------------------------------
tab1, tab2, tab3, tab4 = st.tabs(
    [
        "Prospect Profile",
        "Cost & Opportunity",
        "Forecast & Scenarios",
        "HQ Brief Summary",
    ]
)

# --------------------------------------------------
# TAB 1: PROSPECT PROFILE
# --------------------------------------------------
with tab1:
    st.header("Prospect Profile")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Company", company or "Not entered")
    col2.metric("Industry", industry or "Not entered")
    col3.metric("Turnover Rate", percent(turnover_rate))
    col4.metric("Annual Departures", f"{annual_departures:,}")

    listen1, listen2 = st.columns(2)
    listen1.metric("Current Listening Maturity", f"{maturity_score:.0f}/100")
    listen2.metric("Retention Action Plan", retention_plan)

    turnover_chart_data = pd.DataFrame(
        {
            "Cost Category": ["Annual Turnover Cost"],
            "Amount": [annual_turnover_cost],
        }
    )

    listening_chart_data = pd.DataFrame(
        {
            "Cost Category": [
                "Replaceable Software",
                "Reducible HR Effort",
                "Replaceable External Support",
            ],
            "Amount": [software_cost, internal_cost, external_cost],
        }
    )

    if annual_turnover_cost > 0 or listening_chart_data["Amount"].sum() > 0:
        st.subheader("Current Annual Cost Inputs")

        turnover_col, listening_col = st.columns([1, 2])

        with turnover_col:
            if annual_turnover_cost > 0:
                fig_turnover = px.bar(
                    turnover_chart_data,
                    x="Cost Category",
                    y="Amount",
                    text="Amount",
                    title="Annual Turnover Cost",
                    color_discrete_sequence=[HSD_NAVY],
                )
                fig_turnover.update_traces(
                    texttemplate="$%{text:,.0f}",
                    textposition="outside",
                )
                fig_turnover.update_layout(
                    xaxis_title="",
                    yaxis_title="Annual Amount",
                    showlegend=False,
                    yaxis=dict(range=[0, annual_turnover_cost * 1.18]),
                )
                fig_turnover = apply_hsd_theme(fig_turnover)
                st.plotly_chart(fig_turnover, use_container_width=True)
            else:
                st.info("Enter turnover inputs to display turnover cost.")

        with listening_col:
            if listening_chart_data["Amount"].sum() > 0:
                max_listening_cost = max(listening_chart_data["Amount"].max(), 1)
                fig_listening_costs = px.bar(
                    listening_chart_data,
                    x="Amount",
                    y="Cost Category",
                    orientation="h",
                    text="Amount",
                    title="Current Listening-Related Costs",
                    color="Cost Category",
                    color_discrete_sequence=[
                        HSD_BLUE,
                        HSD_MEDIUM_BLUE,
                        HSD_SKY_BLUE,
                    ],
                )
                fig_listening_costs.update_traces(
                    texttemplate="$%{text:,.0f}",
                    textposition="outside",
                )
                fig_listening_costs.update_layout(
                    xaxis_title="Annual Amount",
                    yaxis_title="",
                    showlegend=False,
                    xaxis=dict(range=[0, max_listening_cost * 1.28]),
                )
                fig_listening_costs = apply_hsd_theme(fig_listening_costs)
                st.plotly_chart(fig_listening_costs, use_container_width=True)
            else:
                st.info("Enter listening-related costs to display this comparison.")

        st.caption(
            "The charts use separate scales because turnover cost is much larger than the listening-related costs. "
            "The dollar labels show the exact values."
        )
    else:
        st.info("Enter cost inputs to display the current annual cost charts.")

# --------------------------------------------------
# TAB 2: COST & OPPORTUNITY
# --------------------------------------------------
with tab2:
    st.header("Cost & Opportunity")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Annual Turnover Cost", money(annual_turnover_cost))
    c2.metric("Replaceable / Reducible Spend", money(current_listening_spend))
    c3.metric("Status-Quo Exposure", money(status_quo_exposure))
    c4.metric("Annual Benefit Range", money_range(low["annual_benefit"], high["annual_benefit"]))

    left, right = st.columns(2)

    with left:
        cost_data = pd.DataFrame(
            {
                "Cost Category": ["Turnover Cost", "Current Listening Spend"],
                "Amount": [annual_turnover_cost, current_listening_spend],
            }
        )
        if cost_data["Amount"].sum() > 0:
            fig_cost = px.pie(
                cost_data,
                names="Cost Category",
                values="Amount",
                title="Status-Quo Exposure Breakdown",
                color_discrete_sequence=[HSD_NAVY, HSD_SKY_BLUE],
            )
            fig_cost = apply_hsd_theme(fig_cost)
            st.plotly_chart(fig_cost, use_container_width=True)
        else:
            st.info("Enter turnover and listening-cost inputs to display this chart.")

    with right:
        benefit_data = pd.DataFrame(
            {
                "Scenario": ["Low", "Midpoint", "High"],
                "Annual Benefit": [low["annual_benefit"], base["annual_benefit"], high["annual_benefit"]],
            }
        )
        fig_benefit = px.bar(
            benefit_data,
            x="Scenario",
            y="Annual Benefit",
            text="Annual Benefit",
            title="Estimated Annual Benefit Range",
            color="Scenario",
            color_discrete_sequence=[HSD_SKY_BLUE, HSD_MEDIUM_BLUE, HSD_NAVY],
        )
        fig_benefit.update_traces(texttemplate="$%{text:,.0f}", textposition="outside")
        fig_benefit.update_layout(xaxis_title="", yaxis_title="Annual Benefit", showlegend=False)
        fig_benefit = apply_hsd_theme(fig_benefit)
        st.plotly_chart(fig_benefit, use_container_width=True)

    with st.expander("See the calculation formulas"):
        st.markdown(
            """
            **Annual turnover cost** = Annual employee departures × Average cost per employee departure  
            **Status-quo exposure** = Annual turnover cost + Current listening spend expected to be replaced/reduced  
            **Annual HSD benefit** = Avoided turnover cost under the selected improvement scenario + Replaceable/reducible listening spend  
            **Net benefit** = Annual HSD benefit − Annual HSD investment  
            **Net ROI %** = Net benefit ÷ HSD investment × 100  
            **Benefit-cost ratio** = Annual HSD benefit ÷ HSD investment  
            **Payback months** = HSD investment ÷ Estimated monthly benefit
            """
        )

# --------------------------------------------------
# TAB 3: FORECAST & SCENARIOS
# --------------------------------------------------
with tab3:
    st.header("Forecast & Scenarios")

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Improvement Range", f"{improvement_low:.1f}% - {improvement_high:.1f}%")
    s2.metric("Net Benefit Range", money_range(low["net_benefit"], high["net_benefit"]))
    s3.metric("Net ROI Range", roi_range(low["net_roi_pct"], high["net_roi_pct"]))
    s4.metric("Payback Range", months_range(low["payback_months"], high["payback_months"]))

    scenario_df = pd.DataFrame(
        {
            "Scenario": ["Low", "Midpoint", "High"],
            "Improvement %": [improvement_low, improvement_base, improvement_high],
            "Annual Benefit": [low["annual_benefit"], base["annual_benefit"], high["annual_benefit"]],
            "Net Benefit": [low["net_benefit"], base["net_benefit"], high["net_benefit"]],
            "Net ROI %": [low["net_roi_pct"], base["net_roi_pct"], high["net_roi_pct"]],
            "Benefit-Cost Ratio": [
                low["benefit_cost_ratio"],
                base["benefit_cost_ratio"],
                high["benefit_cost_ratio"],
            ],
            "Payback Months": [low["payback_months"], base["payback_months"], high["payback_months"]],
            "Annual Cost With HSD": [
                low["annual_cost_with_hsd"],
                base["annual_cost_with_hsd"],
                high["annual_cost_with_hsd"],
            ],
        }
    )

    display_df = scenario_df.copy()
    display_df["Improvement %"] = display_df["Improvement %"].map(lambda x: f"{x:.1f}%")
    display_df["Annual Benefit"] = display_df["Annual Benefit"].map(money)
    display_df["Net Benefit"] = display_df["Net Benefit"].map(money)
    display_df["Net ROI %"] = display_df["Net ROI %"].map(roi_text)
    display_df["Benefit-Cost Ratio"] = display_df["Benefit-Cost Ratio"].map(multiple)
    display_df["Payback Months"] = display_df["Payback Months"].map(months)
    display_df["Annual Cost With HSD"] = display_df["Annual Cost With HSD"].map(money)
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    col_left, col_right = st.columns(2)

    with col_left:
        scenario_cost_data = pd.DataFrame(
            {
                "Scenario": ["Current Status Quo", "HSD Low", "HSD Midpoint", "HSD High"],
                "Annual Cost": [
                    status_quo_exposure,
                    low["annual_cost_with_hsd"],
                    base["annual_cost_with_hsd"],
                    high["annual_cost_with_hsd"],
                ],
            }
        )
        fig_scenario = px.bar(
            scenario_cost_data,
            x="Scenario",
            y="Annual Cost",
            text="Annual Cost",
            title="Current Annual Cost vs HSD Scenarios",
            color="Scenario",
            color_discrete_sequence=[HSD_NAVY, HSD_SKY_BLUE, HSD_MEDIUM_BLUE, HSD_BLUE],
        )
        fig_scenario.update_traces(texttemplate="$%{text:,.0f}", textposition="outside")
        fig_scenario.update_layout(xaxis_title="", yaxis_title="Estimated Annual Cost", showlegend=False)
        fig_scenario = apply_hsd_theme(fig_scenario)
        st.plotly_chart(fig_scenario, use_container_width=True)

    with col_right:
        month_numbers = list(range(1, 13))
        forecast_data = pd.DataFrame(
            {
                "Month": month_numbers,
                "Current Status Quo": [(status_quo_exposure / 12) * month for month in month_numbers],
                "HSD Low": [(low["annual_cost_with_hsd"] / 12) * month for month in month_numbers],
                "HSD Midpoint": [(base["annual_cost_with_hsd"] / 12) * month for month in month_numbers],
                "HSD High": [(high["annual_cost_with_hsd"] / 12) * month for month in month_numbers],
            }
        )
        fig_forecast = px.line(
            forecast_data,
            x="Month",
            y=["Current Status Quo", "HSD Low", "HSD Midpoint", "HSD High"],
            markers=True,
            title="12-Month Cumulative Cost Scenario",
            color_discrete_sequence=[HSD_NAVY, HSD_SKY_BLUE, HSD_MEDIUM_BLUE, HSD_BLUE],
        )
        fig_forecast.update_layout(xaxis_title="Month", yaxis_title="Cumulative Cost")
        fig_forecast = apply_hsd_theme(fig_forecast)
        st.plotly_chart(fig_forecast, use_container_width=True)

    st.caption(
        "The 12-month chart spreads annual costs evenly by month. Actual billing and savings timing may differ."
    )

# --------------------------------------------------
# TAB 4: HQ BRIEF SUMMARY + WORD DOWNLOAD
# --------------------------------------------------
with tab4:
    st.header("HQ Brief Summary")

    st.markdown(
        f"""
        <div class="hsd-card">
            <h2>{company_html}</h2>
            <p><b>Industry:</b> {industry_html}</p>
            <p><b>Estimated Turnover Rate:</b> {turnover_rate:.1f}%</p>
            <p><b>Annual Employee Departures:</b> {annual_departures:,}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("Status-Quo Exposure", money(status_quo_exposure))
    col_b.metric("Annual Benefit Range", money_range(low["annual_benefit"], high["annual_benefit"]))
    col_c.metric("Net ROI Range", roi_range(low["net_roi_pct"], high["net_roi_pct"]))
    col_d.metric("Benefit-Cost Range", multiple_range(low["benefit_cost_ratio"], high["benefit_cost_ratio"]))

    st.markdown(
        f"""
        <div class="hsd-blue-box">
            <h3>Recommended Sales Message</h3>
            <p>
                Based on the information entered for <b>{company_html if company else "the prospect"}</b>, the estimated annual
                status-quo exposure is <b>{money(status_quo_exposure)}</b>.
            </p>
            <p>
                Under an HSD improvement range of <b>{improvement_low:.1f}% to {improvement_high:.1f}%</b>,
                the estimated annual benefit is <b>{money_range(low["annual_benefit"], high["annual_benefit"])}</b>.
            </p>
            <p>
                HSD can help strengthen employee reach, improve the quality of listening insights, and connect
                feedback to supported retention action. The range should be validated with the prospect and HSD
                subject-matter experts before it is used as a formal financial commitment.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    docx_bytes = create_hq_brief_docx(
        company=company or "New Prospect Company",
        industry=industry,
        turnover_rate=turnover_rate,
        annual_departures=int(annual_departures),
        cost_per_departure=cost_per_departure,
        current_listening_spend=current_listening_spend,
        hsd_investment=hsd_investment,
        maturity_score=maturity_score,
        retention_plan=retention_plan,
        improvement_low=improvement_low,
        improvement_high=improvement_high,
        status_quo_exposure=status_quo_exposure,
        annual_benefit_low=low["annual_benefit"],
        annual_benefit_base=base["annual_benefit"],
        annual_benefit_high=high["annual_benefit"],
        net_benefit_low=low["net_benefit"],
        net_benefit_base=base["net_benefit"],
        net_benefit_high=high["net_benefit"],
        roi_low=low["net_roi_pct"],
        roi_base=base["net_roi_pct"],
        roi_high=high["net_roi_pct"],
        benefit_cost_low=low["benefit_cost_ratio"],
        benefit_cost_base=base["benefit_cost_ratio"],
        benefit_cost_high=high["benefit_cost_ratio"],
        payback_low=low["payback_months"],
        payback_base=base["payback_months"],
        payback_high=high["payback_months"],
    )

    st.download_button(
        label="Download Report",
        data=docx_bytes,
        file_name=f"{safe_filename(company)}_Employee_Listening_Enhancement_Brief.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

    st.warning(
        "This dashboard is a directional pre-sales model. Final client-facing numbers should be validated with "
        "real HR, finance, procurement, turnover, and employee-listening data."
    )
