from __future__ import annotations

from datetime import date
from io import BytesIO
from pathlib import Path
import base64
import html
import re
import textwrap

import pandas as pd
import plotly.express as px
import streamlit as st
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
    """Use the local HSD logo when available, with a remote fallback."""
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
# HSD STYLE COLORS
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

# Public-facing directional estimates by employee-count band.
# These ranges intentionally avoid displaying exact internal pricing.
PYTHIA_ESTIMATED_PRICING = {
    "Under 500": {"annual": (4500, 5500), "setup": (2000, 3000)},
    "500 to 749": {"annual": (5500, 6500), "setup": (3000, 4000)},
    "750 to 999": {"annual": (5500, 6500), "setup": (3000, 4000)},
    "1,000 to 1,499": {"annual": (5500, 6500), "setup": (4000, 5000)},
    "1,500 to 1,999": {"annual": (5500, 6500), "setup": (4000, 5000)},
    "2,000 to 2,499": {"annual": (5500, 6500), "setup": (4000, 5000)},
    "2,500 to 2,999": {"annual": (6500, 7500), "setup": (4000, 5000)},
    "3,000 to 3,499": {"annual": (6500, 7500), "setup": (4000, 5000)},
    "3,500 to 3,999": {"annual": (6500, 7500), "setup": (4000, 5000)},
    "4,000 to 4,499": {"annual": (6500, 7500), "setup": (4000, 5000)},
    "4,500 to 4,999": {"annual": (6500, 7500), "setup": (4500, 5500)},
    "5,000 to 5,999": {"annual": (8500, 9500), "setup": (4500, 5500)},
    "6,000 to 6,999": {"annual": (8500, 9500), "setup": (4500, 5500)},
    "7,000 to 7,999": {"annual": (8500, 9500), "setup": (4500, 5500)},
    "8,000 to 8,999": {"annual": (8500, 9500), "setup": (4500, 5500)},
    "9,000 to 9,999": {"annual": (8500, 9500), "setup": (4500, 5500)},
    "10,000 to 12,499": {"annual": (9500, 10500), "setup": (5500, 6500)},
    "12,500 to 14,999": {"annual": (9500, 10500), "setup": (5500, 6500)},
    "15,000 to 19,999": {"annual": (9500, 10500), "setup": (5500, 6500)},
    "20,000 to 24,999": {"annual": (9500, 10500), "setup": (5500, 6500)},
    "25,000 to 29,999": {"annual": (9500, 10500), "setup": (5500, 6500)},
    "30,000 to 34,999": {"annual": (9500, 10500), "setup": (5500, 6500)},
    "35,000 to 39,999": {"annual": (9500, 10500), "setup": (5500, 6500)},
}

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

    section[data-testid="stSidebar"] > div {{
        padding-bottom: 110px;
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

    .hsd-results {{
        background: {HSD_WHITE};
        border: 1px solid {HSD_BORDER};
        border-radius: 16px;
        padding: 18px 20px;
        margin: 18px 0 22px 0;
        box-shadow: 0 2px 10px rgba(27, 42, 74, 0.08);
    }}

    .hsd-results-title {{
        color: {HSD_NAVY};
        font-size: 17px;
        font-weight: 700;
        margin-bottom: 14px;
    }}

    .hsd-results-grid {{
        display: grid;
        grid-template-columns: 1fr 1fr 1fr 1.65fr;
        align-items: stretch;
        gap: 0;
    }}

    .hsd-result-stat {{
        display: flex;
        align-items: center;
        gap: 12px;
        padding: 8px 18px 8px 8px;
        border-right: 1px solid {HSD_BORDER};
        min-height: 86px;
    }}

    .hsd-result-icon {{
        width: 48px;
        height: 48px;
        min-width: 48px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 23px;
        font-weight: 700;
    }}

    .hsd-icon-blue {{
        background: #E8F1FF;
        color: #1769E0;
    }}

    .hsd-icon-green {{
        background: #E6F7F2;
        color: #07946F;
    }}

    .hsd-icon-purple {{
        background: #F0EAFE;
        color: #7047D9;
    }}

    .hsd-result-value {{
        font-size: 26px;
        line-height: 1;
        font-weight: 800;
        margin-bottom: 7px;
    }}

    .hsd-value-blue {{
        color: #1769E0;
    }}

    .hsd-value-green {{
        color: #07946F;
    }}

    .hsd-value-purple {{
        color: #7047D9;
    }}

    .hsd-result-label {{
        color: {HSD_TEXT};
        font-size: 13px;
        line-height: 1.35;
    }}

    .hsd-result-quote {{
        background: #EEF6FF;
        border-radius: 13px;
        margin-left: 18px;
        padding: 15px 18px;
        color: {HSD_NAVY};
        min-height: 86px;
    }}

    .hsd-quote-mark {{
        color: {HSD_MEDIUM_BLUE};
        font-size: 32px;
        font-weight: 800;
        line-height: 0.7;
        margin-bottom: 5px;
    }}

    .hsd-quote-text {{
        font-size: 13px;
        line-height: 1.45;
        margin-bottom: 7px;
    }}

    .hsd-quote-source {{
        font-size: 12px;
        font-weight: 700;
        text-align: right;
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

    .hsd-prospect-line {{
        background: {HSD_WHITE};
        border: 1px solid {HSD_BORDER};
        border-radius: 14px;
        padding: 15px 18px;
        margin-bottom: 16px;
        box-shadow: 0 2px 8px rgba(27, 42, 74, 0.06);
        color: {HSD_NAVY};
        font-size: 15px;
    }}

    .hsd-prospect-line strong {{
        font-size: 19px;
    }}

    .hsd-insight-box {{
        background: linear-gradient(135deg, #F4F8FF, #EBF4FF);
        border: 1px solid {HSD_SKY_BLUE};
        border-left: 6px solid {HSD_BLUE};
        border-radius: 14px;
        padding: 18px 20px;
        margin-top: 16px;
        color: {HSD_TEXT};
    }}

    .hsd-insight-box h3 {{
        margin-top: 0;
        margin-bottom: 10px;
        color: {HSD_NAVY};
    }}

    .hsd-insight-box p {{
        margin: 7px 0;
        line-height: 1.5;
    }}

    .hsd-pricing-box {{
        background: {HSD_WHITE};
        border: 1px solid {HSD_BORDER};
        border-radius: 14px;
        padding: 18px;
        box-shadow: 0 2px 8px rgba(27, 42, 74, 0.06);
        min-height: 180px;
    }}

    .hsd-pricing-row {{
        display: flex;
        justify-content: space-between;
        gap: 18px;
        padding: 10px 0;
        border-bottom: 1px solid {HSD_BORDER};
    }}

    .hsd-pricing-row:last-child {{
        border-bottom: none;
    }}

    .hsd-pricing-label {{
        color: {HSD_MUTED};
    }}

    .hsd-pricing-value {{
        color: {HSD_NAVY};
        font-weight: 700;
        text-align: right;
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

    @media (max-width: 1050px) {{
        .hsd-results-grid {{
            grid-template-columns: 1fr 1fr;
            gap: 12px;
        }}

        .hsd-result-stat {{
            border-right: none;
            border-bottom: 1px solid {HSD_BORDER};
        }}

        .hsd-result-quote {{
            grid-column: 1 / -1;
            margin-left: 0;
        }}
    }}

    @media (max-width: 800px) {{
        .hsd-header {{
            display: block;
        }}
        .hsd-logo-wrap {{
            width: fit-content;
            margin-bottom: 16px;
        }}

        .hsd-results-grid {{
            grid-template-columns: 1fr;
        }}

        .hsd-result-quote {{
            grid-column: auto;
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


def compact_money(value: float) -> str:
    """Use compact values for dashboard cards while keeping full values elsewhere."""
    absolute = abs(value)
    sign = "-" if value < 0 else ""

    if absolute >= 1_000_000_000:
        formatted = f"{absolute / 1_000_000_000:.1f}".rstrip("0").rstrip(".")
        return f"{sign}${formatted}B"
    if absolute >= 1_000_000:
        formatted = f"{absolute / 1_000_000:.1f}".rstrip("0").rstrip(".")
        return f"{sign}${formatted}M"
    if absolute >= 1_000:
        formatted = f"{absolute / 1_000:.1f}".rstrip("0").rstrip(".")
        return f"{sign}${formatted}K"
    return f"{sign}${absolute:,.0f}"


def signed_money(value: float) -> str:
    if abs(value) < 0.01:
        return "$0"
    sign = "+" if value > 0 else "-"
    return f"{sign}${abs(value):,.0f}"


def percent(value: float) -> str:
    return f"{value:,.1f}%"


def money_range(low: float, high: float) -> str:
    return money(low) if abs(low - high) < 0.01 else f"{money(low)} - {money(high)}"


def metric_money_range(low: float, high: float) -> str:
    """Use a compact range that fits safely inside Streamlit metric cards."""
    if abs(low - high) < 0.01:
        return compact_money(low)

    low_text = compact_money(low).replace("$", "")
    high_text = compact_money(high).replace("$", "")
    return f"USD {low_text}–{high_text}"


def signed_money_value(value: float) -> str:
    """Format positive savings and negative additional cost clearly."""
    if value < 0:
        return f"-${abs(value):,.0f}"
    return money(value)


def signed_money_range(low: float, high: float) -> str:
    low_value, high_value = sorted([low, high])
    if abs(low_value - high_value) < 0.01:
        return signed_money_value(low_value)
    return f"{signed_money_value(low_value)} - {signed_money_value(high_value)}"


def metric_signed_money_range(low: float, high: float) -> str:
    low_value, high_value = sorted([low, high])
    if abs(low_value - high_value) < 0.01:
        return compact_money(low_value)
    return f"{compact_money(low_value)} to {compact_money(high_value)}"


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
    employee_count_range: str,
    pythia_annual_low: float,
    pythia_annual_high: float,
    pythia_setup_low: float,
    pythia_setup_high: float,
    turnover_rate: float,
    annual_departures: int,
    cost_per_departure: float,
    software_cost: float,
    internal_cost: float,
    external_cost: float,
    current_listening_cost: float,
    annual_turnover_cost: float,
    current_cost_exposure: float,
    first_year_hsd_low: float,
    first_year_hsd_mid: float,
    first_year_hsd_high: float,
    first_year_savings_low: float,
    first_year_savings_mid: float,
    first_year_savings_high: float,
    ongoing_savings_low: float,
    ongoing_savings_mid: float,
    ongoing_savings_high: float,
    maturity_score: float,
    retention_plan: str,
) -> bytes:
    """Create a company-specific two-page Word brief."""
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(9)
    styles["Normal"].paragraph_format.space_after = Pt(3)

    if HSD_LOGO_PATH.exists():
        logo_p = doc.add_paragraph()
        logo_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo_run = logo_p.add_run()
        logo_run.add_picture(str(HSD_LOGO_PATH), width=Inches(1.45))

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title.add_run(f"{company} Employee Listening Enhancement Brief")
    title_run.bold = True
    title_run.font.name = "Arial"
    title_run.font.size = Pt(19)
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
    overview_run = overview.add_run(
        f"This brief summarizes the cost information entered for {company} and compares the current "
        "employee-listening program cost with the directional Pythia platform and setup estimates. "
        "Potential savings are shown only as a full-replacement scenario and are not guaranteed."
    )
    overview_run.font.name = "Arial"
    overview_run.font.size = Pt(9)

    add_doc_heading(doc, "Prospect profile")
    profile_table = doc.add_table(rows=5, cols=2)
    profile_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    profile_table.style = "Table Grid"
    profile_rows = [
        ("Industry", industry or "Not entered"),
        ("Employee count range", employee_count_range),
        ("Estimated turnover rate", percent(turnover_rate)),
        ("Annual employee departures", f"{annual_departures:,}"),
        ("Listening maturity / retention plan", f"{maturity_score:.0f}/100 / {retention_plan}"),
    ]
    for row, (label, value) in zip(profile_table.rows, profile_rows):
        set_cell_text(row.cells[0], label, bold=True, color="1B2A4A")
        shade_cell(row.cells[0], "EBF4FF")
        set_cell_text(row.cells[1], value)

    add_doc_heading(doc, "Current annual cost context")
    cost_table = doc.add_table(rows=7, cols=2)
    cost_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cost_table.style = "Table Grid"
    cost_rows = [
        ("Average cost per employee departure", money(cost_per_departure)),
        ("Estimated annual turnover cost", money(annual_turnover_cost)),
        ("Current listening software cost", money(software_cost)),
        ("Current internal HR effort cost", money(internal_cost)),
        ("Current external support cost", money(external_cost)),
        ("Current listening program cost", money(current_listening_cost)),
        ("Total current cost exposure", money(current_cost_exposure)),
    ]
    for row, (label, value) in zip(cost_table.rows, cost_rows):
        set_cell_text(row.cells[0], label, bold=True, color="1B2A4A")
        shade_cell(row.cells[0], "EBF4FF")
        set_cell_text(row.cells[1], value, bold=label in {"Current listening program cost", "Total current cost exposure"})

    add_doc_heading(doc, "Directional Pythia cost and potential savings")
    savings_table = doc.add_table(rows=6, cols=2)
    savings_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    savings_table.style = "Table Grid"
    savings_rows = [
        ("Annual platform estimate", money_range(pythia_annual_low, pythia_annual_high) if pythia_annual_high > 0 else "Not selected"),
        ("One-time setup estimate", money_range(pythia_setup_low, pythia_setup_high) if pythia_setup_high > 0 else "Not selected"),
        ("Estimated first-year HSD cost", money_range(first_year_hsd_low, first_year_hsd_high) if first_year_hsd_high > 0 else "Not selected"),
        ("Potential first-year savings", signed_money_range(first_year_savings_low, first_year_savings_high) if first_year_hsd_high > 0 and current_listening_cost > 0 else "Not available"),
        ("Potential ongoing annual savings", signed_money_range(ongoing_savings_low, ongoing_savings_high) if pythia_annual_high > 0 and current_listening_cost > 0 else "Not available"),
        ("Savings assumption", "All entered current listening costs are replaced or avoided"),
    ]
    for row, (label, value) in zip(savings_table.rows, savings_rows):
        set_cell_text(row.cells[0], label, bold=True, color="1B2A4A")
        shade_cell(row.cells[0], "EBF4FF")
        set_cell_text(row.cells[1], value, bold="savings" in label.lower())

    sales = doc.add_paragraph()
    sales.paragraph_format.space_before = Pt(6)
    sales_run = sales.add_run(
        f"Sales message: {company} has entered current employee-listening costs of "
        f"{money(current_listening_cost)}. Based on the selected employee range, the estimated first-year "
        f"Pythia cost is {money_range(first_year_hsd_low, first_year_hsd_high) if first_year_hsd_high > 0 else 'not selected'}. "
        f"Under a full-replacement scenario, potential first-year savings are "
        f"{signed_money_range(first_year_savings_low, first_year_savings_high) if first_year_hsd_high > 0 and current_listening_cost > 0 else 'not available'}."
    )
    sales_run.bold = True
    sales_run.font.name = "Arial"
    sales_run.font.size = Pt(9)
    sales_run.font.color.rgb = RGBColor(27, 42, 74)

    doc.add_page_break()
    add_doc_heading(doc, "Savings scenario detail", size=16)
    scenario_table = doc.add_table(rows=4, cols=5)
    scenario_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    scenario_table.style = "Table Grid"
    headers = ["Scenario", "First-year HSD cost", "First-year savings", "Ongoing annual cost", "Ongoing savings"]
    for i, header in enumerate(headers):
        set_cell_text(scenario_table.rows[0].cells[i], header, bold=True, color="FFFFFF")
        shade_cell(scenario_table.rows[0].cells[i], "1B2A4A")

    scenarios = [
        ("Low cost", first_year_hsd_low, first_year_savings_high, pythia_annual_low, ongoing_savings_high),
        ("Midpoint", first_year_hsd_mid, first_year_savings_mid, (pythia_annual_low + pythia_annual_high) / 2, ongoing_savings_mid),
        ("High cost", first_year_hsd_high, first_year_savings_low, pythia_annual_high, ongoing_savings_low),
    ]
    for row, values in zip(scenario_table.rows[1:], scenarios):
        display_values = [values[0], money(values[1]), signed_money_value(values[2]), money(values[3]), signed_money_value(values[4])]
        for i, value in enumerate(display_values):
            set_cell_text(row.cells[i], value, bold=i == 0)
            if i == 0:
                shade_cell(row.cells[i], "EBF4FF")

    add_doc_heading(doc, "Recommended HSD approach")
    add_doc_bullet(doc, "Validate which current software, HR effort, and external support costs would actually be reduced or removed.")
    add_doc_bullet(doc, "Confirm the final HSD scope and proposal price before presenting savings externally.")
    add_doc_bullet(doc, "Treat first-year savings separately because the setup fee occurs only once.")
    add_doc_bullet(doc, "Do not include turnover reduction in savings until HSD approves an impact assumption.")

    add_doc_heading(doc, "Calculation methodology")
    methodology = [
        "Current listening program cost = software cost + internal HR effort cost + external support cost.",
        "Estimated first-year HSD cost = annual Pythia platform estimate + one-time setup estimate.",
        "Potential first-year savings = current listening program cost - estimated first-year HSD cost.",
        "Potential ongoing annual savings = current listening program cost - annual Pythia platform estimate.",
        "The savings scenario assumes all entered current listening costs are replaced or avoided by HSD.",
    ]
    for item in methodology:
        add_doc_bullet(doc, item)

    add_doc_heading(doc, "Important interpretation")
    disclaimer = doc.add_paragraph()
    disclaimer_run = disclaimer.add_run(
        "Potential savings are directional and depend on whether the client can actually eliminate or reduce the entered "
        "software, internal HR effort, and external support costs. The calculation does not include turnover reduction, "
        "ROI, or productivity gains. Final pricing, scope, and client-facing claims should be validated by HSD leadership."
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
# SIDEBAR INPUTS
# --------------------------------------------------
st.sidebar.title("Build Prospect Profile")
company = st.sidebar.text_input("Company Name", "New Prospect Company")
industry = st.sidebar.text_input("Industry", placeholder="e.g., Insurance")

employee_count_range = st.sidebar.selectbox(
    "Employee Count Range",
    ["Select employee range", *PYTHIA_ESTIMATED_PRICING.keys()],
    index=0,
    help="Select the prospect's approximate total employee-count range.",
)

pythia_selected = employee_count_range != "Select employee range"
if pythia_selected:
    pythia_annual_low, pythia_annual_high = PYTHIA_ESTIMATED_PRICING[employee_count_range]["annual"]
    pythia_setup_low, pythia_setup_high = PYTHIA_ESTIMATED_PRICING[employee_count_range]["setup"]
else:
    pythia_annual_low = pythia_annual_high = 0.0
    pythia_setup_low = pythia_setup_high = 0.0

st.sidebar.markdown("---")
st.sidebar.subheader("Turnover & Cost Inputs")

turnover_rate = st.sidebar.number_input(
    "Estimated Turnover Rate %",
    min_value=0.0,
    max_value=100.0,
    value=0.0,
    step=0.5,
    help="This is shown for context. Annual departures are entered separately because headcount is not used.",
)

annual_departures = st.sidebar.number_input(
    "Annual Employee Departures",
    min_value=0,
    value=0,
    step=1,
    help="Enter the number of employees who leave during a typical year.",
)

cost_per_departure = st.sidebar.number_input(
    "Average Cost per Employee Departure",
    min_value=0.0,
    value=0.0,
    step=1000.0,
    help="This may include recruiting, onboarding, vacancy time, training, lost productivity, and knowledge loss.",
)

software_cost = st.sidebar.number_input(
    "Current Annual Listening Software Cost",
    min_value=0.0,
    value=0.0,
    step=5000.0,
)
internal_cost = st.sidebar.number_input(
    "Current Internal HR Effort Cost",
    min_value=0.0,
    value=0.0,
    step=5000.0,
)
external_cost = st.sidebar.number_input(
    "Current External Support Cost",
    min_value=0.0,
    value=0.0,
    step=5000.0,
)

st.sidebar.markdown("---")
st.sidebar.subheader("Listening Inputs")
maturity_score = st.sidebar.number_input(
    "Current Listening Maturity Score",
    min_value=0.0,
    max_value=100.0,
    value=0.0,
    step=1.0,
    help="Use this only when the score comes from an agreed assessment or questionnaire.",
)
retention_plan = st.sidebar.radio(
    "Retention Action Plan",
    ["Not entered", "No", "In progress", "Yes"],
    index=0,
    help="Select the option that best describes the prospect's current retention planning status.",
)

# --------------------------------------------------
# CALCULATIONS
# --------------------------------------------------
annual_turnover_cost = annual_departures * cost_per_departure
current_listening_cost = software_cost + internal_cost + external_cost
current_cost_exposure = annual_turnover_cost + current_listening_cost

pythia_annual_mid = (pythia_annual_low + pythia_annual_high) / 2
pythia_setup_mid = (pythia_setup_low + pythia_setup_high) / 2

first_year_hsd_low = pythia_annual_low + pythia_setup_low
first_year_hsd_mid = pythia_annual_mid + pythia_setup_mid
first_year_hsd_high = pythia_annual_high + pythia_setup_high

# Conservative savings use the higher HSD cost; optimistic savings use the lower HSD cost.
first_year_savings_low = current_listening_cost - first_year_hsd_high
first_year_savings_mid = current_listening_cost - first_year_hsd_mid
first_year_savings_high = current_listening_cost - first_year_hsd_low

ongoing_savings_low = current_listening_cost - pythia_annual_high
ongoing_savings_mid = current_listening_cost - pythia_annual_mid
ongoing_savings_high = current_listening_cost - pythia_annual_low

savings_available = pythia_selected and current_listening_cost > 0

savings_scenarios = pd.DataFrame(
    {
        "Scenario": ["Low", "Midpoint", "High"],
        "Potential First-Year Savings": [
            first_year_savings_low,
            first_year_savings_mid,
            first_year_savings_high,
        ],
    }
)
savings_scenarios["Display Label"] = savings_scenarios["Potential First-Year Savings"].map(signed_money_value)

company_html = html.escape(company or "New Prospect Company")
industry_html = html.escape(industry or "Not entered")
employee_range_html = html.escape(employee_count_range if pythia_selected else "Not selected")

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
            <p>A directional pre-sales tool for comparing current listening spend with Pythia pricing and potential savings.</p>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.caption("Click the HSD logo to open the HSD Metrics website. Results use only the values entered in the sidebar.")

st.markdown(
    textwrap.dedent(
        """
        <div class="hsd-results">
            <div class="hsd-results-title">Proven Results from Our Clients</div>
            <div class="hsd-results-grid">
                <div class="hsd-result-stat">
                    <div class="hsd-result-icon hsd-icon-blue">👥</div>
                    <div>
                        <div class="hsd-result-value hsd-value-blue">97%</div>
                        <div class="hsd-result-label">Client retention rate</div>
                    </div>
                </div>
                <div class="hsd-result-stat">
                    <div class="hsd-result-icon hsd-icon-green">↘</div>
                    <div>
                        <div class="hsd-result-value hsd-value-green">80%</div>
                        <div class="hsd-result-label">Of clients report a decrease in turnover</div>
                    </div>
                </div>
                <div class="hsd-result-stat">
                    <div class="hsd-result-icon hsd-icon-purple">💬</div>
                    <div>
                        <div class="hsd-result-value hsd-value-purple">4M+</div>
                        <div class="hsd-result-label">Employees surveyed across industries</div>
                    </div>
                </div>
                <div class="hsd-result-quote">
                    <div class="hsd-quote-mark">“</div>
                    <div class="hsd-quote-text">
                        We had tried two other platforms before HSD. Our response rates went
                        from 22% to over 80% in the first cycle.
                    </div>
                    <div class="hsd-quote-source">
                        — Chief People Officer,<br>Healthcare Organization
                    </div>
                </div>
            </div>
        </div>
        """
    ),
    unsafe_allow_html=True,
)

# --------------------------------------------------
# CLEANER DASHBOARD TABS
# --------------------------------------------------
tab1, tab2, tab3 = st.tabs(
    [
        "Executive Summary",
        "Cost Details",
        "HQ Brief Summary",
    ]
)

# --------------------------------------------------
# DYNAMIC INTERPRETATION TEXT
# --------------------------------------------------
if not pythia_selected:
    savings_sentence = "Select an employee-count range to calculate the Pythia cost and potential savings."
elif current_listening_cost <= 0:
    savings_sentence = "Enter the client’s current listening-program costs to calculate potential savings."
elif first_year_savings_high < 0:
    savings_sentence = (
        f"The estimated first-year Pythia cost is {signed_money_range(abs(first_year_savings_high), abs(first_year_savings_low))} "
        "above the entered current listening-program cost."
    )
elif first_year_savings_low >= 0:
    savings_sentence = (
        f"Potential first-year savings are {signed_money_range(first_year_savings_low, first_year_savings_high)} "
        "under a full-replacement scenario."
    )
else:
    savings_sentence = (
        "The first-year result ranges from a possible additional cost to possible savings, depending on the final Pythia price."
    )

pythia_sentence = (
    f"For the selected employee range of {employee_range_html}, the directional annual Pythia estimate is "
    f"{money_range(pythia_annual_low, pythia_annual_high)}, with a one-time setup estimate of "
    f"{money_range(pythia_setup_low, pythia_setup_high)}."
    if pythia_selected
    else "Select an employee-count range to generate the directional Pythia platform and setup estimates."
)

# --------------------------------------------------
# TAB 1: EXECUTIVE SUMMARY
# --------------------------------------------------
with tab1:
    st.header("Executive Summary")

    st.markdown(
        f"""
        <div class="hsd-prospect-line">
            <strong>{company_html}</strong>
            &nbsp;&nbsp;|&nbsp;&nbsp;
            {industry_html}
            &nbsp;&nbsp;|&nbsp;&nbsp;
            Employee range: {employee_range_html}
        </div>
        """,
        unsafe_allow_html=True,
    )

    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Annual Turnover Cost", compact_money(annual_turnover_cost))
    kpi2.metric("Current Listening Program Cost", compact_money(current_listening_cost))
    kpi3.metric(
        "Potential First-Year Savings*",
        metric_signed_money_range(first_year_savings_low, first_year_savings_high)
        if savings_available
        else "Not available",
    )

    visual_left, visual_right = st.columns([1.15, 1])

    with visual_left:
        comparison_data = pd.DataFrame(
            {
                "Cost": [
                    "Current Listening Program",
                    "Pythia First-Year Low",
                    "Pythia First-Year Midpoint",
                    "Pythia First-Year High",
                ],
                "Annual Amount": [
                    current_listening_cost,
                    first_year_hsd_low,
                    first_year_hsd_mid,
                    first_year_hsd_high,
                ],
            }
        )

        if pythia_selected and comparison_data["Annual Amount"].sum() > 0:
            fig_summary_comparison = px.bar(
                comparison_data,
                x="Cost",
                y="Annual Amount",
                text="Annual Amount",
                title="Current Listening Cost vs Estimated First-Year Pythia Cost",
                color="Cost",
                color_discrete_sequence=[
                    HSD_NAVY,
                    HSD_SKY_BLUE,
                    HSD_MEDIUM_BLUE,
                    HSD_BLUE,
                ],
            )
            fig_summary_comparison.update_traces(
                texttemplate="$%{text:,.0f}",
                textposition="outside",
            )
            fig_summary_comparison.update_layout(
                xaxis_title="",
                yaxis_title="Annual Cost",
                showlegend=False,
            )
            fig_summary_comparison = apply_hsd_theme(fig_summary_comparison)
            st.plotly_chart(fig_summary_comparison, use_container_width=True)
        else:
            st.info("Select an employee range and enter current listening costs to display the comparison.")

    with visual_right:
        st.markdown(
            f"""
            <div class="hsd-pricing-box">
                <h3>Pythia Cost & Savings Estimate</h3>
                <div class="hsd-pricing-row">
                    <span class="hsd-pricing-label">Employee count range</span>
                    <span class="hsd-pricing-value">{employee_range_html}</span>
                </div>
                <div class="hsd-pricing-row">
                    <span class="hsd-pricing-label">Annual platform estimate</span>
                    <span class="hsd-pricing-value">
                        {metric_money_range(pythia_annual_low, pythia_annual_high) if pythia_selected else "Not selected"}
                    </span>
                </div>
                <div class="hsd-pricing-row">
                    <span class="hsd-pricing-label">One-time setup estimate</span>
                    <span class="hsd-pricing-value">
                        {metric_money_range(pythia_setup_low, pythia_setup_high) if pythia_selected else "Not selected"}
                    </span>
                </div>
                <div class="hsd-pricing-row">
                    <span class="hsd-pricing-label">First-year Pythia cost</span>
                    <span class="hsd-pricing-value">
                        {metric_money_range(first_year_hsd_low, first_year_hsd_high) if pythia_selected else "Not selected"}
                    </span>
                </div>
                <div class="hsd-pricing-row">
                    <span class="hsd-pricing-label">Potential ongoing annual savings*</span>
                    <span class="hsd-pricing-value">
                        {metric_signed_money_range(ongoing_savings_low, ongoing_savings_high) if savings_available else "Not available"}
                    </span>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown(
        f"""
        <div class="hsd-insight-box">
            <h3>What the entered information suggests</h3>
            <p>
                <b>Turnover exposure:</b> {annual_departures:,} annual departures at
                {money(cost_per_departure)} per departure produce an estimated annual
                turnover cost of <b>{money(annual_turnover_cost)}</b>.
            </p>
            <p>
                <b>Current listening spend:</b> The entered software, internal HR effort,
                and external support costs total <b>{money(current_listening_cost)}</b> annually.
            </p>
            <p><b>Potential savings:</b> {savings_sentence}</p>
            <p><b>Pythia estimate:</b> {pythia_sentence}</p>
            <p>
                <b>Important assumption:</b> Potential savings assume that all entered current
                listening costs can be removed or avoided after adopting HSD. Validate this with
                the client before using the figure in a proposal.
            </p>
            <p>
                <b>Listening context:</b> Current maturity is {maturity_score:.0f}/100,
                and the retention action plan is marked as “{html.escape(retention_plan)}.”
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

# --------------------------------------------------
# TAB 2: COST DETAILS
# --------------------------------------------------
with tab2:
    st.header("Cost Details")

    detail_left, detail_right = st.columns(2)

    with detail_left:
        listening_breakdown = pd.DataFrame(
            {
                "Cost Category": ["Software", "Internal HR Effort", "External Support"],
                "Amount": [software_cost, internal_cost, external_cost],
            }
        )

        if listening_breakdown["Amount"].sum() > 0:
            fig_breakdown = px.pie(
                listening_breakdown,
                names="Cost Category",
                values="Amount",
                hole=0.55,
                title="Current Listening Program Cost Breakdown",
                color="Cost Category",
                color_discrete_sequence=[HSD_BLUE, HSD_MEDIUM_BLUE, HSD_SKY_BLUE],
            )
            fig_breakdown.update_traces(
                textposition="inside",
                texttemplate="$%{value:,.0f}<br>%{percent}",
                hovertemplate="<b>%{label}</b><br>Annual cost: $%{value:,.0f}<br>Share: %{percent}<extra></extra>",
                sort=False,
            )
            fig_breakdown.add_annotation(
                text=f"<b>{money(current_listening_cost)}</b><br>Total",
                x=0.5,
                y=0.5,
                showarrow=False,
                font=dict(size=16, color=HSD_NAVY),
            )
            fig_breakdown = apply_hsd_theme(fig_breakdown)
            st.plotly_chart(fig_breakdown, use_container_width=True)
        else:
            st.info("Enter software, HR effort, or external-support costs to display the donut chart.")

    with detail_right:
        if savings_available:
            fig_savings = px.bar(
                savings_scenarios,
                x="Scenario",
                y="Potential First-Year Savings",
                text="Display Label",
                title="Potential First-Year Savings Scenarios",
                color="Scenario",
                color_discrete_sequence=[HSD_SKY_BLUE, HSD_MEDIUM_BLUE, HSD_NAVY],
            )
            fig_savings.update_traces(textposition="outside")
            fig_savings.update_layout(
                xaxis_title="",
                yaxis_title="Savings / (Additional Cost)",
                showlegend=False,
            )
            fig_savings = apply_hsd_theme(fig_savings)
            st.plotly_chart(fig_savings, use_container_width=True)
        else:
            st.info("Select an employee range and enter current listening costs to display savings scenarios.")

    exact_values = pd.DataFrame(
        {
            "Measure": [
                "Estimated turnover rate",
                "Annual employee departures",
                "Average cost per employee departure",
                "Annual turnover cost",
                "Current listening software cost",
                "Current internal HR effort cost",
                "Current external support cost",
                "Current listening program cost",
                "Total current cost exposure",
                "Pythia annual platform estimate",
                "Pythia one-time setup estimate",
                "Estimated first-year Pythia cost",
                "Potential first-year savings",
                "Potential ongoing annual savings",
            ],
            "Value": [
                percent(turnover_rate),
                f"{annual_departures:,}",
                money(cost_per_departure),
                money(annual_turnover_cost),
                money(software_cost),
                money(internal_cost),
                money(external_cost),
                money(current_listening_cost),
                money(current_cost_exposure),
                money_range(pythia_annual_low, pythia_annual_high) if pythia_selected else "Not selected",
                money_range(pythia_setup_low, pythia_setup_high) if pythia_selected else "Not selected",
                money_range(first_year_hsd_low, first_year_hsd_high) if pythia_selected else "Not selected",
                signed_money_range(first_year_savings_low, first_year_savings_high) if savings_available else "Not available",
                signed_money_range(ongoing_savings_low, ongoing_savings_high) if savings_available else "Not available",
            ],
        }
    )

    with st.expander("View exact values"):
        st.dataframe(exact_values, use_container_width=True, hide_index=True)

    with st.expander("View calculation formulas"):
        st.markdown(
            """
            **Current listening program cost** = Software cost + Internal HR effort cost + External support cost  
            **Estimated first-year Pythia cost** = Annual platform estimate + One-time setup estimate  
            **Potential first-year savings** = Current listening program cost − Estimated first-year Pythia cost  
            **Potential ongoing annual savings** = Current listening program cost − Annual Pythia platform estimate
            """
        )

    st.warning(
        "Potential savings assume that all entered current listening costs are replaced or avoided by HSD. "
        "The calculation does not include turnover reduction or productivity benefits. Validate the assumption with the client and HSD leadership."
    )

# --------------------------------------------------
# TAB 3: HQ BRIEF SUMMARY + DOWNLOAD
# --------------------------------------------------
with tab3:
    st.header("HQ Brief Summary")

    st.markdown(
        f"""
        <div class="hsd-blue-box">
            <h3>Recommended Sales Summary</h3>
            <p>
                <b>{company_html}</b> is a prospect in the <b>{industry_html}</b> industry,
                with a selected employee range of <b>{employee_range_html}</b>.
            </p>
            <p>
                Based on the entered values, estimated annual turnover cost is
                <b>{money(annual_turnover_cost)}</b>, while the current employee-listening
                program cost is <b>{money(current_listening_cost)}</b>.
            </p>
            <p>{pythia_sentence}</p>
            <p>
                <b>Potential first-year savings:</b>
                {signed_money_range(first_year_savings_low, first_year_savings_high) if savings_available else "Not available"}.
                {savings_sentence}
            </p>
            <p>
                The savings estimate assumes that all entered current listening costs can be
                eliminated or avoided. Confirm the final scope and replacement assumptions before
                including the estimate in a formal proposal.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    docx_bytes = create_hq_brief_docx(
        company=company or "New Prospect Company",
        industry=industry,
        employee_count_range=employee_count_range if pythia_selected else "Not selected",
        pythia_annual_low=pythia_annual_low,
        pythia_annual_high=pythia_annual_high,
        pythia_setup_low=pythia_setup_low,
        pythia_setup_high=pythia_setup_high,
        turnover_rate=turnover_rate,
        annual_departures=int(annual_departures),
        cost_per_departure=cost_per_departure,
        software_cost=software_cost,
        internal_cost=internal_cost,
        external_cost=external_cost,
        current_listening_cost=current_listening_cost,
        annual_turnover_cost=annual_turnover_cost,
        current_cost_exposure=current_cost_exposure,
        first_year_hsd_low=first_year_hsd_low,
        first_year_hsd_mid=first_year_hsd_mid,
        first_year_hsd_high=first_year_hsd_high,
        first_year_savings_low=first_year_savings_low,
        first_year_savings_mid=first_year_savings_mid,
        first_year_savings_high=first_year_savings_high,
        ongoing_savings_low=ongoing_savings_low,
        ongoing_savings_mid=ongoing_savings_mid,
        ongoing_savings_high=ongoing_savings_high,
        maturity_score=maturity_score,
        retention_plan=retention_plan,
    )

    st.download_button(
        label="Download Report",
        data=docx_bytes,
        file_name=f"{safe_filename(company)}_Employee_Listening_Enhancement_Brief.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
    )

    st.caption(
        "Directional pre-sales estimate. Final pricing, scope, replacement assumptions, and client-facing statements should be validated by HSD leadership."
    )
