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
    hsd_cost_low: float,
    hsd_cost_mid: float,
    hsd_cost_high: float,
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
        f"This brief summarizes the cost information entered for {company} and compares the company’s "
        "current employee-listening costs with an estimated annual HSD service cost range. The report "
        "does not assume that HSD will automatically replace the company’s current services."
    )
    overview_run.font.name = "Arial"
    overview_run.font.size = Pt(9)

    add_doc_heading(doc, "Prospect profile")
    profile_table = doc.add_table(rows=7, cols=2)
    profile_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    profile_table.style = "Table Grid"
    profile_rows = [
        ("Industry", industry or "Not entered"),
        ("Employee count range", employee_count_range),
        ("Estimated annual Pythia platform cost", money_range(pythia_annual_low, pythia_annual_high) if pythia_annual_high > 0 else "Not selected"),
        ("Estimated one-time Pythia setup cost", money_range(pythia_setup_low, pythia_setup_high) if pythia_setup_high > 0 else "Not selected"),
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

    add_doc_heading(doc, "Estimated annual HSD service cost")
    hsd_table = doc.add_table(rows=4, cols=2)
    hsd_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hsd_table.style = "Table Grid"
    hsd_rows = [
        ("Low estimate", money(hsd_cost_low)),
        ("Midpoint estimate", money(hsd_cost_mid)),
        ("High estimate", money(hsd_cost_high)),
        ("Current listening program cost for comparison", money(current_listening_cost)),
    ]
    for row, (label, value) in zip(hsd_table.rows, hsd_rows):
        set_cell_text(row.cells[0], label, bold=True, color="1B2A4A")
        shade_cell(row.cells[0], "EBF4FF")
        set_cell_text(row.cells[1], value, bold=True)

    sales = doc.add_paragraph()
    sales.paragraph_format.space_before = Pt(6)
    sales_run = sales.add_run(
        f"Sales message: Based on the entered information, {company} has an estimated annual turnover "
        f"cost of {money(annual_turnover_cost)} and current employee-listening costs of "
        f"{money(current_listening_cost)}. The estimated annual HSD service cost is "
        f"{money_range(hsd_cost_low, hsd_cost_high)}. For the selected employee-count band, the directional "
        f"Pythia platform estimate is {money_range(pythia_annual_low, pythia_annual_high) if pythia_annual_high > 0 else 'not selected'}, "
        f"plus an estimated one-time setup cost of {money_range(pythia_setup_low, pythia_setup_high) if pythia_setup_high > 0 else 'not selected'}."
    )
    sales_run.bold = True
    sales_run.font.name = "Arial"
    sales_run.font.size = Pt(9)
    sales_run.font.color.rgb = RGBColor(27, 42, 74)

    # PAGE 2
    doc.add_page_break()
    add_doc_heading(doc, "HSD service cost scenarios", size=16)
    scenario_table = doc.add_table(rows=4, cols=5)
    scenario_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    scenario_table.style = "Table Grid"
    headers = [
        "Scenario",
        "Annual HSD cost",
        "Monthly equivalent",
        "Difference vs current listening cost",
        "Meaning",
    ]
    for i, header in enumerate(headers):
        set_cell_text(scenario_table.rows[0].cells[i], header, bold=True, color="FFFFFF")
        shade_cell(scenario_table.rows[0].cells[i], "1B2A4A")

    scenarios = [
        ("Low", hsd_cost_low),
        ("Midpoint", hsd_cost_mid),
        ("High", hsd_cost_high),
    ]
    for row, (scenario_name, annual_cost) in zip(scenario_table.rows[1:], scenarios):
        difference = annual_cost - current_listening_cost
        if difference > 0:
            meaning = "Estimated HSD cost is above current listening spend."
        elif difference < 0:
            meaning = "Estimated HSD cost is below current listening spend."
        else:
            meaning = "Estimated HSD cost matches current listening spend."
        values = [
            scenario_name,
            money(annual_cost),
            money(annual_cost / 12),
            signed_money(difference),
            meaning,
        ]
        for i, value in enumerate(values):
            set_cell_text(row.cells[i], value, bold=i == 0)
            if i == 0:
                shade_cell(row.cells[i], "EBF4FF")

    add_doc_heading(doc, "Recommended HSD approach")
    add_doc_bullet(doc, "Reach employees through multiple listening methods, including groups that may be missed by email-only surveys.")
    add_doc_bullet(doc, "Gather feedback across important employee moments instead of relying only on one annual survey.")
    add_doc_bullet(doc, "Analyze feedback and connect the findings to business and retention priorities.")
    add_doc_bullet(doc, "Give leaders clear priorities and support for follow-through.")

    add_doc_heading(doc, "Calculation methodology")
    methodology = [
        "Annual turnover cost = annual employee departures x average cost per employee departure.",
        "Current listening program cost = software cost + internal HR effort cost + external support cost.",
        "Total current cost exposure = annual turnover cost + current listening program cost.",
        "HSD midpoint service cost = average of the low and high HSD service cost estimates.",
        "Difference vs current listening cost = HSD service cost estimate - current listening program cost.",
    ]
    for item in methodology:
        add_doc_bullet(doc, item)

    add_doc_heading(doc, "Important interpretation")
    disclaimer = doc.add_paragraph()
    disclaimer_run = disclaimer.add_run(
        "The current software, internal HR, and external support costs are shown for comparison only. "
        "The model does not assume that HSD will replace all existing services. Savings, ROI, and payback "
        "are not calculated because an approved HSD impact assumption has not been entered. Final pricing "
        "and client-facing statements should be validated by HSD leadership."
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
st.sidebar.subheader("Estimated HSD Service Cost Range")
raw_hsd_cost_low = st.sidebar.number_input(
    "Low Estimated Annual HSD Service Cost ($)",
    min_value=0.0,
    value=0.0,
    step=5000.0,
)
raw_hsd_cost_high = st.sidebar.number_input(
    "High Estimated Annual HSD Service Cost ($)",
    min_value=0.0,
    value=0.0,
    step=5000.0,
)

hsd_cost_low = min(raw_hsd_cost_low, raw_hsd_cost_high)
hsd_cost_high = max(raw_hsd_cost_low, raw_hsd_cost_high)
hsd_cost_mid = (hsd_cost_low + hsd_cost_high) / 2

if raw_hsd_cost_high < raw_hsd_cost_low:
    st.sidebar.warning("The app has reordered the HSD service cost values from low to high.")

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

hsd_cost_scenarios = pd.DataFrame(
    {
        "Scenario": ["Low", "Midpoint", "High"],
        "Annual HSD Service Cost": [hsd_cost_low, hsd_cost_mid, hsd_cost_high],
    }
)
hsd_cost_scenarios["Monthly Equivalent"] = hsd_cost_scenarios["Annual HSD Service Cost"] / 12
hsd_cost_scenarios["Difference vs Current Listening Cost"] = (
    hsd_cost_scenarios["Annual HSD Service Cost"] - current_listening_cost
)

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
            <p>A directional pre-sales tool for comparing current cost exposure with an estimated HSD service cost range.</p>
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
if current_listening_cost > 0:
    midpoint_difference = hsd_cost_mid - current_listening_cost
    if hsd_cost_high <= 0:
        cost_comparison_sentence = (
            "Enter the estimated HSD service cost range to compare it with the current listening-program cost."
        )
    elif midpoint_difference < 0:
        cost_comparison_sentence = (
            f"The midpoint HSD estimate is {money(abs(midpoint_difference))} below the current "
            "listening-program cost. This is a price comparison only, not a confirmed savings claim."
        )
    elif midpoint_difference > 0:
        cost_comparison_sentence = (
            f"The midpoint HSD estimate is {money(midpoint_difference)} above the current "
            "listening-program cost. The difference should be reviewed against the proposed HSD scope."
        )
    else:
        cost_comparison_sentence = (
            "The midpoint HSD estimate is equal to the current listening-program cost."
        )
else:
    cost_comparison_sentence = (
        "Current listening-program costs have not been entered, so a direct cost comparison is not yet available."
    )

pythia_sentence = (
    f"For the selected employee range of {employee_range_html}, the directional Pythia estimate is "
    f"{money_range(pythia_annual_low, pythia_annual_high)} annually, plus "
    f"{money_range(pythia_setup_low, pythia_setup_high)} as a one-time setup estimate."
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

    # Only the three most important financial figures are shown as cards.
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Annual Turnover Cost", compact_money(annual_turnover_cost))
    kpi2.metric("Current Listening Program Cost", compact_money(current_listening_cost))
    kpi3.metric(
        "Estimated HSD Service Cost",
        metric_money_range(hsd_cost_low, hsd_cost_high),
    )

    visual_left, visual_right = st.columns([1.15, 1])

    with visual_left:
        comparison_data = pd.DataFrame(
            {
                "Cost": [
                    "Current Listening Program",
                    "HSD Low",
                    "HSD Midpoint",
                    "HSD High",
                ],
                "Annual Amount": [
                    current_listening_cost,
                    hsd_cost_low,
                    hsd_cost_mid,
                    hsd_cost_high,
                ],
            }
        )

        if comparison_data["Annual Amount"].sum() > 0:
            fig_summary_comparison = px.bar(
                comparison_data,
                x="Cost",
                y="Annual Amount",
                text="Annual Amount",
                title="Current Listening Cost vs Estimated HSD Cost",
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
            st.info("Enter current listening costs and the HSD service-cost range to display the comparison.")

    with visual_right:
        st.markdown(
            f"""
            <div class="hsd-pricing-box">
                <h3>Pythia Pricing Estimate</h3>
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
                <b>Listening-program cost:</b> The entered software, internal HR effort,
                and external support costs total <b>{money(current_listening_cost)}</b> annually.
            </p>
            <p><b>HSD cost comparison:</b> {cost_comparison_sentence}</p>
            <p><b>Pythia estimate:</b> {pythia_sentence}</p>
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
                "Cost Category": [
                    "Software",
                    "Internal HR Effort",
                    "External Support",
                ],
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
                color_discrete_sequence=[
                    HSD_BLUE,
                    HSD_MEDIUM_BLUE,
                    HSD_SKY_BLUE,
                ],
            )
            fig_breakdown.update_traces(
                textposition="inside",
                texttemplate="$%{value:,.0f}<br>%{percent}",
                hovertemplate=(
                    "<b>%{label}</b><br>Annual cost: $%{value:,.0f}"
                    "<br>Share: %{percent}<extra></extra>"
                ),
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
        if hsd_cost_high > 0:
            fig_hsd_range = px.bar(
                hsd_cost_scenarios,
                x="Scenario",
                y="Annual HSD Service Cost",
                text="Annual HSD Service Cost",
                title="Estimated Annual HSD Service Cost Range",
                color="Scenario",
                color_discrete_sequence=[
                    HSD_SKY_BLUE,
                    HSD_MEDIUM_BLUE,
                    HSD_NAVY,
                ],
            )
            fig_hsd_range.update_traces(
                texttemplate="$%{text:,.0f}",
                textposition="outside",
            )
            fig_hsd_range.update_layout(
                xaxis_title="",
                yaxis_title="Annual HSD Service Cost",
                showlegend=False,
            )
            fig_hsd_range = apply_hsd_theme(fig_hsd_range)
            st.plotly_chart(fig_hsd_range, use_container_width=True)
        else:
            st.info("Enter the low and high HSD service-cost estimates to display the scenario chart.")

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
                "Low HSD service cost",
                "Midpoint HSD service cost",
                "High HSD service cost",
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
                money(hsd_cost_low),
                money(hsd_cost_mid),
                money(hsd_cost_high),
            ],
        }
    )

    with st.expander("View exact values"):
        st.dataframe(exact_values, use_container_width=True, hide_index=True)

    with st.expander("View calculation formulas"):
        st.markdown(
            """
            **Annual turnover cost** = Annual employee departures × Average cost per employee departure  
            **Current listening program cost** = Software cost + Internal HR effort cost + External support cost  
            **Total current cost exposure** = Annual turnover cost + Current listening program cost  
            **HSD midpoint cost** = (Low HSD cost + High HSD cost) ÷ 2  
            **Difference for comparison** = HSD service cost − Current listening program cost
            """
        )

    st.caption(
        "The current listening-program cost and HSD cost are compared for discussion only. "
        "The dashboard does not claim that HSD will replace every current service or that the difference is guaranteed savings."
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
            <p>
                The estimated annual HSD service cost is
                <b>{money_range(hsd_cost_low, hsd_cost_high)}</b>.
                {cost_comparison_sentence}
            </p>
            <p>{pythia_sentence}</p>
            <p>
                These figures are directional and should be validated with the client
                and HSD leadership before being included in a formal proposal.
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
        hsd_cost_low=hsd_cost_low,
        hsd_cost_mid=hsd_cost_mid,
        hsd_cost_high=hsd_cost_high,
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
        "Directional pre-sales estimate. Final pricing, scope, and client-facing statements should be validated by HSD leadership."
    )
