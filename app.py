# ================================================================
# Mike's True Productivity Machine
# Streamlit App - Version 1
# ================================================================
# Inputs:
#   1. Hours Worked - user input
#   2. Caseload Excel Export - Column A = Name
#   3. Services Report Excel Export - Main Engine
#
# Services Report column map:
#   A = Client Name
#   B = DOS / Date-Time, ignored for now
#   C = Procedure
#   F = Location
#   J = Status
#   K = Units, ignored completely
#   M = ServiceUnits / Duration / Minutes
# ================================================================

import json
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import streamlit as st


# ================================================================
# Page Setup
# ================================================================

st.set_page_config(
    page_title="Mike's True Productivity Machine",
    page_icon="⚙️",
    layout="wide",
)


# ================================================================
# App Constants
# ================================================================

NON_BILLABLE_PROCEDURES = {
    "Non-billable Attempted Contact",
    "Client Non Billable Srvc Must Document",
}

SUCCESSFUL_ENGAGEMENT_PROCEDURES = {
    "Psychosocial Rehab - Individual",
    "TCM/ICC",
    "Psychosocial Rehabilitation Group",
}

SC_LOCATIONS = {
    "Telehealth - Telephone Audio Only Not in Client Home",
    "Telehealth - Telephone Audio Only Client Home",
}

NO_SHOW_CANCEL_STATUSES = {"No Show", "Cancel", "Cancelled", "Canceled", "Cancelation", "Cancellation"}

DEFAULT_UNIT_GRID = [
    (0, 7, 0, 0),
    (8, 22, 1, 15),
    (23, 37, 2, 30),
    (38, 52, 3, 45),
    (53, 67, 4, 60),
    (68, 82, 5, 75),
    (83, 97, 6, 90),
    (98, 112, 7, 105),
    (113, 127, 8, 120),
    (128, 142, 9, 135),
    (143, 157, 10, 150),
    (158, 172, 11, 165),
    (173, 187, 12, 180),
    (188, 202, 13, 195),
    (203, 217, 14, 210),
    (218, 232, 15, 225),
    (233, float("inf"), 16, 240),
]

SC_UNIT_GRID = [
    (0, 7, 0, 0),
    (8, 22, 1, 14),
    (23, 37, 2, 28),
    (38, 52, 3, 42),
    (53, 67, 4, 56),
    (68, 82, 5, 70),
    (83, 97, 6, 84),
    (98, 112, 7, 98),
    (113, 127, 8, 112),
    (128, 142, 9, 126),
    (143, 157, 10, 140),
    (158, 172, 11, 154),
    (173, 187, 12, 168),
    (188, 202, 13, 182),
    (203, 217, 14, 196),
    (218, 232, 15, 210),
    (233, 247, 16, 224),
    (248, 262, 17, 238),
    (263, 277, 18, 252),
    (278, 292, 19, 266),
    (293, 307, 20, 280),
    (308, float("inf"), 21, 294),
]


# ================================================================
# Styling
# ================================================================

st.markdown(
    """
    <style>
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(59, 130, 246, 0.20), transparent 34%),
            radial-gradient(circle at bottom right, rgba(249, 115, 22, 0.16), transparent 32%),
            linear-gradient(135deg, #07111f 0%, #0f172a 45%, #111827 100%);
        color: #e5e7eb;
    }
    .main-title {
        font-size: 2.35rem;
        font-weight: 900;
        letter-spacing: 0.02em;
        color: #f8fafc;
        margin-bottom: 0.15rem;
    }
    .sub-title {
        color: #cbd5e1;
        font-size: 1.02rem;
        margin-bottom: 1.1rem;
    }
    .section-title {
        font-size: 1.45rem;
        font-weight: 800;
        color: #f8fafc;
        margin-top: 1.6rem;
        margin-bottom: 0.65rem;
        padding-top: 0.55rem;
        border-top: 1px solid rgba(148, 163, 184, 0.28);
    }
    .metric-card {
        background: rgba(15, 23, 42, 0.76);
        border: 1px solid rgba(148, 163, 184, 0.26);
        border-radius: 18px;
        padding: 18px 16px;
        box-shadow: 0 12px 30px rgba(0, 0, 0, 0.23);
        min-height: 122px;
    }
    .metric-label {
        color: #cbd5e1;
        font-size: 0.88rem;
        font-weight: 700;
        line-height: 1.15rem;
        min-height: 2.2rem;
    }
    .metric-value {
        color: #ffffff;
        font-size: 2rem;
        font-weight: 900;
        margin-top: 0.35rem;
        line-height: 2.15rem;
    }
    .metric-help {
        color: #94a3b8;
        font-size: 0.75rem;
        margin-top: 0.35rem;
    }
    div[data-testid="stFileUploader"] {
        background: rgba(15, 23, 42, 0.62);
        border: 1px solid rgba(96, 165, 250, 0.30);
        border-radius: 18px;
        padding: 10px;
    }
    .small-note {
        color: #cbd5e1;
        font-size: 0.92rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ================================================================
# Helper Functions
# ================================================================

def clean_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def normalize_key(value: Any) -> str:
    return clean_text(value).casefold()


def parse_minutes(value: Any) -> float:
    """Parse duration values like '55.00 Minutes', '71 Minutes', or numeric values."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = clean_text(value)
    match = re.search(r"(-?\d+(?:\.\d+)?)", text)
    if not match:
        return 0.0
    return float(match.group(1))


def safe_percent(numerator: float, denominator: float) -> float:
    if denominator == 0:
        return 0.0
    return (numerator / denominator) * 100


def fmt_number(value: float, decimals: int = 2) -> str:
    try:
        return f"{float(value):,.{decimals}f}"
    except Exception:
        return "0.00"


def fmt_int(value: float) -> str:
    try:
        return f"{int(round(float(value))):,}"
    except Exception:
        return "0"


def metric_card(label: str, value: Any, help_text: str = "") -> None:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-help">{help_text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def read_excel_first_sheet(uploaded_file: Any) -> pd.DataFrame:
    return pd.read_excel(uploaded_file, sheet_name=0, dtype=object)


def get_col_by_index(df: pd.DataFrame, index: int, label: str) -> pd.Series:
    if df.shape[1] <= index:
        raise ValueError(f"Missing expected column {label}. The uploaded file does not have enough columns.")
    return df.iloc[:, index]


def calculate_from_grid(minutes: float, grid: List[Tuple[float, float, int, int]]) -> Tuple[int, int]:
    for low, high, units, productive_minutes in grid:
        if low <= minutes <= high:
            return units, productive_minutes
    return 0, 0


def classify_modifier(procedure: str, location: str) -> str:
    procedure_clean = clean_text(procedure)
    location_clean = clean_text(location)

    if location_clean in SC_LOCATIONS:
        return "SC"
    if procedure_clean == "Psychosocial Rehabilitation Group":
        return "HQ"
    # HK is intentionally display-ready for future expansion.
    # Current version treats HK as default unless a future rule is added here.
    return "DEFAULT"


def calculate_units(minutes: float, modifier: str) -> Tuple[int, int, str]:
    if modifier == "SC":
        units, productive_minutes = calculate_from_grid(minutes, SC_UNIT_GRID)
        return units, productive_minutes, "SC custom grid"
    units, productive_minutes = calculate_from_grid(minutes, DEFAULT_UNIT_GRID)
    return units, productive_minutes, "Default grid"


def build_caseload_clients(caseload_df: pd.DataFrame) -> pd.DataFrame:
    names = get_col_by_index(caseload_df, 0, "A / Name")
    out = pd.DataFrame({"client_name": names.map(clean_text)})
    out = out[out["client_name"].ne("")].copy()
    out["client_key"] = out["client_name"].map(normalize_key)
    out = out.drop_duplicates(subset=["client_key"]).reset_index(drop=True)
    return out


def build_services_engine(services_df: pd.DataFrame) -> pd.DataFrame:
    engine = pd.DataFrame(
        {
            "client_name": get_col_by_index(services_df, 0, "A / Client Name").map(clean_text),
            "procedure": get_col_by_index(services_df, 2, "C / Procedure").map(clean_text),
            "location": get_col_by_index(services_df, 5, "F / Location").map(clean_text),
            "status": get_col_by_index(services_df, 9, "J / Status").map(clean_text),
            "duration_minutes": get_col_by_index(services_df, 12, "M / ServiceUnits").map(parse_minutes),
        }
    )
    engine = engine[engine["client_name"].ne("") | engine["procedure"].ne("")].copy()
    engine["client_key"] = engine["client_name"].map(normalize_key)
    engine["procedure_key"] = engine["procedure"].map(normalize_key)
    engine["status_key"] = engine["status"].map(normalize_key)

    non_billable_keys = {normalize_key(x) for x in NON_BILLABLE_PROCEDURES}
    successful_keys = {normalize_key(x) for x in SUCCESSFUL_ENGAGEMENT_PROCEDURES}
    no_show_cancel_keys = {normalize_key(x) for x in NO_SHOW_CANCEL_STATUSES}

    engine["is_non_billable"] = engine["procedure_key"].isin(non_billable_keys)
    engine["is_successful_engagement_procedure"] = engine["procedure_key"].isin(successful_keys)
    engine["is_no_show_cancel"] = engine["status_key"].isin(no_show_cancel_keys)
    engine["is_billable"] = (~engine["is_non_billable"]) & (~engine["is_no_show_cancel"]) & engine["procedure"].ne("")

    engine["modifier"] = engine.apply(lambda r: classify_modifier(r["procedure"], r["location"]), axis=1)

    calc = engine.apply(
        lambda r: calculate_units(r["duration_minutes"], r["modifier"]) if r["is_billable"] else (0, 0, "Not billable"),
        axis=1,
        result_type="expand",
    )
    calc.columns = ["calculated_units", "productive_minutes", "unit_rule"]
    engine = pd.concat([engine, calc], axis=1)
    return engine


# ================================================================
# App Header
# ================================================================

st.markdown('<div class="main-title">Mike\'s True Productivity Machine</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-title">One-stop productivity, engagement, and service accountability dashboard.</div>',
    unsafe_allow_html=True,
)


# ================================================================
# Inputs
# ================================================================

st.markdown('<div class="section-title">Inputs</div>', unsafe_allow_html=True)

input_cols = st.columns([1, 1, 1])
with input_cols[0]:
    hours_worked = st.number_input("Enter Hours Worked", min_value=0.0, step=0.01, format="%.2f")
with input_cols[1]:
    caseload_file = st.file_uploader("Upload Caseload Export", type=["xlsx", "xls"], key="caseload_upload")
with input_cols[2]:
    services_file = st.file_uploader("Upload Services Report", type=["xlsx", "xls"], key="services_upload")

st.caption("Services Report uses Column A = Client Name, Column C = Procedure, Column F = Location, Column J = Status, and Column M = Duration/Minutes. Column K Units is ignored.")

if not caseload_file or not services_file:
    st.info("Upload both files to run the dashboard.")
    st.stop()

try:
    caseload_raw = read_excel_first_sheet(caseload_file)
    services_raw = read_excel_first_sheet(services_file)
    caseload_clients_df = build_caseload_clients(caseload_raw)
    services_engine = build_services_engine(services_raw)
except Exception as exc:
    st.error(f"ERROR LOADING/PROCESSING FILE: {exc}. Tell Mike.")
    st.stop()


# ================================================================
# Core Calculations
# ================================================================

minutes_worked = hours_worked * 60

caseload_keys = set(caseload_clients_df["client_key"])
services_clients_df = services_engine[["client_name", "client_key"]].drop_duplicates(subset=["client_key"])
services_keys = set(services_clients_df["client_key"])

non_billable_rows = services_engine[services_engine["is_non_billable"]].copy()
success_rows = services_engine[services_engine["is_successful_engagement_procedure"] & (~services_engine["is_no_show_cancel"])].copy()
billable_rows = services_engine[services_engine["is_billable"]].copy()
no_show_cancel_rows = services_engine[services_engine["is_no_show_cancel"]].copy()

non_billable_client_keys = set(non_billable_rows["client_key"])
success_client_keys = set(success_rows["client_key"])

attempts_only_keys = non_billable_client_keys - success_client_keys
no_engagement_no_attempts_keys = caseload_keys - services_keys

attempts_only_df = (
    non_billable_rows[non_billable_rows["client_key"].isin(attempts_only_keys)][["client_name", "client_key"]]
    .drop_duplicates(subset=["client_key"])
    .sort_values("client_name")
)

no_engagement_df = (
    caseload_clients_df[caseload_clients_df["client_key"].isin(no_engagement_no_attempts_keys)][["client_name", "client_key"]]
    .drop_duplicates(subset=["client_key"])
    .sort_values("client_name")
)

total_caseload = len(caseload_clients_df)
total_services_rendered = len(services_engine[services_engine["procedure"].ne("")])
successful_engagements = len(success_rows)
non_billable_services_rendered = len(non_billable_rows)
attempts_only_count = len(attempts_only_df)
no_engagement_count = len(no_engagement_df)
no_show_cancel_count = len(no_show_cancel_rows)

minutes_billed = float(billable_rows["duration_minutes"].sum())
non_billable_total = float(non_billable_rows["duration_minutes"].sum())
units_billed = int(billable_rows["calculated_units"].sum())
productive_minutes_total = int(billable_rows["productive_minutes"].sum())

billable_minutes_pct = safe_percent(minutes_billed, minutes_worked)
non_billable_pct = safe_percent(non_billable_total, minutes_worked)
billable_units_pct = safe_percent(productive_minutes_total, minutes_worked)


# ================================================================
# KPI Cards
# ================================================================

st.markdown('<div class="section-title">Dashboard Summary</div>', unsafe_allow_html=True)

row1 = st.columns(4)
with row1[0]:
    metric_card("Total Caseload", fmt_int(total_caseload), "Unique clients from Caseload file")
with row1[1]:
    metric_card("Total Services Rendered", fmt_int(total_services_rendered), "All service/procedure rows")
with row1[2]:
    metric_card("Successful Engagements", fmt_int(successful_engagements), "Current approved engagement procedures")
with row1[3]:
    metric_card("Non-Billable Services Rendered", fmt_int(non_billable_services_rendered), "Non-billable service rows")

row2 = st.columns(3)
with row2[0]:
    metric_card("Attempts Only", fmt_int(attempts_only_count), "Non-billable clients without successful engagement")
with row2[1]:
    metric_card("No Engagement/No Attempts", fmt_int(no_engagement_count), "Caseload clients absent from Services Report")
with row2[2]:
    metric_card("No Show/Cancelled Appointments", fmt_int(no_show_cancel_count), "Rows marked No Show or Cancel")


# ================================================================
# Monthly Break Down
# ================================================================

st.markdown('<div class="section-title">Monthly Break Down</div>', unsafe_allow_html=True)

mrow1 = st.columns(4)
with mrow1[0]:
    metric_card("Hours Worked", fmt_number(hours_worked, 2), "User input")
with mrow1[1]:
    metric_card("Minutes Worked", fmt_number(minutes_worked, 1), "Hours × 60")
with mrow1[2]:
    metric_card("Minutes Billed", fmt_number(minutes_billed, 1), "Billable rows, Column M")
with mrow1[3]:
    metric_card("Billable Minutes %", f"{fmt_number(billable_minutes_pct, 2)}%", "Minutes billed ÷ minutes worked")

mrow2 = st.columns(4)
with mrow2[0]:
    metric_card("Units Billed", fmt_int(units_billed), "App-calculated, Column K ignored")
with mrow2[1]:
    metric_card("Billable Units %", f"{fmt_number(billable_units_pct, 2)}%", "Productive minutes ÷ minutes worked")
with mrow2[2]:
    metric_card("Non-Billable Total", fmt_number(non_billable_total, 1), "Non-billable minutes, Column M")
with mrow2[3]:
    metric_card("Non-Billable %", f"{fmt_number(non_billable_pct, 2)}%", "Non-billable minutes ÷ minutes worked")

mrow3 = st.columns(4)
with mrow3[0]:
    metric_card("Documentation Total", "Pending", "Placeholder for future logic")
with mrow3[1]:
    metric_card("Documentation %", "Pending", "Placeholder for future logic")
with mrow3[2]:
    metric_card("Travel Total", "Pending", "Placeholder for future logic")
with mrow3[3]:
    metric_card("Travel %", "Pending", "Placeholder for future logic")


# ================================================================
# Charts
# ================================================================

st.markdown('<div class="section-title">Visual Breakdown</div>', unsafe_allow_html=True)

chart_cols = st.columns(2)
with chart_cols[0]:
    engagement_chart = pd.DataFrame(
        {
            "Category": ["Successful Engagement", "Attempts Only", "No Engagement/No Attempts"],
            "Count": [len(success_client_keys), attempts_only_count, no_engagement_count],
        }
    )
    engagement_chart = engagement_chart[engagement_chart["Count"] > 0]
    if not engagement_chart.empty:
        fig = px.pie(engagement_chart, names="Category", values="Count", title="Caseload Engagement Outcome")
        fig.update_traces(textinfo="percent+label")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No engagement outcome data available yet.")

with chart_cols[1]:
    time_chart = pd.DataFrame(
        {
            "Category": ["Billable Minutes", "Non-Billable Minutes", "Unaccounted Minutes"],
            "Minutes": [minutes_billed, non_billable_total, max(minutes_worked - minutes_billed - non_billable_total, 0)],
        }
    )
    time_chart = time_chart[time_chart["Minutes"] > 0]
    if not time_chart.empty:
        fig2 = px.pie(time_chart, names="Category", values="Minutes", title="Monthly Time Breakdown")
        fig2.update_traces(textinfo="percent+label")
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Enter hours worked to view time breakdown.")

modifier_breakdown = billable_rows.groupby("modifier", dropna=False).agg(
    Services=("procedure", "count"),
    Minutes=("duration_minutes", "sum"),
    Units=("calculated_units", "sum"),
    ProductiveMinutes=("productive_minutes", "sum"),
).reset_index()

if not modifier_breakdown.empty:
    st.markdown("### Modifier / Rate Category Breakdown")
    st.dataframe(modifier_breakdown, use_container_width=True, hide_index=True)


# ================================================================
# Review Lists
# ================================================================

st.markdown('<div class="section-title">Review Lists</div>', unsafe_allow_html=True)

list_cols = st.columns(2)
with list_cols[0]:
    with st.expander(f"Attempts Only Clients ({attempts_only_count})", expanded=False):
        if attempts_only_df.empty:
            st.success("No Attempts Only clients found.")
        else:
            st.dataframe(attempts_only_df[["client_name"]], use_container_width=True, hide_index=True)

with list_cols[1]:
    with st.expander(f"No Engagement/No Attempts Clients ({no_engagement_count})", expanded=False):
        if no_engagement_df.empty:
            st.success("No Engagement/No Attempts clients found.")
        else:
            st.dataframe(no_engagement_df[["client_name"]], use_container_width=True, hide_index=True)

with st.expander("Calculated Billable Services Audit", expanded=False):
    display_cols = [
        "client_name",
        "procedure",
        "location",
        "status",
        "duration_minutes",
        "modifier",
        "calculated_units",
        "productive_minutes",
        "unit_rule",
    ]
    st.dataframe(billable_rows[display_cols], use_container_width=True, hide_index=True)

with st.expander("Non-Billable Services Audit", expanded=False):
    nb_cols = ["client_name", "procedure", "location", "status", "duration_minutes"]
    st.dataframe(non_billable_rows[nb_cols], use_container_width=True, hide_index=True)


# ================================================================
# Downloads
# ================================================================

audit_payload: Dict[str, Any] = {
    "app_name": "Mike's True Productivity Machine",
    "inputs": {
        "hours_worked": hours_worked,
        "minutes_worked": minutes_worked,
    },
    "dashboard_summary": {
        "total_caseload": total_caseload,
        "total_services_rendered": total_services_rendered,
        "successful_engagements": successful_engagements,
        "non_billable_services_rendered": non_billable_services_rendered,
        "attempts_only": attempts_only_count,
        "no_engagement_no_attempts": no_engagement_count,
        "no_show_cancelled_appointments": no_show_cancel_count,
    },
    "monthly_breakdown": {
        "hours_worked": hours_worked,
        "minutes_worked": minutes_worked,
        "minutes_billed": minutes_billed,
        "billable_minutes_percent": billable_minutes_pct,
        "units_billed": units_billed,
        "productive_minutes_total": productive_minutes_total,
        "billable_units_percent": billable_units_pct,
        "non_billable_total_minutes": non_billable_total,
        "non_billable_percent": non_billable_pct,
        "documentation_total": "pending",
        "documentation_percent": "pending",
        "travel_total": "pending",
        "travel_percent": "pending",
    },
    "rules": {
        "ignored_column": "Services Report Column K / Units",
        "service_minutes_column": "Services Report Column M / ServiceUnits",
        "sc_locations": list(SC_LOCATIONS),
        "non_billable_procedures": list(NON_BILLABLE_PROCEDURES),
        "successful_engagement_procedures": list(SUCCESSFUL_ENGAGEMENT_PROCEDURES),
    },
}

st.markdown('<div class="section-title">Downloads</div>', unsafe_allow_html=True)

download_cols = st.columns(3)
with download_cols[0]:
    st.download_button(
        "Download Audit JSON",
        data=json.dumps(audit_payload, indent=2),
        file_name="true_productivity_audit.json",
        mime="application/json",
    )
with download_cols[1]:
    st.download_button(
        "Download Billable Audit CSV",
        data=billable_rows.to_csv(index=False).encode("utf-8"),
        file_name="billable_services_audit.csv",
        mime="text/csv",
    )
with download_cols[2]:
    st.download_button(
        "Download No Engagement List CSV",
        data=no_engagement_df[["client_name"]].to_csv(index=False).encode("utf-8"),
        file_name="no_engagement_no_attempts.csv",
        mime="text/csv",
    )

