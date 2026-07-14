#!/usr/bin/env python3
"""
Excel workbook generator for the Meta Ads Performance Intelligence report.

Builds meta_ads_performance_intelligence.xlsx from three already-generated,
already-validated inputs:

    data/meta_campaign_daily.csv       -- Layer 1 raw campaign daily data
    data/performance_findings.json     -- deterministic analysis (Layer 2)
    reports/meta_performance_report.md -- Claude's executive markdown report

This script performs its own aggregation for two presentation-only views
(a full-period per-campaign rollup for the Campaign Performance sheet, and a
full-period daily rollup for the dashboard's trend charts) but reuses
performance_analyzer.py's existing aggregate_window()/safe_divide()/
safe_pct_change() helpers rather than reimplementing that math -- every
ratio in this workbook is computed exactly the same way as the rest of the
pipeline. No metric definitions are duplicated or altered here.

The workbook contains exactly four worksheets, in this order:
    1. Raw Data
    2. Executive Dashboard
    3. Campaign Performance
    4. AI Analysis
"""

import argparse
import json
import os
import re
import sys
from typing import Any, Dict, List, Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from currency_utils import get_currency_excel_number_format
from performance_analyzer import aggregate_window, safe_divide, safe_pct_change

DEFAULT_CSV_PATH = os.path.join("data", "meta_campaign_daily.csv")
DEFAULT_FINDINGS_PATH = os.path.join("data", "performance_findings.json")
DEFAULT_REPORT_PATH = os.path.join("reports", "meta_performance_report.md")
DEFAULT_OUTPUT_PATH = "meta_ads_performance_intelligence.xlsx"

REQUIRED_WORKSHEET_ORDER = ["Raw Data", "Executive Dashboard", "Campaign Performance", "AI Analysis"]

# Columns exactly as produced by meta_collector.py's OUTPUT_COLUMNS. This
# script does not assume any other schema for the raw CSV.
RAW_DATA_COLUMNS = [
    "date", "campaign_id", "campaign_name", "objective", "spend", "impressions", "reach",
    "clicks", "ctr", "cpc", "cpm", "frequency", "purchases", "purchase_value", "CPA", "ROAS",
    "add_to_carts", "initiate_checkouts", "landing_page_views",
    "click_to_landing_page_rate", "landing_page_to_purchase_rate",
]
RAW_DATA_TEXT_COLUMNS = {"campaign_id", "campaign_name", "objective"}
RAW_DATA_DATE_COLUMNS = {"date"}

# Meta's raw `ctr` field is already a percent-number (e.g. 2.35 meaning
# 2.35%), not a 0-1 fraction -- use a custom format that doesn't re-multiply
# by 100. click_to_landing_page_rate / landing_page_to_purchase_rate are
# true 0-1 fractions computed by meta_collector.py's own safe_divide, so
# they use Excel's native percentage format.
#
# Currency columns (spend, purchase_value, cpc, cpm, CPA) are NOT listed
# here -- their number_format is built dynamically per the account's actual
# currency via get_currency_excel_number_format() (see RAW_DATA_CURRENCY_COLUMNS
# below). This workbook never hardcodes "$"/USD.
RAW_DATA_NUMBER_FORMATS = {
    "impressions": '#,##0',
    "reach": '#,##0',
    "clicks": '#,##0',
    "purchases": '#,##0',
    "add_to_carts": '#,##0',
    "initiate_checkouts": '#,##0',
    "landing_page_views": '#,##0',
    "frequency": '0.00',
    "ROAS": '0.00',
    "ctr": '0.00"%"',
    "click_to_landing_page_rate": '0.0%',
    "landing_page_to_purchase_rate": '0.0%',
}
RAW_DATA_CURRENCY_COLUMNS = {"spend", "purchase_value", "cpc", "cpm", "CPA"}

CHANGE_NUMBER_FORMAT = '+0.0"%";-0.0"%";0.0"%"'

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
ORANGE_FILL = PatternFill("solid", fgColor="FCE4D6")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
GRAY_FILL = PatternFill("solid", fgColor="D9D9D9")

# KPI definitions for the Executive Dashboard: (label, metric key in the
# last_7_days/previous_7_days dicts, pct-change key, format type, direction).
# direction=None means the metric is not treated as inherently good or bad
# when it increases (e.g. spend) -- no conditional coloring is applied.
KPI_DEFINITIONS = [
    ("Total Spend", "spend", "spend_pct_change", "currency", None),
    ("Purchase Value / Revenue", "purchase_value", "purchase_value_pct_change", "currency", "positive_increase"),
    ("Purchases", "purchases", "purchases_pct_change", "integer", "positive_increase"),
    ("ROAS", "roas", "roas_pct_change", "decimal", "positive_increase"),
    ("CPA", "cpa", "cpa_pct_change", "currency", "negative_increase"),
    ("CTR", "ctr", "ctr_pct_change", "true_fraction_percent", "positive_increase"),
    ("CPC", "cpc", "cpc_pct_change", "currency", "negative_increase"),
    ("CPM", "cpm", "cpm_pct_change", "currency", "negative_increase"),
    ("Add to Carts", "add_to_carts", "add_to_carts_pct_change", "integer", "positive_increase"),
    ("Initiated Checkouts", "initiated_checkouts", "initiated_checkouts_pct_change", "integer", "positive_increase"),
    (
        "Checkout to Purchase Rate",
        "checkout_to_purchase_rate",
        "checkout_to_purchase_rate_pct_change",
        "true_fraction_percent",
        "positive_increase",
    ),
]

CAMPAIGN_PERFORMANCE_COLUMNS = [
    "Campaign ID", "Campaign Name", "Objective", "Spend", "Impressions", "Clicks", "CTR", "CPC", "CPM",
    "Purchases", "Purchase Value", "CPA", "ROAS", "Add to Carts", "Initiated Checkouts",
    "Checkout to Purchase Rate", "Frequency", "AI Status",
]
# As with Raw Data, currency columns are formatted dynamically per the
# account's actual currency (see CAMPAIGN_PERFORMANCE_CURRENCY_COLUMNS) --
# never hardcoded here.
CAMPAIGN_PERFORMANCE_NUMBER_FORMATS = {
    "Impressions": '#,##0',
    "Clicks": '#,##0',
    "Purchases": '#,##0',
    "Add to Carts": '#,##0',
    "Initiated Checkouts": '#,##0',
    "CTR": '0.0%',
    "Checkout to Purchase Rate": '0.0%',
    "ROAS": '0.00',
    "Frequency": '0.00',
}
CAMPAIGN_PERFORMANCE_CURRENCY_COLUMNS = {"Spend", "Purchase Value", "CPC", "CPM", "CPA"}
AI_STATUS_FILLS = {
    "TRACKING WARNING": RED_FILL,
    "EFFICIENCY RISK": ORANGE_FILL,
    "CREATIVE FATIGUE": YELLOW_FILL,
    "SCALE": GREEN_FILL,
    "INSUFFICIENT DATA": GRAY_FILL,
}

# Calculation area (chart source data) lives in its own column block, well
# separated from the KPI table and charts so it reads as a distinct section.
CALC_COL_OFFSET = 10  # column J
CHARTS_START_ROW = 18
CHART_ROW_SPACING = 20

REQUIRED_AI_ANALYSIS_SECTIONS = [
    "EXECUTIVE SUMMARY",
    "DATA INTEGRITY & DECISION CONFIDENCE",
    "CRITICAL ISSUES",
    "HIGH PRIORITY RISKS",
    "SCALING OPPORTUNITIES",
    "CREATIVE FATIGUE SIGNALS",
    "FUNNEL ANALYSIS",
    "CAMPAIGNS TO WATCH",
    "NEXT 24 HOURS ACTION PLAN",
]
ACTION_SECTIONS = {
    "CRITICAL ISSUES",
    "HIGH PRIORITY RISKS",
    "SCALING OPPORTUNITIES",
    "CREATIVE FATIGUE SIGNALS",
    "CAMPAIGNS TO WATCH",
    "NEXT 24 HOURS ACTION PLAN",
}
RECOMMENDATION_FIELDS = ["Priority", "Decision Confidence", "Campaign", "Evidence", "Observation", "Recommended Action"]
RECOMMENDATION_FIELD_PATTERN = re.compile(
    r'^\**\s*(Priority|Decision Confidence|Campaign|Evidence|Observation|Recommended Action)\s*:\**\s*(.*)$',
    re.IGNORECASE,
)


class ExcelReportError(Exception):
    """Raised when the Excel workbook cannot be generated reliably."""


# ---------------------------------------------------------------------------
# Loading inputs
# ---------------------------------------------------------------------------

def load_raw_data(csv_path: str) -> pd.DataFrame:
    try:
        df = pd.read_csv(csv_path, dtype={"campaign_id": str})
    except FileNotFoundError as exc:
        raise ExcelReportError(
            f"CSV file '{csv_path}' not found. Run meta_collector.py first."
        ) from exc

    missing_columns = [col for col in RAW_DATA_COLUMNS if col not in df.columns]
    if missing_columns:
        raise ExcelReportError(
            f"CSV file '{csv_path}' is missing expected column(s): {', '.join(missing_columns)}."
        )
    if df.empty:
        raise ExcelReportError(f"CSV file '{csv_path}' contains no rows.")

    df["date"] = pd.to_datetime(df["date"]).dt.date
    numeric_columns = [c for c in RAW_DATA_COLUMNS if c not in RAW_DATA_TEXT_COLUMNS | RAW_DATA_DATE_COLUMNS]
    for col in numeric_columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    return df


def load_findings(path: str) -> dict:
    try:
        with open(path) as findings_file:
            return json.load(findings_file)
    except FileNotFoundError as exc:
        raise ExcelReportError(
            f"Findings file '{path}' not found. Run performance_analyzer.py first."
        ) from exc
    except json.JSONDecodeError as exc:
        raise ExcelReportError(f"Findings file '{path}' is not valid JSON: {exc}") from exc


def load_ai_report(path: str) -> str:
    try:
        with open(path) as report_file:
            return report_file.read()
    except FileNotFoundError as exc:
        raise ExcelReportError(
            f"AI report file '{path}' not found. Run ai_analyst.py first."
        ) from exc


# ---------------------------------------------------------------------------
# Aggregation (reuses performance_analyzer.py's math -- no reimplementation)
# ---------------------------------------------------------------------------

def aggregate_campaign_totals(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """Full-collected-period per-campaign rollup for the Campaign Performance sheet.

    Ratios are derived from aggregated totals via performance_analyzer.py's
    own aggregate_window(), exactly as the rest of the pipeline computes
    them -- not reimplemented or averaged from daily ratios.
    """
    min_date = df["date"].min()
    max_date = df["date"].max()
    campaigns = df[["campaign_id", "campaign_name"]].drop_duplicates()

    results = []
    for _, row in campaigns.iterrows():
        campaign_id = row["campaign_id"]
        campaign_name = row["campaign_name"]
        campaign_rows = df[df["campaign_id"] == campaign_id]

        totals = aggregate_window(campaign_rows, min_date, max_date)

        objective_series = campaign_rows["objective"].dropna()
        objective = str(objective_series.iloc[0]) if not objective_series.empty else ""

        # Reach is not additive across days (it's per-day unique reach), so a
        # true lifetime Frequency (impressions / unique reach) cannot be
        # derived from summed daily reach. The arithmetic mean of the daily
        # Frequency values is used as a pragmatic estimate instead.
        frequency_series = pd.to_numeric(campaign_rows["frequency"], errors="coerce").dropna()
        frequency = float(frequency_series.mean()) if not frequency_series.empty else None

        results.append(
            {
                "campaign_id": campaign_id,
                "campaign_name": campaign_name,
                "objective": objective,
                "frequency": frequency,
                **totals,
            }
        )
    return results


def aggregate_daily_totals(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """Full-period, account-wide daily rollup for the dashboard's trend charts."""
    grouped = (
        df.groupby("date")
        .agg(
            spend=("spend", "sum"),
            purchase_value=("purchase_value", "sum"),
            purchases=("purchases", "sum"),
        )
        .reset_index()
        .sort_values("date")
    )

    results = []
    for _, row in grouped.iterrows():
        spend = float(row["spend"])
        purchase_value = float(row["purchase_value"])
        purchases = float(row["purchases"])
        results.append(
            {
                "date": row["date"],
                "spend": spend,
                "purchase_value": purchase_value,
                "roas": safe_divide(purchase_value, spend),
                "cpa": safe_divide(spend, purchases),
                "purchases": purchases,
            }
        )
    return results


# ---------------------------------------------------------------------------
# Styling helpers
# ---------------------------------------------------------------------------

def style_header_row(ws, row_idx: int, num_cols: int) -> None:
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")


def autosize_columns(ws, max_width: int = 32, min_width: int = 8) -> None:
    # Iterate by column index rather than cell.column_letter -- MergedCell
    # objects (non-anchor cells inside a merged range) don't expose
    # column_letter, but get_column_letter(cell.column) works for any cell.
    for col_cells in ws.columns:
        if not col_cells:
            continue
        col_letter = get_column_letter(col_cells[0].column)
        length = 0
        for cell in col_cells:
            if cell.value is not None:
                length = max(length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(length + 2, min_width), max_width)


# ---------------------------------------------------------------------------
# Sheet 1: Raw Data
# ---------------------------------------------------------------------------

def build_raw_data_sheet(wb: Workbook, df: pd.DataFrame, currency_code: Optional[str] = None):
    ws = wb.create_sheet("Raw Data")
    ws.append(RAW_DATA_COLUMNS)
    style_header_row(ws, 1, len(RAW_DATA_COLUMNS))

    for _, row in df.iterrows():
        values = []
        for col in RAW_DATA_COLUMNS:
            value = row[col]
            if col in RAW_DATA_DATE_COLUMNS:
                values.append(value)
            elif col in RAW_DATA_TEXT_COLUMNS:
                values.append("" if pd.isna(value) else str(value))
            else:
                values.append(None if pd.isna(value) else float(value))
        ws.append(values)

    last_row = ws.max_row
    col_index = {name: idx + 1 for idx, name in enumerate(RAW_DATA_COLUMNS)}

    for r in range(2, last_row + 1):
        ws.cell(row=r, column=col_index["date"]).number_format = "YYYY-MM-DD"
        ws.cell(row=r, column=col_index["campaign_id"]).number_format = "@"

    for col_name, fmt in RAW_DATA_NUMBER_FORMATS.items():
        c_idx = col_index[col_name]
        for r in range(2, last_row + 1):
            ws.cell(row=r, column=c_idx).number_format = fmt

    currency_format = get_currency_excel_number_format(currency_code)
    for col_name in RAW_DATA_CURRENCY_COLUMNS:
        c_idx = col_index[col_name]
        for r in range(2, last_row + 1):
            ws.cell(row=r, column=c_idx).number_format = currency_format

    ws.freeze_panes = "A2"

    last_col_letter = get_column_letter(len(RAW_DATA_COLUMNS))
    table = Table(displayName="RawDataTable", ref=f"A1:{last_col_letter}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(table)

    autosize_columns(ws, max_width=24)
    return ws


# ---------------------------------------------------------------------------
# Sheet 2: Executive Dashboard
# ---------------------------------------------------------------------------

def _kpi_number_format(format_type: str, currency_code: Optional[str]) -> str:
    if format_type == "currency":
        return get_currency_excel_number_format(currency_code)
    return {
        "integer": '#,##0',
        "decimal": '0.00',
        "true_fraction_percent": '0.0%',
    }[format_type]


def build_dashboard_header(ws, findings: dict) -> None:
    meta = findings["analysis_metadata"]
    confidence = findings["account_integrity"]["decision_confidence"]

    ws.cell(row=1, column=1, value="Meta Ads Performance Intelligence -- Executive Dashboard")
    ws.cell(row=1, column=1).font = Font(bold=True, size=16)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)

    subtitle = (
        f"Last 7 Days: {meta['last_7_day_window']['start']} to {meta['last_7_day_window']['end']}   |   "
        f"Previous 7 Days: {meta['previous_7_day_window']['start']} to {meta['previous_7_day_window']['end']}   |   "
        f"Account Decision Confidence: {confidence}"
    )
    ws.cell(row=2, column=1, value=subtitle)
    ws.cell(row=2, column=1).font = Font(italic=True, size=10)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)


def build_kpi_table(ws, account_summary: dict, start_row: int, currency_code: Optional[str] = None) -> int:
    last7 = account_summary["last_7_days"]
    previous7 = account_summary["previous_7_days"]
    changes = dict(account_summary["changes"])
    # compute_changes() (performance_analyzer.py) intentionally covers only 9
    # metrics -- it is not modified here. The 2 additional pct-changes the
    # dashboard needs (purchase_value, cpm) are derived the same way, locally,
    # using the same unmodified safe_pct_change() helper.
    changes["purchase_value_pct_change"] = safe_pct_change(last7.get("purchase_value"), previous7.get("purchase_value"))
    changes["cpm_pct_change"] = safe_pct_change(last7.get("cpm"), previous7.get("cpm"))

    headers = ["Metric", "Current 7 Days", "Previous 7 Days", "% Change"]
    for i, header in enumerate(headers, start=1):
        ws.cell(row=start_row, column=i, value=header)
    style_header_row(ws, start_row, len(headers))

    for offset, (label, metric_key, change_key, format_type, direction) in enumerate(KPI_DEFINITIONS, start=1):
        row = start_row + offset
        fmt = _kpi_number_format(format_type, currency_code)

        ws.cell(row=row, column=1, value=label)
        current_cell = ws.cell(row=row, column=2, value=last7.get(metric_key))
        previous_cell = ws.cell(row=row, column=3, value=previous7.get(metric_key))
        change_cell = ws.cell(row=row, column=4, value=changes.get(change_key))
        current_cell.number_format = fmt
        previous_cell.number_format = fmt
        change_cell.number_format = CHANGE_NUMBER_FORMAT

        if direction == "positive_increase":
            good_rule = CellIsRule(operator="greaterThan", formula=["0"], fill=GREEN_FILL)
            bad_rule = CellIsRule(operator="lessThan", formula=["0"], fill=RED_FILL)
        elif direction == "negative_increase":
            good_rule = CellIsRule(operator="lessThan", formula=["0"], fill=GREEN_FILL)
            bad_rule = CellIsRule(operator="greaterThan", formula=["0"], fill=RED_FILL)
        else:
            good_rule = bad_rule = None

        if good_rule is not None:
            ws.conditional_formatting.add(change_cell.coordinate, good_rule)
            ws.conditional_formatting.add(change_cell.coordinate, bad_rule)

    return start_row + len(KPI_DEFINITIONS)


def build_chart_data_area(
    ws,
    start_row: int,
    daily_totals: List[Dict[str, Any]],
    account_last7: dict,
    campaign_totals: List[Dict[str, Any]],
    currency_code: Optional[str] = None,
) -> Dict[str, int]:
    """Write the deterministic chart source data into its own column block.

    Every chart built from this workbook references these cells via
    openpyxl Reference() objects -- nothing is hardcoded into a chart.
    """
    col = CALC_COL_OFFSET
    row = start_row
    currency_format = get_currency_excel_number_format(currency_code)

    ws.cell(row=row, column=col, value="Chart Source Data (auto-generated -- do not edit)")
    ws.cell(row=row, column=col).font = Font(bold=True, size=12)
    row += 2

    daily_header_row = row
    for i, header in enumerate(["Date", "Spend", "Purchase Value", "ROAS", "CPA", "Purchases"]):
        ws.cell(row=row, column=col + i, value=header)
    style_header_row(ws, row, 6)
    row += 1
    daily_first_data_row = row
    for entry in daily_totals:
        ws.cell(row=row, column=col, value=entry["date"]).number_format = "YYYY-MM-DD"
        ws.cell(row=row, column=col + 1, value=entry["spend"]).number_format = currency_format
        ws.cell(row=row, column=col + 2, value=entry["purchase_value"]).number_format = currency_format
        ws.cell(row=row, column=col + 3, value=entry["roas"]).number_format = '0.00'
        ws.cell(row=row, column=col + 4, value=entry["cpa"]).number_format = currency_format
        ws.cell(row=row, column=col + 5, value=entry["purchases"]).number_format = '#,##0'
        row += 1
    daily_last_data_row = row - 1
    row += 2

    ws.cell(row=row, column=col, value="Conversion Funnel (Last 7 Days)")
    ws.cell(row=row, column=col).font = Font(bold=True, size=12)
    row += 1
    ws.cell(row=row, column=col, value="Stage")
    ws.cell(row=row, column=col + 1, value="Count")
    style_header_row(ws, row, 2)
    funnel_header_row = row
    row += 1
    funnel_first_row = row
    for label, value in (
        ("Clicks", account_last7.get("clicks")),
        ("Add to Cart", account_last7.get("add_to_carts")),
        ("Initiated Checkout", account_last7.get("initiated_checkouts")),
        ("Purchases", account_last7.get("purchases")),
    ):
        ws.cell(row=row, column=col, value=label)
        ws.cell(row=row, column=col + 1, value=value).number_format = '#,##0'
        row += 1
    funnel_last_row = row - 1
    row += 2

    ws.cell(row=row, column=col, value="Top Campaigns by Purchase Value")
    ws.cell(row=row, column=col).font = Font(bold=True, size=12)
    row += 1
    ws.cell(row=row, column=col, value="Campaign")
    ws.cell(row=row, column=col + 1, value="Purchase Value")
    style_header_row(ws, row, 2)
    top_value_header_row = row
    row += 1
    top_value_first_row = row
    top_by_value = sorted(campaign_totals, key=lambda c: (c["purchase_value"] or 0.0), reverse=True)[:10]
    for entry in top_by_value:
        ws.cell(row=row, column=col, value=entry["campaign_name"])
        ws.cell(row=row, column=col + 1, value=entry["purchase_value"]).number_format = currency_format
        row += 1
    top_value_last_row = row - 1
    row += 2

    ws.cell(row=row, column=col, value="Top Campaigns by ROAS")
    ws.cell(row=row, column=col).font = Font(bold=True, size=12)
    row += 1
    ws.cell(row=row, column=col, value="Campaign")
    ws.cell(row=row, column=col + 1, value="ROAS")
    style_header_row(ws, row, 2)
    top_roas_header_row = row
    row += 1
    top_roas_first_row = row
    ranked_by_roas = sorted(
        (c for c in campaign_totals if c["roas"] is not None), key=lambda c: c["roas"], reverse=True
    )[:10]
    for entry in ranked_by_roas:
        ws.cell(row=row, column=col, value=entry["campaign_name"])
        ws.cell(row=row, column=col + 1, value=entry["roas"]).number_format = '0.00'
        row += 1
    top_roas_last_row = row - 1

    return {
        "daily_header_row": daily_header_row,
        "daily_first_data_row": daily_first_data_row,
        "daily_last_data_row": daily_last_data_row,
        "funnel_header_row": funnel_header_row,
        "funnel_first_row": funnel_first_row,
        "funnel_last_row": funnel_last_row,
        "top_value_header_row": top_value_header_row,
        "top_value_first_row": top_value_first_row,
        "top_value_last_row": top_value_last_row,
        "top_roas_header_row": top_roas_header_row,
        "top_roas_first_row": top_roas_first_row,
        "top_roas_last_row": top_roas_last_row,
    }


def _add_line_chart(ws, title: str, y_title: str, data_col: int, layout: Dict[str, int], anchor: str) -> None:
    min_row = layout["daily_first_data_row"]
    max_row = layout["daily_last_data_row"]
    if max_row < min_row:
        return
    chart = LineChart()
    chart.title = title
    chart.y_axis.title = y_title
    chart.x_axis.title = "Date"
    data_ref = Reference(ws, min_col=data_col, max_col=data_col, min_row=layout["daily_header_row"], max_row=max_row)
    cats_ref = Reference(ws, min_col=CALC_COL_OFFSET, min_row=min_row, max_row=max_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.width = 24
    chart.height = 10
    ws.add_chart(chart, anchor)


def _add_spend_vs_revenue_chart(ws, layout: Dict[str, int], anchor: str) -> None:
    min_row = layout["daily_first_data_row"]
    max_row = layout["daily_last_data_row"]
    if max_row < min_row:
        return
    chart = LineChart()
    chart.title = "Daily Spend vs Purchase Value Trend"
    chart.y_axis.title = "Amount"
    chart.x_axis.title = "Date"
    data_ref = Reference(
        ws, min_col=CALC_COL_OFFSET + 1, max_col=CALC_COL_OFFSET + 2, min_row=layout["daily_header_row"], max_row=max_row
    )
    cats_ref = Reference(ws, min_col=CALC_COL_OFFSET, min_row=min_row, max_row=max_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.width = 26
    chart.height = 11
    ws.add_chart(chart, anchor)


def _add_bar_chart(
    ws, title: str, x_title: str, name_col: int, value_col: int, header_row: int, first_row: int, last_row: int, anchor: str
) -> None:
    if last_row < first_row:
        return
    chart = BarChart()
    chart.type = "bar"
    chart.title = title
    chart.x_axis.title = x_title
    data_ref = Reference(ws, min_col=value_col, max_col=value_col, min_row=header_row, max_row=last_row)
    cats_ref = Reference(ws, min_col=name_col, min_row=first_row, max_row=last_row)
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.width = 22
    chart.height = 10
    ws.add_chart(chart, anchor)


def build_executive_dashboard_sheet(
    wb: Workbook, findings: dict, daily_totals: List[Dict[str, Any]], campaign_totals: List[Dict[str, Any]]
):
    ws = wb.create_sheet("Executive Dashboard")
    currency_code = findings.get("analysis_metadata", {}).get("account_currency")

    build_dashboard_header(ws, findings)
    build_kpi_table(ws, findings["account_summary"], start_row=4, currency_code=currency_code)

    layout = build_chart_data_area(
        ws, start_row=1, daily_totals=daily_totals,
        account_last7=findings["account_summary"]["last_7_days"], campaign_totals=campaign_totals,
        currency_code=currency_code,
    )

    anchor_row = CHARTS_START_ROW
    _add_spend_vs_revenue_chart(ws, layout, f"A{anchor_row}")
    anchor_row += CHART_ROW_SPACING
    _add_line_chart(ws, "Daily ROAS Trend", "ROAS", CALC_COL_OFFSET + 3, layout, f"A{anchor_row}")
    anchor_row += CHART_ROW_SPACING
    _add_line_chart(ws, "Daily CPA Trend", "CPA", CALC_COL_OFFSET + 4, layout, f"A{anchor_row}")
    anchor_row += CHART_ROW_SPACING
    _add_line_chart(ws, "Daily Purchases Trend", "Purchases", CALC_COL_OFFSET + 5, layout, f"A{anchor_row}")
    anchor_row += CHART_ROW_SPACING
    _add_bar_chart(
        ws, "Conversion Funnel (Last 7 Days)", "Count",
        CALC_COL_OFFSET, CALC_COL_OFFSET + 1,
        layout["funnel_header_row"], layout["funnel_first_row"], layout["funnel_last_row"],
        f"A{anchor_row}",
    )
    anchor_row += CHART_ROW_SPACING
    _add_bar_chart(
        ws, "Top Campaigns by Purchase Value", "Purchase Value",
        CALC_COL_OFFSET, CALC_COL_OFFSET + 1,
        layout["top_value_header_row"], layout["top_value_first_row"], layout["top_value_last_row"],
        f"A{anchor_row}",
    )
    anchor_row += CHART_ROW_SPACING
    _add_bar_chart(
        ws, "Top Campaigns by ROAS", "ROAS",
        CALC_COL_OFFSET, CALC_COL_OFFSET + 1,
        layout["top_roas_header_row"], layout["top_roas_first_row"], layout["top_roas_last_row"],
        f"A{anchor_row}",
    )

    autosize_columns(ws, max_width=30)
    return ws


# ---------------------------------------------------------------------------
# Sheet 3: Campaign Performance
# ---------------------------------------------------------------------------

def build_campaign_performance_sheet(wb: Workbook, campaign_totals: List[Dict[str, Any]], findings: dict):
    ws = wb.create_sheet("Campaign Performance")
    ws.append(CAMPAIGN_PERFORMANCE_COLUMNS)
    style_header_row(ws, 1, len(CAMPAIGN_PERFORMANCE_COLUMNS))

    currency_code = findings.get("analysis_metadata", {}).get("account_currency")
    ai_status_by_campaign = {c["campaign_id"]: c["ai_status"] for c in findings.get("campaigns", [])}
    sorted_totals = sorted(campaign_totals, key=lambda c: c["spend"], reverse=True)

    for entry in sorted_totals:
        ai_status = ai_status_by_campaign.get(entry["campaign_id"], "MONITOR")
        ws.append(
            [
                entry["campaign_id"], entry["campaign_name"], entry["objective"],
                entry["spend"], entry["impressions"], entry["clicks"],
                entry["ctr"], entry["cpc"], entry["cpm"],
                entry["purchases"], entry["purchase_value"], entry["cpa"], entry["roas"],
                entry["add_to_carts"], entry["initiated_checkouts"], entry["checkout_to_purchase_rate"],
                entry["frequency"], ai_status,
            ]
        )

    last_row = ws.max_row
    col_index = {name: idx + 1 for idx, name in enumerate(CAMPAIGN_PERFORMANCE_COLUMNS)}

    for r in range(2, last_row + 1):
        ws.cell(row=r, column=col_index["Campaign ID"]).number_format = "@"
    for col_name, fmt in CAMPAIGN_PERFORMANCE_NUMBER_FORMATS.items():
        c_idx = col_index[col_name]
        for r in range(2, last_row + 1):
            ws.cell(row=r, column=c_idx).number_format = fmt

    currency_format = get_currency_excel_number_format(currency_code)
    for col_name in CAMPAIGN_PERFORMANCE_CURRENCY_COLUMNS:
        c_idx = col_index[col_name]
        for r in range(2, last_row + 1):
            ws.cell(row=r, column=c_idx).number_format = currency_format

    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(CAMPAIGN_PERFORMANCE_COLUMNS))
    table = Table(displayName="CampaignPerformanceTable", ref=f"A1:{last_col_letter}{last_row}")
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(table)

    if last_row >= 2:
        roas_col_letter = get_column_letter(col_index["ROAS"])
        roas_range = f"{roas_col_letter}2:{roas_col_letter}{last_row}"
        ws.conditional_formatting.add(roas_range, CellIsRule(operator="greaterThanOrEqual", formula=["3"], fill=GREEN_FILL))
        ws.conditional_formatting.add(roas_range, CellIsRule(operator="lessThan", formula=["1"], fill=RED_FILL))

        status_col_letter = get_column_letter(col_index["AI Status"])
        status_range = f"{status_col_letter}2:{status_col_letter}{last_row}"
        for status, fill in AI_STATUS_FILLS.items():
            ws.conditional_formatting.add(status_range, CellIsRule(operator="equal", formula=[f'"{status}"'], fill=fill))

    autosize_columns(ws, max_width=26)
    return ws


# ---------------------------------------------------------------------------
# Sheet 4: AI Analysis
# ---------------------------------------------------------------------------

def parse_ai_report_sections(markdown_text: str) -> Dict[str, str]:
    """Split Claude's markdown into an ordered dict of section name -> body text.

    Matches on '## SECTION NAME' headings (case-insensitive) against the
    required section list. Anything before the first recognized heading is
    treated as part of the Executive Summary so content is never dropped,
    even if Claude's formatting drifts slightly from the requested structure.
    """
    sections: Dict[str, List[str]] = {name: [] for name in REQUIRED_AI_ANALYSIS_SECTIONS}
    current = "EXECUTIVE SUMMARY"

    for line in markdown_text.splitlines():
        stripped = line.strip()
        if stripped.startswith("#"):
            heading_text = stripped.lstrip("#").strip().upper()
            matched = next((name for name in REQUIRED_AI_ANALYSIS_SECTIONS if heading_text == name), None)
            if matched:
                current = matched
                continue
            if stripped.lstrip("#").strip().startswith("Meta Ads Performance"):
                continue
        sections[current].append(line)

    return {name: "\n".join(body_lines).strip() for name, body_lines in sections.items()}


def parse_recommendation_blocks(section_text: str) -> List[Dict[str, str]]:
    """Split one section's body into labeled recommendation blocks.

    Blocks are separated by blank lines. A block that doesn't match the
    expected Priority/Decision Confidence/... field structure is kept as
    free text rather than discarded, so no content from Claude is ever lost.
    """
    if not section_text.strip():
        return []

    blocks: List[Dict[str, str]] = []
    for raw_block in re.split(r"\n\s*\n", section_text.strip()):
        fields: Dict[str, str] = {}
        unmatched_lines: List[str] = []
        for line in raw_block.splitlines():
            match = RECOMMENDATION_FIELD_PATTERN.match(line.strip())
            if match:
                field_name = match.group(1).strip()
                canonical = next(f for f in RECOMMENDATION_FIELDS if f.lower() == field_name.lower())
                fields[canonical] = match.group(2).strip()
            elif line.strip():
                unmatched_lines.append(line.strip())

        if fields:
            for field_name in RECOMMENDATION_FIELDS:
                fields.setdefault(field_name, "")
            if unmatched_lines:
                fields["Observation"] = (fields["Observation"] + " " + " ".join(unmatched_lines)).strip()
            blocks.append(fields)
        elif unmatched_lines:
            blocks.append({"free_text": " ".join(unmatched_lines)})

    return blocks


def build_ai_analysis_sheet(wb: Workbook, markdown_text: str):
    ws = wb.create_sheet("AI Analysis")
    sections = parse_ai_report_sections(markdown_text)

    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 95

    row = 1
    ws.cell(row=row, column=1, value="Meta Ads Performance Intelligence -- AI Analysis")
    ws.cell(row=row, column=1).font = Font(bold=True, size=16)
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    row += 2

    for section_name in REQUIRED_AI_ANALYSIS_SECTIONS:
        body = sections.get(section_name, "")

        header_cell = ws.cell(row=row, column=1, value=section_name)
        header_cell.font = Font(bold=True, size=13, color="FFFFFF")
        header_cell.fill = HEADER_FILL
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        ws.row_dimensions[row].height = 22
        row += 1

        if not body:
            ws.cell(row=row, column=1, value="No content generated for this section.")
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            row += 2
            continue

        if section_name in ACTION_SECTIONS:
            blocks = parse_recommendation_blocks(body)
            if not blocks:
                cell = ws.cell(row=row, column=1, value=body)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
                ws.row_dimensions[row].height = 60
                row += 2
                continue
            for block in blocks:
                if "free_text" in block:
                    cell = ws.cell(row=row, column=1, value=block["free_text"])
                    cell.alignment = Alignment(wrap_text=True, vertical="top")
                    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
                    ws.row_dimensions[row].height = 45
                    row += 1
                    continue
                for field_name in RECOMMENDATION_FIELDS:
                    label_cell = ws.cell(row=row, column=1, value=f"{field_name}:")
                    label_cell.font = Font(bold=True)
                    label_cell.alignment = Alignment(vertical="top")
                    value_cell = ws.cell(row=row, column=2, value=block.get(field_name, ""))
                    value_cell.alignment = Alignment(wrap_text=True, vertical="top")
                    ws.row_dimensions[row].height = 32
                    row += 1
                row += 1  # blank separator between recommendation blocks
        else:
            cell = ws.cell(row=row, column=1, value=body)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
            approx_lines = max(1, body.count("\n") + len(body) // 90)
            ws.row_dimensions[row].height = min(400, max(30, approx_lines * 15))
            row += 2

    ws.freeze_panes = "A1"
    return ws


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def generate_workbook(csv_path: str, findings_path: str, report_path: str, output_path: str) -> str:
    df = load_raw_data(csv_path)
    findings = load_findings(findings_path)
    markdown_text = load_ai_report(report_path)

    campaign_totals = aggregate_campaign_totals(df)
    daily_totals = aggregate_daily_totals(df)
    currency_code = findings.get("analysis_metadata", {}).get("account_currency")

    wb = Workbook()
    wb.remove(wb.active)

    build_raw_data_sheet(wb, df, currency_code=currency_code)
    build_executive_dashboard_sheet(wb, findings, daily_totals, campaign_totals)
    build_campaign_performance_sheet(wb, campaign_totals, findings)
    build_ai_analysis_sheet(wb, markdown_text)

    if list(wb.sheetnames) != REQUIRED_WORKSHEET_ORDER:
        raise ExcelReportError(
            f"Internal error: worksheet order {list(wb.sheetnames)} does not match "
            f"the required order {REQUIRED_WORKSHEET_ORDER}."
        )

    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    wb.save(output_path)
    return output_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate the Meta Ads Performance Intelligence Excel workbook.")
    parser.add_argument("--csv", default=DEFAULT_CSV_PATH, help="Path to meta_campaign_daily.csv.")
    parser.add_argument("--findings", default=DEFAULT_FINDINGS_PATH, help="Path to performance_findings.json.")
    parser.add_argument("--report", default=DEFAULT_REPORT_PATH, help="Path to meta_performance_report.md.")
    parser.add_argument("--output", default=DEFAULT_OUTPUT_PATH, help="Output .xlsx path.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    try:
        generate_workbook(args.csv, args.findings, args.report, args.output)
    except ExcelReportError as exc:
        print(f"Excel report error: {exc}")
        sys.exit(1)
    print(f"Saved Excel workbook to {args.output}")


if __name__ == "__main__":
    main()
