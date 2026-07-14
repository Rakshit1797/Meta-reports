#!/usr/bin/env python3
"""
Validation for the Meta Ads Performance Intelligence Excel workbook.

Runs after excel_report.py and before email delivery. Every check here is
structural/content-presence validation -- it never recalculates or second-
guesses any metric, it only confirms the workbook was built as expected.

If validation fails, this script exits with code 1 and prints a clear,
human-readable error for each failed check. No email is sent when this
script fails (send_email_report.py is only invoked after this succeeds).
"""

import argparse
import os
import sys
from typing import List, Optional, Tuple

import openpyxl

from excel_report import REQUIRED_WORKSHEET_ORDER

REQUIRED_CHART_TITLES = [
    "Daily Spend vs Purchase Value Trend",
    "Daily ROAS Trend",
    "Daily CPA Trend",
    "Daily Purchases Trend",
    "Conversion Funnel (Last 7 Days)",
    "Top Campaigns by Purchase Value",
    "Top Campaigns by ROAS",
]

DEFAULT_WORKBOOK_PATH = "meta_ads_performance_intelligence.xlsx"
DEFAULT_OTHER_OUTPUTS = [
    os.path.join("data", "meta_campaign_daily.csv"),
    os.path.join("data", "collection_metadata.json"),
    os.path.join("data", "performance_findings.json"),
    os.path.join("reports", "meta_performance_report.md"),
]


class ExcelValidationError(Exception):
    """Raised when the workbook (or a companion output file) fails validation."""


def _chart_title_text(chart) -> Optional[str]:
    try:
        return chart.title.tx.rich.p[0].r[0].t
    except Exception:
        return None


def _chart_series_references(chart) -> List[Optional[str]]:
    refs = []
    for series in chart.series:
        ref = None
        if series.val is not None and series.val.numRef is not None:
            ref = series.val.numRef.f
        refs.append(ref)
    return refs


def check_file_exists(path: str) -> Tuple[bool, str]:
    if not os.path.isfile(path):
        return False, f"Required output file '{path}' does not exist."
    return True, ""


def check_file_not_empty(path: str) -> Tuple[bool, str]:
    if os.path.getsize(path) == 0:
        return False, f"Required output file '{path}' is empty."
    return True, ""


def check_workbook_opens(path: str) -> Tuple[bool, str, Optional[openpyxl.Workbook]]:
    try:
        wb = openpyxl.load_workbook(path)
    except Exception as exc:
        return False, f"Workbook '{path}' could not be opened with openpyxl: {exc}", None
    return True, "", wb


def check_worksheet_set_and_order(wb: openpyxl.Workbook) -> List[str]:
    errors = []
    actual = list(wb.sheetnames)
    if set(actual) != set(REQUIRED_WORKSHEET_ORDER):
        errors.append(
            f"Workbook worksheets {actual} do not match the required set {REQUIRED_WORKSHEET_ORDER}."
        )
    elif actual != REQUIRED_WORKSHEET_ORDER:
        errors.append(
            f"Worksheet order is {actual}, but must be exactly {REQUIRED_WORKSHEET_ORDER}."
        )
    return errors


def check_raw_data_sheet(wb: openpyxl.Workbook) -> List[str]:
    errors = []
    if "Raw Data" not in wb.sheetnames:
        return ["'Raw Data' worksheet is missing."]
    ws = wb["Raw Data"]
    if ws.max_row < 2:
        errors.append("'Raw Data' worksheet contains no data rows (only a header, or is empty).")
    return errors


def check_executive_dashboard_sheet(wb: openpyxl.Workbook) -> List[str]:
    errors = []
    if "Executive Dashboard" not in wb.sheetnames:
        return ["'Executive Dashboard' worksheet is missing."]
    ws = wb["Executive Dashboard"]

    kpi_values_found = False
    for row in ws.iter_rows(min_row=1, max_row=20, max_col=4):
        for cell in row:
            if isinstance(cell.value, (int, float)):
                kpi_values_found = True
                break
        if kpi_values_found:
            break
    if not kpi_values_found:
        errors.append("'Executive Dashboard' worksheet has no numeric KPI values in its summary area.")

    found_titles = {_chart_title_text(chart) for chart in ws._charts}
    missing_charts = [title for title in REQUIRED_CHART_TITLES if title not in found_titles]
    if missing_charts:
        errors.append(f"'Executive Dashboard' is missing required chart(s): {', '.join(missing_charts)}.")

    for chart in ws._charts:
        title = _chart_title_text(chart) or "<untitled chart>"
        refs = _chart_series_references(chart)
        if not refs:
            errors.append(f"Chart '{title}' has no data series.")
            continue
        for ref in refs:
            if not ref:
                errors.append(f"Chart '{title}' has a series with no worksheet cell reference (possible hardcoded data).")

    return errors


def check_campaign_performance_sheet(wb: openpyxl.Workbook) -> List[str]:
    errors = []
    if "Campaign Performance" not in wb.sheetnames:
        return ["'Campaign Performance' worksheet is missing."]
    ws = wb["Campaign Performance"]
    if ws.max_row < 2:
        errors.append("'Campaign Performance' worksheet contains no campaign rows.")
    return errors


def check_ai_analysis_sheet(wb: openpyxl.Workbook) -> List[str]:
    errors = []
    if "AI Analysis" not in wb.sheetnames:
        return ["'AI Analysis' worksheet is missing."]
    ws = wb["AI Analysis"]

    has_content = False
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=2):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip():
                has_content = True
                break
        if has_content:
            break
    if not has_content:
        errors.append("'AI Analysis' worksheet contains no report content.")
    return errors


def validate_workbook(workbook_path: str, other_output_paths: List[str]) -> None:
    """Run every validation check. Raises ExcelValidationError with all failures if any fail."""
    errors: List[str] = []

    ok, message = check_file_exists(workbook_path)
    if not ok:
        raise ExcelValidationError(message)
    ok, message = check_file_not_empty(workbook_path)
    if not ok:
        raise ExcelValidationError(message)

    for path in other_output_paths:
        ok, message = check_file_exists(path)
        if not ok:
            errors.append(message)
            continue
        ok, message = check_file_not_empty(path)
        if not ok:
            errors.append(message)

    ok, message, wb = check_workbook_opens(workbook_path)
    if not ok:
        raise ExcelValidationError(message)

    errors.extend(check_worksheet_set_and_order(wb))
    errors.extend(check_raw_data_sheet(wb))
    errors.extend(check_executive_dashboard_sheet(wb))
    errors.extend(check_campaign_performance_sheet(wb))
    errors.extend(check_ai_analysis_sheet(wb))

    if errors:
        raise ExcelValidationError("\n".join(f"  - {err}" for err in errors))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate the Meta Ads Performance Intelligence Excel workbook.")
    parser.add_argument("--workbook", default=DEFAULT_WORKBOOK_PATH, help="Path to the .xlsx workbook to validate.")
    parser.add_argument(
        "--other-outputs",
        nargs="*",
        default=DEFAULT_OTHER_OUTPUTS,
        help="Other pipeline output files that must exist and be non-empty.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    try:
        validate_workbook(args.workbook, args.other_outputs)
    except ExcelValidationError as exc:
        print("Excel workbook validation FAILED:")
        print(str(exc))
        sys.exit(1)

    print(f"Excel workbook validation passed: '{args.workbook}' is well-formed.")


if __name__ == "__main__":
    main()
