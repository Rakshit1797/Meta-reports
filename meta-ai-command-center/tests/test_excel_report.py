#!/usr/bin/env python3
"""
Unit tests for excel_report.py.

Builds small synthetic CSV/findings/report fixtures on disk (no live Meta
API or Anthropic API calls) and verifies the generated workbook's structure,
content, and chart wiring. Run with:

    python -m unittest discover -s tests -t . -v
"""

import json
import os
import shutil
import sys
import tempfile
import unittest
from datetime import date, timedelta

import openpyxl
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excel_report import (  # noqa: E402
    REQUIRED_WORKSHEET_ORDER,
    aggregate_campaign_totals,
    aggregate_daily_totals,
    generate_workbook,
    parse_ai_report_sections,
    parse_recommendation_blocks,
)
# REQUIRED_CHART_TITLES lives in validate_excel_report.py, not excel_report.py.
from validate_excel_report import REQUIRED_CHART_TITLES  # noqa: E402


def _build_synthetic_csv(path: str) -> None:
    rows = []
    end = date(2026, 7, 14)
    start = end - timedelta(days=13)

    def add_campaign(campaign_id, name, objective, last7_vals, previous7_vals):
        for offset in range(14):
            d = start + timedelta(days=offset)
            vals = last7_vals if offset >= 7 else previous7_vals
            rows.append(
                {
                    "date": d.isoformat(), "campaign_id": campaign_id, "campaign_name": name,
                    "objective": objective, **vals,
                }
            )

    healthy = {
        "spend": 100, "impressions": 3000, "reach": 2500, "clicks": 80, "ctr": 2.5, "cpc": 1.25,
        "cpm": 33.3, "frequency": 1.1, "purchases": 5, "purchase_value": 500,
        "add_to_carts": 20, "initiate_checkouts": 10, "landing_page_views": 50,
    }
    zero_purchase = {
        "spend": 300, "impressions": 4000, "reach": 3000, "clicks": 90, "ctr": 2.2, "cpc": 3.3,
        "cpm": 75, "frequency": 1.0, "purchases": 0, "purchase_value": 0,
        "add_to_carts": 5, "initiate_checkouts": 2, "landing_page_views": 40,
    }

    add_campaign("300000000001", "Steady Campaign", "OUTCOME_SALES", healthy, healthy)
    add_campaign("300000000002", "Zero Purchase Campaign", "OUTCOME_TRAFFIC", zero_purchase, healthy)

    df = pd.DataFrame(rows)

    def safe_div(n, d):
        return (n / d) if d else None

    df["CPA"] = df.apply(lambda r: safe_div(r["spend"], r["purchases"]), axis=1)
    df["ROAS"] = df.apply(lambda r: safe_div(r["purchase_value"], r["spend"]), axis=1)
    df["click_to_landing_page_rate"] = df.apply(lambda r: safe_div(r["landing_page_views"], r["clicks"]), axis=1)
    df["landing_page_to_purchase_rate"] = df.apply(lambda r: safe_div(r["purchases"], r["landing_page_views"]), axis=1)

    column_order = [
        "date", "campaign_id", "campaign_name", "objective", "spend", "impressions", "reach", "clicks",
        "ctr", "cpc", "cpm", "frequency", "purchases", "purchase_value", "CPA", "ROAS", "add_to_carts",
        "initiate_checkouts", "landing_page_views", "click_to_landing_page_rate", "landing_page_to_purchase_rate",
    ]
    df[column_order].to_csv(path, index=False)


SAMPLE_REPORT_MARKDOWN = """# Meta Ads Performance Intelligence Report

## EXECUTIVE SUMMARY
- Account performance was mixed over the last 7 days.

## DATA INTEGRITY & DECISION CONFIDENCE
FACT: No integrity warnings were detected this period.

## CRITICAL ISSUES

Priority: critical
Decision Confidence: HIGH
Campaign: Zero Purchase Campaign
Evidence: Spend of 2100.00 with 0 purchases in the last 7 days.
Observation: FACT: This campaign spent with zero purchases.
Recommended Action: Pause and review targeting immediately.

## HIGH PRIORITY RISKS
No issues detected in this category for the current period.

## SCALING OPPORTUNITIES
No issues detected in this category for the current period.

## CREATIVE FATIGUE SIGNALS
No issues detected in this category for the current period.

## FUNNEL ANALYSIS
FACT: Funnel metrics were within normal range account-wide.

## CAMPAIGNS TO WATCH
No issues detected in this category for the current period.

## NEXT 24 HOURS ACTION PLAN
1. Review the Zero Purchase Campaign immediately.
"""


class ExcelReportTestCase(unittest.TestCase):
    """Base class that builds a full synthetic pipeline fixture in a temp dir."""

    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.csv_path = os.path.join(self.tmp_dir, "meta_campaign_daily.csv")
        self.findings_path = os.path.join(self.tmp_dir, "performance_findings.json")
        self.report_path = os.path.join(self.tmp_dir, "meta_performance_report.md")
        self.output_path = os.path.join(self.tmp_dir, "meta_ads_performance_intelligence.xlsx")

        _build_synthetic_csv(self.csv_path)

        # Run the real (unmodified) performance_analyzer.py to produce findings --
        # no live API calls are involved, this is pure local computation.
        from performance_analyzer import analyze, load_dataset

        df = load_dataset(self.csv_path)
        results = analyze(df)
        with open(self.findings_path, "w") as f:
            json.dump(results, f, indent=2, default=str)

        with open(self.report_path, "w") as f:
            f.write(SAMPLE_REPORT_MARKDOWN)

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def set_account_currency(self, currency_code):
        """Rewrite the findings fixture's account_currency, then regenerate the workbook."""
        with open(self.findings_path) as f:
            findings = json.load(f)
        findings["analysis_metadata"]["account_currency"] = currency_code
        with open(self.findings_path, "w") as f:
            json.dump(findings, f, indent=2, default=str)
        generate_workbook(self.csv_path, self.findings_path, self.report_path, self.output_path)
        return openpyxl.load_workbook(self.output_path)


class TestWorkbookCreation(ExcelReportTestCase):
    def test_workbook_is_created_and_openable(self):
        generate_workbook(self.csv_path, self.findings_path, self.report_path, self.output_path)
        self.assertTrue(os.path.isfile(self.output_path))
        wb = openpyxl.load_workbook(self.output_path)
        self.assertIsNotNone(wb)


class TestWorksheetNamesAndOrder(ExcelReportTestCase):
    def test_exactly_four_worksheets_in_required_order(self):
        generate_workbook(self.csv_path, self.findings_path, self.report_path, self.output_path)
        wb = openpyxl.load_workbook(self.output_path)
        self.assertEqual(list(wb.sheetnames), REQUIRED_WORKSHEET_ORDER)
        self.assertEqual(len(wb.sheetnames), 4)


class TestRawDataPreservation(ExcelReportTestCase):
    def test_raw_data_row_count_matches_csv_exactly(self):
        generate_workbook(self.csv_path, self.findings_path, self.report_path, self.output_path)
        wb = openpyxl.load_workbook(self.output_path)
        ws = wb["Raw Data"]
        with open(self.csv_path) as csv_file:
            csv_row_count = sum(1 for _ in csv_file) - 1  # minus header
        self.assertEqual(ws.max_row - 1, csv_row_count)

    def test_raw_data_campaign_id_preserved_as_text(self):
        generate_workbook(self.csv_path, self.findings_path, self.report_path, self.output_path)
        wb = openpyxl.load_workbook(self.output_path)
        ws = wb["Raw Data"]
        header = [c.value for c in ws[1]]
        campaign_id_col = header.index("campaign_id") + 1
        value = ws.cell(row=2, column=campaign_id_col).value
        self.assertIsInstance(value, str)
        self.assertEqual(value, "300000000001")

    def test_raw_data_values_not_modified(self):
        generate_workbook(self.csv_path, self.findings_path, self.report_path, self.output_path)
        wb = openpyxl.load_workbook(self.output_path)
        ws = wb["Raw Data"]
        header = [c.value for c in ws[1]]
        spend_col = header.index("spend") + 1
        first_data_spend = ws.cell(row=2, column=spend_col).value
        df = pd.read_csv(self.csv_path)
        self.assertAlmostEqual(float(first_data_spend), float(df["spend"].iloc[0]))


class TestCampaignMetricAggregation(unittest.TestCase):
    def test_ctr_computed_from_totals_not_averaged(self):
        df = pd.DataFrame(
            [
                {"date": date(2026, 1, 1), "campaign_id": "1", "campaign_name": "A", "objective": "X",
                 "spend": 100.0, "impressions": 1000.0, "clicks": 100.0, "purchases": 1.0,
                 "purchase_value": 100.0, "add_to_carts": 5.0, "initiate_checkouts": 2.0, "frequency": 1.0},
                {"date": date(2026, 1, 2), "campaign_id": "1", "campaign_name": "A", "objective": "X",
                 "spend": 100.0, "impressions": 9000.0, "clicks": 90.0, "purchases": 1.0,
                 "purchase_value": 100.0, "add_to_carts": 5.0, "initiate_checkouts": 2.0, "frequency": 1.0},
            ]
        )
        totals = aggregate_campaign_totals(df)
        self.assertEqual(len(totals), 1)
        # total clicks = 190, total impressions = 10000 -> ctr = 0.019
        # (naive average of daily ratios would give (0.1 + 0.01)/2 = 0.055 -- wrong)
        self.assertAlmostEqual(totals[0]["ctr"], 190.0 / 10000.0)

    def test_daily_totals_aggregate_across_campaigns(self):
        df = pd.DataFrame(
            [
                {"date": date(2026, 1, 1), "campaign_id": "1", "campaign_name": "A", "objective": "X",
                 "spend": 100.0, "impressions": 1000.0, "clicks": 10.0, "purchases": 1.0,
                 "purchase_value": 100.0, "add_to_carts": 1.0, "initiate_checkouts": 1.0, "frequency": 1.0},
                {"date": date(2026, 1, 1), "campaign_id": "2", "campaign_name": "B", "objective": "X",
                 "spend": 50.0, "impressions": 500.0, "clicks": 5.0, "purchases": 1.0,
                 "purchase_value": 50.0, "add_to_carts": 1.0, "initiate_checkouts": 1.0, "frequency": 1.0},
            ]
        )
        daily = aggregate_daily_totals(df)
        self.assertEqual(len(daily), 1)
        self.assertAlmostEqual(daily[0]["spend"], 150.0)
        self.assertAlmostEqual(daily[0]["purchase_value"], 150.0)


class TestZeroDenominatorHandling(unittest.TestCase):
    def test_zero_spend_and_zero_purchases_yield_none_not_crash(self):
        df = pd.DataFrame(
            [
                {"date": date(2026, 1, 1), "campaign_id": "1", "campaign_name": "A", "objective": "X",
                 "spend": 0.0, "impressions": 0.0, "clicks": 0.0, "purchases": 0.0,
                 "purchase_value": 0.0, "add_to_carts": 0.0, "initiate_checkouts": 0.0, "frequency": None},
            ]
        )
        totals = aggregate_campaign_totals(df)
        self.assertEqual(len(totals), 1)
        entry = totals[0]
        self.assertIsNone(entry["ctr"])
        self.assertIsNone(entry["cpc"])
        self.assertIsNone(entry["roas"])
        self.assertIsNone(entry["cpa"])
        self.assertIsNone(entry["checkout_to_purchase_rate"])


class TestChartExistence(ExcelReportTestCase):
    def test_all_required_charts_exist_with_worksheet_references(self):
        generate_workbook(self.csv_path, self.findings_path, self.report_path, self.output_path)
        wb = openpyxl.load_workbook(self.output_path)
        ws = wb["Executive Dashboard"]

        found_titles = []
        for chart in ws._charts:
            try:
                found_titles.append(chart.title.tx.rich.p[0].r[0].t)
            except Exception:
                found_titles.append(None)

        for required_title in REQUIRED_CHART_TITLES:
            self.assertIn(required_title, found_titles)

        for chart in ws._charts:
            for series in chart.series:
                self.assertIsNotNone(series.val)
                self.assertIsNotNone(series.val.numRef)
                self.assertTrue(series.val.numRef.f)  # non-empty cell reference string


class TestCurrencyNumberFormats(ExcelReportTestCase):
    """Covers requirement: Excel currency formatting must dynamically use the
    account currency -- never hardcoded '$', and never assumed to be USD when
    currency metadata is missing."""

    RAW_DATA_CURRENCY_COLUMNS = ["spend", "purchase_value", "cpc", "cpm", "CPA"]
    CAMPAIGN_PERFORMANCE_CURRENCY_COLUMNS = ["Spend", "Purchase Value", "CPC", "CPM", "CPA"]

    def _raw_data_formats(self, wb):
        ws = wb["Raw Data"]
        header = [c.value for c in ws[1]]
        return {
            col: ws.cell(row=2, column=header.index(col) + 1).number_format
            for col in self.RAW_DATA_CURRENCY_COLUMNS
        }

    def _campaign_performance_formats(self, wb):
        ws = wb["Campaign Performance"]
        header = [c.value for c in ws[1]]
        return {
            col: ws.cell(row=2, column=header.index(col) + 1).number_format
            for col in self.CAMPAIGN_PERFORMANCE_CURRENCY_COLUMNS
        }

    def _kpi_currency_format(self, wb):
        # Executive Dashboard KPI table: header row 4, "Total Spend" row 5,
        # column 2 is the "Current 7 Days" value cell.
        ws = wb["Executive Dashboard"]
        return ws.cell(row=5, column=2).number_format

    def test_inr_uses_rupee_symbol_number_format(self):
        wb = self.set_account_currency("INR")
        for col, fmt in self._raw_data_formats(wb).items():
            self.assertEqual(fmt, '"₹"#,##0.00', f"Raw Data column {col} should use INR format")
        for col, fmt in self._campaign_performance_formats(wb).items():
            self.assertEqual(fmt, '"₹"#,##0.00', f"Campaign Performance column {col} should use INR format")
        self.assertEqual(self._kpi_currency_format(wb), '"₹"#,##0.00')

    def test_usd_uses_dollar_sign_number_format(self):
        wb = self.set_account_currency("USD")
        for col, fmt in self._raw_data_formats(wb).items():
            self.assertEqual(fmt, '"$"#,##0.00', f"Raw Data column {col} should use USD format")
        for col, fmt in self._campaign_performance_formats(wb).items():
            self.assertEqual(fmt, '"$"#,##0.00', f"Campaign Performance column {col} should use USD format")
        self.assertEqual(self._kpi_currency_format(wb), '"$"#,##0.00')

    def test_missing_currency_metadata_falls_back_without_assuming_usd(self):
        wb = self.set_account_currency(None)
        for col, fmt in self._raw_data_formats(wb).items():
            self.assertEqual(fmt, '#,##0.00', f"Raw Data column {col} should have no currency symbol")
            self.assertNotIn("$", fmt)
        for col, fmt in self._campaign_performance_formats(wb).items():
            self.assertEqual(fmt, '#,##0.00', f"Campaign Performance column {col} should have no currency symbol")
            self.assertNotIn("$", fmt)
        kpi_fmt = self._kpi_currency_format(wb)
        self.assertEqual(kpi_fmt, '#,##0.00')
        self.assertNotIn("$", kpi_fmt)

    def test_unknown_currency_code_appends_iso_code_not_dollar(self):
        wb = self.set_account_currency("AED")
        for col, fmt in self._raw_data_formats(wb).items():
            self.assertEqual(fmt, '#,##0.00" AED"', f"Raw Data column {col} should append ISO code")
            self.assertNotIn("$", fmt)
        for col, fmt in self._campaign_performance_formats(wb).items():
            self.assertEqual(fmt, '#,##0.00" AED"', f"Campaign Performance column {col} should append ISO code")
            self.assertNotIn("$", fmt)


class TestAiAnalysisParsing(unittest.TestCase):
    def test_sections_split_correctly(self):
        sections = parse_ai_report_sections(SAMPLE_REPORT_MARKDOWN)
        self.assertIn("EXECUTIVE SUMMARY", sections)
        self.assertIn("mixed", sections["EXECUTIVE SUMMARY"])
        self.assertIn("Zero Purchase Campaign", sections["CRITICAL ISSUES"])
        self.assertEqual(sections["HIGH PRIORITY RISKS"], "No issues detected in this category for the current period.")

    def test_recommendation_blocks_parsed_into_fields(self):
        sections = parse_ai_report_sections(SAMPLE_REPORT_MARKDOWN)
        blocks = parse_recommendation_blocks(sections["CRITICAL ISSUES"])
        self.assertEqual(len(blocks), 1)
        block = blocks[0]
        self.assertEqual(block["Priority"], "critical")
        self.assertEqual(block["Decision Confidence"], "HIGH")
        self.assertEqual(block["Campaign"], "Zero Purchase Campaign")
        self.assertIn("2100.00", block["Evidence"])

    def test_malformed_block_falls_back_to_free_text_without_dropping_content(self):
        blocks = parse_recommendation_blocks("Just some unstructured narrative text with no field labels.")
        self.assertEqual(len(blocks), 1)
        self.assertIn("free_text", blocks[0])
        self.assertIn("unstructured narrative", blocks[0]["free_text"])


if __name__ == "__main__":
    unittest.main()
