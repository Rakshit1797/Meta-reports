#!/usr/bin/env python3
"""
Unit tests for send_email_report.py.

Every Resend call is mocked -- no real email is ever sent, and no network
call is made. Run with:

    python -m unittest discover -s tests -t . -v
"""

import os
import sys
import tempfile
import unittest
from unittest import mock

import openpyxl
import yaml

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import send_email_report as ser  # noqa: E402


def _make_findings(account_currency, spend=1000.0, purchase_value=3000.0, cpa=50.0, landing_page_views=200.0):
    return {
        "analysis_metadata": {
            "analysis_end_date": "2026-07-14",
            "account_currency": account_currency,
            "last_7_day_window": {"start": "2026-07-08", "end": "2026-07-14"},
        },
        "account_summary": {
            "last_7_days": {
                "spend": spend, "purchase_value": purchase_value, "purchases": 20.0, "roas": 3.0, "cpa": cpa,
                "landing_page_views": landing_page_views,
            }
        },
        "account_integrity": {"decision_confidence": "HIGH"},
        "findings": [
            {"severity": "critical", "category": "wasted_spend"},
            {"severity": "positive", "category": "scaling_opportunity"},
            {"severity": "positive", "category": "scaling_opportunity"},
            {"severity": "high", "category": "tracking_risk"},
        ],
        "management_signals": [
            {"signal": "Wasted spend detected", "business_implication": "Budget is being lost.", "confidence": "HIGH"},
        ],
        "management_decisions": [
            {"priority": "critical", "decision": "Pause underperforming campaign", "evidence": "Zero purchases.",
             "commercial_implication": "Avoid further loss.", "confidence": "HIGH"},
        ],
    }


# USD is the explicit fixture for USD-currency tests -- it is never assumed
# as a default anywhere in send_email_report.py itself.
SAMPLE_FINDINGS = _make_findings("USD")


class TestSnapshotConstruction(unittest.TestCase):
    def test_snapshot_values_are_deterministic(self):
        snapshot = ser.build_snapshot(SAMPLE_FINDINGS)
        self.assertEqual(snapshot["decision_confidence"], "HIGH")
        self.assertEqual(snapshot["total_spend"], "$1,000.00")
        self.assertEqual(snapshot["purchase_value"], "$3,000.00")
        self.assertEqual(snapshot["purchases"], "20")
        self.assertEqual(snapshot["roas"], "3.00")
        self.assertEqual(snapshot["cpa"], "$50.00")
        self.assertEqual(snapshot["website_landings"], "200")
        self.assertEqual(snapshot["reporting_period"], "2026-07-08 to 2026-07-14")

    def test_null_metrics_render_as_na(self):
        findings = {
            "analysis_metadata": {"account_currency": "USD", "last_7_day_window": {"start": "2026-07-08", "end": "2026-07-14"}},
            "account_summary": {"last_7_days": {"spend": 0.0, "purchase_value": 0.0, "purchases": 0.0,
                                                   "roas": None, "cpa": None, "landing_page_views": None}},
            "account_integrity": {"decision_confidence": "LOW"},
            "findings": [],
        }
        snapshot = ser.build_snapshot(findings)
        self.assertEqual(snapshot["roas"], "N/A")
        self.assertEqual(snapshot["cpa"], "N/A")
        self.assertEqual(snapshot["website_landings"], "N/A")

    def test_subject_format(self):
        subject = ser.build_subject(SAMPLE_FINDINGS)
        self.assertEqual(subject, "Meta Ads Executive Performance Brief | 2026-07-14")

    def test_html_body_contains_required_sections_and_snapshot(self):
        snapshot = ser.build_snapshot(SAMPLE_FINDINGS)
        html = ser.build_html_body(
            snapshot, verdict="Account performance is stable.",
            signals=SAMPLE_FINDINGS["management_signals"], decisions=SAMPLE_FINDINGS["management_decisions"],
        )
        self.assertIn(
            "Full campaign portfolio, funnel analysis, trend analysis, AI management brief, and raw Meta "
            "data are attached in the Excel performance pack.",
            html,
        )
        self.assertIn("$1,000.00", html)
        self.assertIn("HIGH", html)
        self.assertIn("Executive Verdict", html)
        self.assertIn("Account performance is stable.", html)
        self.assertIn("Top 3 Management Signals", html)
        self.assertIn("Wasted spend detected", html)
        self.assertIn("Decisions Required This Week", html)
        self.assertIn("Pause underperforming campaign", html)

    def test_html_body_falls_back_to_no_evidence_wording_when_no_signals_or_decisions(self):
        snapshot = ser.build_snapshot(SAMPLE_FINDINGS)
        html = ser.build_html_body(snapshot, verdict="", signals=[], decisions=[])
        self.assertIn(ser.NO_CONTENT_FALLBACK, html)


class TestLpvEmailBehavior(unittest.TestCase):
    """Covers requirement: email must show LPV as N/A when unavailable, never a fabricated zero."""

    def test_website_landings_shows_na_when_unavailable(self):
        findings = _make_findings("USD", landing_page_views=None)
        snapshot = ser.build_snapshot(findings)
        self.assertEqual(snapshot["website_landings"], "N/A")

    def test_website_landings_shows_real_count_when_available(self):
        findings = _make_findings("USD", landing_page_views=350.0)
        snapshot = ser.build_snapshot(findings)
        self.assertEqual(snapshot["website_landings"], "350")

    def test_html_body_never_shows_zero_for_unavailable_lpv(self):
        findings = _make_findings("USD", landing_page_views=None)
        snapshot = ser.build_snapshot(findings)
        html = ser.build_html_body(snapshot)
        self.assertNotIn(">0<", html)
        self.assertIn("N/A", html)


class TestCurrencyFormatting(unittest.TestCase):
    """Covers requirement: never hardcode USD/'$' -- currency must come from
    the collected account_currency, INR must render with ₹, and a missing
    currency must never be silently assumed to be USD."""

    def test_usd_currency_uses_dollar_sign(self):
        snapshot = ser.build_snapshot(_make_findings("USD", spend=10282.30))
        self.assertEqual(snapshot["total_spend"], "$10,282.30")

    def test_inr_currency_uses_rupee_symbol(self):
        snapshot = ser.build_snapshot(_make_findings("INR", spend=10282.30, purchase_value=51000.0, cpa=250.5))
        self.assertEqual(snapshot["total_spend"], "₹10,282.30")
        self.assertEqual(snapshot["purchase_value"], "₹51,000.00")
        self.assertEqual(snapshot["cpa"], "₹250.50")
        self.assertNotIn("$", snapshot["total_spend"])

    def test_missing_currency_metadata_falls_back_without_assuming_usd(self):
        findings = _make_findings(None, spend=10282.30)
        snapshot = ser.build_snapshot(findings)
        self.assertEqual(snapshot["total_spend"], "10,282.30")
        self.assertNotIn("$", snapshot["total_spend"])

    def test_currency_code_without_known_symbol_appends_iso_code(self):
        snapshot = ser.build_snapshot(_make_findings("AED", spend=10282.30))
        self.assertEqual(snapshot["total_spend"], "10,282.30 AED")
        self.assertNotIn("$", snapshot["total_spend"])

    def test_email_html_body_uses_inr_symbol_not_dollar(self):
        snapshot = ser.build_snapshot(_make_findings("INR", spend=10282.30))
        html = ser.build_html_body(snapshot)
        self.assertIn("₹10,282.30", html)
        self.assertNotIn("$10,282.30", html)


class TestConfigValidation(unittest.TestCase):
    def test_missing_all_env_vars_raises_clear_error(self):
        with mock.patch.dict(os.environ, {}, clear=True):
            for var in ("RESEND_API_KEY", "REPORT_EMAIL_TO", "REPORT_EMAIL_FROM"):
                os.environ.pop(var, None)
            with self.assertRaises(ser.EmailDeliveryError) as ctx:
                ser.load_config()
            self.assertIn("RESEND_API_KEY", str(ctx.exception))
            self.assertIn("REPORT_EMAIL_TO", str(ctx.exception))
            self.assertIn("REPORT_EMAIL_FROM", str(ctx.exception))

    def test_all_env_vars_present_succeeds(self):
        env = {"RESEND_API_KEY": "sk-fake", "REPORT_EMAIL_TO": "to@example.com", "REPORT_EMAIL_FROM": "from@example.com"}
        with mock.patch.dict(os.environ, env, clear=True):
            config = ser.load_config()
            self.assertEqual(config["api_key"], "sk-fake")
            self.assertEqual(config["to"], "to@example.com")
            self.assertEqual(config["from"], "from@example.com")


class TestAiVerdictLoading(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()

    def test_missing_report_falls_back_without_raising(self):
        verdict = ser.load_ai_verdict(os.path.join(self.tmp_dir, "does_not_exist.md"))
        self.assertEqual(verdict, ser.NO_CONTENT_FALLBACK)

    def test_verdict_extracted_from_real_report(self):
        report_path = os.path.join(self.tmp_dir, "report.md")
        with open(report_path, "w") as f:
            f.write("## EXECUTIVE VERDICT\nAccount performance is winning this period.\n\n## WHAT MATERIALLY CHANGED\nSomething.\n")
        verdict = ser.load_ai_verdict(report_path)
        self.assertEqual(verdict, "Account performance is winning this period.")


class TestMockedEmailDelivery(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = tempfile.mkdtemp()
        self.workbook_path = os.path.join(self.tmp_dir, "meta_ads_performance_intelligence.xlsx")
        wb = openpyxl.Workbook()
        wb.save(self.workbook_path)
        self.report_path = os.path.join(self.tmp_dir, "meta_performance_report.md")
        with open(self.report_path, "w") as f:
            f.write("## EXECUTIVE VERDICT\nStable performance this period.\n")

    def test_send_report_email_attaches_workbook_without_live_call(self):
        config = {"api_key": "sk-fake", "to": "test@example.com", "from": "noreply@example.com"}
        with mock.patch("resend.Emails.send") as mock_send:
            mock_send.return_value = {"id": "mock-email-id"}
            email_id = ser.send_report_email(config, SAMPLE_FINDINGS, self.workbook_path, self.report_path)

        self.assertEqual(email_id, "mock-email-id")
        mock_send.assert_called_once()
        sent_params = mock_send.call_args[0][0]
        self.assertEqual(sent_params["to"], ["test@example.com"])
        self.assertEqual(sent_params["from"], "noreply@example.com")
        self.assertEqual(sent_params["subject"], "Meta Ads Executive Performance Brief | 2026-07-14")
        self.assertEqual(len(sent_params["attachments"]), 1)
        self.assertEqual(sent_params["attachments"][0]["filename"], "meta_ads_performance_intelligence.xlsx")
        self.assertIsInstance(sent_params["attachments"][0]["content"], list)

    def test_missing_workbook_raises_before_any_send_attempt(self):
        config = {"api_key": "sk-fake", "to": "test@example.com", "from": "noreply@example.com"}
        with mock.patch("resend.Emails.send") as mock_send:
            with self.assertRaises(ser.EmailDeliveryError):
                ser.send_report_email(config, SAMPLE_FINDINGS, os.path.join(self.tmp_dir, "does_not_exist.xlsx"), self.report_path)
            mock_send.assert_not_called()

    def test_resend_error_is_wrapped_clearly_and_never_leaks_api_key(self):
        import resend

        config = {"api_key": "sk-super-secret-value", "to": "test@example.com", "from": "noreply@example.com"}
        with mock.patch("resend.Emails.send") as mock_send:
            mock_send.side_effect = resend.exceptions.ResendError(
                code=401, error_type="authentication_error", message="Invalid API key",
                suggested_action="Check your key",
            )
            with self.assertRaises(ser.EmailDeliveryError) as ctx:
                ser.send_report_email(config, SAMPLE_FINDINGS, self.workbook_path, self.report_path)
            self.assertNotIn("sk-super-secret-value", str(ctx.exception))


class TestEmailNotSentWhenValidationFails(unittest.TestCase):
    """The pipeline-level guarantee that email only runs after Excel validation succeeds
    is enforced by GitHub Actions step ordering (default failure propagation), not by
    send_email_report.py itself. This test asserts that ordering is actually in place.
    """

    def test_workflow_runs_validation_before_email_and_email_has_no_always_condition(self):
        workflow_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
            ".github", "workflows", "meta-collector.yml",
        )
        with open(workflow_path) as f:
            workflow = yaml.safe_load(f)

        steps = workflow["jobs"]["collect"]["steps"]
        step_names = [s.get("name") for s in steps]

        validate_index = step_names.index("Validate Excel report")
        email_index = step_names.index("Send Excel report email")
        self.assertLess(
            validate_index, email_index,
            "'Validate Excel report' must run before 'Send Excel report email'",
        )

        email_step = steps[email_index]
        self.assertNotEqual(
            email_step.get("if"), "always()",
            "The email step must not use if: always() -- it should be skipped if an earlier step fails",
        )


if __name__ == "__main__":
    unittest.main()
