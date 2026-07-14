#!/usr/bin/env python3
"""
Email delivery for the Meta Ads Performance Intelligence workbook.

Runs only after the Excel workbook has been generated and validated. Sends
the workbook as an attachment via the Resend API. Every number in the
email's executive snapshot is read directly from data/performance_findings.json
(management_signals and management_decisions are deterministic output from
performance_analyzer.py) -- the only AI-authored text in the email is the
Executive Verdict, taken verbatim (word-limited) from Claude's own markdown
report. This script performs no calculations of its own beyond simple
counting/formatting of already-computed values.

Required environment variables:
    RESEND_API_KEY     -- GitHub Actions secret (never logged or printed)
    REPORT_EMAIL_TO    -- GitHub Actions repository variable
    REPORT_EMAIL_FROM  -- GitHub Actions repository variable
"""

import argparse
import json
import os
import sys

import resend

from currency_utils import format_currency_text
from excel_report import NO_CONTENT_FALLBACK, limit_words, parse_ai_report_sections

DEFAULT_FINDINGS_PATH = os.path.join("data", "performance_findings.json")
DEFAULT_REPORT_PATH = os.path.join("reports", "meta_performance_report.md")
DEFAULT_ATTACHMENT_PATH = "meta_ads_performance_intelligence.xlsx"
ATTACHMENT_FILENAME = "meta_ads_performance_intelligence.xlsx"


class EmailDeliveryError(Exception):
    """Raised when the report email cannot be built or sent."""


def load_config() -> dict:
    """Read Resend credentials/addresses exclusively from environment variables."""
    api_key = os.getenv("RESEND_API_KEY")
    to_address = os.getenv("REPORT_EMAIL_TO")
    from_address = os.getenv("REPORT_EMAIL_FROM")

    missing = [
        name
        for name, value in (
            ("RESEND_API_KEY", api_key),
            ("REPORT_EMAIL_TO", to_address),
            ("REPORT_EMAIL_FROM", from_address),
        )
        if not value
    ]
    if missing:
        raise EmailDeliveryError(
            f"Missing required environment variable(s): {', '.join(missing)}. "
            "RESEND_API_KEY must be configured as a GitHub Actions repository "
            "secret; REPORT_EMAIL_TO and REPORT_EMAIL_FROM must be configured as "
            "GitHub Actions repository variables."
        )
    return {"api_key": api_key, "to": to_address, "from": from_address}


def load_findings(path: str) -> dict:
    try:
        with open(path) as findings_file:
            return json.load(findings_file)
    except FileNotFoundError as exc:
        raise EmailDeliveryError(
            f"Findings file '{path}' not found. Run performance_analyzer.py first."
        ) from exc
    except json.JSONDecodeError as exc:
        raise EmailDeliveryError(f"Findings file '{path}' is not valid JSON: {exc}") from exc


def load_ai_verdict(report_path: str) -> str:
    """Best-effort read of the AI markdown report's Executive Verdict section.

    Never raises -- if the report is missing or unreadable, the email falls
    back to the same deterministic "no material evidence" wording used
    everywhere else in the pipeline rather than fabricating a verdict.
    """
    try:
        with open(report_path) as report_file:
            markdown_text = report_file.read()
    except FileNotFoundError:
        return NO_CONTENT_FALLBACK
    sections = parse_ai_report_sections(markdown_text)
    verdict = sections.get("EXECUTIVE VERDICT", "").strip()
    return limit_words(verdict, 100) if verdict else NO_CONTENT_FALLBACK


def _fmt_ratio(value) -> str:
    return "N/A" if value is None else f"{value:.2f}"


def _fmt_count(value) -> str:
    return "N/A" if value is None else f"{value:,.0f}"


def build_snapshot(findings: dict) -> dict:
    """Pull the email's executive snapshot from already-computed deterministic values.

    No metric here is calculated by this script -- every number traces back
    to performance_analyzer.py's output. Website Landings (LPV) is read only
    from last_7_days["landing_page_views"]; if that is null, it renders as
    "N/A", never a fabricated zero and never clicks.
    """
    meta = findings.get("analysis_metadata", {})
    last7 = findings.get("account_summary", {}).get("last_7_days", {})
    currency_code = meta.get("account_currency")

    reporting_period = (
        f"{meta.get('last_7_day_window', {}).get('start', 'N/A')} to "
        f"{meta.get('last_7_day_window', {}).get('end', 'N/A')}"
    )

    return {
        "reporting_period": reporting_period,
        "total_spend": format_currency_text(last7.get("spend"), currency_code),
        "purchase_value": format_currency_text(last7.get("purchase_value"), currency_code),
        "purchases": _fmt_count(last7.get("purchases")),
        "roas": _fmt_ratio(last7.get("roas")),
        "cpa": format_currency_text(last7.get("cpa"), currency_code),
        "website_landings": _fmt_count(last7.get("landing_page_views")),
        "decision_confidence": findings.get("account_integrity", {}).get("decision_confidence", "N/A"),
    }


def build_subject(findings: dict) -> str:
    report_date = findings.get("analysis_metadata", {}).get("analysis_end_date", "")
    return f"Meta Ads Executive Performance Brief | {report_date}"


def build_html_body(snapshot: dict, verdict: str = "", signals=None, decisions=None) -> str:
    signals = signals or []
    decisions = decisions or []

    rows = [
        ("Reporting Period", snapshot["reporting_period"]),
        ("Spend", snapshot["total_spend"]),
        ("Tracked Sales", snapshot["purchase_value"]),
        ("ROAS", snapshot["roas"]),
        ("Purchases", snapshot["purchases"]),
        ("CPA", snapshot["cpa"]),
        ("Website Landings", snapshot["website_landings"]),
        ("Decision Confidence", snapshot["decision_confidence"]),
    ]
    rows_html = "".join(
        f'<tr>'
        f'<td style="padding:8px 16px;border-bottom:1px solid #e5e5e5;color:#444;font-family:Arial,sans-serif;">{label}</td>'
        f'<td style="padding:8px 16px;border-bottom:1px solid #e5e5e5;color:#111;font-weight:600;font-family:Arial,sans-serif;">{value}</td>'
        f'</tr>'
        for label, value in rows
    )

    signals_html = "".join(
        f'<li style="margin-bottom:8px;"><strong>{s.get("signal", "")}</strong> -- {s.get("business_implication", "")} '
        f'<em>(Confidence: {s.get("confidence", "N/A")})</em></li>'
        for s in signals[:3]
    ) or f'<li>{NO_CONTENT_FALLBACK}</li>'

    decisions_html = "".join(
        f'<li style="margin-bottom:8px;"><strong>[{d.get("priority", "").upper()}]</strong> {d.get("decision", "")} '
        f'<em>(Confidence: {d.get("confidence", "N/A")})</em></li>'
        for d in decisions[:5]
    ) or f'<li>{NO_CONTENT_FALLBACK}</li>'

    return f"""
<html>
  <body style="font-family:Arial,sans-serif;color:#222;">
    <h2 style="margin-bottom:4px;">Meta Ads Executive Performance Brief</h2>
    <table style="border-collapse:collapse;width:100%;max-width:480px;margin-top:12px;">
      {rows_html}
    </table>

    <h3 style="margin-top:24px;">Executive Verdict</h3>
    <p>{verdict}</p>

    <h3 style="margin-top:24px;">Top 3 Management Signals</h3>
    <ul>{signals_html}</ul>

    <h3 style="margin-top:24px;">Decisions Required This Week</h3>
    <ul>{decisions_html}</ul>

    <p style="margin-top:20px;">Full campaign portfolio, funnel analysis, trend analysis, AI management brief, and raw Meta data are attached in the Excel performance pack.</p>
  </body>
</html>
""".strip()


def build_attachment(workbook_path: str) -> dict:
    if not os.path.isfile(workbook_path):
        raise EmailDeliveryError(f"Workbook '{workbook_path}' does not exist -- cannot attach it to the email.")
    with open(workbook_path, "rb") as workbook_file:
        content = list(workbook_file.read())
    return {"filename": ATTACHMENT_FILENAME, "content": content}


def send_report_email(config: dict, findings: dict, workbook_path: str, report_path: str = DEFAULT_REPORT_PATH) -> str:
    """Send the report email via Resend. Returns the sent email's ID.

    Never logs or prints config["api_key"].
    """
    resend.api_key = config["api_key"]

    snapshot = build_snapshot(findings)
    verdict = load_ai_verdict(report_path)
    signals = findings.get("management_signals", [])
    decisions = findings.get("management_decisions", [])

    params = {
        "from": config["from"],
        "to": [config["to"]],
        "subject": build_subject(findings),
        "html": build_html_body(snapshot, verdict, signals, decisions),
        "attachments": [build_attachment(workbook_path)],
    }

    try:
        response = resend.Emails.send(params)
    except resend.exceptions.ResendError as exc:
        raise EmailDeliveryError(f"Resend API error ({exc.error_type}): {exc.message}") from exc
    except Exception as exc:  # network failures, etc. -- never include config["api_key"]
        raise EmailDeliveryError(f"Failed to send report email: {exc}") from exc

    return response.get("id", "") if isinstance(response, dict) else getattr(response, "id", "")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Email the generated Meta Ads Excel report via Resend.")
    parser.add_argument("--findings", default=DEFAULT_FINDINGS_PATH, help="Path to performance_findings.json.")
    parser.add_argument("--report", default=DEFAULT_REPORT_PATH, help="Path to meta_performance_report.md.")
    parser.add_argument("--workbook", default=DEFAULT_ATTACHMENT_PATH, help="Path to the Excel workbook to attach.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    try:
        config = load_config()
        findings = load_findings(args.findings)
        email_id = send_report_email(config, findings, args.workbook, args.report)
    except EmailDeliveryError as exc:
        print(f"Email delivery error: {exc}")
        sys.exit(1)

    print(f"Report email sent (id: {email_id}).")


if __name__ == "__main__":
    main()
