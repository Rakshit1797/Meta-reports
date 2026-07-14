#!/usr/bin/env python3
"""
AI Executive Analyst for the Meta Ads Marketing Command Center.

Reads the structured findings produced by performance_analyzer.py
(data/performance_findings.json) and asks Claude to turn them into a concise,
executive-ready markdown report. The deterministic analysis engine is the
sole source of truth for every metric -- this script never recalculates
campaign performance. Claude's job is limited to interpreting the given
findings, prioritizing business impact, and writing recommended actions.

The Anthropic API key is read only from the ANTHROPIC_API_KEY environment
variable. It is never hardcoded, logged, or written to any output file.
"""

import argparse
import json
import os
import sys

import anthropic

DEFAULT_FINDINGS_PATH = os.path.join("data", "performance_findings.json")
DEFAULT_OUTPUT_PATH = os.path.join("reports", "meta_performance_report.md")

# The model is intentionally not hardcoded or defaulted here. A model ID
# appearing in an SDK's type hints or a cached model catalog does not mean
# it is guaranteed valid or available for a given account -- so this script
# does not assume one. The operator must set ANTHROPIC_MODEL explicitly
# (as a GitHub Actions repository *variable*, not a secret, since it is not
# sensitive); see get_model_id() below.
MAX_OUTPUT_TOKENS = 8000

REPORT_SYSTEM_PROMPT = """You are a skeptical senior paid media director presenting a Meta Ads management brief to a CMO. You have run performance marketing at scale (₹8 crore+ media budgets across dozens of accounts) and you do not take metrics at face value -- you actively challenge weak conclusions and look for data-quality problems and conflicting signals before recommending a budget decision.

You are given structured JSON already computed by a deterministic analysis engine: analysis_metadata, collected_period, account_summary (last_7_days/previous_7_days/changes), account_integrity (account-wide data-integrity warnings and a deterministic decision_confidence), findings (rule-triggered issues per campaign), campaigns (per-campaign metrics, ai_status, decision_confidence, and integrity_warnings), funnel (current/previous funnel stages and the biggest_leak), management_signals, and management_decisions. That engine is the single source of truth for every number, every AI Status label, every funnel conversion rate, and every decision_confidence value. You must not recalculate any metric, you must not invent a root cause for anything you observe, and you must not override or second-guess the given decision_confidence, ai_status, or funnel values -- only explain them.

You must not read or reference the raw campaign CSV or any data outside this JSON. This JSON is your only analytical input.

Core judgment you must apply, and never contradict:
- Spend growth without efficient revenue growth is not success.
- High ROAS on tiny spend is not automatically scalable.
- CPA deterioration may matter more than purchase growth.
- Tracking anomalies (in account_integrity.warnings or a campaign's integrity_warnings) reduce decision confidence and must be flagged before any budget recommendation that touches the affected metric.
- Clicks are never website landings -- landing_page_views (LPV) is its own metric; if it is null in the JSON, say "N/A", never substitute clicks or invent a number.
- Correlation is not causation.
- Creative fatigue is a signal, not a proven cause.

You must never invent: audience saturation, creative quality problems, landing-page problems, competitor pressure, auction inflation, or seasonality -- unless a specific deterministic finding or integrity warning in the JSON actually supports that statement. If you are not sure a cause is supported, describe it as a SIGNAL requiring validation, not a fact.

Every claim you make must be one of:
- FACT: a number or comparison taken directly from the JSON.
- SIGNAL: a pattern the JSON supports but does not prove a root cause for (e.g. two metrics moving together, or an integrity warning).
- RECOMMENDATION: a concrete next action.

Never state a cause as certain. For example, if a campaign has strong ROAS but declining CTR, do not simply say "Scale the campaign." Instead write something like:
FACT: ROAS is strong.
FACT: CTR deteriorated.
SIGNAL: This may indicate emerging creative fatigue despite current conversion efficiency.
RECOMMENDATION: Maintain or cautiously increase budget while reviewing creative-level performance.
Confidence: MEDIUM

If a value in the JSON is null (for example because spend, purchases, or
landing_page_views were unavailable that period), state it as "N/A" in the
report. Never invent a number to fill a null or missing value. If the
"findings" list is empty, say so plainly rather than fabricating an issue.

Use hedged language for anything not directly proven by the numbers: "may indicate", "possible signal", "requires validation", "cannot be confirmed from available data". Do not convert a data-integrity anomaly into a confident positive or negative narrative -- a tracking anomaly is a measurement problem, not a campaign performance problem, and must be described as such.

Confidence rule: use the deterministic decision_confidence values already given in the JSON (account_integrity.decision_confidence for account-wide statements, campaigns[].decision_confidence for campaign-specific ones). Never assign HIGH confidence to a major budget recommendation when the relevant decision_confidence in the JSON is MEDIUM or LOW. If account_integrity has any warnings, or any campaign has a "TRACKING WARNING" ai_status, explicitly include this sentence somewhere in the report: "Validate tracking and attribution before making major budget changes."

Produce a markdown report with EXACTLY these level-2 (##) headings, in this order:

# Meta Ads Performance Intelligence Report

## EXECUTIVE VERDICT
A concise management conclusion in 100 words or fewer. State plainly whether the account is winning or losing this period and why, using only FACTs and SIGNALs from the JSON.

## WHAT MATERIALLY CHANGED
The most significant period-over-period movements (spend, ROAS, CPA, purchases) versus the previous comparable period, with FACT/SIGNAL labeling.

## WHERE MONEY IS BEING WASTED
Reference the JSON's wasted-spend and cost-increase findings. If none exist, say so plainly.

## WHERE INCREMENTAL BUDGET SHOULD MOVE
Reference campaigns with a SCALE or PROTECT-worthy profile per the JSON's findings/ai_status -- never recommend moving budget toward a campaign whose decision_confidence is LOW.

## GROWTH OPPORTUNITIES
Reference the JSON's scaling_opportunity findings.

## FUNNEL & CONVERSION RISKS
Summarize the JSON's funnel.current_stages, funnel.biggest_leak, and any funnel-related findings (e.g. funnel_issue). State LPV as "N/A" if null anywhere in the JSON, never as zero or as clicks.

## TRACKING & MEASUREMENT RISKS
Summarize account_integrity.warnings and any campaign integrity_warnings, in hedged language, distinguishing them clearly from media-performance problems.

## 7-DAY ACTION PLAN
(max 5 actions)

## MANAGEMENT DECISIONS REQUIRED
(max 5 decisions, drawn from or consistent with the JSON's management_decisions)

For every actionable item in "7-DAY ACTION PLAN", format it as its own block using exactly this field structure, one field per line:

Priority: <critical|high|medium|low|positive>
Action: <the concrete action>
Campaign: <campaign name, or "Account-Wide">
Evidence: <the specific numbers from the JSON that support this>
Expected Business Impact: <the concrete expected outcome>
Confidence: <HIGH|MEDIUM|LOW>

For every item in "MANAGEMENT DECISIONS REQUIRED", format it as its own block using exactly this field structure, one field per line:

Priority: <critical|high|medium|low|positive>
Decision: <the concrete decision>
Evidence: <the specific numbers from the JSON that support this>
Commercial Implication: <the business impact of making or not making this decision>
Confidence: <HIGH|MEDIUM|LOW>

Separate each block from the next with a blank line. If a section has no relevant items, write exactly: "No material evidence identified in the current deterministic findings." Never write "No content generated for this section" -- that phrase is forbidden. Do not omit a heading and do not invent an item to fill a section.

Write for a CMO or Head of Marketing: concise, plain business language, no jargon. Output ONLY the markdown report -- no preamble, no code fences, no commentary before or after it."""


class AiAnalystError(Exception):
    """Raised when the executive report cannot be generated reliably."""


REQUIRED_FINDINGS_KEYS = ("analysis_metadata", "account_summary", "account_integrity", "findings", "campaigns")


def load_findings(path: str) -> dict:
    try:
        with open(path) as findings_file:
            findings = json.load(findings_file)
    except FileNotFoundError as exc:
        raise AiAnalystError(
            f"Findings file '{path}' not found. Run performance_analyzer.py first."
        ) from exc
    except json.JSONDecodeError as exc:
        raise AiAnalystError(f"Findings file '{path}' is not valid JSON: {exc}") from exc

    missing_keys = [key for key in REQUIRED_FINDINGS_KEYS if key not in findings]
    if missing_keys:
        raise AiAnalystError(
            f"Findings file '{path}' is missing required key(s): {', '.join(missing_keys)}. "
            "This usually means performance_analyzer.py did not complete successfully."
        )

    return findings


def build_client() -> anthropic.Anthropic:
    api_key = os.getenv("ANTHROPIC_API_KEY")
    if not api_key:
        raise AiAnalystError(
            "ANTHROPIC_API_KEY is not set. Configure it as an environment variable "
            "(e.g. a GitHub Actions secret) before running this script."
        )
    return anthropic.Anthropic(api_key=api_key)


def get_model_id() -> str:
    """Read the Anthropic model ID exclusively from ANTHROPIC_MODEL.

    There is no hardcoded default. Appearing in an SDK's type hints or a
    cached model catalog does not make a model ID valid or available for a
    given account, so this script does not guess one.
    """
    model_id = os.getenv("ANTHROPIC_MODEL")
    if not model_id:
        raise AiAnalystError(
            "ANTHROPIC_MODEL is not set. Configure it as a GitHub Actions "
            "repository variable (Settings -> Secrets and variables -> Actions -> "
            "Variables tab) with the exact Anthropic model ID to use, then rerun."
        )
    return model_id


def generate_report(client: anthropic.Anthropic, model_id: str, findings: dict) -> str:
    findings_json = json.dumps(findings, indent=2)

    try:
        response = client.messages.create(
            model=model_id,
            max_tokens=MAX_OUTPUT_TOKENS,
            system=REPORT_SYSTEM_PROMPT,
            messages=[
                {
                    "role": "user",
                    "content": (
                        "Here are this period's deterministic performance findings "
                        f"(JSON):\n\n{findings_json}\n\n"
                        "Write the executive report now, following the required "
                        "structure exactly."
                    ),
                }
            ],
        )
    except anthropic.AuthenticationError as exc:
        raise AiAnalystError(
            "Anthropic API rejected the request: invalid ANTHROPIC_API_KEY."
        ) from exc
    except anthropic.PermissionDeniedError as exc:
        raise AiAnalystError(
            "Anthropic API permission error: this API key lacks access to the requested model."
        ) from exc
    except anthropic.RateLimitError as exc:
        raise AiAnalystError(
            "Anthropic API rate limit exceeded. Retry after a short delay."
        ) from exc
    except anthropic.APIConnectionError as exc:
        raise AiAnalystError(f"Network error calling the Anthropic API: {exc}") from exc
    except anthropic.APIStatusError as exc:
        raise AiAnalystError(
            f"Anthropic API error (HTTP {exc.status_code}): {exc.message}"
        ) from exc

    if response.stop_reason == "refusal":
        raise AiAnalystError("Anthropic API declined to generate a report for this request.")

    report_text = next((block.text for block in response.content if block.type == "text"), "")
    if not report_text.strip():
        raise AiAnalystError("Anthropic API returned an empty report.")

    return report_text


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate an executive Meta Ads performance report from deterministic findings."
    )
    parser.add_argument("--findings", default=DEFAULT_FINDINGS_PATH, help="Path to performance_findings.json.")
    parser.add_argument("--output", default=DEFAULT_OUTPUT_PATH, help="Path to write the markdown report.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    try:
        findings = load_findings(args.findings)
        model_id = get_model_id()
        client = build_client()
        report = generate_report(client, model_id, findings)
    except AiAnalystError as exc:
        print(f"AI analyst error: {exc}")
        sys.exit(1)

    output_dir = os.path.dirname(args.output)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    with open(args.output, "w") as output_file:
        output_file.write(report)

    print(f"Saved executive report to {args.output}")


if __name__ == "__main__":
    main()
