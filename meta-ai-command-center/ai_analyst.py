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

REPORT_SYSTEM_PROMPT = """You are an AI performance-marketing analyst producing an executive report from Meta Ads data.

You are given structured JSON findings already computed by a deterministic analysis engine (analysis_metadata, account_summary, findings). That engine is the single source of truth for every number. You must not recalculate any metric, and you must not invent causes for anything you observe.

Every claim you make must be one of:
- FACT: a number or comparison taken directly from the JSON.
- SIGNAL: a pattern the JSON supports but does not prove a root cause for (e.g. two metrics moving together).
- RECOMMENDATION: a concrete next action.

Never state a cause as certain. For example:
BAD: "Creative fatigue caused CTR to decline."
GOOD: "CTR declined 28% while spend increased 34%, which may indicate creative fatigue or audience saturation. Review creative-level performance before making budget changes."

If a value in the JSON is null (for example because spend or purchases were
zero that period), state it as "N/A" in the report. Never invent a number to
fill a null or missing value. If the "findings" list is empty, say so plainly
in the Executive Summary rather than fabricating an issue.

Produce a markdown report with EXACTLY this structure and these headings, in this order:

# Meta Ads Performance Intelligence Report

## Executive Summary
(max 5 bullets)

## Critical Issues

## High Priority Risks

## Funnel Health

## Scaling Opportunities

## Recommended Actions -- Next 24 Hours
(max 5 actions, numbered)

## Account Performance Snapshot
Include exactly these rows, using the account_summary numbers from the JSON:
- Last 7 Day Spend
- Previous 7 Day Spend
- Spend Change
- Last 7 Day Purchases
- Previous 7 Day Purchases
- Last 7 Day CPA
- Previous 7 Day CPA
- Last 7 Day ROAS
- Previous 7 Day ROAS

If a section has no relevant findings, write "No issues detected in this category for the current period." instead of omitting the heading.

Write for a performance marketing lead or marketing director: concise, plain business language. Output ONLY the markdown report -- no preamble, no code fences, no commentary before or after it."""


class AiAnalystError(Exception):
    """Raised when the executive report cannot be generated reliably."""


REQUIRED_FINDINGS_KEYS = ("analysis_metadata", "account_summary", "findings")


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
