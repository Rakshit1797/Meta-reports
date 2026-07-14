#!/usr/bin/env python3
"""
Deterministic Meta Ads performance analysis engine.

Reads the cleaned campaign-level daily dataset produced by meta_collector.py
(data/meta_campaign_daily.csv) and produces structured findings comparing the
last 7 days against the previous 7 days, per campaign and for the account as
a whole. This engine is the sole source of truth for every number in the
findings JSON -- the AI layer (ai_analyst.py) only interprets these numbers,
it never recalculates them.

Column names are read exactly as produced by meta_collector.py's
OUTPUT_COLUMNS: date, campaign_id, campaign_name, spend, impressions, clicks,
purchases, purchase_value, add_to_carts, initiate_checkouts,
landing_page_views. This script does not assume any other column names.
"""

import argparse
import json
import os
import sys
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

import pandas as pd

DEFAULT_INPUT_PATH = os.path.join("data", "meta_campaign_daily.csv")
DEFAULT_OUTPUT_PATH = os.path.join("data", "performance_findings.json")
DEFAULT_METADATA_PATH = os.path.join("data", "collection_metadata.json")

WINDOW_LENGTH_DAYS = 7

# Rule thresholds
WASTED_SPEND_MIN_SPEND = 2000.0
CPA_DETERIORATION_MIN_PURCHASES = 3
CPA_DETERIORATION_MIN_INCREASE_PCT = 40.0
ROAS_DECLINE_MIN_DECREASE_PCT = 35.0
FUNNEL_LEAKAGE_MIN_CHECKOUTS = 20
FUNNEL_LEAKAGE_MAX_CONVERSION_RATE = 0.05
CREATIVE_FATIGUE_MIN_SPEND_INCREASE_PCT = 20.0
CREATIVE_FATIGUE_MIN_CTR_DECREASE_PCT = 20.0
SCALING_MIN_PURCHASES = 5
SCALING_MIN_ROAS = 2.0
SCALING_MIN_CPA_IMPROVEMENT_PCT = 20.0

SEVERITY_SORT_ORDER = {"critical": 0, "high": 1, "medium": 2, "positive": 3, "low": 4}

# --- Data integrity / decision confidence thresholds -----------------------
# These are additive to the rule engine above -- they do not change any of
# the 7 existing rules' thresholds or severities.
EXTREME_CHANGE_THRESHOLD_PCT = 300.0
TRAFFIC_SPEND_INCONSISTENCY_CHANGE_PCT = 100.0
TRAFFIC_SPEND_INCONSISTENCY_STABLE_SPEND_PCT = 10.0
PURCHASE_SURGE_THRESHOLD_PCT = 100.0
PURCHASE_SURGE_TRAFFIC_FLAT_PCT = 10.0
CONFLICTING_SIGNAL_CTR_DECLINE_PCT = 20.0
MIN_CLICKS_FOR_RELIABLE_SAMPLE = 50
MIN_IMPRESSIONS_FOR_RELIABLE_SAMPLE = 500

# AI Status priority hierarchy (highest priority first). A campaign's status
# is the highest-priority entry among the statuses implied by its findings.
AI_STATUS_PRIORITY = [
    "TRACKING WARNING",
    "EFFICIENCY RISK",
    "CREATIVE FATIGUE",
    "SCALE",
    "MONITOR",
    "INSUFFICIENT DATA",
]

# Maps an existing rule's category to the AI Status it implies. Existing
# rule categories/severities are unchanged -- this only adds a label.
CATEGORY_TO_AI_STATUS = {
    "tracking_risk": "TRACKING WARNING",
    "wasted_spend": "EFFICIENCY RISK",
    "cost_increase": "EFFICIENCY RISK",
    "performance_decline": "EFFICIENCY RISK",
    "funnel_issue": "EFFICIENCY RISK",
    "creative_fatigue": "CREATIVE FATIGUE",
    "scaling_opportunity": "SCALE",
}

REQUIRED_COLUMNS = [
    "date",
    "campaign_id",
    "campaign_name",
    "spend",
    "impressions",
    "clicks",
    "purchases",
    "purchase_value",
    "add_to_carts",
    "initiate_checkouts",
]


class AnalysisError(Exception):
    """Raised when the deterministic analysis cannot be completed reliably."""


# ---------------------------------------------------------------------------
# Safe math helpers
# ---------------------------------------------------------------------------

def safe_divide(numerator: float, denominator: float) -> Optional[float]:
    """Divide two numbers, returning None instead of raising or producing infinity."""
    if not denominator:
        return None
    return numerator / denominator


def safe_pct_change(current: Optional[float], previous: Optional[float]) -> Optional[float]:
    """Percentage change from previous to current.

    Returns None (never inf/NaN) when previous is zero or either value is
    missing -- a percentage change against a zero or undefined baseline is
    not meaningful.
    """
    if current is None or previous is None or not previous:
        return None
    return (current - previous) / abs(previous) * 100.0


# ---------------------------------------------------------------------------
# Loading + window aggregation
# ---------------------------------------------------------------------------

def load_dataset(path: str) -> pd.DataFrame:
    try:
        # campaign_id is an opaque identifier, never used arithmetically --
        # read it as text so it matches how the rest of the pipeline (and
        # Excel) treats it, and to avoid float64 precision loss on long
        # numeric IDs if pandas would otherwise infer a numeric dtype.
        df = pd.read_csv(path, dtype={"campaign_id": str})
    except FileNotFoundError as exc:
        raise AnalysisError(
            f"Input file '{path}' not found. Run meta_collector.py first."
        ) from exc

    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        raise AnalysisError(
            f"Input file '{path}' is missing expected column(s): {', '.join(missing_columns)}. "
            "This script does not guess column names -- confirm the collector's output schema."
        )

    if df.empty:
        raise AnalysisError(f"Input file '{path}' contains no rows.")

    df["date"] = pd.to_datetime(df["date"]).dt.date
    return df


def load_account_currency(metadata_path: str) -> Optional[str]:
    """Best-effort read of the account currency written by meta_collector.py.

    Never raises -- this is enrichment only. If the metadata file is
    missing, unreadable, or doesn't have the field, this returns None and
    downstream consumers must not assume a currency (e.g. USD).
    """
    try:
        with open(metadata_path) as metadata_file:
            metadata = json.load(metadata_file)
    except (FileNotFoundError, json.JSONDecodeError):
        return None
    return metadata.get("account_currency")


def get_analysis_windows(df: pd.DataFrame) -> Dict[str, Any]:
    """Determine the last-7-day and previous-7-day date windows from the dataset's max date."""
    analysis_end_date = df["date"].max()
    last_start = analysis_end_date - timedelta(days=WINDOW_LENGTH_DAYS - 1)
    previous_end = last_start - timedelta(days=1)
    previous_start = previous_end - timedelta(days=WINDOW_LENGTH_DAYS - 1)

    return {
        "analysis_end_date": analysis_end_date,
        "last_7_day_window": {"start": last_start, "end": analysis_end_date},
        "previous_7_day_window": {"start": previous_start, "end": previous_end},
    }


def aggregate_window(rows: pd.DataFrame, start, end) -> Dict[str, float]:
    """Sum raw metrics for rows within [start, end], then derive ratio metrics from the sums.

    Ratios (CTR, CPC, CPM, CPA, ROAS, funnel rates) are computed from summed
    totals, not averaged from daily per-row ratios -- averaging daily ratios
    would misweight low-volume days.
    """
    window_rows = rows[(rows["date"] >= start) & (rows["date"] <= end)]

    spend = float(window_rows["spend"].sum())
    impressions = float(window_rows["impressions"].sum())
    clicks = float(window_rows["clicks"].sum())
    purchases = float(window_rows["purchases"].sum())
    purchase_value = float(window_rows["purchase_value"].sum())
    add_to_carts = float(window_rows["add_to_carts"].sum())
    initiated_checkouts = float(window_rows["initiate_checkouts"].sum())

    spend_per_impression = safe_divide(spend, impressions)

    return {
        "spend": spend,
        "impressions": impressions,
        "clicks": clicks,
        "ctr": safe_divide(clicks, impressions),
        "cpc": safe_divide(spend, clicks),
        "cpm": spend_per_impression * 1000 if spend_per_impression is not None else None,
        "purchases": purchases,
        "purchase_value": purchase_value,
        "roas": safe_divide(purchase_value, spend),
        "cpa": safe_divide(spend, purchases),
        "add_to_carts": add_to_carts,
        "initiated_checkouts": initiated_checkouts,
        "atc_rate": safe_divide(add_to_carts, clicks),
        "checkout_rate": safe_divide(initiated_checkouts, add_to_carts),
        "checkout_to_purchase_rate": safe_divide(purchases, initiated_checkouts),
    }


def compute_changes(last7: Dict[str, float], previous7: Dict[str, float]) -> Dict[str, Optional[float]]:
    """Percentage change for the metrics the rule engine and report need."""
    return {
        "spend_pct_change": safe_pct_change(last7["spend"], previous7["spend"]),
        "ctr_pct_change": safe_pct_change(last7["ctr"], previous7["ctr"]),
        "cpc_pct_change": safe_pct_change(last7["cpc"], previous7["cpc"]),
        "purchases_pct_change": safe_pct_change(last7["purchases"], previous7["purchases"]),
        "cpa_pct_change": safe_pct_change(last7["cpa"], previous7["cpa"]),
        "roas_pct_change": safe_pct_change(last7["roas"], previous7["roas"]),
        "add_to_carts_pct_change": safe_pct_change(last7["add_to_carts"], previous7["add_to_carts"]),
        "initiated_checkouts_pct_change": safe_pct_change(
            last7["initiated_checkouts"], previous7["initiated_checkouts"]
        ),
        "checkout_to_purchase_rate_pct_change": safe_pct_change(
            last7["checkout_to_purchase_rate"], previous7["checkout_to_purchase_rate"]
        ),
    }


# ---------------------------------------------------------------------------
# Rule engine
# ---------------------------------------------------------------------------

def rule_wasted_spend(campaign_id, campaign_name, last7, previous7, changes):
    if last7["spend"] >= WASTED_SPEND_MIN_SPEND and last7["purchases"] == 0:
        return {
            "finding_id": f"wasted_spend_{campaign_id}",
            "campaign_id": campaign_id,
            "campaign_name": campaign_name,
            "severity": "critical",
            "category": "wasted_spend",
            "title": "Wasted spend with zero purchases",
            "observation": (
                f"This campaign spent {last7['spend']:.2f} over the last 7 days "
                "with zero canonical purchases recorded."
            ),
            "evidence": {
                "last_7_day_spend": last7["spend"],
                "last_7_day_purchases": last7["purchases"],
            },
            "recommended_action": (
                "Pause or urgently review this campaign's targeting, creative, and "
                "landing page before spending further."
            ),
        }
    return None


def rule_cpa_deterioration(campaign_id, campaign_name, last7, previous7, changes):
    cpa_change = changes["cpa_pct_change"]
    if (
        last7["purchases"] >= CPA_DETERIORATION_MIN_PURCHASES
        and previous7["purchases"] >= CPA_DETERIORATION_MIN_PURCHASES
        and cpa_change is not None
        and cpa_change >= CPA_DETERIORATION_MIN_INCREASE_PCT
    ):
        return {
            "finding_id": f"cpa_deterioration_{campaign_id}",
            "campaign_id": campaign_id,
            "campaign_name": campaign_name,
            "severity": "high",
            "category": "cost_increase",
            "title": "CPA increased significantly",
            "observation": (
                f"CPA increased {cpa_change:.1f}% (from {previous7['cpa']:.2f} to "
                f"{last7['cpa']:.2f}) while purchase volume stayed comparable "
                f"({previous7['purchases']:.0f} to {last7['purchases']:.0f})."
            ),
            "evidence": {
                "last_7_day_cpa": last7["cpa"],
                "previous_7_day_cpa": previous7["cpa"],
                "cpa_pct_change": cpa_change,
                "last_7_day_purchases": last7["purchases"],
                "previous_7_day_purchases": previous7["purchases"],
            },
            "recommended_action": (
                "Investigate bid strategy, audience saturation, and creative "
                "performance behind the CPA increase."
            ),
        }
    return None


def rule_roas_decline(campaign_id, campaign_name, last7, previous7, changes):
    roas_change = changes["roas_pct_change"]
    if (
        previous7["roas"] is not None
        and last7["roas"] is not None
        and roas_change is not None
        and roas_change <= -ROAS_DECLINE_MIN_DECREASE_PCT
    ):
        return {
            "finding_id": f"roas_decline_{campaign_id}",
            "campaign_id": campaign_id,
            "campaign_name": campaign_name,
            "severity": "high",
            "category": "performance_decline",
            "title": "ROAS declined",
            "observation": (
                f"ROAS declined {abs(roas_change):.1f}% (from {previous7['roas']:.2f} to "
                f"{last7['roas']:.2f})."
            ),
            "evidence": {
                "last_7_day_roas": last7["roas"],
                "previous_7_day_roas": previous7["roas"],
                "roas_pct_change": roas_change,
            },
            "recommended_action": (
                "Review recent creative, audience, and bidding changes; consider "
                "reallocating budget until ROAS recovers."
            ),
        }
    return None


def rule_funnel_leakage(campaign_id, campaign_name, last7, previous7, changes):
    conversion_rate = last7["checkout_to_purchase_rate"]
    if last7["initiated_checkouts"] >= FUNNEL_LEAKAGE_MIN_CHECKOUTS and (
        conversion_rate is None or conversion_rate < FUNNEL_LEAKAGE_MAX_CONVERSION_RATE
    ):
        rate_display = f"{conversion_rate * 100:.1f}%" if conversion_rate is not None else "0% (no purchases)"
        return {
            "finding_id": f"funnel_leakage_{campaign_id}",
            "campaign_id": campaign_id,
            "campaign_name": campaign_name,
            "severity": "high",
            "category": "funnel_issue",
            "title": "High checkout volume with low purchase conversion",
            "observation": (
                f"{last7['initiated_checkouts']:.0f} checkouts were initiated in the last "
                f"7 days, but only {rate_display} converted to a purchase."
            ),
            "evidence": {
                "last_7_day_initiated_checkouts": last7["initiated_checkouts"],
                "last_7_day_checkout_to_purchase_rate": conversion_rate,
            },
            "recommended_action": (
                "Audit the checkout flow, payment methods, and post-checkout tracking "
                "for drop-off causes."
            ),
        }
    return None


def rule_creative_fatigue(campaign_id, campaign_name, last7, previous7, changes):
    spend_change = changes["spend_pct_change"]
    ctr_change = changes["ctr_pct_change"]
    if (
        spend_change is not None
        and ctr_change is not None
        and spend_change >= CREATIVE_FATIGUE_MIN_SPEND_INCREASE_PCT
        and ctr_change <= -CREATIVE_FATIGUE_MIN_CTR_DECREASE_PCT
    ):
        return {
            "finding_id": f"creative_fatigue_{campaign_id}",
            "campaign_id": campaign_id,
            "campaign_name": campaign_name,
            "severity": "medium",
            "category": "creative_fatigue",
            "title": "Possible creative fatigue or audience saturation signal",
            "observation": (
                f"CTR declined {abs(ctr_change):.1f}% while spend increased "
                f"{spend_change:.1f}%. This is a possible signal of creative fatigue "
                "or audience saturation, not a confirmed cause."
            ),
            "evidence": {
                "spend_pct_change": spend_change,
                "ctr_pct_change": ctr_change,
            },
            "recommended_action": (
                "Review creative-level performance and refresh ad assets before making "
                "further budget changes."
            ),
        }
    return None


def rule_scaling_opportunity(campaign_id, campaign_name, last7, previous7, changes):
    cpa_change = changes["cpa_pct_change"]
    if (
        last7["purchases"] >= SCALING_MIN_PURCHASES
        and last7["roas"] is not None
        and last7["roas"] >= SCALING_MIN_ROAS
        and cpa_change is not None
        and cpa_change <= -SCALING_MIN_CPA_IMPROVEMENT_PCT
    ):
        return {
            "finding_id": f"scaling_opportunity_{campaign_id}",
            "campaign_id": campaign_id,
            "campaign_name": campaign_name,
            "severity": "positive",
            "category": "scaling_opportunity",
            "title": "Strong performance -- scaling opportunity",
            "observation": (
                f"{last7['purchases']:.0f} purchases at a {last7['roas']:.2f} ROAS in the "
                f"last 7 days, with CPA improving {abs(cpa_change):.1f}% versus the prior period."
            ),
            "evidence": {
                "last_7_day_purchases": last7["purchases"],
                "last_7_day_roas": last7["roas"],
                "cpa_pct_change": cpa_change,
            },
            "recommended_action": (
                "Consider incrementally increasing budget while monitoring CPA and ROAS "
                "for signs of diminishing returns."
            ),
        }
    return None


def rule_tracking_risk(campaign_id, campaign_name, last7, previous7, changes):
    if last7["purchases"] > 0 and last7["purchase_value"] == 0:
        return {
            "finding_id": f"tracking_risk_{campaign_id}",
            "campaign_id": campaign_id,
            "campaign_name": campaign_name,
            "severity": "high",
            "category": "tracking_risk",
            "title": "Purchases recorded with zero purchase value",
            "observation": (
                f"{last7['purchases']:.0f} purchase(s) were recorded in the last 7 days "
                "with a total purchase value of 0, suggesting a possible pixel or "
                "conversion-value tracking issue."
            ),
            "evidence": {
                "last_7_day_purchases": last7["purchases"],
                "last_7_day_purchase_value": last7["purchase_value"],
            },
            "recommended_action": (
                "Verify the Meta pixel / Conversions API purchase value parameter is "
                "firing correctly for this campaign."
            ),
        }
    return None


RULES = [
    rule_wasted_spend,
    rule_cpa_deterioration,
    rule_roas_decline,
    rule_funnel_leakage,
    rule_creative_fatigue,
    rule_scaling_opportunity,
    rule_tracking_risk,
]


def generate_findings_for_campaign(campaign_id, campaign_name, last7, previous7, changes) -> List[Dict[str, Any]]:
    findings = []
    for rule in RULES:
        finding = rule(campaign_id, campaign_name, last7, previous7, changes)
        if finding is not None:
            findings.append(finding)
    return findings


def sort_findings(findings: List[Dict[str, Any]], last7_spend_by_campaign: Dict[Any, float]) -> List[Dict[str, Any]]:
    return sorted(
        findings,
        key=lambda f: (
            SEVERITY_SORT_ORDER.get(f["severity"], len(SEVERITY_SORT_ORDER)),
            -last7_spend_by_campaign.get(f["campaign_id"], 0.0),
        ),
    )


# ---------------------------------------------------------------------------
# Data integrity, AI Status, and decision confidence
#
# This section is additive: it never changes the 7 rules above, their
# thresholds, or their severities. It only detects data-quality anomalies
# and derives a status/confidence label from the rule engine's own output,
# so both the AI analyst and the Excel report have a deterministic,
# auditable classification to work from instead of inventing one.
# ---------------------------------------------------------------------------

def detect_integrity_warnings(
    entity_id: str,
    entity_name: str,
    last7: Dict[str, Any],
    previous7: Dict[str, Any],
    changes: Dict[str, Optional[float]],
) -> List[Dict[str, Any]]:
    """Flag suspicious data patterns for one entity (a campaign or the whole account).

    Every warning uses hedged language ("may indicate", "possible signal",
    "requires validation") -- this function never asserts a root cause, only
    that something looks inconsistent and should be checked.
    """
    warnings: List[Dict[str, Any]] = []

    def add(warning_type: str, severity: str, message: str) -> None:
        warnings.append(
            {
                "type": warning_type,
                "entity_id": entity_id,
                "entity_name": entity_name,
                "severity": severity,
                "message": message,
            }
        )

    if last7["initiated_checkouts"] > last7["add_to_carts"]:
        add(
            "checkouts_exceed_add_to_carts",
            "high",
            f"Initiated checkouts ({last7['initiated_checkouts']:.0f}) exceed add-to-carts "
            f"({last7['add_to_carts']:.0f}) in the last 7 days. This may indicate an event-ordering "
            "or tracking anomaly and requires validation before drawing funnel conclusions.",
        )

    if last7["checkout_to_purchase_rate"] is not None and last7["checkout_to_purchase_rate"] > 1.0:
        add(
            "conversion_rate_above_100_percent",
            "high",
            f"Checkout-to-purchase rate is {last7['checkout_to_purchase_rate'] * 100:.1f}%, above 100%. "
            "This cannot be confirmed from available data and may indicate duplicate or "
            "misattributed purchase events.",
        )

    if last7["purchases"] > last7["initiated_checkouts"]:
        add(
            "purchases_exceed_checkouts",
            "high",
            f"Purchases ({last7['purchases']:.0f}) exceed initiated checkouts "
            f"({last7['initiated_checkouts']:.0f}) in the last 7 days, which requires validation "
            "of the checkout and purchase event tracking.",
        )

    for metric_key, metric_label, change_key in (
        ("spend", "spend", "spend_pct_change"),
        ("purchases", "purchases", "purchases_pct_change"),
    ):
        change = changes.get(change_key)
        if change is not None and abs(change) >= EXTREME_CHANGE_THRESHOLD_PCT:
            add(
                f"extreme_discontinuity_{metric_key}",
                "medium",
                f"{metric_label.capitalize()} changed {change:.1f}% versus the previous 7 days, "
                "an extreme period-over-period discontinuity that may indicate a data collection "
                "issue, a major account change, or a possible signal worth investigating.",
            )

    impressions_change = safe_pct_change(last7["impressions"], previous7["impressions"])
    clicks_change = safe_pct_change(last7["clicks"], previous7["clicks"])
    spend_change = changes.get("spend_pct_change")

    if (
        impressions_change is not None
        and spend_change is not None
        and abs(impressions_change) >= TRAFFIC_SPEND_INCONSISTENCY_CHANGE_PCT
        and abs(spend_change) < TRAFFIC_SPEND_INCONSISTENCY_STABLE_SPEND_PCT
    ):
        add(
            "impressions_spend_inconsistency",
            "medium",
            f"Impressions changed {impressions_change:.1f}% while spend changed only "
            f"{spend_change:.1f}%. This possible signal may indicate a delivery, auction, or "
            "tracking change and requires validation.",
        )

    if (
        clicks_change is not None
        and spend_change is not None
        and abs(clicks_change) >= TRAFFIC_SPEND_INCONSISTENCY_CHANGE_PCT
        and abs(spend_change) < TRAFFIC_SPEND_INCONSISTENCY_STABLE_SPEND_PCT
    ):
        add(
            "clicks_spend_inconsistency",
            "medium",
            f"Clicks changed {clicks_change:.1f}% while spend changed only {spend_change:.1f}%. "
            "This possible signal may indicate a creative, audience, or tracking change and "
            "requires validation.",
        )

    purchases_change = changes.get("purchases_pct_change")
    if (
        purchases_change is not None
        and purchases_change >= PURCHASE_SURGE_THRESHOLD_PCT
        and clicks_change is not None
        and clicks_change < PURCHASE_SURGE_TRAFFIC_FLAT_PCT
    ):
        add(
            "purchase_surge_without_traffic_growth",
            "medium",
            f"Purchases increased {purchases_change:.1f}% while clicks changed only "
            f"{clicks_change:.1f}%. This possible signal may indicate an attribution or "
            "tracking anomaly and cannot be confirmed from available data alone.",
        )

    roas_change = changes.get("roas_pct_change")
    ctr_change = changes.get("ctr_pct_change")
    if (
        roas_change is not None
        and roas_change > 0
        and ctr_change is not None
        and ctr_change <= -CONFLICTING_SIGNAL_CTR_DECLINE_PCT
    ):
        add(
            "conflicting_roas_ctr_signal",
            "low",
            f"ROAS improved {roas_change:.1f}% while CTR declined {abs(ctr_change):.1f}%. These "
            "conflicting signals may indicate early creative fatigue that current conversion "
            "efficiency is masking, and requires monitoring rather than a confident conclusion.",
        )

    return warnings


def classify_decision_confidence(
    integrity_warnings: List[Dict[str, Any]],
    tracking_risk_present: bool,
    sample_size_ok: bool,
) -> str:
    """Deterministically classify HIGH / MEDIUM / LOW confidence.

    - LOW: a tracking-risk finding is present, or 2+ high-severity integrity
      warnings exist -- decisions should not be made with high confidence.
    - MEDIUM: any integrity warning exists, or the sample size is too small
      to trust ratio metrics.
    - HIGH: no integrity warnings, a confirmed tracking issue, or sample
      size concern.
    """
    high_severity_warning_count = sum(1 for w in integrity_warnings if w["severity"] in ("high", "critical"))

    if tracking_risk_present or high_severity_warning_count >= 2:
        return "LOW"
    if integrity_warnings or not sample_size_ok:
        return "MEDIUM"
    return "HIGH"


def compute_ai_status(campaign_findings: List[Dict[str, Any]], has_sufficient_data: bool) -> str:
    """Derive a single AI Status from a campaign's findings using the priority hierarchy."""
    if not has_sufficient_data:
        return "INSUFFICIENT DATA"

    implied_statuses = {
        CATEGORY_TO_AI_STATUS[f["category"]]
        for f in campaign_findings
        if f["category"] in CATEGORY_TO_AI_STATUS
    }
    if not implied_statuses:
        return "MONITOR"

    for status in AI_STATUS_PRIORITY:
        if status in implied_statuses:
            return status
    return "MONITOR"


# ---------------------------------------------------------------------------
# Orchestration
# ---------------------------------------------------------------------------

def analyze(df: pd.DataFrame, account_currency: Optional[str] = None) -> Dict[str, Any]:
    windows = get_analysis_windows(df)
    last_start = windows["last_7_day_window"]["start"]
    last_end = windows["last_7_day_window"]["end"]
    previous_start = windows["previous_7_day_window"]["start"]
    previous_end = windows["previous_7_day_window"]["end"]

    all_findings: List[Dict[str, Any]] = []
    last7_spend_by_campaign: Dict[Any, float] = {}
    campaign_summaries: List[Dict[str, Any]] = []
    all_integrity_warnings: List[Dict[str, Any]] = []

    campaigns = df[["campaign_id", "campaign_name"]].drop_duplicates()

    for _, campaign_row in campaigns.iterrows():
        campaign_id = campaign_row["campaign_id"]
        campaign_name = campaign_row["campaign_name"]
        campaign_rows = df[df["campaign_id"] == campaign_id]

        last7 = aggregate_window(campaign_rows, last_start, last_end)
        previous7 = aggregate_window(campaign_rows, previous_start, previous_end)
        changes = compute_changes(last7, previous7)

        last7_spend_by_campaign[campaign_id] = last7["spend"]
        campaign_findings = generate_findings_for_campaign(campaign_id, campaign_name, last7, previous7, changes)
        all_findings.extend(campaign_findings)

        campaign_integrity_warnings = detect_integrity_warnings(
            campaign_id, campaign_name, last7, previous7, changes
        )
        all_integrity_warnings.extend(campaign_integrity_warnings)

        has_sufficient_data = (last7["impressions"] + previous7["impressions"]) > 0
        tracking_risk_present = any(f["category"] == "tracking_risk" for f in campaign_findings)
        sample_size_ok = (
            last7["clicks"] >= MIN_CLICKS_FOR_RELIABLE_SAMPLE
            and last7["impressions"] >= MIN_IMPRESSIONS_FOR_RELIABLE_SAMPLE
        )

        campaign_summaries.append(
            {
                "campaign_id": campaign_id,
                "campaign_name": campaign_name,
                "last_7_days": last7,
                "previous_7_days": previous7,
                "changes": changes,
                "ai_status": compute_ai_status(campaign_findings, has_sufficient_data),
                "decision_confidence": classify_decision_confidence(
                    campaign_integrity_warnings, tracking_risk_present, sample_size_ok
                ),
                "integrity_warnings": campaign_integrity_warnings,
            }
        )

    all_findings = sort_findings(all_findings, last7_spend_by_campaign)

    account_last7 = aggregate_window(df, last_start, last_end)
    account_previous7 = aggregate_window(df, previous_start, previous_end)
    account_changes = compute_changes(account_last7, account_previous7)

    account_integrity_warnings = detect_integrity_warnings(
        "ACCOUNT", "Account-Wide", account_last7, account_previous7, account_changes
    )
    account_tracking_risk_present = any(f["category"] == "tracking_risk" for f in all_findings)
    account_sample_size_ok = (
        account_last7["clicks"] >= MIN_CLICKS_FOR_RELIABLE_SAMPLE
        and account_last7["impressions"] >= MIN_IMPRESSIONS_FOR_RELIABLE_SAMPLE
    )
    account_decision_confidence = classify_decision_confidence(
        account_integrity_warnings, account_tracking_risk_present, account_sample_size_ok
    )

    return {
        "analysis_metadata": {
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "analysis_end_date": windows["analysis_end_date"].isoformat(),
            "last_7_day_window": {
                "start": last_start.isoformat(),
                "end": last_end.isoformat(),
            },
            "previous_7_day_window": {
                "start": previous_start.isoformat(),
                "end": previous_end.isoformat(),
            },
            "campaigns_analyzed": int(len(campaigns)),
            "findings_generated": len(all_findings),
            "integrity_warnings_generated": len(all_integrity_warnings) + len(account_integrity_warnings),
            "account_currency": account_currency,
        },
        "account_summary": {
            "last_7_days": account_last7,
            "previous_7_days": account_previous7,
            "changes": account_changes,
        },
        "account_integrity": {
            "decision_confidence": account_decision_confidence,
            "warnings": account_integrity_warnings,
        },
        "findings": all_findings,
        "campaigns": campaign_summaries,
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Run the deterministic Meta Ads performance analysis engine."
    )
    parser.add_argument("--input", default=DEFAULT_INPUT_PATH, help="Path to the cleaned campaign daily CSV.")
    parser.add_argument("--output", default=DEFAULT_OUTPUT_PATH, help="Path to write the findings JSON.")
    parser.add_argument(
        "--metadata",
        default=DEFAULT_METADATA_PATH,
        help="Path to collection_metadata.json (used only to read the account currency, if present).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    try:
        df = load_dataset(args.input)
        account_currency = load_account_currency(args.metadata)
        results = analyze(df, account_currency=account_currency)
    except AnalysisError as exc:
        print(f"Analysis error: {exc}")
        sys.exit(1)

    output_dir = os.path.dirname(args.output)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    with open(args.output, "w") as output_file:
        json.dump(results, output_file, indent=2, default=str)

    print(
        f"Analyzed {results['analysis_metadata']['campaigns_analyzed']} campaign(s), "
        f"generated {results['analysis_metadata']['findings_generated']} finding(s)."
    )
    print(f"Saved findings to {args.output}")


if __name__ == "__main__":
    main()
