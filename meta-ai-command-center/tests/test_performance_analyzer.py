#!/usr/bin/env python3
"""
Unit tests for the deterministic performance analysis rule engine.

These tests use hand-built synthetic metric dictionaries -- no live Meta API
data and no CSV files are read. Run with:

    python -m unittest discover -s tests -t . -v
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from performance_analyzer import (  # noqa: E402
    compute_changes,
    generate_findings_for_campaign,
    rule_cpa_deterioration,
    rule_creative_fatigue,
    rule_funnel_leakage,
    rule_roas_decline,
    rule_scaling_opportunity,
    rule_tracking_risk,
    rule_wasted_spend,
    safe_divide,
    safe_pct_change,
    sort_findings,
)


def metrics(spend=0.0, impressions=0.0, clicks=0.0, purchases=0.0, purchase_value=0.0,
            add_to_carts=0.0, initiated_checkouts=0.0):
    """Build a metrics dict the same way aggregate_window() would, from raw sums."""
    return {
        "spend": spend,
        "impressions": impressions,
        "clicks": clicks,
        "ctr": safe_divide(clicks, impressions),
        "cpc": safe_divide(spend, clicks),
        "cpm": (safe_divide(spend, impressions) * 1000) if safe_divide(spend, impressions) is not None else None,
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


class TestSafePercentChange(unittest.TestCase):
    def test_zero_previous_returns_none_not_infinity(self):
        self.assertIsNone(safe_pct_change(100.0, 0.0))

    def test_missing_values_return_none(self):
        self.assertIsNone(safe_pct_change(None, 10.0))
        self.assertIsNone(safe_pct_change(10.0, None))

    def test_normal_change(self):
        self.assertAlmostEqual(safe_pct_change(150.0, 100.0), 50.0)
        self.assertAlmostEqual(safe_pct_change(50.0, 100.0), -50.0)


class TestWastedSpendRule(unittest.TestCase):
    def test_triggers_on_high_spend_zero_purchases(self):
        last7 = metrics(spend=2500, impressions=5000, clicks=100, purchases=0, purchase_value=0)
        previous7 = metrics(spend=1000, impressions=4000, clicks=80, purchases=3, purchase_value=300)
        changes = compute_changes(last7, previous7)
        finding = rule_wasted_spend("C1", "Test", last7, previous7, changes)
        self.assertIsNotNone(finding)
        self.assertEqual(finding["severity"], "critical")
        self.assertEqual(finding["category"], "wasted_spend")

    def test_does_not_trigger_below_spend_threshold(self):
        last7 = metrics(spend=1999, impressions=5000, clicks=100, purchases=0, purchase_value=0)
        previous7 = metrics(spend=1000, impressions=4000, clicks=80, purchases=3, purchase_value=300)
        changes = compute_changes(last7, previous7)
        self.assertIsNone(rule_wasted_spend("C1", "Test", last7, previous7, changes))


class TestCpaDeteriorationRule(unittest.TestCase):
    def test_triggers_on_cpa_increase(self):
        last7 = metrics(spend=200, purchases=5, clicks=50, impressions=1000)
        previous7 = metrics(spend=100, purchases=5, clicks=50, impressions=1000)
        changes = compute_changes(last7, previous7)
        finding = rule_cpa_deterioration("C2", "Test", last7, previous7, changes)
        self.assertIsNotNone(finding)
        self.assertEqual(finding["severity"], "high")

    def test_does_not_trigger_with_low_purchase_volume(self):
        last7 = metrics(spend=200, purchases=2, clicks=50, impressions=1000)
        previous7 = metrics(spend=100, purchases=2, clicks=50, impressions=1000)
        changes = compute_changes(last7, previous7)
        self.assertIsNone(rule_cpa_deterioration("C2", "Test", last7, previous7, changes))


class TestRoasDeclineRule(unittest.TestCase):
    def test_triggers_on_roas_decline(self):
        last7 = metrics(spend=100, purchase_value=400)  # ROAS 4
        previous7 = metrics(spend=100, purchase_value=1000)  # ROAS 10
        changes = compute_changes(last7, previous7)
        finding = rule_roas_decline("C3", "Test", last7, previous7, changes)
        self.assertIsNotNone(finding)
        self.assertEqual(finding["severity"], "high")

    def test_no_trigger_when_previous_roas_missing(self):
        last7 = metrics(spend=100, purchase_value=400)
        previous7 = metrics(spend=0, purchase_value=0)  # ROAS undefined
        changes = compute_changes(last7, previous7)
        self.assertIsNone(rule_roas_decline("C3", "Test", last7, previous7, changes))


class TestFunnelLeakageRule(unittest.TestCase):
    def test_triggers_on_low_conversion_high_checkouts(self):
        last7 = metrics(initiated_checkouts=25, purchases=1)
        previous7 = metrics(initiated_checkouts=10, purchases=2)
        changes = compute_changes(last7, previous7)
        finding = rule_funnel_leakage("C4", "Test", last7, previous7, changes)
        self.assertIsNotNone(finding)
        self.assertEqual(finding["severity"], "high")

    def test_no_trigger_below_checkout_threshold(self):
        last7 = metrics(initiated_checkouts=10, purchases=0)
        previous7 = metrics(initiated_checkouts=10, purchases=2)
        changes = compute_changes(last7, previous7)
        self.assertIsNone(rule_funnel_leakage("C4", "Test", last7, previous7, changes))


class TestCreativeFatigueRule(unittest.TestCase):
    def test_triggers_on_spend_up_ctr_down(self):
        last7 = metrics(spend=150, clicks=60, impressions=2000)
        previous7 = metrics(spend=100, clicks=100, impressions=2000)
        changes = compute_changes(last7, previous7)
        finding = rule_creative_fatigue("C5", "Test", last7, previous7, changes)
        self.assertIsNotNone(finding)
        self.assertEqual(finding["severity"], "medium")
        self.assertIn("signal", finding["observation"].lower())
        self.assertNotIn("caused", finding["observation"].lower())

    def test_no_trigger_when_ctr_stable(self):
        last7 = metrics(spend=150, clicks=100, impressions=2000)
        previous7 = metrics(spend=100, clicks=100, impressions=2000)
        changes = compute_changes(last7, previous7)
        self.assertIsNone(rule_creative_fatigue("C5", "Test", last7, previous7, changes))


class TestScalingOpportunityRule(unittest.TestCase):
    def test_triggers_on_strong_performance(self):
        last7 = metrics(spend=100, purchases=6, purchase_value=1200)  # ROAS 12, CPA 16.67
        previous7 = metrics(spend=200, purchases=6, purchase_value=1200)  # CPA 33.33
        changes = compute_changes(last7, previous7)
        finding = rule_scaling_opportunity("C6", "Test", last7, previous7, changes)
        self.assertIsNotNone(finding)
        self.assertEqual(finding["severity"], "positive")

    def test_no_trigger_with_low_roas(self):
        last7 = metrics(spend=100, purchases=6, purchase_value=100)  # ROAS 1
        previous7 = metrics(spend=200, purchases=6, purchase_value=100)
        changes = compute_changes(last7, previous7)
        self.assertIsNone(rule_scaling_opportunity("C6", "Test", last7, previous7, changes))


class TestTrackingRiskRule(unittest.TestCase):
    def test_triggers_on_purchases_with_zero_value(self):
        last7 = metrics(purchases=3, purchase_value=0)
        previous7 = metrics(purchases=2, purchase_value=200)
        changes = compute_changes(last7, previous7)
        finding = rule_tracking_risk("C7", "Test", last7, previous7, changes)
        self.assertIsNotNone(finding)
        self.assertEqual(finding["severity"], "high")

    def test_no_trigger_when_no_purchases(self):
        last7 = metrics(purchases=0, purchase_value=0)
        previous7 = metrics(purchases=2, purchase_value=200)
        changes = compute_changes(last7, previous7)
        self.assertIsNone(rule_tracking_risk("C7", "Test", last7, previous7, changes))


class TestZeroPreviousPeriodSafety(unittest.TestCase):
    def test_all_changes_none_when_previous_period_is_empty(self):
        last7 = metrics(spend=100, purchases=5, purchase_value=500, clicks=50, impressions=1000,
                         add_to_carts=10, initiated_checkouts=5)
        previous7 = metrics()  # everything zero/undefined
        changes = compute_changes(last7, previous7)
        for metric_name, value in changes.items():
            self.assertIsNone(value, f"{metric_name} should be None, got {value}")

    def test_rule_engine_does_not_raise_with_empty_previous_period(self):
        last7 = metrics(spend=100, purchases=5, purchase_value=500, clicks=50, impressions=1000,
                         add_to_carts=10, initiated_checkouts=5)
        previous7 = metrics()
        changes = compute_changes(last7, previous7)
        # Should not raise TypeError from comparing None to a number.
        findings = generate_findings_for_campaign("C8", "Zero Previous", last7, previous7, changes)
        for finding in findings:
            for value in finding["evidence"].values():
                if isinstance(value, float):
                    self.assertNotIn(value, (float("inf"), float("-inf")))
                    self.assertFalse(value != value, "NaN found in evidence")  # NaN != NaN


class TestFindingsSorting(unittest.TestCase):
    def test_sorts_by_severity_priority(self):
        findings = [
            {"campaign_id": "A", "severity": "low"},
            {"campaign_id": "B", "severity": "critical"},
            {"campaign_id": "C", "severity": "high"},
            {"campaign_id": "D", "severity": "positive"},
            {"campaign_id": "E", "severity": "medium"},
        ]
        spend_by_campaign = {"A": 100, "B": 100, "C": 100, "D": 100, "E": 100}
        ordered = sort_findings(findings, spend_by_campaign)
        self.assertEqual(
            [f["severity"] for f in ordered],
            ["critical", "high", "medium", "positive", "low"],
        )

    def test_ties_broken_by_higher_spend_first(self):
        findings = [
            {"campaign_id": "LOW_SPEND", "severity": "high"},
            {"campaign_id": "HIGH_SPEND", "severity": "high"},
        ]
        spend_by_campaign = {"LOW_SPEND": 100, "HIGH_SPEND": 900}
        ordered = sort_findings(findings, spend_by_campaign)
        self.assertEqual([f["campaign_id"] for f in ordered], ["HIGH_SPEND", "LOW_SPEND"])


class TestNoDuplicateFindings(unittest.TestCase):
    def test_each_rule_fires_at_most_once_per_campaign(self):
        # Metrics designed to trigger several rules simultaneously for one campaign.
        last7 = metrics(spend=3000, purchases=6, purchase_value=1200, clicks=100,
                         impressions=2000, add_to_carts=30, initiated_checkouts=25)
        previous7 = metrics(spend=1500, purchases=6, purchase_value=1200, clicks=150,
                             impressions=2000, add_to_carts=30, initiated_checkouts=25)
        changes = compute_changes(last7, previous7)
        findings = generate_findings_for_campaign("DUP_TEST", "Duplicate Test", last7, previous7, changes)

        finding_ids = [f["finding_id"] for f in findings]
        self.assertEqual(len(finding_ids), len(set(finding_ids)), "duplicate finding_id produced")

        categories = [f["category"] for f in findings]
        self.assertEqual(len(categories), len(set(categories)), "same rule fired twice for one campaign")


if __name__ == "__main__":
    unittest.main()
