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

import pandas as pd

from performance_analyzer import (  # noqa: E402
    add_moving_averages,
    aggregate_daily_metrics,
    aggregate_monthly_totals,
    aggregate_weekly_totals,
    classify_decision_confidence,
    classify_portfolio_decision,
    compute_ai_status,
    compute_biggest_funnel_leak,
    compute_campaign_shares,
    compute_changes,
    compute_funnel_stages,
    compute_management_decisions,
    compute_top_management_signals,
    compute_top_sales_contributors,
    compute_wasted_spend_candidates,
    detect_integrity_warnings,
    generate_findings_for_campaign,
    get_collected_period,
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


class TestDataIntegrityAnomalyDetection(unittest.TestCase):
    def test_checkouts_exceeding_add_to_carts_is_flagged(self):
        last7 = metrics(initiated_checkouts=20, add_to_carts=5, clicks=100, impressions=2000)
        previous7 = metrics(initiated_checkouts=18, add_to_carts=6, clicks=100, impressions=2000)
        changes = compute_changes(last7, previous7)
        warnings = detect_integrity_warnings("C1", "Test", last7, previous7, changes)
        types = [w["type"] for w in warnings]
        self.assertIn("checkouts_exceed_add_to_carts", types)

    def test_conversion_rate_above_100_percent_is_flagged(self):
        last7 = metrics(purchases=10, initiated_checkouts=8)
        previous7 = metrics(purchases=5, initiated_checkouts=8)
        changes = compute_changes(last7, previous7)
        warnings = detect_integrity_warnings("C2", "Test", last7, previous7, changes)
        types = [w["type"] for w in warnings]
        self.assertIn("conversion_rate_above_100_percent", types)
        self.assertIn("purchases_exceed_checkouts", types)

    def test_extreme_discontinuity_is_flagged(self):
        last7 = metrics(spend=1000, purchases=5)
        previous7 = metrics(spend=100, purchases=5)
        changes = compute_changes(last7, previous7)
        warnings = detect_integrity_warnings("C3", "Test", last7, previous7, changes)
        types = [w["type"] for w in warnings]
        self.assertIn("extreme_discontinuity_spend", types)

    def test_healthy_campaign_produces_no_warnings(self):
        last7 = metrics(spend=100, impressions=2000, clicks=100, purchases=5, purchase_value=500,
                         add_to_carts=20, initiated_checkouts=10)
        previous7 = metrics(spend=100, impressions=2000, clicks=100, purchases=5, purchase_value=500,
                             add_to_carts=20, initiated_checkouts=10)
        changes = compute_changes(last7, previous7)
        warnings = detect_integrity_warnings("C4", "Test", last7, previous7, changes)
        self.assertEqual(warnings, [])

    def test_conflicting_roas_ctr_signal_is_flagged(self):
        last7 = metrics(spend=100, purchase_value=500, clicks=50, impressions=2000)  # ROAS 5, CTR 2.5%
        previous7 = metrics(spend=100, purchase_value=300, clicks=100, impressions=2000)  # ROAS 3, CTR 5%
        changes = compute_changes(last7, previous7)
        warnings = detect_integrity_warnings("C5", "Test", last7, previous7, changes)
        types = [w["type"] for w in warnings]
        self.assertIn("conflicting_roas_ctr_signal", types)

    def test_warnings_use_hedged_language(self):
        last7 = metrics(initiated_checkouts=20, add_to_carts=5)
        previous7 = metrics(initiated_checkouts=18, add_to_carts=6)
        changes = compute_changes(last7, previous7)
        warnings = detect_integrity_warnings("C6", "Test", last7, previous7, changes)
        hedge_phrases = ("may indicate", "requires validation", "cannot be confirmed")
        for warning in warnings:
            self.assertTrue(
                any(phrase in warning["message"].lower() for phrase in hedge_phrases),
                f"warning message not hedged: {warning['message']}",
            )
            self.assertNotIn("caused by", warning["message"].lower())


class TestDecisionConfidenceClassification(unittest.TestCase):
    def test_no_warnings_no_tracking_risk_good_sample_is_high(self):
        confidence = classify_decision_confidence([], tracking_risk_present=False, sample_size_ok=True)
        self.assertEqual(confidence, "HIGH")

    def test_tracking_risk_present_is_low(self):
        confidence = classify_decision_confidence([], tracking_risk_present=True, sample_size_ok=True)
        self.assertEqual(confidence, "LOW")

    def test_two_high_severity_warnings_is_low(self):
        warnings = [{"severity": "high"}, {"severity": "high"}]
        confidence = classify_decision_confidence(warnings, tracking_risk_present=False, sample_size_ok=True)
        self.assertEqual(confidence, "LOW")

    def test_one_warning_is_medium(self):
        warnings = [{"severity": "medium"}]
        confidence = classify_decision_confidence(warnings, tracking_risk_present=False, sample_size_ok=True)
        self.assertEqual(confidence, "MEDIUM")

    def test_small_sample_size_is_medium(self):
        confidence = classify_decision_confidence([], tracking_risk_present=False, sample_size_ok=False)
        self.assertEqual(confidence, "MEDIUM")


class TestAiStatusPriorityHierarchy(unittest.TestCase):
    def test_insufficient_data_overrides_everything(self):
        findings = [{"category": "scaling_opportunity"}, {"category": "tracking_risk"}]
        status = compute_ai_status(findings, has_sufficient_data=False)
        self.assertEqual(status, "INSUFFICIENT DATA")

    def test_no_findings_is_monitor(self):
        status = compute_ai_status([], has_sufficient_data=True)
        self.assertEqual(status, "MONITOR")

    def test_tracking_warning_outranks_all_others(self):
        findings = [
            {"category": "scaling_opportunity"},
            {"category": "creative_fatigue"},
            {"category": "cost_increase"},
            {"category": "tracking_risk"},
        ]
        status = compute_ai_status(findings, has_sufficient_data=True)
        self.assertEqual(status, "TRACKING WARNING")

    def test_efficiency_risk_outranks_creative_fatigue_and_scale(self):
        findings = [
            {"category": "scaling_opportunity"},
            {"category": "creative_fatigue"},
            {"category": "wasted_spend"},
        ]
        status = compute_ai_status(findings, has_sufficient_data=True)
        self.assertEqual(status, "EFFICIENCY RISK")

    def test_creative_fatigue_outranks_scale(self):
        findings = [{"category": "scaling_opportunity"}, {"category": "creative_fatigue"}]
        status = compute_ai_status(findings, has_sufficient_data=True)
        self.assertEqual(status, "CREATIVE FATIGUE")

    def test_scale_alone(self):
        findings = [{"category": "scaling_opportunity"}]
        status = compute_ai_status(findings, has_sufficient_data=True)
        self.assertEqual(status, "SCALE")

    def test_all_cost_increase_variants_map_to_efficiency_risk(self):
        for category in ("wasted_spend", "cost_increase", "performance_decline", "funnel_issue"):
            with self.subTest(category=category):
                status = compute_ai_status([{"category": category}], has_sufficient_data=True)
                self.assertEqual(status, "EFFICIENCY RISK")


class TestFunnelStages(unittest.TestCase):
    def test_funnel_calculations_are_sequential_conversion_rates(self):
        window_totals = {
            "impressions": 10000.0, "reach": 8000.0, "clicks": 400.0, "landing_page_views": 300.0,
            "add_to_carts": 60.0, "initiated_checkouts": 30.0, "purchases": 12.0,
        }
        stages = compute_funnel_stages(window_totals)
        by_key = {s["metric_key"]: s for s in stages}

        self.assertEqual(by_key["impressions"]["stage_conversion_rate"], None)  # first stage has no prior
        self.assertAlmostEqual(by_key["reach"]["stage_conversion_rate"], 8000.0 / 10000.0)
        self.assertAlmostEqual(by_key["clicks"]["stage_conversion_rate"], 400.0 / 8000.0)
        self.assertAlmostEqual(by_key["landing_page_views"]["stage_conversion_rate"], 300.0 / 400.0)
        self.assertAlmostEqual(by_key["add_to_carts"]["stage_conversion_rate"], 60.0 / 300.0)
        self.assertAlmostEqual(by_key["initiated_checkouts"]["stage_conversion_rate"], 30.0 / 60.0)
        self.assertAlmostEqual(by_key["purchases"]["stage_conversion_rate"], 12.0 / 30.0)
        # Every non-first stage documents its own formula.
        for key in ("reach", "clicks", "landing_page_views", "add_to_carts", "initiated_checkouts", "purchases"):
            self.assertTrue(by_key[key]["formula"])

    def test_funnel_stage_none_when_lpv_unavailable(self):
        window_totals = {
            "impressions": 10000.0, "reach": 8000.0, "clicks": 400.0, "landing_page_views": None,
            "add_to_carts": 60.0, "initiated_checkouts": 30.0, "purchases": 12.0,
        }
        stages = compute_funnel_stages(window_totals)
        by_key = {s["metric_key"]: s for s in stages}
        self.assertIsNone(by_key["landing_page_views"]["volume"])
        self.assertIsNone(by_key["landing_page_views"]["stage_conversion_rate"])
        # Clicks were never substituted for the missing LPV volume.
        self.assertNotEqual(by_key["landing_page_views"]["volume"], window_totals["clicks"])

    def test_biggest_funnel_leak_is_lowest_conversion_stage(self):
        current = compute_funnel_stages({
            "impressions": 10000.0, "reach": 9000.0, "clicks": 4500.0, "landing_page_views": 4000.0,
            "add_to_carts": 100.0, "initiated_checkouts": 90.0, "purchases": 80.0,
        })
        previous = compute_funnel_stages({
            "impressions": 10000.0, "reach": 9000.0, "clicks": 4500.0, "landing_page_views": 4000.0,
            "add_to_carts": 200.0, "initiated_checkouts": 180.0, "purchases": 160.0,
        })
        leak = compute_biggest_funnel_leak(current, previous)
        self.assertEqual(leak["stage"], "Add to Cart")  # 100/4000 = 2.5%, the weakest step
        self.assertIsNotNone(leak["movement_pct"])

    def test_tracking_anomaly_lpv_exceeds_clicks_is_flagged_as_tracking_not_performance(self):
        last7 = metrics(spend=1000.0, impressions=5000.0, clicks=100.0, purchases=5.0, purchase_value=500.0)
        last7["landing_page_views"] = 150.0  # 50% above clicks -- beyond tolerance
        previous7 = metrics(spend=1000.0, impressions=5000.0, clicks=100.0, purchases=5.0, purchase_value=500.0)
        previous7["landing_page_views"] = 100.0
        changes = compute_changes(last7, previous7)
        warnings = detect_integrity_warnings("C1", "Test", last7, previous7, changes)
        types = [w["type"] for w in warnings]
        self.assertIn("lpv_exceeds_clicks", types)
        warning = next(w for w in warnings if w["type"] == "lpv_exceeds_clicks")
        self.assertIn("tracking anomaly", warning["message"])
        self.assertNotIn("campaign performance", warning["message"].replace("not a campaign performance problem", ""))


class TestCampaignShares(unittest.TestCase):
    def test_spend_and_sales_share_computed_against_account_totals(self):
        campaign_totals = [
            {"spend": 300.0, "purchase_value": 900.0},
            {"spend": 700.0, "purchase_value": 100.0},
        ]
        account_totals = {"spend": 1000.0, "purchase_value": 1000.0}
        compute_campaign_shares(campaign_totals, account_totals)
        self.assertAlmostEqual(campaign_totals[0]["spend_share"], 0.3)
        self.assertAlmostEqual(campaign_totals[0]["sales_share"], 0.9)
        self.assertAlmostEqual(campaign_totals[1]["spend_share"], 0.7)
        self.assertAlmostEqual(campaign_totals[1]["sales_share"], 0.1)


class TestPortfolioDecisionHierarchy(unittest.TestCase):
    def test_scale_with_sufficient_confidence(self):
        decision = classify_portfolio_decision(spend=5000.0, purchases=10, roas=4.0, tracking_confidence="HIGH")
        self.assertEqual(decision, "SCALE")

    def test_scale_blocked_under_low_tracking_confidence(self):
        decision = classify_portfolio_decision(spend=5000.0, purchases=10, roas=4.0, tracking_confidence="LOW")
        self.assertEqual(decision, "INVESTIGATE TRACKING")
        self.assertNotEqual(decision, "SCALE")

    def test_cut_blocked_for_insufficient_sample(self):
        # Terrible ROAS, but spend is far too small to draw a CUT conclusion from.
        decision = classify_portfolio_decision(spend=50.0, purchases=0, roas=0.1, tracking_confidence="HIGH")
        self.assertEqual(decision, "INSUFFICIENT DATA")
        self.assertNotEqual(decision, "CUT")

    def test_cut_with_real_spend_and_poor_roas(self):
        decision = classify_portfolio_decision(spend=3000.0, purchases=2, roas=0.2, tracking_confidence="HIGH")
        self.assertEqual(decision, "CUT")

    def test_investigate_tracking_takes_priority_over_everything(self):
        # Even a campaign that looks like a strong scale candidate must be
        # INVESTIGATE TRACKING when tracking confidence is LOW.
        decision = classify_portfolio_decision(spend=100000.0, purchases=500, roas=10.0, tracking_confidence="LOW")
        self.assertEqual(decision, "INVESTIGATE TRACKING")


class TestWastedSpendAndTopContributors(unittest.TestCase):
    def test_wasted_spend_candidates_require_explicit_criteria(self):
        campaign_totals = [
            {"campaign_name": "A", "spend": 3000.0, "purchases": 0.0},
            {"campaign_name": "B", "spend": 3000.0, "purchases": 5.0},
            {"campaign_name": "C", "spend": 100.0, "purchases": 0.0},
        ]
        wasted = compute_wasted_spend_candidates(campaign_totals)
        names = [c["campaign_name"] for c in wasted]
        self.assertEqual(names, ["A"])  # B has purchases, C is below the spend threshold

    def test_top_sales_contributors_sorted_descending(self):
        campaign_totals = [
            {"campaign_name": "A", "purchase_value": 100.0},
            {"campaign_name": "B", "purchase_value": 500.0},
            {"campaign_name": "C", "purchase_value": 300.0},
        ]
        top = compute_top_sales_contributors(campaign_totals, max_count=2)
        self.assertEqual([c["campaign_name"] for c in top], ["B", "C"])


class TestManagementSignalsAndDecisions(unittest.TestCase):
    def test_management_signals_structure(self):
        findings = [
            {"campaign_id": "1", "title": "Wasted spend", "observation": "Spend with zero purchases."},
        ]
        campaign_summaries = [{"campaign_id": "1", "decision_confidence": "HIGH"}]
        signals = compute_top_management_signals(findings, campaign_summaries)
        self.assertEqual(len(signals), 1)
        self.assertEqual(set(signals[0].keys()), {"signal", "business_implication", "confidence"})
        self.assertEqual(signals[0]["confidence"], "HIGH")

    def test_management_decisions_structure_and_max_five(self):
        findings = [
            {"campaign_id": str(i), "severity": "high", "recommended_action": f"Action {i}",
             "observation": f"Evidence {i}", "title": f"Title {i}"}
            for i in range(8)
        ]
        campaign_summaries = [{"campaign_id": str(i), "decision_confidence": "MEDIUM"} for i in range(8)]
        decisions = compute_management_decisions(findings, campaign_summaries)
        self.assertEqual(len(decisions), 5)
        self.assertEqual(
            set(decisions[0].keys()), {"priority", "decision", "evidence", "commercial_implication", "confidence"}
        )


class TestMonthlyWeeklyDailyAggregation(unittest.TestCase):
    def _build_df(self, num_days=40):
        from datetime import date, timedelta
        rows = []
        start = date(2026, 1, 1)
        for offset in range(num_days):
            d = start + timedelta(days=offset)
            rows.append({
                "date": d, "campaign_id": "1", "campaign_name": "A", "objective": "X",
                "spend": 10.0, "impressions": 1000.0, "reach": 900.0, "clicks": 50.0,
                "purchases": 1.0, "purchase_value": 20.0, "add_to_carts": 5.0,
                "initiate_checkouts": 2.0, "landing_page_views": 30.0,
            })
        return pd.DataFrame(rows)

    def test_monthly_aggregation_groups_by_calendar_month(self):
        df = self._build_df(40)  # spans Jan and Feb 2026
        monthly = aggregate_monthly_totals(df)
        months = [m["month"] for m in monthly]
        self.assertEqual(months, sorted(months))
        self.assertIn("2026-01", months)
        self.assertIn("2026-02", months)
        jan = next(m for m in monthly if m["month"] == "2026-01")
        self.assertAlmostEqual(jan["spend"], 10.0 * 31)

    def test_weekly_aggregation_groups_by_monday_start_week(self):
        df = self._build_df(14)
        weekly = aggregate_weekly_totals(df)
        self.assertGreaterEqual(len(weekly), 2)
        for entry in weekly:
            self.assertIn("week_start", entry)
            self.assertIn("week_end", entry)

    def test_daily_metrics_one_row_per_day(self):
        df = self._build_df(10)
        daily = aggregate_daily_metrics(df)
        self.assertEqual(len(daily), 10)
        self.assertAlmostEqual(daily[0]["spend"], 10.0)

    def test_seven_day_moving_average_uses_available_days_only(self):
        df = self._build_df(10)
        daily = aggregate_daily_metrics(df)
        # Every day has roas = purchase_value/spend = 20/10 = 2.0, so the
        # moving average should also be exactly 2.0 throughout, including
        # the early days where fewer than 7 days are available.
        add_moving_averages(daily, ["roas"], window=7)
        for i, day in enumerate(daily):
            self.assertAlmostEqual(day["roas_7d_avg"], 2.0, msg=f"day index {i}")
        # The 3rd day's average is over exactly 3 real days, not padded/fabricated.
        add_moving_averages(daily, ["spend"], window=7)
        self.assertAlmostEqual(daily[2]["spend_7d_avg"], 10.0)


class TestCollectedPeriod(unittest.TestCase):
    def test_collected_period_is_the_csv_range_not_account_lifetime(self):
        from datetime import date
        import pandas as pd
        df = pd.DataFrame([
            {"date": date(2026, 3, 1)},
            {"date": date(2026, 3, 5)},
            {"date": date(2026, 3, 10)},
        ])
        period = get_collected_period(df)
        self.assertEqual(period["min_date"], date(2026, 3, 1))
        self.assertEqual(period["max_date"], date(2026, 3, 10))


if __name__ == "__main__":
    unittest.main()
