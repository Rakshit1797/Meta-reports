#!/usr/bin/env python3
"""
Meta Ads campaign-level daily data collector.

Connects to the Meta Marketing API Ads Insights endpoint, pulls
campaign-level daily performance data for a configurable date range,
cleans up Meta's overlapping/duplicate conversion event types, computes
standard performance-marketing metrics (CPA, ROAS, funnel rates), and
writes a clean CSV to disk.

This script only performs data collection and cleaning. It does not do
any AI analysis, anomaly detection, or optimisation recommendations.
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime, timezone
from typing import Any, Dict, List, Optional

import pandas as pd
import requests
from dotenv import load_dotenv

load_dotenv()

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Meta Graph API version. Meta deprecates old versions on a rolling basis,
# so this is overridable via env var without touching the code.
GRAPH_API_VERSION = os.getenv("META_API_VERSION", "v21.0")
GRAPH_API_BASE_URL = f"https://graph.facebook.com/{GRAPH_API_VERSION}"

# Default historical date range for the initial collection run.
# Not hardcoded into the request logic -- these are only fallback defaults
# and can be overridden with --since / --until on the command line.
DEFAULT_SINCE = "2026-01-14"
DEFAULT_UNTIL = "2026-07-14"

DEFAULT_OUTPUT_PATH = os.path.join("data", "meta_campaign_daily.csv")

# Fields requested from the Ads Insights endpoint.
INSIGHTS_FIELDS = [
    "campaign_id",
    "campaign_name",
    "objective",
    "spend",
    "impressions",
    "reach",
    "clicks",
    "ctr",
    "cpc",
    "cpm",
    "frequency",
    "actions",
    "action_values",
    "date_start",
    "date_stop",
]

# Canonical Meta action_type values used to avoid double-counting.
# Meta's `actions` / `action_values` arrays contain several overlapping
# representations of the same conversion event (e.g. "purchase",
# "omni_purchase", "onsite_web_purchase", "offsite_conversion.fb_pixel_purchase").
# We deliberately read exactly one canonical action_type per metric instead
# of summing the array, otherwise the same conversion gets counted multiple
# times.
PURCHASE_ACTION_TYPE = "offsite_conversion.fb_pixel_purchase"
ADD_TO_CART_ACTION_TYPE = "offsite_conversion.fb_pixel_add_to_cart"
INITIATE_CHECKOUT_ACTION_TYPE = "offsite_conversion.fb_pixel_initiate_checkout"
LANDING_PAGE_VIEW_ACTION_TYPE = "landing_page_view"

OUTPUT_COLUMNS = [
    "date",
    "campaign_id",
    "campaign_name",
    "objective",
    "spend",
    "impressions",
    "reach",
    "clicks",
    "ctr",
    "cpc",
    "cpm",
    "frequency",
    "purchases",
    "purchase_value",
    "CPA",
    "ROAS",
    "add_to_carts",
    "initiate_checkouts",
    "landing_page_views",
    "click_to_landing_page_rate",
    "landing_page_to_purchase_rate",
]

REQUEST_TIMEOUT_SECONDS = 30
MAX_RETRIES = 3
RETRY_BACKOFF_SECONDS = 5

# Meta error codes that indicate a rate limit / throttling situation.
# 4  = application request limit reached
# 17 = user request limit reached
# 32 = page request limit reached
# 613 = custom rate limit (e.g. ads management/insights rate limit)
RATE_LIMIT_ERROR_CODES = {4, 17, 32, 613}


class MetaApiError(Exception):
    """Raised when the Meta Marketing API returns an error or an invalid response."""


# ---------------------------------------------------------------------------
# Credentials / configuration loading
# ---------------------------------------------------------------------------

def load_credentials() -> Dict[str, str]:
    """Load the Meta access token and ad account id from environment variables.

    Never hardcode credentials in code -- they must come from the
    environment (typically via a local .env file that is never committed).
    """
    access_token = os.getenv("META_ACCESS_TOKEN")
    ad_account_id = os.getenv("META_AD_ACCOUNT_ID")

    if not access_token:
        raise MetaApiError(
            "META_ACCESS_TOKEN is not set. Create a .env file (see .env.example) "
            "in the project root and set META_ACCESS_TOKEN before running this script."
        )
    if not ad_account_id:
        raise MetaApiError(
            "META_AD_ACCOUNT_ID is not set. Create a .env file (see .env.example) "
            "in the project root and set META_AD_ACCOUNT_ID before running this script."
        )

    if not ad_account_id.startswith("act_"):
        ad_account_id = f"act_{ad_account_id}"

    return {"access_token": access_token, "ad_account_id": ad_account_id}


# ---------------------------------------------------------------------------
# API request + pagination handling
# ---------------------------------------------------------------------------

def _raise_for_meta_error(error: Dict[str, Any]) -> None:
    """Translate a Meta API error object into a clear, human-readable exception.

    The access token itself is never included in the message.
    """
    code = error.get("code")
    message = error.get("message", "Unknown error")
    error_type = error.get("type", "")

    if code == 190:
        raise MetaApiError(
            "Meta access token is invalid or has expired (error code 190). "
            "Generate a new token in Meta Business Suite / Graph API Explorer "
            "and update META_ACCESS_TOKEN in your .env file."
        )
    if code in (10, 200, 299) or "permission" in message.lower():
        raise MetaApiError(
            f"Meta API permission error: {message}. "
            "Confirm your access token has the 'ads_read' permission for this ad account."
        )
    if code == 100:
        raise MetaApiError(f"Meta API rejected the request parameters: {message}")

    raise MetaApiError(f"Meta API error (code {code}, type '{error_type}'): {message}")


def _request_with_retries(url: str, params: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Perform a GET request against the Graph API with retries for transient failures."""
    last_exception: Optional[Exception] = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.get(url, params=params, timeout=REQUEST_TIMEOUT_SECONDS)
        except requests.exceptions.RequestException as exc:
            last_exception = exc
            print(f"Network error on attempt {attempt}/{MAX_RETRIES}: {exc}")
            time.sleep(RETRY_BACKOFF_SECONDS * attempt)
            continue

        try:
            response_json = response.json()
        except ValueError as exc:
            raise MetaApiError(
                f"Meta API returned a non-JSON response (HTTP {response.status_code}). "
                "This may indicate a temporary Meta outage or an invalid request."
            ) from exc

        if response.status_code == 200 and "error" not in response_json:
            return response_json

        error = response_json.get("error", {})
        error_code = error.get("code")

        if error_code in RATE_LIMIT_ERROR_CODES or response.status_code == 429:
            wait_seconds = RETRY_BACKOFF_SECONDS * attempt
            print(
                f"Rate limited by the Meta API. Waiting {wait_seconds}s before "
                f"retrying (attempt {attempt}/{MAX_RETRIES})..."
            )
            time.sleep(wait_seconds)
            continue

        if error:
            _raise_for_meta_error(error)

        raise MetaApiError(
            f"Meta API returned an unexpected HTTP {response.status_code} response."
        )

    raise MetaApiError(
        f"Failed to reach the Meta API after {MAX_RETRIES} attempts due to network "
        f"failures. Last error: {last_exception}"
    )


def fetch_campaign_insights(
    access_token: str,
    ad_account_id: str,
    since: str,
    until: str,
    page_limit: int = 500,
) -> "tuple[List[Dict[str, Any]], int]":
    """Fetch every page of campaign-level daily insights for the given date range.

    Follows paging.next until Meta reports no further pages, combining all
    results into a single list of raw records. Returns (records, pages_fetched)
    -- pagination is only "complete" once this function returns without
    raising, since the while loop only exits when paging.next is absent.
    """
    url = f"{GRAPH_API_BASE_URL}/{ad_account_id}/insights"
    params: Optional[Dict[str, Any]] = {
        "level": "campaign",
        "time_increment": 1,
        "fields": ",".join(INSIGHTS_FIELDS),
        "time_range": json.dumps({"since": since, "until": until}),
        "access_token": access_token,
        "limit": page_limit,
    }

    all_records: List[Dict[str, Any]] = []
    page_number = 0

    while url:
        page_number += 1
        print(f"Fetching page {page_number}...")

        response_json = _request_with_retries(url, params)
        # The first request uses `params`; every subsequent request uses the
        # full `paging.next` URL, which already contains its own query string.
        params = None

        if "error" in response_json:
            _raise_for_meta_error(response_json["error"])

        data = response_json.get("data")
        if data is None:
            raise MetaApiError(
                f"Meta API response is missing the 'data' field on page {page_number}. "
                "This usually means the request was malformed or the API changed "
                "unexpectedly -- treat pagination as incomplete and investigate before retrying."
            )

        all_records.extend(data)

        paging = response_json.get("paging", {})
        url = paging.get("next")

    return all_records, page_number


# ---------------------------------------------------------------------------
# Cleaning + metric calculation
# ---------------------------------------------------------------------------

def get_action_value(action_list: Optional[List[Dict[str, Any]]], action_type: str) -> float:
    """Return the numeric value for one exact canonical action_type.

    Meta's `actions` and `action_values` arrays contain several overlapping
    entries for the same underlying conversion event. We match on a single
    canonical action_type and ignore the rest, instead of summing the array,
    to avoid double- or triple-counting the same conversion.
    """
    if not action_list:
        return 0.0
    for action in action_list:
        if action.get("action_type") == action_type:
            try:
                return float(action.get("value", 0))
            except (TypeError, ValueError):
                return 0.0
    return 0.0


def safe_divide(numerator: float, denominator: float) -> Optional[float]:
    """Divide two numbers, returning None (blank in the CSV) instead of raising or infinity."""
    if not denominator:
        return None
    return numerator / denominator


def clean_insights_records(records: List[Dict[str, Any]]) -> pd.DataFrame:
    """Convert raw Meta Insights records into a clean, typed DataFrame.

    - Extracts canonical conversion metrics from `actions` / `action_values`.
    - Converts all numeric fields to proper numeric types.
    - Computes CPA, ROAS, and funnel conversion rates with safe division.
    """
    cleaned_rows = []
    skipped_count = 0

    for record in records:
        if not record.get("campaign_id") or not record.get("date_start"):
            skipped_count += 1
            continue

        actions = record.get("actions", [])
        action_values = record.get("action_values", [])

        spend = float(record.get("spend", 0) or 0)
        impressions = int(float(record.get("impressions", 0) or 0))
        reach = int(float(record.get("reach", 0) or 0))
        clicks = int(float(record.get("clicks", 0) or 0))
        ctr = float(record.get("ctr", 0) or 0)
        cpc = float(record.get("cpc", 0) or 0)
        cpm = float(record.get("cpm", 0) or 0)
        frequency = float(record.get("frequency", 0) or 0)

        purchases = get_action_value(actions, PURCHASE_ACTION_TYPE)
        purchase_value = get_action_value(action_values, PURCHASE_ACTION_TYPE)
        add_to_carts = get_action_value(actions, ADD_TO_CART_ACTION_TYPE)
        initiate_checkouts = get_action_value(actions, INITIATE_CHECKOUT_ACTION_TYPE)
        landing_page_views = get_action_value(actions, LANDING_PAGE_VIEW_ACTION_TYPE)

        cleaned_rows.append(
            {
                "date": record.get("date_start"),
                "campaign_id": record.get("campaign_id"),
                "campaign_name": record.get("campaign_name"),
                "objective": record.get("objective"),
                "spend": spend,
                "impressions": impressions,
                "reach": reach,
                "clicks": clicks,
                "ctr": ctr,
                "cpc": cpc,
                "cpm": cpm,
                "frequency": frequency,
                "purchases": purchases,
                "purchase_value": purchase_value,
                "CPA": safe_divide(spend, purchases),
                "ROAS": safe_divide(purchase_value, spend),
                "add_to_carts": add_to_carts,
                "initiate_checkouts": initiate_checkouts,
                "landing_page_views": landing_page_views,
                "click_to_landing_page_rate": safe_divide(landing_page_views, clicks),
                "landing_page_to_purchase_rate": safe_divide(purchases, landing_page_views),
            }
        )

    if skipped_count:
        print(
            f"Warning: skipped {skipped_count} record(s) missing campaign_id or date_start."
        )

    return pd.DataFrame(cleaned_rows, columns=OUTPUT_COLUMNS)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Collect and clean campaign-level daily Meta Ads Insights data."
    )
    parser.add_argument(
        "--since",
        default=DEFAULT_SINCE,
        help=f"Start date in YYYY-MM-DD format (default: {DEFAULT_SINCE}).",
    )
    parser.add_argument(
        "--until",
        default=DEFAULT_UNTIL,
        help=f"End date in YYYY-MM-DD format (default: {DEFAULT_UNTIL}).",
    )
    parser.add_argument(
        "--output",
        default=DEFAULT_OUTPUT_PATH,
        help=f"Output CSV path (default: {DEFAULT_OUTPUT_PATH}).",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    try:
        credentials = load_credentials()
    except MetaApiError as exc:
        print(f"Configuration error: {exc}")
        sys.exit(1)

    print(f"Fetching Meta Ads campaign-level daily insights from {args.since} to {args.until}...")

    try:
        raw_records, pages_fetched = fetch_campaign_insights(
            access_token=credentials["access_token"],
            ad_account_id=credentials["ad_account_id"],
            since=args.since,
            until=args.until,
        )
    except MetaApiError as exc:
        print(f"Meta API error: {exc}")
        sys.exit(1)

    print(
        f"Retrieved {len(raw_records)} raw daily campaign record(s) across "
        f"{pages_fetched} page(s). Pagination complete."
    )

    dataset = clean_insights_records(raw_records)
    print(f"Cleaned dataset contains {len(dataset)} row(s).")

    output_dir = os.path.dirname(args.output)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    dataset.to_csv(args.output, index=False)
    print(f"Saved cleaned dataset to {args.output}")

    # Sidecar metadata for downstream validation (e.g. CI). Contains no
    # credentials -- only run bookkeeping.
    metadata_path = os.path.join(output_dir or ".", "collection_metadata.json")
    metadata = {
        "since": args.since,
        "until": args.until,
        "pages_fetched": pages_fetched,
        "raw_record_count": len(raw_records),
        "cleaned_row_count": len(dataset),
        "generated_at": datetime.now(timezone.utc).isoformat(),
    }
    with open(metadata_path, "w") as metadata_file:
        json.dump(metadata, metadata_file, indent=2)
    print(f"Saved collection metadata to {metadata_path}")


if __name__ == "__main__":
    main()
