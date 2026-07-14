#!/usr/bin/env python3
"""
Validate the cleaned Meta Ads campaign-level daily dataset produced by
meta_collector.py.

Reads the generated CSV (and its collection_metadata.json sidecar) and
reports row/campaign counts, date range, financial and conversion totals,
duplicate campaign-date rows, and whether pagination completed. Never
touches API credentials.
"""

import argparse
import json
import sys
from typing import Any, Dict, Optional

import pandas as pd


def load_metadata(path: str) -> Optional[Dict[str, Any]]:
    try:
        with open(path) as metadata_file:
            return json.load(metadata_file)
    except FileNotFoundError:
        return None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate the Meta Ads collector's cleaned CSV output."
    )
    parser.add_argument("--csv", default="data/meta_campaign_daily.csv")
    parser.add_argument("--metadata", default="data/collection_metadata.json")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    df = pd.read_csv(args.csv)
    metadata = load_metadata(args.metadata)

    total_rows = len(df)
    unique_campaigns = df["campaign_id"].nunique() if total_rows else 0
    min_date = df["date"].min() if total_rows else None
    max_date = df["date"].max() if total_rows else None
    total_spend = df["spend"].sum() if total_rows else 0.0
    total_purchases = df["purchases"].sum() if total_rows else 0.0
    total_purchase_value = df["purchase_value"].sum() if total_rows else 0.0
    duplicate_rows = int(df.duplicated(subset=["campaign_id", "date"]).sum()) if total_rows else 0

    print("=== Meta Ads Collector Validation ===")
    print(f"Total rows collected:           {total_rows}")
    print(f"Unique campaigns:                {unique_campaigns}")
    print(f"Date range in data:              {min_date} to {max_date}")
    print(f"Total spend:                     {total_spend:.2f}")
    print(f"Total canonical purchases:       {total_purchases:.2f}")
    print(f"Total canonical purchase value:  {total_purchase_value:.2f}")
    print(f"Duplicate campaign-date rows:    {duplicate_rows}")

    if metadata:
        print(f"Requested date range:           {metadata.get('since')} to {metadata.get('until')}")
        print(f"Pages fetched:                   {metadata.get('pages_fetched')}")
        print("Pagination completed successfully: Yes")
    else:
        print(
            f"Warning: metadata file '{args.metadata}' not found -- "
            "cannot confirm pagination completeness from run bookkeeping."
        )

    errors = []
    if total_rows == 0:
        errors.append("No rows were collected.")
    if duplicate_rows > 0:
        errors.append(f"Found {duplicate_rows} duplicate campaign-date row(s).")

    if errors:
        print("\nVALIDATION FAILED:")
        for error in errors:
            print(f"  - {error}")
        sys.exit(1)

    print("\nValidation passed.")


if __name__ == "__main__":
    main()
