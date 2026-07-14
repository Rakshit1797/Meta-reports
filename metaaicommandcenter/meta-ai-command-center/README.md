# Meta Ads Marketing Command Center

## 1. What this project is

The Meta Ads Marketing Command Center is an AI-powered paid media monitoring
and reporting system built on top of the Meta Marketing API. The long-term
goal is a command center that automatically collects Meta Ads performance
data, cleans it, and (in later phases) analyzes it for anomalies and
optimisation opportunities.

**This repository currently contains only the first layer: data collection
and data cleaning.** There is no AI analysis, anomaly detection, or
dashboard yet -- those are separate, future phases.

## 2. What `meta_collector.py` does

`meta_collector.py` is a standalone Python script that:

1. Connects to the Meta Marketing API's Ads Insights endpoint for a single
   ad account.
2. Requests **campaign-level, daily** performance data (`level=campaign`,
   `time_increment=1`) for a configurable date range.
3. Automatically follows `paging.next` until every page of results has been
   downloaded, and combines them into one dataset.
4. Cleans Meta's `actions` / `action_values` arrays down to a single,
   canonical conversion event per metric (see section 5).
5. Converts every numeric field to a proper numeric type.
6. Calculates CPA, ROAS, and funnel conversion rates safely (no
   divide-by-zero crashes).
7. Writes the result to `data/meta_campaign_daily.csv`.

It does not perform any AI analysis, anomaly detection, or optimisation
recommendations -- that is intentionally out of scope for this phase.

## 3. Project architecture

```
meta-ai-command-center/
├── meta_collector.py       # Data collection + cleaning script (this phase)
├── data/
│   └── meta_campaign_daily.csv   # Generated output (not committed to git)
├── config/                 # Reserved for future configuration files
├── README.md
├── requirements.txt
├── .env.example
└── .gitignore
```

The script is organized into small, single-purpose functions:

- `load_credentials()` -- reads `META_ACCESS_TOKEN` / `META_AD_ACCOUNT_ID`
  from the environment.
- `fetch_campaign_insights()` -- builds the API request and drives
  pagination.
- `_request_with_retries()` -- performs a single HTTP request with retry/
  backoff logic for network failures and rate limits.
- `_raise_for_meta_error()` -- turns a Meta API error object into a clear,
  human-readable exception (without ever exposing the access token).
- `get_action_value()` -- pulls one canonical conversion metric out of
  Meta's `actions` / `action_values` arrays.
- `safe_divide()` -- division helper that returns a blank value instead of
  crashing or producing infinity.
- `clean_insights_records()` -- turns raw API records into a typed,
  metric-complete pandas DataFrame.
- `main()` -- CLI entry point that wires everything together.

## 4. How data flows from the Meta Marketing API into the CSV

```
Meta Ads Insights API (level=campaign, time_increment=1)
        │
        │  page 1 (data + paging.next)
        ▼
fetch_campaign_insights()  ──▶ follows paging.next until exhausted
        │
        │  combined list of raw daily campaign records
        ▼
clean_insights_records()
        │  - extract canonical purchase / add-to-cart / initiate-checkout /
        │    landing-page-view metrics from actions & action_values
        │  - cast every numeric field to float/int
        │  - compute CPA, ROAS, click_to_landing_page_rate,
        │    landing_page_to_purchase_rate with safe division
        ▼
pandas DataFrame (one row per campaign per day)
        │
        ▼
data/meta_campaign_daily.csv
```

Each row in the final CSV represents **one campaign, on one day**, with the
columns listed in section 9.

## 5. How duplicate Meta conversion events are prevented

Meta's `actions` and `action_values` arrays often contain multiple
overlapping entries for what is really the *same* conversion, for example:

- `purchase`
- `omni_purchase`
- `onsite_web_purchase`
- `onsite_web_app_purchase`
- `web_in_store_purchase`
- `offsite_conversion.fb_pixel_purchase`

Summing all of these would massively over-count real purchases, revenue,
add-to-carts, and checkouts.

To avoid this, `get_action_value()` never sums the array. Instead it scans
the array for **one exact, canonical `action_type`** and returns only that
value:

| Metric              | Canonical `action_type` used                        |
|---------------------|------------------------------------------------------|
| Purchases            | `offsite_conversion.fb_pixel_purchase`               |
| Purchase value       | `offsite_conversion.fb_pixel_purchase` (from `action_values`) |
| Add to cart          | `offsite_conversion.fb_pixel_add_to_cart`            |
| Initiate checkout    | `offsite_conversion.fb_pixel_initiate_checkout`      |
| Landing page views   | `landing_page_view`                                  |

If the canonical action type isn't present for a given campaign/day, the
metric is simply `0` -- it is never inferred from another action type.

## 6. Installing dependencies

From inside the `meta-ai-command-center/` directory:

```bash
python3 -m venv .venv
source .venv/bin/activate      # on Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## 7. Configuring environment variables

**Never put your Meta access token in code, in the CSV, in logs, or in a
chat window.** Configure it locally instead:

1. Copy the example file:
   ```bash
   cp .env.example .env
   ```
2. Open `.env` in a text editor (not in a chat with anyone, including an
   AI assistant) and fill in your real values:
   ```
   META_ACCESS_TOKEN=your_real_access_token
   META_AD_ACCOUNT_ID=act_1647415089594272
   ```
3. `.env` is already listed in `.gitignore`, so it will never be committed
   to version control.

The script loads `.env` automatically via `python-dotenv` -- you do not
need to `export` the variables manually, though you can if you prefer to
manage them at the shell/OS level instead of a file.

## 8. Running the collector

With `.env` configured and dependencies installed:

```bash
python meta_collector.py
```

By default this fetches campaign-level daily data from `2026-01-14` to
`2026-07-14` and writes it to `data/meta_campaign_daily.csv`.

To use a different date range or output path without editing the code:

```bash
python meta_collector.py --since 2026-06-01 --until 2026-06-30 --output data/meta_campaign_daily.csv
```

Available flags:

| Flag        | Description                                  | Default        |
|-------------|-----------------------------------------------|----------------|
| `--since`   | Start date (`YYYY-MM-DD`)                     | `2026-01-14`   |
| `--until`   | End date (`YYYY-MM-DD`)                       | `2026-07-14`   |
| `--output`  | Output CSV path                               | `data/meta_campaign_daily.csv` |

## 9. Expected CSV output

`data/meta_campaign_daily.csv` contains one row per campaign per day, with
these columns:

| Column | Description |
|---|---|
| `date` | The day this row's metrics apply to |
| `campaign_id` | Meta campaign ID |
| `campaign_name` | Meta campaign name |
| `objective` | Campaign objective |
| `spend` | Amount spent |
| `impressions` | Impressions |
| `reach` | Reach |
| `clicks` | Clicks |
| `ctr` | Click-through rate |
| `cpc` | Cost per click |
| `cpm` | Cost per 1,000 impressions |
| `frequency` | Average impressions per person |
| `purchases` | Canonical purchase count (`offsite_conversion.fb_pixel_purchase`) |
| `purchase_value` | Canonical purchase value |
| `CPA` | `spend / purchases` (blank if `purchases` is 0) |
| `ROAS` | `purchase_value / spend` (blank if `spend` is 0) |
| `add_to_carts` | Canonical add-to-cart count |
| `initiate_checkouts` | Canonical initiate-checkout count |
| `landing_page_views` | Canonical landing page view count |
| `click_to_landing_page_rate` | `landing_page_views / clicks` (blank if `clicks` is 0) |
| `landing_page_to_purchase_rate` | `purchases / landing_page_views` (blank if `landing_page_views` is 0) |

The access token is never written to this file, printed to the console, or
logged anywhere.

## 10. Running via GitHub Actions

The collector also runs automatically through
`.github/workflows/meta-collector.yml`:

- **Trigger:** once a day on a schedule (`06:00 UTC`), and on demand via
  the "Run workflow" button (`workflow_dispatch`), with optional `since` /
  `until` inputs to override the default date range for that run.
- **Secrets:** `META_ACCESS_TOKEN` and `META_AD_ACCOUNT_ID` must be
  configured as repository secrets (Settings → Secrets and variables →
  Actions). The workflow passes them to `meta_collector.py` as environment
  variables and never logs them.
- **Validation:** after collection, `validate_output.py` checks total rows,
  unique campaigns, the collected date range, total spend, total canonical
  purchases and purchase value, duplicate campaign-date rows, and
  pagination completeness (via the `collection_metadata.json` sidecar
  written by the collector). The job fails if no rows were collected or if
  duplicate campaign-date rows are found.
- **Output:** `data/meta_campaign_daily.csv` is uploaded as a workflow
  artifact for each run. It is never committed to the repository -- both
  the CSV and the metadata sidecar are covered by `.gitignore`.
