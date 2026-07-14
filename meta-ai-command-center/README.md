# Meta Ads Marketing Command Center

## 1. What this project is

The Meta Ads Marketing Command Center is an AI-powered paid media monitoring
and reporting system built on top of the Meta Marketing API. It has two
layers today:

- **Layer 1 -- Data collection and cleaning** (`meta_collector.py`): pulls
  campaign-level daily performance data from the Meta Marketing API into a
  clean CSV.
- **Layer 2 -- Performance analysis and executive reporting**
  (`performance_analyzer.py` + `ai_analyst.py`): turns that CSV into
  structured, rule-based findings, then asks Claude to turn those findings
  into a concise executive report.

There is still no dashboard or automated optimisation/bidding layer -- those
remain out of scope for now.

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
├── meta_collector.py         # Layer 1: data collection + cleaning
├── performance_analyzer.py   # Layer 2: deterministic rule-based analysis engine
├── ai_analyst.py              # Layer 2: AI executive report writer (Claude)
├── validate_output.py         # Validates the collector's CSV output
├── tests/
│   └── test_performance_analyzer.py   # Unit tests for the rule engine
├── data/
│   ├── meta_campaign_daily.csv        # Generated (not committed to git)
│   └── performance_findings.json      # Generated (not committed to git)
├── reports/
│   └── meta_performance_report.md     # Generated (not committed to git)
├── config/                    # Reserved for future configuration files
├── README.md
├── requirements.txt
├── .env.example
└── .gitignore

.github/workflows/
└── meta-collector.yml         # Daily + on-demand pipeline (Layer 1 + Layer 2)
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

---

# Layer 2 -- Performance Analysis & AI Executive Reporting

Layer 2 turns the Layer 1 CSV into an executive-ready report in two strictly
separated stages: a **deterministic analysis engine** that owns every number,
and an **AI executive analyst** that only interprets and writes about numbers
it's given.

```
Layer 1 output (data/meta_campaign_daily.csv)
        │
        ▼
performance_analyzer.py   (deterministic -- no AI, no API calls)
        │  last 7 days vs. previous 7 days, per campaign + account-wide
        │  rule engine produces structured findings
        ▼
data/performance_findings.json
        │
        ▼
ai_analyst.py   (calls the Anthropic API once)
        │  interprets findings, writes FACT / SIGNAL / RECOMMENDATION report
        ▼
reports/meta_performance_report.md
```

## 10. Deterministic vs. AI responsibilities

This split is deliberate and is the most important design decision in Layer 2:

| | Deterministic engine (`performance_analyzer.py`) | AI analyst (`ai_analyst.py`) |
|---|---|---|
| Computes metrics and percentage changes | ✅ Yes -- the only source of truth | ❌ Never recalculates anything |
| Decides which findings fire | ✅ Yes -- fixed, auditable rule thresholds | ❌ Cannot invent or suppress a finding |
| Explains *why* something might be happening | ❌ No causal language at all | ✅ Yes, but only as a labeled SIGNAL, never a stated FACT |
| Writes the executive summary / prioritization / recommended actions | ❌ No | ✅ Yes |
| Calls an external API | ❌ No | ✅ Yes (Anthropic API, one call per run) |

If `performance_analyzer.py` and `ai_analyst.py` ever disagree on a number,
`performance_analyzer.py` is correct -- the AI is instructed to use the given
JSON as-is and never to recompute or override it.

## 11. `performance_analyzer.py` -- the analysis engine

Reads `data/meta_campaign_daily.csv` and, per campaign and for the account as
a whole:

1. Determines a **last 7 days** window (the dataset's most recent date, and
   the 6 days before it) and a **previous 7 days** window (the 7 days before
   that), based on the data's own max date -- not today's date.
2. Sums raw metrics (spend, impressions, clicks, purchases, purchase value,
   add-to-carts, initiated checkouts) within each window, then derives ratio
   metrics (CTR, CPC, CPM, CPA, ROAS, ATC rate, checkout rate,
   checkout-to-purchase rate) from those sums -- never by averaging daily
   ratios, which would misweight low-volume days.
3. Computes the percentage change between the two windows for spend, CTR,
   CPC, purchases, CPA, ROAS, add-to-carts, initiated checkouts, and
   checkout-to-purchase rate, using `safe_pct_change()`: if the previous
   value is zero or undefined, the change is `null`, never `inf` or `NaN`.
4. Runs the seven rules below against each campaign's numbers.
5. Sorts findings by severity (`critical` → `high` → `medium` → `positive`
   → `low`), then by that campaign's last-7-day spend, descending.
6. Writes everything -- metadata, an account-wide summary, and the sorted
   findings -- to `data/performance_findings.json`.

This script makes no network calls and does not require any API key.

## 12. Performance rules (deterministic anomaly detection)

| Rule | Category | Severity | Trigger |
|---|---|---|---|
| Wasted spend | `wasted_spend` | critical | Last 7-day spend ≥ 2,000 **and** last 7-day purchases = 0 |
| CPA deterioration | `cost_increase` | high | Last & previous 7-day purchases both ≥ 3 **and** CPA increased ≥ 40% |
| ROAS decline | `performance_decline` | high | Both periods have a defined ROAS **and** it declined ≥ 35% |
| Funnel leakage | `funnel_issue` | high | Last 7-day initiated checkouts ≥ 20 **and** checkout-to-purchase rate < 5% |
| Creative fatigue signal | `creative_fatigue` | medium | Spend increased ≥ 20% **and** CTR declined ≥ 20% -- reported explicitly as a possible signal, never a proven cause |
| Scaling opportunity | `scaling_opportunity` | positive | Last 7-day purchases ≥ 5 **and** ROAS ≥ 2 **and** CPA improved ≥ 20% |
| Tracking risk | `tracking_risk` | high | Last 7-day purchases > 0 **and** last 7-day purchase value = 0 |

Every finding carries `finding_id`, `campaign_id`, `campaign_name`,
`severity`, `category`, `title`, `observation`, `evidence` (the exact numbers
that triggered it), and `recommended_action`. Rules are pure functions in
`performance_analyzer.py` and are covered by `tests/test_performance_analyzer.py`.

## 13. `data/performance_findings.json` structure

```json
{
  "analysis_metadata": {
    "generated_at": "...",
    "analysis_end_date": "...",
    "last_7_day_window": { "start": "...", "end": "..." },
    "previous_7_day_window": { "start": "...", "end": "..." },
    "campaigns_analyzed": 0,
    "findings_generated": 0
  },
  "account_summary": {
    "last_7_days": { "...": "..." },
    "previous_7_days": { "...": "..." },
    "changes": { "...": "..." }
  },
  "findings": [ { "finding_id": "...", "...": "..." } ]
}
```

## 14. `ai_analyst.py` -- the AI executive analyst

Reads `data/performance_findings.json` and makes **one**, deliberately plain
call to the Anthropic Messages API -- just `model`, `max_tokens`, `system`,
and `messages`. No thinking configuration or effort tuning is set; those are
optional API features, not requirements for this task, and leaving them out
keeps the call simple and easy to reason about.

**The model ID is read exclusively from the `ANTHROPIC_MODEL` environment
variable -- there is no hardcoded default in the script.** A model ID
appearing in an SDK's type hints or a cached model catalog is not proof that
it's valid or available for a given account, so this script does not assume
one on your behalf. If `ANTHROPIC_MODEL` is unset, the script fails
immediately with a message telling you to configure it (see section 17)
rather than guessing a model to call.

The system prompt given to Claude enforces the boundaries above: it must not
recalculate metrics, must not invent a cause for anything, and must label
every claim as FACT, SIGNAL, or RECOMMENDATION. The report follows a fixed
structure: Executive Summary (max 5 bullets), Critical Issues, High Priority
Risks, Funnel Health, Scaling Opportunities, Recommended Actions -- Next 24
Hours (max 5 actions), and an Account Performance Snapshot table.

`ANTHROPIC_API_KEY` is read only from that environment variable -- never
hardcoded, never logged, never written to the report. If the key or the
model ID is missing, or the Anthropic API call fails for any reason (invalid
key, rate limit, network error, refusal), the script prints a clear error and
exits non-zero **without writing any report file** -- it never produces a
placeholder or fabricated report on failure.

## 15. Running Layer 2 locally

With Layer 1 already run (so `data/meta_campaign_daily.csv` exists) and both
`ANTHROPIC_API_KEY` and `ANTHROPIC_MODEL` set in your environment (same rule
as `META_ACCESS_TOKEN` in section 7 -- never paste either into a chat):

```bash
export ANTHROPIC_MODEL=your-chosen-model-id
python performance_analyzer.py
python ai_analyst.py
```

Both scripts accept `--input`/`--output` (analyzer) and `--findings`/`--output`
(analyst) flags if you want non-default paths.

## 16. Testing

```bash
python -m unittest discover -s tests -t . -v
```

The suite in `tests/test_performance_analyzer.py` uses hand-built synthetic
metrics only -- it never calls the live Meta API or the Anthropic API. It
covers all seven rules firing and not firing, safe handling of a zero/missing
previous period (no `inf`/`NaN`), severity + spend sorting, and that no rule
produces duplicate findings for the same campaign.

## 17. Running via GitHub Actions

The full pipeline runs automatically through
`.github/workflows/meta-collector.yml`:

1. **Trigger:** once a day on a schedule (`06:00 UTC`), and on demand via
   the "Run workflow" button (`workflow_dispatch`), with optional `since` /
   `until` inputs to override the collector's default date range for that
   run.
2. **Collect:** `meta_collector.py` fetches campaign-level daily data.
3. **Validate:** `validate_output.py` checks total rows, unique campaigns,
   date range, total spend, total canonical purchases and purchase value,
   duplicate campaign-date rows, and pagination completeness. The job fails
   here if no rows were collected or duplicates are found -- Layer 2 never
   runs against unvalidated data.
4. **Analyze:** `performance_analyzer.py` runs the deterministic rule engine.
   If it fails for any reason, the workflow fails immediately.
5. **Report:** `ai_analyst.py` calls the Anthropic API to write the executive
   report. If `ANTHROPIC_MODEL` isn't configured, or the API call fails for
   any reason, this step fails with a clear error and the workflow fails --
   it does not fall back to a fake report.
6. **Upload artifacts:** `data/meta_campaign_daily.csv` is uploaded as one
   artifact; `data/performance_findings.json` and
   `reports/meta_performance_report.md` are uploaded together as a second
   artifact. None of these three files are committed to the repository --
   all are covered by `.gitignore`.

### Required GitHub secrets and variables

Configure these under Settings → Secrets and variables → Actions:

| Name | Kind | Used by |
|---|---|---|
| `META_ACCESS_TOKEN` | Secret | `meta_collector.py` |
| `META_AD_ACCOUNT_ID` | Secret | `meta_collector.py` |
| `ANTHROPIC_API_KEY` | Secret | `ai_analyst.py` |
| `ANTHROPIC_MODEL` | **Variable** (Variables tab, not Secrets) | `ai_analyst.py` |

`ANTHROPIC_MODEL` is a repository *variable* rather than a secret because a
model ID is not sensitive -- it's configuration, not a credential. The
workflow passes it as `ANTHROPIC_MODEL: ${{ vars.ANTHROPIC_MODEL }}` (note
`vars.*`, not `secrets.*`). All four are passed to the relevant step as
environment variables and are never printed or logged by any script in this
repository.

## 18. Security approach

- Credentials are read exclusively from environment variables
  (`META_ACCESS_TOKEN`, `META_AD_ACCOUNT_ID`, `ANTHROPIC_API_KEY`) -- never
  hardcoded, never logged, never written to a CSV, JSON, or markdown output.
- The Anthropic model ID is also read exclusively from an environment
  variable (`ANTHROPIC_MODEL`), separately from the API key, since it's
  configuration rather than a secret -- but it is still never hardcoded or
  guessed: the script has no default and fails clearly if it's unset.
- `.env` and all generated data (`data/*.csv`, `data/*.json`,
  `reports/*.md`) are excluded from version control via `.gitignore`.
- The deterministic engine and the AI analyst are hard-separated: the AI
  never sees raw API responses or credentials, only the already-computed,
  already-validated findings JSON.
- The AI is explicitly constrained to label claims as FACT / SIGNAL /
  RECOMMENDATION and forbidden from asserting unproven causes.
- Every stage fails loudly (non-zero exit, clear message) rather than
  silently producing partial or fabricated output.

## 19. Limitations

- The rule thresholds (e.g. 2,000 spend, 40% CPA increase, 35% ROAS decline)
  are fixed starting points, not tuned to any specific account -- revisit
  them as real data comes in.
- The analysis window is always "last 7 days vs. previous 7 days" ending on
  the dataset's max date; it does not yet support custom window lengths or
  day-of-week seasonality adjustments.
- The AI analyst makes one non-streaming API call per run; very large
  findings sets (many campaigns with many simultaneous findings) could
  approach context/output limits and may need chunking in a future revision.
- This is still not an anomaly-detection or forecasting system in the
  statistical sense -- it is a fixed, transparent rule engine plus an LLM
  writer. There is no historical baselining, seasonality modeling, or
  learned thresholds yet.
- No dashboard, alerting/notification integration, or automated
  bid/budget optimisation exists yet -- those remain future phases.
