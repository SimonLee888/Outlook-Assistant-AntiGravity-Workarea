# Live Data Integration Lab

This package contains sanitized source samples derived from real operational
automation work.

## What is included

- Moraware job activity extraction through a .NET bridge.
- RFMS session bootstrap and estimate enrichment.
- Microsoft Graph timecard access with OAuth client credentials.
- Samsara historical vehicle polling in safe time windows.
- Excel cut-log normalization for downstream ETL.
- SQL duplicate cleanup logic for post-load quality control.

## Requirements

- Python 3.11 or newer
- A virtual environment
- Your own API credentials for the services you want to test
- Your own local or remote database target if you want to persist output

## Step by step

1. Extract the package into a clean local folder.
2. Create and activate a virtual environment.
3. Install the dependencies from `requirements.txt`.
4. Open the script you want to test first.
5. Replace placeholder values like `***` with your own safe local configuration.
6. Start with a very small scope such as one API call or a reduced date range.
7. Review the output shape before adapting the script to your own schema or storage.

## Suggested starting points

- Start with `rfms_estimate_bridge.py` if you want a simple API session flow.
- Start with `teams_timecards_sync.py` if you want an OAuth client-credentials example.
- Start with `excel_cutlog_normalizer.py` if you want a practical spreadsheet cleanup pattern.
- Start with `samsara_history_backfill.py` if you want a time-windowed historical polling example.

## Sanitization notes

- All credentials, tokens, tenant IDs, and internal identifiers were replaced.
- Absolute workstation paths and internal network details were removed.
- No raw business data, spreadsheets, exports, logs, or binaries are included.

These files are intended as portfolio-grade engineering samples, not as a drop-in
production deployment.
