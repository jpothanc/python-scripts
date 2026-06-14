# API Test Monitor

Call a configured list of REST APIs at regular intervals, log **API name**, **HTTP status**, and **response time**, and track **timeouts** and errors. Output is a CSV file suitable for analysis with [`api-log-analyzer`](../api-log-analyzer/).

## Requirements

- Python 3.10+
- Dependencies: `requests`, `PyYAML`

On Windows, use the `py` launcher if `python` is not available:

```powershell
py --version
```

## Installation

```powershell
cd c:\Projects\Python\python-scripts\api-test
py -m pip install -r requirements.txt
```

## Quick start

```powershell
py monitor.py
```

Uses `config.yaml` in the current folder. Results are written to `results.csv`.

For a short test run, edit `config.yaml` first:

```yaml
settings:
  duration_hours: 0.05    # ~3 minutes
  interval_seconds: 30
```

## What it does

1. Reads API endpoints from `config.yaml` (up to four per run).
2. Calls each API in turn every `interval_seconds`.
3. Records for every call:
   - Timestamp (UTC)
   - API name
   - URL and HTTP method
   - HTTP status code
   - Response time (milliseconds)
   - Result: `success`, `timeout`, or `error`
4. Runs for `duration_hours` (default 24 hours).
5. Prints a summary when finished (total success / timeout / error counts).

Each request uses a **10 second timeout** by default (`timeout_seconds` in config).

## Configuration (`config.yaml`)

```yaml
settings:
  duration_hours: 24        # how long to run
  interval_seconds: 60      # wait between full rounds
  timeout_seconds: 10       # per-request timeout
  output_file: results.csv  # CSV log path

apis:
  - name: httpbin_get
    url: https://httpbin.org/get
    method: GET

  - name: my_api
    url: https://api.example.com/health
    method: GET
    headers:
      Authorization: Bearer your-token
```

Paths in `output_file` are relative to the config file’s folder unless absolute.

### Settings

| Setting | Description |
|---------|-------------|
| `duration_hours` | Total run time in hours (`24` = one day) |
| `interval_seconds` | Seconds between each full round of checks |
| `timeout_seconds` | Request timeout in seconds |
| `output_file` | CSV file to write |

### API entries

| Field | Required | Description |
|-------|----------|-------------|
| `name` | Yes | Short name used in logs and reports |
| `url` | Yes | Full URL to call |
| `method` | No | HTTP method (default `GET`) |
| `headers` | No | Optional request headers |

## CSV output

Default file: `results.csv`

| Column | Description |
|--------|-------------|
| `timestamp_utc` | When the call was made |
| `api_name` | Name from config |
| `url` | Request URL |
| `method` | HTTP method |
| `http_status` | Status code (empty on timeout/error) |
| `response_time_ms` | Elapsed time in milliseconds |
| `result` | `success`, `timeout`, or `error` |
| `error_message` | Details when not successful |

Example rows:

```csv
timestamp_utc,api_name,url,method,http_status,response_time_ms,result,error_message
2026-06-09T09:00:01+00:00,httpbin_get,https://httpbin.org/get,GET,200,85.20,success,
2026-06-09T09:02:01+00:00,httpbin_delay_15s,https://httpbin.org/delay/15,GET,,10000.00,timeout,Request exceeded 10s timeout
```

## Command-line usage

```powershell
py monitor.py [-c config.yaml]
```

| Option | Description |
|--------|-------------|
| `-c`, `--config` | Path to YAML config (default: `config.yaml`) |

## While it runs

- Progress prints in the terminal (status, response time, timeouts).
- Each result is appended to the CSV immediately.
- Press **Ctrl+C** to stop after the current round; a summary is still printed.

## Terminal summary

At the end of a run:

```
=== Summary ===
Total checks: 96
Successful:   90
Timeouts:     4
Errors:       2

Per API:
  httpbin_get: success=24, timeout=0, error=0
  ...
```

## Excel report (optional)

To turn `results.csv` into an Excel workbook with time buckets and highlighted timeouts, use **api-log-analyzer**:

```powershell
cd c:\Projects\Python\python-scripts\api-log-analyzer
py -m pip install -r requirements.txt
py analyze_api_logs.py -i ..\api-test\results.csv -o api_report.xlsx
```

See [`api-log-analyzer/README.md`](../api-log-analyzer/README.md) for details.

## Project layout

```text
api-test/
├── README.md
├── config.yaml
├── monitor.py
├── requirements.txt
└── results.csv          # created when you run the monitor
```

## Typical workflow

1. Edit `config.yaml` with your APIs and run duration.
2. Run `py monitor.py` (leave it running for the day if needed).
3. Review `results.csv` or generate an Excel report with `api-log-analyzer`.
4. Use timeout and response-time counts to spot slow or failing APIs.

## Troubleshooting

| Issue | What to try |
|-------|-------------|
| `python` not found | Use `py` instead of `python` |
| No results in CSV | Check URLs and network access |
| Many timeouts | Increase `timeout_seconds` or fix slow endpoints |
| File locked | Close Excel if `results.csv` is open elsewhere |
