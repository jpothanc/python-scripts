# API Log Analyzer

Read API monitor log files and produce an Excel report with a **Summary** sheet (response-time buckets per API) and a **Records** sheet (every call, with timeouts highlighted).

Works with CSV output from [`api-test/monitor.py`](../api-test/monitor.py) or any CSV with the same columns.

## Log format

Each row in the CSV should contain at least:

| Column | Description |
|--------|-------------|
| `timestamp_utc` | When the call was made (ISO format) |
| `api_name` | Name of the API |
| `http_status` | HTTP status code (empty on timeout) |
| `response_time_ms` | Time taken in milliseconds |
| `result` | `success`, `timeout`, or `error` |

Example (written by the API monitor):

```csv
timestamp_utc,api_name,url,method,http_status,response_time_ms,result,error_message
2026-06-09T09:00:01+00:00,httpbin_get,https://httpbin.org/get,GET,200,85.20,success,
2026-06-09T09:02:01+00:00,httpbin_delay_15s,https://httpbin.org/delay/15,GET,,10000.00,timeout,Request exceeded 10s timeout
```

## Step 1 — Collect logs (api-test monitor)

From the `api-test` folder, run the monitor to call APIs and write `results.csv`:

```powershell
cd c:\Projects\Python\python-scripts\api-test
py -m pip install -r requirements.txt
py monitor.py
```

This logs **API name**, **HTTP status**, and **time taken** for every call. Timeouts are recorded with `result=timeout`.

Copy or point the analyzer at that file, e.g. `..\api-test\results.csv`.

## Step 2 — Create Excel report

```powershell
cd c:\Projects\Python\python-scripts\api-log-analyzer
py -m pip install -r requirements.txt
py analyze_api_logs.py
```

Uses `config.yaml` by default (`sample_logs/api_calls.csv` → `api_report.xlsx`).

Analyze the monitor output directly:

```powershell
py analyze_api_logs.py -i ..\api-test\results.csv -o api_report.xlsx
```

## Excel output

### Summary sheet

Counts per **API name** across time buckets:

| Bucket | Meaning |
|--------|---------|
| `<10ms` | Under 10 ms |
| `10-20ms` | 10 ms up to 20 ms |
| `20-50ms` | … |
| `50-100ms` | … |
| `100-200ms` | … |
| `200-500ms` | … |
| `500ms-1s` | … |
| `1s-1min` | … |
| `>1min` | Over 1 minute |
| `timeout` | Request timed out |

Includes an **ALL APIS** total row when multiple APIs are present. **Timeout** counts use a red highlight.

### Records sheet

Every log row with:

- Timestamp  
- API Name  
- HTTP Status  
- Time (ms)  
- Time Bucket  
- Result  

**Timeout rows are highlighted in red** for quick scanning.

## Configuration (`config.yaml`)

```yaml
input_file: sample_logs/api_calls.csv
output_file: api_report.xlsx
```

Paths are relative to the config file’s folder unless absolute.

## Command-line options

```text
py analyze_api_logs.py [-c config.yaml] [-i input.csv] [-o output.xlsx] [input.csv]
```

| Option | Description |
|--------|-------------|
| `-c`, `--config` | YAML config file |
| `-i`, `--input` | Input CSV log file |
| `-o`, `--output` | Output Excel file |
| positional | Input CSV (alternative to `-i`) |

## Project layout

```text
api-log-analyzer/
├── README.md
├── analyze_api_logs.py
├── config.yaml
├── requirements.txt
└── sample_logs/
    └── api_calls.csv
```

## Typical workflow

1. Configure APIs in `api-test/config.yaml`.
2. Run `py monitor.py` for a day (or your desired duration).
3. Run `py analyze_api_logs.py -i ..\api-test\results.csv`.
4. Open `api_report.xlsx` — use **Summary** for bucket counts and **Records** for line-by-line detail (timeouts in red).

## Requirements

- Python 3.10+
- `openpyxl`, `PyYAML`

On Windows use `py` if `python` is not available:

```powershell
py --version
```
