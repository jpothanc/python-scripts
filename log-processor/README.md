# Log Processor

Process up to four timestamp-only log files, extract lines matching a configurable pattern, assign calendar dates from a start date, and export results to a single Excel workbook with one sheet per log file.

## Requirements

- Python 3.10+
- Dependencies: `openpyxl`, `PyYAML`

On Windows, use the `py` launcher if `python` is not available:

```powershell
py --version
```

## Installation

```powershell
cd c:\Projects\Python\python-scripts\log-processor
py -m pip install -r requirements.txt
```

## Quick start

With the default `config.yaml` in the current directory:

```powershell
py process_logs.py
```

This writes `log_analysis.xlsx` using the sample logs under `sample_logs/`.

## How it works

1. Reads each log file line by line.
2. Matches lines containing your **line pattern** (configured in YAML or via `--pattern`).
3. Extracts the numeric value where `x` appears in the pattern (for example, millis).
4. Converts the value to **seconds** using `value_divisor` (default `1000` for millis).
5. Optionally keeps only lines where duration is **greater than** `min_seconds`.
6. Parses the **time** on each line (no date in the log).
7. Assigns a **calendar date** starting from the start date; when the time moves backward (e.g. `23:59` then `00:02`), the date advances by one day.
8. Writes one Excel sheet per log file (sheet name = file name without extension).
9. Adds a **Summary** sheet (first tab) in a **side-by-side** layout: two columns per log file (**DateTime**, **Seconds**) plus one **Bucket** filter column (e.g. 5 columns for 2 logs).

### Example log line

```
23:58:10 INFO DB retrieval took 120 millis
```

With pattern `DB retrieval took x millis` and start date `2026-06-01`, the first rows might be:

| DateTime            | Seconds | Bucket | >10s | >30s | >1min | >2min | >3min | >4min | >5min |
|---------------------|---------|--------|------|------|-------|-------|-------|-------|-------|
| 2026-06-01 23:58:10 | 0.12    | <=10s  |      |      |       |       |       |       |       |
| 2026-06-02 00:02:30 | 125.0   | >2min  | Yes  | Yes  | Yes   | Yes   |       |       |       |

**Bucket** — one label per row (<=10s, >10s, >30s, …, >5 min). **>10s** … **>5min** — `Yes` when duration exceeds that threshold (use Excel AutoFilter on `Yes`).

## Configuration (`config.yaml`)

```yaml
settings:
  start_date: 2026-06-01      # optional if you pass --start-date
  output_file: log_analysis.xlsx

line_pattern: "DB retrieval took x millis"

value_divisor: 1000           # 1000 = millis → seconds, 1 = already seconds

min_seconds: 0                # only include lines with duration > this (seconds)

log_files:                    # optional if you pass -f or positional paths
  - sample_logs/server-a.log
  - sample_logs/server-b.log
```

Paths in `log_files` and `output_file` are resolved relative to the config file’s directory unless you use an absolute path.

### Line pattern

Use `x` (or `X`) where the number appears in the log:

| Log text                         | `line_pattern`                      |
|----------------------------------|-------------------------------------|
| `DB retrieval took 150 millis`   | `DB retrieval took x millis`        |
| `Cache lookup duration: 42ms`    | `Cache lookup duration: xms`        |
| `Query completed in 3.5 sec`     | `Query completed in x sec` + `value_divisor: 1` |

### Minimum duration filter

Set `min_seconds` in config or pass `--min-seconds` to export only lines **greater than** that threshold (after converting to seconds):

```yaml
min_seconds: 1.0    # keep only DB calls taking more than 1 second
```

```powershell
py process_logs.py --min-seconds 1
```

## Command-line usage

```text
py process_logs.py [-c CONFIG] [-d YYYY-MM-DD] [--pattern PATTERN]
                   [--min-seconds SECONDS] [-o OUTPUT] [-f PATH ...] [log_files ...]
```

| Option | Description |
|--------|-------------|
| `-c`, `--config` | YAML config file (default: `config.yaml` in the current directory if it exists) |
| `-d`, `--start-date` | First calendar date (`YYYY-MM-DD`). Overrides `settings.start_date` in config |
| `--pattern` | Line pattern with `x` as the numeric placeholder. Overrides `line_pattern` in config |
| `--min-seconds` | Only include lines with duration **greater than** this many seconds. Overrides `min_seconds` in config |
| `-o`, `--output` | Output Excel path. Overrides `settings.output_file` in config |
| `-f`, `--log-file` | Log file path (repeat up to 4 times). Overrides `log_files` in config |
| positional `log_files` | Log paths at the end of the command (alternative to `-f`) |

CLI values override the config when both are provided. You must supply:

- **Start date** — `--start-date` or `settings.start_date` in config
- **At least one log file** — `-f`, positional paths, or `log_files` in config
- **Line pattern** — `--pattern` or `line_pattern` in config (unless using a config file that defines it)

### Examples

**Config only:**

```powershell
py process_logs.py
```

**Override start date and log files:**

```powershell
py process_logs.py -d 2026-06-15 -f sample_logs\server-a.log -f sample_logs\server-b.log
```

**Only slow DB calls (> 1 second):**

```powershell
py process_logs.py --min-seconds 1
```

**Override output file:**

```powershell
py process_logs.py -d 2026-06-01 -o weekly_report.xlsx
```

**Positional log paths:**

```powershell
py process_logs.py -d 2026-06-01 sample_logs\server-a.log sample_logs\server-b.log
```

**No config file — all options on the command line:**

```powershell
py process_logs.py `
  -d 2026-06-01 `
  --pattern "DB retrieval took x millis" `
  -f C:\logs\server-a.log `
  -f C:\logs\server-b.log `
  -o report.xlsx
```

**Custom config path:**

```powershell
py process_logs.py -c C:\configs\prod.yaml -d 2026-06-01
```

## Output

- Single `.xlsx` file (default: `log_analysis.xlsx`).
- **Summary** sheet (first tab): side-by-side columns per log file, e.g. `server-a DateTime | server-a Seconds | server-b DateTime | server-b Seconds | Bucket`. Rows align by entry order (1st match from each log on row 1, etc.). **Bucket** uses the slowest duration on that row. AutoFilter enabled on **Bucket**.
- One worksheet per input log file with **DateTime**, **Seconds**, **Bucket**, and threshold flags **>10s** … **>5min**.

## Project layout

```text
log-processor/
├── README.md
├── config.yaml           # Default settings and sample paths
├── process_logs.py       # Main script
├── requirements.txt
└── sample_logs/          # Example logs for testing
    ├── server-a.log
    └── server-b.log
```

## Limits

- Maximum **4** log files per run.
- Log lines must contain a parseable time (`HH:MM:SS` or with fractional seconds).
- Lines without a matching pattern or without a timestamp are skipped.

## Troubleshooting

| Issue | What to try |
|-------|-------------|
| `python` not found | Use `py` instead of `python` |
| No matching lines | Check `line_pattern` matches your log text; ensure `X` is where the number is |
| Wrong dates | Verify `--start-date`; midnight rollover assumes logs are in chronological order |