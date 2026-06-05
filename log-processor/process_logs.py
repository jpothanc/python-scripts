#!/usr/bin/env python3
"""Process timestamp-only log files and export matched metrics to Excel."""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.child import INVALID_TITLE_REGEX
from openpyxl.worksheet.worksheet import Worksheet

MAX_LOG_FILES = 4

TIME_PATTERN = re.compile(
    r"(?<!\d)"
    r"(\d{1,2}):(\d{2}):(\d{2})"
    r"(?:[.,](\d{1,6}))?"
    r"(?!\d)"
)


PLACEHOLDER_SPLIT = re.compile(r"(?i)x")


@dataclass(frozen=True)
class LogEntry:
    entry_datetime: datetime
    seconds: float


@dataclass(frozen=True)
class ProcessorConfig:
    start_date: date | None
    line_pattern: str
    value_divisor: float
    min_seconds: float
    output_file: Path
    log_files: list[Path]


def pattern_to_regex(pattern: str) -> re.Pattern[str]:
    if not PLACEHOLDER_SPLIT.search(pattern):
        raise ValueError("line_pattern must contain 'x' as the numeric placeholder.")

    parts = PLACEHOLDER_SPLIT.split(pattern)
    regex_body = r"(\d+(?:\.\d+)?)".join(re.escape(part) for part in parts)
    return re.compile(regex_body, re.IGNORECASE)


def parse_time_from_line(line: str) -> time | None:
    match = TIME_PATTERN.search(line)
    if not match:
        return None

    hour, minute, second, fraction = match.groups()
    microsecond = 0
    if fraction:
        microsecond = int(fraction.ljust(6, "0")[:6])

    return time(
        hour=int(hour),
        minute=int(minute),
        second=int(second),
        microsecond=microsecond,
    )


def assign_dates(entries: list[tuple[time, float]], start_date: date) -> list[LogEntry]:
    if not entries:
        return []

    current_date = start_date
    previous_time: time | None = None
    results: list[LogEntry] = []

    for entry_time, seconds in entries:
        if previous_time is not None and entry_time < previous_time:
            current_date += timedelta(days=1)

        results.append(
            LogEntry(
                entry_datetime=datetime.combine(current_date, entry_time),
                seconds=seconds,
            )
        )
        previous_time = entry_time

    return results


def process_log_file(
    log_path: Path,
    line_pattern: re.Pattern[str],
    start_date: date,
    value_divisor: float,
    min_seconds: float,
) -> list[LogEntry]:
    raw_entries: list[tuple[time, float]] = []

    with log_path.open(encoding="utf-8", errors="replace") as log_file:
        for line in log_file:
            pattern_match = line_pattern.search(line)
            if not pattern_match:
                continue

            entry_time = parse_time_from_line(line)
            if entry_time is None:
                continue

            raw_value = float(pattern_match.group(1))
            raw_entries.append((entry_time, raw_value / value_divisor))

    entries = assign_dates(raw_entries, start_date)
    if min_seconds <= 0:
        return entries
    return [entry for entry in entries if entry.seconds > min_seconds]


def sanitize_sheet_name(file_path: Path, used_names: set[str]) -> str:
    base_name = file_path.stem[:31]
    base_name = INVALID_TITLE_REGEX.sub("_", base_name).strip("'") or "Sheet"

    candidate = base_name
    suffix = 1
    while candidate in used_names:
        tail = f"_{suffix}"
        candidate = f"{base_name[: 31 - len(tail)]}{tail}"
        suffix += 1

    used_names.add(candidate)
    return candidate


SUMMARY_SHEET_NAME = "Summary"


def duration_bucket(seconds: float) -> str:
    """Single mutually exclusive bucket label for Excel filtering."""
    if seconds > 300:
        return ">5 min"
    if seconds > 240:
        return ">4 min"
    if seconds > 180:
        return ">3 min"
    if seconds > 120:
        return ">2 min"
    if seconds > 60:
        return ">1 min"
    if seconds > 30:
        return ">30s"
    if seconds > 10:
        return ">10s"
    return "<=10s"


def entry_row_values(entry: LogEntry) -> list:
    return [
        entry.entry_datetime,
        round(entry.seconds, 3),
        duration_bucket(entry.seconds),
    ]


def apply_datetime_format(sheet: Worksheet, datetime_column: int) -> None:
    for row in range(2, sheet.max_row + 1):
        sheet.cell(row=row, column=datetime_column).number_format = "yyyy-mm-dd hh:mm:ss"


def style_bucket_column(sheet: Worksheet, bucket_column: int) -> None:
    center = Alignment(horizontal="center")
    for row in range(1, sheet.max_row + 1):
        sheet.cell(row=row, column=bucket_column).alignment = center


def apply_sheet_layout(
    sheet: Worksheet,
    *,
    datetime_column: int,
    bucket_column: int,
    datetime_width: int = 22,
    seconds_width: int = 12,
    bucket_width: int = 10,
) -> None:
    apply_datetime_format(sheet, datetime_column)
    style_bucket_column(sheet, bucket_column)
    sheet.column_dimensions[get_column_letter(datetime_column)].width = datetime_width
    sheet.column_dimensions[get_column_letter(bucket_column - 1)].width = seconds_width
    sheet.column_dimensions[get_column_letter(bucket_column)].width = bucket_width
    if sheet.max_row > 1:
        sheet.auto_filter.ref = sheet.dimensions


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    entries: list[LogEntry],
) -> Worksheet:
    sheet = workbook.create_sheet(title=sheet_name)
    sheet.append(["DateTime", "Seconds", "Bucket"])

    for entry in entries:
        sheet.append(entry_row_values(entry))

    apply_sheet_layout(sheet, datetime_column=1, bucket_column=3)
    return sheet


def write_summary_sheet(
    workbook: Workbook,
    rows: list[tuple[str, LogEntry]],
) -> Worksheet:
    sheet = workbook.create_sheet(title=SUMMARY_SHEET_NAME, index=0)
    sheet.append(["Source", "DateTime", "Seconds", "Bucket"])

    for source, entry in sorted(rows, key=lambda item: (item[1].entry_datetime, item[0])):
        sheet.append([source, *entry_row_values(entry)])

    apply_sheet_layout(
        sheet,
        datetime_column=2,
        bucket_column=4,
        datetime_width=22,
        seconds_width=12,
        bucket_width=10,
    )
    sheet.column_dimensions["A"].width = 18

    return sheet


def parse_start_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Invalid start date '{value}'. Use YYYY-MM-DD (e.g. 2026-06-01)."
        ) from exc


def resolve_path(path_value: str, base_dir: Path) -> Path:
    path = Path(path_value)
    return path if path.is_absolute() else base_dir / path


def load_config_file(config_path: Path) -> ProcessorConfig:
    with config_path.open(encoding="utf-8") as config_file:
        raw = yaml.safe_load(config_file)

    if not isinstance(raw, dict):
        raise ValueError("Config root must be a mapping.")

    settings = raw.get("settings", {})
    if not isinstance(settings, dict):
        raise ValueError("'settings' must be a mapping.")

    line_pattern = raw.get("line_pattern")
    if not line_pattern:
        raise ValueError("'line_pattern' is required in the config file.")

    log_files_raw = raw.get("log_files", [])
    if not isinstance(log_files_raw, list):
        raise ValueError("'log_files' must be a list when provided in config.")

    start_date_raw = settings.get("start_date")
    start_date = date.fromisoformat(str(start_date_raw)) if start_date_raw else None

    base_dir = config_path.parent
    output_raw = settings.get("output_file", "log_analysis.xlsx")
    output_file = resolve_path(str(output_raw), base_dir)

    value_divisor = float(raw.get("value_divisor", 1000))
    if value_divisor <= 0:
        raise ValueError("'value_divisor' must be greater than zero.")

    min_seconds = float(raw.get("min_seconds", 0))
    if min_seconds < 0:
        raise ValueError("'min_seconds' must be zero or greater.")

    log_files = [resolve_path(str(item), base_dir) for item in log_files_raw]

    return ProcessorConfig(
        start_date=start_date,
        line_pattern=str(line_pattern),
        value_divisor=value_divisor,
        min_seconds=min_seconds,
        output_file=output_file,
        log_files=log_files,
    )


def merge_config(
    file_config: ProcessorConfig | None,
    cli_start_date: date | None,
    cli_pattern: str | None,
    cli_output: str | None,
    cli_log_files: list[Path] | None,
    cli_min_seconds: float | None,
    config_base_dir: Path,
) -> ProcessorConfig:
    if file_config is None:
        if cli_start_date is None or not cli_pattern or not cli_log_files:
            raise ValueError(
                "Provide -c config.yaml, or pass --start-date, --pattern, and "
                "-f/--log-file (or positional log paths)."
            )
        output_file = Path(cli_output or "log_analysis.xlsx")
        if not output_file.is_absolute():
            output_file = config_base_dir / output_file
        return ProcessorConfig(
            start_date=cli_start_date,
            line_pattern=cli_pattern,
            value_divisor=1000.0,
            min_seconds=cli_min_seconds if cli_min_seconds is not None else 0.0,
            output_file=output_file,
            log_files=[path.resolve() for path in cli_log_files],
        )

    line_pattern = cli_pattern or file_config.line_pattern
    output_file = file_config.output_file
    if cli_output:
        output_file = Path(cli_output)
        if not output_file.is_absolute():
            output_file = config_base_dir / output_file

    log_files = (
        [path.resolve() for path in cli_log_files]
        if cli_log_files
        else file_config.log_files
    )
    if not log_files:
        raise ValueError(
            "At least one log file is required. Pass -f/--log-file or positional "
            "paths, or set log_files in the config file."
        )

    merged_start_date = cli_start_date or file_config.start_date
    if merged_start_date is None:
        raise ValueError(
            "Start date is required. Pass --start-date YYYY-MM-DD or set "
            "settings.start_date in the config file."
        )

    min_seconds = (
        cli_min_seconds if cli_min_seconds is not None else file_config.min_seconds
    )

    return ProcessorConfig(
        start_date=merged_start_date,
        line_pattern=line_pattern,
        value_divisor=file_config.value_divisor,
        min_seconds=min_seconds,
        output_file=output_file.resolve(),
        log_files=log_files,
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Process up to four timestamp-only log files, extract a configurable "
            "line pattern, assign dates from a start date, and write one Excel workbook."
        )
    )
    parser.add_argument(
        "-c",
        "--config",
        type=Path,
        default=None,
        help="YAML config file (default: config.yaml in the current directory).",
    )
    parser.add_argument(
        "-d",
        "--start-date",
        type=parse_start_date,
        default=None,
        metavar="YYYY-MM-DD",
        help=(
            "First calendar date for log timestamps (e.g. 2026-06-01). "
            "Overrides settings.start_date in config when both are set."
        ),
    )
    parser.add_argument(
        "--pattern",
        default=None,
        help="Override config: line pattern with 'x' as the numeric placeholder.",
    )
    parser.add_argument(
        "--min-seconds",
        type=float,
        default=None,
        metavar="SECONDS",
        help=(
            "Only include lines where duration is greater than this many seconds. "
            "Overrides min_seconds in config (default: 0, include all)."
        ),
    )
    parser.add_argument(
        "-o",
        "--output",
        default=None,
        help="Override config: output Excel file path.",
    )
    parser.add_argument(
        "-f",
        "--log-file",
        action="append",
        type=Path,
        default=None,
        metavar="PATH",
        help=(
            "Log file to process (repeat up to 4 times). "
            "Overrides log_files in config when provided."
        ),
    )
    parser.add_argument(
        "log_files",
        nargs="*",
        type=Path,
        help="Log file paths (alternative to -f/--log-file). Up to four files.",
    )
    return parser.parse_args()


def collect_cli_log_files(args: argparse.Namespace) -> list[Path] | None:
    paths: list[Path] = []
    if args.log_file:
        paths.extend(args.log_file)
    if args.log_files:
        paths.extend(args.log_files)
    return paths or None


def main() -> int:
    args = parse_args()
    config_path = args.config

    if config_path is None:
        default_config = Path("config.yaml")
        if default_config.exists():
            config_path = default_config

    file_config: ProcessorConfig | None = None
    config_base_dir = Path.cwd()

    if config_path is not None:
        if not config_path.exists():
            print(f"Config file not found: {config_path}", file=sys.stderr)
            return 1
        try:
            file_config = load_config_file(config_path.resolve())
            config_base_dir = config_path.resolve().parent
        except (ValueError, yaml.YAMLError) as exc:
            print(f"Invalid config: {exc}", file=sys.stderr)
            return 1

    try:
        config = merge_config(
            file_config=file_config,
            cli_start_date=args.start_date,
            cli_pattern=args.pattern,
            cli_output=args.output,
            cli_log_files=collect_cli_log_files(args),
            cli_min_seconds=args.min_seconds,
            config_base_dir=config_base_dir,
        )
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if len(config.log_files) > MAX_LOG_FILES:
        print(
            f"At most {MAX_LOG_FILES} log files are supported; "
            f"received {len(config.log_files)}.",
            file=sys.stderr,
        )
        return 1

    try:
        line_pattern = pattern_to_regex(config.line_pattern)
    except ValueError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    missing = [path for path in config.log_files if not path.exists()]
    if missing:
        for path in missing:
            print(f"Log file not found: {path}", file=sys.stderr)
        return 1

    workbook = Workbook()
    workbook.remove(workbook.active)

    used_sheet_names: set[str] = {SUMMARY_SHEET_NAME}
    total_rows = 0
    summary_rows: list[tuple[str, LogEntry]] = []

    for log_path in config.log_files:
        resolved = log_path.resolve()
        entries = process_log_file(
            resolved,
            line_pattern,
            config.start_date,
            config.value_divisor,
            config.min_seconds,
        )
        sheet_name = sanitize_sheet_name(resolved, used_sheet_names)
        write_sheet(workbook, sheet_name, entries)
        source_label = resolved.stem
        summary_rows.extend((source_label, entry) for entry in entries)
        total_rows += len(entries)
        print(f"{resolved.name}: {len(entries)} matching line(s) -> sheet '{sheet_name}'")

    if not summary_rows:
        print("No matching lines found in any log file.", file=sys.stderr)
        return 1

    write_summary_sheet(workbook, summary_rows)
    print(f"Summary: {len(summary_rows)} row(s) -> sheet '{SUMMARY_SHEET_NAME}'")

    workbook.save(config.output_file)
    min_filter = (
        f"min_seconds: >{config.min_seconds}"
        if config.min_seconds > 0
        else "min_seconds: none"
    )
    print(
        f"\nPattern: {config.line_pattern!r}"
        f"\n{min_filter}"
        f"\nWrote {config.output_file} ({len(workbook.sheetnames)} sheet(s), {total_rows} row(s))."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
