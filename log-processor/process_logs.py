#!/usr/bin/env python3
"""Process timestamp-only log files and export matched metrics to Excel."""

from __future__ import annotations

import argparse
import re
import statistics
import sys
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
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


@dataclass(frozen=True)
class SourceInsights:
    counts: dict[str, int]
    total: int
    avg_seconds: float
    median_seconds: float
    p95_seconds: float
    max_seconds: float
    pct_over_1min: float
    pct_over_5min: float
    first_seen: datetime | None
    last_seen: datetime | None


@dataclass(frozen=True)
class SummaryContext:
    start_date: date
    line_pattern: str
    min_seconds: float
    output_file: Path
    log_files: list[Path]
    generated_at: datetime


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


ALL_SHEET_NAME = "All"
SUMMARY_SHEET_NAME = "Summary"

# Summary sheet styling
SUMMARY_TITLE_FILL = PatternFill("solid", fgColor="1E3A5F")
SUMMARY_TITLE_FONT = Font(bold=True, size=16, color="FFFFFF")
SUMMARY_LABEL_FILLS = (
    PatternFill("solid", fgColor="2563EB"),
    PatternFill("solid", fgColor="3B82F6"),
    PatternFill("solid", fgColor="0EA5E9"),
    PatternFill("solid", fgColor="06B6D4"),
    PatternFill("solid", fgColor="14B8A6"),
    PatternFill("solid", fgColor="10B981"),
)
SUMMARY_VALUE_FILLS = (
    PatternFill("solid", fgColor="DBEAFE"),
    PatternFill("solid", fgColor="E0F2FE"),
    PatternFill("solid", fgColor="CFFAFE"),
    PatternFill("solid", fgColor="CCFBF1"),
    PatternFill("solid", fgColor="D1FAE5"),
    PatternFill("solid", fgColor="DCFCE7"),
)
SUMMARY_LABEL_FONT = Font(bold=True, color="FFFFFF", size=11)
SUMMARY_VALUE_FONT = Font(color="1E293B", size=11)
TABLE_HEADER_FILL = PatternFill("solid", fgColor="6D28D9")
TABLE_HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
TABLE_ROW_FILL_A = PatternFill("solid", fgColor="FFFFFF")
TABLE_ROW_FILL_B = PatternFill("solid", fgColor="F5F3FF")
TABLE_TOTAL_FILL = PatternFill("solid", fgColor="FEF3C7")
TABLE_TOTAL_FONT = Font(bold=True, color="92400E", size=10)
BUCKET_HOT_FILL = PatternFill("solid", fgColor="FDE68A")
THIN_BORDER = Border(
    left=Side(style="thin", color="CBD5E1"),
    right=Side(style="thin", color="CBD5E1"),
    top=Side(style="thin", color="CBD5E1"),
    bottom=Side(style="thin", color="CBD5E1"),
)

BUCKET_LABELS: tuple[str, ...] = (
    "<=10s",
    ">10s",
    ">30s",
    ">1min",
    ">2min",
    ">3min",
    ">4min",
    ">5min",
)


def duration_bucket(seconds: float) -> str:
    """Single mutually exclusive bucket label for Excel filtering."""
    if seconds > 300:
        return ">5min"
    if seconds > 240:
        return ">4min"
    if seconds > 180:
        return ">3min"
    if seconds > 120:
        return ">2min"
    if seconds > 60:
        return ">1min"
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


def count_buckets(entries: list[LogEntry]) -> dict[str, int]:
    counts = dict.fromkeys(BUCKET_LABELS, 0)
    for entry in entries:
        label = duration_bucket(entry.seconds)
        if label in counts:
            counts[label] += 1
    return counts


def format_bucket_details(counts: dict[str, int]) -> str:
    parts = [f"{label}: {counts[label]} times" for label in BUCKET_LABELS if counts[label] > 0]
    return ", ".join(parts) if parts else "no entries"


def percentile(seconds: list[float], pct: float) -> float:
    if not seconds:
        return 0.0
    if len(seconds) == 1:
        return seconds[0]
    ordered = sorted(seconds)
    rank = (len(ordered) - 1) * pct / 100
    low = int(rank)
    high = min(low + 1, len(ordered) - 1)
    weight = rank - low
    return ordered[low] * (1 - weight) + ordered[high] * weight


def build_source_insights(entries: list[LogEntry]) -> SourceInsights:
    if not entries:
        empty_counts = dict.fromkeys(BUCKET_LABELS, 0)
        return SourceInsights(
            counts=empty_counts,
            total=0,
            avg_seconds=0.0,
            median_seconds=0.0,
            p95_seconds=0.0,
            max_seconds=0.0,
            pct_over_1min=0.0,
            pct_over_5min=0.0,
            first_seen=None,
            last_seen=None,
        )

    durations = [entry.seconds for entry in entries]
    datetimes = [entry.entry_datetime for entry in entries]
    total = len(entries)
    over_1min = sum(1 for value in durations if value > 60)
    over_5min = sum(1 for value in durations if value > 300)

    return SourceInsights(
        counts=count_buckets(entries),
        total=total,
        avg_seconds=statistics.mean(durations),
        median_seconds=statistics.median(durations),
        p95_seconds=percentile(durations, 95),
        max_seconds=max(durations),
        pct_over_1min=round(100 * over_1min / total, 1),
        pct_over_5min=round(100 * over_5min / total, 1),
        first_seen=min(datetimes),
        last_seen=max(datetimes),
    )


def style_cell(
    cell,
    *,
    fill: PatternFill | None = None,
    font: Font | None = None,
    alignment: Alignment | None = None,
    border: Border | None = None,
) -> None:
    if fill is not None:
        cell.fill = fill
    if font is not None:
        cell.font = font
    if alignment is not None:
        cell.alignment = alignment
    if border is not None:
        cell.border = border


def write_summary_info_block(sheet: Worksheet, context: SummaryContext) -> int:
    """Write run context at the top; return the row number where the data table starts."""
    sheet.merge_cells("A1:B1")
    title_cell = sheet["A1"]
    title_cell.value = "Log Analysis Summary"
    style_cell(
        title_cell,
        fill=SUMMARY_TITLE_FILL,
        font=SUMMARY_TITLE_FONT,
        alignment=Alignment(horizontal="center", vertical="center"),
        border=THIN_BORDER,
    )
    sheet.row_dimensions[1].height = 28

    min_filter = (
        f">{context.min_seconds}s only"
        if context.min_seconds > 0
        else "none (all matching lines)"
    )
    log_file_names = ", ".join(path.name for path in context.log_files)

    info_rows = [
        ("Generated", context.generated_at.strftime("%Y-%m-%d %H:%M:%S")),
        ("Start date", context.start_date.isoformat()),
        ("Line pattern", context.line_pattern),
        ("Min filter", min_filter),
        ("Output file", context.output_file.name),
        ("Log files", log_file_names),
    ]

    for index, (label, value) in enumerate(info_rows):
        row = index + 2
        color_index = index % len(SUMMARY_LABEL_FILLS)
        label_cell = sheet.cell(row=row, column=1, value=label)
        value_cell = sheet.cell(row=row, column=2, value=value)
        style_cell(
            label_cell,
            fill=SUMMARY_LABEL_FILLS[color_index],
            font=SUMMARY_LABEL_FONT,
            alignment=Alignment(horizontal="left", vertical="center"),
            border=THIN_BORDER,
        )
        style_cell(
            value_cell,
            fill=SUMMARY_VALUE_FILLS[color_index],
            font=SUMMARY_VALUE_FONT,
            alignment=Alignment(horizontal="left", vertical="center", wrap_text=True),
            border=THIN_BORDER,
        )
        sheet.row_dimensions[row].height = 22

    sheet.column_dimensions["A"].width = 18
    sheet.column_dimensions["B"].width = 54

    return len(info_rows) + 3


def write_all_sheet(
    workbook: Workbook,
    rows: list[tuple[str, LogEntry]],
) -> Worksheet:
    sheet = workbook.create_sheet(title=ALL_SHEET_NAME, index=0)
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


def write_bucket_summary_sheet(
    workbook: Workbook,
    source_entries: list[tuple[str, list[LogEntry]]],
    context: SummaryContext,
) -> Worksheet:
    sheet = workbook.create_sheet(title=SUMMARY_SHEET_NAME, index=0)
    sheet.sheet_properties.tabColor = "6D28D9"
    table_start = write_summary_info_block(sheet, context)

    metric_headers = [
        "Avg (s)",
        "Median (s)",
        "P95 (s)",
        "Max (s)",
        "%>1 min",
        "%>5 min",
        "First seen",
        "Last seen",
        "Details",
    ]
    headers = ["Source", *BUCKET_LABELS, "Total", *metric_headers]
    bucket_col_start = 2
    bucket_col_end = bucket_col_start + len(BUCKET_LABELS) - 1

    for col, header in enumerate(headers, start=1):
        cell = sheet.cell(row=table_start, column=col, value=header)
        style_cell(
            cell,
            fill=TABLE_HEADER_FILL,
            font=TABLE_HEADER_FONT,
            alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
            border=THIN_BORDER,
        )
    sheet.row_dimensions[table_start].height = 24

    center = Alignment(horizontal="center", vertical="center")
    data_row = table_start + 1
    source_row_index = 0

    def style_data_row(row_number: int, *, is_total: bool = False) -> None:
        row_fill = TABLE_TOTAL_FILL if is_total else (
            TABLE_ROW_FILL_B if source_row_index % 2 else TABLE_ROW_FILL_A
        )
        row_font = TABLE_TOTAL_FONT if is_total else Font(color="334155", size=10)
        for col in range(1, len(headers) + 1):
            cell = sheet.cell(row=row_number, column=col)
            style_cell(
                cell,
                fill=row_fill,
                font=row_font,
                alignment=center if col > 1 else Alignment(horizontal="left", vertical="center"),
                border=THIN_BORDER,
            )

    for source, entries in source_entries:
        insights = build_source_insights(entries)
        row_values = [
            source,
            *[insights.counts[label] for label in BUCKET_LABELS],
            insights.total,
            round(insights.avg_seconds, 3),
            round(insights.median_seconds, 3),
            round(insights.p95_seconds, 3),
            round(insights.max_seconds, 3),
            insights.pct_over_1min,
            insights.pct_over_5min,
            insights.first_seen,
            insights.last_seen,
            format_bucket_details(insights.counts),
        ]
        for col, value in enumerate(row_values, start=1):
            sheet.cell(row=data_row, column=col, value=value)
        style_data_row(data_row)
        for col in range(bucket_col_start, bucket_col_end + 1):
            cell = sheet.cell(row=data_row, column=col)
            if isinstance(cell.value, int) and cell.value > 0:
                cell.fill = BUCKET_HOT_FILL
        source_row_index += 1
        data_row += 1

    all_entries = [entry for _, entries in source_entries for entry in entries]
    if len(source_entries) > 1 and all_entries:
        combined = build_source_insights(all_entries)
        row_values = [
            "ALL SOURCES",
            *[combined.counts[label] for label in BUCKET_LABELS],
            combined.total,
            round(combined.avg_seconds, 3),
            round(combined.median_seconds, 3),
            round(combined.p95_seconds, 3),
            round(combined.max_seconds, 3),
            combined.pct_over_1min,
            combined.pct_over_5min,
            combined.first_seen,
            combined.last_seen,
            format_bucket_details(combined.counts),
        ]
        for col, value in enumerate(row_values, start=1):
            sheet.cell(row=data_row, column=col, value=value)
        style_data_row(data_row, is_total=True)
        for col in range(bucket_col_start, bucket_col_end + 1):
            cell = sheet.cell(row=data_row, column=col)
            if isinstance(cell.value, int) and cell.value > 0:
                cell.fill = PatternFill("solid", fgColor="FCD34D")
        data_row += 1

    first_seen_col = len(headers) - 2
    last_seen_col = len(headers) - 1
    for row in range(table_start + 1, data_row):
        sheet.cell(row=row, column=first_seen_col).number_format = "yyyy-mm-dd hh:mm:ss"
        sheet.cell(row=row, column=last_seen_col).number_format = "yyyy-mm-dd hh:mm:ss"
        sheet.row_dimensions[row].height = 20

    for col in range(2, len(headers)):
        sheet.column_dimensions[get_column_letter(col)].width = 10
    sheet.column_dimensions["A"].width = 18
    sheet.column_dimensions[get_column_letter(len(headers))].width = 40

    sheet.freeze_panes = sheet.cell(row=table_start + 1, column=1)

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

    used_sheet_names: set[str] = {SUMMARY_SHEET_NAME, ALL_SHEET_NAME}
    total_rows = 0
    all_rows: list[tuple[str, LogEntry]] = []
    source_entries: list[tuple[str, list[LogEntry]]] = []

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
        source_entries.append((source_label, entries))
        all_rows.extend((source_label, entry) for entry in entries)
        total_rows += len(entries)
        print(f"{resolved.name}: {len(entries)} matching line(s) -> sheet '{sheet_name}'")

    if not all_rows:
        print("No matching lines found in any log file.", file=sys.stderr)
        return 1

    write_all_sheet(workbook, all_rows)
    summary_context = SummaryContext(
        start_date=config.start_date,
        line_pattern=config.line_pattern,
        min_seconds=config.min_seconds,
        output_file=config.output_file,
        log_files=config.log_files,
        generated_at=datetime.now(),
    )
    write_bucket_summary_sheet(workbook, source_entries, summary_context)
    print(f"All: {len(all_rows)} row(s) -> sheet '{ALL_SHEET_NAME}'")
    print(f"Summary: bucket counts for {len(source_entries)} source(s) -> sheet '{SUMMARY_SHEET_NAME}'")

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
