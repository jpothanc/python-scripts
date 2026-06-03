#!/usr/bin/env python3
"""Monitor REST APIs at configured intervals and log results to CSV."""

from __future__ import annotations

import argparse
import csv
import signal
import sys
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests
import yaml
from requests.exceptions import RequestException, Timeout

CSV_HEADERS = [
    "timestamp_utc",
    "api_name",
    "url",
    "method",
    "http_status",
    "response_time_ms",
    "result",
    "error_message",
]


@dataclass(frozen=True)
class ApiTarget:
    name: str
    url: str
    method: str = "GET"
    headers: dict[str, str] | None = None


@dataclass(frozen=True)
class MonitorSettings:
    duration_hours: float
    interval_seconds: float
    timeout_seconds: float
    output_file: Path


@dataclass(frozen=True)
class CheckResult:
    timestamp_utc: str
    api_name: str
    url: str
    method: str
    http_status: str
    response_time_ms: str
    result: str
    error_message: str


class ApiMonitor:
    def __init__(self, settings: MonitorSettings, apis: list[ApiTarget]) -> None:
        self.settings = settings
        self.apis = apis
        self._stop_requested = False
        self._session = requests.Session()

    def request_stop(self) -> None:
        self._stop_requested = True

    def run(self) -> int:
        end_time = time.monotonic() + (self.settings.duration_hours * 3600)
        round_number = 0

        print(
            f"Starting monitor for {self.settings.duration_hours} hour(s). "
            f"Interval: {self.settings.interval_seconds}s. "
            f"Timeout: {self.settings.timeout_seconds}s. "
            f"Output: {self.settings.output_file}"
        )

        with self.settings.output_file.open("w", newline="", encoding="utf-8") as csv_file:
            writer = csv.DictWriter(csv_file, fieldnames=CSV_HEADERS)
            writer.writeheader()
            csv_file.flush()

            while not self._stop_requested and time.monotonic() < end_time:
                round_number += 1
                round_started = datetime.now(timezone.utc).isoformat()
                print(f"\nRound {round_number} started at {round_started}")

                for api in self.apis:
                    if self._stop_requested or time.monotonic() >= end_time:
                        break

                    result = self._check_api(api)
                    writer.writerow(result.__dict__)
                    csv_file.flush()
                    self._print_result(result)

                if self._stop_requested or time.monotonic() >= end_time:
                    break

                remaining = end_time - time.monotonic()
                sleep_for = min(self.settings.interval_seconds, remaining)
                if sleep_for > 0:
                    print(f"Sleeping for {sleep_for:.1f}s before next round...")
                    time.sleep(sleep_for)

        self._print_summary()
        return 0 if not self._stop_requested else 130

    def _check_api(self, api: ApiTarget) -> CheckResult:
        started = time.perf_counter()
        timestamp = datetime.now(timezone.utc).isoformat()

        try:
            response = self._session.request(
                method=api.method.upper(),
                url=api.url,
                headers=api.headers,
                timeout=self.settings.timeout_seconds,
            )
            elapsed_ms = (time.perf_counter() - started) * 1000
            return CheckResult(
                timestamp_utc=timestamp,
                api_name=api.name,
                url=api.url,
                method=api.method.upper(),
                http_status=str(response.status_code),
                response_time_ms=f"{elapsed_ms:.2f}",
                result="success",
                error_message="",
            )
        except Timeout:
            elapsed_ms = (time.perf_counter() - started) * 1000
            return CheckResult(
                timestamp_utc=timestamp,
                api_name=api.name,
                url=api.url,
                method=api.method.upper(),
                http_status="",
                response_time_ms=f"{elapsed_ms:.2f}",
                result="timeout",
                error_message=f"Request exceeded {self.settings.timeout_seconds}s timeout",
            )
        except RequestException as exc:
            elapsed_ms = (time.perf_counter() - started) * 1000
            return CheckResult(
                timestamp_utc=timestamp,
                api_name=api.name,
                url=api.url,
                method=api.method.upper(),
                http_status="",
                response_time_ms=f"{elapsed_ms:.2f}",
                result="error",
                error_message=str(exc),
            )

    def _print_result(self, result: CheckResult) -> None:
        if result.result == "success":
            print(
                f"  [{result.api_name}] {result.http_status} "
                f"in {result.response_time_ms}ms"
            )
        else:
            print(
                f"  [{result.api_name}] {result.result.upper()} "
                f"after {result.response_time_ms}ms - {result.error_message}"
            )

    def _print_summary(self) -> None:
        if not self.settings.output_file.exists():
            print("No results file found.")
            return

        counts: dict[str, int] = {"success": 0, "timeout": 0, "error": 0}
        per_api: dict[str, dict[str, int]] = {}

        with self.settings.output_file.open(newline="", encoding="utf-8") as csv_file:
            reader = csv.DictReader(csv_file)
            for row in reader:
                result = row["result"]
                counts[result] = counts.get(result, 0) + 1

                api_name = row["api_name"]
                if api_name not in per_api:
                    per_api[api_name] = {"success": 0, "timeout": 0, "error": 0}
                per_api[api_name][result] = per_api[api_name].get(result, 0) + 1

        total = sum(counts.values())
        print("\n=== Summary ===")
        print(f"Total checks: {total}")
        print(f"Successful:   {counts.get('success', 0)}")
        print(f"Timeouts:     {counts.get('timeout', 0)}")
        print(f"Errors:       {counts.get('error', 0)}")

        if per_api:
            print("\nPer API:")
            for api_name in sorted(per_api):
                stats = per_api[api_name]
                print(
                    f"  {api_name}: "
                    f"success={stats.get('success', 0)}, "
                    f"timeout={stats.get('timeout', 0)}, "
                    f"error={stats.get('error', 0)}"
                )


def load_config(config_path: Path) -> tuple[MonitorSettings, list[ApiTarget]]:
    with config_path.open(encoding="utf-8") as config_file:
        raw = yaml.safe_load(config_file)

    if not isinstance(raw, dict):
        raise ValueError("Config root must be a mapping.")

    settings_raw = raw.get("settings", {})
    apis_raw = raw.get("apis", [])

    if not isinstance(settings_raw, dict):
        raise ValueError("'settings' must be a mapping.")
    if not isinstance(apis_raw, list) or not apis_raw:
        raise ValueError("'apis' must be a non-empty list.")

    output_file = Path(settings_raw.get("output_file", "results.csv"))
    if not output_file.is_absolute():
        output_file = config_path.parent / output_file

    settings = MonitorSettings(
        duration_hours=float(settings_raw.get("duration_hours", 24)),
        interval_seconds=float(settings_raw.get("interval_seconds", 60)),
        timeout_seconds=float(settings_raw.get("timeout_seconds", 10)),
        output_file=output_file,
    )

    apis: list[ApiTarget] = []
    for index, api_raw in enumerate(apis_raw, start=1):
        if not isinstance(api_raw, dict):
            raise ValueError(f"API entry #{index} must be a mapping.")

        name = api_raw.get("name")
        url = api_raw.get("url")
        if not name or not url:
            raise ValueError(f"API entry #{index} requires 'name' and 'url'.")

        headers = api_raw.get("headers")
        if headers is not None and not isinstance(headers, dict):
            raise ValueError(f"API '{name}' headers must be a mapping.")

        apis.append(
            ApiTarget(
                name=str(name),
                url=str(url),
                method=str(api_raw.get("method", "GET")),
                headers=headers,
            )
        )

    return settings, apis


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Call configured REST APIs on an interval and log results to CSV."
    )
    parser.add_argument(
        "-c",
        "--config",
        default="config.yaml",
        help="Path to YAML config file (default: config.yaml)",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    config_path = Path(args.config).resolve()

    if not config_path.exists():
        print(f"Config file not found: {config_path}", file=sys.stderr)
        return 1

    try:
        settings, apis = load_config(config_path)
    except (ValueError, yaml.YAMLError) as exc:
        print(f"Invalid config: {exc}", file=sys.stderr)
        return 1

    monitor = ApiMonitor(settings, apis)

    def handle_signal(signum: int, _frame: Any) -> None:
        print(f"\nReceived signal {signum}. Finishing current round and exiting...")
        monitor.request_stop()

    signal.signal(signal.SIGINT, handle_signal)
    if hasattr(signal, "SIGTERM"):
        signal.signal(signal.SIGTERM, handle_signal)

    return monitor.run()


if __name__ == "__main__":
    raise SystemExit(main())
