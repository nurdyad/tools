#!/usr/bin/env python3
from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import subprocess
import sys
import time
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import quote_plus

import requests
from dotenv import load_dotenv
from robocorp import windows
from robocorp.windows import ActionNotPossible, ElementNotFound


COUNT_PATTERN = re.compile(r"\((\d+)\s*,\s*\d+\)\s*$")
PAGE_PATTERN = re.compile(r"Page\s+(\d+)\s+of\s+(\d+)", re.IGNORECASE)
DOC_TYPE_ROW_PATTERN = re.compile(r"Document Type Row (\d+)")
ODS_PATTERN = re.compile(r"^[A-Za-z0-9]{5,8}$")

CDB_SWITCHER_EXE = r"C:\ProgramData\SDS\Version6\Applications\SDS Tools\SDS.Client.ConfigurationSwitcher.exe"
CDB_SWITCHER_NAME = "name: EMIS Configuration Switcher"
EMIS_AUTH_WINDOW = "name: Authentication"
EMIS_WINDOW_REGEX = "regex:EMIS Web Health Care System.*"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Standalone EMIS letter summary tool with automatic practice login."
    )
    parser.add_argument(
        "--practice",
        default=os.getenv("PRACTICE_NAME", ""),
        help="Practice name (or ODS code).",
    )
    parser.add_argument(
        "--queues",
        default="Awaiting Filing",
        help='Comma-separated queue names (e.g. "Unmatched,Awaiting Filing") or "all".',
    )
    parser.add_argument(
        "--output",
        default="",
        help="Path to JSON report. Defaults to ./output/emis_letter_type_summary_<timestamp>.json",
    )
    parser.add_argument(
        "--csv-output",
        default="",
        help="Optional path to CSV report.",
    )
    parser.add_argument("--timeout", type=int, default=120, help="Wait timeout seconds.")
    parser.add_argument("--quiet", action="store_true", help="Suppress per-page progress logs.")
    parser.add_argument(
        "--strict-sidebar-match",
        action="store_true",
        help="Exit non-zero if scanned queue count differs from sidebar total.",
    )

    parser.add_argument("--skip-login", action="store_true", help="Attach to existing EMIS session.")
    parser.add_argument("--close-on-exit", action="store_true", help="Close EMIS window at end.")
    parser.add_argument(
        "--kill-existing",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Kill EmisWeb.exe/WINWORD.exe before login (default true).",
    )
    parser.add_argument(
        "--clear-cache",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Clear EMIS cache before login (default true).",
    )

    parser.add_argument("--ods-code", default=os.getenv("ODS_CODE", ""), help="Manual ODS override.")
    parser.add_argument(
        "--mailroom-api-url",
        default=os.getenv("MAILROOM_API_URL", ""),
        help="Mailroom API URL.",
    )
    parser.add_argument(
        "--mailroom-api-key",
        default=os.getenv("MAILROOM_API_KEY", ""),
        help="Mailroom API key.",
    )
    parser.add_argument("--cdb", default=os.getenv("EMIS_CDB", ""), help="Manual CDB override.")
    parser.add_argument(
        "--emis-username",
        default=os.getenv("EMIS_USERNAME", ""),
        help="Manual EMIS username override.",
    )
    parser.add_argument(
        "--emis-password",
        default=os.getenv("EMIS_PASSWORD", ""),
        help="Manual EMIS password override.",
    )

    return parser.parse_args()


def setup_logging(logs_dir: Path) -> logging.Logger:
    logs_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = logs_dir / f"emis_letter_summary_{timestamp}.log"

    logger = logging.getLogger("emis_letter_summary")
    logger.setLevel(logging.INFO)
    logger.handlers = []

    formatter = logging.Formatter("%(asctime)s %(levelname)s %(message)s")

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    logger.info("Log file: %s", log_path)
    return logger


def queue_name_without_count(raw_name: str) -> str:
    return COUNT_PATTERN.sub("", raw_name).strip()


def parse_sidebar_total(raw_name: str) -> int | None:
    match = COUNT_PATTERN.search(raw_name)
    return int(match.group(1)) if match else None


def normalise_letter_type(raw_type: str | None) -> str:
    value = (raw_type or "").strip()
    return value if value else "(blank)"


def normalize_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", value.lower())


def looks_like_ods(value: str) -> bool:
    candidate = value.strip().upper()
    if not ODS_PATTERN.fullmatch(candidate):
        return False
    # ODS codes always contain digits; this avoids treating practice names
    # like "Ashlea" as ODS codes.
    return any(ch.isdigit() for ch in candidate)


def mailroom_get(url: str, api_key: str) -> Any:
    response = requests.get(url, headers={"Authorization": f"Bearer {api_key}"}, timeout=30)
    if response.status_code != 200:
        raise RuntimeError(f"Mailroom call failed [{response.status_code}] {url}: {response.text}")
    return response.json()


def extract_practice_records(payload: Any) -> list[dict[str, Any]]:
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    if isinstance(payload, dict):
        for key in ("practices", "items", "results", "data"):
            value = payload.get(key)
            if isinstance(value, list):
                return [item for item in value if isinstance(item, dict)]
    return []


def practice_display_candidates(record: dict[str, Any]) -> list[str]:
    keys = [
        "name",
        "practice_name",
        "display_name",
        "docman_practice_display_name",
        "title",
    ]
    values = []
    for key in keys:
        value = record.get(key)
        if isinstance(value, str) and value.strip():
            values.append(value.strip())
    return values


def practice_ods(record: dict[str, Any]) -> str:
    for key in ("ods_code", "odsCode", "ods", "code", "practice_id", "practiceId"):
        value = record.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def resolve_ods_from_practice_name(
    practice_name: str,
    mailroom_api_url: str,
    mailroom_api_key: str,
    logger: logging.Logger,
) -> str:
    name = practice_name.strip()
    if not name:
        raise ValueError("Practice name is required. Use --practice or PRACTICE_NAME env.")
    if looks_like_ods(name):
        return name.upper()

    if not mailroom_api_url or not mailroom_api_key:
        raise ValueError(
            "MAILROOM_API_URL and MAILROOM_API_KEY are required to resolve practice name to ODS."
        )

    base = mailroom_api_url.rstrip("/")
    encoded = quote_plus(name)
    endpoints = [
        f"{base}/practices?search={encoded}",
        f"{base}/practices?name={encoded}",
        f"{base}/practices/search?name={encoded}",
        f"{base}/practices/search?query={encoded}",
        f"{base}/practices",
    ]

    requested_norm = normalize_text(name)
    best_score = -1
    best_ods = ""
    best_label = ""

    for endpoint in endpoints:
        try:
            payload = mailroom_get(endpoint, mailroom_api_key)
        except Exception as error:
            logger.info("Practice lookup endpoint failed: %s (%s)", endpoint, error)
            continue

        records = extract_practice_records(payload)
        if not records:
            continue

        for record in records:
            ods = practice_ods(record)
            if not ods:
                continue
            for label in practice_display_candidates(record):
                label_norm = normalize_text(label)
                if not label_norm:
                    continue

                if label_norm == requested_norm:
                    score = 100
                elif requested_norm in label_norm or label_norm in requested_norm:
                    score = 80
                else:
                    score = 0

                if score > best_score:
                    best_score = score
                    best_ods = ods
                    best_label = label

    if not best_ods:
        raise RuntimeError(
            "Could not resolve practice name to ODS from Mailroom. "
            "Try passing --ods-code explicitly."
        )

    logger.info("Resolved practice '%s' -> ODS '%s' using match '%s'", name, best_ods, best_label)
    return best_ods


def fetch_ehr_settings(ods_code: str, mailroom_api_url: str, mailroom_api_key: str) -> dict[str, Any]:
    if not mailroom_api_url or not mailroom_api_key:
        raise ValueError("MAILROOM_API_URL and MAILROOM_API_KEY are required for EHR settings lookup.")
    endpoint = f"{mailroom_api_url.rstrip('/')}/practices/{ods_code}/ehr_settings"
    payload = mailroom_get(endpoint, mailroom_api_key)
    if not isinstance(payload, dict):
        raise RuntimeError("Unexpected payload from ehr_settings endpoint")
    return payload


def resolve_login_details(args: argparse.Namespace, logger: logging.Logger) -> tuple[str, str, str, str]:
    if args.skip_login:
        return "", "", "", "attached-session"

    if args.cdb and args.emis_username and args.emis_password:
        return args.cdb.strip(), args.emis_username.strip(), args.emis_password.strip(), "manual"

    ods = args.ods_code.strip()
    if not ods:
        ods = resolve_ods_from_practice_name(
            practice_name=args.practice,
            mailroom_api_url=args.mailroom_api_url,
            mailroom_api_key=args.mailroom_api_key,
            logger=logger,
        )

    settings = fetch_ehr_settings(ods, args.mailroom_api_url, args.mailroom_api_key)
    cdb = (settings.get("practice_cdb") or "").strip()
    emis_web = settings.get("emis_web") or {}
    username = (emis_web.get("username") or "").strip()
    password = (emis_web.get("password") or "").strip()

    missing = [
        name
        for name, value in (
            ("practice_cdb", cdb),
            ("emis_web.username", username),
            ("emis_web.password", password),
        )
        if not value
    ]
    if missing:
        raise RuntimeError("Missing required EHR settings fields: " + ", ".join(missing))

    return cdb, username, password, f"mailroom:{ods}"


def safe_run(command: list[str]) -> None:
    try:
        subprocess.run(command, check=False, capture_output=True, text=True)
    except Exception:
        pass


def kill_existing_processes() -> None:
    safe_run(["taskkill", "/f", "/im", "EmisWeb.exe"])
    safe_run(["taskkill", "/f", "/im", "WINWORD.exe"])


def clear_emis_cache() -> None:
    safe_run(["cmd", "/c", "rmdir", "/S", "/Q", r"C:\ProgramData\EMIS\ResourcePublisherSQLite"])


def start_emis_with_cdb(cdb: str) -> None:
    switcher = windows.find_window(CDB_SWITCHER_NAME, timeout=5, raise_error=False)
    if switcher is None:
        subprocess.Popen([CDB_SWITCHER_EXE], close_fds=True)
        switcher = windows.find_window(CDB_SWITCHER_NAME, timeout=15, raise_error=True)

    switcher.send_keys(locator="id:1001", keys=cdb)
    switcher.click("id:btnGo")


def login_to_emis(username: str, password: str, timeout: int) -> None:
    auth_window = windows.find_window(EMIS_AUTH_WINDOW, timeout=timeout, raise_error=False)
    if auth_window is None:
        return
    auth_window.send_keys(locator="id:textBoxUserName", keys=username)
    auth_window.send_keys(locator="id:textBoxPassword", keys=password)
    auth_window.click("id:buttonLogin")


def find_emis_window(timeout: int, cdb: str = ""):
    if cdb:
        specific = windows.find_window(
            f"regex:EMIS Web Health Care System.*{re.escape(cdb)}.*",
            timeout=timeout,
            raise_error=False,
            foreground=True,
        )
        if specific is not None:
            return specific
    return windows.find_window(
        EMIS_WINDOW_REGEX,
        timeout=timeout,
        raise_error=True,
        foreground=True,
    )


def add_quick_nav_button(emis_window, button_name: str) -> None:
    file_button = emis_window.find("control:ButtonControl and name:File")
    buttons = file_button.get_parent().find_many("control:ButtonControl")
    buttons[-1].click()
    windows.desktop().click("name:Customize Quick Access Toolbar...", timeout=2)

    custom = emis_window.find_child_window("id:CustomizeDialog")
    while True:
        try:
            custom.click(locator=f"name: {button_name}")
            break
        except ActionNotPossible:
            scroll_bar = custom.find("id:NonClientVerticalScrollBar", timeout=1)
            scroll_bar.click("id:DownPageButton", timeout=1)
    custom.click(locator="id:buttonAdd")
    custom.click(locator="id:buttonOK")


def ensure_workflow_manager_visible(emis_window) -> None:
    try:
        emis_window.find('control:ButtonControl and name:"Workflow Manager"', timeout=2)
    except ElementNotFound:
        add_quick_nav_button(emis_window, "Workflow Manager")


def open_workflow_manager(emis_window) -> None:
    emis_window.click('control:ButtonControl and name:"Workflow Manager"')
    emis_window.find("id:WorkflowManagerPage", timeout=60)


def open_document_management_panel(emis_window):
    panel = emis_window.find("id:DocumentManagement_itemPanel", timeout=2, raise_error=False)
    if panel is not None:
        return panel
    emis_window.click('type:Pane and name:"Document Management"')
    return emis_window.find("id:DocumentManagement_itemPanel", timeout=15)


def get_available_queue_names(panel) -> list[str]:
    names = []
    seen = set()
    for item in panel.find_many("type:ListItem", search_strategy="all"):
        raw_name = (item.name or "").strip()
        if not raw_name:
            continue
        base = queue_name_without_count(raw_name)
        if not base:
            continue
        key = base.lower()
        if key in seen:
            continue
        seen.add(key)
        names.append(base)
    return names


def resolve_target_queues(panel, queues_arg: str) -> list[str]:
    available = get_available_queue_names(panel)
    lookup = {x.lower(): x for x in available}

    if queues_arg.strip().lower() == "all":
        return available

    requested = [x.strip() for x in queues_arg.split(",") if x.strip()]
    if not requested:
        raise ValueError("No queues provided")

    missing = [x for x in requested if x.lower() not in lookup]
    if missing:
        raise ValueError(
            f"Queue(s) not found: {', '.join(missing)}. Available: {', '.join(available)}"
        )

    return [lookup[x.lower()] for x in requested]


def select_queue(panel, queue_name: str) -> str:
    for item in panel.find_many("type:ListItem", search_strategy="all"):
        raw_name = (item.name or "").strip()
        if not raw_name:
            continue
        if queue_name_without_count(raw_name).lower() != queue_name.lower():
            continue
        item.click()
        return raw_name
    raise ValueError(f"Could not select queue '{queue_name}'")


def read_page_state(emis_window):
    nav = emis_window.find("id:pageNavigatorWorkflow", search_depth=9, raise_error=False)
    if nav is None:
        return 1, 1, None, None

    page_input = nav.find("type:Edit", raise_error=False)
    raw = ""
    if page_input is not None:
        try:
            raw = page_input.ui_automation_control.GetValuePattern().Value
        except Exception:
            raw = page_input.get_value()

    match = PAGE_PATTERN.search(raw or "")
    if match:
        page = int(match.group(1))
        total = int(match.group(2))
    else:
        page, total = 1, 1

    buttons = nav.find_many("type:Button")
    prev_btn = buttons[0] if len(buttons) >= 1 else None
    next_btn = buttons[2] if len(buttons) >= 3 else None
    return page, total, prev_btn, next_btn


def control_enabled(control) -> bool:
    if control is None:
        return False
    try:
        return control.ui_automation_control.IsEnabled == 1
    except Exception:
        return True


def wait_for_page_change(emis_window, previous_page: int, timeout_seconds: float = 8.0) -> bool:
    start = time.time()
    while time.time() - start < timeout_seconds:
        current_page, _, _, _ = read_page_state(emis_window)
        if current_page != previous_page:
            return True
        time.sleep(0.2)
    return False


def move_to_first_page(emis_window, max_steps: int = 100) -> None:
    steps = 0
    while steps < max_steps:
        current_page, _, prev_btn, _ = read_page_state(emis_window)
        if current_page <= 1:
            return
        if not control_enabled(prev_btn):
            return
        prev_btn.click()
        wait_for_page_change(emis_window, previous_page=current_page)
        steps += 1


def scan_document_type_rows(table) -> list[tuple[int, str]]:
    rows = []
    seen_indexes = set()
    for row in table.find_many('regex:"Document Type Row [0-9]+"', search_strategy="all"):
        match = DOC_TYPE_ROW_PATTERN.search(row.name or "")
        if not match:
            continue
        idx = int(match.group(1))
        if idx in seen_indexes:
            continue
        seen_indexes.add(idx)
        try:
            value = normalise_letter_type(row.get_value())
        except Exception:
            value = "(blank)"
        rows.append((idx, value))
    return rows


def sort_counter(counter: Counter) -> dict[str, int]:
    keys = sorted(counter.keys(), key=lambda x: (-counter[x], x.lower()))
    return {k: counter[k] for k in keys}


def summarise_queue(emis_window, panel, queue_name: str, quiet: bool, logger: logging.Logger) -> dict[str, Any]:
    selected_sidebar_text = select_queue(panel, queue_name)
    sidebar_total = parse_sidebar_total(selected_sidebar_text)
    time.sleep(0.4)

    table = emis_window.find("id:workflowDatagridView", timeout=30)
    move_to_first_page(emis_window)
    table = emis_window.find("id:workflowDatagridView", timeout=15)

    pages_scanned = 0
    counts = Counter()
    seen_signatures = set()
    no_growth = 0

    while pages_scanned < 200:
        current_page, total_pages, _, next_btn = read_page_state(emis_window)
        visible_rows = scan_document_type_rows(table)

        before = len(seen_signatures)
        for row_idx, letter_type in visible_rows:
            signature = (current_page, row_idx, letter_type)
            if signature in seen_signatures:
                continue
            seen_signatures.add(signature)
            counts[letter_type] += 1

        pages_scanned += 1
        added = len(seen_signatures) - before
        if not quiet:
            logger.info(
                "[%s] page %s/%s: +%s unique rows (%s total)",
                queue_name,
                current_page,
                total_pages,
                added,
                len(seen_signatures),
            )

        if added == 0:
            no_growth += 1
        else:
            no_growth = 0

        if sidebar_total is not None and len(seen_signatures) >= sidebar_total:
            break
        if no_growth >= 3:
            break
        if current_page >= total_pages:
            break
        if not control_enabled(next_btn):
            break

        next_btn.click()
        if not wait_for_page_change(emis_window, previous_page=current_page):
            logger.info("[%s] Page did not advance, stopping scan.", queue_name)
            break
        table = emis_window.find("id:workflowDatagridView", timeout=15)

    rows_scanned = sum(counts.values())
    return {
        "queue": queue_name,
        "sidebar_text": selected_sidebar_text,
        "sidebar_total": sidebar_total,
        "rows_scanned": rows_scanned,
        "pages_scanned": pages_scanned,
        "sidebar_match": sidebar_total is None or sidebar_total == rows_scanned,
        "letter_type_counts": sort_counter(counts),
    }


def format_summary(report: dict[str, Any]) -> str:
    lines = []
    lines.append(f"Generated at: {report['generated_at']}")
    lines.append(f"Credentials source: {report['credentials_source']}")
    lines.append("")

    for queue_report in report["queues"]:
        lines.append(f"Queue: {queue_report['queue']}")
        if queue_report["sidebar_total"] is not None:
            lines.append(f"Sidebar total: {queue_report['sidebar_total']}")
        lines.append(f"Rows scanned: {queue_report['rows_scanned']}")
        lines.append(f"Pages scanned: {queue_report['pages_scanned']}")
        lines.append(f"Sidebar match: {queue_report['sidebar_match']}")
        for letter_type, count in queue_report["letter_type_counts"].items():
            lines.append(f"  {letter_type}: {count}")
        lines.append("")

    lines.append("Overall totals:")
    lines.append(f"Rows scanned: {report['overall']['rows_scanned']}")
    for letter_type, count in report["overall"]["letter_type_counts"].items():
        lines.append(f"  {letter_type}: {count}")

    return "\n".join(lines)


def write_json(report: dict[str, Any], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(report, indent=2), encoding="utf-8")


def write_csv(report: dict[str, Any], path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(["queue", "letter_type", "count"])
        for queue_report in report["queues"]:
            queue_name = queue_report["queue"]
            for letter_type, count in queue_report["letter_type_counts"].items():
                writer.writerow([queue_name, letter_type, count])


def default_output_path(base_dir: Path) -> Path:
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    return base_dir / "output" / f"emis_letter_type_summary_{timestamp}.json"


def main() -> int:
    base_dir = Path(__file__).resolve().parent
    load_dotenv(base_dir / ".env", override=False)
    args = parse_args()

    logger = setup_logging(base_dir / "logs")

    try:
        cdb, username, password, cred_source = resolve_login_details(args, logger)

        if not args.skip_login:
            logger.info("Preparing EMIS login...")
            if args.kill_existing:
                kill_existing_processes()
            if args.clear_cache:
                clear_emis_cache()
            start_emis_with_cdb(cdb)
            login_to_emis(username, password, timeout=args.timeout)

        emis_window = find_emis_window(timeout=args.timeout, cdb=cdb)
        ensure_workflow_manager_visible(emis_window)
        open_workflow_manager(emis_window)
        panel = open_document_management_panel(emis_window)

        target_queues = resolve_target_queues(panel, args.queues)
        logger.info("Target queues: %s", ", ".join(target_queues))

        queue_reports = []
        overall_counter = Counter()

        for queue_name in target_queues:
            queue_report = summarise_queue(emis_window, panel, queue_name, args.quiet, logger)
            queue_reports.append(queue_report)
            overall_counter.update(queue_report["letter_type_counts"])

        report = {
            "generated_at": datetime.now(timezone.utc).isoformat(),
            "credentials_source": cred_source,
            "queues": queue_reports,
            "overall": {
                "rows_scanned": sum(overall_counter.values()),
                "letter_type_counts": sort_counter(overall_counter),
            },
        }

        output_path = Path(args.output) if args.output else default_output_path(base_dir)
        write_json(report, output_path)
        if args.csv_output:
            write_csv(report, Path(args.csv_output))

        print("")
        print(format_summary(report))
        print("")
        print(f"JSON written to: {output_path}")
        if args.csv_output:
            print(f"CSV written to: {args.csv_output}")

        if args.close_on_exit:
            try:
                emis_window.close_window()
            except Exception:
                pass

        if args.strict_sidebar_match:
            mismatched = [q for q in queue_reports if not q["sidebar_match"]]
            if mismatched:
                logger.error("Sidebar mismatch detected in %s queue(s).", len(mismatched))
                return 2

        return 0

    except Exception as error:
        logger.exception("Run failed: %s", error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

