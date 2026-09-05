from __future__ import annotations

from pathlib import Path
import re

from bs4 import BeautifulSoup

from .html_parser import read_html


DEFAULT_REPORT_CREATOR = "Trần Đinh Nhất Đăng"
DEFAULT_REPORT_APPROVER = "Hồ Quốc Trí"
DEFAULT_REPORT_VERSION = "1.0"


def extract_edb360_metadata(input_root: str | Path) -> dict[str, str]:
    root = Path(input_root)
    metadata: dict[str, str] = {}
    _merge(metadata, _extract_index_metadata(root))
    _merge(metadata, _extract_system_under_observation(root))
    _merge(metadata, _extract_identification(root))
    return metadata


def build_edb360_text_mapping(metadata: dict[str, str] | None = None) -> dict[str, str]:
    values = {
        "creator": DEFAULT_REPORT_CREATOR,
        "approver": DEFAULT_REPORT_APPROVER,
        "version": DEFAULT_REPORT_VERSION,
        **(metadata or {}),
    }
    database_name = values.get("database_name") or values.get("db_name") or values.get("db_unique_name") or ""
    if database_name:
        values.setdefault("database_display_name", database_name)
        values.setdefault("system_name", database_name)
        values.setdefault("customer_name", database_name)

    placeholders = {
        "customer_name": "{{customer_name}}",
        "system_name": "{{system_name}}",
        "database_name": "{{database_name}}",
        "database_display_name": "{{database_display_name}}",
        "dbid": "{{dbid}}",
        "db_unique_name": "{{db_unique_name}}",
        "oracle_version": "{{oracle_version}}",
        "database_size": "{{database_size}}",
        "database_configuration": "{{database_configuration}}",
        "instance_summary": "{{instance_summary}}",
        "host_summary": "{{host_summary}}",
        "creator": "{{creator}}",
        "approver": "{{approver}}",
        "version": "{{version}}",
        "collection_date": "{{collection_date}}",
        "report_period": "{{report_period}}",
        "report_timestamp": "{{report_timestamp}}",
        "history_days": "{{history_days}}",
    }

    mapping: dict[str, str] = {}
    for key, placeholder in placeholders.items():
        value = str(values.get(key, "") or "").strip()
        if value:
            mapping[placeholder] = value
    return mapping


def _extract_index_metadata(root: Path) -> dict[str, str]:
    index_file = next(iter(sorted(root.rglob("*_index.html"))), None)
    if not index_file:
        return {}

    text = _page_text(index_file)
    metadata: dict[str, str] = {}
    period_match = re.search(r"between\s+([0-9T:\-]+)\s+and\s+([0-9T:\-]+)", text, re.IGNORECASE)
    if period_match:
        start, end = period_match.groups()
        metadata["report_period"] = f"{start} to {end}"
        metadata.setdefault("collection_date", _month_year_from_timestamp(end))
    days_match = re.search(r"\bDays:\s*([0-9]+)", text, re.IGNORECASE)
    if days_match:
        metadata["history_days"] = days_match.group(1)
    timestamp_match = re.search(r"\bTimestamp:\s*([0-9T:\-]+)", text, re.IGNORECASE)
    if timestamp_match:
        metadata["report_timestamp"] = timestamp_match.group(1)
        metadata.setdefault("collection_date", _month_year_from_timestamp(timestamp_match.group(1)))
    return metadata


def _extract_system_under_observation(root: Path) -> dict[str, str]:
    path = next(iter(sorted(root.rglob("*system_under_observation.html"))), None)
    if not path:
        return {}

    text = _page_text(path)
    metadata: dict[str, str] = {}
    _assign_regex(metadata, "database_name", text, r"Database name:\s*([^\s]+)")
    _assign_regex(metadata, "oracle_version", text, r"Oracle Database version:\s*([^\s]+)")
    _assign_regex(metadata, "database_size", text, r"Database size:\s*([0-9.,]+\s*[A-Za-z]+)")
    _assign_regex(metadata, "database_configuration", text, r"Database configuration:\s*(.+?)(?:\s+\d+\s+Instance|\s+Operating system:|$)")
    return metadata


def _extract_identification(root: Path) -> dict[str, str]:
    path = next(iter(sorted(root.rglob("*identification.html"))), None)
    if not path:
        return {}

    html = read_html(path)
    soup = BeautifulSoup(html, "html.parser")
    table = soup.find("table")
    if table is None:
        return {}

    rows = []
    for tr in table.find_all("tr"):
        cells = [cell.get_text(" ", strip=True) for cell in tr.find_all(["th", "td"])]
        if cells:
            rows.append(cells)
    if len(rows) < 2:
        return {}

    header = [item.strip().upper() for item in rows[0]]
    first = rows[1]
    row = {header[index]: first[index] for index in range(min(len(header), len(first)))}
    metadata = {
        "dbid": row.get("DBID", ""),
        "database_name": row.get("DBNAME", ""),
        "db_unique_name": row.get("DB_UNIQUE_NAME", ""),
    }
    instances = []
    hosts = []
    for data_row in rows[1:]:
        data = {header[index]: data_row[index] for index in range(min(len(header), len(data_row)))}
        instance = data.get("INSTANCE_NAME", "")
        host = data.get("HOST_NAME", "")
        if instance:
            instances.append(instance)
        if host:
            hosts.append(host)
    if instances:
        metadata["instance_summary"] = ", ".join(dict.fromkeys(instances))
    if hosts:
        metadata["host_summary"] = ", ".join(dict.fromkeys(hosts))
    return {key: value for key, value in metadata.items() if value}


def _page_text(path: Path) -> str:
    soup = BeautifulSoup(read_html(path), "html.parser")
    return re.sub(r"\s+", " ", soup.get_text(" ", strip=True))


def _assign_regex(metadata: dict[str, str], key: str, text: str, pattern: str) -> None:
    match = re.search(pattern, text, re.IGNORECASE)
    if match:
        metadata[key] = match.group(1).strip()


def _month_year_from_timestamp(value: str) -> str:
    try:
        from datetime import datetime

        return datetime.fromisoformat(value).strftime("%b-%Y")
    except ValueError:
        return value[:7]


def _merge(target: dict[str, str], source: dict[str, str]) -> None:
    for key, value in source.items():
        if value and not target.get(key):
            target[key] = value
