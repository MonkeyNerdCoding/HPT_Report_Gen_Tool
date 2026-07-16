from __future__ import annotations

import html
import re
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

import yaml
from bs4 import BeautifulSoup


PLACEHOLDER_PATTERN = re.compile(r"\{\{[^{}]+\}\}|<[^<>\s]+>")
MAPPING_FIELDS = (
    "placeholder",
    "source_key",
    "content_type",
    "source_file",
    "section",
    "table_index",
    "chart_variant",
    "required",
    "on_missing",
    "width_inches",
    "table_header_vertical",
    "description",
    "template",
    "location",
    "purpose",
    "source_html",
    "report_section",
)


def load_placeholder_items(mapping_path: str | Path) -> list[dict[str, Any]]:
    raw = _load_mapping(mapping_path)
    return [_normalize_item(item) for item in raw.get("placeholders", [])]


def save_placeholder_items(mapping_path: str | Path, items: list[dict[str, Any]]) -> None:
    path = Path(mapping_path)
    normalized = [_ordered_item(_normalize_item(item)) for item in items]
    with path.open("w", encoding="utf-8") as handle:
        yaml.safe_dump(
            {"placeholders": normalized},
            handle,
            sort_keys=False,
            allow_unicode=True,
            default_flow_style=False,
        )


def add_placeholder_items(mapping_path: str | Path, new_items: list[dict[str, Any]]) -> list[dict[str, Any]]:
    existing = load_placeholder_items(mapping_path)
    known = {item["placeholder"] for item in existing}
    added: list[dict[str, Any]] = []
    for item in new_items:
        normalized = _normalize_item(item)
        if not normalized.get("placeholder") or normalized["placeholder"] in known:
            continue
        existing.append(normalized)
        known.add(normalized["placeholder"])
        added.append(normalized)
    if added:
        save_placeholder_items(mapping_path, existing)
    return added


def upsert_placeholder_item(mapping_path: str | Path, item: dict[str, Any], original_placeholder: str = "") -> dict[str, Any]:
    normalized = _normalize_item(item)
    placeholder = normalized.get("placeholder", "")
    if not placeholder:
        raise ValueError("placeholder is required")

    existing = load_placeholder_items(mapping_path)
    lookup = original_placeholder or placeholder
    matched_index = next((index for index, row in enumerate(existing) if row.get("placeholder") == lookup), None)
    if matched_index is not None:
        merged = dict(existing[matched_index])
        merged.update(normalized)
        normalized = _normalize_item(merged)
        placeholder = normalized["placeholder"]

    duplicate = any(
        row.get("placeholder") == placeholder and index != matched_index
        for index, row in enumerate(existing)
    )
    if duplicate:
        raise ValueError(f"placeholder already exists: {placeholder}")

    if matched_index is None:
        existing.append(normalized)
    else:
        existing[matched_index] = normalized
    save_placeholder_items(mapping_path, existing)
    return normalized


def delete_placeholder_item(mapping_path: str | Path, placeholder: str) -> bool:
    existing = load_placeholder_items(mapping_path)
    remaining = [item for item in existing if item.get("placeholder") != placeholder]
    if len(remaining) == len(existing):
        return False
    save_placeholder_items(mapping_path, remaining)
    return True


def scan_placeholder_file(file_path: str | Path, mapping_path: str | Path) -> dict[str, Any]:
    path = Path(file_path)
    yaml_items = load_placeholder_items(mapping_path)
    yaml_by_placeholder = {item["placeholder"]: item for item in yaml_items}
    found = _extract_placeholders(path)
    counts = Counter(found)
    unique_found = sorted(counts)

    found_rows = [
        {
            **_scan_row(placeholder, yaml_by_placeholder.get(placeholder)),
            "count": counts[placeholder],
            "status": "mapped" if placeholder in yaml_by_placeholder else "new",
        }
        for placeholder in unique_found
    ]
    missing_rows = [
        {**_scan_row(item["placeholder"], item), "count": 0, "status": "missing_in_file"}
        for item in yaml_items
        if item["placeholder"] not in counts
    ]
    new_rows = [row for row in found_rows if row["status"] == "new"]

    return {
        "success": True,
        "file_name": path.name,
        "file_type": path.suffix.lower().lstrip("."),
        "found": found_rows,
        "missing_in_file": missing_rows,
        "new_in_file": new_rows,
        "summary": {
            "found": len(found_rows),
            "missing_in_file": len(missing_rows),
            "new_in_file": len(new_rows),
            "total_yaml": len(yaml_items),
        },
    }


def extract_docx_placeholders(file_path: str | Path) -> list[str]:
    placeholders: list[str] = []
    with zipfile.ZipFile(file_path, "r") as archive:
        for name in archive.namelist():
            if not name.startswith("word/") or not name.endswith(".xml"):
                continue
            raw_xml = archive.read(name).decode("utf-8", errors="ignore")
            text = _xml_text_content(raw_xml)
            compact_text = re.sub(r"\s+", "", text)
            text_matches = [_normalize_placeholder(match) for match in PLACEHOLDER_PATTERN.findall(text)]
            compact_matches = [_normalize_placeholder(match) for match in PLACEHOLDER_PATTERN.findall(compact_text)]
            placeholders.extend(text_matches)
            known = set(text_matches)
            placeholders.extend(match for match in compact_matches if match not in known)
    return placeholders


def placeholder_to_key(placeholder: str) -> str:
    value = placeholder.strip()
    if value.startswith("{{") and value.endswith("}}"):
        return value[2:-2].strip()
    return value.strip("<>").strip()


def infer_placeholder_type(mapping_key: str) -> str:
    key = mapping_key.lower()
    if key.startswith("chart_") or "chart" in key:
        return "chart"
    if key.startswith("image_") or "image" in key:
        return "image"
    if key.startswith("text_"):
        return "text"
    return "table"


def build_placeholder_from_token(placeholder: str, file_name: str = "") -> dict[str, Any]:
    mapping_key = placeholder_to_key(placeholder)
    return _normalize_item(
        {
            "placeholder": placeholder,
            "source_key": mapping_key,
            "content_type": infer_placeholder_type(mapping_key),
            "description": mapping_key.replace("_", " ").title(),
            "template": file_name,
            "location": "",
            "purpose": "",
        }
    )


def _extract_placeholders(path: Path) -> list[str]:
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return extract_docx_placeholders(path)
    if suffix in {".html", ".htm"}:
        text = path.read_text(encoding="utf-8", errors="ignore")
        return [_normalize_placeholder(match) for match in PLACEHOLDER_PATTERN.findall(_html_text_content(text))]
    if suffix == ".pdf":
        return _extract_pdf_placeholders(path)
    raise ValueError("Only .docx, .pdf, .html and .htm files are supported.")


def _extract_pdf_placeholders(path: Path) -> list[str]:
    try:
        from pypdf import PdfReader  # type: ignore
    except ImportError as exc:
        raise ValueError("PDF scan requires pypdf. Install it with: pip install pypdf") from exc

    reader = PdfReader(str(path))
    text_parts: list[str] = []
    for page in reader.pages:
        text_parts.append(page.extract_text() or "")
    text = "\n".join(text_parts)
    compact_text = re.sub(r"\s+", "", text)
    matches = [_normalize_placeholder(match) for match in PLACEHOLDER_PATTERN.findall(text)]
    known = set(matches)
    matches.extend(
        _normalize_placeholder(match)
        for match in PLACEHOLDER_PATTERN.findall(compact_text)
        if _normalize_placeholder(match) not in known
    )
    return matches


def _scan_row(placeholder: str, item: dict[str, Any] | None) -> dict[str, Any]:
    item = item or build_placeholder_from_token(placeholder)
    return {
        "placeholder": placeholder,
        "source_key": item.get("source_key", placeholder_to_key(placeholder)),
        "content_type": item.get("content_type", infer_placeholder_type(placeholder_to_key(placeholder))),
        "source_file": item.get("source_file", ""),
        "section": item.get("section", ""),
        "source_html": item.get("source_html", ""),
        "report_section": item.get("report_section", ""),
        "description": item.get("description", ""),
        "template": item.get("template", ""),
        "location": item.get("location", ""),
        "purpose": item.get("purpose", ""),
    }


def _load_mapping(mapping_path: str | Path) -> dict[str, Any]:
    path = Path(mapping_path)
    if not path.exists():
        return {"placeholders": []}
    with path.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle) or {}
    if not isinstance(raw, dict):
        return {"placeholders": []}
    raw.setdefault("placeholders", [])
    return raw


def _normalize_item(item: dict[str, Any]) -> dict[str, Any]:
    placeholder = str(item.get("placeholder", "")).strip()
    source_key = str(item.get("source_key") or placeholder_to_key(placeholder)).strip()
    normalized = dict(item)
    normalized["placeholder"] = placeholder
    normalized["source_key"] = source_key
    normalized["content_type"] = str(item.get("content_type") or infer_placeholder_type(source_key)).strip()
    normalized.setdefault("source_file", "")
    normalized.setdefault("section", "")
    normalized.setdefault("source_html", "")
    normalized.setdefault("report_section", "")
    normalized.setdefault("description", source_key.replace("_", " ").title() if source_key else placeholder)
    normalized.setdefault("template", normalized.get("source_file", ""))
    normalized.setdefault("location", normalized.get("section", ""))
    normalized.setdefault("purpose", "")
    return {key: value for key, value in normalized.items() if value not in (None, "")}


def _ordered_item(item: dict[str, Any]) -> dict[str, Any]:
    ordered = {field: item[field] for field in MAPPING_FIELDS if field in item}
    for key, value in item.items():
        if key not in ordered:
            ordered[key] = value
    return ordered


def _normalize_placeholder(value: str) -> str:
    value = value.strip()
    if value.startswith("{{") and value.endswith("}}"):
        inner = re.sub(r"\s+", "", value[2:-2])
        return f"{{{{{inner}}}}}"
    return value


def _xml_text_content(raw_xml: str) -> str:
    without_tags = re.sub(r"<[^>]+>", "", raw_xml)
    return html.unescape(without_tags)


def _html_text_content(raw_html: str) -> str:
    soup = BeautifulSoup(raw_html, "html.parser")
    return soup.get_text(" ")
