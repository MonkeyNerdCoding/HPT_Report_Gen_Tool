from __future__ import annotations

import json
import os
import re
import unicodedata
import urllib.error
import urllib.request
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as DocumentObject
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P

from extraction.html_parser import parse_html_file
from extraction.table_extractor import extract_tables


MISSING_ASSESSMENT = "Không đủ dữ liệu để đánh giá"
NO_ACTION = "Không cần điều chỉnh thêm"

SYSTEM_SCHEMAS = {
    "SYS",
    "SYSTEM",
    "DVSYS",
    "GSMADMIN_INTERNAL",
    "GSMCATUSER",
    "XDB",
    "CTXSYS",
    "MDSYS",
    "ORDSYS",
    "WMSYS",
    "DBSNMP",
    "OUTLN",
    "AUDSYS",
    "ORACLE_OCM",
    "OJVMSYS",
    "OLAPSYS",
    "APPQOSSYS",
    "LBACSYS",
    "GGSYS",
}
COMPILE_CANDIDATE_TYPES = {
    "PACKAGE",
    "PACKAGE BODY",
    "PROCEDURE",
    "FUNCTION",
    "VIEW",
    "TRIGGER",
    "TYPE",
    "TYPE BODY",
    "MATERIALIZED VIEW",
}

DEFAULT_REVIEW_SECTIONS = [
    {"section": "4.1.1. Control file", "keywords": ["control_file", "control file"]},
    {"section": "4.1.2. Online redo log", "keywords": ["redo", "online redo", "log group", "cau hinh redo log"]},
    {"section": "4.1.3.2. Mức độ sử dụng tablespace", "keywords": ["tablespace_usage", "tablespace usage", "muc do su dung tablespace"]},
    {"section": "4.1.4. Các table không có index", "keywords": ["tables_without_indexes", "without index", "khong co index"]},
    {"section": "4.1.5. Các table không có khóa chính", "keywords": ["tables_without_primary", "primary key", "khong co khoa chinh"]},
    {"section": "5.3. Invalid Object", "keywords": ["invalid_objects", "invalid object"]},
    {"section": "5.4. Table và index có statistics cũ", "keywords": ["stale", "statistics", "verify_stats", "statistics cu"]},
    {"section": "6.1. Oracle Foreground Process", "keywords": ["foreground", "process", "aas_per_wait", "cpu busy"]},
    {"section": "6.2. Cấu hình vùng nhớ cho CSDL", "keywords": ["memory_configuration", "memory configuration", "cau hinh vung nho"]},
    {"section": "6.3. Memory Statistics", "keywords": ["memory_statistics", "memory statistics"]},
    {"section": "6.4. SGA Statistics", "keywords": ["sga_statistics", "sga statistics"]},
    {"section": "6.5. Buffer Cache Hit", "keywords": ["buffer_cache_hit", "buffer cache hit"]},
    {"section": "6.6. Library Cache Hit", "keywords": ["library_cache_hit", "library cache hit"]},
    {"section": "6.7. PGA", "keywords": ["pga", "pga_statistics"]},
    {"section": "7.1. Patching và backup", "keywords": ["patch", "backup", "rman", "registry_sql_patch"]},
]


@dataclass
class ReviewSectionData:
    section: str
    source_files: list[str]
    row_count: int
    headers: list[str]
    sample_rows: list[list[str]]
    notes: list[str]
    assessment_hint: str = ""
    recommendation_hint: str = ""
    details: dict[str, Any] | None = None


class AIReviewError(Exception):
    pass


def generate_ai_review(source_path: Path, provider: str = "gemini") -> dict[str, Any]:
    extracted = extract_review_data(source_path)
    try:
        rows = GeminiReviewProvider().generate(extracted)
        return {"rows": rows, "provider": provider, "used_ai": True, "debug": ""}
    except AIReviewError as exc:
        rows = build_rule_based_review(extracted)
        return {"rows": rows, "provider": "local_rules", "used_ai": False, "debug": str(exc)}


def extract_review_data(source_path: Path) -> list[ReviewSectionData]:
    if source_path.is_dir():
        return extract_from_directory(source_path)
    suffix = source_path.suffix.lower()
    if suffix in {".html", ".htm"}:
        return extract_from_html_files([source_path])
    if suffix == ".docx":
        return extract_from_docx(source_path)
    raise AIReviewError("AI Review source must be an extracted folder, .html/.htm, or .docx file.")


def extract_from_directory(source_root: Path) -> list[ReviewSectionData]:
    html_files = sorted([path for path in source_root.rglob("*") if path.suffix.lower() in {".html", ".htm"}])
    if html_files:
        return extract_from_html_files(html_files)
    docx_files = sorted([path for path in source_root.rglob("*.docx")])
    if docx_files:
        return extract_from_docx(docx_files[0])
    return empty_sections()


def extract_from_html_files(html_files: list[Path]) -> list[ReviewSectionData]:
    buckets = {item["section"]: [] for item in DEFAULT_REVIEW_SECTIONS}
    notes = {item["section"]: [] for item in DEFAULT_REVIEW_SECTIONS}

    for path in html_files:
        try:
            page, soup, _html = parse_html_file(path)
            tables = extract_tables(page, soup)
        except Exception as exc:
            continue
        searchable = " ".join([path.name, page.title, page.heading, page.logical_key, *page.keys]).lower()
        for config in DEFAULT_REVIEW_SECTIONS:
            if any(keyword.lower() in searchable for keyword in config["keywords"]):
                if not tables:
                    notes[config["section"]].append(page.title or path.name)
                for table in tables[:2]:
                    buckets[config["section"]].append((path, table.rows, table.no_rows_selected))

    result: list[ReviewSectionData] = []
    for config in DEFAULT_REVIEW_SECTIONS:
        section = config["section"]
        entries = buckets[section]
        rows = []
        detail_rows = []
        files = []
        no_rows = False
        for path, table_rows, no_rows_selected in entries[:4]:
            files.append(path.name)
            rows.extend(table_rows[:8])
            detail_rows.extend(table_rows)
            no_rows = no_rows or no_rows_selected
        headers = rows[0] if rows else []
        sample_rows = rows[1:6] if len(rows) > 1 else []
        result.append(
            ReviewSectionData(
                section=section,
                source_files=files[:4],
                row_count=max(0, len(rows) - 1) if rows else 0,
                headers=headers[:10],
                sample_rows=[row[:10] for row in sample_rows],
                notes=["0 rows selected"] if no_rows else notes[section][:3],
                details=build_section_details(section, detail_rows),
            )
        )
    return result


def extract_from_docx(docx_path: Path) -> list[ReviewSectionData]:
    document = Document(docx_path)
    result = {item.section: item for item in empty_sections()}
    current_context: list[str] = []
    pending_section = ""
    last_section = ""
    capture_assessment_for = ""
    capture_recommendation_for = ""

    for block in iter_docx_blocks(document):
        if isinstance(block, Paragraph):
            text = clean_spaces(block.text)
            if not text:
                continue
            lowered = normalize_text(text)
            if lowered.startswith("danh gia"):
                capture_assessment_for = last_section
                capture_recommendation_for = ""
                value = text.split(":", 1)[1].strip() if ":" in text else ""
                if value and last_section:
                    result[last_section].assessment_hint = append_sentence(result[last_section].assessment_hint, value)
                continue
            if lowered.startswith("khuyen nghi"):
                capture_recommendation_for = last_section
                capture_assessment_for = ""
                value = text.split(":", 1)[1].strip() if ":" in text else ""
                if value and last_section:
                    result[last_section].recommendation_hint = append_sentence(result[last_section].recommendation_hint, value)
                continue
            matched = match_section(text)
            if matched:
                pending_section = matched
                last_section = matched
                current_context.append(text)
                current_context = current_context[-5:]
                capture_assessment_for = ""
                capture_recommendation_for = ""
                continue
            if is_docx_heading_text(text):
                capture_assessment_for = ""
                capture_recommendation_for = ""
                current_context.append(text)
                current_context = current_context[-5:]
                continue
            if capture_assessment_for:
                result[capture_assessment_for].assessment_hint = append_sentence(result[capture_assessment_for].assessment_hint, text)
                continue
            if capture_recommendation_for:
                result[capture_recommendation_for].recommendation_hint = append_sentence(result[capture_recommendation_for].recommendation_hint, text)
                continue
            current_context.append(text)
            current_context = current_context[-5:]
            continue

        rows = table_to_rows(block)
        if not rows or is_review_summary_table(rows):
            continue
        header_section = match_section_by_headers(rows)
        if header_section == "__skip__":
            continue
        table_text = " ".join([" ".join(row) for row in rows[:3]])
        section = header_section or pending_section or match_section_from_context(current_context) or match_section(table_text)
        if not section:
            continue
        item = result[section]
        table_details = build_section_details(section, rows)
        if not item.headers:
            item.source_files = [docx_path.name]
            item.row_count = max(0, len(rows) - 1)
            item.headers = rows[0][:10]
            item.sample_rows = [row[:10] for row in rows[1:8]]
            item.details = table_details
        else:
            item.row_count += max(0, len(rows) - 1)
            item.sample_rows.extend([row[:10] for row in rows[1:4]])
            item.sample_rows = item.sample_rows[:8]
            item.details = merge_section_details(item.details or {}, table_details)
        last_section = section
        pending_section = ""
        capture_assessment_for = ""
        capture_recommendation_for = ""

    apply_paragraph_hints(document, result)
    return list(result.values())


def empty_sections() -> list[ReviewSectionData]:
    return [
        ReviewSectionData(
            section=config["section"],
            source_files=[],
            row_count=0,
            headers=[],
            sample_rows=[],
            notes=[],
        )
        for config in DEFAULT_REVIEW_SECTIONS
    ]


def iter_docx_blocks(document: DocumentObject):
    for child in document.element.body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def table_to_rows(table: Table) -> list[list[str]]:
    return [[clean_spaces(cell.text) for cell in row.cells] for row in table.rows if any(clean_spaces(cell.text) for cell in row.cells)]


def clean_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def normalize_text(value: str) -> str:
    text = unicodedata.normalize("NFD", value)
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    text = text.replace("đ", "d").replace("Đ", "D")
    text = re.sub(r"[^a-zA-Z0-9_ ]+", " ", text)
    return re.sub(r"\s+", " ", text).strip().lower()


def match_section(text: str) -> str:
    normalized = normalize_text(text)
    if "memory statistics" in normalized:
        return "6.3. Memory Statistics"
    if "sga statistics" in normalized:
        return "6.4. SGA Statistics"
    if "pga statistics" in normalized:
        return "6.7. PGA"
    for config in DEFAULT_REVIEW_SECTIONS:
        if any(normalize_text(keyword) in normalized for keyword in config["keywords"]):
            return config["section"]
    return ""


def match_section_from_context(context: list[str]) -> str:
    for text in reversed(context[-4:]):
        matched = match_section(text)
        if matched:
            return matched
    return ""


def match_section_by_headers(rows: list[list[str]]) -> str:
    headers = {normalize_text(cell) for cell in (rows[0] if rows else [])}
    header_line = " ".join(headers)
    if "file_name" in headers and "file_id" in headers and "bytes" in headers:
        return "__skip__"
    if "resource_name" in headers and "profile" in headers and "limit" in headers:
        return "__skip__"
    if "username" in headers and "account_status" in headers:
        return "__skip__"
    if "table_name" in headers and "privilege" in headers:
        return "__skip__"
    if "granted_role" in headers and "grantee" in headers:
        return "__skip__"
    if "database_mode" in header_line and "recovery" in header_line and "protection" in header_line:
        return "__skip__"
    if "dest_id" in header_line and "destination" in header_line and "target" in header_line:
        return "__skip__"
    if {"tablespace_name", "pct_used"} & headers:
        return "4.1.3.2. Mức độ sử dụng tablespace"
    if "group" in header_line and "members" in header_line:
        return "4.1.2. Online redo log"
    if "group" in header_line and "member" in header_line and "type" in header_line:
        return "4.1.2. Online redo log"
    if {"parameter", "value"}.issubset(headers) and any("control" in normalize_text(cell) for row in rows[:3] for cell in row):
        return "4.1.1. Control file"
    if {"owner", "object_name", "object_type", "status"}.issubset(headers):
        return "5.3. Invalid Object"
    if {"patch_uid", "action", "status"}.issubset(headers):
        return "7.1. Patching và backup"
    return ""


def is_review_summary_table(rows: list[list[str]]) -> bool:
    if not rows:
        return False
    header = [normalize_text(cell) for cell in rows[0][:3]]
    return header == ["muc", "danh gia", "khuyen nghi"]


def append_sentence(current: str, value: str) -> str:
    value = clean_spaces(value)
    if not value:
        return current
    if not current:
        return value
    return clean_spaces(f"{current} {value}")


def is_docx_heading_text(text: str) -> bool:
    normalized = normalize_text(text)
    if len(normalized) < 3:
        return False
    if normalized.startswith(("danh gia", "khuyen nghi")):
        return False
    letters = [char for char in text if char.isalpha()]
    if len(letters) < 3:
        return False
    uppercase_ratio = sum(1 for char in letters if char.upper() == char) / len(letters)
    return uppercase_ratio > 0.75 and len(text) <= 90


def apply_paragraph_hints(document: DocumentObject, result: dict[str, ReviewSectionData]) -> None:
    current_section = ""
    mode = ""
    for paragraph in document.paragraphs:
        text = clean_spaces(paragraph.text)
        if not text:
            continue
        lowered = normalize_text(text)
        if lowered.startswith("danh gia"):
            mode = "assessment"
            value = text.split(":", 1)[1].strip() if ":" in text else ""
            if value and current_section:
                result[current_section].assessment_hint = value
                apply_combined_cache_hint(value, result)
            continue
        if lowered.startswith("khuyen nghi"):
            mode = "recommendation"
            value = text.split(":", 1)[1].strip() if ":" in text else ""
            if value and current_section:
                result[current_section].recommendation_hint = value
            continue
        matched = match_section(text)
        if matched:
            current_section = matched
            mode = ""
            continue
        if is_docx_heading_text(text):
            mode = ""
            continue
        if mode == "assessment" and current_section:
            result[current_section].assessment_hint = append_sentence(result[current_section].assessment_hint, text)
            apply_combined_cache_hint(text, result)
        elif mode == "recommendation" and current_section:
            result[current_section].recommendation_hint = append_sentence(result[current_section].recommendation_hint, text)


def apply_combined_cache_hint(text: str, result: dict[str, ReviewSectionData]) -> None:
    normalized = normalize_text(text)
    if "buffer cache hit" in normalized and "library cache hit" in normalized:
        result["6.5. Buffer Cache Hit"].assessment_hint = text
        result["6.6. Library Cache Hit"].assessment_hint = text


def build_section_details(section: str, rows: list[list[str]]) -> dict[str, Any]:
    if not rows:
        return {}
    headers = [normalize_text(cell) for cell in rows[0]]
    data_rows = rows[1:]
    if "tablespace" in normalize_text(section):
        return build_tablespace_details(headers, data_rows)
    if "khong co index" in normalize_text(section) or "khong co khoa chinh" in normalize_text(section):
        return build_schema_scope_details(headers, data_rows)
    if section == "5.3. Invalid Object":
        return build_invalid_object_details(headers, data_rows)
    return {}


def build_tablespace_details(headers: list[str], rows: list[list[str]]) -> dict[str, Any]:
    name_index = find_header_index(headers, ["tablespace_name", "tablespace"])
    pct_index = find_header_index(headers, ["pct_used", "used_percent", "percent_used"])
    size_index = find_header_index(headers, ["size_gb"])
    used_index = find_header_index(headers, ["used_gb"])
    if name_index < 0 or pct_index < 0:
        return {}
    tablespaces = []
    for row in rows:
        if len(row) <= max(name_index, pct_index):
            continue
        name = clean_spaces(row[name_index])
        if not name:
            continue
        pct = to_number(row[pct_index])
        if pct is None:
            continue
        item = {"name": name, "pct_used": pct}
        if size_index >= 0 and len(row) > size_index:
            item["size_gb"] = to_number(row[size_index])
        if used_index >= 0 and len(row) > used_index:
            item["used_gb"] = to_number(row[used_index])
        tablespaces.append(item)
    over_80 = [item for item in tablespaces if item["name"].upper() != "TOTAL" and item["pct_used"] >= 80]
    return {
        "threshold_pct": 80,
        "tablespace_count": len([item for item in tablespaces if item["name"].upper() != "TOTAL"]),
        "over_threshold": over_80,
        "all_tablespaces": tablespaces,
        "instruction": "If using the 80% threshold, list every tablespace in over_threshold and do not mention tablespaces below the threshold.",
    }


def build_schema_scope_details(headers: list[str], rows: list[list[str]]) -> dict[str, Any]:
    owner_index = find_header_index(headers, ["owner", "schema"])
    if owner_index < 0:
        return {"scope_note": "OWNER/schema column not found; report count only as raw extracted rows, not application risk."}
    owners: dict[str, int] = {}
    for row in rows:
        if len(row) <= owner_index:
            continue
        owner = clean_spaces(row[owner_index]).upper()
        if owner:
            owners[owner] = owners.get(owner, 0) + 1
    system_counts = {owner: count for owner, count in owners.items() if owner in SYSTEM_SCHEMAS}
    app_counts = {owner: count for owner, count in owners.items() if owner not in SYSTEM_SCHEMAS}
    return {
        "raw_row_count": sum(owners.values()),
        "system_schema_row_count": sum(system_counts.values()),
        "application_schema_row_count": sum(app_counts.values()),
        "system_schema_counts": system_counts,
        "application_schema_counts": app_counts,
        "scope_note": (
            "System schemas are separated from application schemas. "
            "Do not present raw total as application risk when system_schema_row_count is non-zero."
        ),
    }


def build_invalid_object_details(headers: list[str], rows: list[list[str]]) -> dict[str, Any]:
    owner_index = find_header_index(headers, ["owner"])
    type_index = find_header_index(headers, ["object_type", "type"])
    if type_index < 0:
        return {}
    type_counts: dict[str, int] = {}
    owner_counts: dict[str, int] = {}
    for row in rows:
        if len(row) > type_index:
            object_type = clean_spaces(row[type_index]).upper()
            if object_type:
                type_counts[object_type] = type_counts.get(object_type, 0) + 1
        if owner_index >= 0 and len(row) > owner_index:
            owner = clean_spaces(row[owner_index]).upper()
            if owner:
                owner_counts[owner] = owner_counts.get(owner, 0) + 1
    synonym_count = type_counts.get("SYNONYM", 0)
    compile_candidate_count = sum(count for object_type, count in type_counts.items() if object_type in COMPILE_CANDIDATE_TYPES)
    return {
        "raw_row_count": sum(type_counts.values()),
        "object_type_counts": type_counts,
        "owner_counts": owner_counts,
        "synonym_count": synonym_count,
        "compile_candidate_count": compile_candidate_count,
        "instruction": (
            "If most invalid objects are SYNONYM, recommend reviewing references and root cause first. "
            "Recommend compile/recompile only for packages, procedures, functions, views, triggers, types, or materialized views when present."
        ),
    }


def merge_section_details(current: dict[str, Any], incoming: dict[str, Any]) -> dict[str, Any]:
    if not current:
        return incoming
    if not incoming:
        return current
    merged = dict(current)
    for key, value in incoming.items():
        if isinstance(value, int | float):
            merged[key] = merged.get(key, 0) + value
        elif isinstance(value, list):
            merged[key] = list(merged.get(key, [])) + value
        elif isinstance(value, dict):
            combined = dict(merged.get(key, {}))
            for sub_key, sub_value in value.items():
                if isinstance(sub_value, int | float):
                    combined[sub_key] = combined.get(sub_key, 0) + sub_value
                else:
                    combined[sub_key] = sub_value
            merged[key] = combined
        elif key not in merged:
            merged[key] = value
    return merged


class GeminiReviewProvider:
    endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent"

    def __init__(self) -> None:
        self.api_key = load_env_value("GEMINI_API_KEY")

    def generate(self, sections: list[ReviewSectionData]) -> list[dict[str, str]]:
        if not self.api_key:
            raise AIReviewError("GEMINI_API_KEY is not configured.")
        payload = self._build_payload(sections, strict=False)
        raw = self._call(payload)
        try:
            return validate_review_json(extract_json_array(raw))
        except AIReviewError:
            retry_payload = self._build_payload(sections, strict=True, previous_response=raw)
            retry_raw = self._call(retry_payload)
            return validate_review_json(extract_json_array(retry_raw))

    def _build_payload(
        self,
        sections: list[ReviewSectionData],
        strict: bool,
        previous_response: str = "",
    ) -> dict[str, Any]:
        prompt = build_prompt(sections, strict=strict, previous_response=previous_response)
        return {
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {
                "temperature": 0.2,
                "responseMimeType": "application/json",
            },
        }

    def _call(self, payload: dict[str, Any]) -> str:
        request = urllib.request.Request(
            self.endpoint,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={
                "Content-Type": "application/json",
                "X-goog-api-key": self.api_key,
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=45) as response:
                body = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise AIReviewError(f"Gemini API error: {exc.code} {detail[:400]}") from exc
        except Exception as exc:
            raise AIReviewError(f"Gemini API request failed: {exc}") from exc

        candidates = body.get("candidates") or []
        parts = ((candidates[0].get("content") or {}).get("parts") or []) if candidates else []
        text = "".join(str(part.get("text", "")) for part in parts)
        if not text.strip():
            raise AIReviewError("Gemini returned an empty response.")
        return text


def build_prompt(sections: list[ReviewSectionData], strict: bool = False, previous_response: str = "") -> str:
    extracted = [
        {
            "section": item.section,
            "source_files": item.source_files,
            "row_count": item.row_count,
            "headers": item.headers,
            "sample_rows": item.sample_rows,
            "notes": item.notes,
            "assessment_hint": item.assessment_hint,
            "recommendation_hint": item.recommendation_hint,
            "details": item.details or {},
        }
        for item in sections
    ]
    strict_text = ""
    if strict:
        strict_text = (
            "\nPrevious response was invalid. Return a raw JSON array only, with no markdown, no comments, "
            "and exactly these keys: section, assessment, recommendation.\n"
            f"Previous response: {previous_response[:1200]}\n"
        )
    return (
        "You are generating an Oracle Health Check review table.\n"
        "Return JSON only.\n"
        "Do not use markdown.\n"
        "Do not write long paragraphs.\n"
        "Each assessment must be 1-3 concise sentences.\n"
        "Each recommendation must be 1-2 concise sentences.\n"
        "Only use the provided extracted data.\n"
        "Do not invent numbers, versions, paths, or risks.\n"
        f'If data is missing, write "{MISSING_ASSESSMENT}".\n'
        f'If no action is needed, write "{NO_ACTION}".\n'
        "If row_count = 0, assess that no abnormal data was recorded or there is not enough data.\n"
        "If tablespace usage is below 80%, assess it is safe and recommend no adjustment.\n"
        "For Tablespace 4.1.3.2, use details.over_threshold as the source of truth. If threshold_pct is 80, list every tablespace in over_threshold and do not omit any item such as USERS or TRANS_TBS when present. Do not mention tablespaces that are below the threshold.\n"
        "For Tables without Index 4.1.4 and Tables without Primary Key 4.1.5, do not use hard-coded quantities. Use details.application_schema_row_count, details.system_schema_row_count, and schema counts. If system schemas are present and there is no explicit exclusion rule, state the evaluated scope clearly instead of presenting raw total as application risk.\n"
        "For Invalid Objects 5.3, use details.object_type_counts. If SYNONYM is the majority, recommend reviewing invalid references/root cause and only compile/recompile suitable object types such as Package, Procedure, Function, View, Trigger, Type, or Materialized View when present.\n"
        "Keep existing wording for other sections when it already matches the extracted data.\n"
        "If redo log group has only one member, assess it is not multiplexed and recommend adding members.\n"
        "If control files are on the same path or mount point, assess they are not distributed and recommend separating storage.\n"
        f"{strict_text}\n"
        "Extracted summarized data:\n"
        f"{json.dumps(extracted, ensure_ascii=False)}\n"
        "Required output schema:\n"
        '[{"section":"...","assessment":"...","recommendation":"..."}]'
    )


def extract_json_array(raw_text: str) -> Any:
    text = raw_text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text, flags=re.IGNORECASE).strip()
        text = re.sub(r"```$", "", text).strip()
    start = text.find("[")
    end = text.rfind("]")
    if start == -1 or end == -1 or end <= start:
        raise AIReviewError("AI response did not contain a JSON array.")
    try:
        return json.loads(text[start : end + 1])
    except json.JSONDecodeError as exc:
        raise AIReviewError(f"AI response JSON is invalid: {exc}") from exc


def validate_review_json(value: Any) -> list[dict[str, str]]:
    if not isinstance(value, list):
        raise AIReviewError("AI response must be a JSON array.")
    allowed_sections = {item["section"] for item in DEFAULT_REVIEW_SECTIONS}
    rows: list[dict[str, str]] = []
    for item in value:
        if not isinstance(item, dict):
            raise AIReviewError("Each AI review row must be an object.")
        section = str(item.get("section", "")).strip()
        if section not in allowed_sections:
            raise AIReviewError(f"Unexpected AI review section: {section}")
        rows.append(
            {
                "section": section,
                "assessment": concise_text(item.get("assessment", "")),
                "recommendation": concise_text(item.get("recommendation", "")),
            }
        )
    existing = {row["section"] for row in rows}
    for section in allowed_sections:
        if section not in existing:
            rows.append({"section": section, "assessment": MISSING_ASSESSMENT, "recommendation": ""})
    rows.sort(key=lambda row: [item["section"] for item in DEFAULT_REVIEW_SECTIONS].index(row["section"]))
    return rows


def concise_text(value: Any) -> str:
    text = re.sub(r"\s+", " ", str(value or "")).strip()
    return text[:700]


def build_rule_based_review(sections: list[ReviewSectionData]) -> list[dict[str, str]]:
    rows = []
    for item in sections:
        assessment, recommendation = assess_section(item)
        rows.append({"section": item.section, "assessment": assessment, "recommendation": recommendation})
    return rows


def assess_section(item: ReviewSectionData) -> tuple[str, str]:
    if not item.source_files and not item.headers and not item.sample_rows:
        if item.assessment_hint:
            return concise_text(item.assessment_hint), concise_text(item.recommendation_hint or NO_ACTION)
        return MISSING_ASSESSMENT, ""
    if item.section == "4.1.1. Control file":
        return assess_control_files(item)
    if item.section == "4.1.2. Online redo log":
        return assess_redo_log(item)
    if item.section == "4.1.3.2. Mức độ sử dụng tablespace":
        return assess_tablespace(item)
    if item.section == "4.1.4. Các table không có index":
        return assess_count_issue(item, "Có nhiều bảng chưa có index", "Phối hợp với team ứng dụng để xác định các bảng cần tạo index, ưu tiên bảng lớn hoặc truy vấn thường xuyên.")
    if item.section == "4.1.5. Các table không có khóa chính":
        return assess_count_issue(item, "Có nhiều bảng chưa có Primary Key", "Xem xét bổ sung Primary Key cho các bảng quan trọng để đảm bảo toàn vẹn dữ liệu và hỗ trợ tối ưu truy vấn.")
    if item.section == "5.3. Invalid Object":
        return assess_count_issue(item, "Hệ thống có object ở trạng thái INVALID", "Thực hiện recompile các object INVALID và kiểm tra lại lỗi phụ thuộc sau khi xử lý.")
    if item.section == "5.4. Table và index có statistics cũ":
        return assess_count_issue(item, "Có table hoặc index có statistics cũ", "Kiểm tra job gather statistics và cập nhật statistics cho các object quan trọng.")
    if item.section == "7.1. Patching và backup":
        return assess_patch_backup(item)
    if item.assessment_hint:
        return concise_text(item.assessment_hint), concise_text(item.recommendation_hint or NO_ACTION)
    if item.row_count == 0:
        return "Không ghi nhận dữ liệu bất thường hoặc không phát hiện vấn đề từ dữ liệu đã trích xuất.", NO_ACTION
    return "Dữ liệu đã được trích xuất, chưa ghi nhận dấu hiệu bất thường rõ ràng từ bảng nguồn.", "Tiếp tục theo dõi và đối chiếu với ngưỡng vận hành thực tế trước khi chốt báo cáo."


def assess_control_files(item: ReviewSectionData) -> tuple[str, str]:
    values = flatten_sample_values(item)
    paths = [value for value in values if "/" in value or "\\" in value]
    mounts = {storage_root(path) for path in paths if storage_root(path)}
    if len(paths) >= 2 and len(mounts) <= 1:
        return (
            f"Database có {len(paths)} control files nhưng các file đang nằm cùng vùng lưu trữ {next(iter(mounts), '')}.",
            "Nên phân bố control files trên nhiều phân vùng hoặc disk khác nhau để tăng khả năng chịu lỗi.",
        )
    if len(paths) >= 2:
        return f"Database có {len(paths)} control files và đã có dấu hiệu phân tán vị trí lưu trữ.", NO_ACTION
    return MISSING_ASSESSMENT, ""


def assess_redo_log(item: ReviewSectionData) -> tuple[str, str]:
    headers = [normalize_text(cell) for cell in item.headers]
    member_index = find_header_index(headers, ["members"])
    if member_index >= 0:
        members = [to_number(row[member_index]) for row in item.sample_rows if len(row) > member_index]
        risky = [value for value in members if value == 1]
        if risky:
            return (
                f"Có {len(risky)} redo log group chỉ có 1 member, redo log chưa được multiplex đầy đủ.",
                "Thêm ít nhất 1 member cho mỗi redo log group và đặt trên disk/vị trí lưu trữ khác.",
            )
        if members:
            return "Các redo log group đã có nhiều hơn 1 member theo dữ liệu trích xuất.", NO_ACTION
    return MISSING_ASSESSMENT, ""


def assess_tablespace(item: ReviewSectionData) -> tuple[str, str]:
    headers = [normalize_text(cell) for cell in item.headers]
    name_index = find_header_index(headers, ["tablespace_name", "tablespace"])
    pct_index = find_header_index(headers, ["pct_used", "used_percent", "percent_used"])
    if pct_index < 0:
        return MISSING_ASSESSMENT, ""
    usages = []
    for row in item.sample_rows:
        if len(row) <= pct_index:
            continue
        name = row[name_index] if name_index >= 0 and len(row) > name_index else "tablespace"
        pct = to_number(row[pct_index])
        if pct is not None:
            usages.append((name, pct))
    if not usages:
        return MISSING_ASSESSMENT, ""
    high = [(name, pct) for name, pct in usages if name.lower() != "total" and pct >= 90]
    warning = [(name, pct) for name, pct in usages if name.lower() != "total" and 80 <= pct < 90]
    total = next((pct for name, pct in usages if name.lower() == "total"), None)
    if high:
        names = ", ".join(f"{name} ({pct:g}%)" for name, pct in high[:4])
        return (
            f"Một số tablespace đang sử dụng cao: {names}. Tổng mức sử dụng là {total:g}%." if total is not None else f"Một số tablespace đang sử dụng cao: {names}.",
            "Theo dõi tăng trưởng và mở rộng datafile/tablespace hoặc dọn dẹp dữ liệu nếu xu hướng tiếp tục tăng.",
        )
    if warning:
        names = ", ".join(f"{name} ({pct:g}%)" for name, pct in warning[:4])
        return f"Một số tablespace ở ngưỡng cần theo dõi: {names}.", "Theo dõi tăng trưởng định kỳ và chuẩn bị phương án mở rộng khi vượt ngưỡng cảnh báo."
    return "Dung lượng các tablespace đang ở ngưỡng an toàn.", NO_ACTION


def assess_count_issue(item: ReviewSectionData, issue_text: str, recommendation: str) -> tuple[str, str]:
    if item.row_count <= 0:
        return "Không ghi nhận dữ liệu bất thường hoặc không phát hiện vấn đề từ dữ liệu đã trích xuất.", NO_ACTION
    return f"{issue_text}; dữ liệu trích xuất ghi nhận khoảng {item.row_count} dòng liên quan.", recommendation


def assess_patch_backup(item: ReviewSectionData) -> tuple[str, str]:
    text = normalize_text(" ".join(flatten_sample_values(item)))
    if "completed" in text and ("success" in text or "apply" in text):
        return "Dữ liệu patch và backup ghi nhận trạng thái thành công/hoàn tất trong các dòng đã trích xuất.", "Tiếp tục duy trì lịch backup định kỳ và cập nhật patch theo khuyến nghị từ Oracle."
    if item.row_count > 0:
        return "Đã ghi nhận dữ liệu patch hoặc backup, cần kiểm tra chi tiết trạng thái từng dòng.", "Xác nhận các job backup gần nhất hoàn tất và đối chiếu mức patch với khuyến nghị hiện hành."
    return MISSING_ASSESSMENT, ""


def flatten_sample_values(item: ReviewSectionData) -> list[str]:
    return [cell for row in item.sample_rows for cell in row if cell]


def storage_root(path: str) -> str:
    clean = path.replace("\\", "/")
    parts = [part for part in clean.split("/") if part]
    return f"/{parts[0]}" if parts else ""


def find_header_index(headers: list[str], names: list[str]) -> int:
    normalized_names = {normalize_text(name) for name in names}
    for index, header in enumerate(headers):
        if header in normalized_names:
            return index
    for index, header in enumerate(headers):
        if any(name in header for name in normalized_names):
            return index
    return -1


def to_number(value: str) -> float | None:
    cleaned = str(value or "").strip().replace(",", ".")
    match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    return float(match.group(0)) if match else None


def load_env_value(name: str) -> str:
    value = os.getenv(name, "").strip()
    if value:
        return value.strip('"').strip("'")
    env_path = Path(__file__).resolve().parent / ".env"
    if not env_path.is_file():
        return ""
    for line in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
        if not line.strip() or line.lstrip().startswith("#") or "=" not in line:
            continue
        key, raw_value = line.split("=", 1)
        if key.strip() == name:
            return raw_value.strip().strip('"').strip("'")
    return ""
