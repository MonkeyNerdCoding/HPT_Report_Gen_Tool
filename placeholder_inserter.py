from __future__ import annotations

import shutil
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph

from models import MappingRule


@dataclass
class PlaceholderInsertReport:
    inserted: list[str] = field(default_factory=list)
    already_present: list[str] = field(default_factory=list)
    missing_anchors: list[str] = field(default_factory=list)
    backup_path: Path | None = None


def insert_mapping_placeholders(
    word_file: str | Path,
    rules: list[MappingRule],
    create_backup: bool = True,
) -> PlaceholderInsertReport:
    path = Path(word_file)
    doc = Document(str(path))
    report = PlaceholderInsertReport()

    if create_backup:
        backup = path.with_name(f"{path.stem}.placeholder_backup{path.suffix}")
        if not backup.exists():
            shutil.copy2(path, backup)
        report.backup_path = backup

    anchors = _build_anchor_index(doc)
    previous_anchor: Paragraph | None = None
    ordered_placeholders = [rule.placeholder for rule in rules]

    for placeholder in ordered_placeholders:
        existing = _find_placeholder(doc, placeholder)
        if existing is not None:
            report.already_present.append(placeholder)
            previous_anchor = existing
            continue

        anchor = _find_best_anchor(doc, anchors, placeholder, previous_anchor)
        if anchor is None:
            report.missing_anchors.append(placeholder)
            continue

        previous_anchor = _insert_after(anchor, placeholder)
        report.inserted.append(placeholder)

    if report.inserted:
        doc.save(str(path))

    return report


def _find_placeholder(doc: DocumentObject, placeholder: str) -> Paragraph | None:
    for paragraph in doc.paragraphs:
        if placeholder in paragraph.text:
            return paragraph
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if placeholder in paragraph.text:
                        return paragraph
    return None


def _insert_after(paragraph: Paragraph, text: str) -> Paragraph:
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    inserted = Paragraph(new_p, paragraph._parent)
    inserted.style = "Normal"
    inserted.add_run(text)
    return inserted


def _build_anchor_index(doc: DocumentObject) -> dict[str, Paragraph]:
    anchors: dict[str, Paragraph] = {}
    for paragraph in doc.paragraphs:
        text = _normalize(paragraph.text)
        if text and text not in anchors:
            anchors[text] = paragraph
    return anchors


def _find_best_anchor(
    doc: DocumentObject,
    anchors: dict[str, Paragraph],
    placeholder: str,
    previous_anchor: Paragraph | None,
) -> Paragraph | None:
    strategy = _PLACEHOLDER_STRATEGIES.get(placeholder, ())
    for candidate in strategy:
        if isinstance(candidate, tuple):
            anchor = _section_tail(doc, candidate[0], candidate[1])
            if anchor is not None:
                return anchor
            continue

        anchor = anchors.get(_normalize(candidate))
        if anchor is not None:
            return anchor

    if placeholder in _DATABASE_FALLBACK_PLACEHOLDERS:
        return previous_anchor or anchors.get(_normalize("QUẢN LÝ DATABASE"))

    if placeholder in _MEMORY_FALLBACK_PLACEHOLDERS:
        return (
            anchors.get(_normalize("CẤU HÌNH VÙNG NHỚ CHO CSDL"))
            or anchors.get(_normalize("HIỆU SUẤT BỘ NHỚ"))
            or previous_anchor
        )

    return previous_anchor


def _section_tail(
    doc: DocumentObject,
    heading_text: str,
    next_heading_texts: tuple[str, ...],
) -> Paragraph | None:
    heading = _normalize(heading_text)
    next_headings = {_normalize(text) for text in next_heading_texts}
    in_section = False
    last_nonempty: Paragraph | None = None

    for paragraph in doc.paragraphs:
        text = _normalize(paragraph.text)
        if text == heading:
            in_section = True
            last_nonempty = paragraph
            continue
        if in_section and text in next_headings:
            return last_nonempty
        if in_section and text:
            last_nonempty = paragraph

    return last_nonempty if in_section else None


def _normalize(value: str) -> str:
    text = unicodedata.normalize("NFC", value or "")
    text = " ".join(text.strip().upper().split())
    return text


_DATABASE_FALLBACK_PLACEHOLDERS = {
    "<invalid_obj>",
    "<db_job>",
    "<sche_job>",
    "<table_stats>",
    "<index_stats>",
}

_MEMORY_FALLBACK_PLACEHOLDERS = {
    "<SGA_chart>",
    "<PGA_chart>",
    "<MEMORY_chart>",
}

_PLACEHOLDER_STRATEGIES: dict[str, tuple[str | tuple[str, tuple[str, ...]], ...]] = {
    "<tbs_usage>": ("MỨC ĐỘ SỬ DỤNG TABLESPACES",),
    "<data_file>": ("DATA FILE",),
    "<no_index_table>": (
        "CÁC BẢNG KHÔNG CÓ INDEX",
        "CÁC TABLES KHÔNG CÓ INDEX",
        "CÁC TABLE KHÔNG CÓ INDEX",
    ),
    "<no_pk_table>": (
        "CÁC BẢNG KHÔNG CÓ KHÓA CHÍNH",
        "CÁC BẢNG KHÔNG CÓ KHOÁ CHÍNH",
        "CÁC TABLE KHÔNG CÓ KHÓA CHÍNH",
        "CÁC TABLE KHÔNG CÓ KHOÁ CHÍNH",
        "CÁC TABLES KHÔNG CÓ KHÓA CHÍNH",
        "CÁC TABLES KHÔNG CÓ KHOÁ CHÍNH",
    ),
    "<invalid_obj>": ("INVALID OBJECT",),
    "<db_job>": ("Jobs",),
    "<sche_job>": ("Scheduler jobs",),
    "<table_stats>": ("Tables with Stale Stats", "TABLE VÀ INDEX CÓ STATISTICS CŨ"),
    "<index_stats>": ("Indexes with Stale Stats", "TABLE VÀ INDEX CÓ STATISTICS CŨ"),
    "<buffer_chart>": ("BUFFER CACHE HIT",),
    "<library_chart>": ("LIBRARY CACHE HIT",),
    "<SGA_chart>": ("SGA STATISTICS", "SYSTEM GLOBAL AREA"),
    "<PGA_chart>": ("PGA",),
    "<MEMORY_chart>": ("MEMORY STATISTICS",),
    "<log_switch_chart>": (
        "TẦN SUẤT LOG SWITCH",
        ("ONLINE REDO LOG", ("CẤU HÌNH LƯU TRỮ",)),
    ),
}
