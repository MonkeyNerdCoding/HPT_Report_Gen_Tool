from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal


ContentType = Literal["table", "chart", "image", "text"]


@dataclass
class ReportPage:
    path: Path
    title: str = ""
    heading: str = ""
    section: str = ""
    logical_key: str = ""
    keys: set[str] = field(default_factory=set)
    warnings: list[str] = field(default_factory=list)


@dataclass
class ExtractedContent:
    content_type: ContentType
    source_path: Path
    title: str = ""
    section: str = ""
    logical_key: str = ""
    keys: set[str] = field(default_factory=set)
    index: int = 0
    variant: str = ""


@dataclass
class TableContent(ExtractedContent):
    rows: list[list[str]] = field(default_factory=list)
    no_rows_selected: bool = False

    def __init__(
        self,
        source_path: Path,
        rows: list[list[str]],
        title: str = "",
        section: str = "",
        logical_key: str = "",
        keys: set[str] | None = None,
        index: int = 0,
        no_rows_selected: bool = False,
    ) -> None:
        super().__init__("table", source_path, title, section, logical_key, keys or set(), index)
        self.rows = rows
        self.no_rows_selected = no_rows_selected


@dataclass
class ImageContent(ExtractedContent):
    image_path: Path = Path()

    def __init__(
        self,
        content_type: ContentType,
        source_path: Path,
        image_path: Path,
        title: str = "",
        section: str = "",
        logical_key: str = "",
        keys: set[str] | None = None,
        index: int = 0,
        variant: str = "",
    ) -> None:
        super().__init__(content_type, source_path, title, section, logical_key, keys or set(), index, variant)
        self.image_path = image_path


@dataclass
class MappingRule:
    placeholder: str
    source_key: str = ""
    content_type: ContentType = "table"
    source_file: str = ""
    section: str = ""
    table_index: int | None = None
    chart_variant: str = ""
    required: bool = False
    on_missing: str = "leave"
    width_inches: float | None = None
    table_header_vertical: bool = False


@dataclass
class GenerationReport:
    inserted: list[str] = field(default_factory=list)
    missing_content: list[str] = field(default_factory=list)
    missing_placeholders: list[str] = field(default_factory=list)
    ambiguous: list[str] = field(default_factory=list)
    skipped: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)

    def print_summary(self) -> None:
        print("\nGeneration summary")
        print(f"  Inserted: {len(self.inserted)}")
        for item in self.inserted:
            print(f"    + {item}")

        print(f"  Missing content: {len(self.missing_content)}")
        for item in self.missing_content:
            print(f"    ! {item}")

        print(f"  Missing placeholders: {len(self.missing_placeholders)}")
        for item in self.missing_placeholders:
            print(f"    ! {item}")

        print(f"  Ambiguous mappings: {len(self.ambiguous)}")
        for item in self.ambiguous:
            print(f"    ! {item}")

        print(f"  Warnings: {len(self.warnings)}")
        for item in self.warnings:
            print(f"    ! {item}")
