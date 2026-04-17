from __future__ import annotations

from pathlib import Path
import yaml

from models import MappingRule


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT = BASE_DIR / "badinh1_0925"
DEFAULT_TEMPLATE = BASE_DIR / "BADINH_VES081.docx"
DEFAULT_OUTPUT = BASE_DIR / "output.docx"
DEFAULT_MAPPING = BASE_DIR / "mapping" / "report_mapping.yaml"
DEFAULT_CHART_OUTPUT_DIR = BASE_DIR / "generated_charts"


def load_mapping_rules(mapping_path: str | Path) -> list[MappingRule]:
    path = Path(mapping_path)
    with path.open("r", encoding="utf-8") as handle:
        raw = yaml.safe_load(handle) or {}

    rules: list[MappingRule] = []
    for item in raw.get("placeholders", []):
        rules.append(
            MappingRule(
                placeholder=item["placeholder"],
                source_key=item.get("source_key", ""),
                content_type=item.get("content_type", "table"),
                source_file=item.get("source_file", ""),
                section=item.get("section", ""),
                table_index=item.get("table_index"),
                chart_variant=item.get("chart_variant", ""),
                required=item.get("required", False),
                on_missing=item.get("on_missing", "leave"),
                width_inches=item.get("width_inches"),
            )
        )
    return rules

