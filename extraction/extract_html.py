from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from models import ExtractedContent, GenerationReport
from models import MappingRule
from utils.normalize import content_key_aliases

from .chart_extractor import extract_rendered_chart, extract_static_images
from .html_discovery import discover_html_files
from .html_parser import parse_html_file
from .table_extractor import extract_tables

CancelCheck = Callable[[], bool]


def extract_content_from_input(
    input_path: str | Path,
    chart_output_dir: str | Path,
    report: GenerationReport,
    chart_rules: list[MappingRule] | None = None,
    cancel_check: CancelCheck | None = None,
) -> list[ExtractedContent]:
    contents: list[ExtractedContent] = []
    html_files = discover_html_files(input_path)

    for path in html_files:
        if cancel_check and cancel_check():
            break
        page, soup, html = parse_html_file(path)
        tables = extract_tables(page, soup)
        contents.extend(tables)
        contents.extend(extract_static_images(page, soup))

        chart = None
        if _should_extract_chart(page, html, chart_rules):
            chart = extract_rendered_chart(page, html, Path(chart_output_dir), report)
        if chart:
            contents.append(chart)

        if not tables and not chart:
            report.skipped.append(f"No table/chart content found in {path.name}")

    return contents


def _should_extract_chart(page, html: str, chart_rules: list[MappingRule] | None) -> bool:
    if chart_rules is None:
        return True
    if not chart_rules:
        return False

    page_name = page.path.name.lower()
    page_keys = {alias for key in page.keys for alias in content_key_aliases(key)}
    page_keys.update(content_key_aliases(page.logical_key))
    page_text = html[:4000]

    for rule in chart_rules:
        if rule.source_file and _matches_source_file(page_name, Path(rule.source_file).name.lower()):
            return True
        if rule.source_key and content_key_aliases(rule.source_key) & page_keys:
            if not rule.chart_variant or rule.chart_variant.lower() in page_name or rule.chart_variant in page_text:
                return True
    return False


def _matches_source_file(actual_name: str, expected_name: str) -> bool:
    actual = actual_name.lower()
    expected = expected_name.lower()
    if actual == expected:
        return True
    if actual.endswith(f"_{expected}"):
        return True

    expected_stem = Path(expected).stem
    actual_stem = Path(actual).stem
    return bool(expected_stem and actual_stem.endswith(f"_{expected_stem}"))

