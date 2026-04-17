from __future__ import annotations

from pathlib import Path

from models import ExtractedContent, GenerationReport

from .chart_extractor import extract_rendered_chart, extract_static_images
from .html_discovery import discover_html_files
from .html_parser import parse_html_file
from .table_extractor import extract_tables


def extract_content_from_input(
    input_path: str | Path,
    chart_output_dir: str | Path,
    report: GenerationReport,
) -> list[ExtractedContent]:
    contents: list[ExtractedContent] = []
    html_files = discover_html_files(input_path)

    for path in html_files:
        page, soup, html = parse_html_file(path)
        tables = extract_tables(page, soup)
        contents.extend(tables)
        contents.extend(extract_static_images(page, soup))

        chart = extract_rendered_chart(page, html, Path(chart_output_dir), report)
        if chart:
            contents.append(chart)

        if not tables and not chart:
            report.skipped.append(f"No table/chart content found in {path.name}")

    return contents

