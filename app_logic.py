from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from config import DEFAULT_MAPPING, load_mapping_rules
from extraction.extract_html import extract_content_from_input
from mapping.content_registry import ContentRegistry
from mapping.mapper import resolve_mappings
from models import GenerationReport
from rendering.word_renderer import render_report


LogCallback = Callable[[str], None]


def generate_report(
    html_root_folder: str,
    word_file: str,
    output_file_path: str,
    log_callback: LogCallback | None = None,
) -> str:
    """Generate a report using the GUI Save As workflow."""
    return generate_report_to_file(
        html_input=_validate_html_root_folder(html_root_folder),
        word_file=word_file,
        output_file=output_file_path,
        mapping_file=DEFAULT_MAPPING,
        chart_output_dir=Path(output_file_path).parent / "generated_charts",
        log_callback=log_callback,
    )


def generate_report_to_file(
    html_input: str | Path,
    word_file: str | Path,
    output_file: str | Path,
    mapping_file: str | Path = DEFAULT_MAPPING,
    chart_output_dir: str | Path | None = None,
    validate_only: bool = False,
    log_callback: LogCallback | None = None,
) -> str:
    """Generate a report to an explicit output file path."""
    input_path = _validate_html_input(html_input)
    template_path = _validate_word_file(word_file)
    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if output_path.suffix.lower() != ".docx":
        raise ValueError(f"Output file must be a .docx file: {output_path}")

    mapping_path = Path(mapping_file)
    if not mapping_path.is_file():
        raise FileNotFoundError(f"Mapping file does not exist: {mapping_path}")

    chart_dir = Path(chart_output_dir) if chart_output_dir else output_path.parent / "generated_charts"
    report = GenerationReport()

    log = _make_logger(log_callback)
    log("Validating inputs...")
    log(f"HTML input: {input_path}")
    log(f"Template: {template_path}")
    log(f"Mapping: {mapping_path}")
    log(f"Output: {output_path}")

    log("Loading mapping rules...")
    rules = load_mapping_rules(mapping_path)
    log(f"Loaded mapping rules: {len(rules)}")

    log("Extracting tables and charts from HTML...")
    contents = extract_content_from_input(input_path, chart_dir, report)
    log(f"Extracted content blocks: {len(contents)}")
    if not contents:
        raise ValueError(f"No table, chart, or image content was found under HTML source folder: {input_path}")

    log("Resolving mappings...")
    registry = ContentRegistry(contents)
    resolved = resolve_mappings(rules, registry, report)
    log(f"Resolved mappings: {len(resolved)} / {len(rules)}")

    if validate_only:
        log("Validation completed. No Word file was written.")
    else:
        log("Rendering Word report...")
        render_report(template_path, output_path, resolved, rules, report)
        log(f"Done. Output saved to: {output_path}")

    _log_summary(report, log)
    return str(output_path)


def _make_logger(log_callback: LogCallback | None) -> LogCallback:
    def log(message: str) -> None:
        print(message)
        if log_callback:
            log_callback(message)

    return log


def _validate_html_root_folder(path_value: str | Path) -> Path:
    path = Path(path_value)
    if not path.exists():
        raise FileNotFoundError(f"HTML root folder does not exist: {path}")
    if not path.is_dir():
        raise NotADirectoryError(f"HTML source must be a root folder: {path}")
    return path


def _validate_html_input(path_value: str | Path) -> Path:
    path = Path(path_value)
    if not path.exists():
        raise FileNotFoundError(f"HTML input does not exist: {path}")
    if path.is_file() and path.suffix.lower() not in {".html", ".htm"}:
        raise ValueError(f"HTML input file must be .html or .htm: {path}")
    if not path.is_file() and not path.is_dir():
        raise ValueError(f"HTML input must be a file or folder: {path}")
    return path


def _validate_word_file(path_value: str | Path) -> Path:
    path = Path(path_value)
    if not path.exists():
        raise FileNotFoundError(f"Word file does not exist: {path}")
    if not path.is_file():
        raise ValueError(f"Word file path is not a file: {path}")
    if path.suffix.lower() != ".docx":
        raise ValueError(f"Word file must be a .docx file: {path}")
    return path


def _log_summary(report: GenerationReport, log: LogCallback) -> None:
    log("")
    log("Generation summary")
    log(f"Inserted: {len(report.inserted)}")
    for item in report.inserted:
        log(f"  + {item}")

    log(f"Missing content: {len(report.missing_content)}")
    for item in report.missing_content:
        log(f"  ! {item}")

    log(f"Missing placeholders: {len(report.missing_placeholders)}")
    for item in report.missing_placeholders:
        log(f"  ! {item}")

    log(f"Ambiguous mappings: {len(report.ambiguous)}")
    for item in report.ambiguous:
        log(f"  ! {item}")

    log(f"Warnings: {len(report.warnings)}")
    for item in report.warnings:
        log(f"  ! {item}")

    if report.skipped:
        log(f"Skipped: {len(report.skipped)}")
        for item in report.skipped:
            log(f"  ! {item}")
