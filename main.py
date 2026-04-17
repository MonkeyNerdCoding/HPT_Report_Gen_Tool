from __future__ import annotations

import argparse

from app_logic import generate_report_to_file
from config import DEFAULT_CHART_OUTPUT_DIR, DEFAULT_INPUT, DEFAULT_MAPPING, DEFAULT_OUTPUT, DEFAULT_TEMPLATE


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate a Word report from Oracle Health Check / EDB360 HTML output."
    )
    parser.add_argument("--input", default=str(DEFAULT_INPUT), help="HTML file or folder of HTML files.")
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE), help="Word template .docx path.")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT), help="Output .docx path.")
    parser.add_argument("--mapping", default=str(DEFAULT_MAPPING), help="YAML mapping file path.")
    parser.add_argument(
        "--chart-output-dir",
        default=str(DEFAULT_CHART_OUTPUT_DIR),
        help="Folder used for rendered chart images.",
    )
    parser.add_argument(
        "--validate-only",
        action="store_true",
        help="Extract and resolve mappings without writing a Word document.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    generate_report_to_file(
        html_input=args.input,
        word_file=args.template,
        output_file=args.output,
        mapping_file=args.mapping,
        chart_output_dir=args.chart_output_dir,
        validate_only=args.validate_only,
    )


if __name__ == "__main__":
    main()
