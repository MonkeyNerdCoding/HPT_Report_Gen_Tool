from __future__ import annotations

from datetime import datetime
from pathlib import Path
import ast
import re

from bs4 import BeautifulSoup

from models import GenerationReport, ImageContent, ReportPage
from utils.normalize import strip_chart_suffix


LOGO_NAMES = {"edb360_img.jpg", "edb360_favicon.ico"}


def extract_static_images(page: ReportPage, soup: BeautifulSoup) -> list[ImageContent]:
    contents: list[ImageContent] = []
    for index, img in enumerate(soup.find_all("img")):
        src = img.get("src", "").strip()
        if not src:
            continue
        image_path = (page.path.parent / src).resolve()
        if image_path.name.lower() in LOGO_NAMES:
            continue
        if not image_path.exists():
            continue
        contents.append(
            ImageContent(
                "image",
                page.path,
                image_path,
                title=page.title,
                section=page.section,
                logical_key=page.logical_key,
                keys=set(page.keys),
                index=index,
            )
        )
    return contents


def detect_chart_variant(path: Path, html: str) -> str:
    name = path.stem.lower()
    if "pie_chart" in name or "PieChart" in html:
        return "pie"
    if "bar_chart" in name or "BarChart" in html:
        return "bar"
    if "line_chart" in name or "LineChart" in html:
        return "line"
    if "google.visualization" in html:
        return "chart"
    return ""


def is_google_chart_page(html: str) -> bool:
    return "google.visualization" in html or "arrayToDataTable" in html


def extract_rendered_chart(
    page: ReportPage,
    html: str,
    chart_output_dir: Path,
    report: GenerationReport,
) -> ImageContent | None:
    if not is_google_chart_page(html):
        return None

    fallback = render_google_chart_with_matplotlib(page, html, chart_output_dir, report)
    if fallback:
        return fallback

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        report.skipped.append(
            f"Chart detected but Playwright is not installed, skipped render: {page.path.name}"
        )
        return None

    chart_output_dir.mkdir(parents=True, exist_ok=True)
    variant = detect_chart_variant(page.path, html)
    image_path = chart_output_dir / f"{page.path.stem}.png"

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page_browser = browser.new_page(viewport={"width": 1000, "height": 700})
            page_browser.goto(page.path.resolve().as_uri(), wait_until="networkidle", timeout=30000)
            locator = page_browser.locator(".google-chart").first
            if locator.count() == 0:
                locator = page_browser.locator("svg").first
            locator.screenshot(path=str(image_path))
            browser.close()
    except Exception as exc:
        report.warnings.append(f"Could not render chart {page.path.name}: {exc}")
        return None

    keys = set(page.keys)
    keys.add(strip_chart_suffix(page.logical_key))
    return ImageContent(
        "chart",
        page.path,
        image_path,
        title=page.title,
        section=page.section,
        logical_key=strip_chart_suffix(page.logical_key),
        keys=keys,
        variant=variant,
    )


def render_google_chart_with_matplotlib(
    page: ReportPage,
    html: str,
    chart_output_dir: Path,
    report: GenerationReport,
) -> ImageContent | None:
    variant = detect_chart_variant(page.path, html)
    if variant != "line":
        return None

    parsed = parse_array_to_data_table(html)
    if not parsed:
        report.skipped.append(f"Chart has no plottable data, skipped render: {page.path.name}")
        return None

    headers, rows = parsed
    if len(headers) < 2 or not rows:
        report.skipped.append(f"Chart has no plottable data, skipped render: {page.path.name}")
        return None

    chart_output_dir.mkdir(parents=True, exist_ok=True)
    image_path = chart_output_dir / f"{page.path.stem}.png"
    title = extract_option_text(html, "title") or page.title
    y_label = extract_axis_title(html, "vAxis")

    render_line_chart_png(headers, rows, title, y_label, image_path)

    keys = set(page.keys)
    keys.add(strip_chart_suffix(page.logical_key))
    return ImageContent(
        "chart",
        page.path,
        image_path,
        title=page.title,
        section=page.section,
        logical_key=strip_chart_suffix(page.logical_key),
        keys=keys,
        variant=variant,
    )


def parse_array_to_data_table(html: str) -> tuple[list[str], list[list[object]]] | None:
    match = re.search(r"arrayToDataTable\(\s*\[(.*?)\]\s*\)", html, re.DOTALL)
    if not match:
        return None

    headers: list[str] = []
    rows: list[list[object]] = []
    for raw_line in match.group(1).splitlines():
        line = raw_line.strip().rstrip(",")
        if not line:
            continue
        if line.startswith(","):
            line = line[1:].strip()
        if not line.startswith("[") or not line.endswith("]"):
            continue

        if "new Date" in line:
            row = parse_date_row(line)
            if row:
                rows.append(row)
            continue

        try:
            parsed = ast.literal_eval(line)
        except Exception:
            continue
        if parsed and all(isinstance(item, str) for item in parsed):
            headers = [str(item) for item in parsed]

    if not headers:
        return None
    return headers, rows


def render_line_chart_png(
    headers: list[str],
    rows: list[list[object]],
    title: str,
    y_label: str,
    image_path: Path,
) -> None:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    fig, ax = plt.subplots(figsize=(12, 6.75))

    x_values = [row[0] for row in rows]
    has_plottable_data = False

    for series_index, series_name in enumerate(headers[1:]):
        y_values: list[float | None] = []
        for row in rows:
            value = row[series_index + 1] if len(row) > series_index + 1 else None
            y_values.append(float(value) if isinstance(value, (int, float)) else None)

        # Filter out None values for plotting
        valid_pairs = [
            (x, y) for x, y in zip(x_values, y_values) if y is not None
        ]
        if not valid_pairs:
            continue

        xs, ys = zip(*valid_pairs)
        ax.plot(xs, ys, marker="o", markersize=3, linewidth=2, label=series_name)
        has_plottable_data = True

    if not has_plottable_data:
        ax.text(
            0.5, 0.5,
            "No plottable data",
            transform=ax.transAxes,
            ha="center",
            va="center",
            fontsize=14,
            color="gray",
        )
    else:
        ax.grid(True, linestyle="--", alpha=0.5)
        if len(headers) > 2:
            ax.legend(loc="best", fontsize=9)

    if title:
        ax.set_title(title, fontsize=13, pad=10)
    if y_label:
        ax.set_ylabel(y_label, fontsize=10)

    # Handle datetime x-axis formatting
    if x_values and isinstance(x_values[0], datetime):
        fig.autofmt_xdate()

    plt.savefig(image_path, bbox_inches="tight", dpi=150)
    plt.close(fig)


def parse_date_row(line: str) -> list[object] | None:
    date_match = re.search(r"new Date\(([^)]*)\)", line)
    if not date_match:
        return None

    parts = [int(part.strip()) for part in date_match.group(1).split(",")]
    while len(parts) < 6:
        parts.append(0)
    year, month, day, hour, minute, second = parts[:6]
    date_value = datetime(year, month + 1, day, hour, minute, second)

    tail = line[date_match.end() :].strip()
    tail = tail.lstrip(", ").rstrip("]")
    values: list[object] = [date_value]
    for value in tail.split(","):
        value = value.strip()
        if not value or value.lower() == "null":
            values.append(None)
            continue
        try:
            values.append(float(value))
        except ValueError:
            values.append(None)
    return values


def extract_option_text(html: str, option_name: str) -> str:
    match = re.search(rf"{re.escape(option_name)}\s*:\s*'([^']+)'", html)
    return match.group(1) if match else ""


def extract_axis_title(html: str, axis_name: str) -> str:
    match = re.search(rf"{re.escape(axis_name)}\s*:\s*\{{[^}}]*title\s*:\s*'([^']+)'", html, re.DOTALL)
    return match.group(1) if match else ""