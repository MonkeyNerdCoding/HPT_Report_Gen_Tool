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
    x_label = extract_axis_title(html, "hAxis")

    render_line_chart_png(headers, rows, title, y_label, image_path, x_label=x_label)

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
    x_label: str = "",
) -> None:
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import matplotlib.dates as mdates
    import matplotlib.ticker as mticker

    fig, ax = plt.subplots(figsize=(12, 6.75))

    x_values = [row[0] for row in rows]
    import itertools
    default_colors = ['#1976D2', '#E64A19', '#FB8C00', '#2E7D32', '#8E24AA', '#039BE5']
    color_cycle = itertools.cycle(default_colors)
    has_plottable_data = False
    all_y_values: list[float] = []

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
        all_y_values.extend(ys)
        plot_xs, plot_ys = smooth_line_for_display(list(xs), list(ys))
        # Use color index based on series position for consistent color mapping
        series_color = default_colors[series_index % len(default_colors)]
        ax.plot(plot_xs, plot_ys, marker=None, linewidth=1.8, label=series_name, color=series_color)
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
        if len(headers) > 1:
            ax.legend(loc="best", fontsize=9)

    if title:
        ax.set_title(title, fontsize=13, pad=10)
    if y_label:
        ax.set_ylabel(y_label, fontsize=12, fontstyle="italic" if _is_library_cache_chart(title, y_label) else "normal")
    if x_label:
        ax.set_xlabel(x_label, fontsize=11, fontstyle="italic" if _is_library_cache_chart(title, y_label) else "normal")

    if _is_library_cache_chart(title, y_label) and all_y_values:
        ax.yaxis.set_major_formatter(mticker.StrMethodFormatter("{x:,.0f}"))
        upper = max(all_y_values)
        if upper > 0:
            rounded_upper = _round_axis_upper(upper)
            ax.set_ylim(bottom=0, top=rounded_upper)
            ax.yaxis.set_major_locator(mticker.MaxNLocator(nbins=8))

    # Handle datetime x-axis formatting
    if x_values and isinstance(x_values[0], datetime):
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%B %d,\n%Y"))
        fig.autofmt_xdate()

    plt.savefig(image_path, bbox_inches="tight", dpi=150)
    plt.close(fig)


def _is_library_cache_chart(title: str, y_label: str) -> bool:
    text = f"{title} {y_label}".lower()
    return "library cache hit ratio" in text


def _round_axis_upper(value: float) -> float:
    if value <= 100:
        return 100
    magnitude = 10 ** max(len(str(int(value))) - 2, 0)
    return ((int(value + magnitude - 1) // magnitude) * magnitude)


def smooth_line_for_display(
    xs: list[object],
    ys: list[float],
    points_per_segment: int = 12,
) -> tuple[list[object], list[float]]:
    if len(xs) < 4 or len(xs) != len(ys):
        return xs, ys

    import matplotlib.dates as mdates

    x_is_datetime = isinstance(xs[0], datetime)
    numeric_xs = [mdates.date2num(x) if isinstance(x, datetime) else float(x) for x in xs]
    if any(numeric_xs[index] >= numeric_xs[index + 1] for index in range(len(numeric_xs) - 1)):
        return xs, ys

    smooth_xs: list[float] = []
    smooth_ys: list[float] = []
    for index in range(len(xs) - 1):
        x0 = numeric_xs[max(index - 1, 0)]
        x1 = numeric_xs[index]
        x2 = numeric_xs[index + 1]
        x3 = numeric_xs[min(index + 2, len(xs) - 1)]
        y0 = ys[max(index - 1, 0)]
        y1 = ys[index]
        y2 = ys[index + 1]
        y3 = ys[min(index + 2, len(ys) - 1)]

        steps = points_per_segment if index < len(xs) - 2 else points_per_segment + 1
        for step in range(steps):
            t = step / points_per_segment
            smooth_xs.append(_catmull_rom_value(x0, x1, x2, x3, t))
            smooth_ys.append(_catmull_rom_value(y0, y1, y2, y3, t))

    if x_is_datetime:
        return [mdates.num2date(value).replace(tzinfo=None) for value in smooth_xs], smooth_ys
    return smooth_xs, smooth_ys


def _catmull_rom_value(p0: float, p1: float, p2: float, p3: float, t: float) -> float:
    return 0.5 * (
        (2 * p1)
        + (-p0 + p2) * t
        + (2 * p0 - 5 * p1 + 4 * p2 - p3) * t * t
        + (-p0 + 3 * p1 - 3 * p2 + p3) * t * t * t
    )


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
