from __future__ import annotations

from pathlib import Path
import re

from bs4 import BeautifulSoup

from models import ReportPage
from utils.normalize import normalize_key, strip_chart_suffix


def read_html(path: Path) -> str:
    for encoding in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def parse_html_file(path: Path) -> tuple[ReportPage, BeautifulSoup, str]:
    html = read_html(path)
    soup = BeautifulSoup(html, "html.parser")

    title = soup.title.get_text(" ", strip=True) if soup.title else ""
    h1 = soup.find("h1")
    heading = h1.get_text(" ", strip=True) if h1 else ""

    section, logical_from_name = extract_section_and_logical_name(path.stem)
    logical_key = normalize_key(title) or normalize_key(logical_from_name) or normalize_key(path.stem)

    keys = {
        normalize_key(title),
        normalize_key(heading),
        normalize_key(path.stem),
        normalize_key(logical_from_name),
        strip_chart_suffix(normalize_key(logical_from_name)),
        strip_chart_suffix(logical_key),
    }
    if section:
        keys.add(section.lower())
        keys.add(section.lower().replace(".", "_"))
    keys.discard("")

    page = ReportPage(
        path=path,
        title=title,
        heading=heading,
        section=section,
        logical_key=logical_key,
        keys=keys,
    )
    return page, soup, html


def extract_section_and_logical_name(stem: str) -> tuple[str, str]:
    matches = list(re.finditer(r"(?:^|_)([0-9]+[a-z]?)_([0-9]+)_(.+)$", stem, re.IGNORECASE))
    if not matches:
        return "", stem

    match = matches[-1]
    section = f"{match.group(1)}.{match.group(2)}"
    logical_name = match.group(3)
    return section, logical_name

