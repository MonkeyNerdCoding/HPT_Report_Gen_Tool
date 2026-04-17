from bs4 import BeautifulSoup
import os

def extract_tables_from_html(folder_path):
    """
    Scan all .html files in folder_path, extract first <table> from each,
    and return a dict mapping placeholder name to table element.
    Example: "Tablespace Usage" -> ("<tbs_usage>", table_element)
    """
    table_map = {}

    for filename in os.listdir(folder_path):
        if not filename.lower().endswith(".html"):
            continue

        file_path = os.path.join(folder_path, filename)
        with open(file_path, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f, "html.parser")

        table = soup.find("table")
        if not table:
            continue

        # Get logical name from filename (remove prefix and underscores)
        # Example: "00010_edb360_..._tablespace_usage.html" → "Tablespace Usage"
        parts = filename.split("_")
        logical_name = " ".join(parts[-2].replace(".html", "").split("_")).title()

        # Placeholder derived from logical name
        placeholder = "<" + logical_name.lower().replace(" ", "_") + ">"
        table_map[placeholder] = table

    return table_map


def html_table_to_matrix(table):
    """
    Convert <table> HTML to a 2D list of cell texts.
    """
    rows = []
    for tr in table.find_all("tr"):
        cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
        rows.append(cells)
    return rows


# New implementation compatibility exports.
from extraction.extract_html import extract_content_from_input as _extract_content_from_input
from extraction.table_extractor import html_table_to_matrix as _html_table_to_matrix
from models import GenerationReport, TableContent


def extract_tables_from_html(folder_path):
    """
    Backward-compatible wrapper for the original API.

    New code should use extraction.extract_html.extract_content_from_input(),
    which returns structured TableContent/ImageContent objects.
    """
    report = GenerationReport()
    contents = _extract_content_from_input(folder_path, "generated_charts", report)
    table_map = {}
    for content in contents:
        if isinstance(content, TableContent):
            key = content.logical_key or content.title
            table_map[key] = content.rows
    return table_map


html_table_to_matrix = _html_table_to_matrix
