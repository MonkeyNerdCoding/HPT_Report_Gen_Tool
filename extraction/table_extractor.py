from __future__ import annotations

from bs4 import BeautifulSoup, Tag

from models import ReportPage, TableContent
from utils.normalize import strip_chart_suffix


def extract_tables(page: ReportPage, soup: BeautifulSoup) -> list[TableContent]:
    tables = [table for table in soup.find_all("table") if looks_like_data_table(table)]
    tables.sort(key=lambda table: len(table.find_all("tr")), reverse=True)

    no_rows_selected = "0 rows selected" in soup.get_text(" ", strip=True).lower()
    if not tables and no_rows_selected:
        return [
            TableContent(
                source_path=page.path,
                rows=[],
                title=page.title,
                section=page.section,
                logical_key=page.logical_key,
                keys=set(page.keys),
                no_rows_selected=True,
            )
        ]

    contents: list[TableContent] = []
    for index, table in enumerate(tables):
        rows = html_table_to_matrix(table)
        keys = set(page.keys)
        if index:
            keys.add(f"{page.logical_key}_{index}")
        keys.add(strip_chart_suffix(page.logical_key))
        contents.append(
            TableContent(
                source_path=page.path,
                rows=rows,
                title=page.title,
                section=page.section,
                logical_key=page.logical_key,
                keys=keys,
                index=index,
                no_rows_selected=no_rows_selected,
            )
        )
    return contents


def looks_like_data_table(table: Tag) -> bool:
    rows = table.find_all("tr")
    if not rows:
        return False
    if table.find(["th", "td"]) is None:
        return False
    return True


def html_table_to_matrix(table: Tag) -> list[list[str]]:
    rows: list[list[str]] = []
    pending_rowspans: dict[int, list[object]] = {}

    for tr in table.find_all("tr"):
        row: list[str] = []
        col_index = 0

        while col_index in pending_rowspans:
            text, remaining = pending_rowspans[col_index]
            row.append(str(text))
            remaining = int(remaining) - 1
            if remaining:
                pending_rowspans[col_index] = [text, remaining]
            else:
                del pending_rowspans[col_index]
            col_index += 1

        for cell in tr.find_all(["td", "th"]):
            while col_index in pending_rowspans:
                text, remaining = pending_rowspans[col_index]
                row.append(str(text))
                remaining = int(remaining) - 1
                if remaining:
                    pending_rowspans[col_index] = [text, remaining]
                else:
                    del pending_rowspans[col_index]
                col_index += 1

            text = cell.get_text(" ", strip=True)
            colspan = int(cell.get("colspan", 1) or 1)
            rowspan = int(cell.get("rowspan", 1) or 1)
            for offset in range(colspan):
                row.append(text if offset == 0 else "")
                if rowspan > 1:
                    pending_rowspans[col_index + offset] = [text if offset == 0 else "", rowspan - 1]
            col_index += colspan

        if row:
            rows.append(row)

    width = max((len(row) for row in rows), default=0)
    for row in rows:
        row.extend([""] * (width - len(row)))
    return rows
