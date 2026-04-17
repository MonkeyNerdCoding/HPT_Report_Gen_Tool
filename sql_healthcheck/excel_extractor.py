from __future__ import annotations

from pathlib import Path

import pandas as pd

from models import TableContent
from utils.normalize import normalize_key


def extract_tables_from_excel(excel_path: str | Path) -> list[TableContent]:
    """Adapt every Excel sheet into the TableContent shape used by word_renderer."""
    path = Path(excel_path)
    workbook = pd.read_excel(path, sheet_name=None, dtype=object)
    contents: list[TableContent] = []

    for index, (sheet_name, dataframe) in enumerate(workbook.items()):
        dataframe = dataframe.where(pd.notna(dataframe), "")
        rows = [list(map(str, dataframe.columns.tolist()))]
        rows.extend([list(map(str, row)) for row in dataframe.to_numpy().tolist()])
        keys = {
            sheet_name,
            normalize_key(sheet_name),
            sheet_name.lower().replace(" ", "_"),
        }
        contents.append(
            TableContent(
                source_path=path,
                rows=rows,
                title=sheet_name,
                logical_key=normalize_key(sheet_name),
                keys=keys,
                index=index,
            )
        )

    return contents
