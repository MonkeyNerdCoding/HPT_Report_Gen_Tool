from __future__ import annotations

import os
from collections.abc import Callable, Iterable
from datetime import datetime
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from docx.table import _Cell
from docx.text.paragraph import Paragraph


LogCallback = Callable[[str], None]
DEFAULT_SQL_MAX_RENDER_ROWS = int(os.getenv("SQLHEALCHECK_MAX_RENDER_ROWS", "200"))


SQL_TABLE_MAPPING = {
    "<volume_info>": {
        "sheet": "Volume Info",
        "columns": [0, 1, 2, 3, 4, 5],
        "transpose": True,
    },
    "<file_size>": {
        "sheet": "File Sizes and Space",
        "columns": [0, 1, 2, 3, 4, 5, 7],
        "max_rows": 50,
    },
    "<fileio>": {
        "sheet": "IO Stats By File",
        "max_rows": 50,
        "vertical_header": True,
        "vertical_body": True,
        "horizontal_columns": ["Database Name", "Logical Name", "type_desc", "Physical Name", "file_id"],
        "header_height": 2.0,
        "row_height": 1.8,
        "column_widths": [
            2.5,
            2.5,
            1.2,
            2.0,
            10.0,
            1.5,
            1.5,
            1.5,
            1.5,
            1.5,
            1.5,
            1.5,
            1.5,
            1.5,
            1.5,
            1.5,
            1.5,
            1.5,
            1.5,
            1.5,
        ],
    },
    "<conn_count>": {
        "sheet": "Connection Counts by IP Address",
        "max_rows": 50,
    },
    "<cpu_usage>": {
        "sheet": "CPU Usage by Database",
        "columns": [0, 1, 3],
        "max_rows": 50,
    },
    "<io_usage>": {
        "sheet": "IO Usage By Database",
        "columns": [0, 1, 3],
        "max_rows": 50,
    },
    "<buffer_usage>": {
        "sheet": "Total Buffer Usage by Database",
        "columns": [0, 1, 3],
        "max_rows": 50,
    },
    "<top_worker>": {
        "sheet": "Top Worker Time Queries",
        "columns": [0, 1, 2, 4],
        "max_rows": 50,
    },
    "<missing_index>": {
        "sheet": "Missing Indexes",
        "columns": [2, 5, 6, 7, 9],
        "max_rows": 50,
    },
    "<agent_job>": {
        "sheet": "SQL Server Agent Jobs",
        "columns": [0, 1, 2, 3, 4, 8, 9],
        "max_rows": 50,
    },
    "<recent_bk>": {
        "sheet": "Recent Full Backups",
        "columns": [2, 3, 4, 5, 11],
        "max_rows": 50,
    },
    "<collect_date>": {},
}


SQL_CHART_MAPPING = {
    "<cpu_usage_chart>": {
        "sheet": "CPU Usage by Database",
        "title": "Chart 1. CPU Usage by Database",
        "label_col": 1,
        "value_col": 3,
        "top_n": 10,
    },
    "<io_usage_chart>": {
        "sheet": "IO Usage By Database",
        "title": "Chart 2. IO Usage By Database",
        "label_col": 1,
        "value_col": 3,
        "top_n": 10,
    },
    "<buffer_usage_chart>": {
        "sheet": "Total Buffer Usage by Database",
        "title": "Chart 3. Total Buffer Usage by Database",
        "label_col": 1,
        "value_col": 3,
        "top_n": 10,
    },
}


SQL_SCALAR_MAPPING = {
    "<volume_inf_per>": {
        "sheet": "Volume Info",
        "value_column": "Space Free %",
        "selector_column": "Space Free %",
        "selector": "min",
        "format": "{:.2f}",
    },
    "<volume_inf_size>": {
        "sheet": "Volume Info",
        "value_column": "Available Size (GB)",
        "selector_column": "Space Free %",
        "selector": "min",
        "format": "{:.2f}",
    },
}


def render_excel_report(
    excel_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
    mapping_path: str | Path | None = None,
    log_callback: LogCallback | None = None,
    max_table_rows: int | None = DEFAULT_SQL_MAX_RENDER_ROWS,
    lightweight_tables: bool = True,
    slow_step_seconds: float = 10.0,
) -> str:
    """Render SQLHealcheck Excel data into the Word template using SQL-specific mappings."""
    del mapping_path, lightweight_tables, slow_step_seconds

    excel = Path(excel_path)
    template = Path(template_path)
    output = Path(output_path)
    log = _make_logger(log_callback)

    if not excel.is_file():
        raise FileNotFoundError(f"Excel input does not exist: {excel}")
    if not template.is_file():
        raise FileNotFoundError(f"Word template does not exist: {template}")
    if template.suffix.lower() != ".docx":
        raise ValueError(f"Word template must be a .docx file: {template}")
    if output.suffix.lower() != ".docx":
        raise ValueError(f"Output report must be a .docx file: {output}")

    output.parent.mkdir(parents=True, exist_ok=True)
    log(f"SQLHealcheck Excel input: {excel}")
    log(f"SQLHealcheck template: {template}")
    log("Rendering SQLHealcheck Word report with fixed SQL placeholder mapping...")

    generated = generate_sql_healthcheck_report(
        excel_file=excel,
        template_file=template,
        output_file=output,
        mapping=SQL_TABLE_MAPPING,
        chart_mapping=SQL_CHART_MAPPING,
        scalar_mapping=SQL_SCALAR_MAPPING,
        log_callback=log_callback,
        max_table_rows=max_table_rows,
    )
    log(f"Word report created: {generated}")
    return str(generated)


def generate_sql_healthcheck_report(
    excel_file: str | Path,
    template_file: str | Path,
    output_file: str | Path,
    mapping: dict,
    chart_mapping: dict | None = None,
    scalar_mapping: dict | None = None,
    log_callback: LogCallback | None = None,
    max_table_rows: int | None = DEFAULT_SQL_MAX_RENDER_ROWS,
) -> str:
    log = _make_logger(log_callback)
    excel_path = Path(excel_file)
    output_path = Path(output_file)
    xls = pd.ExcelFile(excel_path)
    doc = Document(template_file)
    temp_images: list[Path] = []

    for placeholder, config in mapping.items():
        try:
            if placeholder == "<collect_date>":
                current_date = datetime.now().strftime("%m.%Y")
                replaced = replace_placeholder_text(doc, placeholder, current_date)
                log(f"Replaced {placeholder} with {current_date}" if replaced else f"Placeholder not found: {placeholder}")
                continue

            if not config:
                continue

            sheet_name = config["sheet"]
            if sheet_name not in xls.sheet_names:
                log(f"Could not process {placeholder}: sheet not found '{sheet_name}'")
                continue

            dataframe = pd.read_excel(xls, sheet_name=sheet_name)
            dataframe = _prepare_dataframe(dataframe, config)

            max_rows = config.get("max_rows")
            if max_rows is None and max_table_rows is not None:
                max_rows = max_table_rows
            if max_rows and len(dataframe) > max_rows:
                dataframe = dataframe.head(max_rows)

            inserted = replace_placeholder_with_table(doc, placeholder, dataframe, config)
            if inserted:
                log(f"Replaced {placeholder} with sheet '{sheet_name}' (rows={len(dataframe)})")
            else:
                log(f"Placeholder not found: {placeholder}")

        except Exception as exc:
            log(f"Could not process {placeholder}: {exc}")

    if chart_mapping:
        for placeholder, config in chart_mapping.items():
            try:
                sheet_name = config["sheet"]
                if sheet_name not in xls.sheet_names:
                    log(f"Could not create chart for {placeholder}: sheet not found '{sheet_name}'")
                    continue

                dataframe = pd.read_excel(xls, sheet_name=sheet_name)
                temp_image = output_path.parent / f"temp_chart_{placeholder.strip('<>').replace('_', '')}.png"
                temp_images.append(temp_image)

                if create_pie_chart(
                    dataframe,
                    config.get("title", sheet_name),
                    temp_image,
                    config.get("label_col", 0),
                    config.get("value_col", 1),
                    config.get("top_n", 10),
                    log,
                ):
                    inserted = replace_placeholder_with_image(doc, placeholder, temp_image)
                    log(f"Inserted chart for {placeholder}" if inserted else f"Placeholder not found: {placeholder}")
            except Exception as exc:
                log(f"Could not create chart for {placeholder}: {exc}")

    if scalar_mapping:
        for placeholder, config in scalar_mapping.items():
            try:
                value = extract_scalar_value(xls, config)
                replaced = replace_placeholder_text(doc, placeholder, value)
                log(f"Replaced {placeholder} with {value}" if replaced else f"Placeholder not found: {placeholder}")
            except Exception as exc:
                log(f"Could not process {placeholder}: {exc}")

    try:
        doc.save(output_path)
        log(f"Report generated: {output_path}")
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = output_path.with_name(f"{output_path.stem}_{timestamp}{output_path.suffix}")
        doc.save(output_path)
        log(f"File is open. Saved as: {output_path}")
    finally:
        for temp_image in temp_images:
            try:
                temp_image.unlink(missing_ok=True)
            except Exception:
                pass

    return str(output_path)


def extract_scalar_value(xls: pd.ExcelFile, config: dict) -> str:
    sheet_name = config["sheet"]
    if sheet_name not in xls.sheet_names:
        raise ValueError(f"sheet not found '{sheet_name}'")

    dataframe = pd.read_excel(xls, sheet_name=sheet_name)
    if dataframe.empty:
        return ""

    value_column = config["value_column"]
    selector_column = config.get("selector_column")
    selected_row = dataframe.iloc[0]

    if selector_column and selector_column in dataframe.columns:
        numeric_selector = pd.to_numeric(dataframe[selector_column], errors="coerce")
        if numeric_selector.notna().any():
            if config.get("selector") == "max":
                selected_row = dataframe.loc[numeric_selector.idxmax()]
            else:
                selected_row = dataframe.loc[numeric_selector.idxmin()]

    value = selected_row[value_column] if value_column in dataframe.columns else ""
    numeric_value = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    if pd.notna(numeric_value) and config.get("format"):
        return config["format"].format(float(numeric_value))
    return "" if pd.isna(value) else str(value)


def _prepare_dataframe(dataframe: pd.DataFrame, config: dict) -> pd.DataFrame:
    dataframe = dataframe.where(pd.notna(dataframe), "")

    if config.get("transpose", False):
        if config.get("columns"):
            selected_cols = [dataframe.columns[i] for i in config["columns"] if i < len(dataframe.columns)]
            dataframe = dataframe[selected_cols]

        original_first_col = dataframe.columns[0]
        dataframe = dataframe.T.reset_index()

        if len(dataframe.columns) > 1:
            new_columns = [original_first_col] + [str(value) for value in dataframe.iloc[0, 1:].tolist()]
            dataframe.columns = new_columns
            dataframe = dataframe.iloc[1:].reset_index(drop=True)

        dataframe = dataframe.loc[:, ~dataframe.columns.astype(str).str.lower().str.contains("nan", na=False)]
        dataframe = dataframe.loc[:, dataframe.columns.astype(str).str.strip() != ""]
        return dataframe

    if config.get("columns"):
        selected = [dataframe.columns[i] for i in config["columns"] if i < len(dataframe.columns)]
        dataframe = dataframe[selected]

    return dataframe


def replace_placeholder_text(doc: DocumentObject, placeholder: str, replacement: str) -> bool:
    replaced = False
    for paragraph in iter_all_paragraphs(doc):
        if placeholder in paragraph.text:
            _replace_text_in_paragraph(paragraph, placeholder, replacement)
            replaced = True
    return replaced


def replace_placeholder_with_table(doc: DocumentObject, placeholder: str, dataframe: pd.DataFrame, config: dict) -> bool:
    for paragraph in iter_all_paragraphs(doc):
        if placeholder not in paragraph.text:
            continue

        table = doc.add_table(rows=1, cols=len(dataframe.columns))
        table.autofit = True
        set_table_borders(table)

        header_height = config.get("header_height", 1.8)
        set_row_height(table.rows[0], header_height)

        use_vertical_header = config.get("vertical_header", False)
        for column_index, column_name in enumerate(dataframe.columns):
            cell = table.rows[0].cells[column_index]
            cell.text = str(column_name)
            set_cell_bg(cell, "0066CC")
            format_cell(cell, bold=True, font_color=RGBColor(255, 255, 255))
            if use_vertical_header:
                set_cell_text_direction(cell, "tbRl")

        horizontal_columns = config.get("horizontal_columns", [])
        vertical_body = config.get("vertical_body", False)
        row_height = config.get("row_height", 1.8)

        for _, row in dataframe.iterrows():
            new_row = table.add_row()
            set_row_height(new_row, row_height)
            for column_index, value in enumerate(row):
                cell = new_row.cells[column_index]
                cell.text = str(value)
                format_cell(cell)

                if vertical_body and dataframe.columns[column_index] not in horizontal_columns:
                    set_cell_text_direction(cell, "tbRl")

        if config.get("column_widths"):
            for column_index, width_cm in enumerate(config["column_widths"]):
                if column_index < len(table.columns):
                    set_column_width(table.columns[column_index], width_cm)

        paragraph._element.getparent().replace(paragraph._element, table._element)
        return True
    return False


def replace_placeholder_with_image(doc: DocumentObject, placeholder: str, image_path: Path) -> bool:
    for paragraph in iter_all_paragraphs(doc):
        if placeholder not in paragraph.text:
            continue

        paragraph.text = paragraph.text.replace(placeholder, "")
        run = paragraph.add_run()
        run.add_picture(str(image_path), width=Inches(5.5))
        return True
    return False


def iter_all_paragraphs(doc: DocumentObject) -> Iterable[Paragraph]:
    yield from _iter_paragraphs(doc)
    for section in doc.sections:
        yield from _iter_paragraphs(section.header)
        yield from _iter_paragraphs(section.footer)


def _iter_paragraphs(parent) -> Iterable[Paragraph]:
    for paragraph in parent.paragraphs:
        yield paragraph
    for table in parent.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from _iter_paragraphs(cell)


def _replace_text_in_paragraph(paragraph: Paragraph, placeholder: str, replacement: str) -> None:
    for run in paragraph.runs:
        if placeholder in run.text:
            run.text = run.text.replace(placeholder, replacement)

    if placeholder in paragraph.text:
        paragraph.text = paragraph.text.replace(placeholder, replacement)


def create_pie_chart(
    dataframe: pd.DataFrame,
    title: str,
    output_image: Path,
    label_col_idx: int = 0,
    value_col_idx: int = 1,
    top_n: int = 10,
    log: LogCallback | None = None,
) -> bool:
    log = log or print
    try:
        df_chart = dataframe.head(top_n).copy()
        labels = df_chart.iloc[:, label_col_idx].astype(str).tolist()
        values = pd.to_numeric(df_chart.iloc[:, value_col_idx], errors="coerce").fillna(0).tolist()

        valid_data = [
            (label, value)
            for label, value in zip(labels, values, strict=False)
            if value > 0 and label.strip() != "" and label.lower() not in ["nan", "none"]
        ]
        if not valid_data:
            log(f"No valid data for chart: {title}")
            return False

        labels, values = zip(*valid_data, strict=False)
        fig, ax = plt.subplots(figsize=(10, 8), facecolor="white")
        colors = [
            "#5B9BD5",
            "#ED7D31",
            "#A5A5A5",
            "#FFC000",
            "#70AD47",
            "#4472C4",
            "#C55A11",
            "#7030A0",
            "#44546A",
            "#264478",
        ]
        wedges, _texts = ax.pie(
            values,
            labels=None,
            startangle=90,
            colors=colors[: len(values)],
            explode=[0.02] * len(values),
        )
        ax.set_title(title, fontsize=16, fontweight="bold", pad=30, color="#333333")

        num_cols = 3 if len(labels) > 6 else (2 if len(labels) > 3 else 1)
        legend = ax.legend(labels, loc="upper center", bbox_to_anchor=(0.5, -0.05), ncol=num_cols, frameon=False, fontsize=11)

        for index, _wedge in enumerate(wedges):
            if index < len(legend.get_patches()):
                legend.get_patches()[index].set_facecolor(colors[index % len(colors)])

        plt.axis("equal")
        plt.tight_layout()
        plt.savefig(output_image, dpi=150, bbox_inches="tight", facecolor="white")
        plt.close(fig)
        log(f"Created chart: {output_image}")
        return True
    except Exception as exc:
        log(f"Error creating chart: {exc}")
        plt.close()
        return False


def set_cell_bg(cell: _Cell, fill_color: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill_color)
    tc_pr.append(shading)


def set_cell_text_direction(cell: _Cell, direction: str = "lrTb") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    text_direction = OxmlElement("w:textDirection")
    text_direction.set(qn("w:val"), direction)
    tc_pr.append(text_direction)


def set_row_height(row, height_cm: float) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tr_height = OxmlElement("w:trHeight")
    tr_height.set(qn("w:val"), str(int(height_cm * 567)))
    tr_height.set(qn("w:hRule"), "exact")
    tr_pr.append(tr_height)


def format_cell(cell: _Cell, bold: bool = False, font_color: RGBColor | None = None) -> None:
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Cambria"
            run.font.size = Pt(12)
            run.bold = bold
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Cambria")
            if font_color:
                run.font.color.rgb = font_color


def set_table_borders(table) -> None:
    table_properties = table._element.tblPr
    borders = OxmlElement("w:tblBorders")
    for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")
        borders.append(border)
    table_properties.append(borders)


def set_column_width(column, width_cm: float) -> None:
    for cell in column.cells:
        cell.width = Inches(width_cm / 2.54)


def _make_logger(log_callback: LogCallback | None) -> LogCallback:
    def log(message: str) -> None:
        print(message)
        if log_callback:
            log_callback(message)

    return log
