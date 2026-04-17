from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Inches, Pt, RGBColor
from docx.table import _Cell
from docx.text.paragraph import Paragraph

from models import ExtractedContent, GenerationReport, ImageContent, MappingRule, TableContent


def render_report(
    template_path: str | Path,
    output_path: str | Path,
    resolved: dict[str, tuple[MappingRule, ExtractedContent]],
    rules: list[MappingRule],
    report: GenerationReport,
) -> None:
    doc = Document(template_path)

    # Thay thế {{collection_date}} nếu có trong template (bao gồm cả body, header, footer)
    # Mục đích: nhiều template đặt placeholder trong footer/header nên cần xử lý
    # cả các section.header/footer chứ không chỉ body.
    
    from datetime import datetime
    collection_date = datetime.now().strftime("%b-%Y")
    replaced_collection_date = False
    # Thay thế trong body
    for paragraph in iter_paragraphs(doc):
        if "{{collection_date}}" in paragraph.text:
            paragraph.text = paragraph.text.replace("{{collection_date}}", collection_date)
            replaced_collection_date = True
    # Thay thế trong header và footer
    for section in doc.sections:
        # Header
        # Header: duyệt các paragraph trong header của từng section
        for paragraph in section.header.paragraphs:
            if "{{collection_date}}" in paragraph.text:
                paragraph.text = paragraph.text.replace("{{collection_date}}", collection_date)
                replaced_collection_date = True
        # Footer
        # Footer: duyệt các paragraph trong footer của từng section
        for paragraph in section.footer.paragraphs:
            if "{{collection_date}}" in paragraph.text:
                paragraph.text = paragraph.text.replace("{{collection_date}}", collection_date)
                replaced_collection_date = True

    for placeholder, (rule, content) in resolved.items():
        inserted = replace_placeholder(doc, placeholder, rule, content)
        if inserted:
            report.inserted.append(f"{placeholder} <- {content.source_path.name}")
        else:
            report.missing_placeholders.append(f"{placeholder}: placeholder not found in template")

    for rule in rules:
        if rule.placeholder not in resolved:
            if not document_contains_placeholder(doc, rule.placeholder):
                report.missing_placeholders.append(f"{rule.placeholder}: configured but not found in template")

    doc.save(output_path)


def document_contains_placeholder(doc: DocumentObject, placeholder: str) -> bool:
    # Kiểm tra xem placeholder có tồn tại trong body hay trong các cell của table
    return any(placeholder in paragraph.text for paragraph in iter_paragraphs(doc))


def replace_placeholder(
    doc: DocumentObject,
    placeholder: str,
    rule: MappingRule,
    content: ExtractedContent,
) -> bool:
    # Duyệt tất cả paragraph (cả trong tables) để tìm placeholder
    for paragraph in iter_paragraphs(doc):
        if placeholder not in paragraph.text:
            continue

        # Xoá nội dung paragraph hiện tại trước khi chèn nội dung mới
        clear_paragraph(paragraph)
        if isinstance(content, TableContent):
            insert_table_after_paragraph(doc, paragraph, content)
        elif isinstance(content, ImageContent):
            insert_image_after_paragraph(doc, paragraph, content, rule.width_inches)
        else:
            paragraph.text = ""
        remove_paragraph(paragraph)
        return True
    return False


def iter_paragraphs(parent: DocumentObject | _Cell):
    # Trả về từng Paragraph trong một DocumentObject hoặc _Cell.
    # Lưu ý: nếu một cell chứa table lồng nhau thì hàm này cũng sẽ đệ quy.
    for paragraph in parent.paragraphs:
        yield paragraph
    for table in parent.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from iter_paragraphs(cell)


def clear_paragraph(paragraph: Paragraph) -> None:
    # Xoá nội dung text trong mọi run của paragraph (giữ nguyên style/run)
    for run in paragraph.runs:
        run.text = ""


def remove_paragraph(paragraph: Paragraph) -> None:
    element = paragraph._element
    parent = element.getparent()
    if parent is not None:
        parent.remove(element)
    # Lưu ý: remove_paragraph dùng thao tác trực tiếp lên XML tree để gỡ phần tử paragraph


def insert_table_after_paragraph(doc: DocumentObject, paragraph: Paragraph, content: TableContent) -> None:
    # Tạo table mới và chèn ngay sau paragraph chứa placeholder
    rows = content.rows
    if not rows and content.no_rows_selected:
        rows = [["No rows selected"]]
    if not rows:
        rows = [["No data found"]]

    row_count = len(rows)
    col_count = max(len(row) for row in rows)
    table = doc.add_table(rows=row_count, cols=col_count)
    table.style = "Table Grid"

    for row_index, row_data in enumerate(rows):
        for col_index in range(col_count):
            cell = table.cell(row_index, col_index)
            cell.text = row_data[col_index] if col_index < len(row_data) else ""
            # Định dạng cell: font, size, màu; và tô màu header nếu là hàng đầu
            format_cell(cell, is_header=row_index == 0 and row_count > 1)

    set_table_autofit(table)
    paragraph._p.addnext(table._tbl)


def insert_image_after_paragraph(
    doc: DocumentObject,
    paragraph: Paragraph,
    content: ImageContent,
    width_inches: float | None,
) -> None:
    # Chèn ảnh vào một paragraph mới rồi đặt paragraph đó ngay sau paragraph gốc
    image_paragraph = doc.add_paragraph()
    run = image_paragraph.add_run()
    kwargs = {}
    if width_inches:
        kwargs["width"] = Inches(width_inches)
    run.add_picture(str(content.image_path), **kwargs)
    paragraph._p.addnext(image_paragraph._p)


def format_cell(cell: _Cell, is_header: bool = False) -> None:
    # Đặt font và kích thước chuẩn cho nội dung trong cell
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.name = "Cambria"
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(0, 0, 0)
            run._element.rPr.rFonts.set(qn("w:eastAsia"), "Cambria")

    if is_header:
        # Nếu là header: tô nền và đổi chữ sang màu trắng, chữ in đậm
        shading = parse_xml(r'<w:shd {} w:fill="0066CC"/>'.format(nsdecls("w")))
        cell._tc.get_or_add_tcPr().append(shading)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)


def set_table_autofit(table) -> None:
    tbl_pr = table._tbl.tblPr
    tbl_layout = OxmlElement("w:tblLayout")
    tbl_layout.set(qn("w:type"), "autofit")
    tbl_pr.append(tbl_layout)
    # Ghi chú: thiết lập autofit giúp table co giãn theo nội dung/khổ trang

