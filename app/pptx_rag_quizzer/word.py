from utils import convert_image_to_png_or_jpg
import uuid
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Union
from docx import Document as docx_lib
from docx.text.paragraph import Paragraph
from docx.table import Table, _Cell
from docx.oxml.ns import qn
from models.models import (
    WordDocument,
    WordSection,
    WordText,
    WordImage,
    Type,
)
import traceback

NAMESPACES = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

W_PARAGRAPH = qn("w:p")
W_TABLE = qn("w:tbl")
W_TEXT = qn("w:t")
W_DRAWING = qn("w:drawing")


def parse_word_document(file_object, file_name: str) -> WordDocument:
    """
    Parses an in-memory Word (.docx) file to extract ordered text and image content.

    Args:
        file_object (io.BytesIO): Input DOCX stream.
        file_name (str): Name of the uploaded file.

    Returns:
        WordDocument: Structured representation of the document content.
    """
    document = docx_lib(file_object)

    sections: List[WordSection] = []
    current_items: List[Union[WordText, WordImage]] = []
    current_section_number = 1
    current_section_title: Optional[str] = None
    section_has_content = False
    order_number = 0

    def finalize_current_section():
        nonlocal current_items, section_has_content
        if not current_items:
            return
        sections.append(
            WordSection(
                id=str(uuid.uuid4()),
                section_number=current_section_number,
                title=current_section_title,
                items=list(current_items),
            )
        )
        current_items = []
        section_has_content = False

    def add_text_item(text: str):
        nonlocal order_number, section_has_content
        cleaned = " ".join(text.split())
        if not cleaned:
            return
        current_items.append(
            WordText(
                id=str(uuid.uuid4()),
                section_number=current_section_number,
                content=cleaned,
                type=Type.text,
                order_number=order_number,
            )
        )
        order_number += 1
        section_has_content = True

    def add_image_item(image_bytes: bytes, extension: str):
        nonlocal order_number, section_has_content
        current_items.append(
            WordImage(
                id=str(uuid.uuid4()),
                section_number=current_section_number,
                content="none",
                type=Type.image,
                order_number=order_number,
                image_bytes=image_bytes,
                extension=extension,
            )
        )
        order_number += 1
        section_has_content = True

    def extract_image_from_inline(inline) -> Optional[Tuple[bytes, str]]:
        rel_ids = inline.xpath(".//a:blip/@r:embed", namespaces=NAMESPACES)
        for rel_id in rel_ids:
            image_part = document.part.related_parts.get(rel_id)
            if not image_part:
                print(f"⚠️  Missing image part for rId={rel_id}")
                continue
            blob = image_part.blob
            ext = Path(image_part.partname).suffix.lstrip(".")
            converted_bytes, converted_ext = convert_image_to_png_or_jpg(blob, ext)
            if converted_bytes is None:
                print(
                    f"⚠️  Skipped unsupported image format ({ext or 'unknown'}) "
                    f"in section {current_section_number}"
                )
                continue
            return converted_bytes, converted_ext
        return None

    def process_paragraph(paragraph: Paragraph):
        nonlocal current_section_number, current_section_title, order_number, section_has_content

        heading_text = _safe_text(paragraph.text)
        if _is_heading(paragraph) and heading_text:
            if current_items:
                finalize_current_section()
                current_section_number += 1
            elif sections:
                current_section_number = sections[-1].section_number + 1
            else:
                current_section_number = 1
            current_section_title = heading_text
            order_number = 0
            section_has_content = False
            add_text_item(heading_text)
            return

        text_buffer: List[str] = []
        for child in paragraph._element:
            if child.tag == W_TEXT:
                text_value = _safe_text(child.text)
                if text_value:
                    text_buffer.append(text_value)
            elif child.tag == W_DRAWING:
                if text_buffer:
                    add_text_item(" ".join(text_buffer))
                    text_buffer = []
                inlines = child.xpath(".//wp:inline | .//wp:anchor", namespaces=NAMESPACES)
                for inline in inlines:
                    try:
                        image_payload = extract_image_from_inline(inline)
                        if image_payload:
                            add_image_item(*image_payload)
                    except Exception as image_error:
                        print(
                            "Skipping image due to error | "
                            f"section_number={current_section_number} | "
                            f"order_number={order_number} | error={image_error}"
                        )
                        print(traceback.format_exc())
                        continue

        if text_buffer:
            add_text_item(" ".join(text_buffer))

    def process_table(table: Table):
        for row in table.rows:
            for cell in row.cells:
                process_cell(cell)

    def process_cell(cell: _Cell):
        for child in cell._tc:
            if child.tag == W_PARAGRAPH:
                process_paragraph(Paragraph(child, cell))
            elif child.tag == W_TABLE:
                process_table(Table(child, cell))

    for element in document.element.body:
        try:
            if element.tag == W_PARAGRAPH:
                process_paragraph(Paragraph(element, document))
            elif element.tag == W_TABLE:
                process_table(Table(element, document))
        except Exception as parse_error:
            print(
                "Skipping block due to error | "
                f"section_number={current_section_number} | "
                f"order_number={order_number} | error={parse_error}"
            )
            print(traceback.format_exc())
            continue

    if current_items:
        finalize_current_section()

    if not sections:
        sections.append(
            WordSection(
                id=str(uuid.uuid4()),
                section_number=current_section_number,
                title=current_section_title,
                items=[],
            )
        )

    return WordDocument(
        id=str(uuid.uuid4()),
        name=file_name,
        sections=sections,
    )


def rebuild_word_document_with_accessible_features(word_document_model, word_file):
    """
    Applies alt text descriptions from the parsed model back into a DOCX file.
    Notes are not generated for Word documents.
    """
    document = docx_lib(word_file)

    image_lookup: Dict[Tuple[int, int], WordImage] = {}
    for section in word_document_model.sections:
        for item in section.items:
            if item.type == Type.image:
                image_lookup[(section.section_number, item.order_number)] = item  # type: ignore[arg-type]

    current_section_number = 1
    order_number = 0
    section_has_content = False

    def increment_for_text(text: str):
        nonlocal order_number, section_has_content
        cleaned = " ".join(text.split())
        if not cleaned:
            return
        order_number += 1
        section_has_content = True

    def apply_alt_text(inline):
        nonlocal order_number, section_has_content
        key = (current_section_number, order_number)
        matching_image = image_lookup.get(key)

        if matching_image and getattr(matching_image, "content", None):
            alt_text = matching_image.content.strip()
            if alt_text:
                try:
                    doc_pr = inline.xpath("./wp:docPr", namespaces=NAMESPACES)
                    if doc_pr:
                        doc_pr[0].set("descr", alt_text)
                    else:
                        print(
                            "⚠️  Unable to locate docPr node for image | "
                            f"section_number={current_section_number} | "
                            f"order_number={order_number}"
                        )
                except Exception as alt_error:
                    print(
                        "❌ Error setting alt text | "
                        f"section_number={current_section_number} | "
                        f"order_number={order_number} | "
                        f"error={alt_error}"
                    )
                    print(traceback.format_exc())
        elif matching_image:
            print(
                "⚠️  Skipping alt text update due to empty description | "
                f"section_number={current_section_number} | "
                f"order_number={order_number}"
            )
        else:
            print(
                "⚠️  No matching image found for alt text | "
                f"section_number={current_section_number} | "
                f"order_number={order_number}"
            )

        order_number += 1
        section_has_content = True

    def process_paragraph(paragraph: Paragraph):
        nonlocal current_section_number, order_number, section_has_content

        heading_text = _safe_text(paragraph.text)
        if _is_heading(paragraph) and heading_text:
            if section_has_content:
                current_section_number += 1
            order_number = 0
            section_has_content = False
            increment_for_text(heading_text)
            return

        text_buffer: List[str] = []
        for child in paragraph._element:
            if child.tag == W_TEXT:
                text_value = _safe_text(child.text)
                if text_value:
                    text_buffer.append(text_value)
            elif child.tag == W_DRAWING:
                if text_buffer:
                    increment_for_text(" ".join(text_buffer))
                    text_buffer = []
                inlines = child.xpath(".//wp:inline | .//wp:anchor", namespaces=NAMESPACES)
                for inline in inlines:
                    apply_alt_text(inline)

        if text_buffer:
            increment_for_text(" ".join(text_buffer))

    def process_table(table: Table):
        for row in table.rows:
            for cell in row.cells:
                process_cell(cell)

    def process_cell(cell: _Cell):
        for child in cell._tc:
            if child.tag == W_PARAGRAPH:
                process_paragraph(Paragraph(child, cell))
            elif child.tag == W_TABLE:
                process_table(Table(child, cell))

    for element in document.element.body:
        if element.tag == W_PARAGRAPH:
            process_paragraph(Paragraph(element, document))
        elif element.tag == W_TABLE:
            process_table(Table(element, document))

    return document


def _is_heading(paragraph: Paragraph) -> bool:
    try:
        style = paragraph.style
        if style and style.name:
            return style.name.lower().startswith("heading")
    except Exception:
        return False
    return False


def _safe_text(value: Optional[str]) -> str:
    return (value or "").strip()



