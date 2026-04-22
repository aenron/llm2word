from io import BytesIO
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt

from app.schemas import AgendaDocRequest, RenderStyle


def _set_run_font(run, font_name: str, size_pt: float, bold: bool = False) -> None:
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)


def _paragraph_space_after_pt(body_size_pt: float, line_spacing: float) -> Pt:
    return Pt(body_size_pt * line_spacing * 0.5)


def _set_paragraph_spacing(paragraph, line_spacing: float, body_size_pt: float) -> None:
    pformat = paragraph.paragraph_format
    pformat.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pformat.line_spacing = line_spacing
    pformat.space_before = Pt(0)
    pformat.space_after = _paragraph_space_after_pt(body_size_pt, line_spacing)


def _chars_to_pt(chars: float, body_size_pt: float) -> Pt:
    return Pt(chars * body_size_pt)


def _indent_for_level(level: int, style: RenderStyle) -> Pt:
    if level <= 1:
        return _chars_to_pt(style.indent_level1_chars, style.body_size_pt)
    if level == 2:
        return _chars_to_pt(style.indent_level2_chars, style.body_size_pt)
    return _chars_to_pt(style.indent_level3_chars, style.body_size_pt)


def _set_document_default_style(document: Document, payload: AgendaDocRequest) -> None:
    normal_style = document.styles["Normal"]
    normal_style.font.name = payload.style.body_font
    normal_style.font.size = Pt(payload.style.body_size_pt)
    normal_style._element.rPr.rFonts.set(
        qn("w:eastAsia"), payload.style.body_font)


def _add_title_paragraph(document: Document, payload: AgendaDocRequest) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_paragraph_spacing(
        paragraph,
        payload.style.line_spacing,
        payload.style.title_size_pt,
    )
    paragraph.paragraph_format.space_after = Pt(30)

    run = paragraph.add_run(payload.title)
    _set_run_font(
        run,
        font_name=payload.style.title_font,
        size_pt=payload.style.title_size_pt,
        bold=payload.style.title_bold,
    )


def _add_meta_paragraphs(document: Document, payload: AgendaDocRequest) -> None:
    for item in payload.meta:
        paragraph = document.add_paragraph()
        _set_paragraph_spacing(
            paragraph,
            payload.style.line_spacing,
            payload.style.body_size_pt,
        )

        label_run = paragraph.add_run(f"{item.label}：")
        _set_run_font(
            label_run,
            font_name=payload.style.body_font,
            size_pt=payload.style.body_size_pt,
            bold=payload.style.label_bold,
        )

        value_run = paragraph.add_run(item.value)
        _set_run_font(
            value_run,
            font_name=payload.style.body_font,
            size_pt=payload.style.body_size_pt,
        )


def _add_agenda_paragraphs(document: Document, payload: AgendaDocRequest) -> None:
    heading = document.add_paragraph()
    _set_paragraph_spacing(
        heading,
        payload.style.line_spacing,
        payload.style.body_size_pt,
    )
    heading_run = heading.add_run("议程：")
    _set_run_font(
        heading_run,
        font_name=payload.style.body_font,
        size_pt=payload.style.body_size_pt,
        bold=True,
    )

    for item in payload.agenda:
        paragraph = document.add_paragraph()
        _set_paragraph_spacing(
            paragraph,
            payload.style.line_spacing,
            payload.style.body_size_pt,
        )
        paragraph.paragraph_format.left_indent = _indent_for_level(
            item.level, payload.style)

        if item.leading_bold and item.text.startswith(item.leading_bold):
            bold_run = paragraph.add_run(item.leading_bold)
            _set_run_font(
                bold_run,
                font_name=payload.style.body_font,
                size_pt=payload.style.body_size_pt,
                bold=True,
            )

            remain = item.text[len(item.leading_bold):]
            normal_run = paragraph.add_run(remain)
            _set_run_font(
                normal_run,
                font_name=payload.style.body_font,
                size_pt=payload.style.body_size_pt,
            )
            continue

        content_run = paragraph.add_run(item.text)
        _set_run_font(
            content_run,
            font_name=payload.style.body_font,
            size_pt=payload.style.body_size_pt,
        )


def render_agenda_docx(payload: AgendaDocRequest, template_path: Path) -> bytes:
    _ = template_path
    document = Document()
    _set_document_default_style(document, payload)
    _add_title_paragraph(document, payload)
    _add_meta_paragraphs(document, payload)
    _add_agenda_paragraphs(document, payload)

    output = BytesIO()
    document.save(output)
    output.seek(0)
    return output.getvalue()
