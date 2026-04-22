from io import BytesIO
from pathlib import Path

from docx.enum.text import WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Pt
from docxtpl import DocxTemplate, RichText

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


def _build_meta_subdoc(tpl: DocxTemplate, payload: AgendaDocRequest):
    subdoc = tpl.new_subdoc()

    for item in payload.meta:
        paragraph = subdoc.add_paragraph()
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

    return subdoc


def _build_agenda_subdoc(tpl: DocxTemplate, payload: AgendaDocRequest):
    subdoc = tpl.new_subdoc()

    heading = subdoc.add_paragraph()
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
        paragraph = subdoc.add_paragraph()
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

    return subdoc


def render_agenda_docx(payload: AgendaDocRequest, template_path: Path) -> bytes:
    tpl = DocxTemplate(template_path)

    title_rich = RichText()
    title_rich.add(
        payload.title,
        font=payload.style.title_font,
        size=payload.style.title_size_pt * 2,
        bold=payload.style.title_bold,
    )

    context = {
        "title": title_rich,
        "meta_subdoc": _build_meta_subdoc(tpl, payload),
        "agenda_subdoc": _build_agenda_subdoc(tpl, payload),
    }

    tpl.render(context)

    output = BytesIO()
    tpl.save(output)
    output.seek(0)
    return output.getvalue()
