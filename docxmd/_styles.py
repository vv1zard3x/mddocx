"""GOST-like document styling for python-docx.

All formatting decisions live here. The conversion module imports
:func:`apply_gost_styles` and the helper functions to attach formatting
to individual runs/paragraphs when style inheritance is not enough.

References: GOST 7.32-2017 (technical reports), GOST R 2.105.
"""

from __future__ import annotations

from typing import Final

from docx.document import Document as _Doc
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt, RGBColor

TNR: Final = "Times New Roman"
MONO: Final = "Courier New"

# A4, GOST 7.32 margins (left 30mm for binding, right 15mm, top/bottom 20mm).
PAGE_WIDTH_MM: Final = 210
PAGE_HEIGHT_MM: Final = 297
MARGIN_LEFT_MM: Final = 30
MARGIN_RIGHT_MM: Final = 15
MARGIN_TOP_MM: Final = 20
MARGIN_BOTTOM_MM: Final = 20

# Usable width for images / tables.
MAX_CONTENT_WIDTH_CM: Final = (
    PAGE_WIDTH_MM - MARGIN_LEFT_MM - MARGIN_RIGHT_MM
) / 10  # 16.5 cm

BODY_SIZE_PT: Final = 12
H1_SIZE_PT: Final = 14
CODE_SIZE_PT: Final = 10
FIRST_LINE_INDENT_CM: Final = 1.25
LINE_SPACING: Final = 1.5  # полуторный, как требует ГОСТ для основного текста

CODE_SHADING_HEX: Final = "F2F2F2"
QUOTE_BORDER_HEX: Final = "B0B0B0"
LINK_COLOR_HEX: Final = "0563C1"


def apply_gost_styles(doc: _Doc) -> None:
    """Configure page geometry and rewrite built-in/added styles.

    Idempotent; safe to call once per document right after :func:`Document`.
    """

    for section in doc.sections:
        section.page_width = Mm(PAGE_WIDTH_MM)
        section.page_height = Mm(PAGE_HEIGHT_MM)
        section.left_margin = Mm(MARGIN_LEFT_MM)
        section.right_margin = Mm(MARGIN_RIGHT_MM)
        section.top_margin = Mm(MARGIN_TOP_MM)
        section.bottom_margin = Mm(MARGIN_BOTTOM_MM)

    styles = doc.styles

    normal = styles["Normal"]
    _set_font(normal, TNR, size_pt=BODY_SIZE_PT)
    pf = normal.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = LINE_SPACING
    pf.first_line_indent = Cm(FIRST_LINE_INDENT_CM)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    h1 = styles["Heading 1"]
    _set_font(h1, TNR, size_pt=H1_SIZE_PT, bold=True, color=RGBColor(0, 0, 0))
    pf = h1.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.line_spacing = LINE_SPACING
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(18)
    pf.space_after = Pt(12)
    pf.keep_with_next = True
    pf.page_break_before = False

    for level in range(2, 10):
        try:
            h = styles[f"Heading {level}"]
        except KeyError:
            continue
        _set_font(
            h,
            TNR,
            size_pt=BODY_SIZE_PT,
            bold=True,
            italic=False,
            color=RGBColor(0, 0, 0),
        )
        pf = h.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.line_spacing = LINE_SPACING
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(12)
        pf.space_after = Pt(6)
        pf.keep_with_next = True

    quote = _get_or_create_style(doc, "Quote", WD_STYLE_TYPE.PARAGRAPH)
    _set_font(quote, TNR, size_pt=BODY_SIZE_PT, italic=True)
    pf = quote.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = LINE_SPACING
    pf.first_line_indent = Cm(0)
    pf.left_indent = Cm(1.25)
    pf.right_indent = Cm(0)
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    _set_style_border(quote, left=True, color=QUOTE_BORDER_HEX, size=18, space=12)

    code = _get_or_create_style(doc, "Code Block", WD_STYLE_TYPE.PARAGRAPH)
    _set_font(code, MONO, size_pt=CODE_SIZE_PT)
    pf = code.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.line_spacing = 1.15
    pf.first_line_indent = Cm(0)
    pf.left_indent = Cm(0.5)
    pf.right_indent = Cm(0.5)
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    _set_style_shading(code, CODE_SHADING_HEX)

    # Figure caption: per GOST 7.32 — regular weight, regular style, plain text.
    # Centered is customary in technical reports; GOST does not forbid it.
    caption = _get_or_create_style(doc, "Caption", WD_STYLE_TYPE.PARAGRAPH)
    _set_font(
        caption,
        TNR,
        size_pt=BODY_SIZE_PT,
        bold=False,
        italic=False,
        color=RGBColor(0, 0, 0),
    )
    pf = caption.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.line_spacing = LINE_SPACING
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(12)
    pf.keep_with_next = False

    # Table caption: per GOST 7.32 § 6.6.3 — placed above the table, left
    # aligned, no indent, regular weight & style ("Таблица N — Название").
    table_caption = _get_or_create_style(doc, "Table Caption", WD_STYLE_TYPE.PARAGRAPH)
    _set_font(
        table_caption,
        TNR,
        size_pt=BODY_SIZE_PT,
        bold=False,
        italic=False,
        color=RGBColor(0, 0, 0),
    )
    pf = table_caption.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.line_spacing = LINE_SPACING
    pf.first_line_indent = Cm(0)
    pf.space_before = Pt(12)
    pf.space_after = Pt(3)
    pf.keep_with_next = True

    for sname in (
        "List Bullet",
        "List Bullet 2",
        "List Bullet 3",
        "List Number",
        "List Number 2",
        "List Number 3",
    ):
        try:
            s = styles[sname]
        except KeyError:
            continue
        _set_font(s, TNR, size_pt=BODY_SIZE_PT)
        pf = s.paragraph_format
        pf.line_spacing = LINE_SPACING
        pf.first_line_indent = Cm(0)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def _get_or_create_style(doc: _Doc, name: str, style_type):
    try:
        return doc.styles[name]
    except KeyError:
        return doc.styles.add_style(name, style_type)


def _set_font(
    style,
    name: str,
    *,
    size_pt: int | float | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    color: RGBColor | None = None,
) -> None:
    """Set font attributes on a style, including the East-Asian/complex variants.

    python-docx only sets ``w:rFonts/@w:ascii`` when ``font.name`` is assigned,
    so Cyrillic falls back to the Word default unless we also touch
    ``hAnsi``, ``cs`` and ``eastAsia``. Same trap exists for ``w:color``:
    theme colors override explicit RGB unless wiped.
    """

    style.font.name = name
    if size_pt is not None:
        style.font.size = Pt(size_pt)
    if bold is not None:
        style.font.bold = bold
    if italic is not None:
        style.font.italic = italic
    if color is not None:
        style.font.color.rgb = color

    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    # Theme-based font references take precedence over explicit ones in many
    # renderers; strip them so the new font name actually wins.
    for theme_attr in ("w:asciiTheme", "w:hAnsiTheme", "w:cstheme", "w:eastAsiaTheme"):
        key = qn(theme_attr)
        if key in rfonts.attrib:
            del rfonts.attrib[key]
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), name)

    if color is not None:
        color_el = rpr.find(qn("w:color"))
        if color_el is not None:
            for theme_attr in ("w:themeColor", "w:themeTint", "w:themeShade"):
                key = qn(theme_attr)
                if key in color_el.attrib:
                    del color_el.attrib[key]


def _set_style_shading(style, fill_hex: str) -> None:
    pPr = style.element.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    pPr.append(shd)


def _set_style_border(
    style,
    *,
    top: bool = False,
    left: bool = False,
    bottom: bool = False,
    right: bool = False,
    color: str = "888888",
    size: int = 12,
    space: int = 4,
) -> None:
    pPr = style.element.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    edges = {"top": top, "left": left, "bottom": bottom, "right": right}
    for edge, enabled in edges.items():
        if not enabled:
            continue
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), str(size))
        b.set(qn("w:space"), str(space))
        b.set(qn("w:color"), color)
        pBdr.append(b)


def set_run_font(run, name: str) -> None:
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    rfonts = rPr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rPr.insert(0, rfonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(attr), name)


def set_run_shading(run, fill_hex: str) -> None:
    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill_hex)
    rPr.append(shd)


def set_paragraph_horizontal_rule(paragraph) -> None:
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "auto")
    pBdr.append(bottom)
    pPr.append(pBdr)
