"""Custom OOXML numbering definitions for hierarchical lists and headings.

python-docx doesn't expose a high-level API for ``numbering.xml`` so we
operate on the raw XML element of the numbering part. Three definitions
are installed per document:

* one ``abstractNum`` shared by all bullet lists (counters are irrelevant
  for bullets, so they can share a single ``numId``);
* one ``abstractNum`` whose 9 levels render as ``%1.``, ``%1.%2.``,
  ``%1.%2.%3.`` etc. Each *top-level* ordered list in the markdown gets
  its own ``numId`` referencing this abstract, with explicit
  ``<w:lvlOverride><w:startOverride w:val="1"/>`` for every level so
  Word and LibreOffice both restart counters between independent lists.
  Google Docs already restarts; Word/LibreOffice require this hint when
  the abstract's ``multiLevelType`` is ``hybridMultilevel``;
* one ``abstractNum`` whose 5 levels are bound to ``Heading 2``…
  ``Heading 6`` via ``<w:pStyle>``. Word renders ``1``, ``1.1``,
  ``1.1.1`` etc. in front of each section heading automatically; the
  ``multilevel`` type cascades counter resets correctly (a fresh
  ``Heading 2`` zeroes the ``Heading 3`` counter, and so on).

The same abstracts power the ``Tab`` / ``Shift+Tab`` demotion behavior in
Word — both gestures rewrite ``w:ilvl`` against the same ``numId``.
"""

from __future__ import annotations

from dataclasses import dataclass

from docx.document import Document as _Doc
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

BULLET_ABSTRACT_ID = 1900  # high IDs to avoid collisions with the template
ORDERED_ABSTRACT_ID = 1901
HEADING_ABSTRACT_ID = 1902
BULLET_NUM_ID = 1900
HEADING_NUM_ID = 1901

# Bullet glyphs by depth. Level 0 follows GOST 7.32 §6.5.2 ("дефис"); deeper
# levels keep typographic bullets — GOST 2.105 would mandate lowercase letters
# then arabic numerals there, but that turns nested "bullet" lists into hybrid
# bullet/ordered structures, which surprises markdown authors more than it
# helps. Stick to a single-shape contract until someone needs strict 2.105.
_BULLET_GLYPHS = ["-", "\u25e6", "\u25aa", "\u2022", "\u25e6", "\u25aa", "\u2022", "\u25e6", "\u25aa"]

# How many heading levels carry automatic numbering. Bound to Heading 2..6 so
# H1 stays untouched — H1 is reserved for structural elements per GOST 7.32
# (Реферат, Содержание, Введение, Заключение, Список источников, Приложение).
_HEADING_DEPTH = 5


@dataclass
class ListContext:
    """One frame on the active-list stack maintained by the builder."""

    kind: str  # "bullet" | "ordered"
    num_id: int
    ilvl: int


def install_numbering(doc: _Doc) -> None:
    """Inject our abstractNum/num definitions into the document's numbering part.

    Must be called exactly once per document, before any list paragraphs are
    appended. Idempotency is enforced via :data:`BULLET_ABSTRACT_ID`.
    """

    numbering = doc.part.numbering_part.element

    for existing in numbering.findall(qn("w:abstractNum")):
        if existing.get(qn("w:abstractNumId")) == str(BULLET_ABSTRACT_ID):
            return  # already installed

    _insert_abstract_num(numbering, _build_bullet_abstract())
    _insert_abstract_num(numbering, _build_ordered_abstract())
    _insert_abstract_num(numbering, _build_heading_abstract())
    numbering.append(_build_num(BULLET_NUM_ID, BULLET_ABSTRACT_ID))
    numbering.append(_build_num(HEADING_NUM_ID, HEADING_ABSTRACT_ID))


def allocate_ordered_num_id(doc: _Doc) -> int:
    """Create a fresh ``<w:num>`` pointing at the hierarchical ordered abstract.

    Each top-level ordered list calls this once so its counter starts at 1
    independently of any earlier list. The returned ``<w:num>`` carries
    explicit ``<w:lvlOverride><w:startOverride w:val="1"/>`` for every
    level, which is what Word and LibreOffice actually consult when
    deciding whether to restart numbering across list instances.
    """

    numbering = doc.part.numbering_part.element
    used: set[int] = set()
    for num in numbering.findall(qn("w:num")):
        try:
            used.add(int(num.get(qn("w:numId"))))
        except (TypeError, ValueError):
            continue
    candidate = max(used | {2000}) + 1  # start the dynamic range well above presets
    numbering.append(
        _build_num(candidate, ORDERED_ABSTRACT_ID, restart_levels=range(9))
    )
    return candidate


def apply_numbering(paragraph, num_id: int, ilvl: int) -> None:
    """Attach a ``<w:numPr>`` to *paragraph* so Word renders it as a list item."""

    pPr = paragraph._p.get_or_add_pPr()
    existing = pPr.find(qn("w:numPr"))
    if existing is not None:
        pPr.remove(existing)
    num_pr = OxmlElement("w:numPr")
    ilvl_el = OxmlElement("w:ilvl")
    ilvl_el.set(qn("w:val"), str(ilvl))
    num_pr.append(ilvl_el)
    num_id_el = OxmlElement("w:numId")
    num_id_el.set(qn("w:val"), str(num_id))
    num_pr.append(num_id_el)
    # numPr must appear after pStyle but before most other pPr children per
    # the schema; put it at the front to be safe — Word reorders on save.
    pPr.insert(0, num_pr)


# --------------------------------------------------------------------------- #
# XML construction
# --------------------------------------------------------------------------- #


def _insert_abstract_num(numbering, element) -> None:
    """Insert an ``<w:abstractNum>`` before the first ``<w:num>`` (schema order)."""
    first_num = numbering.find(qn("w:num"))
    if first_num is None:
        numbering.append(element)
    else:
        first_num.addprevious(element)


# Indent schedule (twips). Both schedules put the first-level marker at the
# body red-line position (709 twips ≈ 1.25 cm) so list items honor GOST 7.32
# "абзацный отступ". Numbers indent by 360 twips per level; ordered text
# indents by 540 twips per level so the hanging grows enough to fit
# widening multi-digit prefixes like "1.1.1.1.".
_RED_LINE_TWIPS = 709
_LEVEL_NUM_STEP = 360       # how much the marker shifts right per nesting level
_LEVEL_TEXT_STEP_ORDERED = 540
_LEVEL_TEXT_STEP_BULLET = 360
_HANGING_STEP = 180         # extra hanging per level for wider numbers (ordered)
_HANGING_BASE = 360         # base hanging at level 0 for both kinds


def _build_bullet_abstract():
    abs_num = OxmlElement("w:abstractNum")
    abs_num.set(qn("w:abstractNumId"), str(BULLET_ABSTRACT_ID))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "hybridMultilevel")
    abs_num.append(multi)
    for ilvl in range(9):
        num_x = _RED_LINE_TWIPS + _LEVEL_NUM_STEP * ilvl
        hanging = _HANGING_BASE
        left = num_x + hanging + _LEVEL_TEXT_STEP_BULLET * 0  # text right after the bullet
        lvl = _build_level(
            ilvl=ilvl,
            num_fmt="bullet",
            lvl_text=_BULLET_GLYPHS[ilvl],
            left_indent_twips=left,
            hanging_twips=hanging,
        )
        abs_num.append(lvl)
    return abs_num


def _build_ordered_abstract():
    abs_num = OxmlElement("w:abstractNum")
    abs_num.set(qn("w:abstractNumId"), str(ORDERED_ABSTRACT_ID))
    multi = OxmlElement("w:multiLevelType")
    # hybridMultilevel + per-num lvlOverride/startOverride is the canonical
    # Word recipe for "each list instance restarts at 1". Plain ``multilevel``
    # let Word and LibreOffice apply a "continue across instances" heuristic
    # that visually breaks adjacent independent lists; Google Docs ignores
    # the heuristic, which is why the bug only surfaced in the desktop
    # editors.
    multi.set(qn("w:val"), "hybridMultilevel")
    abs_num.append(multi)
    for ilvl in range(9):
        # "%1.", "%1.%2.", "%1.%2.%3.", ...
        pattern = "".join(f"%{i + 1}." for i in range(ilvl + 1))
        num_x = _RED_LINE_TWIPS + _LEVEL_NUM_STEP * ilvl
        hanging = _HANGING_BASE + _HANGING_STEP * ilvl
        left = num_x + hanging
        lvl = _build_level(
            ilvl=ilvl,
            num_fmt="decimal",
            lvl_text=pattern,
            left_indent_twips=left,
            hanging_twips=hanging,
        )
        abs_num.append(lvl)
    return abs_num


def _build_heading_abstract():
    """Multilevel definition for auto-numbered headings 2..6.

    Each level is bound to the matching ``Heading N`` style via ``<w:pStyle>``,
    so any paragraph carrying that style picks up the corresponding prefix
    ("1", "1.1", "1.1.1", …). ``multiLevelType="multilevel"`` makes Word
    cascade restarts: a new Heading 2 zeroes the Heading 3 counter, and so
    on down the chain.
    """

    abs_num = OxmlElement("w:abstractNum")
    abs_num.set(qn("w:abstractNumId"), str(HEADING_ABSTRACT_ID))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abs_num.append(multi)
    for ilvl in range(_HEADING_DEPTH):
        pattern = ".".join(f"%{i + 1}" for i in range(ilvl + 1))
        heading_style_id = f"Heading{ilvl + 2}"  # ilvl 0 -> Heading2, …
        lvl = _build_heading_level(
            ilvl=ilvl,
            pattern=pattern,
            heading_style_id=heading_style_id,
        )
        abs_num.append(lvl)
    return abs_num


def _build_level(
    *,
    ilvl: int,
    num_fmt: str,
    lvl_text: str,
    left_indent_twips: int,
    hanging_twips: int,
):
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), str(ilvl))

    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    lvl.append(start)

    fmt = OxmlElement("w:numFmt")
    fmt.set(qn("w:val"), num_fmt)
    lvl.append(fmt)

    text = OxmlElement("w:lvlText")
    text.set(qn("w:val"), lvl_text)
    lvl.append(text)

    jc = OxmlElement("w:lvlJc")
    jc.set(qn("w:val"), "left")
    lvl.append(jc)

    pPr = OxmlElement("w:pPr")
    # Explicit numbering tab at `left` so the text after the prefix always
    # aligns to `left` regardless of marker width. Without this Word falls
    # back to default tab stops (every 720 twips) when the marker overflows
    # the hanging area, producing inconsistent gaps between numbered levels.
    tabs = OxmlElement("w:tabs")
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "num")
    tab.set(qn("w:pos"), str(left_indent_twips))
    tabs.append(tab)
    pPr.append(tabs)

    ind = OxmlElement("w:ind")
    ind.set(qn("w:left"), str(left_indent_twips))
    ind.set(qn("w:hanging"), str(hanging_twips))
    pPr.append(ind)
    lvl.append(pPr)

    return lvl


def _build_heading_level(*, ilvl: int, pattern: str, heading_style_id: str):
    """One <w:lvl> entry for the heading abstract.

    Headings sit at the body's red-line position (1.25 cm), so the first
    character of the heading — the marker prefix — lines up with where
    paragraph text would otherwise start. ``<w:suff w:val="space"/>``
    keeps "1.1 Section name" tight; the default ``tab`` suffix would
    leave a wide gap that varies with prefix width.
    """

    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), str(ilvl))

    start = OxmlElement("w:start")
    start.set(qn("w:val"), "1")
    lvl.append(start)

    fmt = OxmlElement("w:numFmt")
    fmt.set(qn("w:val"), "decimal")
    lvl.append(fmt)

    pStyle = OxmlElement("w:pStyle")
    pStyle.set(qn("w:val"), heading_style_id)
    lvl.append(pStyle)

    suff = OxmlElement("w:suff")
    suff.set(qn("w:val"), "space")
    lvl.append(suff)

    text = OxmlElement("w:lvlText")
    text.set(qn("w:val"), pattern)
    lvl.append(text)

    jc = OxmlElement("w:lvlJc")
    jc.set(qn("w:val"), "left")
    lvl.append(jc)

    pPr = OxmlElement("w:pPr")
    ind = OxmlElement("w:ind")
    # firstLine puts the marker (and the first line of the heading text) at
    # 1.25 cm from the left margin; wrapped lines fall back to the margin.
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:firstLine"), str(_RED_LINE_TWIPS))
    pPr.append(ind)
    lvl.append(pPr)

    return lvl


def _build_num(num_id: int, abstract_id: int, *, restart_levels=None):
    """Create ``<w:num>`` referencing *abstract_id*.

    When *restart_levels* is an iterable of ``ilvl`` values, an explicit
    ``<w:lvlOverride w:ilvl="N"><w:startOverride w:val="1"/></w:lvlOverride>``
    is emitted for each. Word and LibreOffice rely on these overrides to
    restart counters across independent list instances backed by the same
    abstract definition.
    """

    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abs_ref = OxmlElement("w:abstractNumId")
    abs_ref.set(qn("w:val"), str(abstract_id))
    num.append(abs_ref)
    if restart_levels is not None:
        for ilvl in restart_levels:
            lvl_override = OxmlElement("w:lvlOverride")
            lvl_override.set(qn("w:ilvl"), str(ilvl))
            start_override = OxmlElement("w:startOverride")
            start_override.set(qn("w:val"), "1")
            lvl_override.append(start_override)
            num.append(lvl_override)
    return num
