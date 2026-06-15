from __future__ import annotations

from .constants import NS
from .xml_utils import qn

NOTE_PREFIX = "註"
CONTROL_SPACE_CHARS = "\ufeff\u200b\u200c\u200d\u2060"


def normalize_note_text(text: str) -> str:
    cleaned = str(text or "")
    for char in CONTROL_SPACE_CHARS:
        cleaned = cleaned.replace(char, "")
    return " ".join(cleaned.split()).lstrip()


def is_note_text(text: str) -> bool:
    return normalize_note_text(text).startswith(NOTE_PREFIX)


def numbering_lvl_text_starts_with_note(
    num_id,
    ilvl,
    numbering_format_lookup=None,
) -> bool:
    if num_id is None or ilvl is None or not numbering_format_lookup:
        return False
    try:
        ilvl_int = int(ilvl)
    except Exception:
        return False
    level_format = numbering_format_lookup.get((str(num_id), ilvl_int), {})
    return is_note_text(str(level_format.get("lvlText") or ""))


def _paragraph_num_identity(p) -> tuple[str | None, int | None]:
    ilvl_el = p.find("./w:pPr/w:numPr/w:ilvl", NS)
    num_id_el = p.find("./w:pPr/w:numPr/w:numId", NS)
    num_id = num_id_el.get(qn("val")) if num_id_el is not None else None
    try:
        ilvl = int(ilvl_el.get(qn("val"))) if ilvl_el is not None else 0
    except Exception:
        ilvl = None
    return num_id, ilvl


def _paragraph_style_id(p) -> str | None:
    style_el = p.find("./w:pPr/w:pStyle", NS)
    return style_el.get(qn("val")) if style_el is not None else None


def note_source_for_paragraph(
    p,
    text: str,
    numbering_format_lookup=None,
    style_numbering_lookup=None,
) -> str | None:
    if is_note_text(text):
        return "text"

    num_id, ilvl = _paragraph_num_identity(p)
    if numbering_lvl_text_starts_with_note(num_id, ilvl, numbering_format_lookup):
        return "numPr"

    style_id = _paragraph_style_id(p)
    if style_id and style_numbering_lookup:
        style_num_id, style_ilvl = style_numbering_lookup.get(style_id, (None, None))
        if numbering_lvl_text_starts_with_note(style_num_id, style_ilvl, numbering_format_lookup):
            return "styleNumPr"

    return None
