from __future__ import annotations

from .constants import NS
from .numbering import is_toc_style_definition, style_name_value
from .xml_utils import qn, remove_character_indent_attrs

def remove_character_indent_attrs_from_styles_root_excluding_protected(
    root,
    excluded_style_ids: set[str] | None = None,
    change_logs: list[str] | None = None,
) -> int:
    removed = 0
    excluded_style_ids = excluded_style_ids or set()
    for style in root.xpath("./w:style[@w:type='paragraph']", namespaces=NS):
        style_id = style.get(qn("styleId")) or ""
        style_name = style_name_value(style)
        if style_id in excluded_style_ids:
            if change_logs is not None:
                change_logs.append(
                    "CHAR_INDENT_SANITIZE_SKIP_EXCLUDED_STYLE: "
                    f"styleId={style_id}; reason=used_by_chapter_three"
                )
            continue
        if is_toc_style_definition(style_id, style_name):
            continue
        for ind in style.xpath(".//w:ind", namespaces=NS):
            removed += remove_character_indent_attrs(ind)
    return removed


def remove_character_indent_attrs_from_numbering_root_excluding_protected(
    root,
    excluded_numbering_pairs: set[tuple[str, int]],
    excluded_num_ids: set[str],
    excluded_abstract_ids: set[str],
) -> int:
    removed = 0
    num_to_abstract_id: dict[str, str] = {}
    for num in root.xpath("./w:num", namespaces=NS):
        num_id = num.get(qn("numId"))
        abstract_el = num.find("w:abstractNumId", NS)
        abstract_id = abstract_el.get(qn("val")) if abstract_el is not None else None
        if num_id is not None and abstract_id is not None:
            num_to_abstract_id[num_id] = abstract_id

    def should_skip(num_id: str | None, ilvl: int | None, abstract_id: str | None) -> bool:
        if abstract_id is not None and abstract_id in excluded_abstract_ids:
            return True
        if num_id is not None and num_id in excluded_num_ids:
            return True
        if num_id is not None and ilvl is not None and (num_id, ilvl) in excluded_numbering_pairs:
            return True
        return False

    for lvl in root.xpath("./w:abstractNum/w:lvl", namespaces=NS):
        abstract_num = lvl.getparent()
        abstract_id = abstract_num.get(qn("abstractNumId")) if abstract_num is not None else None
        try:
            ilvl = int(lvl.get(qn("ilvl")))
        except Exception:
            ilvl = None
        if should_skip(None, ilvl, abstract_id):
            continue
        for ind in lvl.xpath(".//w:ind", namespaces=NS):
            removed += remove_character_indent_attrs(ind)

    for lvl in root.xpath("./w:num/w:lvlOverride/w:lvl", namespaces=NS):
        override = lvl.getparent()
        num = override.getparent() if override is not None else None
        num_id = num.get(qn("numId")) if num is not None else None
        abstract_id = num_to_abstract_id.get(num_id or "")
        try:
            ilvl = int(override.get(qn("ilvl"))) if override is not None else None
        except Exception:
            ilvl = None
        if should_skip(num_id, ilvl, abstract_id):
            continue
        for ind in lvl.xpath(".//w:ind", namespaces=NS):
            removed += remove_character_indent_attrs(ind)

    return removed