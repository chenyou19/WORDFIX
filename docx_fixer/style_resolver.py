from __future__ import annotations

from lxml import etree

from .constants import NS
from .xml_utils import qn


def half_points_to_pt(value: str | None) -> float | None:
    try:
        return int(value or "") / 2
    except (TypeError, ValueError):
        return None


def rpr_font_size_pt(rPr) -> float | None:
    if rPr is None:
        return None

    for tag in ("sz", "szCs"):
        size = rPr.find(f"w:{tag}", NS)
        if size is None:
            continue
        value = half_points_to_pt(size.get(qn("val")))
        if value is not None:
            return value

    return None


def build_style_font_size_lookup(styles_xml: bytes | None) -> dict[str, object]:
    lookup: dict[str, object] = {
        "paragraph": {},
        "character": {},
        "docDefaults": None,
    }
    if not styles_xml:
        return lookup

    try:
        root = etree.fromstring(styles_xml)
    except Exception:
        return lookup

    doc_default_rpr = root.find("./w:docDefaults/w:rPrDefault/w:rPr", NS)
    doc_default_size = rpr_font_size_pt(doc_default_rpr)
    lookup["docDefaults"] = doc_default_size

    raw: dict[str, dict[str, dict[str, object]]] = {
        "paragraph": {},
        "character": {},
    }
    for style in root.xpath("./w:style", namespaces=NS):
        style_type = style.get(qn("type"))
        if style_type not in raw:
            continue

        style_id = style.get(qn("styleId"))
        if not style_id:
            continue

        based_on_el = style.find("w:basedOn", NS)
        based_on = based_on_el.get(qn("val")) if based_on_el is not None else None
        raw[style_type][style_id] = {
            "size": rpr_font_size_pt(style.find("w:rPr", NS)),
            "basedOn": based_on,
        }

    def resolve(style_type: str, style_id: str, resolving: set[str]) -> float | None:
        resolved = lookup[style_type]
        if isinstance(resolved, dict) and style_id in resolved:
            return resolved[style_id]
        if style_id in resolving:
            return None

        item = raw[style_type].get(style_id)
        if item is None:
            return None

        resolving.add(style_id)
        size = item.get("size")
        if size is None:
            based_on = item.get("basedOn")
            if isinstance(based_on, str) and based_on:
                size = resolve(style_type, based_on, resolving)
        if size is None:
            size = doc_default_size
        resolving.remove(style_id)

        if isinstance(resolved, dict) and size is not None:
            resolved[style_id] = size
        return size if isinstance(size, float) else None

    for style_type, styles in raw.items():
        for style_id in styles:
            resolve(style_type, style_id, set())

    return lookup
