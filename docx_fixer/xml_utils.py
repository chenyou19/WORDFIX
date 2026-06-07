from __future__ import annotations

from lxml import etree

from .constants import NS, W_NS

CHAR_INDENT_ATTRS = [
    "leftChars",
    "startChars",
    "rightChars",
    "endChars",
    "firstLineChars",
    "hangingChars",
]

def qn(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def get_or_add(parent, tag: str, first: bool = False):
    child = parent.find(f"w:{tag}", NS)
    if child is None:
        child = etree.Element(qn(tag))
        if first:
            parent.insert(0, child)
        else:
            parent.append(child)
    return child


def paragraph_text(p) -> str:
    texts = p.xpath(".//w:t/text()", namespaces=NS)
    return "".join(texts)


def remove_character_indent_attrs(ind) -> int:
    """
    Remove only character-based indent attributes from one w:ind element.

    Twips-based attributes such as w:left, w:start, w:hanging, and w:firstLine
    are intentionally preserved.
    """
    removed = 0
    for attr in CHAR_INDENT_ATTRS:
        attr_name = qn(attr)
        if attr_name in ind.attrib:
            ind.attrib.pop(attr_name, None)
            removed += 1
    return removed


def remove_character_indent_attrs_from_root(
    root,
    exclude_paragraph_ids: set[int] | None = None,
    change_logs: list[str] | None = None,
    part_name: str = "word/document.xml",
) -> int:
    """
    Remove character-based indent attributes from every w:ind under root.

    Returns the number of attributes actually removed.
    """
    removed = 0
    if exclude_paragraph_ids is not None:
        for paragraph_index, p in enumerate(root.xpath(".//w:p", namespaces=NS), start=1):
            if id(p) in exclude_paragraph_ids:
                if change_logs is not None and p.xpath(".//w:ind", namespaces=NS):
                    text = paragraph_text(p)
                    preview = " ".join((text or "").split())
                    if len(preview) > 80:
                        preview = preview[:77] + "..."
                    change_logs.append(
                        "CHAR_INDENT_SANITIZE_SKIP_TOC: "
                        f"part={part_name} paragraph_index={paragraph_index} text={preview}"
                    )
                continue
            for ind in p.xpath(".//w:ind", namespaces=NS):
                removed += remove_character_indent_attrs(ind)
        return removed

    for ind in root.xpath(".//w:ind", namespaces=NS):
        removed += remove_character_indent_attrs(ind)
    return removed
