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


def remove_character_indent_attrs_from_root(root) -> int:
    """
    Remove character-based indent attributes from every w:ind under root.

    Returns the number of attributes actually removed.
    """
    removed = 0
    for ind in root.xpath(".//w:ind", namespaces=NS):
        removed += remove_character_indent_attrs(ind)
    return removed
