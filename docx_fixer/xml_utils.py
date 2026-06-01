from __future__ import annotations

from lxml import etree

from .constants import NS, W_NS

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
