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


# WordprocessingML child order for CT_Lvl (a numbering w:lvl). Word silently
# ignores or repairs out-of-order children, so w:suff written after w:lvlText /
# w:pPr / w:rPr makes the "trailing character" show as None in the UI. Anything
# we add or rebuild on a w:lvl must be placed using this order.
LEVEL_CHILD_ORDER = (
    qn("start"),
    qn("numFmt"),
    qn("lvlRestart"),
    qn("pStyle"),
    qn("isLgl"),
    qn("suff"),
    qn("lvlText"),
    qn("lvlPicBulletId"),
    qn("legacy"),
    qn("lvlJc"),
    qn("pPr"),
    qn("rPr"),
)

# WordprocessingML child order for paragraph properties (CT_PPrBase). The part
# that matters for numbering geometry is w:tabs before w:ind; the full sequence
# is listed so any added child lands at a schema-valid position.
PPR_CHILD_ORDER = (
    qn("pStyle"),
    qn("keepNext"),
    qn("keepLines"),
    qn("pageBreakBefore"),
    qn("framePr"),
    qn("widowControl"),
    qn("numPr"),
    qn("suppressLineNumbers"),
    qn("pBdr"),
    qn("shd"),
    qn("tabs"),
    qn("suppressAutoHyphens"),
    qn("kinsoku"),
    qn("wordWrap"),
    qn("overflowPunct"),
    qn("topLinePunct"),
    qn("autoSpaceDE"),
    qn("autoSpaceDN"),
    qn("bidi"),
    qn("adjustRightInd"),
    qn("snapToGrid"),
    qn("spacing"),
    qn("ind"),
    qn("contextualSpacing"),
    qn("mirrorIndents"),
    qn("suppressOverlap"),
    qn("jc"),
    qn("textDirection"),
    qn("textAlignment"),
    qn("textboxTightWrap"),
    qn("outlineLvl"),
    qn("divId"),
    qn("cnfStyle"),
)


def insert_child_in_schema_order(parent, child, ordered_tags: tuple[str, ...]) -> None:
    """Insert `child` into `parent` at the position dictated by `ordered_tags`.

    Any existing children sharing `child`'s tag are removed first (so there are
    never duplicates, and a mis-ordered existing element is repositioned rather
    than left in place). The element is inserted before the first sibling whose
    tag sorts strictly later in `ordered_tags`; siblings with unknown/extension
    tags are skipped as ordering anchors and otherwise left untouched. A child
    whose tag is not in `ordered_tags` is appended.
    """
    tag = child.tag
    for existing in parent.findall(tag):
        parent.remove(existing)

    if tag in ordered_tags:
        tag_pos = ordered_tags.index(tag)
        for sibling in parent:
            sib_tag = sibling.tag
            if sib_tag in ordered_tags and ordered_tags.index(sib_tag) > tag_pos:
                sibling.addprevious(child)
                return

    parent.append(child)


def get_or_add_in_schema_order(parent, tag: str, ordered_tags: tuple[str, ...]):
    """Find or create `w:tag` and (re)place it at the schema-correct position.

    Existing content is preserved (the same element is moved, not recreated);
    duplicates are collapsed to one.
    """
    child = parent.find(f"w:{tag}", NS)
    if child is None:
        child = etree.Element(qn(tag))
    insert_child_in_schema_order(parent, child, ordered_tags)
    return child


def child_is_in_schema_order(parent, child, ordered_tags: tuple[str, ...]) -> bool:
    """Whether `child` currently sits at a schema-valid position among siblings."""
    tag = child.tag
    if tag not in ordered_tags:
        return True
    tag_pos = ordered_tags.index(tag)
    seen_self = False
    for sibling in parent:
        if sibling is child:
            seen_self = True
            continue
        sib_tag = sibling.tag
        if sib_tag not in ordered_tags:
            continue
        sib_pos = ordered_tags.index(sib_tag)
        if not seen_self and sib_pos > tag_pos:
            return False
        if seen_self and sib_pos < tag_pos:
            return False
    return True


def children_in_schema_order(parent, ordered_tags: tuple[str, ...]) -> bool:
    """Whether all known children of `parent` appear in non-decreasing order."""
    last_pos = -1
    for child in parent:
        if child.tag not in ordered_tags:
            continue
        pos = ordered_tags.index(child.tag)
        if pos < last_pos:
            return False
        last_pos = pos
    return True


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
        if attr_name in ind.attrib and ind.get(attr_name) != "0":
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
                        "CHAR_INDENT_SANITIZE_SKIP_EXCLUDED: "
                        f"part={part_name} paragraph_index={paragraph_index} text={preview}"
                    )
                continue
            for ind in p.xpath(".//w:ind", namespaces=NS):
                removed += remove_character_indent_attrs(ind)
        return removed

    for ind in root.xpath(".//w:ind", namespaces=NS):
        removed += remove_character_indent_attrs(ind)
    return removed
