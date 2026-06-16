from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from lxml import etree

from .constants import NS
from .numbering import has_auto_numbering, paragraph_style_id
from .outline import (
    collect_all_toc_paragraph_ids,
    detect_manual_numbering_prefix,
    get_auto_number_identity,
)
from .xml_utils import paragraph_text, qn

CHAPTER_THREE_SKIP_TITLE = "價格形成之主要因素分析"
CHAPTER_THREE_SKIP_VISIBLE_PREFIXES = (
    "參、價格形成之主要因素分析",
)

_TRADITIONAL_LEGAL_CHAPTER_NUMBERS = {
    1: "壹",
    2: "貳",
    3: "參",
    4: "肆",
    5: "伍",
    6: "陸",
    7: "柒",
    8: "捌",
    9: "玖",
    10: "拾",
}

_TRADITIONAL_COUNTING_CHAPTER_NUMBERS = {
    1: "一",
    2: "二",
    3: "三",
    4: "四",
    5: "五",
    6: "六",
    7: "七",
    8: "八",
    9: "九",
    10: "十",
}


def _effective_paragraph_numbering_identity(
    p,
    style_numbering_lookup,
) -> tuple[str | None, int | None]:
    num_id, ilvl = get_auto_number_identity(p)
    if num_id is not None and ilvl is not None:
        return num_id, ilvl

    style_id = paragraph_style_id(p)
    if style_id and style_numbering_lookup:
        return style_numbering_lookup.get(style_id, (None, None))

    return num_id, ilvl


def _outline_level_from_identity(
    num_id,
    ilvl,
    numbering_level_lookup,
) -> int | None:
    if num_id is None or ilvl is None:
        return None

    # The level lookup only contains numbering pairs whose numFmt + lvlText
    # matched a supported heading pattern. ilvl alone must never become an
    # outline level, otherwise leftover numPr on body paragraphs would be
    # misclassified as headings.
    if not numbering_level_lookup:
        return None

    return numbering_level_lookup.get((num_id, ilvl))


def _chapter_number_token_from_format(num_fmt: str | None, ordinal: int) -> str | None:
    if ordinal <= 0:
        return None

    fmt = (num_fmt or "").strip()
    if fmt in {"ideographLegalTraditional", "chineseLegalSimplified"}:
        return _TRADITIONAL_LEGAL_CHAPTER_NUMBERS.get(ordinal)
    if fmt in {"taiwaneseCountingThousand", "ideographTraditional", "chineseCounting"}:
        return _TRADITIONAL_COUNTING_CHAPTER_NUMBERS.get(ordinal)
    return None


def _count_same_stream_first_level_headings_before_paragraph(
    p,
    *,
    num_id,
    ilvl,
    numbering_level_lookup,
    style_numbering_lookup,
) -> int:
    count = 0
    paragraphs = p.xpath("preceding::w:p[not(ancestor::w:tbl)]", namespaces=NS)
    for candidate in [*paragraphs, p]:
        candidate_num_id, candidate_ilvl = _effective_paragraph_numbering_identity(
            candidate,
            style_numbering_lookup,
        )
        if (candidate_num_id, candidate_ilvl) != (num_id, ilvl):
            continue

        candidate_level = _outline_level_from_identity(
            candidate_num_id,
            candidate_ilvl,
            numbering_level_lookup,
        )
        if candidate_level == 0:
            count += 1

    return count


def _first_level_heading_prefix_for_paragraph(
    p,
    *,
    numbering_level_lookup,
    numbering_format_lookup,
    style_numbering_lookup,
) -> str | None:
    text = paragraph_text(p).strip()
    manual = detect_manual_numbering_prefix(text)
    if manual is not None:
        level, prefix = manual
        if level == 0:
            return prefix
        return None

    num_id, ilvl = _effective_paragraph_numbering_identity(p, style_numbering_lookup)
    level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)
    if level != 0 or num_id is None or ilvl is None:
        return None

    level_format = numbering_format_lookup.get((num_id, ilvl), {})
    ordinal = _count_same_stream_first_level_headings_before_paragraph(
        p,
        num_id=num_id,
        ilvl=ilvl,
        numbering_level_lookup=numbering_level_lookup,
        style_numbering_lookup=style_numbering_lookup,
    )
    token = _chapter_number_token_from_format(level_format.get("numFmt"), ordinal)
    lvl_text = level_format.get("lvlText")
    if token is None or lvl_text is None or "%1" not in lvl_text:
        return None

    return lvl_text.replace("%1", token)


def _compact_heading_text(text: str) -> str:
    return "".join((text or "").split())


def is_chapter_three_start_marker(
    p,
    text: str,
    *,
    numbering_level_lookup=None,
    numbering_format_lookup=None,
    style_numbering_lookup=None,
) -> bool:
    del numbering_format_lookup  # Kept for parity with chapter prefix helpers.
    compact = _compact_heading_text(text)
    visible_prefixes = tuple(
        _compact_heading_text(prefix)
        for prefix in CHAPTER_THREE_SKIP_VISIBLE_PREFIXES
    )
    if compact.startswith(visible_prefixes):
        return True

    if not compact.startswith(_compact_heading_text(CHAPTER_THREE_SKIP_TITLE)):
        return False

    level = None
    if has_auto_numbering(p):
        num_id, ilvl = _effective_paragraph_numbering_identity(p, style_numbering_lookup)
        level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)

    if level is None:
        num_id, ilvl = _effective_paragraph_numbering_identity(p, style_numbering_lookup)
        level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)

    if level is None:
        manual = detect_manual_numbering_prefix(text)
        if manual is not None:
            level = manual[0]

    return level == 0


SECTION_THREE_CHAPTER_TOKEN = "參"


def is_section_three_chapter_marker(
    p,
    text=None,
    *,
    numbering_level_lookup=None,
    numbering_format_lookup=None,
    style_numbering_lookup=None,
) -> bool:
    """Generic detector for the body chapter 「參、」 by chapter number.

    Unlike is_chapter_three_start_marker (which is tied to the specific title
    「價格形成之主要因素分析」), this recognises any first-level heading whose
    chapter number resolves to 參 (the 3rd 壹貳參 chapter). It reuses the same
    first-level-heading resolution that already excludes TOC entries, so a 參、
    line inside a table of contents does not trigger protection.
    """
    del text  # Accepted for predicate-signature parity; not needed here.
    prefix = _first_level_heading_prefix_for_paragraph(
        p,
        numbering_level_lookup=numbering_level_lookup,
        numbering_format_lookup=numbering_format_lookup,
        style_numbering_lookup=style_numbering_lookup,
    )
    if not prefix:
        return False
    return prefix.lstrip().startswith(SECTION_THREE_CHAPTER_TOKEN)


def is_table_under_chapter_three(
    tbl,
    numbering_level_lookup,
    numbering_format_lookup,
    style_numbering_lookup,
) -> bool:
    paragraphs = tbl.xpath("preceding::w:p[not(ancestor::w:tbl)]", namespaces=NS)
    for p in reversed(paragraphs):
        text = paragraph_text(p)
        prefix = _first_level_heading_prefix_for_paragraph(
            p,
            numbering_level_lookup=numbering_level_lookup,
            numbering_format_lookup=numbering_format_lookup,
            style_numbering_lookup=style_numbering_lookup,
        )
        if prefix is not None:
            return is_chapter_three_start_marker(
                p,
                text,
                numbering_level_lookup=numbering_level_lookup,
                numbering_format_lookup=numbering_format_lookup,
                style_numbering_lookup=style_numbering_lookup,
            )
    return False


def find_table_first_level_heading(
    tbl,
    numbering_level_lookup,
    numbering_format_lookup,
    style_numbering_lookup,
) -> str | None:
    paragraphs = tbl.xpath("preceding::w:p[not(ancestor::w:tbl)]", namespaces=NS)
    for p in reversed(paragraphs):
        prefix = _first_level_heading_prefix_for_paragraph(
            p,
            numbering_level_lookup=numbering_level_lookup,
            numbering_format_lookup=numbering_format_lookup,
            style_numbering_lookup=style_numbering_lookup,
        )
        if prefix is not None:
            return prefix
    return None


def collect_chapter_three_paragraph_ids(
    root,
    *,
    numbering_level_lookup,
    numbering_format_lookup,
    style_numbering_lookup,
    toc_paragraph_ids=None,
    paragraphs=None,
    start_marker=None,
) -> set[int]:
    """Collect paragraphs from the chapter 參 start marker until the next first-level heading.

    start_marker selects which paragraph begins the protected region. It
    defaults to the title-specific 「參、價格形成之主要因素分析」 detector; pass
    is_section_three_chapter_marker for the generic 「參、不要調整」 behaviour.
    """
    if start_marker is None:
        start_marker = is_chapter_three_start_marker

    skip_ids: set[int] = set()
    toc_ids = toc_paragraph_ids or set()
    in_chapter_three = False

    paragraphs = paragraphs if paragraphs is not None else root.xpath(".//w:p", namespaces=NS)
    for p in paragraphs:
        paragraph_id = id(p)
        if paragraph_id in toc_ids:
            continue

        text = paragraph_text(p)
        is_first_level_heading = False
        prefix = _first_level_heading_prefix_for_paragraph(
            p,
            numbering_level_lookup=numbering_level_lookup,
            numbering_format_lookup=numbering_format_lookup,
            style_numbering_lookup=style_numbering_lookup,
        )
        if prefix is not None:
            is_first_level_heading = True
        else:
            num_id, ilvl = _effective_paragraph_numbering_identity(p, style_numbering_lookup)
            level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)

            if level is None:
                manual = detect_manual_numbering_prefix(text.strip())
                if manual is not None:
                    level = manual[0]

            is_first_level_heading = level == 0

        if start_marker(
            p,
            text,
            numbering_level_lookup=numbering_level_lookup,
            numbering_format_lookup=numbering_format_lookup,
            style_numbering_lookup=style_numbering_lookup,
        ):
            in_chapter_three = True
        elif in_chapter_three and is_first_level_heading:
            in_chapter_three = False

        if in_chapter_three:
            skip_ids.add(paragraph_id)

    return skip_ids

def collect_toc_numbering_exclusions(
    document_root,
    toc_paragraph_ids: set[int],
    style_numbering_lookup: dict[str, tuple[str, int]],
    numbering_xml: bytes | None,
    paragraphs=None,
) -> tuple[set[tuple[str, int]], set[str], set[str]]:
    pairs: set[tuple[str, int]] = set()
    num_ids: set[str] = set()
    abstract_ids: set[str] = set()
    num_to_abstract_id: dict[str, str] = {}

    if numbering_xml:
        try:
            numbering_root = etree.fromstring(numbering_xml)
            for num in numbering_root.xpath("./w:num", namespaces=NS):
                num_id = num.get(qn("numId"))
                abstract_el = num.find("w:abstractNumId", NS)
                abstract_id = abstract_el.get(qn("val")) if abstract_el is not None else None
                if num_id is not None and abstract_id is not None:
                    num_to_abstract_id[num_id] = abstract_id
        except Exception:
            pass

    paragraphs = paragraphs if paragraphs is not None else document_root.xpath(".//w:p", namespaces=NS)
    for p in paragraphs:
        if id(p) not in toc_paragraph_ids:
            continue

        num_id = None
        ilvl = None
        if has_auto_numbering(p):
            num_id, ilvl = get_auto_number_identity(p)
        if num_id is None:
            style_id = paragraph_style_id(p)
            if style_id:
                num_id, ilvl = style_numbering_lookup.get(style_id, (None, None))

        if num_id is None:
            continue
        if ilvl is None:
            ilvl = 0

        num_ids.add(str(num_id))
        pairs.add((str(num_id), int(ilvl)))
        abstract_id = num_to_abstract_id.get(str(num_id))
        if abstract_id is not None:
            abstract_ids.add(abstract_id)

    return pairs, num_ids, abstract_ids


def collect_body_heading_paragraph_ids(
    document_root,
    toc_paragraph_ids: set[int],
    *,
    numbering_level_lookup,
    style_numbering_lookup,
    paragraphs=None,
) -> set[int]:
    del document_root
    heading_ids: set[int] = set()
    toc_ids = toc_paragraph_ids or set()
    paragraphs = paragraphs or []

    for p in paragraphs:
        paragraph_id = id(p)
        if paragraph_id in toc_ids:
            continue
        if p.xpath("ancestor::w:tbl", namespaces=NS):
            continue

        text = paragraph_text(p)
        num_id, ilvl = _effective_paragraph_numbering_identity(p, style_numbering_lookup)
        level = _outline_level_from_identity(num_id, ilvl, numbering_level_lookup)

        if level is None:
            manual = detect_manual_numbering_prefix(text.strip())
            if manual is not None:
                level = manual[0]

        if level is not None and 0 <= level <= 8:
            heading_ids.add(paragraph_id)

    return heading_ids


@dataclass
class ProtectedRegionContext:
    document_toc_paragraph_ids: set[int] = field(default_factory=set)
    document_chapter_three_paragraph_ids: set[int] = field(default_factory=set)
    document_body_heading_paragraph_ids: set[int] = field(default_factory=set)
    toc_numbering_pairs: set[tuple[str, int]] = field(default_factory=set)
    toc_num_ids: set[str] = field(default_factory=set)
    toc_abstract_ids: set[str] = field(default_factory=set)
    chapter_three_numbering_pairs: set[tuple[str, int]] = field(default_factory=set)
    chapter_three_num_ids: set[str] = field(default_factory=set)
    chapter_three_abstract_ids: set[str] = field(default_factory=set)
    chapter_three_style_ids: set[str] = field(default_factory=set)
    body_heading_numbering_pairs: set[tuple[str, int]] = field(default_factory=set)
    body_heading_num_ids: set[str] = field(default_factory=set)
    body_heading_abstract_ids: set[str] = field(default_factory=set)
    section_three_protection_enabled: bool = False
    section_three_detection_source: str = "none"
    document_section_three_note_paragraph_ids: set[int] = field(default_factory=set)
    section_three_note_region_enabled: bool = False
    _document_paragraph_refs: list = field(default_factory=list, repr=False)

    @classmethod
    def from_document(
        cls,
        document_root,
        *,
        protect_chapter_three: bool,
        numbering_level_lookup,
        numbering_format_lookup,
        style_numbering_lookup,
        numbering_xml: bytes | None,
        summary: Any | None = None,
        use_generic_section_three: bool = False,
        collect_section_three_note_region: bool = False,
    ) -> "ProtectedRegionContext":
        paragraphs = document_root.xpath(".//w:p", namespaces=NS)
        toc_paragraph_ids = collect_all_toc_paragraph_ids(
            document_root,
            numbering_level_lookup=numbering_level_lookup,
            style_numbering_lookup=style_numbering_lookup,
            paragraphs=paragraphs,
        )
        toc_pairs, toc_num_ids, toc_abstract_ids = collect_toc_numbering_exclusions(
            document_root,
            toc_paragraph_ids,
            style_numbering_lookup,
            numbering_xml,
            paragraphs=paragraphs,
        )

        context = cls(
            document_toc_paragraph_ids=toc_paragraph_ids,
            toc_numbering_pairs=toc_pairs,
            toc_num_ids=toc_num_ids,
            toc_abstract_ids=toc_abstract_ids,
            _document_paragraph_refs=list(paragraphs),
        )

        context.document_body_heading_paragraph_ids = collect_body_heading_paragraph_ids(
            document_root,
            toc_paragraph_ids,
            numbering_level_lookup=numbering_level_lookup,
            style_numbering_lookup=style_numbering_lookup,
            paragraphs=paragraphs,
        )
        (
            context.body_heading_numbering_pairs,
            context.body_heading_num_ids,
            context.body_heading_abstract_ids,
        ) = collect_toc_numbering_exclusions(
            document_root,
            context.document_body_heading_paragraph_ids,
            style_numbering_lookup,
            numbering_xml,
            paragraphs=paragraphs,
        )

        # The note-skip region ("參、不要表格註記搬移") is decoupled from the
        # layout/color/indent protection: it always uses the generic 參、
        # chapter detector and is collected independently, so enabling note
        # protection never changes which tables get layout/color protected.
        if collect_section_three_note_region:
            context.section_three_note_region_enabled = True
            context.document_section_three_note_paragraph_ids = collect_chapter_three_paragraph_ids(
                document_root,
                numbering_level_lookup=numbering_level_lookup,
                numbering_format_lookup=numbering_format_lookup,
                style_numbering_lookup=style_numbering_lookup,
                toc_paragraph_ids=toc_paragraph_ids,
                paragraphs=paragraphs,
                start_marker=is_section_three_chapter_marker,
            )
            if summary is not None:
                summary.numbering_xml_logs.append(
                    "SECTION_THREE_TABLE_NOTE_SKIP_IDS collected="
                    f"{len(context.document_section_three_note_paragraph_ids)} "
                    "detection_source=generic_section_three_chapter_參"
                )

        if not protect_chapter_three:
            return context

        context.section_three_protection_enabled = use_generic_section_three
        start_marker = (
            is_section_three_chapter_marker
            if use_generic_section_three
            else is_chapter_three_start_marker
        )
        context.section_three_detection_source = (
            "generic_section_three_chapter_參"
            if use_generic_section_three
            else "title_specific_價格形成之主要因素分析"
        )
        context.document_chapter_three_paragraph_ids = collect_chapter_three_paragraph_ids(
            document_root,
            numbering_level_lookup=numbering_level_lookup,
            numbering_format_lookup=numbering_format_lookup,
            style_numbering_lookup=style_numbering_lookup,
            toc_paragraph_ids=toc_paragraph_ids,
            paragraphs=paragraphs,
            start_marker=start_marker,
        )
        if summary is not None:
            summary.numbering_xml_logs.append(
                f"CHAPTER_THREE_SKIP_IDS collected={len(context.document_chapter_three_paragraph_ids)} "
                f"detection_source={context.section_three_detection_source}"
            )

        (
            context.chapter_three_numbering_pairs,
            context.chapter_three_num_ids,
            context.chapter_three_abstract_ids,
        ) = collect_toc_numbering_exclusions(
            document_root,
            context.document_chapter_three_paragraph_ids,
            style_numbering_lookup,
            numbering_xml,
            paragraphs=paragraphs,
        )
        context.chapter_three_style_ids = {
            style_id
            for p in paragraphs
            if id(p) in context.document_chapter_three_paragraph_ids
            for style_id in [paragraph_style_id(p)]
            if style_id
        }
        return context

    @property
    def excluded_numbering_pairs(self) -> set[tuple[str, int]]:
        return set(self.toc_numbering_pairs) | set(self.chapter_three_numbering_pairs)

    @property
    def excluded_num_ids(self) -> set[str]:
        return set(self.toc_num_ids) | set(self.chapter_three_num_ids)

    @property
    def excluded_abstract_ids(self) -> set[str]:
        return set(self.toc_abstract_ids) | set(self.chapter_three_abstract_ids)

    def chapter_three_paragraph_ids_for_part(self, part_name: str) -> set[int] | None:
        if part_name != "word/document.xml" or not self.document_chapter_three_paragraph_ids:
            return None
        return self.document_chapter_three_paragraph_ids

    def toc_paragraph_ids_for_part(self, part_name: str) -> set[int] | None:
        if part_name != "word/document.xml" or not self.document_toc_paragraph_ids:
            return None
        return self.document_toc_paragraph_ids

    def sanitize_excluded_paragraph_ids_for_part(self, part_name: str) -> set[int] | None:
        if part_name != "word/document.xml":
            return None
        ids = set(self.document_toc_paragraph_ids)
        ids.update(self.document_chapter_three_paragraph_ids)
        return ids or None

    def is_paragraph_protected(self, p, part_name: str = "word/document.xml") -> bool:
        return (
            part_name == "word/document.xml"
            and id(p) in self.document_chapter_three_paragraph_ids
        )

    def is_table_protected(self, tbl, part_name: str = "word/document.xml") -> bool:
        if part_name != "word/document.xml" or not self.document_chapter_three_paragraph_ids:
            return False
        return any(
            self.is_paragraph_protected(p, part_name)
            for p in tbl.xpath(".//w:p", namespaces=NS)
        )

    def is_table_in_section_three_for_notes(
        self, tbl, part_name: str = "word/document.xml"
    ) -> bool:
        """Whether the table sits in the generic body 參、 chapter, used only to
        decide whether table note cells should be left in place."""
        if part_name != "word/document.xml" or not self.document_section_three_note_paragraph_ids:
            return False
        return any(
            id(p) in self.document_section_three_note_paragraph_ids
            for p in tbl.xpath(".//w:p", namespaces=NS)
        )

    def protected_reason(self, item_type: str = "content") -> str:
        del item_type
        return (
            "under chapter 參、價格形成之主要因素分析; "
            "table layout and color fixes skipped by option skip_chapter_three_tables"
        )

    @property
    def log_reason(self) -> str:
        return "chapter 參、價格形成之主要因素分析 protected region"
