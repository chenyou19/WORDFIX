from __future__ import annotations

from lxml import etree

from .constants import NS, OUTLINE_LEVEL_FONT_SIZE_PT, TEMPLATE_OUTLINE_INDENTS
from .xml_utils import get_or_add, qn

BULLET_OUTLINE_LEVEL = -1
PAREN_PAIRS = {
    "(": ")",
    "（": "）",
}


def numbering_suffix_for_level(level: int) -> str:
    return "nothing"


def sanitize_numbering_level_suffix_tabs_and_text(lvl) -> bool:
    changed = False
    suff = lvl.find("w:suff", NS)
    if suff is None:
        suff = etree.Element(qn("suff"))
        lvl.append(suff)
        changed = True
    if suff.get(qn("val")) != "nothing":
        suff.set(qn("val"), "nothing")
        changed = True

    pPr = lvl.find("w:pPr", NS)
    tabs = pPr.find("w:tabs", NS) if pPr is not None else None
    if tabs is not None:
        pPr.remove(tabs)
        changed = True

    lvl_text = lvl.find("w:lvlText", NS)
    if lvl_text is not None:
        value = lvl_text.get(qn("val"))
        if value is not None:
            stripped = value.rstrip(" \t\u3000")
            if stripped != value:
                lvl_text.set(qn("val"), stripped)
                changed = True

    return changed


def has_auto_numbering(p) -> bool:
    return p.find("./w:pPr/w:numPr", NS) is not None


def paragraph_style_id(p) -> str | None:
    style_el = p.find("./w:pPr/w:pStyle", NS)
    if style_el is None:
        return None
    return style_el.get(qn("val"))


def normalize_lvl_text(value: str | None) -> str:
    if value is None:
        return ""
    return value.strip().replace("％", "%")


def is_bullet_num_fmt(num_fmt: str | None) -> bool:
    return (num_fmt or "").strip().lower() == "bullet"


def is_parenthesized_lvl_text(text: str) -> bool:
    if len(text) < 2:
        return False

    closing = PAREN_PAIRS.get(text[0])
    return closing is not None and text[-1] == closing


def numbering_pattern_to_outline_level(num_fmt: str | None, lvl_text: str | None):
    """
    Convert Word numbering.xml numFmt + lvlText values to template levels.

    Template automatic numbering usually assigns one numId per marker style,
    so ilvl is almost always 0. Therefore, the actual numbering format must be
    inspected instead of relying only on ilvl.
    """
    fmt = (num_fmt or "").strip()
    text = normalize_lvl_text(lvl_text)

    bracketed = is_parenthesized_lvl_text(text)
    has_dot = text.endswith((".", "．")) or "." in text or "．" in text
    has_ideographic_separator = "、" in text

    # Financial Chinese numerals with an ideographic separator.
    if fmt in {"ideographLegalTraditional", "chineseLegalSimplified"}:
        if has_ideographic_separator:
            return 0

    # Simple Chinese numerals, either plain or parenthesized.
    if fmt in {"taiwaneseCountingThousand", "ideographTraditional", "chineseCounting"}:
        if bracketed:
            return 2
        if has_ideographic_separator:
            return 1

    # Decimal numbering, either dotted or parenthesized.
    if fmt == "decimal":
        if bracketed:
            return 4
        if has_dot or has_ideographic_separator:
            return 3

    # Uppercase letters, either dotted or parenthesized.
    if fmt == "upperLetter":
        if bracketed:
            return 6
        if has_dot or has_ideographic_separator:
            return 5

    # Lowercase letters, either dotted or parenthesized.
    if fmt == "lowerLetter":
        if bracketed:
            return 8
        if has_dot or has_ideographic_separator:
            return 7

    return None


def build_numbering_level_lookup(numbering_xml: bytes | None):
    """
    Build a {(numId, ilvl): outline_level} lookup for template-level matching.
    """
    if not numbering_xml:
        return {}

    try:
        root = etree.fromstring(numbering_xml)
    except Exception:
        return {}

    abstract_levels: dict[str, dict[int, int]] = {}

    for abstract_num in root.xpath("./w:abstractNum", namespaces=NS):
        abstract_id = abstract_num.get(qn("abstractNumId"))
        if abstract_id is None:
            continue

        levels: dict[int, int] = {}
        for lvl in abstract_num.xpath("./w:lvl", namespaces=NS):
            try:
                ilvl = int(lvl.get(qn("ilvl")))
            except Exception:
                continue

            num_fmt_el = lvl.find("w:numFmt", NS)
            lvl_text_el = lvl.find("w:lvlText", NS)
            num_fmt = num_fmt_el.get(qn("val")) if num_fmt_el is not None else None
            lvl_text = lvl_text_el.get(qn("val")) if lvl_text_el is not None else None
            if is_bullet_num_fmt(num_fmt):
                levels[ilvl] = BULLET_OUTLINE_LEVEL
                continue

            outline_level = numbering_pattern_to_outline_level(num_fmt, lvl_text)

            # If numbering.xml cannot identify the format directly but this is
            # a Word automatic numbering level, use ilvl as the outline level.
            # This matches the assumption that every automatic numbering entry
            # in the document represents an outline level.
            if outline_level is None and 0 <= ilvl <= 8:
                outline_level = ilvl

            if outline_level is not None:
                levels[ilvl] = outline_level

        abstract_levels[abstract_id] = levels

    lookup: dict[tuple[str, int], int] = {}

    for num in root.xpath("./w:num", namespaces=NS):
        num_id = num.get(qn("numId"))
        abstract_el = num.find("w:abstractNumId", NS)
        if num_id is None or abstract_el is None:
            continue

        abstract_id = abstract_el.get(qn("val"))
        for ilvl, outline_level in abstract_levels.get(abstract_id, {}).items():
            lookup[(num_id, ilvl)] = outline_level

        # Apply per-num lvlOverride definitions to the lookup as well.
        for override in num.xpath("./w:lvlOverride", namespaces=NS):
            try:
                ilvl = int(override.get(qn("ilvl")))
            except Exception:
                continue

            lvl = override.find("w:lvl", NS)
            if lvl is None:
                continue

            num_fmt_el = lvl.find("w:numFmt", NS)
            lvl_text_el = lvl.find("w:lvlText", NS)
            num_fmt = num_fmt_el.get(qn("val")) if num_fmt_el is not None else None
            lvl_text = lvl_text_el.get(qn("val")) if lvl_text_el is not None else None
            if is_bullet_num_fmt(num_fmt):
                lookup[(num_id, ilvl)] = BULLET_OUTLINE_LEVEL
                continue

            outline_level = numbering_pattern_to_outline_level(num_fmt, lvl_text)
            if outline_level is None and 0 <= ilvl <= 8:
                outline_level = ilvl

            if outline_level is not None:
                lookup[(num_id, ilvl)] = outline_level

    return lookup


def build_style_numbering_lookup(styles_xml: bytes | None) -> dict[str, tuple[str, int]]:
    """
    Build a {styleId: (numId, ilvl)} lookup.

    Some Word documents attach numbering to paragraph styles instead of each
    paragraph's pPr/numPr. basedOn inheritance is resolved so styled paragraphs
    can still be recognized as document numbering levels.
    """
    if not styles_xml:
        return {}

    try:
        root = etree.fromstring(styles_xml)
    except Exception:
        return {}

    direct: dict[str, tuple[str | None, int | None, str | None]] = {}

    for style in root.xpath("./w:style[@w:type='paragraph']", namespaces=NS):
        style_id = style.get(qn("styleId"))
        if not style_id:
            continue

        based_on_el = style.find("w:basedOn", NS)
        based_on = based_on_el.get(qn("val")) if based_on_el is not None else None

        num_pr = style.find("./w:pPr/w:numPr", NS)
        num_id = None
        ilvl = None
        if num_pr is not None:
            num_id_el = num_pr.find("w:numId", NS)
            ilvl_el = num_pr.find("w:ilvl", NS)
            num_id = num_id_el.get(qn("val")) if num_id_el is not None else None
            try:
                ilvl = int(ilvl_el.get(qn("val"))) if ilvl_el is not None else 0
            except Exception:
                ilvl = None

        direct[style_id] = (num_id, ilvl, based_on)

    resolved: dict[str, tuple[str, int]] = {}
    resolving: set[str] = set()

    def resolve(style_id: str) -> tuple[str, int] | None:
        if style_id in resolved:
            return resolved[style_id]
        if style_id in resolving:
            return None

        item = direct.get(style_id)
        if item is None:
            return None

        resolving.add(style_id)
        num_id, ilvl, based_on = item
        result = None

        if num_id is not None and ilvl is not None:
            result = (num_id, ilvl)
        elif based_on:
            result = resolve(based_on)

        resolving.remove(style_id)

        if result is not None:
            resolved[style_id] = result
        return result

    for style_id in direct:
        resolve(style_id)

    return resolved


def apply_numbering_level_outline_format(lvl, level: int, change_logs: list[str] | None = None) -> None:
    """Normalize a numbering level to the configured outline geometry.

    All heading levels use w:suff="nothing". Tab stops are removed so Word
    cannot reopen the file and push the effective indent via a list tab.
    """
    from .indent_settings import twips_to_cm
    from .outline import apply_indent_spec_to_pPr

    spec = TEMPLATE_OUTLINE_INDENTS.get(level)
    if spec is None:
        return

    font_size_pt = OUTLINE_LEVEL_FONT_SIZE_PT.get(level)
    if font_size_pt is not None:
        rPr = get_or_add(lvl, "rPr")
        font_size = str(round(font_size_pt * 2))
        for tag in ("sz", "szCs"):
            size_el = get_or_add(rPr, tag)
            size_el.set(qn("val"), font_size)

    lvl_jc = get_or_add(lvl, "lvlJc")
    lvl_jc.set(qn("val"), "left")

    suff = lvl.find("w:suff", NS)
    if suff is None:
        suff = etree.Element(qn("suff"))
        lvl.append(suff)
    suffix = numbering_suffix_for_level(level)
    suff.set(qn("val"), suffix)

    pPr = get_or_add(lvl, "pPr")
    written = apply_indent_spec_to_pPr(pPr, spec, "heading_numbered", use_tab_stop=False)

    if change_logs is not None:
        number_start = int(spec["left"]) - int(spec["hanging"])
        tab_pos_cm = (
            f"{twips_to_cm(written['tab_pos']):.2f}"
            if written.get("tab_pos") is not None
            else "None"
        )
        change_logs.append(
            "NUMBERING_XML_LEVEL_INDENT: "
            f"level={level}; "
            f"expected_number_start_cm={twips_to_cm(number_start):.2f}; "
            f"expected_hanging_cm={twips_to_cm(spec['hanging']):.2f}; "
            f"expected_heading_left_cm={twips_to_cm(spec['left']):.2f}; "
            f"xml_written_left_cm={twips_to_cm(written.get('left') or spec['left']):.2f}; "
            f"xml_written_hanging_cm={twips_to_cm(written.get('hanging') or spec['hanging']):.2f}; "
            f"suff={suffix}; tab_pos_cm={tab_pos_cm}"
        )

def _calculated_number_start(left: str | None, hanging: str | None) -> str | None:
    if left is None or hanging is None:
        return None
    try:
        return str(int(left) - int(hanging))
    except (TypeError, ValueError):
        return None


def get_numbering_level_format(lvl) -> dict[str, str | None]:
    pPr = lvl.find("w:pPr", NS)
    ind = pPr.find("w:ind", NS) if pPr is not None else None
    tab = pPr.find("./w:tabs/w:tab", NS) if pPr is not None else None
    lvl_jc = lvl.find("w:lvlJc", NS)
    suff = lvl.find("w:suff", NS)
    num_fmt = lvl.find("w:numFmt", NS)
    lvl_text = lvl.find("w:lvlText", NS)

    left = ind.get(qn("left")) if ind is not None else None
    hanging = ind.get(qn("hanging")) if ind is not None else None
    return {
        "left": left,
        "hanging": hanging,
        "number_start": _calculated_number_start(left, hanging),
        "lvlJc": lvl_jc.get(qn("val")) if lvl_jc is not None else None,
        "suff": suff.get(qn("val")) if suff is not None else None,
        "tab_pos": tab.get(qn("pos")) if tab is not None else None,
        "numFmt": num_fmt.get(qn("val")) if num_fmt is not None else None,
        "lvlText": lvl_text.get(qn("val")) if lvl_text is not None else None,
    }


def build_numbering_format_lookup(numbering_xml: bytes | None) -> dict[tuple[str, int], dict[str, str | None]]:
    if not numbering_xml:
        return {}

    try:
        root = etree.fromstring(numbering_xml)
    except Exception:
        return {}

    abstract_formats: dict[str, dict[int, dict[str, str | None]]] = {}
    for abstract_num in root.xpath("./w:abstractNum", namespaces=NS):
        abstract_id = abstract_num.get(qn("abstractNumId"))
        if abstract_id is None:
            continue

        levels: dict[int, dict[str, str | None]] = {}
        for lvl in abstract_num.xpath("./w:lvl", namespaces=NS):
            try:
                ilvl = int(lvl.get(qn("ilvl")))
            except Exception:
                continue
            levels[ilvl] = get_numbering_level_format(lvl)
        abstract_formats[abstract_id] = levels

    lookup: dict[tuple[str, int], dict[str, str | None]] = {}
    for num in root.xpath("./w:num", namespaces=NS):
        num_id = num.get(qn("numId"))
        abstract_el = num.find("w:abstractNumId", NS)
        if num_id is None or abstract_el is None:
            continue

        abstract_id = abstract_el.get(qn("val"))
        for ilvl, fmt in abstract_formats.get(abstract_id, {}).items():
            lookup[(num_id, ilvl)] = dict(fmt)

        for override in num.xpath("./w:lvlOverride", namespaces=NS):
            try:
                ilvl = int(override.get(qn("ilvl")))
            except Exception:
                continue

            lvl = override.find("w:lvl", NS)
            if lvl is not None:
                lookup[(num_id, ilvl)] = get_numbering_level_format(lvl)

    return lookup


def apply_numbering_level_body_text_format(lvl) -> bool:
    pPr = lvl.find("w:pPr", NS)
    if pPr is None:
        return False

    outline_lvl = pPr.find("w:outlineLvl", NS)
    if outline_lvl is None:
        return False

    if outline_lvl.get(qn("val")) == "9":
        return False

    outline_lvl.set(qn("val"), "9")
    return True


def _legacy_apply_numbering_outline_format(
    numbering_xml: bytes | None,
    change_logs: list[str] | None = None,
    excluded_numbering_pairs: set[tuple[str, int]] | None = None,
    excluded_num_ids: set[str] | None = None,
    excluded_abstract_ids: set[str] | None = None,
) -> bytes | None:
    """
    Apply template indentation to recognizable automatic numbering formats.
    If parsing fails, return the original XML to keep document processing alive.
    """
    if not numbering_xml:
        return numbering_xml

    try:
        root = etree.fromstring(numbering_xml)
    except Exception:
        return numbering_xml

    changed = False

    # General abstract numbering definitions.
    for lvl in root.xpath("./w:abstractNum/w:lvl", namespaces=NS):
        changed = sanitize_numbering_level_suffix_tabs_and_text(lvl) or changed
        num_fmt_el = lvl.find("w:numFmt", NS)
        lvl_text_el = lvl.find("w:lvlText", NS)
        num_fmt = num_fmt_el.get(qn("val")) if num_fmt_el is not None else None
        lvl_text = lvl_text_el.get(qn("val")) if lvl_text_el is not None else None
        if is_bullet_num_fmt(num_fmt):
            changed = apply_numbering_level_body_text_format(lvl) or changed
            continue

        outline_level = numbering_pattern_to_outline_level(num_fmt, lvl_text)
        if outline_level is None:
            try:
                ilvl = int(lvl.get(qn("ilvl")))
            except Exception:
                ilvl = None
            if ilvl is not None and 0 <= ilvl <= 8:
                outline_level = ilvl

        if outline_level is not None:
            apply_numbering_level_outline_format(lvl, outline_level, change_logs=change_logs)
            changed = True

    # Per-num override definitions.
    for lvl in root.xpath("./w:num/w:lvlOverride/w:lvl", namespaces=NS):
        changed = sanitize_numbering_level_suffix_tabs_and_text(lvl) or changed
        num_fmt_el = lvl.find("w:numFmt", NS)
        lvl_text_el = lvl.find("w:lvlText", NS)
        num_fmt = num_fmt_el.get(qn("val")) if num_fmt_el is not None else None
        lvl_text = lvl_text_el.get(qn("val")) if lvl_text_el is not None else None
        if is_bullet_num_fmt(num_fmt):
            changed = apply_numbering_level_body_text_format(lvl) or changed
            continue

        outline_level = numbering_pattern_to_outline_level(num_fmt, lvl_text)
        if outline_level is None:
            try:
                ilvl = int(lvl.get(qn("ilvl")))
            except Exception:
                ilvl = None
            if ilvl is not None and 0 <= ilvl <= 8:
                outline_level = ilvl

        if outline_level is not None:
            apply_numbering_level_outline_format(lvl, outline_level, change_logs=change_logs)
            changed = True

    if not changed:
        return numbering_xml

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def apply_numbering_outline_format(
    numbering_xml: bytes | None,
    change_logs: list[str] | None = None,
    excluded_numbering_pairs: set[tuple[str, int]] | None = None,
    excluded_num_ids: set[str] | None = None,
    excluded_abstract_ids: set[str] | None = None,
) -> bytes | None:
    if not numbering_xml:
        return numbering_xml

    try:
        root = etree.fromstring(numbering_xml)
    except Exception:
        return numbering_xml

    changed = False
    excluded_numbering_pairs = excluded_numbering_pairs or set()
    excluded_num_ids = excluded_num_ids or set()
    excluded_abstract_ids = excluded_abstract_ids or set()

    num_to_abstract_id: dict[str, str] = {}
    for num in root.xpath("./w:num", namespaces=NS):
        num_id = num.get(qn("numId"))
        abstract_el = num.find("w:abstractNumId", NS)
        abstract_id = abstract_el.get(qn("val")) if abstract_el is not None else None
        if num_id is not None and abstract_id is not None:
            num_to_abstract_id[num_id] = abstract_id

    def should_skip_numbering(num_id: str | None, ilvl: int | None, abstract_id: str | None) -> bool:
        if abstract_id is not None and abstract_id in excluded_abstract_ids:
            return True
        if num_id is not None and num_id in excluded_num_ids:
            return True
        if num_id is not None and ilvl is not None and (num_id, ilvl) in excluded_numbering_pairs:
            return True
        return False

    def log_skip_numbering(num_id: str | None, ilvl: int | None, abstract_id: str | None) -> None:
        if change_logs is None:
            return
        change_logs.append(
            "NUMBERING_XML_SKIP_TOC_NUMBERING: "
            f"numId={num_id if num_id is not None else 'unknown'}; "
            f"ilvl={ilvl if ilvl is not None else 'unknown'}; "
            f"abstractNumId={abstract_id if abstract_id is not None else 'unknown'}; "
            "reason=used_by_toc"
        )

    for lvl in root.xpath("./w:abstractNum/w:lvl", namespaces=NS):
        abstract_num = lvl.getparent()
        abstract_id = abstract_num.get(qn("abstractNumId")) if abstract_num is not None else None
        try:
            ilvl = int(lvl.get(qn("ilvl")))
        except Exception:
            ilvl = None
        if should_skip_numbering(None, ilvl, abstract_id):
            log_skip_numbering(None, ilvl, abstract_id)
            continue

        changed = sanitize_numbering_level_suffix_tabs_and_text(lvl) or changed
        num_fmt_el = lvl.find("w:numFmt", NS)
        lvl_text_el = lvl.find("w:lvlText", NS)
        num_fmt = num_fmt_el.get(qn("val")) if num_fmt_el is not None else None
        lvl_text = lvl_text_el.get(qn("val")) if lvl_text_el is not None else None
        if is_bullet_num_fmt(num_fmt):
            changed = apply_numbering_level_body_text_format(lvl) or changed
            continue

        outline_level = numbering_pattern_to_outline_level(num_fmt, lvl_text)
        if outline_level is None and ilvl is not None and 0 <= ilvl <= 8:
            outline_level = ilvl

        if outline_level is not None:
            apply_numbering_level_outline_format(lvl, outline_level, change_logs=change_logs)
            changed = True

    for lvl in root.xpath("./w:num/w:lvlOverride/w:lvl", namespaces=NS):
        override = lvl.getparent()
        num = override.getparent() if override is not None else None
        num_id = num.get(qn("numId")) if num is not None else None
        abstract_id = num_to_abstract_id.get(num_id or "")
        try:
            ilvl = int(override.get(qn("ilvl"))) if override is not None else None
        except Exception:
            ilvl = None
        if should_skip_numbering(num_id, ilvl, abstract_id):
            log_skip_numbering(num_id, ilvl, abstract_id)
            continue

        changed = sanitize_numbering_level_suffix_tabs_and_text(lvl) or changed
        num_fmt_el = lvl.find("w:numFmt", NS)
        lvl_text_el = lvl.find("w:lvlText", NS)
        num_fmt = num_fmt_el.get(qn("val")) if num_fmt_el is not None else None
        lvl_text = lvl_text_el.get(qn("val")) if lvl_text_el is not None else None
        if is_bullet_num_fmt(num_fmt):
            changed = apply_numbering_level_body_text_format(lvl) or changed
            continue

        outline_level = numbering_pattern_to_outline_level(num_fmt, lvl_text)
        if outline_level is None and ilvl is not None and 0 <= ilvl <= 8:
            outline_level = ilvl

        if outline_level is not None:
            apply_numbering_level_outline_format(lvl, outline_level, change_logs=change_logs)
            changed = True

    if not changed:
        return numbering_xml

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def _style_direct_num_identity(style) -> tuple[str | None, int | None]:
    num_pr = style.find("./w:pPr/w:numPr", NS)
    if num_pr is None:
        return None, None

    num_id_el = num_pr.find("w:numId", NS)
    ilvl_el = num_pr.find("w:ilvl", NS)
    num_id = num_id_el.get(qn("val")) if num_id_el is not None else None
    try:
        ilvl = int(ilvl_el.get(qn("val"))) if ilvl_el is not None else 0
    except Exception:
        ilvl = None
    return num_id, ilvl


def clean_plain_style_pPr(pPr) -> None:
    from .outline import clear_indent_attrs, remove_paragraph_tabs

    ind = pPr.find("w:ind", NS)
    if ind is not None:
        clear_indent_attrs(ind)
        if not ind.attrib:
            pPr.remove(ind)
    remove_paragraph_tabs(pPr)


def style_name_value(style) -> str:
    name_el = style.find("w:name", NS)
    return name_el.get(qn("val")) if name_el is not None else ""


def is_toc_style_definition(style_id: str, style_name: str) -> bool:
    normalized_id = (style_id or "").replace(" ", "").replace("_", "").upper()
    normalized_name = (style_name or "").replace(" ", "").replace("_", "").upper()
    if normalized_id.startswith("TOC") or normalized_id in {"目錄", "目录"}:
        return True
    if normalized_name.startswith("TOC") or normalized_name in {"目錄", "目录"}:
        return True
    lowered_name = (style_name or "").lower()
    return (
        "table of contents" in lowered_name
        or "contents" in lowered_name
        or "目錄" in style_name
        or "目录" in style_name
    )


def apply_styles_outline_format_to_root(
    root,
    numbering_level_lookup=None,
    style_numbering_lookup=None,
    change_logs: list[str] | None = None,
    excluded_style_ids: set[str] | None = None,
) -> bool:
    from .indent_settings import twips_to_cm
    from .outline import apply_indent_spec_to_pPr

    style_numbering_lookup = style_numbering_lookup or {}
    excluded_style_ids = excluded_style_ids or set()
    changed = False

    for style in root.xpath("./w:style[@w:type='paragraph']", namespaces=NS):
        style_id = style.get(qn("styleId")) or ""
        style_name = style_name_value(style)
        if style_id in excluded_style_ids:
            if change_logs is not None:
                change_logs.append(
                    "STYLES_XML_SKIP_EXCLUDED_STYLE: "
                    f"styleId={style_id}; name={style_name}; reason=used_by_chapter_參"
                )
            continue
        if is_toc_style_definition(style_id, style_name):
            if change_logs is not None:
                change_logs.append(
                    "STYLES_XML_SKIP_TOC_STYLE: "
                    f"styleId={style_id}; name={style_name}; reason=toc_style"
                )
            continue
        pPr = style.find("w:pPr", NS)
        direct_num_id, direct_ilvl = _style_direct_num_identity(style)
        num_id, ilvl = style_numbering_lookup.get(style_id, (direct_num_id, direct_ilvl))
        level = detect_number_level_from_identity(num_id, ilvl, numbering_level_lookup)

        if level == BULLET_OUTLINE_LEVEL:
            if pPr is not None:
                clean_plain_style_pPr(pPr)
                changed = True
            continue

        if level is not None and 0 <= level <= 8:
            spec = TEMPLATE_OUTLINE_INDENTS.get(level)
            if spec is None:
                continue
            if pPr is None:
                pPr = get_or_add(style, "pPr")
            suffix = numbering_suffix_for_level(level)
            written = apply_indent_spec_to_pPr(pPr, spec, "heading_numbered", use_tab_stop=False)
            changed = True
            if change_logs is not None:
                tab_pos_cm = (
                    f"{twips_to_cm(written['tab_pos']):.2f}"
                    if written.get("tab_pos") is not None
                    else "None"
                )
                change_logs.append(
                    "STYLES_XML_NUMBERED_STYLE_INDENT: "
                    f"styleId={style_id}; kind=auto(style); level={level}; "
                    f"expected_number_start_cm={twips_to_cm(spec.get('number_start', int(spec['left']) - int(spec['hanging']))):.2f}; "
                    f"expected_hanging_cm={twips_to_cm(spec['hanging']):.2f}; "
                    f"expected_heading_left_cm={twips_to_cm(spec['left']):.2f}; "
                    f"xml_written_left_cm={twips_to_cm(written.get('left') or spec['left']):.2f}; "
                    f"xml_written_hanging_cm={twips_to_cm(written.get('hanging') or spec['hanging']):.2f}; "
                    f"suff={suffix}; tab_pos_cm={tab_pos_cm}"
                )
            continue

        if pPr is not None:
            clean_plain_style_pPr(pPr)
            changed = True

    return changed


def apply_styles_outline_format(
    styles_xml: bytes | None,
    numbering_level_lookup=None,
    change_logs: list[str] | None = None,
) -> bytes | None:
    if not styles_xml:
        return styles_xml

    try:
        root = etree.fromstring(styles_xml)
    except Exception:
        return styles_xml

    style_numbering_lookup = build_style_numbering_lookup(styles_xml)
    changed = apply_styles_outline_format_to_root(
        root,
        numbering_level_lookup=numbering_level_lookup,
        style_numbering_lookup=style_numbering_lookup,
        change_logs=change_logs,
    )
    if not changed:
        return styles_xml

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def detect_number_level_from_identity(num_id, ilvl, numbering_level_lookup=None):
    if num_id is None or ilvl is None:
        return None

    if numbering_level_lookup is not None:
        outline_level = numbering_level_lookup.get((num_id, ilvl))
        if outline_level == BULLET_OUTLINE_LEVEL:
            return BULLET_OUTLINE_LEVEL
        if outline_level is not None:
            return outline_level
        if numbering_level_lookup:
            return None

    if 0 <= ilvl <= 8:
        return ilvl

    return None


def detect_auto_number_level(p, numbering_level_lookup=None, style_numbering_lookup=None):
    """
    Resolve Word automatic numbering.

    Word automatic numbering markers are not stored in w:t text, so the actual
    numbering format must be looked up from numbering.xml by numId + ilvl.
    """
    ilvl_el = p.find("./w:pPr/w:numPr/w:ilvl", NS)
    num_id_el = p.find("./w:pPr/w:numPr/w:numId", NS)

    if num_id_el is None:
        return None

    num_id = num_id_el.get(qn("val"))
    if num_id is None:
        return None

    try:
        ilvl = int(ilvl_el.get(qn("val"))) if ilvl_el is not None else 0
    except Exception:
        return None

    direct_level = detect_number_level_from_identity(num_id, ilvl, numbering_level_lookup)
    if direct_level == BULLET_OUTLINE_LEVEL:
        return None
    if direct_level is not None:
        return direct_level

    style_id = paragraph_style_id(p)
    if style_id and style_numbering_lookup:
        style_num_id, style_ilvl = style_numbering_lookup.get(style_id, (None, None))
        style_level = detect_number_level_from_identity(
            style_num_id,
            style_ilvl,
            numbering_level_lookup,
        )
        if style_level == BULLET_OUTLINE_LEVEL:
            return None
        if style_level is not None:
            return style_level

    return None

def detect_style_number_level(p, numbering_level_lookup=None, style_numbering_lookup=None):
    style_id = paragraph_style_id(p)
    if not style_id or not style_numbering_lookup:
        return None

    num_id, ilvl = style_numbering_lookup.get(style_id, (None, None))
    level = detect_number_level_from_identity(num_id, ilvl, numbering_level_lookup)
    if level == BULLET_OUTLINE_LEVEL:
        return None
    return level
