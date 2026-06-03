from __future__ import annotations

from lxml import etree

from .constants import NS, OUTLINE_LEVEL_FONT_SIZE_PT, TEMPLATE_OUTLINE_INDENTS
from .xml_utils import get_or_add, qn

BULLET_OUTLINE_LEVEL = -1
PAREN_PAIRS = {
    "(": ")",
    "（": "）",
}


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
    將 Word numbering.xml 的 numFmt + lvlText 轉成範本階層。

    範本中的自動編號多為「每一種符號一組 numId」，所以 ilvl 幾乎都是 0。
    因此不能只看 ilvl，必須看實際編號格式：
    一、／（一）／(一)／1.／（1）／(1)／A.／（A）／(A)／a.／（a）／(a)。
    """
    fmt = (num_fmt or "").strip()
    text = normalize_lvl_text(lvl_text)

    bracketed = is_parenthesized_lvl_text(text)
    has_dot = text.endswith((".", "．")) or "." in text or "．" in text
    has_ideographic_separator = "、" in text

    # 壹、貳、參...
    if fmt in {"ideographLegalTraditional", "chineseLegalSimplified"}:
        if has_ideographic_separator:
            return 0

    # 一、二、三... 或 （一）（二）（三）... / (一)(二)(三)...
    if fmt in {"taiwaneseCountingThousand", "ideographTraditional", "chineseCounting"}:
        if bracketed:
            return 2
        if has_ideographic_separator:
            return 1

    # 1. 或 （1）/ (1)
    if fmt == "decimal":
        if bracketed:
            return 4
        if has_dot or has_ideographic_separator:
            return 3

    # A. 或 （A）/ (A)
    if fmt == "upperLetter":
        if bracketed:
            return 6
        if has_dot or has_ideographic_separator:
            return 5

    # a. 或 （a）/ (a)
    if fmt == "lowerLetter":
        if bracketed:
            return 8
        if has_dot or has_ideographic_separator:
            return 7

    return None


def build_numbering_level_lookup(numbering_xml: bytes | None):
    """
    建立 {(numId, ilvl): outline_level} 對照表，讓自動編號可套用範本階層。
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

            # 若 numbering.xml 的格式無法直接判斷，但它是 Word 自動編號層級，
            # 就把 ilvl 本身視為大綱階層。
            # 這符合「文件中的每個自動編號都是大綱階層」的前提。
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

        # 若有個別 num 的 lvlOverride，也覆蓋進 lookup。
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
    建立 {styleId: (numId, ilvl)} 對照表。

    有些 Word 文件會把編號掛在段落樣式上，而不是每個段落的 pPr/numPr。
    這裡同時處理 basedOn 繼承，讓套用樣式的段落也能被判斷成文件編號階層。
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


def apply_numbering_level_outline_format(lvl, level: int) -> None:
    """
    同步修改 numbering.xml 裡的自動編號層級格式。

    只改本工具可辨識的 0~8 階編號。重點是：
    1. 套用與範本一致的 left/hanging。
    2. 清掉舊 tab stop，避免編號後面出現過長留白。
    3. 將自動編號後綴由 tab/space 改成 nothing，避免灰底／反白延伸出額外色塊。
    """
    from .outline import clear_indent_attrs, normalize_tabs_to_text_position

    spec = TEMPLATE_OUTLINE_INDENTS.get(level)
    if spec is None:
        return

    # 自動編號後面不要用 tab，也不要用 space。
    # tab 會形成一大片灰色留白；space 會形成一小塊灰色方塊。
    # 改成 nothing，讓灰底只包住編號本身。
    font_size_pt = OUTLINE_LEVEL_FONT_SIZE_PT.get(level)
    if font_size_pt is not None:
        rPr = get_or_add(lvl, "rPr")
        font_size = str(round(font_size_pt * 2))
        for tag in ("sz", "szCs"):
            size_el = get_or_add(rPr, tag)
            size_el.set(qn("val"), font_size)

    suff = lvl.find("w:suff", NS)
    if suff is None:
        suff = etree.Element(qn("suff"))
        lvl.append(suff)
    suff.set(qn("val"), "nothing")

    pPr = get_or_add(lvl, "pPr")
    ind = get_or_add(pPr, "ind")
    clear_indent_attrs(ind)
    ind.set(qn("left"), spec["left"])
    ind.set(qn("hanging"), spec["hanging"])

    normalize_tabs_to_text_position(pPr, spec["left"])


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


def apply_numbering_outline_format(numbering_xml: bytes | None) -> bytes | None:
    """
    將 numbering.xml 裡可辨識的自動編號格式同步套用範本縮排。
    若解析失敗，回傳原始 XML，避免中斷整份文件處理。
    """
    if not numbering_xml:
        return numbering_xml

    try:
        root = etree.fromstring(numbering_xml)
    except Exception:
        return numbering_xml

    changed = False

    # 一般 abstract numbering 定義。
    for lvl in root.xpath("./w:abstractNum/w:lvl", namespaces=NS):
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
            apply_numbering_level_outline_format(lvl, outline_level)
            changed = True

    # 個別 num 的 override 定義。
    for lvl in root.xpath("./w:num/w:lvlOverride/w:lvl", namespaces=NS):
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
            apply_numbering_level_outline_format(lvl, outline_level)
            changed = True

    if not changed:
        return numbering_xml

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
    處理 Word 自動編號。

    Word 自動編號的「壹、」「一、」「（一）」「1.」等不在 w:t 文字裡，
    因此必須根據 numId + ilvl 回查 numbering.xml 的實際編號格式。
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
