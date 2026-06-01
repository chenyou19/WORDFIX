from __future__ import annotations

import ctypes
import ctypes.wintypes
import os
import re
import unicodedata

from lxml import etree

from .constants import FINANCIAL_NUM, NS, PREFACE_OUTLINE_INDENTS, SIMPLE_NUM, TEMPLATE_OUTLINE_INDENTS
from .indent_settings import twips_to_cm
from .numbering import (
    detect_auto_number_level,
    detect_style_number_level,
    has_auto_numbering,
    paragraph_style_id,
)
from .stop_controller import StopController
from .xml_utils import get_or_add, paragraph_text, qn

NOTE_MARKER_PREFIXES = ("※",)
STYLE_NUMBERING_MAX_TEXT_LENGTH = 35
HEADING_ENDINGS = ("：", ":")
PROCESSING_START_TITLE = "序言"
PROCESSING_START_VISIBLE_PREFIX = "壹、序言"
DEFAULT_NUMBER_FONT = "Microsoft JhengHei"
DEFAULT_NUMBER_FONT_SIZE_PT = 12.0


def starts_with_note_marker(text: str) -> bool:
    return (text or "").lstrip().startswith(NOTE_MARKER_PREFIXES)


def normalize_visible_text(text: str) -> str:
    return " ".join((text or "").split())


def should_skip_style_numbering(text: str) -> bool:
    """
    避免只因段落樣式帶編號，就把正文誤判成文件大綱。

    樣式編號主要用於短標題或短項目；長篇敘述文字不使用樣式編號回推階層。
    """
    normalized = normalize_visible_text(text)
    if starts_with_note_marker(normalized):
        return True
    return len(normalized) > STYLE_NUMBERING_MAX_TEXT_LENGTH and not normalized.endswith(HEADING_ENDINGS)


def compact_text(text: str) -> str:
    return "".join((text or "").split())


def is_processing_start_marker(
    p,
    text: str,
    numbering_level_lookup=None,
    style_numbering_lookup=None,
) -> bool:
    compact = compact_text(text)
    if compact.startswith(PROCESSING_START_VISIBLE_PREFIX):
        return True

    if not compact.startswith(PROCESSING_START_TITLE):
        return False

    level = None
    if has_auto_numbering(p):
        level = detect_auto_number_level(
            p,
            numbering_level_lookup=numbering_level_lookup,
            style_numbering_lookup=style_numbering_lookup,
        )

    if level is None:
        level = detect_style_number_level(
            p,
            numbering_level_lookup=numbering_level_lookup,
            style_numbering_lookup=style_numbering_lookup,
        )

    if level is None:
        level = detect_outline_level(text)

    return level == 0


def effective_indent_level(level: int, set_outline: bool) -> int:
    """
    在「壹、序言」前，前置段落不會出現壹、階層，因此一、視為階層 0。
    從「壹、序言」開始，使用正式文件階層。
    """
    if set_outline:
        return level
    return max(level - 1, 0)


def is_separator_char(ch: str) -> bool:
    return ch in {
        "、", ".", "．", "。", ":", "：", ";", "；",
        " ", "\t", "　",
    }


def is_strict_separator_char(ch: str) -> bool:
    """英文字母層級用較嚴格分隔符，避免 A 公司 / a test 誤判。"""
    return ch in {"、", ".", "．", "。", ":", "：", ";", "；"}


def is_end_or_separator(text: str, index: int) -> bool:
    if index >= len(text):
        return True
    return is_separator_char(text[index])


def is_end_or_strict_separator(text: str, index: int) -> bool:
    if index >= len(text):
        return True
    return is_strict_separator_char(text[index])


def is_ascii_letter_or_digit(ch: str) -> bool:
    return bool(re.match(r"[A-Za-z0-9]", ch))


def match_parenthesized_numbering(text: str):
    """
    括號型文件編號：
    層級 2：（一）或 (一)
    層級 4：（1）或 (1)
    層級 6：（A）或 (A)
    層級 8：（a）或 (a)

    同時支援半形括號與多位數，例如 (10)。
    """
    patterns = [
        (rf"^[（(][{SIMPLE_NUM}]+[)）]", 2),
        (r"^[（(][0-9]{1,3}[)）]", 4),
        (r"^[（(][A-Z][)）]", 6),
        (r"^[（(][a-z][)）]", 8),
    ]

    for pattern, level in patterns:
        m = re.match(pattern, text)
        if not m:
            continue

        end_index = m.end()

        # 避免 (A)pple、(a)pple、(1)234 這類誤判。
        if end_index < len(text):
            next_char = text[end_index]
            if level in {4, 6, 8} and is_ascii_letter_or_digit(next_char):
                return None

        return level

    return None


def match_plain_numbering(text: str):
    """
    非括號型文件編號：
    層級 0：壹、貳、參...
    層級 1：一、二、三...
    層級 3：1.
    層級 5：A.
    層級 7：a.

    必須帶指定標點，避免 1、一、A、a 這類單獨文字被誤判。
    """
    checks = [
        (rf"^[{FINANCIAL_NUM}]+、", 0),
        (rf"^[{SIMPLE_NUM}]+、", 1),
        (r"^[0-9]{1,3}\.", 3),
        (r"^[A-Z]\.", 5),
        (r"^[a-z]\.", 7),
    ]

    for pattern, level in checks:
        m = re.match(pattern, text)
        if not m:
            continue
        if m.end() < len(text) and is_ascii_letter_or_digit(text[m.end()]):
            return None
        return level

    return None


def detect_manual_numbering_prefix(text: str) -> tuple[int, str] | None:
    if not text:
        return None

    text = text.lstrip()
    if not text:
        return None

    patterns = [
        (rf"^[（(][{SIMPLE_NUM}]+[)）]", 2),
        (r"^[（(][0-9]{1,3}[)）]", 4),
        (r"^[（(][A-Z][)）]", 6),
        (r"^[（(][a-z][)）]", 8),
        (rf"^[{FINANCIAL_NUM}]+、", 0),
        (rf"^[{SIMPLE_NUM}]+、", 1),
        (r"^[0-9]{1,3}\.", 3),
        (r"^[A-Z]\.", 5),
        (r"^[a-z]\.", 7),
    ]

    for pattern, level in patterns:
        m = re.match(pattern, text)
        if not m:
            continue
        if m.end() < len(text) and level in {3, 4, 5, 6, 7, 8} and is_ascii_letter_or_digit(text[m.end()]):
            return None
        return level, m.group(0)

    return None


def detect_outline_level(text: str):
    """
    只有段落開頭符合文件編號格式時才回傳層級。
    編號前只允許空白；若前面有其他文字，就不視為文件編號。

    層級 0：壹、
    層級 1：一、
    層級 2：（一）或 (一)
    層級 3：1.
    層級 4：（1）或 (1)
    層級 5：A.
    層級 6：（A）或 (A)
    層級 7：a.
    層級 8：（a）或 (a)
    """
    if not text:
        return None

    text = text.lstrip()
    if not text:
        return None

    level = match_parenthesized_numbering(text)
    if level is not None:
        return level

    level = match_plain_numbering(text)
    if level is not None:
        return level

    return None


def clear_indent_attrs(ind) -> None:
    """清掉會互相干擾的縮排屬性，避免 left/start 或 firstLine/hanging 並存。"""
    for attr in [
        "left", "start", "right", "end",
        "leftChars", "startChars", "rightChars", "endChars",
        "firstLine", "hanging",
        "firstLineChars", "hangingChars",
    ]:
        ind.attrib.pop(qn(attr), None)


def normalize_tabs_to_text_position(pPr, text_position_twips: str) -> None:
    """
    將段落／編號層級的 tab stop 統一到文字起點。

    Word 自動編號常會在編號後面放一個 tab；如果原本 tab stop 很遠，
    就會出現「（一）」後方一大片留白。這裡先清掉舊 tabs，再補一個
    與 left indent 相同位置的 tab stop，避免文字被舊 tab 推太遠。
    """
    tabs = pPr.find("w:tabs", NS)
    if tabs is not None:
        pPr.remove(tabs)

    tabs = etree.Element(qn("tabs"))
    tab = etree.Element(qn("tab"))
    tab.set(qn("val"), "left")
    tab.set(qn("pos"), text_position_twips)
    tabs.append(tab)
    pPr.append(tabs)


def apply_paragraph_outline_level(pPr, level: int) -> None:
    """設定 Word 段落屬性中的「大綱階層」。

    Word 的「大綱階層 1~9」在 XML 裡是 w:outlineLvl，值為 0~8。
    這不是只有縮排外觀，而是讓段落真的被 Word 視為對應的大綱階層，
    可用於導覽窗格、目錄、大綱檢視等功能。
    """
    if not (0 <= level <= 8):
        return

    outline_lvl = get_or_add(pPr, "outlineLvl")
    outline_lvl.set(qn("val"), str(level))


def apply_paragraph_body_text_level(p) -> bool:
    outline_lvl = p.find("./w:pPr/w:outlineLvl", NS)
    if outline_lvl is None:
        return False

    if outline_lvl.get(qn("val")) == "9":
        return False

    outline_lvl.set(qn("val"), "9")
    return True


def remove_paragraph_outline_level(p) -> bool:
    pPr = p.find("./w:pPr", NS)
    if pPr is None:
        return False

    outline_lvl = pPr.find("w:outlineLvl", NS)
    if outline_lvl is None:
        return False

    pPr.remove(outline_lvl)
    return True


def get_paragraph_outline_level_value(p) -> str:
    outline_lvl = p.find("./w:pPr/w:outlineLvl", NS)
    if outline_lvl is None:
        return "無"
    return outline_lvl.get(qn("val")) or "無"


def get_auto_number_identity(p) -> tuple[str | None, int | None]:
    ilvl_el = p.find("./w:pPr/w:numPr/w:ilvl", NS)
    num_id_el = p.find("./w:pPr/w:numPr/w:numId", NS)

    num_id = num_id_el.get(qn("val")) if num_id_el is not None else None

    try:
        ilvl = int(ilvl_el.get(qn("val"))) if ilvl_el is not None else 0
    except Exception:
        ilvl = None

    return num_id, ilvl


def get_number_prefix_run_properties(p) -> tuple[str, float]:
    for run in p.findall("./w:r", NS):
        text = "".join(run.xpath(".//w:t/text()", namespaces=NS))
        if not text or not text.strip():
            continue

        rPr = run.find("w:rPr", NS)
        font_name = DEFAULT_NUMBER_FONT
        font_size_pt = DEFAULT_NUMBER_FONT_SIZE_PT

        if rPr is not None:
            r_fonts = rPr.find("w:rFonts", NS)
            if r_fonts is not None:
                font_name = (
                    r_fonts.get(qn("eastAsia"))
                    or r_fonts.get(qn("ascii"))
                    or r_fonts.get(qn("hAnsi"))
                    or font_name
                )

            size = rPr.find("w:sz", NS)
            if size is not None:
                try:
                    font_size_pt = int(size.get(qn("val")) or "24") / 2
                except ValueError:
                    pass

        return font_name, font_size_pt

    return DEFAULT_NUMBER_FONT, DEFAULT_NUMBER_FONT_SIZE_PT


def estimate_text_width_cm(text: str, font_size_pt: float) -> float:
    units = 0.0
    for ch in text:
        if ch.isspace():
            units += 0.35
            continue
        width = unicodedata.east_asian_width(ch)
        units += 1.0 if width in {"F", "W"} else 0.55
    return units * font_size_pt / 72 * 2.54


def measure_text_width_cm(text: str, font_name: str, font_size_pt: float) -> float:
    if os.name != "nt":
        return estimate_text_width_cm(text, font_size_pt)

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)

    user32.GetDC.argtypes = [ctypes.wintypes.HWND]
    user32.GetDC.restype = ctypes.wintypes.HDC
    user32.ReleaseDC.argtypes = [ctypes.wintypes.HWND, ctypes.wintypes.HDC]
    user32.ReleaseDC.restype = ctypes.c_int
    gdi32.GetDeviceCaps.argtypes = [ctypes.wintypes.HDC, ctypes.c_int]
    gdi32.GetDeviceCaps.restype = ctypes.c_int
    gdi32.CreateFontW.restype = ctypes.wintypes.HFONT
    gdi32.SelectObject.argtypes = [ctypes.wintypes.HDC, ctypes.wintypes.HGDIOBJ]
    gdi32.SelectObject.restype = ctypes.wintypes.HGDIOBJ
    gdi32.DeleteObject.argtypes = [ctypes.wintypes.HGDIOBJ]
    gdi32.DeleteObject.restype = ctypes.c_int
    gdi32.GetTextExtentPoint32W.argtypes = [
        ctypes.wintypes.HDC,
        ctypes.c_wchar_p,
        ctypes.c_int,
        ctypes.POINTER(ctypes.wintypes.SIZE),
    ]
    gdi32.GetTextExtentPoint32W.restype = ctypes.c_int

    hdc = user32.GetDC(None)
    if not hdc:
        return estimate_text_width_cm(text, font_size_pt)

    try:
        log_pixels_x = gdi32.GetDeviceCaps(hdc, 88)
        log_pixels_y = gdi32.GetDeviceCaps(hdc, 90)
        font_height = -round(font_size_pt * log_pixels_y / 72)
        font = gdi32.CreateFontW(
            font_height,
            0,
            0,
            0,
            400,
            0,
            0,
            0,
            1,
            0,
            0,
            0,
            0,
            font_name,
        )
        if not font:
            return estimate_text_width_cm(text, font_size_pt)

        old_font = gdi32.SelectObject(hdc, font)
        try:
            size = ctypes.wintypes.SIZE()
            ok = gdi32.GetTextExtentPoint32W(hdc, text, len(text), ctypes.byref(size))
            if not ok:
                return estimate_text_width_cm(text, font_size_pt)
            return size.cx / log_pixels_x * 2.54
        finally:
            gdi32.SelectObject(hdc, old_font)
            gdi32.DeleteObject(font)
    finally:
        user32.ReleaseDC(None, hdc)


def record_numbering_measurement(
    summary,
    *,
    text: str,
    p,
    level: int,
    indent_level: int,
    set_outline: bool,
) -> None:
    if summary is None:
        return

    manual = detect_manual_numbering_prefix(text)
    if manual is None:
        return

    _, prefix = manual
    spec = get_outline_indent_spec(indent_level, set_outline=set_outline)
    if spec is None:
        return

    font_name, font_size_pt = get_number_prefix_run_properties(p)
    number_size_cm = measure_text_width_cm(prefix, font_name, font_size_pt)
    left_cm = twips_to_cm(spec["left"])
    number_start_cm = twips_to_cm(spec.get("number_start", int(spec["left"]) - int(spec["hanging"])))
    section = "壹、序言後" if set_outline else "壹、序言前"
    key = f"{section}:{level}:{prefix}"

    current = summary.numbering_measurements.get(key)
    measurement = {
        "section": section,
        "level": level,
        "indent_level": indent_level,
        "prefix": prefix,
        "text_start_cm": left_cm,
        "number_start_cm": number_start_cm,
        "number_size_cm": number_size_cm,
        "font_name": font_name,
        "font_size_pt": font_size_pt,
        "count": 1,
    }
    if current is None:
        summary.numbering_measurements[key] = measurement
    else:
        current["count"] = int(current.get("count", 1)) + 1
        current["number_size_cm"] = max(float(current.get("number_size_cm", 0)), number_size_cm)


def summarize_paragraph_text(text: str, limit: int = 80) -> str:
    normalized = " ".join((text or "").split())
    if len(normalized) <= limit:
        return normalized
    return normalized[:limit - 1] + "…"


def append_paragraph_change_log(
    change_logs: list[str] | None,
    part_name: str,
    paragraph_index: int,
    text: str,
    before_outline: str,
    after_outline: str,
    reason: str,
) -> None:
    if change_logs is None:
        return

    preview = summarize_paragraph_text(text) or "（無文字）"
    change_logs.append(
        f"[{part_name} #{paragraph_index}] {reason}；"
        f"大綱階層：{before_outline} -> {after_outline}；"
        f"段落文字：{preview}"
    )


def get_outline_indent_spec(level: int, set_outline: bool = True) -> dict[str, str] | None:
    indent_specs = TEMPLATE_OUTLINE_INDENTS if set_outline else PREFACE_OUTLINE_INDENTS
    spec = indent_specs.get(level)
    if spec is None and not set_outline:
        spec = TEMPLATE_OUTLINE_INDENTS.get(level)
    return spec


def apply_outline_indent(p, level: int, set_outline: bool = True) -> None:
    """依範本.docx的階層縮排標準套用段落縮排，可選擇是否同步設定 Word 大綱階層。"""
    spec = get_outline_indent_spec(level, set_outline=set_outline)
    if spec is None:
        return

    pPr = get_or_add(p, "pPr", first=True)

    if set_outline:
        # level 0~8 對應 Word UI 的「大綱階層 1~9」。
        apply_paragraph_outline_level(pPr, level)
    else:
        remove_paragraph_outline_level(p)

    ind = get_or_add(pPr, "ind")

    clear_indent_attrs(ind)

    ind.set(qn("left"), spec["left"])
    ind.set(qn("hanging"), spec["hanging"])

    # 修正「編號後面被舊 tab stop 推出奇怪留白」的問題。
    normalize_tabs_to_text_position(pPr, spec["left"])


def apply_body_indent_from_heading(p, heading_level: int, heading_uses_outline: bool) -> bool:
    """讓標題下方的普通內文左縮排對齊該標題的文字起點。"""
    spec = get_outline_indent_spec(heading_level, set_outline=heading_uses_outline)
    if spec is None:
        return False

    pPr = get_or_add(p, "pPr", first=True)
    ind = get_or_add(pPr, "ind")
    clear_indent_attrs(ind)
    ind.set(qn("left"), spec["left"])
    return True


def get_paragraph_style_value(p) -> str:
    """取得段落樣式值，例如 TOC1、TOC2、TOCHeading。"""
    style_el = p.find("./w:pPr/w:pStyle", NS)
    if style_el is None:
        return ""
    return (style_el.get(qn("val")) or "").strip()


def is_toc_style_value(style_value: str) -> bool:
    """判斷段落樣式是否屬於 Word 目錄樣式。"""
    if not style_value:
        return False

    normalized = style_value.replace(" ", "").replace("_", "").upper()

    # Word 目錄項目通常是 TOC1~TOC9，目錄標題常是 TOCHeading。
    if re.fullmatch(r"TOC\d+", normalized):
        return True
    if normalized in {"TOCHEADING", "目錄", "目录"}:
        return True
    if normalized.startswith(("TOC", "目錄", "目录")):
        return True

    return False


def field_instr_is_toc(instr: str) -> bool:
    """判斷 field instruction 是否為 TOC 欄位。"""
    if not instr:
        return False
    return re.search(r"(^|\s)TOC(\s|$|\\)", instr.upper()) is not None


def collect_toc_paragraph_ids(root) -> set[int]:
    """
    收集目錄段落，避免目錄本身被加上大綱階層。

    Word 目錄通常有兩種訊號：
    1. 段落樣式為 TOC1~TOC9 / TOCHeading。
    2. 由 TOC field 產生，內容位於 TOC field 的 begin/separate/end 範圍中。

    這裡同時支援兩種方式；若是 TOC 裡面的 PAGEREF 子欄位，也不會誤把
    TOC field 提早結束。
    """
    toc_ids: set[int] = set()
    field_stack: list[dict[str, object]] = []

    for p in root.xpath(".//w:p", namespaces=NS):
        p_is_toc = is_toc_style_value(get_paragraph_style_value(p))

        # 若目前已在 TOC field 內，這個段落也要跳過。
        if any(frame.get("is_toc") for frame in field_stack):
            p_is_toc = True

        # fldSimple 的簡單欄位形式。
        for fld_simple in p.xpath("ancestor-or-self::w:fldSimple", namespaces=NS):
            if field_instr_is_toc(fld_simple.get(qn("instr")) or ""):
                p_is_toc = True
                break

        # 依文件順序掃描複雜欄位：begin / instrText / separate / end。
        for el in p.iter():
            if el.tag == qn("fldChar"):
                fld_type = el.get(qn("fldCharType"))

                if fld_type == "begin":
                    field_stack.append({"instr": "", "is_toc": False})

                elif fld_type == "separate":
                    if field_stack:
                        instr = str(field_stack[-1].get("instr", ""))
                        if field_instr_is_toc(instr):
                            field_stack[-1]["is_toc"] = True
                            p_is_toc = True

                elif fld_type == "end":
                    if field_stack:
                        ended = field_stack.pop()
                        if ended.get("is_toc"):
                            p_is_toc = True

            elif el.tag == qn("instrText"):
                if field_stack:
                    field_stack[-1]["instr"] = str(field_stack[-1].get("instr", "")) + (el.text or "")
                    if field_instr_is_toc(str(field_stack[-1].get("instr", ""))):
                        # TOC 指令所在段落本身也不要加大綱階層。
                        p_is_toc = True

        if any(frame.get("is_toc") for frame in field_stack):
            p_is_toc = True

        if p_is_toc:
            toc_ids.add(id(p))

    return toc_ids


def is_plain_toc_heading(text: str) -> bool:
    return compact_text(text) in {"目錄", "目录", "目次"}


def collect_plain_toc_range_paragraph_ids(
    paragraphs,
    numbering_level_lookup=None,
    style_numbering_lookup=None,
) -> set[int]:
    toc_ids: set[int] = set()
    in_toc_range = False

    for p in paragraphs:
        text = paragraph_text(p)

        if in_toc_range and is_processing_start_marker(
            p,
            text,
            numbering_level_lookup=numbering_level_lookup,
            style_numbering_lookup=style_numbering_lookup,
        ):
            in_toc_range = False

        if in_toc_range:
            toc_ids.add(id(p))

        if not in_toc_range and is_plain_toc_heading(text):
            toc_ids.add(id(p))
            in_toc_range = True

    return toc_ids


def is_table_paragraph(p) -> bool:
    return bool(p.xpath("ancestor::w:tbl", namespaces=NS))


def increment_paragraph_level_count(summary, level: int) -> None:
    if summary is None:
        return
    if 0 <= level <= 8:
        summary.paragraph_level_counts[level] += 1


def increment_removed_preface_outline_count(summary) -> None:
    if summary is not None:
        summary.removed_preface_outline_paragraphs += 1


def fix_outline_paragraphs(
    root,
    include_tables: bool,
    stop: StopController | None = None,
    numbering_level_lookup=None,
    style_numbering_lookup=None,
    change_logs: list[str] | None = None,
    part_name: str = "word/document.xml",
    summary=None,
    remove_preface_outline: bool = False,
    fix_numbered_paragraphs: bool = True,
) -> int:
    paragraphs = root.xpath(".//w:p", namespaces=NS)
    if summary is not None:
        summary.total_paragraphs += len(paragraphs)

    # 目錄本身不要加大綱階層，避免 TOC 項目被 Word 當成正式章節。
    toc_paragraph_ids = collect_toc_paragraph_ids(root)
    toc_paragraph_ids.update(
        collect_plain_toc_range_paragraph_ids(
            paragraphs,
            numbering_level_lookup=numbering_level_lookup,
            style_numbering_lookup=style_numbering_lookup,
        )
    )

    changed_count = 0
    outline_processing_started = not remove_preface_outline
    current_heading_indent: tuple[int, bool] | None = None

    for paragraph_index, p in enumerate(paragraphs, start=1):
        if stop:
            stop.check()

        try:
            if is_table_paragraph(p):
                if summary is not None:
                    summary.skipped_table_paragraphs += 1
                continue

            text = paragraph_text(p)
            if not text or not text.strip():
                continue

            if id(p) in toc_paragraph_ids:
                if not outline_processing_started and remove_preface_outline:
                    before_outline = get_paragraph_outline_level_value(p)
                    if remove_paragraph_outline_level(p):
                        changed_count += 1
                        increment_removed_preface_outline_count(summary)
                        append_paragraph_change_log(
                            change_logs,
                            part_name,
                            paragraph_index,
                            text,
                            before_outline,
                            "無",
                            "壹、序言前目錄段落，移除既有 Word 大綱階層",
                        )

                if summary is not None:
                    summary.skipped_toc_paragraphs += 1
                continue

            before_outline = get_paragraph_outline_level_value(p)
            if not outline_processing_started and is_processing_start_marker(
                p,
                text,
                numbering_level_lookup=numbering_level_lookup,
                style_numbering_lookup=style_numbering_lookup,
            ):
                outline_processing_started = True

            if not fix_numbered_paragraphs:
                if not outline_processing_started and remove_paragraph_outline_level(p):
                    changed_count += 1
                    increment_removed_preface_outline_count(summary)
                    append_paragraph_change_log(
                        change_logs,
                        part_name,
                        paragraph_index,
                        text,
                        before_outline,
                        "無",
                        "壹、序言前段落，移除既有 Word 大綱階層",
                    )
                continue

            level = None
            reason = None

            if has_auto_numbering(p):
                num_id, ilvl = get_auto_number_identity(p)
                style_id = paragraph_style_id(p)
                level = detect_auto_number_level(
                    p,
                    numbering_level_lookup=numbering_level_lookup,
                    style_numbering_lookup=style_numbering_lookup,
                )
                reason = f"自動編號 numId={num_id} ilvl={ilvl} style={style_id or '無'}"

            if level is None and not should_skip_style_numbering(text):
                style_id = paragraph_style_id(p)
                level = detect_style_number_level(
                    p,
                    numbering_level_lookup=numbering_level_lookup,
                    style_numbering_lookup=style_numbering_lookup,
                )
                if level is not None:
                    reason = f"段落樣式 style={style_id} 對應文件編號階層"

            if level is None:
                level = detect_outline_level(text)
                if level is not None:
                    reason = "段落開頭手動編號"

            if level is None:
                if current_heading_indent is not None:
                    heading_level, heading_uses_outline = current_heading_indent
                    if apply_body_indent_from_heading(p, heading_level, heading_uses_outline):
                        after_outline = before_outline
                        if not outline_processing_started and remove_paragraph_outline_level(p):
                            after_outline = "無"
                            increment_removed_preface_outline_count(summary)

                        changed_count += 1
                        append_paragraph_change_log(
                            change_logs,
                            part_name,
                            paragraph_index,
                            text,
                            before_outline,
                            after_outline,
                            "標題下方內文段落，左縮排對齊上一個標題的文字起點",
                        )
                        continue

                if not outline_processing_started and remove_paragraph_outline_level(p):
                    changed_count += 1
                    increment_removed_preface_outline_count(summary)
                    append_paragraph_change_log(
                        change_logs,
                        part_name,
                        paragraph_index,
                        text,
                        before_outline,
                        "無",
                        "壹、序言前段落，移除既有 Word 大綱階層",
                    )
                    continue

                if summary is not None:
                    summary.unknown_paragraphs += 1
                continue

            set_outline = outline_processing_started
            indent_level = effective_indent_level(level, set_outline=set_outline)

            apply_outline_indent(p, indent_level, set_outline=set_outline)
            current_heading_indent = (indent_level, set_outline)
            record_numbering_measurement(
                summary,
                text=text,
                p=p,
                level=level,
                indent_level=indent_level,
                set_outline=set_outline,
            )
            changed_count += 1
            if set_outline:
                increment_paragraph_level_count(summary, level)
                after_outline = str(level)
                action = f"套用第 {level + 1} 階大綱階層與縮排"
            else:
                if before_outline != "無":
                    increment_removed_preface_outline_count(summary)
                after_outline = "無"
                action = f"壹、序言前段落，只套用第 {indent_level + 1} 階縮排並移除大綱階層"

            append_paragraph_change_log(
                change_logs,
                part_name,
                paragraph_index,
                text,
                before_outline,
                after_outline,
                f"{reason or '編號段落'}，{action}",
            )
        except Exception:
            if summary is not None:
                summary.unknown_paragraphs += 1
            continue

    return changed_count
