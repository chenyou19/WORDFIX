from __future__ import annotations

import re

from lxml import etree

from .constants import DEFAULT_GRAY, NS
from .models import ProcessOptions
from .shading import (
    fix_shading_to_gray,
    fix_shading_to_no_color,
    format_shading_decision,
    get_shading_decision,
    normalize_fill_hex,
)
from .stop_controller import StopController
from .xml_utils import get_or_add, paragraph_text, qn


def table_cell_count(tbl) -> int:
    return sum(len(tr.findall("w:tc", NS)) for tr in tbl.findall("w:tr", NS))


def _cell_grid_span(tc) -> int:
    grid_span = tc.find("w:tcPr/w:gridSpan", NS)
    if grid_span is None:
        return 1

    try:
        return max(int(grid_span.get(qn("val"), "1")), 1)
    except (TypeError, ValueError):
        return 1


def table_column_count(tbl) -> int:
    column_counts: list[int] = []
    for tr in tbl.findall("w:tr", NS):
        cells = tr.findall("w:tc", NS)
        if cells:
            column_counts.append(sum(_cell_grid_span(tc) for tc in cells))

    return max(column_counts, default=0)


def _clear_fixed_width_constraints(tbl) -> None:
    tbl_grid = tbl.find("w:tblGrid", NS)
    if tbl_grid is not None:
        tbl.remove(tbl_grid)

    for tc_pr in tbl.xpath(".//w:tc/w:tcPr", namespaces=NS):
        tc_w = tc_pr.find("w:tcW", NS)
        if tc_w is not None:
            tc_pr.remove(tc_w)


def _apply_table_content_format(tbl, stop: StopController | None = None) -> None:
    for tr in tbl.xpath(".//w:tr", namespaces=NS):
        if stop:
            stop.check()
        trPr = get_or_add(tr, "trPr", first=True)
        trHeight = get_or_add(trPr, "trHeight")
        trHeight.set(qn("val"), "340")
        trHeight.set(qn("hRule"), "atLeast")

    for tc in tbl.xpath(".//w:tc", namespaces=NS):
        if stop:
            stop.check()
        tcPr = get_or_add(tc, "tcPr", first=True)
        vAlign = get_or_add(tcPr, "vAlign")
        vAlign.set(qn("val"), "center")

    for p in tbl.xpath(".//w:p", namespaces=NS):
        if stop:
            stop.check()
        pPr = get_or_add(p, "pPr", first=True)

        p_jc = get_or_add(pPr, "jc")
        p_jc.set(qn("val"), "center")

        spacing = get_or_add(pPr, "spacing")
        spacing.set(qn("before"), "0")
        spacing.set(qn("after"), "0")
        spacing.set(qn("line"), "240")
        spacing.set(qn("lineRule"), "auto")

    for run in tbl.xpath(".//w:r", namespaces=NS):
        if stop:
            stop.check()
        rPr = get_or_add(run, "rPr", first=True)
        sz = get_or_add(rPr, "sz")
        sz.set(qn("val"), "22")
        szCs = get_or_add(rPr, "szCs")
        szCs.set(qn("val"), "22")


TABLE_BORDER_TAGS = ("top", "left", "bottom", "right", "insideH", "insideV")
TABLE_DOUBLE_BORDER_SIZE = "4"
TABLE_BORDER_COLOR = "000000"

# Schema order shared by CT_TblBorders and CT_TcBorders. Border children must be
# kept in this order so Word does not have to recover the part.
_BORDER_SIDE_ORDER = ("top", "left", "bottom", "right", "insideH", "insideV", "tl2br", "tr2bl")


def _get_or_add_border(borders, side: str):
    """Find (or create in schema order) a single border child of a w:tblBorders
    / w:tcBorders container. Creating one side never disturbs the others."""
    existing = borders.find(f"w:{side}", NS)
    if existing is not None:
        return existing

    border = etree.Element(qn(side))
    side_rank = _BORDER_SIDE_ORDER.index(side)
    insert_at = len(borders)
    for index, child in enumerate(borders):
        child_tag = etree.QName(child).localname
        if child_tag in _BORDER_SIDE_ORDER and _BORDER_SIDE_ORDER.index(child_tag) > side_rank:
            insert_at = index
            break
    borders.insert(insert_at, border)
    return border


def _configure_double_black_border(border) -> None:
    border.set(qn("val"), "double")
    border.set(qn("sz"), TABLE_DOUBLE_BORDER_SIZE)
    border.set(qn("space"), "0")
    border.set(qn("color"), TABLE_BORDER_COLOR)


def set_border_double_black(borders, side: str) -> None:
    """Set a single side of a borders container to a black double line.

    Only the named side is updated; any other existing borders are preserved.
    """
    _configure_double_black_border(_get_or_add_border(borders, side))


def set_border_nil(borders, side: str) -> None:
    """Set a single side of a borders container to no border (w:val="nil").

    Only the named side is updated; any other existing borders are preserved.
    """
    border = _get_or_add_border(borders, side)
    for attr in ("sz", "space", "color"):
        attr_name = qn(attr)
        if attr_name in border.attrib:
            del border.attrib[attr_name]
    border.set(qn("val"), "nil")


def _get_or_add_tbl_borders(tbl):
    tblPr = get_or_add(tbl, "tblPr", first=True)
    return get_or_add(tblPr, "tblBorders")


def _get_or_add_tc_borders(tc):
    tcPr = get_or_add(tc, "tcPr", first=True)
    return get_or_add(tcPr, "tcBorders")


def apply_double_black_table_borders(tbl) -> None:
    """Apply a black double-line border to the whole table.

    Rebuilds w:tblBorders so the outer frame (top/left/bottom/right) and the
    inner gridlines (insideH/insideV) are all black double lines. This must run
    after note cells/rows are removed so newly exposed edges still get borders.
    """
    tblPr = get_or_add(tbl, "tblPr", first=True)

    tbl_borders = tblPr.find("w:tblBorders", NS)
    if tbl_borders is not None:
        tblPr.remove(tbl_borders)
    tbl_borders = etree.SubElement(tblPr, qn("tblBorders"))

    for tag in TABLE_BORDER_TAGS:
        _configure_double_black_border(etree.SubElement(tbl_borders, qn(tag)))


def apply_table_format(tbl, stop: StopController | None = None) -> None:
    tblPr = get_or_add(tbl, "tblPr", first=True)

    jc = get_or_add(tblPr, "jc")
    jc.set(qn("val"), "center")

    tblW = get_or_add(tblPr, "tblW")
    tblW.set(qn("type"), "pct")
    tblW.set(qn("w"), "5000")

    _clear_fixed_width_constraints(tbl)

    tblLayout = get_or_add(tblPr, "tblLayout")
    tblLayout.set(qn("type"), "autofit")

    _apply_table_content_format(tbl, stop=stop)


def _rebuild_fixed_column_widths(tbl, width_twips: int) -> list[int]:
    column_count = table_column_count(tbl)
    if column_count <= 0 or width_twips <= 0:
        return []

    base_width = width_twips // column_count
    remainder = width_twips - base_width * column_count
    column_widths = [
        base_width + (1 if index < remainder else 0)
        for index in range(column_count)
    ]

    tblPr = get_or_add(tbl, "tblPr", first=True)
    tbl_grid = etree.Element(qn("tblGrid"))
    for column_width in column_widths:
        grid_col = etree.SubElement(tbl_grid, qn("gridCol"))
        grid_col.set(qn("w"), str(column_width))
    tbl.insert(tbl.index(tblPr) + 1, tbl_grid)

    for tr in tbl.findall("w:tr", NS):
        cursor = 0
        for tc in tr.findall("w:tc", NS):
            span = _cell_grid_span(tc)
            cell_width = sum(column_widths[cursor : cursor + span])
            cursor += span
            if cell_width <= 0:
                continue
            tcPr = get_or_add(tc, "tcPr", first=True)
            tcW = tcPr.find("w:tcW", NS)
            if tcW is None:
                tcW = etree.Element(qn("tcW"))
                tcPr.insert(0, tcW)
            tcW.set(qn("type"), "dxa")
            tcW.set(qn("w"), str(cell_width))

    return column_widths


def apply_special_table_format(
    tbl,
    *,
    left_indent_twips: int,
    width_twips: int,
    stop: StopController | None = None,
) -> None:
    if stop:
        stop.check()

    _clear_fixed_width_constraints(tbl)

    tblPr = get_or_add(tbl, "tblPr", first=True)

    jc = get_or_add(tblPr, "jc")
    jc.set(qn("val"), "left")

    tblW = get_or_add(tblPr, "tblW")
    tblW.set(qn("type"), "dxa")
    tblW.set(qn("w"), str(width_twips))

    tblLayout = get_or_add(tblPr, "tblLayout")
    tblLayout.set(qn("type"), "fixed")

    tblInd = get_or_add(tblPr, "tblInd")
    tblInd.set(qn("type"), "dxa")
    tblInd.set(qn("w"), str(left_indent_twips))

    _rebuild_fixed_column_widths(tbl, width_twips)

    _apply_table_content_format(tbl, stop=stop)


def apply_autofit_contents_right_format(tbl, stop: StopController | None = None) -> None:
    if stop:
        stop.check()

    tblPr = get_or_add(tbl, "tblPr", first=True)

    jc = get_or_add(tblPr, "jc")
    jc.set(qn("val"), "right")

    tblW = get_or_add(tblPr, "tblW")
    tblW.set(qn("type"), "auto")
    tblW.set(qn("w"), "0")

    tblLayout = get_or_add(tblPr, "tblLayout")
    tblLayout.set(qn("type"), "autofit")

    _apply_table_content_format(tbl, stop=stop)


def table_has_special_skip_color(
    tbl,
    special_color_skip_colors: tuple[str, ...] | list[str],
) -> tuple[bool, list[str]]:
    targets = {color for color in special_color_skip_colors}
    if not targets:
        return False, []

    matched: list[str] = []
    for shd in tbl.xpath(".//w:tc/w:tcPr/w:shd", namespaces=NS):
        fill_hex = normalize_fill_hex(shd.get(qn("fill")))
        if fill_hex is not None and fill_hex in targets and fill_hex not in matched:
            matched.append(fill_hex)
    return bool(matched), matched


def clear_matching_special_colors(
    tbl,
    special_color_skip_colors: tuple[str, ...] | list[str],
) -> int:
    targets = {color for color in special_color_skip_colors}
    if not targets:
        return 0

    cleared = 0
    for shd in tbl.xpath(".//w:tc/w:tcPr/w:shd", namespaces=NS):
        fill_hex = normalize_fill_hex(shd.get(qn("fill")))
        if fill_hex is not None and fill_hex in targets:
            fix_shading_to_no_color(shd)
            cleared += 1
    return cleared


def apply_table_color(
    tbl,
    stop: StopController | None = None,
    *,
    keep_colors: tuple[str, ...] | list[str] = (),
    gray_colors: tuple[str, ...] | list[str] = (),
    gray_target: str = DEFAULT_GRAY,
) -> tuple[int, int, list[str]]:
    changed_to_gray = 0
    cleared_colors = 0
    shading_debug_logs: list[str] = []

    for tc in tbl.xpath(".//w:tc", namespaces=NS):
        if stop:
            stop.check()

        tcPr = tc.find("w:tcPr", NS)
        if tcPr is None:
            continue

        for shd in tcPr.findall("w:shd", NS):
            decision = get_shading_decision(
                shd,
                keep_colors=keep_colors,
                gray_colors=gray_colors,
                gray_target=gray_target,
            )
            action = decision["action"]
            shading_debug_logs.append(format_shading_decision(decision))

            if action == "gray":
                fix_shading_to_gray(shd, gray_target)
                changed_to_gray += 1
            elif action == "clear":
                fix_shading_to_no_color(shd)
                cleared_colors += 1

    return changed_to_gray, cleared_colors, shading_debug_logs


def process_table(
    tbl,
    options: ProcessOptions,
    stop: StopController | None = None,
    *,
    special_layout: bool = False,
    special_table_geometry: tuple[int, int] | None = None,
) -> tuple[int, int, list[str]]:
    changed_to_gray = 0
    cleared_colors = 0
    shading_debug_logs: list[str] = []

    if options.fix_table_layout:
        if special_layout:
            if special_table_geometry is not None:
                left_indent_twips, width_twips = special_table_geometry
                apply_special_table_format(
                    tbl,
                    left_indent_twips=left_indent_twips,
                    width_twips=width_twips,
                    stop=stop,
                )
            else:
                apply_autofit_contents_right_format(tbl, stop=stop)
        else:
            apply_table_format(tbl, stop=stop)

    if options.fix_color:
        changed_to_gray, cleared_colors, shading_debug_logs = apply_table_color(
            tbl,
            stop=stop,
            keep_colors=tuple(getattr(options, "table_keep_colors", ()) or ()),
            gray_colors=tuple(getattr(options, "table_gray_colors", ()) or ()),
            gray_target=str(getattr(options, "table_gray_target", DEFAULT_GRAY) or DEFAULT_GRAY),
        )

    return changed_to_gray, cleared_colors, shading_debug_logs


# 「表格最後一列說明格式化」 (enable_table_footer_source_format)
#
# An independent, opt-in table layout post-step. When enabled for a table it
# applies, in this fixed order:
#   1. whole-table font size -> 11 pt
#   2. table outer frame (top/bottom/left/right) -> black double border
#   3. first-row single cell -> top/left/right no border, bottom black double
#   4. last-row 基期：/資料來源：/註記 cells -> 10 pt, aligned, top black double,
#      left/right/bottom no border
# Steps 3 and 4 only do local cell-border updates, so other cells keep any
# borders they already had. Later steps deliberately override earlier ones
# (10 pt over 11 pt, cell borders over the table frame). This never moves,
# deletes or adds cells/rows/paragraphs; it only formats matching last-row cells.
FOOTER_SOURCE_BODY_FONT_SIZE_HALF_POINTS = "22"  # 11 pt
FOOTER_SOURCE_NOTE_FONT_SIZE_HALF_POINTS = "20"  # 10 pt
FOOTER_SOURCE_BASE_PERIOD_PREFIX = "基期："
FOOTER_SOURCE_DATA_SOURCE_PREFIX = "資料來源："

# A last-row note cell starts directly with 「註 + 可選阿拉伯數字 + 全形或半形冒號」.
# Matches 註：/註:/註1：/註1:/註10：/註10:; rejects 備註：/註記：/本註：/說明註：
# (註 must be at the very start, immediately followed by an optional number and
# a colon).
FOOTER_NOTE_PREFIX_PATTERN = re.compile(r"^註(?:\d+)?[：:]")

# Zero-width / control characters plus the Word cell-end mark that can trail the
# visible text of a cell. Stripped before the 基期/資料來源 prefix test so a
# stray newline or invisible character cannot defeat the match.
_FOOTER_CELL_CONTROL_CHARS = "\ufeff\u200b\u200c\u200d\u2060\u0007"


def normalize_footer_source_cell_text(tc) -> str:
    """Merge a cell's paragraph text into one normalized string for prefix tests.

    Joins every paragraph, drops zero-width/control characters and the Word
    cell mark, then collapses runs of whitespace (including newlines) so leading
    or trailing spaces never break the 基期：/資料來源： startswith check.
    """
    text = " ".join(paragraph_text(p) for p in tc.findall("w:p", NS))
    for char in _FOOTER_CELL_CONTROL_CHARS:
        text = text.replace(char, "")
    return " ".join(text.split())


def _set_runs_font_size(scope, half_points: str) -> None:
    for run in scope.xpath(".//w:r", namespaces=NS):
        rPr = get_or_add(run, "rPr", first=True)
        get_or_add(rPr, "sz").set(qn("val"), half_points)
        get_or_add(rPr, "szCs").set(qn("val"), half_points)


def _set_paragraph_alignment(scope, alignment: str) -> None:
    for p in scope.findall("w:p", NS):
        pPr = get_or_add(p, "pPr", first=True)
        get_or_add(pPr, "jc").set(qn("val"), alignment)


def _apply_table_outer_double_black_borders(tbl) -> None:
    borders = _get_or_add_tbl_borders(tbl)
    for side in ("top", "bottom", "left", "right"):
        set_border_double_black(borders, side)


def _apply_first_row_single_cell_borders(tc) -> None:
    borders = _get_or_add_tc_borders(tc)
    set_border_nil(borders, "top")
    set_border_nil(borders, "left")
    set_border_nil(borders, "right")
    set_border_double_black(borders, "bottom")


def _apply_last_row_footer_cell_borders(tc) -> None:
    borders = _get_or_add_tc_borders(tc)
    set_border_double_black(borders, "top")
    set_border_nil(borders, "left")
    set_border_nil(borders, "right")
    set_border_nil(borders, "bottom")


def _classify_footer_cell(text: str) -> tuple[str, str] | None:
    """Return (cell_type, alignment) for a normalized last-row cell, or None.

    註記 / 基期 / 資料來源 share the same last-row footer treatment; only the
    alignment differs (資料來源 is right-aligned, the rest are left-aligned).
    """
    if text.startswith(FOOTER_SOURCE_BASE_PERIOD_PREFIX):
        return "base_period", "left"
    if text.startswith(FOOTER_SOURCE_DATA_SOURCE_PREFIX):
        return "source", "right"
    if FOOTER_NOTE_PREFIX_PATTERN.match(text):
        return "note", "left"
    return None


def apply_table_footer_source_format(tbl, stop: StopController | None = None) -> dict[str, object]:
    """Apply the 「表格最後一列說明格式化」 format to one table.

    Formats matching last-row cells (基期：/資料來源：/註記) plus the table
    frame and a single-cell first row. Never moves, deletes or inserts any
    cell, row or paragraph. Returns a dict describing what was changed.
    """
    result: dict[str, object] = {
        "outer_double_border_applied": False,
        "first_row_single_cell_border_adjusted": False,
        "footer_note_cells_adjusted": 0,
        "footer_note_cell_matches": [],
        "footer_note_cell_debug": [],
    }

    if stop:
        stop.check()

    # 1. whole-table font size -> 11 pt
    _set_runs_font_size(tbl, FOOTER_SOURCE_BODY_FONT_SIZE_HALF_POINTS)

    # 2. table outer frame -> black double border
    _apply_table_outer_double_black_borders(tbl)
    result["outer_double_border_applied"] = True

    rows = tbl.findall("w:tr", NS)
    if not rows:
        return result

    # 3. first-row single cell overrides the outer frame on three edges.
    first_row_cells = rows[0].findall("w:tc", NS)
    if len(first_row_cells) == 1:
        _apply_first_row_single_cell_borders(first_row_cells[0])
        result["first_row_single_cell_border_adjusted"] = True

    # 4. last-row 基期：/資料來源：/註記 cells. Only matched cells are touched.
    # De-duplicate by the underlying XML element so a merged / spanned cell is
    # never formatted or logged twice.
    seen_cells: set[int] = set()
    for tc in rows[-1].findall("w:tc", NS):
        if id(tc) in seen_cells:
            continue
        seen_cells.add(id(tc))
        if stop:
            stop.check()

        text = normalize_footer_source_cell_text(tc)
        classified = _classify_footer_cell(text)
        if classified is None:
            continue
        cell_type, alignment = classified

        _set_runs_font_size(tc, FOOTER_SOURCE_NOTE_FONT_SIZE_HALF_POINTS)
        _set_paragraph_alignment(tc, alignment)
        _apply_last_row_footer_cell_borders(tc)

        result["footer_note_cells_adjusted"] += 1
        result["footer_note_cell_matches"].append(cell_type)
        result["footer_note_cell_debug"].append(f"{cell_type}: {text[:50]}")

    return result
