from __future__ import annotations

import re
from collections import namedtuple

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

# Parent container order matters too. Word may keep out-of-order XML but ignore
# or repair borders during rendering, so table-specific border containers are
# inserted and relocated without changing the global get_or_add() helper.
_TBL_PR_CHILD_ORDER = (
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
    "tblPrChange",
)
_TC_PR_CHILD_ORDER = (
    "cnfStyle",
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFitText",
    "vAlign",
    "hideMark",
    "headers",
    "cellIns",
    "cellDel",
    "cellMerge",
    "tcPrChange",
)


def _local_name(element) -> str:
    return etree.QName(element).localname


def _child_order(parent) -> list[str]:
    if parent is None:
        return []
    return [_local_name(child) for child in parent]


def _child_order_text(parent) -> str:
    order = _child_order(parent)
    return ",".join(order) if order else "none"


def _is_known_child_order_valid(parent, order: tuple[str, ...]) -> bool:
    if parent is None:
        return True
    ranks = {tag: index for index, tag in enumerate(order)}
    previous_rank = -1
    for child in parent:
        rank = ranks.get(_local_name(child))
        if rank is None:
            continue
        if rank < previous_rank:
            return False
        previous_rank = rank
    return True


def _is_child_at_schema_position(parent, tag: str, order: tuple[str, ...]) -> bool:
    if parent is None:
        return False
    ranks = {name: index for index, name in enumerate(order)}
    target_rank = ranks[tag]
    seen_target = False
    for child in parent:
        child_tag = _local_name(child)
        if child_tag == tag:
            seen_target = True
            continue
        child_rank = ranks.get(child_tag)
        if child_rank is None:
            continue
        if not seen_target and child_rank > target_rank:
            return False
        if seen_target and child_rank < target_rank:
            return False
    return seen_target


def _get_or_add_child_in_schema_order(parent, tag: str, order: tuple[str, ...]):
    child = parent.find(f"w:{tag}", NS)
    if child is None:
        child = etree.Element(qn(tag))
    else:
        parent.remove(child)

    ranks = {name: index for index, name in enumerate(order)}
    target_rank = ranks[tag]
    insert_at = len(parent)
    for index, sibling in enumerate(parent):
        sibling_rank = ranks.get(_local_name(sibling))
        if sibling_rank is not None and sibling_rank > target_rank:
            insert_at = index
            break
    parent.insert(insert_at, child)
    return child


def _normalize_known_children_in_schema_order(parent, order: tuple[str, ...]) -> None:
    """Sort existing known OOXML children into schema order.

    Unknown children are preserved and kept after the known children in their
    original relative order. The helper only moves existing nodes; it does not
    create or clone anything.
    """
    if parent is None or len(parent) < 2:
        return

    ranks = {name: index for index, name in enumerate(order)}
    original_children = list(parent)
    known_children = [
        (index, child)
        for index, child in enumerate(original_children)
        if _local_name(child) in ranks
    ]
    unknown_children = [
        (index, child)
        for index, child in enumerate(original_children)
        if _local_name(child) not in ranks
    ]
    if len(known_children) > 1:
        known_children.sort(key=lambda item: (ranks[_local_name(item[1])], item[0]))

    reordered_children = [child for _, child in known_children + unknown_children]
    if reordered_children == original_children:
        return

    for child in original_children:
        parent.remove(child)
    for child in reordered_children:
        parent.append(child)


def _normalize_border_children_in_schema_order(borders) -> None:
    _normalize_known_children_in_schema_order(borders, _BORDER_SIDE_ORDER)


def _normalize_tbl_pr_known_children_in_schema_order(tbl) -> None:
    tbl_pr = tbl.find("w:tblPr", NS)
    if tbl_pr is None:
        return
    _normalize_known_children_in_schema_order(tbl_pr, _TBL_PR_CHILD_ORDER)
    for tbl_borders in tbl_pr.findall("w:tblBorders", NS):
        _normalize_border_children_in_schema_order(tbl_borders)


def _normalize_tc_pr_known_children_in_schema_order(tc) -> None:
    tc_pr = tc.find("w:tcPr", NS)
    if tc_pr is None:
        return
    _normalize_known_children_in_schema_order(tc_pr, _TC_PR_CHILD_ORDER)
    for tc_borders in tc_pr.findall("w:tcBorders", NS):
        _normalize_border_children_in_schema_order(tc_borders)


def _normalize_table_border_related_schema_order(tbl) -> None:
    _normalize_tbl_pr_known_children_in_schema_order(tbl)
    for tc in tbl.xpath(".//w:tc", namespaces=NS):
        _normalize_tc_pr_known_children_in_schema_order(tc)


def _get_or_add_tbl_borders_in_schema_order(tbl):
    tblPr = get_or_add(tbl, "tblPr", first=True)
    return _get_or_add_child_in_schema_order(tblPr, "tblBorders", _TBL_PR_CHILD_ORDER)


def _get_or_add_tc_borders_in_schema_order(tc):
    tcPr = get_or_add(tc, "tcPr", first=True)
    return _get_or_add_child_in_schema_order(tcPr, "tcBorders", _TC_PR_CHILD_ORDER)


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
    return _get_or_add_tbl_borders_in_schema_order(tbl)


def _get_or_add_tc_borders(tc):
    return _get_or_add_tc_borders_in_schema_order(tc)


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
    tbl_borders = _get_or_add_tbl_borders_in_schema_order(tbl)
    for child in list(tbl_borders):
        tbl_borders.remove(child)

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
#   3. first-row:
#      - single-cell title -> top/left/right nil, bottom black double;
#      - otherwise, direct top black double on every first-row cell so Word
#        does not let existing cell borders override the table top frame.
#   4. footer block: collect the contiguous block of footer rows at the BOTTOM
#      of the table (scan upward; a row is a footer row when any of its cells
#      matches 基期：/資料來源：/註記; stop at the first non-matching/blank row),
#      then format the block as a unit:
#        - the TOP footer row gets a black double TOP border across EVERY cell
#          (the single data/footer separator, spanning the whole width);
#        - every other footer row gets its TOP border cleared (so there is no
#          line between consecutive 註1/註2/註3 rows);
#        - matched cells in every footer row -> 10 pt, aligned, left/right/bottom
#          no border.
#   5. final bottom edge decision:
#      - no bottom footer rows -> visible data edge is black double;
#      - bottom footer rows -> table/terminal footer bottom is nil, while the
#        separator above the footer block remains black double.
#   6. outer vertical edge policy:
#      - non-footer rows get direct left/right black double borders on the
#        logical outer cells (gridSpan-aware; vMerge continuation-aware);
#      - footer rows get direct left/right nil on their logical outer cells.
#   7. full table-related schema normalization for tblPr/tcPr/border children.
# All border updates are local (one side at a time), so other cells keep any
# borders they already had. This never moves, deletes or adds any cell, row or
# paragraph; it only formats the matching footer cells (plus the separator).
FOOTER_SOURCE_BODY_FONT_SIZE_HALF_POINTS = "22"  # 11 pt
FOOTER_SOURCE_NOTE_FONT_SIZE_HALF_POINTS = "20"  # 10 pt

# A note cell starts directly with 「註 + 可選空白 + 可選阿拉伯數字 + 可選空白 +
# 全形或半形冒號」. Matches 註：/註:/註1：/註 1：/註　2：/註10:; rejects
# 備註：/註記：/本註：/說明註：（註 必須在最前面，後面緊跟可選空白、可選數字、冒號）。
FOOTER_NOTE_PREFIX_PATTERN = re.compile(r"^註[ \t　]*(?:\d+[ \t　]*)?[：:]")
FOOTER_BASE_PERIOD_PREFIX_PATTERN = re.compile(r"^基期[：:]")
FOOTER_SOURCE_PREFIX_PATTERN = re.compile(r"^資料來源[：:]")

# Paragraph alignment per footer cell type (資料來源 right, the rest left).
FOOTER_ALIGNMENT_BY_TYPE = {"base_period": "left", "source": "right", "note": "left"}

# A collected bottom footer row: its 0-based row index, its de-duplicated cells,
# and the matched cells as (tc, cell_type, normalized_text).
_FooterRow = namedtuple("_FooterRow", ["row_index", "cells", "matches"])

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


def classify_footer_cell_text(text: str) -> str | None:
    """Classify normalized cell text as a footer cell type, or None.

    Returns "base_period", "source", "note", or None when it is not a footer
    cell. Alignment per type is in FOOTER_ALIGNMENT_BY_TYPE.
    """
    if FOOTER_BASE_PERIOD_PREFIX_PATTERN.match(text):
        return "base_period"
    if FOOTER_SOURCE_PREFIX_PATTERN.match(text):
        return "source"
    if FOOTER_NOTE_PREFIX_PATTERN.match(text):
        return "note"
    return None


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


def _apply_rendered_top_double_black_border(tbl) -> int:
    rows = tbl.findall("w:tr", NS)
    if not rows:
        return 0

    first_row_cells = _unique_row_cells(rows[0])
    if len(first_row_cells) == 1:
        return 0

    applied = 0
    for tc in first_row_cells:
        set_border_double_black(_get_or_add_tc_borders(tc), "top")
        applied += 1
    return applied


def _unique_row_cells(tr) -> list:
    """A row's cells, de-duplicated by XML element (merged / spanned cells)."""
    seen: set[int] = set()
    cells = []
    for tc in tr.findall("w:tc", NS):
        if id(tc) in seen:
            continue
        seen.add(id(tc))
        cells.append(tc)
    return cells


def _cell_vmerge_state(tc) -> str:
    v_merge = tc.find("w:tcPr/w:vMerge", NS)
    if v_merge is None:
        return "none"
    value = v_merge.get(qn("val"))
    if value in (None, "", "continue"):
        return "continue"
    return str(value)


def _row_cell_infos(tr) -> list[dict[str, object]]:
    cursor = 0
    infos: list[dict[str, object]] = []
    for cell_index, tc in enumerate(_unique_row_cells(tr)):
        grid_span = _cell_grid_span(tc)
        infos.append(
            {
                "tc": tc,
                "cell_index": cell_index,
                "start_col": cursor,
                "end_col": cursor + grid_span,
                "grid_span": grid_span,
                "vmerge_state": _cell_vmerge_state(tc),
            }
        )
        cursor += grid_span
    return infos


def _find_vmerge_restart_owner(rows: list, target_info: dict[str, object]):
    target_start = int(target_info["start_col"])
    target_end = int(target_info["end_col"])
    for tr in reversed(rows[:-1]):
        for info in _row_cell_infos(tr):
            if int(info["start_col"]) != target_start or int(info["end_col"]) != target_end:
                continue
            state = str(info["vmerge_state"])
            if state == "restart":
                return info["tc"]
            if state == "none":
                return None
            break
    return None


def _find_vmerge_restart_owner_before(
    rows: list,
    row_index: int,
    target_info: dict[str, object],
):
    target_start = int(target_info["start_col"])
    target_end = int(target_info["end_col"])
    for tr in reversed(rows[:row_index]):
        for info in _row_cell_infos(tr):
            if int(info["start_col"]) != target_start or int(info["end_col"]) != target_end:
                continue
            state = str(info["vmerge_state"])
            if state == "restart":
                return info["tc"]
            if state == "none":
                return None
            break
    return None


def _last_row_bottom_edge_target_cells(tbl) -> list:
    rows = tbl.findall("w:tr", NS)
    if not rows:
        return []

    targets = []
    seen: set[int] = set()
    last_row_infos = _row_cell_infos(rows[-1])
    for info in last_row_infos:
        tc = info["tc"]
        if id(tc) not in seen:
            seen.add(id(tc))
            targets.append(tc)
        if str(info["vmerge_state"]) == "continue":
            owner = _find_vmerge_restart_owner(rows, info)
            if owner is not None and id(owner) not in seen:
                seen.add(id(owner))
                targets.append(owner)
    return targets


def _last_row_diagnostics(tbl) -> dict[str, object]:
    rows = tbl.findall("w:tr", NS)
    if not rows:
        return {
            "last_row_physical_cell_count": 0,
            "last_row_grid_span_sum": 0,
            "last_row_vmerge_states": "none",
            "last_row_bottom_edge_target_count": 0,
        }

    last_row_infos = _row_cell_infos(rows[-1])
    targets = _last_row_bottom_edge_target_cells(tbl)
    vmerge_states = [str(info["vmerge_state"]) for info in last_row_infos]
    return {
        "last_row_physical_cell_count": len(last_row_infos),
        "last_row_grid_span_sum": sum(int(info["grid_span"]) for info in last_row_infos),
        "last_row_vmerge_states": "|".join(vmerge_states) if vmerge_states else "none",
        "last_row_bottom_edge_target_count": len(targets),
    }


def _apply_rendered_bottom_double_black_border(tbl) -> int:
    """Force the last visual data edge's cell bottoms to render black double.

    Word may let direct ``w:tcBorders/w:bottom`` values override the table-level
    bottom border. For plain and gridSpan rows, the physical last-row cells own
    the visible edge. When the last row continues a vertical merge, the restart
    owner is also touched so Word has a bottom value on both the visible
    continuation and the merge owner. Only the bottom side is touched.
    """
    applied = 0
    for tc in _last_row_bottom_edge_target_cells(tbl):
        set_border_double_black(_get_or_add_tc_borders(tc), "bottom")
        applied += 1
    return applied


def _apply_footer_terminal_bottom_none(tbl, footer_rows: list) -> int:
    """Clear the visible bottom under a bottom footer block.

    The data/footer separator lives on the top footer row. Once a footer exists,
    the table must not end with a visible bottom border under the explanation
    text, so both table-level bottom and every physical cell in the terminal
    footer row are set to nil.
    """
    if not footer_rows:
        return 0

    set_border_nil(_get_or_add_tbl_borders(tbl), "bottom")
    applied = 0
    for tc in footer_rows[-1].cells:
        set_border_nil(_get_or_add_tc_borders(tc), "bottom")
        applied += 1
    return applied


def _footer_row_index_set(footer_rows: list) -> set[int]:
    return {int(footer_row.row_index) for footer_row in footer_rows}


def _target_descriptor(target: dict[str, object]) -> str:
    return (
        f"r{target['row_index']}c{target['cell_index']}"
        f"[{target['start_col']}-{target['end_col']}]"
        f":{target['vmerge_state']}"
    )


def _outer_vertical_border_targets(tbl, footer_rows: list) -> dict[str, object]:
    rows = tbl.findall("w:tr", NS)
    footer_indices = _footer_row_index_set(footer_rows)
    column_count = table_column_count(tbl)
    targets: dict[str, object] = {
        "data_row_indices": [],
        "footer_row_indices": [footer_row.row_index for footer_row in footer_rows],
        "data_left_targets": [],
        "data_right_targets": [],
        "footer_left_targets": [],
        "footer_right_targets": [],
        "data_left_vmerge_owner_targets": [],
        "data_right_vmerge_owner_targets": [],
    }
    if column_count <= 0:
        return targets

    for row_index, tr in enumerate(rows):
        row_is_footer = row_index in footer_indices
        row_infos = _row_cell_infos(tr)
        first_row_single_cell_title = (
            row_index == 0 and not row_is_footer and len(row_infos) == 1
        )
        if not row_is_footer and not first_row_single_cell_title:
            targets["data_row_indices"].append(row_index)
        if first_row_single_cell_title:
            continue

        left_key = "footer_left_targets" if row_is_footer else "data_left_targets"
        right_key = "footer_right_targets" if row_is_footer else "data_right_targets"
        for info in row_infos:
            target = {
                **info,
                "row_index": row_index,
            }
            if int(info["start_col"]) == 0:
                targets[left_key].append(target)
            if int(info["end_col"]) == column_count:
                targets[right_key].append(target)
    return targets


def _apply_border_to_targets(
    targets: list[dict[str, object]],
    side: str,
    border_setter,
    *,
    rows: list | None = None,
    include_vmerge_owner: bool = False,
) -> tuple[int, list[str]]:
    applied = 0
    owner_descriptors: list[str] = []
    seen: set[tuple[int, str]] = set()
    for target in targets:
        tc = target["tc"]
        key = (id(tc), side)
        if key not in seen:
            border_setter(_get_or_add_tc_borders(tc), side)
            seen.add(key)
            applied += 1
        if (
            include_vmerge_owner
            and rows is not None
            and str(target["vmerge_state"]) == "continue"
        ):
            owner = _find_vmerge_restart_owner_before(
                rows,
                int(target["row_index"]),
                target,
            )
            owner_key = (id(owner), side) if owner is not None else None
            if owner is not None:
                owner_descriptors.append(
                    f"r{target['row_index']}c{target['cell_index']}->{side}:restart_owner"
                )
                if owner_key not in seen:
                    border_setter(_get_or_add_tc_borders(owner), side)
                    seen.add(owner_key)
                    applied += 1
    return applied, owner_descriptors


def _apply_data_rows_outer_vertical_double_borders(
    tbl,
    footer_rows: list,
) -> dict[str, object]:
    targets = _outer_vertical_border_targets(tbl, footer_rows)
    rows = tbl.findall("w:tr", NS)
    left_count, left_owner_targets = _apply_border_to_targets(
        targets["data_left_targets"],
        "left",
        set_border_double_black,
        rows=rows,
        include_vmerge_owner=True,
    )
    right_count, right_owner_targets = _apply_border_to_targets(
        targets["data_right_targets"],
        "right",
        set_border_double_black,
        rows=rows,
        include_vmerge_owner=True,
    )
    return {
        "data_rows_outer_left_double_applied": left_count > 0,
        "data_rows_outer_right_double_applied": right_count > 0,
        "data_rows_outer_left_target_count": len(targets["data_left_targets"]),
        "data_rows_outer_right_target_count": len(targets["data_right_targets"]),
        "data_rows_outer_left_vmerge_owner_target_count": len(left_owner_targets),
        "data_rows_outer_right_vmerge_owner_target_count": len(right_owner_targets),
    }


def _apply_footer_rows_outer_vertical_none_borders(
    tbl,
    footer_rows: list,
) -> dict[str, object]:
    targets = _outer_vertical_border_targets(tbl, footer_rows)
    left_count, _ = _apply_border_to_targets(
        targets["footer_left_targets"],
        "left",
        set_border_nil,
    )
    right_count, _ = _apply_border_to_targets(
        targets["footer_right_targets"],
        "right",
        set_border_nil,
    )
    return {
        "footer_rows_outer_left_none_applied": left_count > 0,
        "footer_rows_outer_right_none_applied": right_count > 0,
        "footer_rows_outer_left_target_count": len(targets["footer_left_targets"]),
        "footer_rows_outer_right_target_count": len(targets["footer_right_targets"]),
    }


def _border_signature(border) -> str:
    if border is None:
        return "missing"
    return "/".join(
        [
            border.get(qn("val"), "missing"),
            border.get(qn("sz"), "missing"),
            border.get(qn("color"), "missing"),
        ]
    )


def _is_nil_border(border) -> bool:
    return border is not None and border.get(qn("val")) == "nil"


def _is_double_black_border(border) -> bool:
    return (
        border is not None
        and border.get(qn("val")) == "double"
        and border.get(qn("sz")) == TABLE_DOUBLE_BORDER_SIZE
        and border.get(qn("color")) == TABLE_BORDER_COLOR
    )


def _target_border_values(targets: list[dict[str, object]], side: str) -> list[str]:
    return [
        _border_signature(target["tc"].find(f"w:tcPr/w:tcBorders/w:{side}", NS))
        for target in targets
    ]


def _format_index_list(values: list[int]) -> str:
    return ",".join(str(value) for value in values) if values else "none"


def _format_target_list(targets: list[dict[str, object]]) -> str:
    return "|".join(_target_descriptor(target) for target in targets) or "none"


def _format_border_value_list(values: list[str]) -> str:
    return "|".join(values) if values else "none"


def _verify_outer_vertical_border_policy(tbl, footer_rows: list) -> dict[str, object]:
    targets = _outer_vertical_border_targets(tbl, footer_rows)
    data_left_targets = targets["data_left_targets"]
    data_right_targets = targets["data_right_targets"]
    footer_left_targets = targets["footer_left_targets"]
    footer_right_targets = targets["footer_right_targets"]

    data_left_values = _target_border_values(data_left_targets, "left")
    data_right_values = _target_border_values(data_right_targets, "right")
    footer_left_values = _target_border_values(footer_left_targets, "left")
    footer_right_values = _target_border_values(footer_right_targets, "right")

    data_left_ok = all(
        _is_double_black_border(target["tc"].find("w:tcPr/w:tcBorders/w:left", NS))
        for target in data_left_targets
    )
    data_right_ok = all(
        _is_double_black_border(target["tc"].find("w:tcPr/w:tcBorders/w:right", NS))
        for target in data_right_targets
    )
    footer_left_ok = all(
        _is_nil_border(target["tc"].find("w:tcPr/w:tcBorders/w:left", NS))
        for target in footer_left_targets
    )
    footer_right_ok = all(
        _is_nil_border(target["tc"].find("w:tcPr/w:tcBorders/w:right", NS))
        for target in footer_right_targets
    )
    data_targets_present = bool(targets["data_row_indices"])
    footer_targets_present = bool(targets["footer_row_indices"])
    verified = (
        (not data_targets_present or (bool(data_left_targets) and data_left_ok))
        and (not data_targets_present or (bool(data_right_targets) and data_right_ok))
        and (not footer_targets_present or (bool(footer_left_targets) and footer_left_ok))
        and (not footer_targets_present or (bool(footer_right_targets) and footer_right_ok))
    )
    detail = ";".join(
        [
            f"data_row_indices={_format_index_list(targets['data_row_indices'])}",
            f"footer_row_indices={_format_index_list(targets['footer_row_indices'])}",
            f"data_left_targets={_format_target_list(data_left_targets)}",
            f"data_right_targets={_format_target_list(data_right_targets)}",
            f"footer_left_targets={_format_target_list(footer_left_targets)}",
            f"footer_right_targets={_format_target_list(footer_right_targets)}",
            f"data_left_border_values={_format_border_value_list(data_left_values)}",
            f"data_right_border_values={_format_border_value_list(data_right_values)}",
            f"footer_left_border_values={_format_border_value_list(footer_left_values)}",
            f"footer_right_border_values={_format_border_value_list(footer_right_values)}",
        ]
    )
    return {
        "outer_vertical_border_policy_xml_verified": verified,
        "outer_vertical_border_policy_verify_detail": detail,
    }


def _schema_order_summary(tbl, cells: list) -> dict[str, object]:
    tbl_pr = tbl.find("w:tblPr", NS)
    tbl_borders = tbl.find("w:tblPr/w:tblBorders", NS)
    tc_pr_orders = []
    valid = _is_known_child_order_valid(tbl_pr, _TBL_PR_CHILD_ORDER)
    valid = valid and _is_known_child_order_valid(tbl_borders, _BORDER_SIDE_ORDER)

    for tc_pr in tbl.xpath(".//w:tc/w:tcPr", namespaces=NS):
        valid = valid and _is_known_child_order_valid(tc_pr, _TC_PR_CHILD_ORDER)
        tc_borders = tc_pr.find("w:tcBorders", NS)
        valid = valid and _is_known_child_order_valid(tc_borders, _BORDER_SIDE_ORDER)

    for tc in cells:
        tc_pr = tc.find("w:tcPr", NS)
        tc_borders = tc.find("w:tcPr/w:tcBorders", NS)
        tc_pr_orders.append(_child_order_text(tc_pr))

    return {
        "table_border_schema_order_valid": valid,
        "tblPr_child_order": _child_order_text(tbl_pr),
        "last_row_tcPr_child_orders": tc_pr_orders,
    }


def _format_bottom_border_verify_detail(
    *,
    tbl,
    cells: list,
    diagnostics: dict[str, object],
    schema_summary: dict[str, object],
    extra_parts: list[str] | None = None,
) -> str:
    tbl_bottom = tbl.find("w:tblPr/w:tblBorders/w:bottom", NS)
    last_row_bottoms = [tc.find("w:tcPr/w:tcBorders/w:bottom", NS) for tc in cells]
    parts = [
        f"tbl_bottom={_border_signature(tbl_bottom)};"
        + "last_row_tc_bottoms="
        + ("|".join(_border_signature(bottom) for bottom in last_row_bottoms) or "none"),
        "table_border_schema_order_valid="
        + ("true" if schema_summary["table_border_schema_order_valid"] else "false"),
        f"tblPr_child_order={schema_summary['tblPr_child_order']}",
        "last_row_tcPr_child_orders="
        + ("|".join(schema_summary["last_row_tcPr_child_orders"]) or "none"),
        f"last_row_physical_cell_count={diagnostics['last_row_physical_cell_count']}",
        f"last_row_grid_span_sum={diagnostics['last_row_grid_span_sum']}",
        f"last_row_vmerge_states={diagnostics['last_row_vmerge_states']}",
        f"last_row_bottom_edge_target_count={diagnostics['last_row_bottom_edge_target_count']}",
    ]
    if extra_parts:
        parts.extend(extra_parts)
    return ";".join(parts)


def _verify_data_bottom_double_black_border(tbl) -> tuple[bool, str, dict[str, object]]:
    target_cells = _last_row_bottom_edge_target_cells(tbl)
    diagnostics = _last_row_diagnostics(tbl)
    schema_summary = _schema_order_summary(tbl, target_cells)
    tbl_bottom = tbl.find("w:tblPr/w:tblBorders/w:bottom", NS)
    target_bottoms = [
        tc.find("w:tcPr/w:tcBorders/w:bottom", NS) for tc in target_cells
    ]
    verified = (
        bool(target_cells)
        and _is_double_black_border(tbl_bottom)
        and all(_is_double_black_border(bottom) for bottom in target_bottoms)
        and bool(schema_summary["table_border_schema_order_valid"])
    )
    detail = _format_bottom_border_verify_detail(
        tbl=tbl,
        cells=target_cells,
        diagnostics=diagnostics,
        schema_summary=schema_summary,
    )
    return verified, detail, {**diagnostics, **schema_summary}


def _verify_footer_terminal_bottom_none(
    tbl,
    footer_rows: list,
) -> tuple[bool, str, dict[str, object]]:
    last_footer_cells = list(footer_rows[-1].cells) if footer_rows else []
    all_footer_cells = []
    seen: set[int] = set()
    for footer_row in footer_rows:
        for tc in footer_row.cells:
            if id(tc) not in seen:
                seen.add(id(tc))
                all_footer_cells.append(tc)

    diagnostics = _last_row_diagnostics(tbl)
    all_footer_schema_summary = _schema_order_summary(tbl, all_footer_cells)
    schema_summary = _schema_order_summary(tbl, last_footer_cells)
    schema_summary["table_border_schema_order_valid"] = all_footer_schema_summary[
        "table_border_schema_order_valid"
    ]
    tbl_bottom = tbl.find("w:tblPr/w:tblBorders/w:bottom", NS)
    last_footer_bottoms = [
        tc.find("w:tcPr/w:tcBorders/w:bottom", NS) for tc in last_footer_cells
    ]
    top_footer_tops = [
        tc.find("w:tcPr/w:tcBorders/w:top", NS) for tc in footer_rows[0].cells
    ] if footer_rows else []
    internal_footer_tops = [
        tc.find("w:tcPr/w:tcBorders/w:top", NS)
        for footer_row in footer_rows[1:]
        for tc in footer_row.cells
    ]
    verified = (
        bool(last_footer_cells)
        and _is_nil_border(tbl_bottom)
        and all(_is_nil_border(bottom) for bottom in last_footer_bottoms)
        and all(_is_double_black_border(top) for top in top_footer_tops)
        and all(_is_nil_border(top) for top in internal_footer_tops)
        and bool(schema_summary["table_border_schema_order_valid"])
    )
    detail = _format_bottom_border_verify_detail(
        tbl=tbl,
        cells=last_footer_cells,
        diagnostics=diagnostics,
        schema_summary=schema_summary,
        extra_parts=[
            "footer_top_tc_tops="
            + ("|".join(_border_signature(top) for top in top_footer_tops) or "none"),
            "footer_internal_top_tc_tops="
            + ("|".join(_border_signature(top) for top in internal_footer_tops) or "none"),
        ],
    )
    return verified, detail, {**diagnostics, **schema_summary}


def _first_row_cells(tbl) -> list:
    rows = tbl.findall("w:tr", NS)
    if not rows:
        return []
    return _unique_row_cells(rows[0])


def _first_row_grid_span_sum(cells: list) -> int:
    return sum(_cell_grid_span(tc) for tc in cells)


def _first_row_tc_border(tc, side: str):
    return tc.find(f"w:tcPr/w:tcBorders/w:{side}", NS)


def _verify_first_row_single_cell_border(tbl) -> tuple[bool, str, dict[str, object]]:
    first_row_cells = _first_row_cells(tbl)
    schema_summary = _schema_order_summary(tbl, first_row_cells)
    schema_valid = bool(schema_summary["table_border_schema_order_valid"])
    if len(first_row_cells) != 1:
        detail = ";".join(
            [
                "mode=not_applicable",
                f"first_row_cell_count={len(first_row_cells)}",
                f"grid_span={_first_row_grid_span_sum(first_row_cells)}",
                f"schema_order_valid={'true' if schema_valid else 'false'}",
            ]
        )
        return False, detail, {
            "first_row_single_cell_border_mode": "not_applicable",
            "first_row_single_cell_title": False,
            "first_row_grid_span_sum": _first_row_grid_span_sum(first_row_cells),
        }

    title_cell = first_row_cells[0]
    top = _first_row_tc_border(title_cell, "top")
    left = _first_row_tc_border(title_cell, "left")
    right = _first_row_tc_border(title_cell, "right")
    bottom = _first_row_tc_border(title_cell, "bottom")
    verified = (
        _is_nil_border(top)
        and _is_nil_border(left)
        and _is_nil_border(right)
        and _is_double_black_border(bottom)
        and schema_valid
    )
    grid_span = _cell_grid_span(title_cell)
    detail = ";".join(
        [
            "mode=title_open_three_sides",
            f"top={_border_signature(top)}",
            f"left={_border_signature(left)}",
            f"right={_border_signature(right)}",
            f"bottom={_border_signature(bottom)}",
            f"grid_span={grid_span}",
            f"schema_order_valid={'true' if schema_valid else 'false'}",
        ]
    )
    return verified, detail, {
        "first_row_single_cell_border_mode": "title_open_three_sides",
        "first_row_single_cell_title": True,
        "first_row_grid_span_sum": grid_span,
    }


def _verify_table_top_border(tbl) -> tuple[bool, str, dict[str, object]]:
    first_row_cells = _first_row_cells(tbl)
    schema_summary = _schema_order_summary(tbl, first_row_cells)
    schema_valid = bool(schema_summary["table_border_schema_order_valid"])
    tbl_top = tbl.find("w:tblPr/w:tblBorders/w:top", NS)
    first_row_tops = [
        tc.find("w:tcPr/w:tcBorders/w:top", NS) for tc in first_row_cells
    ]
    first_row_single_cell_title = len(first_row_cells) == 1
    first_row_grid_span_sum = _first_row_grid_span_sum(first_row_cells)

    if not first_row_cells:
        mode = "not_applied"
        verified = False
    elif first_row_single_cell_title:
        mode = "single_title_nil"
        title_cell = first_row_cells[0]
        verified = (
            _is_double_black_border(tbl_top)
            and _is_nil_border(_first_row_tc_border(title_cell, "top"))
            and _is_nil_border(_first_row_tc_border(title_cell, "left"))
            and _is_nil_border(_first_row_tc_border(title_cell, "right"))
            and _is_double_black_border(_first_row_tc_border(title_cell, "bottom"))
            and schema_valid
        )
    else:
        mode = "data_double"
        verified = (
            _is_double_black_border(tbl_top)
            and all(_is_double_black_border(top) for top in first_row_tops)
            and schema_valid
        )

    detail = ";".join(
        [
            f"table_top_border_mode={mode}",
            f"tbl_top={_border_signature(tbl_top)}",
            "first_row_tc_tops="
            + ("|".join(_border_signature(top) for top in first_row_tops) or "none"),
            "first_row_single_cell_title="
            + ("true" if first_row_single_cell_title else "false"),
            f"first_row_grid_span_sum={first_row_grid_span_sum}",
            "table_border_schema_order_valid="
            + ("true" if schema_valid else "false"),
            f"tblPr_child_order={schema_summary['tblPr_child_order']}",
            "first_row_tcPr_child_orders="
            + ("|".join(schema_summary["last_row_tcPr_child_orders"]) or "none"),
        ]
    )
    return verified, detail, {
        "table_top_border_mode": mode,
        "table_top_border_cell_count": len(first_row_cells) if mode != "not_applied" else 0,
        "first_row_grid_span_sum": first_row_grid_span_sum,
    }


def collect_consecutive_footer_rows_from_bottom(
    tbl, stop: StopController | None = None
) -> list:
    """Collect the contiguous block of footer rows at the bottom of a table.

    Scans rows upward from the bottom; a row is a footer row when at least one
    of its (de-duplicated) cells matches a footer rule. Stops at the first row
    (including a blank row) with no match. Returns the footer rows ordered
    top-to-bottom (topmost footer row first).
    """
    rows = tbl.findall("w:tr", NS)
    footer_rows: list = []
    for row_index in range(len(rows) - 1, -1, -1):
        if stop:
            stop.check()
        cells = _unique_row_cells(rows[row_index])
        matches = []
        for tc in cells:
            text = normalize_footer_source_cell_text(tc)
            cell_type = classify_footer_cell_text(text)
            if cell_type is not None:
                matches.append((tc, cell_type, text))
        if not matches:
            break  # first non-footer row from the bottom stops the upward scan
        footer_rows.append(_FooterRow(row_index=row_index, cells=cells, matches=matches))
    footer_rows.reverse()  # top-to-bottom
    return footer_rows


def apply_footer_block_format(footer_rows: list) -> dict[str, object]:
    """Format a collected, top-to-bottom block of footer rows as a unit.

    The TOP footer row gets a black double top border across every cell (the one
    data/footer separator). Every other footer row has its top border cleared so
    no line appears between consecutive footer rows. Matched cells get 10 pt,
    their alignment, and no left/right/bottom border. The caller makes the
    final footer/data bottom-edge decision after the footer block is known.
    """
    stats: dict[str, object] = {
        "footer_block_top_border_applied": False,
        "footer_internal_top_borders_cleared": 0,
        "footer_note_cells_adjusted": 0,
        "footer_base_period_cells_adjusted": 0,
        "footer_source_cells_adjusted": 0,
        "footer_cell_matches": [],
        "footer_note_cell_matches": [],
        "footer_note_cell_debug": [],
    }
    if not footer_rows:
        return stats

    # A. Top footer row -> black double TOP border across ALL its cells, so the
    # data/footer separator spans the whole table width (not only matched cells).
    for tc in footer_rows[0].cells:
        set_border_double_black(_get_or_add_tc_borders(tc), "top")
    stats["footer_block_top_border_applied"] = True

    # B. Every other footer row -> clear the TOP border across ALL its cells, so
    # no horizontal line appears between consecutive footer rows.
    for footer_row in footer_rows[1:]:
        for tc in footer_row.cells:
            set_border_nil(_get_or_add_tc_borders(tc), "top")
        stats["footer_internal_top_borders_cleared"] += 1

    # C. Matched cells in every footer row -> 10 pt, alignment, no left/right/
    # bottom border. The TOP edge is owned by A/B and intentionally left as-is.
    for footer_row in footer_rows:
        row_types = []
        for tc, cell_type, text in footer_row.matches:
            _set_runs_font_size(tc, FOOTER_SOURCE_NOTE_FONT_SIZE_HALF_POINTS)
            _set_paragraph_alignment(tc, FOOTER_ALIGNMENT_BY_TYPE[cell_type])
            borders = _get_or_add_tc_borders(tc)
            set_border_nil(borders, "left")
            set_border_nil(borders, "right")
            set_border_nil(borders, "bottom")
            stats[f"footer_{cell_type}_cells_adjusted"] += 1
            row_types.append(cell_type)
            stats["footer_note_cell_matches"].append(cell_type)
            stats["footer_note_cell_debug"].append(f"{cell_type}: {text[:50]}")
        stats["footer_cell_matches"].append(",".join(row_types))
    return stats


def apply_table_footer_source_format(tbl, stop: StopController | None = None) -> dict[str, object]:
    """Apply the 「表格最後一列說明格式化」 format to one table.

    Formats matching last-row cells (基期：/資料來源：/註記) plus the table
    frame and a single-cell first row. Never moves, deletes or inserts any
    cell, row or paragraph. Returns a dict describing what was changed.
    """
    result: dict[str, object] = {
        "outer_double_border_applied": False,
        "first_row_single_cell_border_adjusted": False,
        "first_row_single_cell_title": False,
        "first_row_single_cell_border_mode": "not_applicable",
        "first_row_single_cell_border_xml_verified": False,
        "first_row_single_cell_border_verify_detail": "",
        "table_top_border_mode": "not_applied",
        "table_top_border_cell_count": 0,
        "table_top_border_xml_verified": False,
        "table_top_border_verify_detail": "",
        "footer_rows_detected": False,
        "footer_row_count": 0,
        "footer_top_row_index": None,
        "footer_cell_matches": [],
        "footer_block_top_border_applied": False,
        "footer_internal_top_borders_cleared": 0,
        "footer_note_cells_adjusted": 0,
        "footer_base_period_cells_adjusted": 0,
        "footer_source_cells_adjusted": 0,
        "footer_note_cell_matches": [],
        "footer_note_cell_debug": [],
        "table_bottom_border_mode": "not_applied",
        "table_bottom_border_cell_count": 0,
        "table_bottom_border_xml_verified": False,
        "table_bottom_border_verify_detail": "",
        "table_bottom_double_border_applied": False,
        "table_bottom_double_border_cell_count": 0,
        "table_bottom_double_border_xml_verified": False,
        "table_bottom_double_border_verify_detail": "",
        "footer_terminal_bottom_none_applied": False,
        "footer_terminal_bottom_none_cell_count": 0,
        "data_rows_outer_left_double_applied": False,
        "data_rows_outer_right_double_applied": False,
        "data_rows_outer_left_target_count": 0,
        "data_rows_outer_right_target_count": 0,
        "data_rows_outer_left_vmerge_owner_target_count": 0,
        "data_rows_outer_right_vmerge_owner_target_count": 0,
        "footer_rows_outer_left_none_applied": False,
        "footer_rows_outer_right_none_applied": False,
        "footer_rows_outer_left_target_count": 0,
        "footer_rows_outer_right_target_count": 0,
        "outer_vertical_border_policy_xml_verified": False,
        "outer_vertical_border_policy_verify_detail": "",
        "last_row_physical_cell_count": 0,
        "last_row_grid_span_sum": 0,
        "last_row_vmerge_states": "none",
        "last_row_bottom_edge_target_count": 0,
        "table_border_schema_order_valid": False,
        "tblPr_child_order": "",
        "last_row_tcPr_child_orders": [],
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

    # 3. First-row top border: title rows are intentionally open on three
    # sides; normal first rows get direct top borders so existing cell borders
    # cannot override the rendered table top in Word.
    first_row_cells = _unique_row_cells(rows[0])
    if len(first_row_cells) == 1:
        _apply_first_row_single_cell_borders(first_row_cells[0])
        result["first_row_single_cell_border_adjusted"] = True
        result["first_row_single_cell_title"] = True
        result["table_top_border_mode"] = "single_title_nil"
        result["table_top_border_cell_count"] = 1
    else:
        top_cell_count = _apply_rendered_top_double_black_border(tbl)
        result["table_top_border_mode"] = (
            "data_double" if top_cell_count > 0 else "not_applied"
        )
        result["table_top_border_cell_count"] = top_cell_count

    # 4. Footer block: collect the bottom contiguous footer rows, then format the
    # block as a unit (separate scan from apply, so the borders are decided per
    # block rather than per cell -- only the top row gets the separator line).
    footer_rows = collect_consecutive_footer_rows_from_bottom(tbl, stop=stop)
    if footer_rows:
        result["footer_rows_detected"] = True
        result["footer_row_count"] = len(footer_rows)
        result["footer_top_row_index"] = footer_rows[0].row_index
        result.update(apply_footer_block_format(footer_rows))
        none_cell_count = _apply_footer_terminal_bottom_none(tbl, footer_rows)
        result["table_bottom_border_mode"] = "footer_none"
        result["table_bottom_border_cell_count"] = none_cell_count
        result["footer_terminal_bottom_none_applied"] = none_cell_count > 0
        result["footer_terminal_bottom_none_cell_count"] = none_cell_count
        result["table_bottom_double_border_applied"] = False
        result["table_bottom_double_border_cell_count"] = 0
    else:
        bottom_cell_count = _apply_rendered_bottom_double_black_border(tbl)
        result["table_bottom_border_mode"] = (
            "data_double" if bottom_cell_count > 0 else "not_applied"
        )
        result["table_bottom_border_cell_count"] = bottom_cell_count
        result["table_bottom_double_border_applied"] = bottom_cell_count > 0
        result["table_bottom_double_border_cell_count"] = bottom_cell_count

    result.update(_apply_data_rows_outer_vertical_double_borders(tbl, footer_rows))
    result.update(_apply_footer_rows_outer_vertical_none_borders(tbl, footer_rows))

    _normalize_table_border_related_schema_order(tbl)

    if footer_rows:
        verified, detail, diagnostics = _verify_footer_terminal_bottom_none(
            tbl,
            footer_rows,
        )
        detail = f"table_bottom_border_mode=footer_none;{detail}"
        result.update(diagnostics)
        result["table_bottom_border_xml_verified"] = verified
        result["table_bottom_border_verify_detail"] = detail
        result["table_bottom_double_border_xml_verified"] = False
        result["table_bottom_double_border_verify_detail"] = detail
    else:
        verified, detail, diagnostics = _verify_data_bottom_double_black_border(tbl)
        detail = f"table_bottom_border_mode=data_double;{detail}"
        result.update(diagnostics)
        result["table_bottom_border_xml_verified"] = verified
        result["table_bottom_border_verify_detail"] = detail
        result["table_bottom_double_border_xml_verified"] = verified
        result["table_bottom_double_border_verify_detail"] = detail

    result.update(_verify_outer_vertical_border_policy(tbl, footer_rows))

    (
        first_row_verified,
        first_row_detail,
        first_row_diagnostics,
    ) = _verify_first_row_single_cell_border(tbl)
    result.update(first_row_diagnostics)
    result["first_row_single_cell_border_xml_verified"] = first_row_verified
    result["first_row_single_cell_border_verify_detail"] = first_row_detail

    top_verified, top_detail, top_diagnostics = _verify_table_top_border(tbl)
    result.update(top_diagnostics)
    result["table_top_border_xml_verified"] = top_verified
    result["table_top_border_verify_detail"] = top_detail

    return result
