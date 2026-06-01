from __future__ import annotations

from .constants import NS
from .models import ProcessOptions
from .shading import fix_shading_to_gray, fix_shading_to_no_color, get_shading_action
from .stop_controller import StopController
from .xml_utils import get_or_add, qn

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


def apply_table_format(tbl, stop: StopController | None = None) -> None:
    tblPr = get_or_add(tbl, "tblPr", first=True)

    # 表格置中
    jc = get_or_add(tblPr, "jc")
    jc.set(qn("val"), "center")

    # 表格寬度 100%
    tblW = get_or_add(tblPr, "tblW")
    tblW.set(qn("type"), "pct")
    tblW.set(qn("w"), "5000")

    # 固定欄寬
    tblLayout = get_or_add(tblPr, "tblLayout")
    tblLayout.set(qn("type"), "fixed")

    # 外框線雙黑線，內框線不動
    tblBorders = get_or_add(tblPr, "tblBorders")
    for border_name in ["top", "bottom", "left", "right"]:
        border = get_or_add(tblBorders, border_name)
        border.set(qn("val"), "double")
        border.set(qn("sz"), "4")       # 0.5 pt
        border.set(qn("space"), "0")
        border.set(qn("color"), "000000")

    # 列高：0.6 cm 約等於 340 twips
    for tr in tbl.xpath(".//w:tr", namespaces=NS):
        if stop:
            stop.check()
        trPr = get_or_add(tr, "trPr", first=True)
        trHeight = get_or_add(trPr, "trHeight")
        trHeight.set(qn("val"), "340")
        trHeight.set(qn("hRule"), "atLeast")

    # 儲存格垂直置中
    for tc in tbl.xpath(".//w:tc", namespaces=NS):
        if stop:
            stop.check()
        tcPr = get_or_add(tc, "tcPr", first=True)
        vAlign = get_or_add(tcPr, "vAlign")
        vAlign.set(qn("val"), "center")

    # 表格內段落置中、段前段後 0、單行間距
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


def apply_table_color(tbl, stop: StopController | None = None) -> tuple[int, int]:
    """
    顏色規則：
    無色彩保持；BFBFBF／A6A6A6／808080 改成 D9D9D9；
    F2F2F2 與 D9D9D9 保持；其他顏色改成無色彩。
    回傳：(改成灰的數量, 改成無色彩的數量)
    """
    changed_to_gray = 0
    cleared_colors = 0

    for tc in tbl.xpath(".//w:tc", namespaces=NS):
        if stop:
            stop.check()

        tcPr = tc.find("w:tcPr", NS)
        if tcPr is None:
            continue

        for shd in tcPr.findall("w:shd", NS):
            action = get_shading_action(shd)

            if action == "gray":
                fix_shading_to_gray(shd)
                changed_to_gray += 1
            elif action == "clear":
                fix_shading_to_no_color(shd)
                cleared_colors += 1

    return changed_to_gray, cleared_colors


def process_table(
    tbl,
    options: ProcessOptions,
    stop: StopController | None = None,
    *,
    special_layout: bool = False,
) -> tuple[int, int]:
    changed_to_gray = 0
    cleared_colors = 0

    if options.fix_table_layout:
        if special_layout:
            apply_autofit_contents_right_format(tbl, stop=stop)
        else:
            apply_table_format(tbl, stop=stop)

    if options.fix_color:
        changed_to_gray, cleared_colors = apply_table_color(tbl, stop=stop)

    return changed_to_gray, cleared_colors
