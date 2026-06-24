from __future__ import annotations

import math

DEFAULT_SUFFIX = "_fixed"
DEFAULT_GRAY = "D9D9D9"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

FINANCIAL_NUM = "壹貳參肆伍陸柒捌玖拾"
SIMPLE_NUM = "一二三四五六七八九十"

POINTS_PER_CM = 28.3464567
DEFAULT_HEADING_TEXT_START_OFFSET_CM = 0.85

# Central heading font-size rule, keyed by the detected internal outline level
# (0-8). This is the single source of truth shared by both heading writers:
#   * outline.apply_outline_level_font_size -> paragraph visible text runs
#   * numbering.apply_numbering_level_outline_format -> numbering.xml level rPr
# so manual-text headings and auto-numbered headings stay in sync. Both writers
# emit w:sz and w:szCs together (half-points = round(pt * 2)).
#
# Level mapping (detected level -> numbering marker):
#   0: 壹、 (financial)   1: 一、 (simple)   2: （一）   3: 1.   4: （1）
#   5: A.   6: （A）   7: a.   8: （a）
#
# Level 0 (壹、) is intentionally absent: it keeps its original size and is never
# forced. Level 1 (一、) keeps its existing 16 pt rule. Both must not be changed
# to 14 pt by this feature. Levels 2 and deeper are forced to 14 pt.
# User-visible level 2 is internal level 1.
OUTLINE_LEVEL_FONT_SIZE_PT = {
    1: 16.0,
    2: 14.0,
    3: 14.0,
    4: 14.0,
    5: 14.0,
    6: 14.0,
    7: 14.0,
    8: 14.0,
}


def cm_to_points(cm: float) -> float:
    return cm * POINTS_PER_CM


def cm_to_twips(cm: float) -> str:
    return str(round(cm_to_points(cm) * 20))


def make_outline_indent_spec(
    *,
    number_start_cm: float,
    text_indent_cm: float,
    tab_stop_cm: float,
    body_left_cm: float,
) -> dict[str, str]:
    for value, name in (
        (number_start_cm, "number_start_cm"),
        (text_indent_cm, "text_indent_cm"),
        (tab_stop_cm, "tab_stop_cm"),
        (body_left_cm, "body_left_cm"),
    ):
        if not math.isfinite(float(value)):
            raise ValueError(f"{name} must be finite.")
    if text_indent_cm <= number_start_cm:
        raise ValueError("文字縮排必須大於標號起點。")
    return {
        "left": cm_to_twips(text_indent_cm),
        "hanging": cm_to_twips(text_indent_cm - number_start_cm),
        "number_start": cm_to_twips(number_start_cm),
        "heading_text_start": cm_to_twips(tab_stop_cm),
        "body_left": cm_to_twips(body_left_cm),
    }


TEMPLATE_OUTLINE_INDENTS = {
    0: make_outline_indent_spec(#壹
        number_start_cm=-0.04,
        text_indent_cm=1.23,
        tab_stop_cm=2.08,
        body_left_cm=1.23,
    ),
    1: make_outline_indent_spec(#一
        number_start_cm=0.73,
        text_indent_cm=1.86,
        tab_stop_cm=2.71,
        body_left_cm=1.86,
    ),
    2: make_outline_indent_spec(#(一)
        number_start_cm=1.51,
        text_indent_cm=2.99,
        tab_stop_cm=3.84,
        body_left_cm=2.99,
    ),
    3: make_outline_indent_spec(#1.
        number_start_cm=3.11,
        text_indent_cm=3.98,
        tab_stop_cm=4.84,
        body_left_cm=3.98,
    ),
    4: make_outline_indent_spec(#(1)
        number_start_cm=3.74,
        text_indent_cm=4.96,
        tab_stop_cm=5.81,
        body_left_cm=4.98,
    ),
    5: make_outline_indent_spec(#A.
        number_start_cm=5.16,
        text_indent_cm=5.97,
        tab_stop_cm=6.80,
        body_left_cm=5.97,
    ),
    6: make_outline_indent_spec(#(A)
        number_start_cm=5.93,
        text_indent_cm=6.96,
        tab_stop_cm=8.6,
        body_left_cm=6.96,
    ),
    7: make_outline_indent_spec(#a.
        number_start_cm=7.04,
        text_indent_cm=8.1,
        tab_stop_cm=8.72,
        body_left_cm=8.1,
    ),
    8: make_outline_indent_spec(#(a)
        number_start_cm=7.72,
        text_indent_cm=8.6,
        tab_stop_cm=9.81,
        body_left_cm=8.6,
    ),
}


PREFACE_OUTLINE_INDENTS = {
    0: make_outline_indent_spec(  # 一、
        number_start_cm=-0.04,
        text_indent_cm=1.10,
        tab_stop_cm=1.94,
        body_left_cm=1.09,
    ),
    1: make_outline_indent_spec(  # （一）
        number_start_cm=0.73,
        text_indent_cm=2.21,
        tab_stop_cm=3.06,
        body_left_cm=2.21,
    ),
    2: make_outline_indent_spec(  # 1.
        number_start_cm=2.21,
        text_indent_cm=2.70,
        tab_stop_cm=4.30,
        body_left_cm=3.45,
    ),
    3: make_outline_indent_spec(  # （1）
        number_start_cm=2.45,
        text_indent_cm=3.69,
        tab_stop_cm=5.29,
        body_left_cm=4.44,
    ),
    4: make_outline_indent_spec(  # A.
        number_start_cm=4.43,
        text_indent_cm=4.93,
        tab_stop_cm=5.77,
        body_left_cm=4.92,
    ),
    5: make_outline_indent_spec(  # （A）
        number_start_cm=4.67,
        text_indent_cm=5.91,
        tab_stop_cm=6.75,
        body_left_cm=5.90,
    ),
    6: make_outline_indent_spec(  # a.
        number_start_cm=6.39,
        text_indent_cm=6.89,
        tab_stop_cm=7.70,
        body_left_cm=6.85,
    ),
    7: make_outline_indent_spec(  # （a）
        number_start_cm=7.72,
        text_indent_cm=8.96,
        tab_stop_cm=9.81,
        body_left_cm=8.96,
    ),
}

BUILTIN_TEMPLATE_OUTLINE_INDENTS = {
    level: dict(spec)
    for level, spec in TEMPLATE_OUTLINE_INDENTS.items()
}

BUILTIN_PREFACE_OUTLINE_INDENTS = {
    level: dict(spec)
    for level, spec in PREFACE_OUTLINE_INDENTS.items()
}


def validate_template_outline_indents(tolerance: int = 1) -> bool:
    for level, spec in TEMPLATE_OUTLINE_INDENTS.items():
        cur_left = int(spec["left"])
        cur_hanging = int(spec["hanging"])
        cur_number_start = cur_left - cur_hanging
        expected_number_start = int(spec["number_start"])

        assert abs(cur_number_start - expected_number_start) <= tolerance, (
            f"level {level} number start mismatch: "
            f"calculated {cur_number_start} != expected {expected_number_start}"
        )

    return True
