from __future__ import annotations

DEFAULT_SUFFIX = "_fixed"
DEFAULT_GRAY = "D9D9D9"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

FINANCIAL_NUM = "壹貳參肆伍陸柒捌玖拾"
SIMPLE_NUM = "一二三四五六七八九十"

POINTS_PER_CM = 28.3464567

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


def make_outline_indent_spec(number_start_cm: float, hanging_cm: float, body_left_cm: float) -> dict[str, str]:
    return {
        "left": cm_to_twips(number_start_cm + hanging_cm),
        "hanging": cm_to_twips(hanging_cm),
        "number_start": cm_to_twips(number_start_cm),
        "body_left": cm_to_twips(body_left_cm),
    }


TEMPLATE_OUTLINE_INDENTS = {
    0: make_outline_indent_spec(-0.04, 1.27, 1.23),
    1: make_outline_indent_spec(0.73, 1.13, 1.86),
    2: make_outline_indent_spec(1.51, 1.48, 2.99),
    3: make_outline_indent_spec(3.49, 0.50, 3.99),
    4: make_outline_indent_spec(3.74, 1.23, 4.96),
    5: make_outline_indent_spec(5.45, 0.50, 5.95),
    6: make_outline_indent_spec(4.70, 1.23, 5.94),
    7: make_outline_indent_spec(5.94, 0.49, 6.85),
    8: make_outline_indent_spec(7.72, 1.24, 8.96),
}


PREFACE_OUTLINE_INDENTS = {
    0: make_outline_indent_spec(-0.04, 1.14, 1.09),  # 一、
    1: make_outline_indent_spec(0.73, 1.48, 2.21),  # （一）
    2: make_outline_indent_spec(2.21, 0.49, 3.45),  # 1.
    3: make_outline_indent_spec(2.45, 1.24, 4.44),  # （1）
    4: make_outline_indent_spec(4.43, 0.50, 4.92),  # A.
    5: make_outline_indent_spec(4.67, 1.24, 5.90),  # （A）
    6: make_outline_indent_spec(6.39, 0.50, 6.85),  # a.
    7: make_outline_indent_spec(7.72, 1.24, 8.96),  # （a）
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
