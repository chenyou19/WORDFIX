from __future__ import annotations

DEFAULT_SUFFIX = "_fixed"
DEFAULT_GRAY = "D9D9D9"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

FINANCIAL_NUM = "壹貳參肆伍陸柒捌玖拾"
SIMPLE_NUM = "一二三四五六七八九十"

POINTS_PER_CM = 28.3464567

# User-visible level 2 is internal level 1.
OUTLINE_LEVEL_FONT_SIZE_PT = {
    1: 16.0,
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
    1: make_outline_indent_spec(0.70, 1.12, 1.83),
    2: make_outline_indent_spec(1.47, 1.48, 2.96),
    3: make_outline_indent_spec(2.96, 0.50, 3.45),
    4: make_outline_indent_spec(3.21, 1.23, 4.44),
    5: make_outline_indent_spec(4.45, 0.50, 4.94),
    6: make_outline_indent_spec(4.70, 1.23, 5.94),
    7: make_outline_indent_spec(5.94, 0.49, 6.85),
    8: make_outline_indent_spec(7.72, 1.24, 8.96),
}


PREFACE_OUTLINE_INDENTS = {
    level: dict(TEMPLATE_OUTLINE_INDENTS[level])
    for level in range(8)
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
