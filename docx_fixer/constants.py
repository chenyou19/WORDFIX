from __future__ import annotations

DEFAULT_SUFFIX = "_已修改"
DEFAULT_GRAY = "D9D9D9"

CONVERT_TO_GRAY_FILLS = {"BFBFBF", "A6A6A6", "808080"}
KEEP_COLOR_FILLS = {"F2F2F2", DEFAULT_GRAY}

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

FINANCIAL_NUM = "壹貳參肆伍陸柒捌玖拾佰仟萬"
SIMPLE_NUM = "一二三四五六七八九十百千萬"

POINTS_PER_CM = 28.3464567

# User-visible level 2 ("一、") is internal level 1.
OUTLINE_LEVEL_FONT_SIZE_PT = {
    1: 16.0,
}


def cm_to_points(cm: float) -> float:
    return cm * POINTS_PER_CM


def cm_to_twips(cm: float) -> str:
    return str(round(cm_to_points(cm) * 20))


TEMPLATE_OUTLINE_INDENTS = {
    # 壹、
    0: {
        "left": cm_to_twips(1.11),
        "hanging": cm_to_twips(1.15),
        "number_start": cm_to_twips(-0.04),
    },
    # 一、
    1: {
        "left": cm_to_twips(1.8),
        "hanging": cm_to_twips(1.8 - 0.69),
        "number_start": cm_to_twips(0.69),
    },
    # （一）
    2: {
        "left": cm_to_twips(2.32),
        "hanging": cm_to_twips(2.32 - 1.32),
        "number_start": cm_to_twips(1.32),
    },
    # 1.
    3: {
        "left": cm_to_twips(3.79),
        "hanging": cm_to_twips(3.79 - 3.05),
        "number_start": cm_to_twips(3.05),
    },
    # （1）
    4: {
        "left": cm_to_twips(4.76),
        "hanging": cm_to_twips(4.76 - 3.53),
        "number_start": cm_to_twips(3.53),
    },
    # A.
    5: {
        "left": cm_to_twips(5.27),
        "hanging": cm_to_twips(5.27 - 4.52),
        "number_start": cm_to_twips(4.52),
    },
    # （A）
    6: {
        "left": cm_to_twips(6.26),
        "hanging": cm_to_twips(6.26 - 5.02),
        "number_start": cm_to_twips(5.02),
    },
    # a.
    7: {
        "left": cm_to_twips(6.96),
        "hanging": cm_to_twips(6.96 - 6.2),
        "number_start": cm_to_twips(6.2),
    },
    # （a）
    8: {
        "left": cm_to_twips(8.96),
        "hanging": cm_to_twips(1.24),
        "number_start": cm_to_twips(7.72),
    },
}


PREFACE_OUTLINE_INDENTS = {
    # 一、
    0: {"left": cm_to_twips(1.11), "hanging": cm_to_twips(1.15), "number_start": cm_to_twips(-0.04)},
    # （一）
    1: {"left": cm_to_twips(1.54), "hanging": cm_to_twips(0.85), "number_start": cm_to_twips(0.69)},
    # 1.
    2: {"left": cm_to_twips(3.01), "hanging": cm_to_twips(1.01), "number_start": cm_to_twips(2.00)},
    # （1）
    3: {"left": cm_to_twips(4.02), "hanging": cm_to_twips(1.72), "number_start": cm_to_twips(2.30)},
}


def validate_template_outline_indents(tolerance: int = 1) -> bool:
    for level, spec in TEMPLATE_OUTLINE_INDENTS.items():
        cur_left = int(spec["left"])
        cur_hanging = int(spec["hanging"])
        cur_number_start = cur_left - cur_hanging
        expected_number_start = int(spec["number_start"])

        assert abs(cur_number_start - expected_number_start) <= tolerance, (
            f"level {level} 不符合縮排規則："
            f"編號起點 {cur_number_start} != 指定編號起點 {expected_number_start}"
        )

    return True
