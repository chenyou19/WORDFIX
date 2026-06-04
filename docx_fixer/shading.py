from __future__ import annotations

import re

from .constants import CONVERT_TO_GRAY_FILLS, DEFAULT_GRAY, KEEP_COLOR_FILLS
from .xml_utils import qn

SHADING_DEBUG_ATTRS = [
    "fill",
    "val",
    "color",
    "themeFill",
    "themeColor",
    "themeFillTint",
    "themeFillShade",
]


def normalize_fill_hex(value: str | None) -> str | None:
    if not value:
        return None

    value = value.strip().replace("#", "").upper()

    if value in {"AUTO", "NONE"} or len(value) != 6:
        return None

    if not re.fullmatch(r"[0-9A-F]{6}", value):
        return None

    return value


def is_no_color_shading(shd) -> bool:
    fill = shd.get(qn("fill"))

    has_theme_color = any(
        shd.get(qn(attr)) is not None
        for attr in ["themeFill", "themeColor"]
    )

    if has_theme_color:
        return False

    if fill is None:
        return True

    fill = fill.strip().upper()
    return fill in {"", "AUTO", "NONE"}


def remove_theme_color_attrs(shd) -> None:
    for attr in [
        "themeFill", "themeFillTint", "themeFillShade",
        "themeColor", "themeTint", "themeShade",
    ]:
        shd.attrib.pop(qn(attr), None)


def fix_shading_to_gray(shd, gray_hex: str = DEFAULT_GRAY) -> None:
    shd.set(qn("val"), "clear")
    shd.set(qn("fill"), gray_hex)
    shd.set(qn("color"), "auto")
    remove_theme_color_attrs(shd)


def fix_shading_to_no_color(shd) -> None:
    shd.set(qn("val"), "clear")
    shd.set(qn("fill"), "auto")
    shd.set(qn("color"), "auto")
    remove_theme_color_attrs(shd)


def get_shading_decision(shd) -> dict[str, str | None]:
    raw_attrs = {attr: shd.get(qn(attr)) for attr in SHADING_DEBUG_ATTRS}
    fill_hex = normalize_fill_hex(raw_attrs["fill"])
    has_theme_color = any(raw_attrs[attr] is not None for attr in ["themeFill", "themeColor"])

    if is_no_color_shading(shd):
        action = "keep"
        reason = "no color shading"
    elif fill_hex in CONVERT_TO_GRAY_FILLS:
        action = "gray"
        reason = "fill_hex in convert-to-gray list"
    elif fill_hex in KEEP_COLOR_FILLS:
        action = "keep"
        reason = "fill_hex in keep list"
    elif has_theme_color and fill_hex is None:
        action = "keep"
        reason = "theme color unresolved"
    elif fill_hex is None:
        action = "keep"
        reason = "fill hex unavailable"
    else:
        action = "clear"
        reason = "explicit non-gray hex color"

    return {
        "raw_fill": raw_attrs["fill"],
        "raw_val": raw_attrs["val"],
        "raw_color": raw_attrs["color"],
        "raw_themeFill": raw_attrs["themeFill"],
        "raw_themeColor": raw_attrs["themeColor"],
        "raw_themeFillTint": raw_attrs["themeFillTint"],
        "raw_themeFillShade": raw_attrs["themeFillShade"],
        "normalized_fill_hex": fill_hex,
        "action": action,
        "reason": reason,
    }


def format_shading_decision(decision: dict[str, str | None]) -> str:
    return (
        f"raw_fill={decision['raw_fill']}; "
        f"raw_val={decision['raw_val']}; "
        f"raw_color={decision['raw_color']}; "
        f"raw_themeFill={decision['raw_themeFill']}; "
        f"raw_themeColor={decision['raw_themeColor']}; "
        f"raw_themeFillTint={decision['raw_themeFillTint']}; "
        f"raw_themeFillShade={decision['raw_themeFillShade']}; "
        f"normalized_fill_hex={decision['normalized_fill_hex']}; "
        f"final_action={decision['action']}; "
        f"reason={decision['reason']}"
    )


def get_shading_action(shd) -> str:
    return str(get_shading_decision(shd)["action"])
