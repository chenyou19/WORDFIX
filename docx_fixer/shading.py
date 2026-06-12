from __future__ import annotations

import re

from .constants import DEFAULT_GRAY
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


def hex_to_rgb(fill_hex: str | None) -> tuple[int, int, int] | None:
    normalized = normalize_fill_hex(fill_hex)
    if normalized is None:
        return None
    return (
        int(normalized[0:2], 16),
        int(normalized[2:4], 16),
        int(normalized[4:6], 16),
    )


def is_gray_hex(fill_hex: str | None) -> bool:
    rgb = hex_to_rgb(fill_hex)
    if rgb is None:
        return False
    red, green, blue = rgb
    return red == green == blue


def is_darker_than_default_gray(fill_hex: str | None, default_gray: str = DEFAULT_GRAY) -> bool:
    fill_rgb = hex_to_rgb(fill_hex)
    default_rgb = hex_to_rgb(default_gray)
    if fill_rgb is None or default_rgb is None:
        return False
    return fill_rgb[0] < default_rgb[0]


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


def get_shading_decision(
    shd,
    *,
    keep_colors: tuple[str, ...] | list[str] = (),
    gray_colors: tuple[str, ...] | list[str] = (),
    gray_target: str = DEFAULT_GRAY,
) -> dict[str, str | None]:
    raw_attrs = {attr: shd.get(qn(attr)) for attr in SHADING_DEBUG_ATTRS}
    fill_hex = normalize_fill_hex(raw_attrs["fill"])
    has_theme_color = any(raw_attrs[attr] is not None for attr in ["themeFill", "themeColor"])

    matched_keep_color: str | None = None
    matched_gray_color: str | None = None

    if is_no_color_shading(shd):
        action = "keep"
        reason = "no color shading"
    elif fill_hex is not None and fill_hex in keep_colors:
        matched_keep_color = fill_hex
        action = "keep"
        reason = "matched keep color list"
    elif fill_hex is not None and fill_hex in gray_colors:
        matched_gray_color = fill_hex
        action = "gray"
        reason = "matched gray color list"
    elif is_gray_hex(fill_hex):
        if is_darker_than_default_gray(fill_hex, gray_target):
            action = "gray"
            reason = "gray hex darker than default gray"
        else:
            action = "keep"
            reason = "gray hex lighter/equal default gray"
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
        "matched_keep_color": matched_keep_color,
        "matched_gray_color": matched_gray_color,
        "gray_target": gray_target,
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
        f"matched_keep_color={decision.get('matched_keep_color')}; "
        f"matched_gray_color={decision.get('matched_gray_color')}; "
        f"gray_target={decision.get('gray_target')}; "
        f"final_action={decision['action']}; "
        f"reason={decision['reason']}"
    )


def get_shading_action(
    shd,
    *,
    keep_colors: tuple[str, ...] | list[str] = (),
    gray_colors: tuple[str, ...] | list[str] = (),
    gray_target: str = DEFAULT_GRAY,
) -> str:
    return str(
        get_shading_decision(
            shd,
            keep_colors=keep_colors,
            gray_colors=gray_colors,
            gray_target=gray_target,
        )["action"]
    )
