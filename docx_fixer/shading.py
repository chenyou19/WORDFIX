from __future__ import annotations

import re

from .constants import CONVERT_TO_GRAY_FILLS, DEFAULT_GRAY, KEEP_COLOR_FILLS
from .xml_utils import qn

def normalize_fill_hex(value: str | None) -> str | None:
    """
    將 w:fill 色碼標準化成 6 碼大寫 HEX。
    無色彩、auto、none、格式不正確者回傳 None。
    """
    if not value:
        return None

    value = value.strip().replace("#", "").upper()

    if value in {"AUTO", "NONE"} or len(value) != 6:
        return None

    if not re.fullmatch(r"[0-9A-F]{6}", value):
        return None

    return value


def is_no_color_shading(shd) -> bool:
    """
    判斷是否可視為「無色彩」。
    沒有 fill、fill=auto、fill=none，且沒有 themeFill/themeColor 時，視為無色彩。
    """
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
    """指定色碼改成指定灰色。"""
    shd.set(qn("val"), "clear")
    shd.set(qn("fill"), gray_hex)
    shd.set(qn("color"), "auto")
    remove_theme_color_attrs(shd)


def fix_shading_to_no_color(shd) -> None:
    """其他顏色改成無色彩。"""
    shd.set(qn("val"), "clear")
    shd.set(qn("fill"), "auto")
    shd.set(qn("color"), "auto")
    remove_theme_color_attrs(shd)


def get_shading_action(shd) -> str:
    """
    回傳顏色處理動作：
    - keep：無色彩／F2F2F2／DEFAULT_GRAY，保持不動
    - gray：BFBFBF、A6A6A6、808080，改成 DEFAULT_GRAY
    - clear：其他顏色，改成無色彩
    """
    if is_no_color_shading(shd):
        return "keep"

    fill_hex = normalize_fill_hex(shd.get(qn("fill")))

    # 有 themeFill/themeColor 但沒有可判斷的 HEX 時，當作其他顏色清掉。
    if fill_hex is None:
        return "clear"

    if fill_hex in CONVERT_TO_GRAY_FILLS:
        return "gray"

    if fill_hex in KEEP_COLOR_FILLS:
        return "keep"

    return "clear"
