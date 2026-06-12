from __future__ import annotations

import json
import re
from pathlib import Path

from .indent_settings import get_indent_settings_path

TABLE_COLOR_SETTINGS_KEY = "table_color_settings"

_HEX_RE = re.compile(r"[0-9A-F]{6}")


def get_table_color_settings_path() -> Path:
    # Shares indent_defaults.json so a portable EXE keeps every setting in
    # one file next to the executable.
    return get_indent_settings_path()


def built_in_table_color_settings() -> dict[str, object]:
    return {
        "keep_colors": ["DDEBF7"],
        "gray_colors": ["BFBFBF", "C0C0C0", "A6A6A6", "808080"],
        "gray_target": "D9D9D9",
        "special_color_skip_colors": [],
    }


def normalize_hex_color(value: object) -> str:
    if not isinstance(value, str):
        raise ValueError(f"顏色必須是 HEX 字串：{value!r}")
    text = value.strip().replace("#", "").upper()
    if not _HEX_RE.fullmatch(text):
        raise ValueError(f"無效的 HEX 色碼：{value!r}（請輸入 6 碼，例如 DDEBF7）")
    return text


def parse_color_list_text(text: str) -> list[str]:
    colors: list[str] = []
    for token in re.split(r"[,\s;]+", text or ""):
        token = token.strip()
        if not token:
            continue
        normalized = normalize_hex_color(token)
        if normalized not in colors:
            colors.append(normalized)
    return colors


def format_color_list_text(colors: list[str] | tuple[str, ...]) -> str:
    return "\n".join(colors)


def _normalize_color_list(key: str, value: object) -> list[str]:
    if isinstance(value, str):
        return parse_color_list_text(value)
    if isinstance(value, (list, tuple)):
        colors: list[str] = []
        for item in value:
            normalized = normalize_hex_color(item)
            if normalized not in colors:
                colors.append(normalized)
        return colors
    raise ValueError(f"{key} 必須是 HEX 色碼清單")


def normalize_table_color_settings(data: dict | None) -> dict[str, object]:
    normalized = built_in_table_color_settings()
    if data is None:
        return normalized
    if not isinstance(data, dict):
        raise ValueError("表格顏色設定必須是 JSON 物件")

    for key in ("keep_colors", "gray_colors", "special_color_skip_colors"):
        if key in data:
            normalized[key] = _normalize_color_list(key, data[key])
    if "gray_target" in data:
        normalized["gray_target"] = normalize_hex_color(data["gray_target"])
    return normalized


def load_saved_table_color_settings(path: str | Path | None = None) -> dict[str, object]:
    settings_path = Path(path) if path is not None else get_table_color_settings_path()
    if not settings_path.exists():
        return built_in_table_color_settings()

    data = json.loads(settings_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("共用設定檔必須是 JSON 物件")
    return normalize_table_color_settings(data.get(TABLE_COLOR_SETTINGS_KEY))


def save_table_color_settings(settings: dict, path: str | Path | None = None) -> Path:
    normalized = normalize_table_color_settings(settings)
    settings_path = Path(path) if path is not None else get_table_color_settings_path()
    data = {}
    if settings_path.exists():
        loaded = json.loads(settings_path.read_text(encoding="utf-8"))
        if not isinstance(loaded, dict):
            raise ValueError("共用設定檔必須是 JSON 物件")
        data = loaded
    if "body" in data or "preface" in data:
        data = {"indent_settings": data}
    data[TABLE_COLOR_SETTINGS_KEY] = normalized
    settings_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return settings_path


def current_table_color_settings(path: str | Path | None = None) -> dict[str, object]:
    try:
        return load_saved_table_color_settings(path)
    except Exception:
        return built_in_table_color_settings()
