from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .indent_settings import get_indent_settings_path

GUI_DEFAULTS_KEY = "gui_defaults"


def get_gui_defaults_path() -> Path:
    return get_indent_settings_path()


def built_in_gui_defaults() -> dict[str, bool]:
    return {
        "fix_table": True,
        "fix_color": True,
        "fix_paragraph": True,
        "remove_all_outline": True,
        "indent_preface": False,
        "outline_preface": False,
        "level1_level2_body_first_line_indent": True,
        "word_com_check_body_font": False,
        "skip_log_output": True,
        "skip_nested_tables": True,
        "skip_chapter_three_table_layout": True,
        "skip_chapter_three_table_color": True,
        "skip_chapter_three_indents": False,
        "skip_special_color_tables": False,
        "clear_special_colors_after_skip": False,
    }


def _coerce_bool(key: str, value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, int) and value in (0, 1):
        return bool(value)
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in ("true", "1", "yes", "y", "on"):
            return True
        if normalized in ("false", "0", "no", "n", "off"):
            return False
    raise ValueError(f"GUI default field {key!r} must be a bool-like value")


def normalize_gui_defaults(data: dict[str, Any] | None) -> dict[str, bool]:
    if data is None:
        return built_in_gui_defaults()
    if not isinstance(data, dict):
        raise ValueError("GUI defaults must be a JSON object")

    normalized = built_in_gui_defaults()
    for key in normalized:
        if key in data:
            normalized[key] = _coerce_bool(key, data[key])
    return normalized


def load_saved_gui_defaults(path: str | Path | None = None) -> dict[str, bool]:
    defaults_path = Path(path) if path is not None else get_gui_defaults_path()
    if not defaults_path.exists():
        return built_in_gui_defaults()

    data = json.loads(defaults_path.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("Shared defaults file must be a JSON object")
    return normalize_gui_defaults(data.get(GUI_DEFAULTS_KEY))


def save_gui_defaults(settings: dict[str, Any], path: str | Path | None = None) -> Path:
    normalized = normalize_gui_defaults(settings)
    defaults_path = Path(path) if path is not None else get_gui_defaults_path()
    data = {}
    if defaults_path.exists():
        loaded = json.loads(defaults_path.read_text(encoding="utf-8"))
        if not isinstance(loaded, dict):
            raise ValueError("Shared defaults file must be a JSON object")
        data = loaded
    if "body" in data or "preface" in data:
        data = {"indent_settings": data}
    data[GUI_DEFAULTS_KEY] = normalized
    defaults_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return defaults_path
