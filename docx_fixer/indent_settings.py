from __future__ import annotations

import json
import math
import sys
from pathlib import Path

from .constants import (
    BUILTIN_PREFACE_OUTLINE_INDENTS,
    BUILTIN_TEMPLATE_OUTLINE_INDENTS,
    DEFAULT_HEADING_TEXT_START_OFFSET_CM,
    POINTS_PER_CM,
    PREFACE_OUTLINE_INDENTS,
    TEMPLATE_OUTLINE_INDENTS,
    make_outline_indent_spec,
)

INDENT_SETTINGS_FILENAME = "indent_defaults.json"
INDENT_SETTINGS_KEY = "indent_settings"

BODY_LEVEL_LABELS = [
    "壹、",
    "一、",
    "（一）/(一)",
    "1.",
    "（1）/(1)",
    "A.",
    "（A）/(A)",
    "a.",
    "（a）/(a)",
]

PREFACE_LEVEL_LABELS = [
    "一、",
    "（一）/(一)",
    "1.",
    "（1）/(1)",
    "A.",
    "（A）/(A)",
    "a.",
    "（a）/(a)",
]


def get_indent_settings_path() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).with_name(INDENT_SETTINGS_FILENAME)
    return Path.cwd() / INDENT_SETTINGS_FILENAME


def twips_to_cm(value: str | int) -> float:
    return int(value) / 20 / POINTS_PER_CM


def format_cm(value: float) -> str:
    text = f"{value:.2f}"
    return text.rstrip("0").rstrip(".")


def default_heading_text_start_cm(body_left_cm: float) -> float:
    return body_left_cm + DEFAULT_HEADING_TEXT_START_OFFSET_CM


def _finite_cm(value: object, field_name: str) -> float:
    try:
        cm_value = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"{field_name} 必須是有效數字。") from exc
    if not math.isfinite(cm_value):
        raise ValueError(f"{field_name} 必須是有效有限數字。")
    return cm_value


def spec_to_cm_values(spec: dict[str, str]) -> tuple[float, float, float, float]:
    number_start = twips_to_cm(spec.get("number_start", int(spec["left"]) - int(spec["hanging"])))
    hanging = twips_to_cm(spec["hanging"])
    body_left = twips_to_cm(spec.get("body_left", spec["left"]))
    if "heading_text_start" in spec:
        heading_text_start = twips_to_cm(spec["heading_text_start"])
    else:
        heading_text_start = default_heading_text_start_cm(body_left)
    return number_start, hanging, heading_text_start, body_left


def make_indent_spec(
    number_start_cm: float,
    hanging_cm: float,
    body_left_cm: float,
    heading_text_start_cm: float | None = None,
) -> dict[str, str]:
    number_start_cm = _finite_cm(number_start_cm, "標號起點")
    hanging_cm = _finite_cm(hanging_cm, "懸掛")
    body_left_cm = _finite_cm(body_left_cm, "內文起點")
    if heading_text_start_cm is None:
        heading_text_start_cm = default_heading_text_start_cm(body_left_cm)
    heading_text_start_cm = _finite_cm(heading_text_start_cm, "標題文字起點")
    if hanging_cm <= 0:
        raise ValueError("凸排距離必須大於 0。")

    return make_outline_indent_spec(number_start_cm, hanging_cm, body_left_cm, heading_text_start_cm)


def built_in_indent_settings() -> dict[str, list[dict[str, float | int | str]]]:
    return {
        "body": [
            make_settings_row(level, BODY_LEVEL_LABELS[level], spec)
            for level, spec in BUILTIN_TEMPLATE_OUTLINE_INDENTS.items()
        ],
        "preface": [
            make_settings_row(
                level,
                PREFACE_LEVEL_LABELS[level],
                BUILTIN_PREFACE_OUTLINE_INDENTS.get(level, BUILTIN_TEMPLATE_OUTLINE_INDENTS[level]),
            )
            for level in range(len(PREFACE_LEVEL_LABELS))
        ],
    }


def make_settings_row(level: int, label: str, spec: dict[str, str]) -> dict[str, float | int | str]:
    number_start_cm, hanging_cm, heading_text_start_cm, body_left_cm = spec_to_cm_values(spec)
    return {
        "level": level,
        "label": label,
        "number_start_cm": number_start_cm,
        "hanging_cm": hanging_cm,
        "heading_text_start_cm": heading_text_start_cm,
        "body_left_cm": body_left_cm,
    }


def current_indent_settings() -> dict[str, list[dict[str, float | int | str]]]:
    settings = built_in_indent_settings()
    for row in settings["body"]:
        level = int(row["level"])
        row.update(make_settings_row(level, str(row["label"]), TEMPLATE_OUTLINE_INDENTS[level]))

    for row in settings["preface"]:
        level = int(row["level"])
        spec = PREFACE_OUTLINE_INDENTS.get(level, TEMPLATE_OUTLINE_INDENTS[level])
        row.update(make_settings_row(level, str(row["label"]), spec))

    return settings


def normalize_row(row: dict, level: int, label: str) -> dict[str, float | int | str]:
    if "hanging_cm" in row or "body_left_cm" in row:
        number_start_cm = _finite_cm(row["number_start_cm"], "標號起點")
        hanging_cm = _finite_cm(row["hanging_cm"], "懸掛")
        body_left_cm = _finite_cm(row["body_left_cm"], "內文起點")
    else:
        # Backward compatibility for old indent_defaults.json:
        # left_cm was the heading text start. Body text followed that same value.
        left_cm = _finite_cm(row["left_cm"], "起點")
        number_start_cm = _finite_cm(row["number_start_cm"], "標號起點")
        hanging_cm = left_cm - number_start_cm
        body_left_cm = left_cm

    if "heading_text_start_cm" in row:
        heading_text_start_cm = _finite_cm(row["heading_text_start_cm"], "標題文字起點")
    else:
        heading_text_start_cm = default_heading_text_start_cm(body_left_cm)

    make_indent_spec(number_start_cm, hanging_cm, body_left_cm, heading_text_start_cm)
    return {
        "level": level,
        "label": label,
        "number_start_cm": number_start_cm,
        "hanging_cm": hanging_cm,
        "heading_text_start_cm": heading_text_start_cm,
        "body_left_cm": body_left_cm,
    }


def normalize_indent_settings(settings: dict) -> dict[str, list[dict[str, float | int | str]]]:
    settings = extract_indent_settings(settings)
    normalized = {"body": [], "preface": []}
    expected = {
        "body": BODY_LEVEL_LABELS,
        "preface": PREFACE_LEVEL_LABELS,
    }

    for section, labels in expected.items():
        rows = settings.get(section)
        if not isinstance(rows, list):
            raise ValueError(f"缺少 {section} 縮排設定。")

        by_level = {}
        for row in rows:
            if not isinstance(row, dict):
                continue
            try:
                by_level[int(row["level"])] = row
            except (KeyError, TypeError, ValueError):
                continue

        for level, label in enumerate(labels):
            if level not in by_level:
                raise ValueError(f"{section} 缺少第 {level + 1} 階設定。")

            normalized[section].append(normalize_row(by_level[level], level, label))

    return normalized


def extract_indent_settings(data: dict) -> dict:
    if not isinstance(data, dict):
        raise ValueError("Indent settings must be a JSON object")
    if INDENT_SETTINGS_KEY in data:
        settings = data[INDENT_SETTINGS_KEY]
        if not isinstance(settings, dict):
            raise ValueError("indent_settings must be a JSON object")
        return settings
    return data


def apply_indent_settings(settings: dict) -> dict[str, list[dict[str, float | int | str]]]:
    normalized = normalize_indent_settings(settings)

    TEMPLATE_OUTLINE_INDENTS.clear()
    for row in normalized["body"]:
        TEMPLATE_OUTLINE_INDENTS[int(row["level"])] = make_indent_spec(
            float(row["number_start_cm"]),
            float(row["hanging_cm"]),
            float(row["body_left_cm"]),
            float(row["heading_text_start_cm"]),
        )

    PREFACE_OUTLINE_INDENTS.clear()
    for row in normalized["preface"]:
        PREFACE_OUTLINE_INDENTS[int(row["level"])] = make_indent_spec(
            float(row["number_start_cm"]),
            float(row["hanging_cm"]),
            float(row["body_left_cm"]),
            float(row["heading_text_start_cm"]),
        )

    return normalized


def load_saved_indent_settings(path: str | Path | None = None) -> bool:
    settings_path = Path(path) if path is not None else get_indent_settings_path()
    if not settings_path.exists():
        return False

    data = json.loads(settings_path.read_text(encoding="utf-8"))
    if INDENT_SETTINGS_KEY not in data and not ("body" in data and "preface" in data):
        return False
    apply_indent_settings(data)
    return True


def save_indent_settings(settings: dict, path: str | Path | None = None) -> Path:
    normalized = apply_indent_settings(settings)
    settings_path = Path(path) if path is not None else get_indent_settings_path()
    data = {}
    if settings_path.exists():
        loaded = json.loads(settings_path.read_text(encoding="utf-8"))
        if isinstance(loaded, dict):
            data = loaded
    if INDENT_SETTINGS_KEY not in data and ("body" in data or "preface" in data):
        data = {}
    data[INDENT_SETTINGS_KEY] = normalized
    settings_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return settings_path
