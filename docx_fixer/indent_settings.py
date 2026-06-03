from __future__ import annotations

import json
import sys
from pathlib import Path

from .constants import (
    POINTS_PER_CM,
    PREFACE_OUTLINE_INDENTS,
    TEMPLATE_OUTLINE_INDENTS,
    make_outline_indent_spec,
)

INDENT_SETTINGS_FILENAME = "indent_defaults.json"

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


def spec_to_cm_values(spec: dict[str, str]) -> tuple[float, float, float]:
    number_start = twips_to_cm(spec.get("number_start", int(spec["left"]) - int(spec["hanging"])))
    hanging = twips_to_cm(spec["hanging"])
    body_left = twips_to_cm(spec.get("body_left", spec["left"]))
    return number_start, hanging, body_left


def make_indent_spec(number_start_cm: float, hanging_cm: float, body_left_cm: float) -> dict[str, str]:
    if hanging_cm <= 0:
        raise ValueError("凸排距離必須大於 0。")

    return make_outline_indent_spec(number_start_cm, hanging_cm, body_left_cm)


def built_in_indent_settings() -> dict[str, list[dict[str, float | int | str]]]:
    return {
        "body": [
            make_settings_row(level, BODY_LEVEL_LABELS[level], spec)
            for level, spec in TEMPLATE_OUTLINE_INDENTS.items()
        ],
        "preface": [
            make_settings_row(
                level,
                PREFACE_LEVEL_LABELS[level],
                PREFACE_OUTLINE_INDENTS.get(level, TEMPLATE_OUTLINE_INDENTS[level]),
            )
            for level in range(len(PREFACE_LEVEL_LABELS))
        ],
    }


def make_settings_row(level: int, label: str, spec: dict[str, str]) -> dict[str, float | int | str]:
    number_start_cm, hanging_cm, body_left_cm = spec_to_cm_values(spec)
    return {
        "level": level,
        "label": label,
        "number_start_cm": number_start_cm,
        "hanging_cm": hanging_cm,
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
        number_start_cm = float(row["number_start_cm"])
        hanging_cm = float(row["hanging_cm"])
        body_left_cm = float(row["body_left_cm"])
    else:
        # Backward compatibility for old indent_defaults.json:
        # left_cm was the heading text start. Body text followed that same value.
        left_cm = float(row["left_cm"])
        number_start_cm = float(row["number_start_cm"])
        hanging_cm = left_cm - number_start_cm
        body_left_cm = left_cm

    make_indent_spec(number_start_cm, hanging_cm, body_left_cm)
    return {
        "level": level,
        "label": label,
        "number_start_cm": number_start_cm,
        "hanging_cm": hanging_cm,
        "body_left_cm": body_left_cm,
    }


def normalize_indent_settings(settings: dict) -> dict[str, list[dict[str, float | int | str]]]:
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


def apply_indent_settings(settings: dict) -> dict[str, list[dict[str, float | int | str]]]:
    normalized = normalize_indent_settings(settings)

    TEMPLATE_OUTLINE_INDENTS.clear()
    for row in normalized["body"]:
        TEMPLATE_OUTLINE_INDENTS[int(row["level"])] = make_indent_spec(
            float(row["number_start_cm"]),
            float(row["hanging_cm"]),
            float(row["body_left_cm"]),
        )

    PREFACE_OUTLINE_INDENTS.clear()
    for row in normalized["preface"]:
        PREFACE_OUTLINE_INDENTS[int(row["level"])] = make_indent_spec(
            float(row["number_start_cm"]),
            float(row["hanging_cm"]),
            float(row["body_left_cm"]),
        )

    return normalized


def load_saved_indent_settings(path: str | Path | None = None) -> bool:
    settings_path = Path(path) if path is not None else get_indent_settings_path()
    if not settings_path.exists():
        return False

    data = json.loads(settings_path.read_text(encoding="utf-8"))
    apply_indent_settings(data)
    return True


def save_indent_settings(settings: dict, path: str | Path | None = None) -> Path:
    normalized = apply_indent_settings(settings)
    settings_path = Path(path) if path is not None else get_indent_settings_path()
    settings_path.write_text(
        json.dumps(normalized, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return settings_path
