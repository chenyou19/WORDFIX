from __future__ import annotations

import json
import sys
from pathlib import Path

from .constants import (
    POINTS_PER_CM,
    PREFACE_OUTLINE_INDENTS,
    TEMPLATE_OUTLINE_INDENTS,
    cm_to_twips,
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


def spec_to_cm_values(spec: dict[str, str]) -> tuple[float, float]:
    left = twips_to_cm(spec["left"])
    number_start = twips_to_cm(spec.get("number_start", int(spec["left"]) - int(spec["hanging"])))
    return left, number_start


def make_indent_spec(left_cm: float, number_start_cm: float) -> dict[str, str]:
    hanging_cm = left_cm - number_start_cm
    if hanging_cm <= 0:
        raise ValueError("懸掛值必須大於 0，請確認文字起點大於編號起點。")

    return {
        "left": cm_to_twips(left_cm),
        "hanging": cm_to_twips(hanging_cm),
        "number_start": cm_to_twips(number_start_cm),
    }


def built_in_indent_settings() -> dict[str, list[dict[str, float | int | str]]]:
    return {
        "body": [
            {
                "level": level,
                "label": BODY_LEVEL_LABELS[level],
                "left_cm": spec_to_cm_values(spec)[0],
                "number_start_cm": spec_to_cm_values(spec)[1],
            }
            for level, spec in TEMPLATE_OUTLINE_INDENTS.items()
        ],
        "preface": [
            {
                "level": level,
                "label": PREFACE_LEVEL_LABELS[level],
                "left_cm": spec_to_cm_values(PREFACE_OUTLINE_INDENTS.get(level, TEMPLATE_OUTLINE_INDENTS[level]))[0],
                "number_start_cm": spec_to_cm_values(PREFACE_OUTLINE_INDENTS.get(level, TEMPLATE_OUTLINE_INDENTS[level]))[1],
            }
            for level in range(len(PREFACE_LEVEL_LABELS))
        ],
    }


def current_indent_settings() -> dict[str, list[dict[str, float | int | str]]]:
    settings = built_in_indent_settings()
    for row in settings["body"]:
        level = int(row["level"])
        row["left_cm"], row["number_start_cm"] = spec_to_cm_values(TEMPLATE_OUTLINE_INDENTS[level])

    for row in settings["preface"]:
        level = int(row["level"])
        spec = PREFACE_OUTLINE_INDENTS.get(level, TEMPLATE_OUTLINE_INDENTS[level])
        row["left_cm"], row["number_start_cm"] = spec_to_cm_values(spec)

    return settings


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

            row = by_level[level]
            left_cm = float(row["left_cm"])
            number_start_cm = float(row["number_start_cm"])
            make_indent_spec(left_cm, number_start_cm)
            normalized[section].append({
                "level": level,
                "label": label,
                "left_cm": left_cm,
                "number_start_cm": number_start_cm,
            })

    return normalized


def apply_indent_settings(settings: dict) -> dict[str, list[dict[str, float | int | str]]]:
    normalized = normalize_indent_settings(settings)

    TEMPLATE_OUTLINE_INDENTS.clear()
    for row in normalized["body"]:
        TEMPLATE_OUTLINE_INDENTS[int(row["level"])] = make_indent_spec(
            float(row["left_cm"]),
            float(row["number_start_cm"]),
        )

    PREFACE_OUTLINE_INDENTS.clear()
    for row in normalized["preface"]:
        PREFACE_OUTLINE_INDENTS[int(row["level"])] = make_indent_spec(
            float(row["left_cm"]),
            float(row["number_start_cm"]),
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
