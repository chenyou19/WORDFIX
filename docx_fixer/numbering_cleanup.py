from __future__ import annotations

import shutil
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from .numbering import force_clean_numbering_suffix_tabs


def _long_path_compatible_str(path: Path) -> str:
    resolved = str(path.resolve())
    if resolved.startswith("\\\\?\\"):
        return resolved
    if len(resolved) >= 240 and Path(resolved).drive:
        return "\\\\?\\" + resolved
    return resolved


def force_clean_numbering_suffix_tabs_in_docx(
    docx_path: str | Path,
    logs: list[str] | None = None,
    excluded_numbering_pairs: set[tuple[str, int]] | None = None,
    excluded_num_ids: set[str] | None = None,
    excluded_abstract_ids: set[str] | None = None,
    included_numbering_pairs: set[tuple[str, int]] | None = None,
    included_num_ids: set[str] | None = None,
    included_abstract_ids: set[str] | None = None,
    protected_numbering_pairs: set[tuple[str, int]] | None = None,
) -> bool:
    docx_path = Path(docx_path)
    temp_docx = docx_path.with_suffix(docx_path.suffix + ".numbering_suffix_clean.tmp")
    changed = False

    try:
        with ZipFile(docx_path, "r") as zin:
            if "word/numbering.xml" not in zin.namelist():
                if logs is not None:
                    logs.append("FINAL_NUMBERING_SUFFIX_CLEAN_DOCX_SKIPPED reason=missing_numbering_xml")
                return False

            with ZipFile(temp_docx, "w", ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename == "word/numbering.xml":
                        cleaned = force_clean_numbering_suffix_tabs(
                            data,
                            logs=logs,
                            excluded_numbering_pairs=excluded_numbering_pairs,
                            excluded_num_ids=excluded_num_ids,
                            excluded_abstract_ids=excluded_abstract_ids,
                            included_numbering_pairs=included_numbering_pairs,
                            included_num_ids=included_num_ids,
                            included_abstract_ids=included_abstract_ids,
                            protected_numbering_pairs=protected_numbering_pairs,
                        )
                        if cleaned is not None and cleaned != data:
                            data = cleaned
                            changed = True
                    zout.writestr(item, data)

        if changed:
            shutil.move(_long_path_compatible_str(temp_docx), _long_path_compatible_str(docx_path))
        else:
            temp_docx.unlink(missing_ok=True)
        if logs is not None:
            logs.append(f"FINAL_NUMBERING_SUFFIX_CLEAN_DOCX changed={'true' if changed else 'false'}")
        return changed
    except Exception as exc:
        if logs is not None:
            logs.append(f"FINAL_NUMBERING_SUFFIX_CLEAN_DOCX_FAILED reason={exc!r}")
        try:
            if temp_docx.exists():
                temp_docx.unlink()
        except Exception:
            pass
        return False
