from __future__ import annotations

import os
from pathlib import Path

def normalize_path_for_compare(path: str | Path) -> str:
    """
    Convert a path to a comparable absolute path.

    Windows file names are usually case-insensitive; normcase avoids treating
    differently cased spellings of the same path as different files.
    """
    return os.path.normcase(os.path.abspath(os.path.expanduser(str(path))))


def is_same_file_path(path_a: str | Path, path_b: str | Path) -> bool:
    """
    Return whether two paths point to the same file.

    Use strict=False so an output file that does not exist yet can still be
    compared against the source path chosen by the user.
    """
    try:
        return Path(path_a).resolve(strict=False) == Path(path_b).resolve(strict=False)
    except Exception:
        return normalize_path_for_compare(path_a) == normalize_path_for_compare(path_b)
