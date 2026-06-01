from __future__ import annotations

import os
from pathlib import Path

def normalize_path_for_compare(path: str | Path) -> str:
    """
    將路徑轉成可比較的絕對路徑。

    Windows 檔名大小寫通常不敏感；用 normcase 可避免
    D:/A.docx 與 d:/a.docx 被誤判成不同檔案。
    """
    return os.path.normcase(os.path.abspath(os.path.expanduser(str(path))))


def is_same_file_path(path_a: str | Path, path_b: str | Path) -> bool:
    """
    判斷兩個路徑是否指向同一個檔案。

    使用 strict=False，讓輸出檔尚未存在時也能正確比較
    「使用者指定的輸出檔名」是否就是來源檔案。
    """
    try:
        return Path(path_a).resolve(strict=False) == Path(path_b).resolve(strict=False)
    except Exception:
        return normalize_path_for_compare(path_a) == normalize_path_for_compare(path_b)
