from __future__ import annotations

import shutil
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .constants import NS
from .numbering_cleanup import _long_path_compatible_str
from .outline import is_note_paragraph, set_paragraph_jc, summarize_paragraph_text
from .xml_utils import paragraph_text


def should_fix_note_alignment_part(name: str) -> bool:
    if name == "word/document.xml":
        return True
    if name.startswith("word/header") and name.endswith(".xml"):
        return True
    if name.startswith("word/footer") and name.endswith(".xml"):
        return True
    if name in {"word/footnotes.xml", "word/endnotes.xml"}:
        return True
    return False


def force_note_paragraph_left_alignment_in_root(
    root,
    *,
    part_name: str,
    logs: list[str] | None = None,
) -> int:
    fixed_count = 0
    paragraphs = root.xpath(".//w:p[not(ancestor::w:tbl)]", namespaces=NS)
    for paragraph_index, p in enumerate(paragraphs, start=1):
        text = paragraph_text(p)
        if not is_note_paragraph(text):
            continue

        before_jc, after_jc = set_paragraph_jc(p, "left")
        fixed_count += 1
        if logs is not None:
            logs.append(
                "FINAL_NOTE_ALIGNMENT_FIX "
                f"part={part_name} "
                f"paragraph_index={paragraph_index} "
                f"before_jc={before_jc or 'none'} "
                f"after_jc={after_jc} "
                f"text={summarize_paragraph_text(text)}"
            )

    return fixed_count


def force_note_paragraph_left_alignment_in_docx(
    docx_path: str | Path,
    logs: list[str] | None = None,
) -> bool:
    docx_path = Path(docx_path)
    temp_docx = docx_path.with_suffix(docx_path.suffix + ".note_alignment.tmp")
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    changed = False

    try:
        with ZipFile(docx_path, "r") as zin, ZipFile(temp_docx, "w", ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if should_fix_note_alignment_part(item.filename):
                    try:
                        root = etree.fromstring(data, parser)
                    except Exception as exc:
                        if logs is not None:
                            logs.append(
                                "FINAL_NOTE_ALIGNMENT_FIX_FAILED "
                                f"part={item.filename} reason={exc!r}"
                            )
                    else:
                        fixed_count = force_note_paragraph_left_alignment_in_root(
                            root,
                            part_name=item.filename,
                            logs=logs,
                        )
                        if fixed_count:
                            data = etree.tostring(
                                root,
                                xml_declaration=True,
                                encoding="UTF-8",
                                standalone=True,
                            )
                            changed = True
                zout.writestr(item, data)

        if changed:
            shutil.move(_long_path_compatible_str(temp_docx), _long_path_compatible_str(docx_path))
        else:
            temp_docx.unlink(missing_ok=True)
        return changed
    except Exception as exc:
        if logs is not None:
            logs.append(f"FINAL_NOTE_ALIGNMENT_FIX_DOCX_FAILED reason={exc!r}")
        try:
            if temp_docx.exists():
                temp_docx.unlink()
        except Exception:
            pass
        return False
