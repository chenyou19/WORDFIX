from __future__ import annotations

import shutil
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from .constants import NS
from .note_detection import note_source_for_paragraph
from .numbering import build_numbering_format_lookup, build_style_numbering_lookup
from .numbering_cleanup import _long_path_compatible_str
from .outline import set_paragraph_jc, summarize_paragraph_text
from .xml_utils import paragraph_text, qn


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


def _paragraph_jc(p) -> str | None:
    jc = p.find("./w:pPr/w:jc", NS)
    return jc.get(qn("val")) if jc is not None else None


def force_note_paragraph_left_alignment_in_root(
    root,
    *,
    part_name: str,
    numbering_format_lookup=None,
    style_numbering_lookup=None,
    logs: list[str] | None = None,
) -> int:
    fixed_count = 0
    total_candidate_paragraphs = 0
    matched_counts = {"text": 0, "numPr": 0, "styleNumPr": 0}
    skipped_table = 0
    skipped_non_note = 0
    center_after_fix_count = 0
    still_center_records: list[str] = []

    paragraphs = root.xpath(".//w:p", namespaces=NS)
    for paragraph_index, p in enumerate(paragraphs, start=1):
        text = paragraph_text(p)
        source = note_source_for_paragraph(
            p,
            text,
            numbering_format_lookup=numbering_format_lookup,
            style_numbering_lookup=style_numbering_lookup,
        )
        if source is None:
            skipped_non_note += 1
            continue

        total_candidate_paragraphs += 1
        if source in matched_counts:
            matched_counts[source] += 1

        if p.xpath("ancestor::w:tbl", namespaces=NS):
            skipped_table += 1
            continue

        before_jc, after_jc = set_paragraph_jc(p, "left")
        fixed_count += 1
        final_jc = _paragraph_jc(p) or after_jc
        if str(final_jc).lower() == "center":
            center_after_fix_count += 1
            still_center_records.append(
                f"{part_name}:{paragraph_index}:{summarize_paragraph_text(text)}"
            )
        if logs is not None:
            logs.append(
                "FINAL_NOTE_ALIGNMENT_FIX "
                f"part={part_name} "
                f"paragraph_index={paragraph_index} "
                f"source={source} "
                f"before_jc={before_jc or 'none'} "
                f"after_jc={after_jc} "
                f"text={summarize_paragraph_text(text)}"
            )

    if logs is not None:
        logs.append(
            "FINAL_NOTE_ALIGNMENT_SUMMARY "
            f"part={part_name} "
            f"total_candidate_paragraphs={total_candidate_paragraphs} "
            f"matched_text={matched_counts['text']} "
            f"matched_numPr={matched_counts['numPr']} "
            f"matched_styleNumPr={matched_counts['styleNumPr']} "
            f"fixed_count={fixed_count} "
            f"skipped_table={skipped_table} "
            f"skipped_non_note={skipped_non_note} "
            f"center_after_fix_count={center_after_fix_count} "
            f"still_center_records={';'.join(still_center_records) if still_center_records else 'none'}"
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
            names = set(zin.namelist())
            numbering_xml = zin.read("word/numbering.xml") if "word/numbering.xml" in names else None
            styles_xml = zin.read("word/styles.xml") if "word/styles.xml" in names else None
            numbering_format_lookup = build_numbering_format_lookup(numbering_xml)
            style_numbering_lookup = build_style_numbering_lookup(styles_xml)

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
                            numbering_format_lookup=numbering_format_lookup,
                            style_numbering_lookup=style_numbering_lookup,
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
