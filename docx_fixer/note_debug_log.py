from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from .constants import NS
from .note_detection import NOTE_PREFIX, normalize_note_text, note_source_for_paragraph
from .numbering import (
    build_numbering_format_lookup,
    build_numbering_level_lookup,
    build_style_numbering_lookup,
)
from .outline import collect_all_toc_paragraph_ids
from .protected_region import ProtectedRegionContext
from .xml_utils import paragraph_text, qn


def should_debug_note_part(name: str) -> bool:
    if name == "word/document.xml":
        return True
    if name.startswith("word/header") and name.endswith(".xml"):
        return True
    if name.startswith("word/footer") and name.endswith(".xml"):
        return True
    if name in {"word/footnotes.xml", "word/endnotes.xml"}:
        return True
    return False


def _fmt(value) -> str:
    if value is None:
        return "none"
    text = str(value)
    return text.replace("\r", "\\r").replace("\n", "\\n")


def _paragraph_style_id(p) -> str | None:
    style_el = p.find("./w:pPr/w:pStyle", NS)
    return style_el.get(qn("val")) if style_el is not None else None


def _paragraph_num_identity(p) -> tuple[str | None, int | None]:
    ilvl_el = p.find("./w:pPr/w:numPr/w:ilvl", NS)
    num_id_el = p.find("./w:pPr/w:numPr/w:numId", NS)
    num_id = num_id_el.get(qn("val")) if num_id_el is not None else None
    try:
        ilvl = int(ilvl_el.get(qn("val"))) if ilvl_el is not None else 0
    except Exception:
        ilvl = None
    return num_id, ilvl


def _paragraph_jc(p) -> str | None:
    jc = p.find("./w:pPr/w:jc", NS)
    return jc.get(qn("val")) if jc is not None else None


def _ppr_xml(p) -> str:
    pPr = p.find("./w:pPr", NS)
    if pPr is None:
        return "none"
    return etree.tostring(pPr, encoding="unicode")


def _ppr_child_order(p) -> str:
    pPr = p.find("./w:pPr", NS)
    if pPr is None:
        return "none"
    tags = [etree.QName(child).localname for child in pPr]
    return ",".join(tags) or "empty"


def _style_name_value(style) -> str:
    name_el = style.find("w:name", NS)
    return name_el.get(qn("val")) if name_el is not None else ""


def build_style_paragraph_format_lookup(styles_xml: bytes | None) -> dict[str, dict[str, object]]:
    if not styles_xml:
        return {}
    try:
        root = etree.fromstring(styles_xml)
    except Exception:
        return {}

    direct: dict[str, dict[str, object]] = {}
    for style in root.xpath("./w:style[@w:type='paragraph']", namespaces=NS):
        style_id = style.get(qn("styleId"))
        if not style_id:
            continue
        based_on_el = style.find("w:basedOn", NS)
        jc_el = style.find("./w:pPr/w:jc", NS)
        num_id, ilvl = None, None
        num_pr = style.find("./w:pPr/w:numPr", NS)
        if num_pr is not None:
            num_id_el = num_pr.find("w:numId", NS)
            ilvl_el = num_pr.find("w:ilvl", NS)
            num_id = num_id_el.get(qn("val")) if num_id_el is not None else None
            try:
                ilvl = int(ilvl_el.get(qn("val"))) if ilvl_el is not None else 0
            except Exception:
                ilvl = None

        direct[style_id] = {
            "style_id": style_id,
            "style_name": _style_name_value(style),
            "based_on": based_on_el.get(qn("val")) if based_on_el is not None else None,
            "jc_direct": jc_el.get(qn("val")) if jc_el is not None else None,
            "num_id_direct": num_id,
            "ilvl_direct": ilvl,
        }

    resolved: dict[str, dict[str, object]] = {}

    def resolve(style_id: str, seen: tuple[str, ...] = ()) -> dict[str, object]:
        if style_id in resolved:
            return resolved[style_id]
        info = dict(direct.get(style_id, {"style_id": style_id}))
        based_on = info.get("based_on")
        chain = [style_id]
        inherited_jc = None
        inherited_num_id = None
        inherited_ilvl = None
        if isinstance(based_on, str) and based_on not in seen:
            parent = resolve(based_on, (*seen, style_id))
            chain.extend(str(parent.get("style_based_on_chain", "")).split(">"))
            inherited_jc = parent.get("jc_effective")
            inherited_num_id = parent.get("num_id_effective")
            inherited_ilvl = parent.get("ilvl_effective")

        info["jc_effective"] = info.get("jc_direct") or inherited_jc
        info["num_id_effective"] = info.get("num_id_direct") or inherited_num_id
        info["ilvl_effective"] = info.get("ilvl_direct")
        if info["ilvl_effective"] is None:
            info["ilvl_effective"] = inherited_ilvl
        info["style_based_on_chain"] = ">".join([part for part in chain if part])
        resolved[style_id] = info
        return info

    for style_id in direct:
        resolve(style_id)
    return resolved


def _identity_text(num_id, ilvl) -> str:
    if num_id is None:
        return "none"
    return f"{num_id}:{ilvl if ilvl is not None else 'none'}"


def _level_format(num_id, ilvl, numbering_format_lookup) -> dict[str, object]:
    if num_id is None or ilvl is None:
        return {}
    try:
        ilvl_int = int(ilvl)
    except Exception:
        return {}
    return numbering_format_lookup.get((str(num_id), ilvl_int), {})


def _contains_note(value: object) -> bool:
    return NOTE_PREFIX in normalize_note_text(str(value or ""))


def _is_center(value: object) -> bool:
    return str(value or "").lower() == "center"


def _style_is_suspicious(style_id: str | None, style_info: dict[str, object]) -> bool:
    values = [
        style_id or "",
        str(style_info.get("style_name") or ""),
        str(style_info.get("style_based_on_chain") or ""),
    ]
    for value in values:
        lowered = value.lower()
        if "note" in lowered or "remark" in lowered or _contains_note(value) or "備註" in value:
            return True
    return False


def _record_is_candidate(details: dict[str, object]) -> bool:
    if details["note_source"] != "none":
        return True
    if _contains_note(details["raw_text"]) or _contains_note(details["normalized_text"]):
        return True
    if _style_is_suspicious(details["paragraph_style_id"], details["style_info"]):
        return True
    if _contains_note(details["numbering_lvlText"]) or _contains_note(details["style_numbering_lvlText"]):
        return True
    if any(
        _is_center(details.get(key))
        for key in ("paragraph_jc", "style_jc_effective", "numbering_lvlJc", "final_paragraph_jc")
    ):
        return True
    return bool(details["paragraph_has_numPr"] and len(str(details["normalized_text"])) <= 20)


def _style_lookup_details(style_id: str | None, style_format_lookup) -> dict[str, object]:
    if not style_id:
        return {}
    return style_format_lookup.get(style_id, {})


def _collect_chapter_three_ids(document_root, numbering_xml, numbering_level_lookup, numbering_format_lookup, style_numbering_lookup):
    try:
        context = ProtectedRegionContext.from_document(
            document_root,
            protect_chapter_three=True,
            numbering_level_lookup=numbering_level_lookup,
            numbering_format_lookup=numbering_format_lookup,
            style_numbering_lookup=style_numbering_lookup,
            numbering_xml=numbering_xml,
        )
    except Exception:
        return None
    return context.document_chapter_three_paragraph_ids


def collect_note_debug_records_from_docx(docx_path: str | Path, stage: str) -> list[str]:
    docx_path = Path(docx_path)
    parser = etree.XMLParser(remove_blank_text=False, recover=True)
    blocks: list[str] = [
        "===== NOTE DEBUG FILE =====",
        f"stage={stage}",
        f"docx_path={docx_path}",
    ]
    record_index = 0

    with ZipFile(docx_path, "r") as zin:
        names = set(zin.namelist())
        numbering_xml = zin.read("word/numbering.xml") if "word/numbering.xml" in names else None
        styles_xml = zin.read("word/styles.xml") if "word/styles.xml" in names else None
        numbering_format_lookup = build_numbering_format_lookup(numbering_xml)
        numbering_level_lookup = build_numbering_level_lookup(numbering_xml)
        style_numbering_lookup = build_style_numbering_lookup(styles_xml)
        style_format_lookup = build_style_paragraph_format_lookup(styles_xml)

        for part_name in sorted(name for name in names if should_debug_note_part(name)):
            try:
                root = etree.fromstring(zin.read(part_name), parser)
            except Exception as exc:
                record_index += 1
                blocks.extend(
                    [
                        "",
                        f"===== NOTE DEBUG {record_index:03d} =====",
                        f"stage={stage}",
                        f"part={part_name}",
                        "decision=ERROR",
                        f"reason=parse_error:{exc!r}",
                    ]
                )
                continue

            paragraphs = root.xpath(".//w:p", namespaces=NS)
            try:
                toc_ids = collect_all_toc_paragraph_ids(
                    root,
                    numbering_level_lookup=numbering_level_lookup,
                    style_numbering_lookup=style_numbering_lookup,
                    paragraphs=paragraphs,
                )
            except Exception:
                toc_ids = set()

            chapter_three_ids = (
                _collect_chapter_three_ids(
                    root,
                    numbering_xml,
                    numbering_level_lookup,
                    numbering_format_lookup,
                    style_numbering_lookup,
                )
                if part_name == "word/document.xml"
                else None
            )

            for paragraph_index, p in enumerate(paragraphs, start=1):
                raw_text = paragraph_text(p)
                normalized_text = normalize_note_text(raw_text)
                style_id = _paragraph_style_id(p)
                style_info = _style_lookup_details(style_id, style_format_lookup)
                paragraph_num_id, paragraph_ilvl = _paragraph_num_identity(p)
                style_num_id, style_ilvl = style_numbering_lookup.get(style_id or "", (None, None))
                effective_num_id = paragraph_num_id if paragraph_num_id is not None else style_num_id
                effective_ilvl = paragraph_ilvl if paragraph_num_id is not None else style_ilvl
                level_format = _level_format(effective_num_id, effective_ilvl, numbering_format_lookup)
                style_level_format = _level_format(style_num_id, style_ilvl, numbering_format_lookup)
                note_source = note_source_for_paragraph(
                    p,
                    raw_text,
                    numbering_format_lookup=numbering_format_lookup,
                    style_numbering_lookup=style_numbering_lookup,
                ) or "none"
                paragraph_jc = _paragraph_jc(p)
                in_table = bool(p.xpath("ancestor::w:tbl", namespaces=NS))
                in_toc = id(p) in toc_ids
                protected_value = (
                    "unknown"
                    if chapter_three_ids is None
                    else ("True" if id(p) in chapter_three_ids else "False")
                )
                details = {
                    "raw_text": raw_text,
                    "normalized_text": normalized_text,
                    "note_source": note_source,
                    "paragraph_style_id": style_id or "none",
                    "style_info": style_info,
                    "paragraph_has_numPr": paragraph_num_id is not None,
                    "paragraph_jc": paragraph_jc or "none",
                    "final_paragraph_jc": paragraph_jc or "none",
                    "style_jc_effective": style_info.get("jc_effective") or "none",
                    "numbering_lvlText": level_format.get("lvlText") or "none",
                    "style_numbering_lvlText": style_level_format.get("lvlText") or "none",
                    "numbering_lvlJc": level_format.get("lvlJc") or "none",
                }
                if not _record_is_candidate(details):
                    continue

                record_index += 1
                visible_note_guess = (
                    note_source != "none"
                    or _contains_note(raw_text)
                    or _contains_note(level_format.get("lvlText"))
                    or _contains_note(style_level_format.get("lvlText"))
                )
                if in_table:
                    decision = "SKIPPED_TABLE"
                    reason = "paragraph is inside a table; note final pass ignores table paragraphs"
                elif note_source == "none":
                    decision = "SKIPPED_NOT_NOTE"
                    reason = "candidate did not match text, direct numbering, or style numbering note detection"
                else:
                    decision = "FIXED_LEFT"
                    reason = "candidate matched note detection and is outside tables"

                style_num_text = _identity_text(style_num_id, style_ilvl)
                effective_num_text = _identity_text(effective_num_id, effective_ilvl)
                paragraph_num_text = _identity_text(paragraph_num_id, paragraph_ilvl)
                warnings = []
                if _is_center(level_format.get("lvlJc")):
                    warnings.append("WARNING: numbering level alignment is center")
                if _is_center(paragraph_jc):
                    warnings.append("WARNING: final paragraph alignment is center")
                if _is_center(style_info.get("jc_effective")):
                    warnings.append("WARNING: effective paragraph style alignment is center")

                blocks.extend(
                    [
                        "",
                        f"===== NOTE DEBUG {record_index:03d} =====",
                        f"stage={stage}",
                        f"part={part_name}",
                        f"paragraph_index={paragraph_index}",
                        f"in_table={in_table}",
                        f"in_toc={in_toc}",
                        f"in_chapter_three_protected={protected_value}",
                        f"raw_text={_fmt(raw_text)}",
                        f"normalized_text={_fmt(normalized_text)}",
                        f"visible_note_guess={visible_note_guess}",
                        f"note_source={note_source}",
                        f"paragraph_style_id={_fmt(style_id)}",
                        f"paragraph_style_name={_fmt(style_info.get('style_name'))}",
                        f"paragraph_style_numPr={style_num_text}",
                        f"style_jc_direct={_fmt(style_info.get('jc_direct'))}",
                        f"style_jc_effective={_fmt(style_info.get('jc_effective'))}",
                        f"style_based_on_chain={_fmt(style_info.get('style_based_on_chain'))}",
                        f"paragraph_has_numPr={paragraph_num_id is not None}",
                        f"paragraph_numId={_fmt(paragraph_num_id)}",
                        f"paragraph_ilvl={_fmt(paragraph_ilvl)}",
                        f"effective_numId={_fmt(effective_num_id)}",
                        f"effective_ilvl={_fmt(effective_ilvl)}",
                        f"effective_numPr={effective_num_text}",
                        f"numbering_lvlText={_fmt(level_format.get('lvlText'))}",
                        f"numbering_numFmt={_fmt(level_format.get('numFmt'))}",
                        f"numbering_lvlJc={_fmt(level_format.get('lvlJc'))}",
                        f"numbering_suff={_fmt(level_format.get('suff'))}",
                        f"style_numbering_lvlText={_fmt(style_level_format.get('lvlText'))}",
                        f"style_jc={_fmt(style_info.get('jc_effective'))}",
                        f"paragraph_jc={_fmt(paragraph_jc)}",
                        "before_jc=unknown_snapshot_only",
                        f"after_jc={_fmt(paragraph_jc)}",
                        f"final_jc={_fmt(paragraph_jc)}",
                        f"final_paragraph_jc={_fmt(paragraph_jc)}",
                        f"pPr_child_order={_ppr_child_order(p)}",
                        f"pPr_xml={_ppr_xml(p)}",
                        f"final_pPr_xml={_ppr_xml(p)}",
                        f"decision={decision}",
                        f"reason={reason}",
                    ]
                )
                blocks.extend(warnings)

    blocks.append("")
    blocks.append(f"NOTE_DEBUG_RECORD_COUNT={record_index}")
    return blocks


def write_note_debug_log_for_docx(
    docx_path: str | Path,
    log_path: str | Path,
    stage: str,
    *,
    append: bool = False,
) -> Path:
    log_path = Path(log_path)
    lines = collect_note_debug_records_from_docx(docx_path, stage)
    mode = "a" if append and log_path.exists() else "w"
    with log_path.open(mode, encoding="utf-8", newline="\n") as handle:
        if mode == "a":
            handle.write("\n\n")
        handle.write("\n".join(lines))
        handle.write("\n")
    return log_path
