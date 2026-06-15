from __future__ import annotations

from dataclasses import dataclass, field

from lxml import etree

from .constants import NS
from .note_detection import is_note_cell_text, normalize_note_text
from .xml_utils import paragraph_text, qn

# 10 pt in Word half-points.
NOTE_FONT_SIZE_HALF_POINTS = "20"
NOTE_FONT_EAST_ASIA = "標楷體"
NOTE_FONT_LATIN = "DFKai-SB"


@dataclass
class MovedNote:
    note_text: str
    delete_action: str  # "delete_row" or "delete_cell"
    row_index: int
    cell_index: int


@dataclass
class NoteMoveResult:
    note_cells_moved: bool = False
    moved_note_count: int = 0
    deleted_note_cells: int = 0
    deleted_note_rows: int = 0
    inserted_note_paragraphs: int = 0
    moved_notes: list[MovedNote] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def _cell_text(tc) -> str:
    return "".join(paragraph_text(p) for p in tc.findall("w:p", NS))


def _cell_is_blank(tc) -> bool:
    return normalize_note_text(_cell_text(tc)) == ""


def _build_note_paragraph(note_text: str):
    p = etree.Element(qn("p"))
    pPr = etree.SubElement(p, qn("pPr"))

    # Body outline level (Word UI level "本文" == w:outlineLvl val 9). No
    # heading numbering and no first-line indent so later paragraph passes do
    # not treat it as a heading.
    outline = etree.SubElement(pPr, qn("outlineLvl"))
    outline.set(qn("val"), "9")
    jc = etree.SubElement(pPr, qn("jc"))
    jc.set(qn("val"), "left")
    ind = etree.SubElement(pPr, qn("ind"))
    ind.set(qn("firstLine"), "0")

    p_rPr = etree.SubElement(pPr, qn("rPr"))
    _set_note_run_font(p_rPr)

    run = etree.SubElement(p, qn("r"))
    r_rPr = etree.SubElement(run, qn("rPr"))
    _set_note_run_font(r_rPr)
    t = etree.SubElement(run, qn("t"))
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t.text = note_text
    return p


def _set_note_run_font(rPr) -> None:
    rFonts = etree.SubElement(rPr, qn("rFonts"))
    rFonts.set(qn("eastAsia"), NOTE_FONT_EAST_ASIA)
    rFonts.set(qn("ascii"), NOTE_FONT_LATIN)
    rFonts.set(qn("hAnsi"), NOTE_FONT_LATIN)
    sz = etree.SubElement(rPr, qn("sz"))
    sz.set(qn("val"), NOTE_FONT_SIZE_HALF_POINTS)
    szCs = etree.SubElement(rPr, qn("szCs"))
    szCs.set(qn("val"), NOTE_FONT_SIZE_HALF_POINTS)


def move_table_note_cells_below(tbl) -> NoteMoveResult:
    """Move note cells (text starting with 註：/註1：/註一、 ...) out of a table.

    For each matching cell the row is deleted when every other cell in the row
    is blank, otherwise only the note cell is deleted. The note text is then
    inserted as its own paragraph immediately below the table, in the original
    top-to-bottom, left-to-right scan order. Returns statistics for logging.
    """
    result = NoteMoveResult()

    parent = tbl.getparent()
    if parent is None:
        result.warnings.append("note_move_warning=table_has_no_parent")
        return result

    # Collect matches first; mutate the tree afterwards so indices stay stable.
    rows = tbl.findall("w:tr", NS)
    rows_to_delete: list = []
    cells_to_delete: list = []
    for row_index, tr in enumerate(rows):
        cells = tr.findall("w:tc", NS)
        for cell_index, tc in enumerate(cells):
            text = _cell_text(tc)
            if not is_note_cell_text(text):
                continue

            note_text = normalize_note_text(text)
            others_blank = all(
                _cell_is_blank(other)
                for other_index, other in enumerate(cells)
                if other_index != cell_index
            )
            if others_blank:
                action = "delete_row"
                if tr not in rows_to_delete:
                    rows_to_delete.append(tr)
            else:
                action = "delete_cell"
                cells_to_delete.append((tr, tc))

            result.moved_notes.append(
                MovedNote(
                    note_text=note_text,
                    delete_action=action,
                    row_index=row_index,
                    cell_index=cell_index,
                )
            )

    if not result.moved_notes:
        return result

    # Insert note paragraphs immediately after the table, preserving order.
    anchor = tbl
    for note in result.moved_notes:
        paragraph = _build_note_paragraph(note.note_text)
        anchor.addnext(paragraph)
        anchor = paragraph
        result.inserted_note_paragraphs += 1

    for tr, tc in cells_to_delete:
        # The row may already be scheduled for deletion; skip if so.
        if tr in rows_to_delete:
            continue
        if tc.getparent() is tr:
            tr.remove(tc)
            result.deleted_note_cells += 1
        else:
            result.warnings.append("note_move_warning=cell_parent_changed")

    for tr in rows_to_delete:
        if tr.getparent() is tbl:
            tbl.remove(tr)
            result.deleted_note_rows += 1
        else:
            result.warnings.append("note_move_warning=row_parent_changed")

    result.note_cells_moved = True
    result.moved_note_count = len(result.moved_notes)
    return result
