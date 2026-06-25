from __future__ import annotations

from dataclasses import dataclass, field

@dataclass
class ProcessOptions:
    fix_table_layout: bool
    fix_color: bool
    fix_paragraph: bool
    remove_all_outline_levels: bool = False
    indent_preface_paragraphs: bool = False
    outline_preface_paragraphs: bool = False
    normalize_with_word_com: bool = True
    enable_level1_level2_body_first_line_indent: bool = False
    word_com_check_body_font_when_xml_not_14: bool = False
    normalize_body_style_to_none: bool = False
    skip_chapter_three_table_layout: bool = False
    skip_chapter_three_table_color: bool = False
    skip_chapter_three_tables: bool = False
    skip_chapter_three_indents: bool = False
    skip_chapter_three_numbering_suffix_cleanup: bool = True
    skip_all_under_chapter_three: bool = False
    skip_chapter_three_adjustments: bool = False
    skip_log_output: bool = True
    skip_nested_tables: bool = True
    move_table_notes_below: bool = False
    skip_chapter_three_table_notes: bool = True
    force_note_paragraph_left_alignment: bool = False
    enable_double_black_table_borders: bool = False
    enable_table_footer_source_format: bool = False
    # Developer-only diagnostic: when True, fix_docx_fast writes a
    # *_note_debug_log.txt next to the output. Default False so normal GUI/CLI
    # runs never produce the (temp-named) __tmp___note_debug_log.txt file.
    write_note_debug_log: bool = False
    table_keep_colors: tuple[str, ...] = ("D9D9D9", "F2F2F2")
    table_gray_colors: tuple[str, ...] = ("BFBFBF", "C0C0C0", "A6A6A6", "808080")
    table_gray_target: str = "D9D9D9"
    skip_special_color_tables: bool = False
    special_color_skip_colors: tuple[str, ...] = ()
    clear_special_colors_after_skip: bool = False

    def __post_init__(self) -> None:
        # Backward compatibility for older callers. The GUI no longer exposes
        # the combined "skip all" option, and skip_chapter_three_tables is a
        # deprecated alias for skipping both table layout and color.
        if self.skip_chapter_three_tables:
            self.skip_chapter_three_table_layout = True
            self.skip_chapter_three_table_color = True

        # "參、不要調整" (skip_chapter_three_adjustments) is a legacy alias that
        # expands to the granular chapter-three skips below. Table footer-source
        # formatting has its own eligibility check and is not controlled here.
        if self.skip_all_under_chapter_three or self.skip_chapter_three_adjustments:
            self.skip_chapter_three_table_layout = True
            self.skip_chapter_three_table_color = True
            self.skip_chapter_three_indents = True


@dataclass
class ProcessSummary:
    tables: int = 0
    skipped_first_page_tables: int = 0
    skipped_small_tables: int = 0
    skipped_nested_tables: int = 0
    nested_table_color_only_tables: int = 0
    special_color_skipped_tables: int = 0
    section_three_protected_tables: int = 0
    double_border_tables: int = 0
    table_footer_source_format_tables: int = 0
    note_cells_moved_tables: int = 0
    note_move_skipped_by_chapter_three_tables: int = 0
    moved_note_count: int = 0
    deleted_note_cells: int = 0
    deleted_note_rows: int = 0
    inserted_note_paragraphs: int = 0
    special_autofit_right_tables: int = 0
    normal_processed_tables: int = 0
    cross_page_tables: int = 0
    cross_page_resolved_tables: int = 0
    cross_page_still_split_tables: int = 0
    adjusted_cell_padding_tables: int = 0
    adjusted_table_spacing_tables: int = 0
    auto_height_tables: int = 0
    moved_next_page_resolved_tables: int = 0
    cannot_avoid_cross_page_tables: int = 0
    failed_cross_page_tables: int = 0
    changed_to_gray: int = 0
    cleared_colors: int = 0
    paragraphs: int = 0
    total_paragraphs: int = 0
    skipped_toc_paragraphs: int = 0
    skipped_table_paragraphs: int = 0
    removed_all_outline_paragraphs: int = 0
    indented_preface_paragraphs: int = 0
    outlined_preface_paragraphs: int = 0
    unknown_paragraphs: int = 0
    paragraph_level_counts: list[int] = field(default_factory=lambda: [0] * 9)
    paragraph_logs: list[str] = field(default_factory=list)
    numbering_measurements: dict[str, dict[str, object]] = field(default_factory=dict)
    numbering_xml_logs: list[str] = field(default_factory=list)
    numbering_debug_logs: list[str] = field(default_factory=list)
    body_indent_debug_logs: list[str] = field(default_factory=list)
    body_indent_records: list[dict[str, object]] = field(default_factory=list)
    table_log_records: list[dict[str, object]] = field(default_factory=list)
    table_footer_source_format_records: list[dict[str, object]] = field(default_factory=list)
    table_footer_source_format_logs: list[str] = field(default_factory=list)
    word_com_table_autofit_records: list[dict[str, object]] = field(default_factory=list)
    word_com_table_autofit_logs: list[str] = field(default_factory=list)
    heading_suffix_before_records: list[dict[str, object]] = field(default_factory=list)
    heading_suffix_after_records: list[dict[str, object]] = field(default_factory=list)
    word_com_body_indent_logs: list[str] = field(default_factory=list)
    character_indent_attrs_removed: int = 0

    @property
    def changed_colors(self) -> int:
        return self.changed_to_gray + self.cleared_colors

    def _count_word_com_autofit_status(self, status: str) -> int:
        return sum(
            1
            for record in self.table_log_records
            if record.get("word_com_autofit_status") == status
        )

    @property
    def word_com_table_autofit_applied_count(self) -> int:
        return self._count_word_com_autofit_status("word_com")

    @property
    def word_com_table_autofit_fallback_count(self) -> int:
        return self._count_word_com_autofit_status("xml_fallback")

    @property
    def word_com_table_autofit_failed_count(self) -> int:
        return self._count_word_com_autofit_status("failed")
