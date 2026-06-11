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
    skip_all_under_chapter_three: bool = False

    def __post_init__(self) -> None:
        # Backward compatibility for older callers. The GUI no longer exposes
        # the combined "skip all" option, and skip_chapter_three_tables is a
        # deprecated alias for skipping both table layout and color.
        if self.skip_chapter_three_tables:
            self.skip_chapter_three_table_layout = True
            self.skip_chapter_three_table_color = True

        if self.skip_all_under_chapter_three:
            self.skip_chapter_three_table_layout = True
            self.skip_chapter_three_table_color = True
            self.skip_chapter_three_indents = True


@dataclass
class ProcessSummary:
    tables: int = 0
    skipped_first_page_tables: int = 0
    skipped_small_tables: int = 0
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
    heading_suffix_before_records: list[dict[str, object]] = field(default_factory=list)
    heading_suffix_after_records: list[dict[str, object]] = field(default_factory=list)
    word_com_body_indent_logs: list[str] = field(default_factory=list)
    character_indent_attrs_removed: int = 0

    @property
    def changed_colors(self) -> int:
        return self.changed_to_gray + self.cleared_colors
