from __future__ import annotations

from dataclasses import dataclass, field

@dataclass
class ProcessOptions:
    fix_table_layout: bool
    fix_color: bool
    fix_paragraph: bool
    include_tables_in_paragraph: bool
    remove_preface_outline: bool = False


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
    removed_preface_outline_paragraphs: int = 0
    unknown_paragraphs: int = 0
    paragraph_level_counts: list[int] = field(default_factory=lambda: [0] * 9)
    paragraph_logs: list[str] = field(default_factory=list)

    @property
    def changed_colors(self) -> int:
        return self.changed_to_gray + self.cleared_colors
