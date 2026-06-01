"""Word DOCX quick fixer package."""

from .models import ProcessOptions, ProcessSummary
from .docx_processor import fix_docx_fast

__all__ = ["ProcessOptions", "ProcessSummary", "fix_docx_fast"]
