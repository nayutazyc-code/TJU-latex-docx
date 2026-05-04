"""LaTeX to DOCX desktop converter."""

from .converter import ConversionConfig, ConversionResult, convert_project
from .scanner import find_main_tex_candidates

__all__ = [
    "ConversionConfig",
    "ConversionResult",
    "convert_project",
    "find_main_tex_candidates",
]
