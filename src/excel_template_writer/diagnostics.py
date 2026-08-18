"""Stable diagnostics shared by compilation and rendering."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from enum import StrEnum


class DiagnosticCode(StrEnum):
    UNTERMINATED_EXPRESSION = "E1001"
    UNTERMINATED_DIRECTIVE = "E1002"
    EMPTY_EXPRESSION = "E1003"
    INVALID_EXPRESSION = "E1101"
    INVALID_DIRECTIVE = "E1102"
    UNKNOWN_DIRECTIVE_OPTION = "E1103"
    UNMATCHED_BLOCK_MARKER = "E1201"
    INVALID_BLOCK_GEOMETRY = "E1202"
    AMBIGUOUS_BLOCK_PAIRING = "E1203"
    INVALID_MARKER_POSITION = "E1204"
    PARTIAL_BLOCK_OVERLAP = "E1205"
    MISSING_VALUE = "E1301"
    COLLECTION_IN_SCALAR_CELL = "E1302"
    EXPECTED_COLLECTION = "E1303"
    LAYOUT_COLLISION = "E1401"
    OVERLAPPING_ROW_SHIFTS = "E1402"
    MERGE_CROSSES_BLOCK_BOUNDARY = "E2104"
    CONDITIONAL_FORMATTING_REQUIRES_UNSUPPORTED_TRANSFORM = "E2105"
    DATA_VALIDATION_REQUIRES_UNSUPPORTED_TRANSFORM = "E2106"
    TABLE_REQUIRES_UNSUPPORTED_TRANSFORM = "E2107"
    DRAWING_REQUIRES_UNSUPPORTED_TRANSFORM = "E2108"
    HYPERLINK_REQUIRES_UNSUPPORTED_TRANSFORM = "E2109"
    COMMENT_REQUIRES_UNSUPPORTED_TRANSFORM = "E2110"
    CELL_SHIFT_WITH_CUSTOM_ROW_HEIGHT = "E2111"
    FORMULA_REQUIRES_UNSUPPORTED_TRANSFORM = "E3101"
    INPUT_OUTPUT_PATH_CONFLICT = "E3201"
    UNSUPPORTED_WORKBOOK_FORMAT = "E3202"


@dataclass(frozen=True)
class SourceLocation:
    sheet: str
    cell: str
    start: int | None = None
    end: int | None = None

    def __str__(self) -> str:
        suffix = "" if self.start is None else f":{self.start}"
        return f"{self.sheet}!{self.cell}{suffix}"


@dataclass(frozen=True)
class Diagnostic:
    code: DiagnosticCode
    message: str
    location: SourceLocation

    def __str__(self) -> str:
        return f"{self.code} {self.location}: {self.message}"


class TemplateError(Exception):
    """Base exception carrying structured diagnostics."""

    def __init__(self, diagnostics: Iterable[Diagnostic]) -> None:
        self.diagnostics = tuple(diagnostics)
        super().__init__("\n".join(str(item) for item in self.diagnostics))


class TemplateCompilationError(TemplateError):
    pass


class TemplateRenderError(TemplateError):
    pass
