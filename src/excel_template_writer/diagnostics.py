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
    CONTEXT_MUST_BE_MAPPING = "E1501"
    CONTEXT_KEY_MUST_BE_STRING = "E1502"
    UNSUPPORTED_CONTEXT_VALUE = "E1503"
    UNORDERED_CONTEXT_COLLECTION = "E1504"
    NON_FINITE_CONTEXT_NUMBER = "E1505"
    CYCLIC_CONTEXT_VALUE = "E1506"
    TIMEZONE_AWARE_CONTEXT_VALUE = "E1507"
    DUPLICATE_VALUE_ADAPTER = "E1510"
    AMBIGUOUS_VALUE_ADAPTER = "E1511"
    VALUE_ADAPTER_FAILED = "E1512"
    VALUE_ADAPTER_CYCLE = "E1513"
    CONTEXT_RESOURCE_LIMIT_EXCEEDED = "E1601"
    RENDER_RESOURCE_LIMIT_EXCEEDED = "E1602"
    XLSX_GRID_LIMIT_EXCEEDED = "E1603"
    CELL_TEXT_LIMIT_EXCEEDED = "E1604"
    MERGE_CROSSES_BLOCK_BOUNDARY = "E2104"
    CONDITIONAL_FORMATTING_REQUIRES_UNSUPPORTED_TRANSFORM = "E2105"
    DATA_VALIDATION_REQUIRES_UNSUPPORTED_TRANSFORM = "E2106"
    TABLE_REQUIRES_UNSUPPORTED_TRANSFORM = "E2107"
    DRAWING_REQUIRES_UNSUPPORTED_TRANSFORM = "E2108"
    HYPERLINK_REQUIRES_UNSUPPORTED_TRANSFORM = "E2109"
    COMMENT_REQUIRES_UNSUPPORTED_TRANSFORM = "E2110"
    CELL_SHIFT_WITH_CUSTOM_ROW_HEIGHT = "E2111"
    XLSX_PACKAGE_LIMIT_EXCEEDED = "E2201"
    WORKSHEET_COUNT_LIMIT_EXCEEDED = "E2202"
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
        """Return the worksheet and cell location in diagnostic notation."""

        suffix = "" if self.start is None else f":{self.start}"
        return f"{self.sheet}!{self.cell}{suffix}"


@dataclass(frozen=True)
class ContextLocation:
    path: str

    def __str__(self) -> str:
        """Return the canonical context path."""

        return self.path


@dataclass(frozen=True)
class Diagnostic:
    code: DiagnosticCode
    message: str
    location: SourceLocation | ContextLocation

    def __str__(self) -> str:
        """Format the code, location, and message for human-readable output."""

        return f"{self.code} {self.location}: {self.message}"


class TemplateError(Exception):
    """Base exception carrying structured diagnostics."""

    def __init__(self, diagnostics: Iterable[Diagnostic]) -> None:
        """Create an exception from one or more structured diagnostics.

        Args:
            diagnostics: Diagnostics to retain and join into the exception message.
        """

        self.diagnostics = tuple(diagnostics)
        super().__init__("\n".join(str(item) for item in self.diagnostics))


class TemplateCompilationError(TemplateError):
    pass


class TemplateRenderError(TemplateError):
    pass
