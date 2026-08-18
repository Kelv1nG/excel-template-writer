"""Spatial template interpreter for XLSX workbooks."""

from excel_template_writer.compiler import CompilationResult, compile_sheet
from excel_template_writer.model import Coordinate, Rectangle, WorksheetTemplate
from excel_template_writer.render import RenderPlan, RenderResult, render_sheet
from excel_template_writer.values import (
    CanonicalContext,
    CanonicalValue,
    InputValue,
    NormalizationResult,
    NormalizedContext,
    ScalarValue,
    TypeAdapter,
    normalize_context,
    validate_context,
)

__all__ = [
    "CanonicalContext",
    "CanonicalValue",
    "CompilationResult",
    "Coordinate",
    "InputValue",
    "NormalizationResult",
    "NormalizedContext",
    "Rectangle",
    "RenderPlan",
    "RenderResult",
    "ScalarValue",
    "TypeAdapter",
    "WorksheetTemplate",
    "compile_sheet",
    "normalize_context",
    "render_sheet",
    "validate_context",
]

__version__ = "0.1.0"
