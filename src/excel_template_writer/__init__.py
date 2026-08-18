"""Spatial template interpreter for XLSX workbooks."""

from excel_template_writer.compiler import CompilationResult, compile_sheet
from excel_template_writer.model import Coordinate, Rectangle, WorksheetTemplate
from excel_template_writer.render import RenderPlan, RenderResult, render_sheet
from excel_template_writer.values import (
    CanonicalContext,
    CanonicalValue,
    ScalarValue,
    validate_context,
)

__all__ = [
    "CanonicalContext",
    "CanonicalValue",
    "CompilationResult",
    "Coordinate",
    "Rectangle",
    "RenderPlan",
    "RenderResult",
    "ScalarValue",
    "WorksheetTemplate",
    "compile_sheet",
    "render_sheet",
    "validate_context",
]

__version__ = "0.1.0"
