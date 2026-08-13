"""Spatial template interpreter for XLSX workbooks."""

from excel_template_writer.compiler import CompilationResult, compile_sheet
from excel_template_writer.model import Coordinate, Rectangle, WorksheetTemplate
from excel_template_writer.render import RenderPlan, RenderResult, render_sheet

__all__ = [
    "CompilationResult",
    "Coordinate",
    "Rectangle",
    "RenderPlan",
    "RenderResult",
    "WorksheetTemplate",
    "compile_sheet",
    "render_sheet",
]

__version__ = "0.1.0"
