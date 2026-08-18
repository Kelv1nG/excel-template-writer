"""Production XLSX adapter built around the pure compiler and render plan."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excel_template_writer.compiler import compile_sheet
from excel_template_writer.diagnostics import (
    Diagnostic,
    DiagnosticCode,
    SourceLocation,
    TemplateCompilationError,
)
from excel_template_writer.render import RenderPlan, render_sheet
from excel_template_writer.xlsx.reader import read_workbook
from excel_template_writer.xlsx.validation import validate_sheet_features
from excel_template_writer.xlsx.writer import write_workbook


@dataclass(frozen=True)
class WorkbookRenderResult:
    output_path: Path
    diagnostics: tuple[Diagnostic, ...]


def render_workbook(
    template_path: str | Path,
    output_path: str | Path,
    context: Mapping[str, Any],
) -> WorkbookRenderResult:
    """Compile and plan all sheets before atomically writing a separate XLSX file."""

    source_path = Path(template_path).resolve()
    destination_path = Path(output_path).resolve()
    invalid_path = next(
        (path for path in (source_path, destination_path) if path.suffix.lower() != ".xlsx"),
        None,
    )
    if invalid_path is not None:
        raise TemplateCompilationError(
            (
                Diagnostic(
                    DiagnosticCode.UNSUPPORTED_WORKBOOK_FORMAT,
                    "only .xlsx input and output files are supported",
                    SourceLocation("<workbook>", "A1"),
                ),
            )
        )
    if source_path == destination_path:
        raise TemplateCompilationError(
            (
                Diagnostic(
                    DiagnosticCode.INPUT_OUTPUT_PATH_CONFLICT,
                    "the output path must be different from the input template path",
                    SourceLocation("<workbook>", "A1"),
                ),
            )
        )

    snapshot = read_workbook(source_path)
    plans: list[RenderPlan] = []
    diagnostics: list[Diagnostic] = []
    for sheet in snapshot.sheets:
        compilation = compile_sheet(sheet.template)
        if compilation.compiled is None:
            diagnostics.extend(compilation.diagnostics)
            continue
        rendering = render_sheet(compilation.compiled, context)
        if rendering.plan is None:
            diagnostics.extend(rendering.diagnostics)
            continue
        diagnostics.extend(validate_sheet_features(sheet, compilation.compiled, rendering.plan))
        plans.append(rendering.plan)
    if diagnostics:
        raise TemplateCompilationError(diagnostics)
    if len(plans) != len(snapshot.sheets):
        raise RuntimeError("internal error: not every worksheet produced a render plan")

    written_path = write_workbook(snapshot, tuple(plans), destination_path)
    return WorkbookRenderResult(written_path, ())


__all__ = ["WorkbookRenderResult", "render_workbook"]
