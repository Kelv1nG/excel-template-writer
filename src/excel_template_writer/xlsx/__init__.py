"""Production XLSX adapter built around the pure compiler and render plan."""

from __future__ import annotations

from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excel_template_writer.compiler import compile_sheet
from excel_template_writer.diagnostics import (
    Diagnostic,
    DiagnosticCode,
    SourceLocation,
    TemplateCompilationError,
    TemplateRenderError,
)
from excel_template_writer.limits import DEFAULT_RESOURCE_LIMITS, ResourceLimits
from excel_template_writer.render import RenderPlan, render_sheet
from excel_template_writer.values import TypeAdapter, normalize_context
from excel_template_writer.xlsx.model import SheetFeaturePlan
from excel_template_writer.xlsx.package_limits import inspect_xlsx_package
from excel_template_writer.xlsx.reader import read_workbook
from excel_template_writer.xlsx.validation import plan_sheet_features
from excel_template_writer.xlsx.writer import write_workbook


@dataclass(frozen=True)
class WorkbookRenderResult:
    output_path: Path
    diagnostics: tuple[Diagnostic, ...]


def render_workbook(
    template_path: str | Path,
    output_path: str | Path,
    context: object,
    *,
    adapters: Iterable[TypeAdapter[Any]] = (),
    limits: ResourceLimits = DEFAULT_RESOURCE_LIMITS,
) -> WorkbookRenderResult:
    """Compile and plan all sheets before writing a separate XLSX file.

    Args:
        template_path: Existing input ``.xlsx`` template path.
        output_path: Separate destination ``.xlsx`` path.
        context: Raw render-context mapping or normalized context.
        adapters: Explicit converters for caller-owned runtime types.
        limits: Resource ceilings shared across the complete operation.

    Returns:
        The published output path and operation diagnostics.

    Raises:
        TemplateCompilationError: If paths, syntax, geometry, or workbook features are invalid.
        TemplateRenderError: If normalization, evaluation, or resource planning fails.
    """

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

    package_diagnostic = inspect_xlsx_package(
        source_path,
        limits,
        description="input XLSX package",
    )
    if package_diagnostic is not None:
        raise TemplateCompilationError((package_diagnostic,))

    normalized_context = normalize_context(
        context,
        adapters=adapters,
        limits=limits,
    ).require()

    snapshot = read_workbook(source_path)
    total_sheets = len(snapshot.sheets) + len(snapshot.chartsheets)
    if total_sheets > limits.max_worksheets:
        raise TemplateCompilationError(
            (
                Diagnostic(
                    DiagnosticCode.WORKSHEET_COUNT_LIMIT_EXCEEDED,
                    f"workbook exceeds max_worksheets={limits.max_worksheets:,}",
                    SourceLocation("<workbook>", "A1"),
                ),
            )
        )
    if snapshot.chartsheets:
        raise TemplateCompilationError(
            (
                Diagnostic(
                    DiagnosticCode.CHARTSHEET_UNSUPPORTED,
                    "chart sheets are not supported; embed the chart in a worksheet",
                    SourceLocation(snapshot.chartsheets[0], "A1"),
                ),
            )
        )
    worksheet_names = frozenset(sheet.template.name for sheet in snapshot.sheets)
    plans: list[RenderPlan] = []
    feature_plans: list[SheetFeaturePlan] = []
    diagnostics: list[Diagnostic] = []
    planned_cells = 0
    resource_codes = {
        DiagnosticCode.RENDER_RESOURCE_LIMIT_EXCEEDED,
        DiagnosticCode.XLSX_GRID_LIMIT_EXCEEDED,
        DiagnosticCode.CELL_TEXT_LIMIT_EXCEEDED,
    }
    for sheet in snapshot.sheets:
        compilation = compile_sheet(sheet.template)
        if compilation.compiled is None:
            diagnostics.extend(compilation.diagnostics)
            continue
        rendering = render_sheet(compilation.compiled, normalized_context, limits=limits)
        if rendering.plan is None:
            if any(diagnostic.code in resource_codes for diagnostic in rendering.diagnostics):
                raise TemplateRenderError(rendering.diagnostics)
            diagnostics.extend(rendering.diagnostics)
            continue
        planned_cells += len(rendering.plan.cells)
        if planned_cells > limits.max_planned_cells_per_workbook:
            raise TemplateRenderError(
                (
                    Diagnostic(
                        DiagnosticCode.RENDER_RESOURCE_LIMIT_EXCEEDED,
                        "workbook exceeds max_planned_cells_per_workbook="
                        f"{limits.max_planned_cells_per_workbook:,}",
                        SourceLocation(sheet.template.name, "A1"),
                    ),
                )
            )
        feature_plan, feature_diagnostics = plan_sheet_features(
            sheet,
            compilation.compiled,
            rendering.plan,
            worksheet_names=worksheet_names,
        )
        feature_plans.append(feature_plan)
        diagnostics.extend(feature_diagnostics)
        plans.append(rendering.plan)
    if diagnostics:
        raise TemplateCompilationError(diagnostics)
    if len(plans) != len(snapshot.sheets):
        raise RuntimeError("internal error: not every worksheet produced a render plan")
    if len(feature_plans) != len(snapshot.sheets):
        raise RuntimeError("internal error: not every worksheet produced a feature plan")

    written_path = write_workbook(
        snapshot,
        tuple(plans),
        tuple(feature_plans),
        destination_path,
        limits=limits,
    )
    return WorkbookRenderResult(written_path, ())


__all__ = ["WorkbookRenderResult", "render_workbook"]
