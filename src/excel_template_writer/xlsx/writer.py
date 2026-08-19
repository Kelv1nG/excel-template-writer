"""Apply complete render plans to a newly constructed XLSX workbook."""

from __future__ import annotations

import os
import tempfile
from copy import copy
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension, RowDimension
from openpyxl.worksheet.worksheet import Worksheet

from excel_template_writer.diagnostics import TemplateRenderError
from excel_template_writer.limits import ResourceLimits
from excel_template_writer.render import RenderPlan
from excel_template_writer.xlsx.model import (
    CellPresentation,
    ColumnPresentation,
    DimensionPresentation,
    RowPresentation,
    SheetSnapshot,
    WorkbookSnapshot,
)
from excel_template_writer.xlsx.package_limits import inspect_xlsx_package


def _apply_dimension_style(
    destination: RowDimension | ColumnDimension,
    source: DimensionPresentation,
) -> None:
    """Copy formatting shared by row and column dimensions.

    Args:
        destination: Mutable openpyxl dimension receiving formatting.
        source: Detached source dimension presentation.
    """

    destination.font = copy(source.font)
    destination.fill = copy(source.fill)
    destination.border = copy(source.border)
    destination.alignment = copy(source.alignment)
    destination.number_format = source.number_format
    destination.protection = copy(source.protection)


def _apply_row(destination: RowDimension, source: RowPresentation) -> None:
    """Apply detached row presentation to an output row dimension.

    Args:
        destination: Mutable output row dimension.
        source: Detached source row presentation.
    """

    destination.height = source.height
    destination.hidden = source.hidden
    destination.outlineLevel = source.outline_level
    destination.collapsed = source.collapsed
    destination.thickTop = source.thick_top
    destination.thickBot = source.thick_bottom
    _apply_dimension_style(destination, source)


def _apply_column(destination: ColumnDimension, source: ColumnPresentation) -> None:
    """Apply detached column presentation to an output column dimension.

    Args:
        destination: Mutable output column dimension.
        source: Detached source column presentation.
    """

    destination.width = source.width
    destination.hidden = source.hidden
    destination.outlineLevel = source.outline_level
    destination.collapsed = source.collapsed
    destination.bestFit = source.best_fit
    _apply_dimension_style(destination, source)


def _apply_cell(destination: Cell, source: CellPresentation, value: object) -> None:
    """Write one planned value and copy its direct source presentation.

    Args:
        destination: Mutable output cell.
        source: Detached source cell presentation.
        value: Evaluated planned value to write.
    """

    destination.value = value
    if isinstance(value, str) and value.startswith("=") and not source.is_formula:
        destination.data_type = "s"
    destination.font = copy(source.font)
    destination.fill = copy(source.fill)
    destination.border = copy(source.border)
    destination.alignment = copy(source.alignment)
    destination.number_format = source.number_format
    destination.protection = copy(source.protection)
    destination.quotePrefix = source.quote_prefix
    destination.hyperlink = copy(source.hyperlink)
    destination.comment = copy(source.comment)


def _is_identity_plan(source: SheetSnapshot, plan: RenderPlan) -> bool:
    """Return whether a plan leaves every supported coordinate unchanged.

    Args:
        source: Source worksheet snapshot.
        plan: Completed render plan.

    Returns:
        ``True`` when cell and merge coordinates are identical.
    """

    cell_mappings = [(cell.source_coordinate, cell.coordinate) for cell in plan.cells]
    source_merges = set(source.template.merged_ranges)
    planned_merges = {
        merge.rectangle for merge in plan.merges if merge.source_rectangle == merge.rectangle
    }
    return (
        len(cell_mappings) == len(source.cells)
        and all(
            source_coordinate == destination for source_coordinate, destination in cell_mappings
        )
        and source_merges == planned_merges
    )


def _write_sheet(destination: Worksheet, source: SheetSnapshot, plan: RenderPlan) -> None:
    """Apply one complete validated plan to a new worksheet.

    Args:
        destination: Empty mutable destination worksheet.
        source: Detached source worksheet state.
        plan: Complete adapter-neutral render plan.
    """

    destination.sheet_view.showGridLines = source.show_grid_lines
    destination.freeze_panes = source.freeze_panes
    destination.sheet_properties.tabColor = copy(source.tab_color)
    destination.sheet_format.defaultRowHeight = source.default_row_height
    destination.sheet_format.defaultColWidth = source.default_column_width

    for column, presentation in source.columns.items():
        _apply_column(destination.column_dimensions[get_column_letter(column)], presentation)
    for planned_row in plan.rows:
        if planned_row.source_row is None:
            continue
        presentation = source.rows.get(planned_row.source_row)
        if presentation is not None:
            _apply_row(destination.row_dimensions[planned_row.destination_row], presentation)

    for planned_cell in plan.cells:
        presentation = source.cells[planned_cell.source_coordinate]
        cell = destination.cell(planned_cell.coordinate.row, planned_cell.coordinate.column)
        _apply_cell(cell, presentation, planned_cell.value)

    for merge in plan.merges:
        destination.merge_cells(
            start_row=merge.rectangle.top,
            start_column=merge.rectangle.left,
            end_row=merge.rectangle.bottom,
            end_column=merge.rectangle.right,
        )

    if _is_identity_plan(source, plan):
        destination.conditional_formatting = copy(source.conditional_formatting)
        destination.data_validations = copy(source.data_validations)
        destination.auto_filter = copy(source.auto_filter)


def write_workbook(
    snapshot: WorkbookSnapshot,
    plans: tuple[RenderPlan, ...],
    output_path: str | Path,
    *,
    limits: ResourceLimits,
) -> Path:
    """Write atomically and reopen the package before publishing it.

    Args:
        snapshot: Detached source workbook state.
        plans: One complete render plan per source worksheet.
        output_path: Destination path, which must differ from the template path.
        limits: Package ceilings checked before publication.

    Returns:
        Resolved destination path after successful verification.

    Raises:
        TemplateRenderError: If the serialized package exceeds resource limits.
    """

    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    workbook.remove(workbook.active)
    workbook.properties = copy(snapshot.properties)
    workbook.loaded_theme = snapshot.loaded_theme
    for sheet, plan in zip(snapshot.sheets, plans, strict=True):
        destination = workbook.create_sheet(sheet.template.name)
        _write_sheet(destination, sheet, plan)

    handle, temporary_name = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(handle)
    temporary_path = Path(temporary_name)
    try:
        workbook.save(temporary_path)
        package_diagnostic = inspect_xlsx_package(
            temporary_path,
            limits,
            description="rendered XLSX package",
        )
        if package_diagnostic is not None:
            raise TemplateRenderError((package_diagnostic,))
        verified = load_workbook(temporary_path, read_only=True, data_only=False)
        verified.close()
        os.replace(temporary_path, path)
    finally:
        workbook.close()
        temporary_path.unlink(missing_ok=True)
    return path
