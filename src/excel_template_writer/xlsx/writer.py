"""Apply complete render plans to a newly constructed XLSX workbook."""

from __future__ import annotations

import os
import tempfile
from copy import copy, deepcopy
from pathlib import Path
from typing import Any
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.drawing.image import Image
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor, OneCellAnchor, TwoCellAnchor
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.dimensions import ColumnDimension, RowDimension
from openpyxl.worksheet.worksheet import Worksheet

from excel_template_writer.diagnostics import TemplateRenderError
from excel_template_writer.limits import ResourceLimits
from excel_template_writer.model import Coordinate
from excel_template_writer.render import RenderPlan
from excel_template_writer.xlsx.model import (
    CellPresentation,
    ChartSnapshot,
    ColumnPresentation,
    DimensionPresentation,
    ImageSnapshot,
    RowPresentation,
    SheetFeaturePlan,
    SheetSnapshot,
    WorkbookSnapshot,
)
from excel_template_writer.xlsx.package_limits import inspect_xlsx_package

_CHART_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_DRAWING_MAIN_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main"
_OFFICE_RELATIONSHIP_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)


class _PreservedImage(Image):
    def __init__(self, data: bytes, image_format: str, anchor: Any) -> None:
        """Create an openpyxl-compatible image without decoding its media bytes.

        Args:
            data: Exact embedded media payload from the source XLSX package.
            image_format: Supported file extension used for the output media part.
            anchor: Detached DrawingML anchor containing the source picture frame.
        """

        self._payload = data
        self.format = image_format
        self.anchor = anchor

    def _data(self) -> bytes:
        """Return the exact source media payload for package serialization."""

        return self._payload


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


def _apply_drawing_anchor(anchor: Any, planned_coordinates: tuple[Coordinate, ...]) -> None:
    """Apply one prevalidated translation to a detached drawing anchor.

    Args:
        anchor: Mutable openpyxl drawing anchor.
        planned_coordinates: Destination coordinates for its cell anchor markers.

    Raises:
        RuntimeError: If the validated anchor and feature plan disagree.
    """

    if isinstance(anchor, AbsoluteAnchor):
        if planned_coordinates:
            raise RuntimeError("absolute drawing anchor unexpectedly has cell markers")
        return
    if isinstance(anchor, OneCellAnchor):
        markers = (anchor._from,)
    elif isinstance(anchor, TwoCellAnchor):
        markers = (anchor._from, anchor.to)
    else:
        raise RuntimeError("unsupported drawing anchor reached the workbook writer")
    if len(markers) != len(planned_coordinates):
        raise RuntimeError("drawing anchor marker count does not match its feature plan")
    for marker, coordinate in zip(markers, planned_coordinates, strict=True):
        marker.row = coordinate.row - 1
        marker.col = coordinate.column - 1


def _drawing_relationship_id(anchor: ElementTree.Element) -> str | None:
    """Return the chart or image relationship ID used by one drawing anchor.

    Args:
        anchor: Serialized anchor element from an output drawing part.

    Returns:
        Embedded relationship ID, or ``None`` for an unexpected drawing object.
    """

    chart = anchor.find(f".//{{{_CHART_NAMESPACE}}}chart")
    if chart is not None:
        return chart.get(f"{{{_OFFICE_RELATIONSHIP_NAMESPACE}}}id")
    blip = anchor.find(f".//{{{_DRAWING_MAIN_NAMESPACE}}}blip")
    if blip is not None:
        return blip.get(f"{{{_OFFICE_RELATIONSHIP_NAMESPACE}}}embed")
    return None


def _desired_relationship_order(source: SheetSnapshot) -> tuple[str, ...]:
    """Return output relationship IDs in the source drawing stacking order.

    Args:
        source: Worksheet snapshot containing ordered charts and images.

    Returns:
        Relationship IDs assigned by openpyxl, reordered to the source sequence.
    """

    chart_count = sum(isinstance(drawing, ChartSnapshot) for drawing in source.drawings)
    chart_index = 0
    image_index = 0
    relationship_ids: list[str] = []
    for drawing in source.drawings:
        if isinstance(drawing, ChartSnapshot):
            chart_index += 1
            relationship_ids.append(f"rId{chart_index}")
        else:
            image_index += 1
            relationship_ids.append(f"rId{chart_count + image_index}")
    return tuple(relationship_ids)


def _restore_drawing_order(path: Path, snapshot: WorkbookSnapshot) -> None:
    """Restore source stacking order after openpyxl groups drawing objects.

    Args:
        path: Serialized temporary XLSX package to rewrite atomically.
        snapshot: Source workbook with ordered drawing snapshots.

    Raises:
        RuntimeError: If serialized drawing relationships do not match the feature plan.
    """

    if not any(sheet.drawings for sheet in snapshot.sheets):
        return
    handle, repacked_name = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(handle)
    repacked_path = Path(repacked_name)
    try:
        with ZipFile(path, "r") as archive:
            modified_parts: dict[str, bytes] = {}
            drawing_index = 0
            archive_names = frozenset(archive.namelist())
            for sheet in snapshot.sheets:
                if not sheet.drawings:
                    continue
                drawing_index += 1
                part_name = f"xl/drawings/drawing{drawing_index}.xml"
                if part_name not in archive_names:
                    raise RuntimeError("serialized worksheet drawing part is missing")
                root = ElementTree.fromstring(archive.read(part_name))
                anchors_by_relationship: dict[str, ElementTree.Element] = {}
                for anchor in root:
                    relationship_id = _drawing_relationship_id(anchor)
                    if relationship_id is None or relationship_id in anchors_by_relationship:
                        raise RuntimeError("serialized worksheet drawing order is ambiguous")
                    anchors_by_relationship[relationship_id] = anchor
                desired_order = _desired_relationship_order(sheet)
                if set(anchors_by_relationship) != set(desired_order):
                    raise RuntimeError(
                        "serialized worksheet drawings do not match their feature plan"
                    )
                root[:] = [
                    anchors_by_relationship[relationship_id] for relationship_id in desired_order
                ]
                modified_parts[part_name] = ElementTree.tostring(root, encoding="utf-8")

            with ZipFile(repacked_path, "w") as destination:
                destination.comment = archive.comment
                for member in archive.infolist():
                    destination.writestr(
                        member,
                        modified_parts.get(member.filename, archive.read(member.filename)),
                    )
        os.replace(repacked_path, path)
    finally:
        repacked_path.unlink(missing_ok=True)


def _write_sheet(
    destination: Worksheet,
    source: SheetSnapshot,
    plan: RenderPlan,
    feature_plan: SheetFeaturePlan,
) -> None:
    """Apply one complete validated plan to a new worksheet.

    Args:
        destination: Empty mutable destination worksheet.
        source: Detached source worksheet state.
        plan: Complete adapter-neutral render plan.
        feature_plan: Validated adapter plan for charts and other workbook features.
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
        if planned_cell.source_coordinate in source.synthetic_drawing_anchor_cells:
            continue
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

    for source_drawing, planned_drawing in zip(
        source.drawings,
        feature_plan.drawings,
        strict=True,
    ):
        if isinstance(source_drawing, ChartSnapshot):
            chart = deepcopy(source_drawing.chart)
            _apply_drawing_anchor(chart.anchor, planned_drawing.anchor_coordinates)
            destination.add_chart(chart)
        else:
            assert isinstance(source_drawing, ImageSnapshot)
            anchor = deepcopy(source_drawing.anchor)
            _apply_drawing_anchor(anchor, planned_drawing.anchor_coordinates)
            destination.add_image(
                _PreservedImage(
                    source_drawing.data,
                    source_drawing.image_format,
                    anchor,
                )
            )

    if _is_identity_plan(source, plan):
        destination.conditional_formatting = copy(source.conditional_formatting)
        destination.data_validations = copy(source.data_validations)
        destination.auto_filter = copy(source.auto_filter)


def write_workbook(
    snapshot: WorkbookSnapshot,
    plans: tuple[RenderPlan, ...],
    feature_plans: tuple[SheetFeaturePlan, ...],
    output_path: str | Path,
    *,
    limits: ResourceLimits,
) -> Path:
    """Write atomically and reopen the package before publishing it.

    Args:
        snapshot: Detached source workbook state.
        plans: One complete render plan per source worksheet.
        feature_plans: One validated workbook-feature plan per source worksheet.
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
    for sheet, plan, feature_plan in zip(
        snapshot.sheets,
        plans,
        feature_plans,
        strict=True,
    ):
        destination = workbook.create_sheet(sheet.template.name)
        _write_sheet(destination, sheet, plan, feature_plan)

    handle, temporary_name = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(handle)
    temporary_path = Path(temporary_name)
    try:
        workbook.save(temporary_path)
        _restore_drawing_order(temporary_path, snapshot)
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
