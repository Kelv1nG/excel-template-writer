"""Workbook-feature checks that depend on both the AST and the render plan."""

from __future__ import annotations

from collections import defaultdict

from openpyxl.utils.cell import SHEETRANGE_RE, range_to_tuple

from excel_template_writer.ast import CompiledSheet, ForNode, StructuralNode
from excel_template_writer.diagnostics import Diagnostic, DiagnosticCode, SourceLocation
from excel_template_writer.limits import XLSX_MAX_COLUMNS, XLSX_MAX_ROWS
from excel_template_writer.model import Coordinate
from excel_template_writer.render import RenderPlan
from excel_template_writer.xlsx.model import (
    ChartSnapshot,
    ImageSnapshot,
    PlannedDrawing,
    SheetFeaturePlan,
    SheetSnapshot,
)


def _location(sheet: SheetSnapshot, coordinate: Coordinate | None = None) -> SourceLocation:
    """Create a worksheet diagnostic location with an optional coordinate.

    Args:
        sheet: Owning worksheet snapshot.
        coordinate: Affected coordinate, or ``None`` for worksheet-level ``A1``.

    Returns:
        Source location using the worksheet's name.
    """

    return SourceLocation(
        sheet.template.name,
        "A1" if coordinate is None else coordinate.a1,
    )


def _destinations_by_source(plan: RenderPlan) -> dict[Coordinate, list[Coordinate]]:
    """Group planned destination cells by their source coordinate.

    Args:
        plan: Completed worksheet render plan.

    Returns:
        Source coordinates mapped to every planned destination copy.
    """

    destinations: dict[Coordinate, list[Coordinate]] = defaultdict(list)
    for cell in plan.cells:
        destinations[cell.source_coordinate].append(cell.coordinate)
    return destinations


def _walk(
    nodes: tuple[StructuralNode, ...],
    *,
    loop_depth: int = 0,
) -> list[tuple[StructuralNode, int]]:
    """Walk structural nodes while tracking nested repeat depth.

    Args:
        nodes: Sibling nodes to traverse recursively.
        loop_depth: Number of containing repeat nodes.

    Returns:
        Nodes in preorder paired with their containing repeat depth.
    """

    walked: list[tuple[StructuralNode, int]] = []
    for node in nodes:
        walked.append((node, loop_depth))
        child_depth = loop_depth + 1 if isinstance(node, ForNode) else loop_depth
        walked.extend(_walk(node.children, loop_depth=child_depth))
    return walked


def _is_supported_chart_reference(
    reference: str,
    worksheet_names: frozenset[str],
) -> bool:
    """Return whether a chart formula is one direct in-workbook A1 range.

    Args:
        reference: Formula text stored in the chart XML.
        worksheet_names: Exact worksheet names present in the workbook.

    Returns:
        ``True`` for one concrete cell or rectangle on an existing worksheet.
    """

    match = SHEETRANGE_RE.match(reference)
    if match is None or match.end() != len(reference):
        return False
    sheet_name = match.group("quoted") or match.group("notquoted")
    sheet_name = sheet_name.replace("''", "'")
    if sheet_name not in worksheet_names:
        return False
    try:
        _, boundaries = range_to_tuple(reference)
    except ValueError:
        return False
    min_column, min_row, max_column, max_row = boundaries
    return (
        min_column is not None
        and min_row is not None
        and max_column is not None
        and max_row is not None
        and max_column <= XLSX_MAX_COLUMNS
        and max_row <= XLSX_MAX_ROWS
    )


def plan_sheet_features(
    sheet: SheetSnapshot,
    compiled: CompiledSheet,
    plan: RenderPlan,
    *,
    worksheet_names: frozenset[str],
) -> tuple[SheetFeaturePlan, tuple[Diagnostic, ...]]:
    """Validate and plan workbook features against the completed cell layout.

    Args:
        sheet: Detached worksheet feature and presentation snapshot.
        compiled: Compiled worksheet AST.
        plan: Completed render plan.
        worksheet_names: Exact worksheet names available to direct chart references.

    Returns:
        Adapter feature plan plus diagnostics for unsupported transformations.
    """

    diagnostics: list[Diagnostic] = []
    planned_drawings: list[PlannedDrawing] = []
    has_layout = bool(compiled.children)
    destinations = _destinations_by_source(plan)

    if has_layout and sheet.has_conditional_formatting:
        diagnostics.append(
            Diagnostic(
                DiagnosticCode.CONDITIONAL_FORMATTING_REQUIRES_UNSUPPORTED_TRANSFORM,
                "conditional formatting on a structurally transformed sheet is unsupported",
                _location(sheet),
            )
        )
    if has_layout and sheet.has_data_validations:
        diagnostics.append(
            Diagnostic(
                DiagnosticCode.DATA_VALIDATION_REQUIRES_UNSUPPORTED_TRANSFORM,
                "data validation on a structurally transformed sheet is unsupported",
                _location(sheet),
            )
        )
    if sheet.has_tables:
        diagnostics.append(
            Diagnostic(
                DiagnosticCode.TABLE_REQUIRES_UNSUPPORTED_TRANSFORM,
                "native Excel Tables are not supported by the current workbook writer",
                _location(sheet),
            )
        )
    if sheet.has_unsupported_drawings:
        diagnostics.append(
            Diagnostic(
                DiagnosticCode.DRAWING_REQUIRES_UNSUPPORTED_TRANSFORM,
                "linked images, shapes, and unsupported drawing objects cannot be preserved",
                _location(sheet),
            )
        )

    for drawing in sheet.drawings:
        location = _location(
            sheet,
            drawing.anchor_coordinates[0] if drawing.anchor_coordinates else None,
        )
        if isinstance(drawing, ChartSnapshot):
            anchor_code = DiagnosticCode.CHART_ANCHOR_REQUIRES_UNSUPPORTED_TRANSFORM
            drawing_name = "chart"
            if not drawing.has_supported_type or drawing.is_combined or drawing.is_pivot:
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.CHART_TYPE_UNSUPPORTED,
                        f"{drawing.chart_type} is not in the supported worksheet chart profile",
                        location,
                    )
                )
            if not drawing.references or any(
                not _is_supported_chart_reference(reference, worksheet_names)
                for reference in drawing.references
            ):
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.CHART_REFERENCE_UNSUPPORTED,
                        "chart formulas must be direct concrete A1 ranges in this workbook",
                        location,
                    )
                )
        else:
            assert isinstance(drawing, ImageSnapshot)
            anchor_code = DiagnosticCode.IMAGE_ANCHOR_REQUIRES_UNSUPPORTED_TRANSFORM
            drawing_name = "image"
            if not drawing.has_supported_format:
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.IMAGE_FORMAT_UNSUPPORTED,
                        "embedded images must contain valid PNG or JPEG media",
                        location,
                    )
                )
        if not drawing.has_supported_anchor:
            diagnostics.append(
                Diagnostic(
                    anchor_code,
                    f"{drawing_name} anchor type is not supported",
                    location,
                )
            )
            planned_drawings.append(PlannedDrawing(drawing.anchor_coordinates))
            continue
        anchor_destinations: list[Coordinate] = []
        for coordinate in drawing.anchor_coordinates:
            targets = destinations.get(coordinate, [])
            if len(targets) != 1:
                diagnostics.append(
                    Diagnostic(
                        anchor_code,
                        f"{drawing_name} anchor would be removed or copied by structural rendering",
                        _location(sheet, coordinate),
                    )
                )
                anchor_destinations = list(drawing.anchor_coordinates)
                break
            anchor_destinations.append(targets[0])
        else:
            deltas = {
                (
                    destination.row - source.row,
                    destination.column - source.column,
                )
                for source, destination in zip(
                    drawing.anchor_coordinates,
                    anchor_destinations,
                    strict=True,
                )
            }
            if len(deltas) > 1:
                diagnostics.append(
                    Diagnostic(
                        anchor_code,
                        f"{drawing_name} anchor markers would move by different offsets "
                        "and resize it",
                        location,
                    )
                )
                anchor_destinations = list(drawing.anchor_coordinates)
        planned_drawings.append(PlannedDrawing(tuple(anchor_destinations)))

    coordinate_features = (
        (
            sheet.formula_cells,
            DiagnosticCode.FORMULA_REQUIRES_UNSUPPORTED_TRANSFORM,
            "formula would be copied or moved by structural rendering",
        ),
        (
            sheet.hyperlink_cells,
            DiagnosticCode.HYPERLINK_REQUIRES_UNSUPPORTED_TRANSFORM,
            "hyperlink would be copied or moved by structural rendering",
        ),
        (
            sheet.comment_cells,
            DiagnosticCode.COMMENT_REQUIRES_UNSUPPORTED_TRANSFORM,
            "comment would be copied or moved by structural rendering",
        ),
    )
    for coordinates, code, message in coordinate_features:
        for coordinate in coordinates:
            targets = destinations.get(coordinate, [])
            if targets != [coordinate]:
                diagnostics.append(Diagnostic(code, message, _location(sheet, coordinate)))

    custom_rows = {
        row for row, presentation in sheet.rows.items() if presentation.height is not None
    }
    for node, loop_depth in _walk(compiled.children):
        if not isinstance(node, ForNode) or node.shift != "cells":
            continue
        node_rows = range(node.rectangle.top, node.rectangle.bottom + 1)
        for source_row in custom_rows.intersection(node_rows):
            instance_paths = [
                cell.instance_path
                for cell in plan.cells
                if cell.source_coordinate.row == source_row
                and node.rectangle.left <= cell.source_coordinate.column <= node.rectangle.right
                and len(cell.instance_path) > loop_depth
            ]
            indexes_by_parent: dict[tuple[int, ...], set[int]] = defaultdict(set)
            for path in instance_paths:
                indexes_by_parent[path[:loop_depth]].add(path[loop_depth])
            if any(len(indexes) > 1 for indexes in indexes_by_parent.values()):
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.CELL_SHIFT_WITH_CUSTOM_ROW_HEIGHT,
                        'shift="cells" cannot repeat a worksheet-wide custom row height',
                        _location(sheet, Coordinate(source_row, node.rectangle.left)),
                    )
                )
                break
    return SheetFeaturePlan(tuple(planned_drawings)), tuple(diagnostics)
