"""Workbook-feature checks that depend on both the AST and the render plan."""

from __future__ import annotations

from collections import defaultdict

from excel_template_writer.ast import CompiledSheet, ForNode, RegionNode
from excel_template_writer.diagnostics import Diagnostic, DiagnosticCode, SourceLocation
from excel_template_writer.model import Coordinate
from excel_template_writer.render import RenderPlan
from excel_template_writer.xlsx.model import SheetSnapshot


def _location(sheet: SheetSnapshot, coordinate: Coordinate | None = None) -> SourceLocation:
    return SourceLocation(
        sheet.template.name,
        "A1" if coordinate is None else coordinate.a1,
    )


def _destinations_by_source(plan: RenderPlan) -> dict[Coordinate, list[Coordinate]]:
    destinations: dict[Coordinate, list[Coordinate]] = defaultdict(list)
    for cell in plan.cells:
        destinations[cell.source_coordinate].append(cell.coordinate)
    return destinations


def _walk(
    nodes: tuple[RegionNode, ...],
    *,
    loop_depth: int = 0,
) -> list[tuple[RegionNode, int]]:
    walked: list[tuple[RegionNode, int]] = []
    for node in nodes:
        walked.append((node, loop_depth))
        child_depth = loop_depth + 1 if isinstance(node, ForNode) else loop_depth
        walked.extend(_walk(node.children, loop_depth=child_depth))
    return walked


def validate_sheet_features(
    sheet: SheetSnapshot,
    compiled: CompiledSheet,
    plan: RenderPlan,
) -> tuple[Diagnostic, ...]:
    diagnostics: list[Diagnostic] = []
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
    if sheet.has_drawings:
        diagnostics.append(
            Diagnostic(
                DiagnosticCode.DRAWING_REQUIRES_UNSUPPORTED_TRANSFORM,
                "charts, images, and drawing anchors are not supported by the current writer",
                _location(sheet),
            )
        )

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
    return tuple(diagnostics)
