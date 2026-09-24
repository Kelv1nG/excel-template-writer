"""Pure evaluation and layout planning for a compiled worksheet AST."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, replace
from enum import StrEnum
from typing import Any

from excel_template_writer.ast import (
    CellNode,
    CompiledSheet,
    ExpressionPart,
    ForNode,
    IfNode,
    LiteralPart,
    RegionNode,
    StructuralNode,
)
from excel_template_writer.diagnostics import (
    Diagnostic,
    DiagnosticCode,
    DiagnosticSeverity,
    SourceLocation,
    TemplateRenderError,
)
from excel_template_writer.expressions import (
    ArithmeticTypeError,
    DivisionByZeroError,
    ExpressionEvaluationError,
    FilterTypeError,
    MissingValueError,
    NonFiniteExpressionNumberError,
    evaluate_expression,
)
from excel_template_writer.limits import (
    DEFAULT_RESOURCE_LIMITS,
    XLSX_MAX_CELL_TEXT_LENGTH,
    XLSX_MAX_COLUMNS,
    XLSX_MAX_ROWS,
    ResourceLimits,
)
from excel_template_writer.model import Coordinate, Rectangle
from excel_template_writer.values import (
    TypeAdapter,
    is_collection_value,
    is_ordered_collection,
    normalize_context,
)


class MissingOutputPolicy(StrEnum):
    """Choose how missing values in output tags are rendered."""

    ERROR = "error"
    BLANK = "blank"
    PRESERVE = "preserve"


@dataclass(frozen=True)
class PlannedCell:
    coordinate: Coordinate
    value: Any
    source_coordinate: Coordinate
    instance_path: tuple[int, ...] = ()
    presentation_coordinate: Coordinate | None = None
    writable_destination: bool = False


@dataclass(frozen=True)
class PlannedRow:
    destination_row: int
    source_row: int | None
    instance_path: tuple[int, ...] = ()


@dataclass(frozen=True)
class PlannedMerge:
    rectangle: Rectangle
    source_rectangle: Rectangle
    instance_path: tuple[int, ...] = ()


@dataclass(frozen=True)
class RenderPlan:
    sheet: str
    cells: tuple[PlannedCell, ...]
    rows: tuple[PlannedRow, ...]
    merges: tuple[PlannedMerge, ...]
    height: int
    width: int


@dataclass(frozen=True)
class RenderResult:
    plan: RenderPlan | None
    diagnostics: tuple[Diagnostic, ...]

    def require(self) -> RenderPlan:
        """Return the render plan or raise its diagnostics.

        Returns:
            The successfully completed render plan.

        Raises:
            TemplateRenderError: If rendering produced no plan.
        """

        if self.plan is None:
            raise TemplateRenderError(self.diagnostics)
        return self.plan


@dataclass
class _Block:
    cells: dict[Coordinate, PlannedCell]
    rows: dict[int, PlannedRow]
    merges: list[PlannedMerge]
    reservations: set[Coordinate]
    height: int
    width: int


_EVALUATION_FAILED = object()


class _ResourceLimitExceeded(Exception):
    def __init__(self, diagnostic: Diagnostic) -> None:
        """Create a fail-fast signal for one deterministic resource violation.

        Args:
            diagnostic: Limit diagnostic to return from the public render boundary.
        """

        self.diagnostic = diagnostic
        super().__init__(str(diagnostic))


class _Renderer:
    def __init__(
        self,
        compiled: CompiledSheet,
        limits: ResourceLimits,
        missing_output: MissingOutputPolicy,
    ) -> None:
        """Initialize one pure worksheet render operation.

        Args:
            compiled: Immutable worksheet AST to render.
            limits: Resource ceilings for planning this worksheet.
            missing_output: Policy for missing values in output tags.
        """

        self.compiled = compiled
        self.limits = limits
        self.missing_output = missing_output
        self.diagnostics: list[Diagnostic] = []
        self.repeat_iterations = 0

    def diagnostic(
        self,
        code: DiagnosticCode,
        message: str,
        location: SourceLocation,
        *,
        severity: DiagnosticSeverity = DiagnosticSeverity.ERROR,
    ) -> None:
        """Append a recoverable render diagnostic.

        Args:
            code: Stable diagnostic code.
            message: Human-readable failure description.
            location: Worksheet source location responsible for the failure.
            severity: Whether the diagnostic prevents a render plan.
        """

        self.diagnostics.append(Diagnostic(code, message, location, severity))

    def resource_limit(
        self,
        code: DiagnosticCode,
        message: str,
        location: SourceLocation,
    ) -> None:
        """Stop planning immediately with a resource-limit diagnostic.

        Args:
            code: Stable resource-limit diagnostic code.
            message: Human-readable limit description.
            location: Source location active when the limit was exceeded.

        Raises:
            _ResourceLimitExceeded: Always, carrying the constructed diagnostic.
        """

        raise _ResourceLimitExceeded(Diagnostic(code, message, location))

    def check_block_limits(
        self,
        *,
        cells: int,
        height: int,
        location: SourceLocation,
        width: int,
    ) -> None:
        """Validate measured block size against configured and XLSX ceilings.

        Args:
            cells: Number of planned material cells in the block.
            height: Completed block height in rows.
            location: Source location used for any diagnostic.
            width: Completed block width in columns.

        Raises:
            _ResourceLimitExceeded: If any configured or absolute limit is exceeded.
        """

        if cells > self.limits.max_planned_cells_per_sheet:
            self.resource_limit(
                DiagnosticCode.RENDER_RESOURCE_LIMIT_EXCEEDED,
                "rendered block exceeds "
                f"max_planned_cells_per_sheet={self.limits.max_planned_cells_per_sheet:,}",
                location,
            )
        if height > self.limits.max_output_rows_per_sheet:
            self.resource_limit(
                DiagnosticCode.RENDER_RESOURCE_LIMIT_EXCEEDED,
                "rendered block exceeds "
                f"max_output_rows_per_sheet={self.limits.max_output_rows_per_sheet:,}",
                location,
            )
        if width > self.limits.max_output_columns_per_sheet:
            self.resource_limit(
                DiagnosticCode.RENDER_RESOURCE_LIMIT_EXCEEDED,
                "rendered block exceeds "
                f"max_output_columns_per_sheet={self.limits.max_output_columns_per_sheet:,}",
                location,
            )
        if height > XLSX_MAX_ROWS:
            self.resource_limit(
                DiagnosticCode.XLSX_GRID_LIMIT_EXCEEDED,
                f"rendered block exceeds the XLSX row limit of {XLSX_MAX_ROWS:,}",
                location,
            )
        if width > XLSX_MAX_COLUMNS:
            self.resource_limit(
                DiagnosticCode.XLSX_GRID_LIMIT_EXCEEDED,
                f"rendered block exceeds the XLSX column limit of {XLSX_MAX_COLUMNS:,}",
                location,
            )

    def render_cell(
        self,
        cell: CellNode,
        scope: Mapping[str, Any],
        missing_roots: frozenset[str],
        path: tuple[int, ...],
    ) -> PlannedCell:
        """Evaluate one compiled cell in the current lexical scope.

        Args:
            cell: Compiled source cell.
            scope: Current canonical variable mapping.
            missing_roots: Loop roots intentionally absent for an empty placeholder.
            path: Nested repeat instance indexes for provenance.

        Returns:
            A planned cell. Directive-only cells produce a blank planned value so
            their source presentation remains material after marker removal.
        """

        if not cell.parts:
            return PlannedCell(
                cell.coordinate,
                None,
                cell.coordinate,
                path,
                cell.coordinate,
                self.compiled.template.cells[cell.coordinate] is None,
            )
        values: list[Any] = []
        for part in cell.parts:
            if isinstance(part, LiteralPart):
                values.append(part.value)
                continue
            if not isinstance(part, ExpressionPart):
                raise TypeError(f"unsupported cell part: {type(part).__name__}")
            try:
                value = evaluate_expression(part.expression, scope)
            except MissingValueError as error:
                if error.root in missing_roots:
                    value = None
                elif self.missing_output is MissingOutputPolicy.BLANK:
                    self.diagnostic(
                        DiagnosticCode.MISSING_VALUE_RENDERED,
                        str(error),
                        part.span.location,
                        severity=DiagnosticSeverity.WARNING,
                    )
                    value = None
                elif self.missing_output is MissingOutputPolicy.PRESERVE:
                    self.diagnostic(
                        DiagnosticCode.MISSING_VALUE_RENDERED,
                        str(error),
                        part.span.location,
                        severity=DiagnosticSeverity.WARNING,
                    )
                    source_value = self.compiled.template.cells[cell.coordinate]
                    if not isinstance(source_value, str):
                        raise TypeError(
                            "output expression source cell must contain text"
                        ) from error
                    value = source_value[part.span.start : part.span.end]
                else:
                    self.diagnostic(
                        DiagnosticCode.MISSING_VALUE,
                        str(error),
                        part.span.location,
                    )
                    value = None
            except FilterTypeError as error:
                self.diagnostic(
                    DiagnosticCode.FILTER_TYPE_MISMATCH,
                    str(error),
                    part.span.location,
                )
                value = None
            except ArithmeticTypeError as error:
                self.diagnostic(
                    DiagnosticCode.ARITHMETIC_TYPE_MISMATCH,
                    str(error),
                    part.span.location,
                )
                value = None
            except DivisionByZeroError as error:
                self.diagnostic(
                    DiagnosticCode.DIVISION_BY_ZERO,
                    str(error),
                    part.span.location,
                )
                value = None
            except NonFiniteExpressionNumberError as error:
                self.diagnostic(
                    DiagnosticCode.NON_FINITE_EXPRESSION_NUMBER,
                    str(error),
                    part.span.location,
                )
                value = None
            except ExpressionEvaluationError as error:
                self.diagnostic(
                    DiagnosticCode.MISSING_VALUE,
                    str(error),
                    part.span.location,
                )
                value = None
            if is_collection_value(value):
                self.diagnostic(
                    DiagnosticCode.COLLECTION_IN_SCALAR_CELL,
                    "collections must be rendered by a for block or an explicit filter",
                    part.span.location,
                )
                value = None
            values.append(value)
        if len(cell.parts) == 1:
            value = values[0]
        else:
            value = "".join("" if item is None else str(item) for item in values)
        if isinstance(value, str) and len(value) > XLSX_MAX_CELL_TEXT_LENGTH:
            self.resource_limit(
                DiagnosticCode.CELL_TEXT_LIMIT_EXCEEDED,
                f"cell text exceeds the XLSX limit of {XLSX_MAX_CELL_TEXT_LENGTH:,} characters",
                SourceLocation(self.compiled.template.name, cell.coordinate.a1),
            )
        return PlannedCell(
            cell.coordinate,
            value,
            cell.coordinate,
            path,
            cell.coordinate,
            self.compiled.template.cells[cell.coordinate] is None,
        )

    def evaluate_region_expression(
        self,
        node: ForNode | IfNode,
        scope: Mapping[str, Any],
    ) -> Any:
        """Evaluate the controlling expression of a repeat or condition.

        Args:
            node: Repeat or conditional AST node.
            scope: Current canonical lexical scope.

        Returns:
            The expression value, or an internal failure sentinel after diagnostics.
        """

        expression = node.iterable if isinstance(node, ForNode) else node.condition
        try:
            return evaluate_expression(expression, scope)
        except FilterTypeError as error:
            self.diagnostic(
                DiagnosticCode.FILTER_TYPE_MISMATCH,
                str(error),
                node.span.location,
            )
            return _EVALUATION_FAILED
        except ArithmeticTypeError as error:
            self.diagnostic(
                DiagnosticCode.ARITHMETIC_TYPE_MISMATCH,
                str(error),
                node.span.location,
            )
            return _EVALUATION_FAILED
        except DivisionByZeroError as error:
            self.diagnostic(
                DiagnosticCode.DIVISION_BY_ZERO,
                str(error),
                node.span.location,
            )
            return _EVALUATION_FAILED
        except NonFiniteExpressionNumberError as error:
            self.diagnostic(
                DiagnosticCode.NON_FINITE_EXPRESSION_NUMBER,
                str(error),
                node.span.location,
            )
            return _EVALUATION_FAILED
        except ExpressionEvaluationError as error:
            self.diagnostic(DiagnosticCode.MISSING_VALUE, str(error), node.span.location)
            return _EVALUATION_FAILED

    def shift_grid(
        self,
        grid: dict[Coordinate, PlannedCell],
        *,
        bottom: int,
        left: int,
        right: int,
        delta: int,
        shift: str,
        replacement_height: int,
        top: int,
        location: SourceLocation,
    ) -> dict[Coordinate, PlannedCell]:
        """Move planned cells below a replaced child allocation.

        Args:
            grid: Current local destination grid.
            bottom: Original local bottom row of the child.
            left: Local left edge of the child lane.
            right: Local right edge of the child lane.
            delta: Signed change in child height.
            shift: ``"rows"`` for global movement, ``"cells"`` for lane movement,
                or ``"none"`` for a fixed grid.
            replacement_height: Completed child height used to remove contracted rows.
            top: Original local top row of the child.
            location: Source location used for collision diagnostics.

        Returns:
            A new destination grid with affected cells translated.
        """

        if shift == "none":
            return dict(grid)

        shifted: dict[Coordinate, PlannedCell] = {}
        eliminated_start = top + replacement_height
        for coordinate, cell in grid.items():
            in_lane = shift == "rows" or (shift == "cells" and left <= coordinate.column <= right)
            if delta < 0 and in_lane and eliminated_start <= coordinate.row <= bottom:
                continue
            new_coordinate = coordinate
            if in_lane and coordinate.row > bottom:
                new_coordinate = Coordinate(coordinate.row + delta, coordinate.column)
            if new_coordinate in shifted:
                self.diagnostic(
                    DiagnosticCode.LAYOUT_COLLISION,
                    f"two source cells allocate destination {new_coordinate.a1}",
                    location,
                )
                continue
            shifted[new_coordinate] = replace(cell, coordinate=new_coordinate)
        return shifted

    def shift_reservations(
        self,
        reservations: set[Coordinate],
        *,
        bottom: int,
        left: int,
        right: int,
        delta: int,
        shift: str,
        replacement_height: int,
        top: int,
    ) -> set[Coordinate]:
        """Move fixed-footprint reservations with their containing layout.

        Args:
            reservations: Current local reserved coordinates.
            bottom: Original local bottom row of the child.
            left: Local left edge of the child lane.
            right: Local right edge of the child lane.
            delta: Signed change in child height.
            shift: Child shift policy.
            replacement_height: Completed child height used to remove contracted rows.
            top: Original local top row of the child.

        Returns:
            Reserved coordinates after applying the same displacement as planned cells.
        """

        if shift == "none":
            return set(reservations)

        shifted: set[Coordinate] = set()
        eliminated_start = top + replacement_height
        for coordinate in reservations:
            in_lane = shift == "rows" or (shift == "cells" and left <= coordinate.column <= right)
            if delta < 0 and in_lane and eliminated_start <= coordinate.row <= bottom:
                continue
            if in_lane and coordinate.row > bottom:
                coordinate = Coordinate(coordinate.row + delta, coordinate.column)
            shifted.add(coordinate)
        return shifted

    def shift_rows(
        self,
        rows: dict[int, PlannedRow],
        *,
        top: int,
        bottom: int,
        delta: int,
        shift: str,
    ) -> dict[int, PlannedRow]:
        """Translate worksheet-wide row presentation for a row-shift child.

        Args:
            rows: Current local row provenance mapping.
            top: Original local top row of the child.
            bottom: Original local bottom row of the child.
            delta: Signed change in child height.
            shift: Child shift policy.

        Returns:
            Updated row provenance; unchanged for cell-shift children.
        """

        if shift != "rows":
            return rows
        shifted: dict[int, PlannedRow] = {}
        for destination_row, row in rows.items():
            if top <= destination_row <= bottom:
                continue
            new_row = destination_row + delta if destination_row > bottom else destination_row
            shifted[new_row] = replace(row, destination_row=new_row)
        return shifted

    def shift_merges(
        self,
        merges: list[PlannedMerge],
        *,
        bottom: int,
        left: int,
        right: int,
        delta: int,
        shift: str,
    ) -> list[PlannedMerge]:
        """Translate merges wholly affected by a child height change.

        Args:
            merges: Current local merged-range plans.
            bottom: Original local bottom row of the child.
            left: Local left edge of the child lane.
            right: Local right edge of the child lane.
            delta: Signed change in child height.
            shift: Child shift policy.

        Returns:
            Merged-range plans with affected rectangles translated.
        """

        if shift == "none":
            return list(merges)

        shifted: list[PlannedMerge] = []
        for merge in merges:
            in_lane = shift == "rows" or (
                left <= merge.rectangle.left and merge.rectangle.right <= right
            )
            rectangle = merge.rectangle
            if in_lane and rectangle.top > bottom:
                rectangle = rectangle.translated(rows=delta)
            shifted.append(replace(merge, rectangle=rectangle))
        return shifted

    def validate_child_allocations(
        self,
        reservations: set[Coordinate],
        grid: dict[Coordinate, PlannedCell],
        child: _Block,
        *,
        top: int,
        left: int,
        parent_rectangle: Rectangle,
    ) -> None:
        """Reject cells or fixed footprints that claim an existing reservation.

        Args:
            reservations: Fixed coordinates already owned in the parent block.
            grid: Current parent destination grid.
            child: Completed child block in local coordinates.
            top: Parent-local destination top row.
            left: Parent-local destination left column.
            parent_rectangle: Source rectangle owning the local destination grid.
        """

        child_cells = {
            Coordinate(top + coordinate.row - 1, left + coordinate.column - 1)
            for coordinate in child.cells
        }
        child_reservations = {
            Coordinate(top + coordinate.row - 1, left + coordinate.column - 1)
            for coordinate in child.reservations
        }
        collisions = (child_cells | child_reservations).intersection(reservations)
        for destination in child_reservations:
            existing = grid.get(destination)
            if existing is not None and not existing.writable_destination:
                collisions.add(destination)
        if not collisions:
            return
        destination = min(collisions)
        absolute_destination = Coordinate(
            parent_rectangle.top + destination.row - 1,
            parent_rectangle.left + destination.column - 1,
        )
        self.diagnostic(
            DiagnosticCode.LAYOUT_COLLISION,
            f"fixed footprint is occupied at {absolute_destination.a1}",
            SourceLocation(self.compiled.template.name, absolute_destination.a1),
        )

    def add_child_reservations(
        self,
        reservations: set[Coordinate],
        child: _Block,
        *,
        top: int,
        left: int,
    ) -> None:
        """Place a child's fixed-footprint reservations in its parent block.

        Args:
            reservations: Mutable parent reservation set.
            child: Completed child block in local coordinates.
            top: Parent-local destination top row.
            left: Parent-local destination left column.
        """

        reservations.update(
            Coordinate(top + coordinate.row - 1, left + coordinate.column - 1)
            for coordinate in child.reservations
        )

    def add_child_cells(
        self,
        grid: dict[Coordinate, PlannedCell],
        child: _Block,
        *,
        bottom: int,
        top: int,
        left: int,
        right: int,
        shift: str,
        location: SourceLocation,
    ) -> None:
        """Place a completed child grid into its parent allocation.

        Args:
            grid: Mutable parent destination grid.
            child: Completed child block in local coordinates.
            bottom: Original parent-local bottom row of the child.
            top: Parent-local destination top row.
            left: Parent-local destination left column.
            right: Original parent-local right column of the child.
            shift: Child shift policy.
            location: Source location used for collision diagnostics.
        """

        source_area = Rectangle(top, left, bottom, right)
        for coordinate, cell in child.cells.items():
            destination = Coordinate(top + coordinate.row - 1, left + coordinate.column - 1)
            existing = grid.get(destination)
            if existing is not None and not (shift == "none" and existing.writable_destination):
                if coordinate in child.reservations:
                    continue
                self.diagnostic(
                    DiagnosticCode.LAYOUT_COLLISION,
                    f"two source cells allocate destination {destination.a1}",
                    location,
                )
                continue
            presentation_coordinate = cell.presentation_coordinate
            if shift == "none" and not source_area.contains_coordinate(destination):
                presentation_coordinate = (
                    existing.presentation_coordinate if existing is not None else None
                )
            grid[destination] = replace(
                cell,
                coordinate=destination,
                presentation_coordinate=presentation_coordinate,
                writable_destination=False,
            )

    def add_child_rows(
        self,
        rows: dict[int, PlannedRow],
        child: _Block,
        *,
        parent_source_top: int,
        top: int,
        shift: str,
    ) -> None:
        """Place child row provenance when it owns complete worksheet rows.

        Args:
            rows: Mutable parent row-provenance mapping.
            child: Completed child block.
            parent_source_top: Absolute source row corresponding to parent-local row one.
            top: Parent-local destination top row.
            shift: Child shift policy.
        """

        if shift not in {"rows", "none"}:
            return
        for destination_row, row in child.rows.items():
            absolute_row = top + destination_row - 1
            source_row = row.source_row if shift == "rows" else parent_source_top + absolute_row - 1
            rows[absolute_row] = replace(
                row,
                destination_row=absolute_row,
                source_row=source_row,
            )

    def add_child_merges(
        self,
        merges: list[PlannedMerge],
        child: _Block,
        *,
        bottom: int,
        top: int,
        left: int,
        right: int,
        shift: str,
    ) -> None:
        """Place completed child merges and report any overlap.

        Args:
            merges: Mutable parent merged-range plans.
            child: Completed child block.
            bottom: Original parent-local bottom row of the child.
            top: Parent-local destination top row.
            left: Parent-local destination left column.
            right: Original parent-local right column of the child.
            shift: Child shift policy.
        """

        source_area = Rectangle(top, left, bottom, right)
        for merge in child.merges:
            rectangle = merge.rectangle.translated(rows=top - 1, columns=left - 1)
            if shift == "none":
                exact = next(
                    (existing for existing in merges if existing.rectangle == rectangle),
                    None,
                )
                if exact is not None:
                    merges.remove(exact)
                merges.append(
                    replace(
                        merge,
                        rectangle=rectangle,
                        source_rectangle=(
                            merge.source_rectangle if source_area.contains(rectangle) else rectangle
                        ),
                    )
                )
                continue
            if any(rectangle.intersects(existing.rectangle) for existing in merges):
                self.diagnostic(
                    DiagnosticCode.LAYOUT_COLLISION,
                    "rendered merged ranges overlap",
                    SourceLocation(
                        self.compiled.template.name,
                        Coordinate(rectangle.top, rectangle.left).a1,
                    ),
                )
                continue
            merges.append(replace(merge, rectangle=rectangle))

    def render_area(
        self,
        rectangle: Rectangle,
        children: tuple[StructuralNode, ...],
        scope: Mapping[str, Any],
        missing_roots: frozenset[str],
        path: tuple[int, ...],
    ) -> _Block:
        """Measure and render one source rectangle in local coordinates.

        Args:
            rectangle: Exact source rectangle being rendered.
            children: Direct structural children owned by the rectangle.
            scope: Current canonical lexical scope.
            missing_roots: Loop roots intentionally absent for empty placeholders.
            path: Nested repeat instance indexes for provenance.

        Returns:
            A completed local block containing cells, rows, merges, and measured size.
        """

        self.check_block_limits(
            cells=0,
            height=rectangle.height,
            width=rectangle.width,
            location=SourceLocation(
                self.compiled.template.name,
                Coordinate(rectangle.top, rectangle.left).a1,
            ),
        )
        grid: dict[Coordinate, PlannedCell] = {}
        reservations: set[Coordinate] = set()
        rows = {
            local_row: PlannedRow(
                local_row,
                rectangle.top + local_row - 1,
                path,
            )
            for local_row in range(1, rectangle.height + 1)
        }
        merges = [
            PlannedMerge(
                merged.translated(rows=1 - rectangle.top, columns=1 - rectangle.left),
                merged,
                path,
            )
            for merged in self.compiled.template.merged_ranges
            if rectangle.contains(merged)
            and not any(child.rectangle.contains(merged) for child in children)
        ]
        for source_coordinate, cell in self.compiled.cells.items():
            if not rectangle.contains_coordinate(source_coordinate):
                continue
            if any(child.rectangle.contains_coordinate(source_coordinate) for child in children):
                continue
            planned = self.render_cell(cell, scope, missing_roots, path)
            local = Coordinate(
                source_coordinate.row - rectangle.top + 1,
                source_coordinate.column - rectangle.left + 1,
            )
            grid[local] = replace(planned, coordinate=local)

        height = rectangle.height
        for child_node in sorted(
            children,
            key=lambda node: (node.rectangle.top, node.rectangle.left),
            reverse=True,
        ):
            child = self.render_region(child_node, scope, missing_roots, path)
            child_top = child_node.rectangle.top - rectangle.top + 1
            child_left = child_node.rectangle.left - rectangle.left + 1
            child_bottom = child_node.rectangle.bottom - rectangle.top + 1
            child_right = child_node.rectangle.right - rectangle.left + 1
            delta = child.height - child_node.rectangle.height
            shift = child_node.shift
            grid = self.shift_grid(
                grid,
                bottom=child_bottom,
                left=child_left,
                right=child_right,
                delta=delta,
                shift=shift,
                replacement_height=child.height,
                top=child_top,
                location=child_node.span.location,
            )
            reservations = self.shift_reservations(
                reservations,
                bottom=child_bottom,
                left=child_left,
                right=child_right,
                delta=delta,
                shift=shift,
                replacement_height=child.height,
                top=child_top,
            )
            rows = self.shift_rows(
                rows,
                top=child_top,
                bottom=child_bottom,
                delta=delta,
                shift=shift,
            )
            merges = self.shift_merges(
                merges,
                bottom=child_bottom,
                left=child_left,
                right=child_right,
                delta=delta,
                shift=shift,
            )
            self.validate_child_allocations(
                reservations,
                grid,
                child,
                top=child_top,
                left=child_left,
                parent_rectangle=rectangle,
            )
            self.add_child_cells(
                grid,
                child,
                bottom=child_bottom,
                top=child_top,
                left=child_left,
                right=child_right,
                shift=shift,
                location=child_node.span.location,
            )
            self.add_child_reservations(
                reservations,
                child,
                top=child_top,
                left=child_left,
            )
            self.add_child_rows(
                rows,
                child,
                parent_source_top=rectangle.top,
                top=child_top,
                shift=shift,
            )
            self.add_child_merges(
                merges,
                child,
                bottom=child_bottom,
                top=child_top,
                left=child_left,
                right=child_right,
                shift=shift,
            )
            if shift == "rows":
                height += delta
            else:
                height = max(height, child_top + child.height - 1)
        height = max(height, max((coordinate.row for coordinate in grid), default=0))
        height = max(height, max((coordinate.row for coordinate in reservations), default=0))
        height = max(height, max(rows, default=0))
        allocated_cells = len(reservations) + sum(
            coordinate not in reservations for coordinate in grid
        )
        self.check_block_limits(
            cells=allocated_cells,
            height=height,
            width=rectangle.width,
            location=SourceLocation(
                self.compiled.template.name,
                Coordinate(rectangle.top, rectangle.left).a1,
            ),
        )
        for destination_row in range(1, height + 1):
            rows.setdefault(destination_row, PlannedRow(destination_row, None, path))
        return _Block(grid, rows, merges, reservations, max(0, height), rectangle.width)

    def validate_no_shift_merges(
        self,
        node: ForNode,
        *,
        iterations: int,
    ) -> None:
        """Validate fixed destinations against the authored merge topology.

        Args:
            node: Fixed-stride repeat being rendered.
            iterations: Number of physical instances, including an empty placeholder.
        """

        footprint = Rectangle(
            node.rectangle.top,
            node.rectangle.left,
            node.rectangle.top + iterations * node.rectangle.height - 1,
            node.rectangle.right,
        )
        source_merges = {
            merge
            for merge in self.compiled.template.merged_ranges
            if node.rectangle.contains(merge)
        }
        expected_merges = {
            merge.translated(rows=index * node.rectangle.height)
            for index in range(iterations)
            for merge in source_merges
        }
        intersecting_merges = {
            merge for merge in self.compiled.template.merged_ranges if merge.intersects(footprint)
        }
        incompatible_merges = expected_merges.symmetric_difference(intersecting_merges)
        incompatible_coordinates = sorted(
            {Coordinate(merge.top, merge.left) for merge in incompatible_merges}
        )
        for coordinate in incompatible_coordinates:
            self.diagnostic(
                DiagnosticCode.MERGE_CROSSES_BLOCK_BOUNDARY,
                'shift="none" destination merge topology does not match the source body',
                SourceLocation(self.compiled.template.name, coordinate.a1),
            )

    def render_region(
        self,
        node: StructuralNode,
        scope: Mapping[str, Any],
        missing_roots: frozenset[str],
        path: tuple[int, ...],
    ) -> _Block:
        """Render one structural AST node into a completed local block.

        Args:
            node: Explicit region, repeat, or conditional node.
            scope: Current canonical lexical scope.
            missing_roots: Loop roots intentionally absent for empty placeholders.
            path: Nested repeat instance indexes for provenance.

        Returns:
            The measured and evaluated child block.

        Raises:
            TypeError: If an unsupported structural node reaches the renderer.
        """

        if isinstance(node, RegionNode):
            return self.render_area(node.rectangle, node.children, scope, missing_roots, path)

        if isinstance(node, ForNode):
            raw_items = self.evaluate_region_expression(node, scope)
            if raw_items is _EVALUATION_FAILED:
                items: list[Any] = []
            elif not is_ordered_collection(raw_items):
                self.diagnostic(
                    DiagnosticCode.EXPECTED_COLLECTION,
                    "for expression must evaluate to an ordered list or tuple",
                    node.span.location,
                )
                items = []
            else:
                items = list(raw_items)
            iterations = max(1, len(items))
            self.repeat_iterations += iterations
            if self.repeat_iterations > self.limits.max_repeat_iterations_per_sheet:
                self.resource_limit(
                    DiagnosticCode.RENDER_RESOURCE_LIMIT_EXCEEDED,
                    "worksheet exceeds max_repeat_iterations_per_sheet="
                    f"{self.limits.max_repeat_iterations_per_sheet:,}",
                    node.span.location,
                )
            blocks: list[_Block] = []
            rendered_allocations = 0
            rendered_height = 0
            if items:
                for index, item in enumerate(items):
                    child_scope = dict(scope)
                    child_scope[node.variable] = item
                    block = self.render_area(
                        node.rectangle,
                        node.children,
                        child_scope,
                        missing_roots,
                        (*path, index),
                    )
                    blocks.append(block)
                    rendered_allocations += len(block.reservations) + sum(
                        coordinate not in block.reservations for coordinate in block.cells
                    )
                    rendered_height += block.height
                    self.check_block_limits(
                        cells=rendered_allocations,
                        height=rendered_height,
                        width=node.rectangle.width,
                        location=node.span.location,
                    )
            else:
                block = self.render_area(
                    node.rectangle,
                    node.children,
                    scope,
                    missing_roots | {node.variable},
                    (*path, -1),
                )
                blocks.append(block)
            grid: dict[Coordinate, PlannedCell] = {}
            rows: dict[int, PlannedRow] = {}
            merges: list[PlannedMerge] = []
            reservations: set[Coordinate] = set()
            row_offset = 0
            for block in blocks:
                for coordinate, cell in block.cells.items():
                    destination = Coordinate(coordinate.row + row_offset, coordinate.column)
                    grid[destination] = replace(cell, coordinate=destination)
                for destination_row, row in block.rows.items():
                    absolute_row = destination_row + row_offset
                    rows[absolute_row] = replace(row, destination_row=absolute_row)
                for merge in block.merges:
                    merges.append(
                        replace(
                            merge,
                            rectangle=merge.rectangle.translated(rows=row_offset),
                        )
                    )
                reservations.update(
                    Coordinate(coordinate.row + row_offset, coordinate.column)
                    for coordinate in block.reservations
                )
                row_offset += block.height
            if node.shift == "none":
                self.validate_no_shift_merges(node, iterations=len(blocks))
                reservation_cells = row_offset * node.rectangle.width
                self.check_block_limits(
                    cells=reservation_cells,
                    height=row_offset,
                    width=node.rectangle.width,
                    location=node.span.location,
                )
                reservations.update(
                    Coordinate(row, column)
                    for row in range(1, row_offset + 1)
                    for column in range(1, node.rectangle.width + 1)
                )
            return _Block(
                grid,
                rows,
                merges,
                reservations,
                row_offset,
                node.rectangle.width,
            )

        if isinstance(node, IfNode):
            raw_condition = self.evaluate_region_expression(node, scope)
            selected = raw_condition is not _EVALUATION_FAILED and bool(raw_condition)
            branch = node.true_rectangle if selected else node.false_rectangle
            if branch is None:
                return _Block({}, {}, [], set(), 0, node.rectangle.width)
            branch_children = tuple(
                child for child in node.children if branch.contains(child.rectangle)
            )
            return self.render_area(branch, branch_children, scope, missing_roots, path)
        raise TypeError(f"unsupported region node: {type(node).__name__}")


def render_sheet(
    compiled: CompiledSheet,
    context: object,
    *,
    missing_output: MissingOutputPolicy | str = MissingOutputPolicy.ERROR,
    adapters: Iterable[TypeAdapter[Any]] = (),
    limits: ResourceLimits = DEFAULT_RESOURCE_LIMITS,
    preserved_source_rows: Iterable[int] = (),
) -> RenderResult:
    """Evaluate a compiled sheet into an adapter-neutral render plan.

    Args:
        compiled: Immutable compiled worksheet AST.
        context: Raw or already-normalized render context.
        missing_output: Policy for missing values in output tags.
        adapters: Explicit runtime-type adapters used during normalization.
        limits: Resource ceilings for normalization and planning.
        preserved_source_rows: Explicit row coordinates whose layout provenance must be planned
            even when they contain no material cells.

    Returns:
        A complete render plan or structured diagnostics; never a partial plan.

    Raises:
        ValueError: If ``missing_output`` is unsupported or a preserved row is not a positive
            integer.
    """

    try:
        missing_output_policy = MissingOutputPolicy(missing_output)
    except (TypeError, ValueError) as error:
        raise ValueError("missing_output must be 'error', 'blank', or 'preserve'") from error
    normalization = normalize_context(context, adapters=adapters, limits=limits)
    if normalization.context is None:
        return RenderResult(None, normalization.diagnostics)
    preserved_rows = tuple(preserved_source_rows)
    if any(type(row) is not int or row < 1 for row in preserved_rows):
        raise ValueError("preserved_source_rows must contain one-based row coordinates")
    root_rectangle = Rectangle(
        compiled.rectangle.top,
        compiled.rectangle.left,
        max((compiled.rectangle.bottom, *preserved_rows)),
        compiled.rectangle.right,
    )
    renderer = _Renderer(compiled, limits, missing_output_policy)
    try:
        block = renderer.render_area(
            root_rectangle,
            compiled.children,
            normalization.context,
            frozenset(),
            (),
        )
    except _ResourceLimitExceeded as error:
        return RenderResult(None, (error.diagnostic,))
    diagnostics = tuple(renderer.diagnostics)
    if any(diagnostic.severity is DiagnosticSeverity.ERROR for diagnostic in diagnostics):
        return RenderResult(None, diagnostics)
    cells = tuple(cell for _, cell in sorted(block.cells.items()))
    rows = tuple(row for _, row in sorted(block.rows.items()))
    merges = tuple(
        sorted(
            block.merges,
            key=lambda merge: (
                merge.rectangle.top,
                merge.rectangle.left,
                merge.rectangle.bottom,
                merge.rectangle.right,
            ),
        )
    )
    return RenderResult(
        RenderPlan(
            compiled.template.name,
            cells,
            rows,
            merges,
            block.height,
            block.width,
        ),
        diagnostics,
    )
