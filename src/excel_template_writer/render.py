"""Pure evaluation and layout planning for a compiled worksheet AST."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, replace
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
    SourceLocation,
    TemplateRenderError,
)
from excel_template_writer.expressions import (
    ExpressionEvaluationError,
    MissingValueError,
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


@dataclass(frozen=True)
class PlannedCell:
    coordinate: Coordinate
    value: Any
    source_coordinate: Coordinate
    instance_path: tuple[int, ...] = ()


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
    def __init__(self, compiled: CompiledSheet, limits: ResourceLimits) -> None:
        """Initialize one pure worksheet render operation.

        Args:
            compiled: Immutable worksheet AST to render.
            limits: Resource ceilings for planning this worksheet.
        """

        self.compiled = compiled
        self.limits = limits
        self.diagnostics: list[Diagnostic] = []
        self.repeat_iterations = 0

    def diagnostic(
        self,
        code: DiagnosticCode,
        message: str,
        location: SourceLocation,
    ) -> None:
        """Append a recoverable render diagnostic.

        Args:
            code: Stable diagnostic code.
            message: Human-readable failure description.
            location: Worksheet source location responsible for the failure.
        """

        self.diagnostics.append(Diagnostic(code, message, location))

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
    ) -> PlannedCell | None:
        """Evaluate one compiled cell in the current lexical scope.

        Args:
            cell: Compiled source cell.
            scope: Current canonical variable mapping.
            missing_roots: Loop roots intentionally absent for an empty placeholder.
            path: Nested repeat instance indexes for provenance.

        Returns:
            A planned cell, or ``None`` when the source cell has no output parts.
        """

        if not cell.parts:
            return None
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
                else:
                    self.diagnostic(
                        DiagnosticCode.MISSING_VALUE,
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
        return PlannedCell(cell.coordinate, value, cell.coordinate, path)

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
            shift: ``"rows"`` for global movement or ``"cells"`` for lane movement.
            replacement_height: Completed child height used to remove contracted rows.
            top: Original local top row of the child.
            location: Source location used for collision diagnostics.

        Returns:
            A new destination grid with affected cells translated.
        """

        shifted: dict[Coordinate, PlannedCell] = {}
        eliminated_start = top + replacement_height
        for coordinate, cell in grid.items():
            in_lane = shift == "rows" or left <= coordinate.column <= right
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

    def add_child_cells(
        self,
        grid: dict[Coordinate, PlannedCell],
        child: _Block,
        *,
        top: int,
        left: int,
        location: SourceLocation,
    ) -> None:
        """Place a completed child grid into its parent allocation.

        Args:
            grid: Mutable parent destination grid.
            child: Completed child block in local coordinates.
            top: Parent-local destination top row.
            left: Parent-local destination left column.
            location: Source location used for collision diagnostics.
        """

        for coordinate, cell in child.cells.items():
            destination = Coordinate(top + coordinate.row - 1, left + coordinate.column - 1)
            if destination in grid:
                self.diagnostic(
                    DiagnosticCode.LAYOUT_COLLISION,
                    f"two source cells allocate destination {destination.a1}",
                    location,
                )
                continue
            grid[destination] = replace(cell, coordinate=destination)

    def add_child_rows(
        self,
        rows: dict[int, PlannedRow],
        child: _Block,
        *,
        top: int,
        shift: str,
    ) -> None:
        """Place child row provenance when it owns complete worksheet rows.

        Args:
            rows: Mutable parent row-provenance mapping.
            child: Completed child block.
            top: Parent-local destination top row.
            shift: Child shift policy.
        """

        if shift != "rows":
            return
        for destination_row, row in child.rows.items():
            absolute_row = top + destination_row - 1
            rows[absolute_row] = replace(row, destination_row=absolute_row)

    def add_child_merges(
        self,
        merges: list[PlannedMerge],
        child: _Block,
        *,
        top: int,
        left: int,
    ) -> None:
        """Place completed child merges and report any overlap.

        Args:
            merges: Mutable parent merged-range plans.
            child: Completed child block.
            top: Parent-local destination top row.
            left: Parent-local destination left column.
        """

        for merge in child.merges:
            rectangle = merge.rectangle.translated(rows=top - 1, columns=left - 1)
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
            if planned is None:
                continue
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
            self.add_child_cells(
                grid,
                child,
                top=child_top,
                left=child_left,
                location=child_node.span.location,
            )
            self.add_child_rows(rows, child, top=child_top, shift=shift)
            self.add_child_merges(
                merges,
                child,
                top=child_top,
                left=child_left,
            )
            if shift == "rows":
                height += delta
            else:
                height = max(height, child_top + child.height - 1)
        height = max(height, max((coordinate.row for coordinate in grid), default=0))
        height = max(height, max(rows, default=0))
        self.check_block_limits(
            cells=len(grid),
            height=height,
            width=rectangle.width,
            location=SourceLocation(
                self.compiled.template.name,
                Coordinate(rectangle.top, rectangle.left).a1,
            ),
        )
        for destination_row in range(1, height + 1):
            rows.setdefault(destination_row, PlannedRow(destination_row, None, path))
        return _Block(grid, rows, merges, max(0, height), rectangle.width)

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
            rendered_cells = 0
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
                    rendered_cells += len(block.cells)
                    rendered_height += block.height
                    self.check_block_limits(
                        cells=rendered_cells,
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
                row_offset += block.height
            return _Block(grid, rows, merges, row_offset, node.rectangle.width)

        if isinstance(node, IfNode):
            raw_condition = self.evaluate_region_expression(node, scope)
            selected = raw_condition is not _EVALUATION_FAILED and bool(raw_condition)
            branch = node.true_rectangle if selected else node.false_rectangle
            if branch is None:
                return _Block({}, {}, [], 0, node.rectangle.width)
            branch_children = tuple(
                child for child in node.children if branch.contains(child.rectangle)
            )
            return self.render_area(branch, branch_children, scope, missing_roots, path)
        raise TypeError(f"unsupported region node: {type(node).__name__}")


def render_sheet(
    compiled: CompiledSheet,
    context: object,
    *,
    adapters: Iterable[TypeAdapter[Any]] = (),
    limits: ResourceLimits = DEFAULT_RESOURCE_LIMITS,
) -> RenderResult:
    """Evaluate a compiled sheet into an adapter-neutral render plan.

    Args:
        compiled: Immutable compiled worksheet AST.
        context: Raw or already-normalized render context.
        adapters: Explicit runtime-type adapters used during normalization.
        limits: Resource ceilings for normalization and planning.

    Returns:
        A complete render plan or structured diagnostics; never a partial plan.
    """

    normalization = normalize_context(context, adapters=adapters, limits=limits)
    if normalization.context is None:
        return RenderResult(None, normalization.diagnostics)
    renderer = _Renderer(compiled, limits)
    try:
        block = renderer.render_area(
            compiled.rectangle,
            compiled.children,
            normalization.context,
            frozenset(),
            (),
        )
    except _ResourceLimitExceeded as error:
        return RenderResult(None, (error.diagnostic,))
    if renderer.diagnostics:
        return RenderResult(None, tuple(renderer.diagnostics))
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
        (),
    )
