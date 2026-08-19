"""Compilation pipeline from worksheet cells to a spatial AST."""

from __future__ import annotations

from dataclasses import dataclass
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
    TemplateCompilationError,
)
from excel_template_writer.directives import (
    Directive,
    DirectiveSyntaxError,
    ElseDirective,
    EndForDirective,
    EndIfDirective,
    EndRegionDirective,
    ForDirective,
    IfDirective,
    RegionDirective,
    parse_directive,
)
from excel_template_writer.expressions import ExpressionSyntaxError, parse_expression
from excel_template_writer.model import Coordinate, Rectangle, WorksheetTemplate
from excel_template_writer.syntax import (
    DirectiveToken,
    OutputToken,
    SourceSpan,
    TextToken,
    lex_cell,
)


@dataclass(frozen=True)
class CompilationResult:
    compiled: CompiledSheet | None
    diagnostics: tuple[Diagnostic, ...]

    def require(self) -> CompiledSheet:
        """Return the compiled sheet or raise its diagnostics.

        Returns:
            The successfully compiled immutable sheet.

        Raises:
            TemplateCompilationError: If compilation produced no sheet.
        """

        if self.compiled is None:
            raise TemplateCompilationError(self.diagnostics)
        return self.compiled


@dataclass(frozen=True)
class _Marker:
    coordinate: Coordinate
    directive: Directive
    span: SourceSpan
    order: int

    @property
    def is_opener(self) -> bool:
        """Return whether this marker opens a rectangular construct."""

        return isinstance(self.directive, (ForDirective, IfDirective, RegionDirective))

    @property
    def is_closer(self) -> bool:
        """Return whether this marker closes a rectangular construct."""

        return isinstance(self.directive, (EndForDirective, EndIfDirective, EndRegionDirective))


@dataclass(frozen=True)
class _Pair:
    opener: _Marker
    closer: _Marker
    rectangle: Rectangle


@dataclass(frozen=True)
class _RegionSpec:
    pair: _Pair
    else_marker: _Marker | None = None

    @property
    def rectangle(self) -> Rectangle:
        """Return the source rectangle derived from the paired markers."""

        return self.pair.rectangle


def _marker_kinds_match(opener: _Marker, closer: _Marker) -> bool:
    """Return whether two markers are matching directive kinds.

    Args:
        opener: Candidate opening marker.
        closer: Candidate closing marker.

    Returns:
        ``True`` for ``for/endfor``, ``if/endif``, or ``region/endregion``.
    """

    return (
        (
            isinstance(opener.directive, ForDirective)
            and isinstance(closer.directive, EndForDirective)
        )
        or (
            isinstance(opener.directive, IfDirective)
            and isinstance(closer.directive, EndIfDirective)
        )
        or (
            isinstance(opener.directive, RegionDirective)
            and isinstance(closer.directive, EndRegionDirective)
        )
    )


def _can_pair(opener: _Marker, closer: _Marker) -> bool:
    """Return whether two matching markers have valid corner ordering.

    Args:
        opener: Candidate top-left marker.
        closer: Candidate bottom-right marker.

    Returns:
        ``True`` when kinds, coordinates, and same-cell token order are valid.
    """

    if not _marker_kinds_match(opener, closer):
        return False
    start = opener.coordinate
    end = closer.coordinate
    if start.row > end.row or start.column > end.column:
        return False
    return not (start == end and opener.order >= closer.order)


def _compatible_rectangles(rectangles: list[Rectangle]) -> tuple[bool, bool]:
    """Check whether rectangles form a nested-or-disjoint tree.

    Args:
        rectangles: Candidate source rectangles.

    Returns:
        ``(compatible, saw_partial_overlap)`` for the complete collection.
    """

    for index, first in enumerate(rectangles):
        for second in rectangles[index + 1 :]:
            if first == second:
                return False, False
            if first.is_disjoint(second):
                continue
            if first.contains(second, strict=True) or second.contains(first, strict=True):
                continue
            return False, True
    return True, False


def _pair_markers(
    openers: list[_Marker],
    closers: list[_Marker],
) -> tuple[list[_Pair] | None, Diagnostic | None]:
    """Find the unique globally compatible opener-to-closer pairing.

    Args:
        openers: Structural opening markers in worksheet order.
        closers: Structural closing markers in worksheet order.

    Returns:
        Either the unique pair list and no diagnostic, or no list and one diagnostic.
    """

    if len(openers) != len(closers):
        marker = (openers + closers)[0]
        return None, Diagnostic(
            DiagnosticCode.UNMATCHED_BLOCK_MARKER,
            "opening and closing block marker counts do not match",
            marker.span.location,
        )
    candidates = {
        index: [
            close_index for close_index, close in enumerate(closers) if _can_pair(opener, close)
        ]
        for index, opener in enumerate(openers)
    }
    if any(not options for options in candidates.values()):
        marker = next(openers[index] for index, options in candidates.items() if not options)
        return None, Diagnostic(
            DiagnosticCode.UNMATCHED_BLOCK_MARKER,
            "no spatially valid closing marker exists for this block",
            marker.span.location,
        )

    solutions: list[list[_Pair]] = []
    saw_partial_overlap = False

    def search(index: int, used: set[int], pairs: list[_Pair]) -> None:
        """Enumerate compatible pairings with backtracking.

        Args:
            index: Index of the opener currently being paired.
            used: Closing-marker indexes already allocated.
            pairs: Mutable candidate pair stack for the current search branch.
        """

        nonlocal saw_partial_overlap
        if index == len(openers):
            valid, partial = _compatible_rectangles([pair.rectangle for pair in pairs])
            saw_partial_overlap = saw_partial_overlap or partial
            if valid:
                solutions.append(list(pairs))
            return
        opener = openers[index]
        for close_index in candidates[index]:
            if close_index in used:
                continue
            closer = closers[close_index]
            pair = _Pair(opener, closer, Rectangle.between(opener.coordinate, closer.coordinate))
            used.add(close_index)
            pairs.append(pair)
            search(index + 1, used, pairs)
            pairs.pop()
            used.remove(close_index)

    search(0, set(), [])
    if not solutions:
        marker = openers[0] if openers else closers[0]
        code = (
            DiagnosticCode.PARTIAL_BLOCK_OVERLAP
            if saw_partial_overlap
            else DiagnosticCode.INVALID_BLOCK_GEOMETRY
        )
        message = (
            "block rectangles partially overlap"
            if saw_partial_overlap
            else "block markers cannot form a valid nested or disjoint region tree"
        )
        return None, Diagnostic(code, message, marker.span.location)
    if len(solutions) > 1:
        return None, Diagnostic(
            DiagnosticCode.AMBIGUOUS_BLOCK_PAIRING,
            "block markers have more than one valid spatial pairing",
            openers[0].span.location,
        )
    return solutions[0], None


def _associate_else_markers(
    pairs: list[_Pair], else_markers: list[_Marker]
) -> tuple[list[_RegionSpec] | None, Diagnostic | None]:
    """Attach each ``else`` marker to its nearest compatible ``if`` pair.

    Args:
        pairs: Unbranched structural marker pairs.
        else_markers: Branch markers awaiting an owning conditional.

    Returns:
        Region specifications with branch markers, or one ownership diagnostic.
    """

    assignments: dict[_Pair, list[_Marker]] = {pair: [] for pair in pairs}
    if_pairs = [pair for pair in pairs if isinstance(pair.opener.directive, IfDirective)]
    for marker in else_markers:
        candidates = [
            pair
            for pair in if_pairs
            if pair.rectangle.contains_coordinate(marker.coordinate)
            and marker.coordinate.column == pair.rectangle.right
            and marker.coordinate.row < pair.rectangle.bottom
        ]
        if not candidates:
            return None, Diagnostic(
                DiagnosticCode.UNMATCHED_BLOCK_MARKER,
                "else marker is not inside a compatible if rectangle",
                marker.span.location,
            )
        smallest_area = min(pair.rectangle.area for pair in candidates)
        nearest = [pair for pair in candidates if pair.rectangle.area == smallest_area]
        if len(nearest) != 1:
            return None, Diagnostic(
                DiagnosticCode.AMBIGUOUS_BLOCK_PAIRING,
                "else marker has more than one possible owning if block",
                marker.span.location,
            )
        assignments[nearest[0]].append(marker)

    for markers in assignments.values():
        if len(markers) > 1:
            return None, Diagnostic(
                DiagnosticCode.AMBIGUOUS_BLOCK_PAIRING,
                "an if block contains more than one possible else marker",
                markers[1].span.location,
            )
    specs = [
        _RegionSpec(pair, assignments[pair][0] if assignments[pair] else None) for pair in pairs
    ]
    return specs, None


def _validate_marker_position(
    token_index: int,
    tokens: tuple[TextToken | OutputToken | DirectiveToken, ...],
    marker: _Marker,
) -> Diagnostic | None:
    """Validate the token-boundary rule for a structural marker.

    Args:
        token_index: Marker position within the cell's token sequence.
        tokens: Complete token sequence for the source cell.
        marker: Parsed marker being validated.

    Returns:
        A positioning diagnostic, or ``None`` when valid.
    """

    if marker.is_opener:
        valid = all(
            isinstance(token, TextToken) and not token.text.strip()
            for token in tokens[:token_index]
        )
        if not valid:
            return Diagnostic(
                DiagnosticCode.INVALID_MARKER_POSITION,
                "opening directive must be the first non-whitespace template token in its cell",
                marker.span.location,
            )
    if marker.is_closer or isinstance(marker.directive, ElseDirective):
        valid = all(
            isinstance(token, TextToken) and not token.text.strip()
            for token in tokens[token_index + 1 :]
        )
        if not valid:
            return Diagnostic(
                DiagnosticCode.INVALID_MARKER_POSITION,
                "closing directive must be the last non-whitespace template token in its cell",
                marker.span.location,
            )
    return None


def _parent_of(spec: _RegionSpec, specs: list[_RegionSpec]) -> _RegionSpec | None:
    """Return the smallest strict rectangle containing a specification.

    Args:
        spec: Child candidate whose parent is requested.
        specs: All worksheet structural specifications.

    Returns:
        The nearest containing specification, or ``None`` at worksheet level.
    """

    candidates = [
        possible
        for possible in specs
        if possible is not spec and possible.rectangle.contains(spec.rectangle, strict=True)
    ]
    if not candidates:
        return None
    return min(candidates, key=lambda item: item.rectangle.area)


def _children_of(spec: _RegionSpec | None, specs: list[_RegionSpec]) -> list[_RegionSpec]:
    """Return direct children of a specification or the worksheet root.

    Args:
        spec: Parent specification, or ``None`` for top-level nodes.
        specs: All worksheet structural specifications.

    Returns:
        Specifications whose nearest parent is ``spec``.
    """

    return [candidate for candidate in specs if _parent_of(candidate, specs) is spec]


def _spec_shift(spec: _RegionSpec, specs: list[_RegionSpec]) -> str:
    """Resolve a specification's effective row or cell shift policy.

    Args:
        spec: Structural specification to resolve.
        specs: All specifications used to locate inherited parents.

    Returns:
        ``"rows"`` or ``"cells"`` after applying conditional inheritance.
    """

    directive = spec.pair.opener.directive
    if isinstance(directive, (ForDirective, RegionDirective)):
        return directive.shift
    parent = _parent_of(spec, specs)
    return "rows" if parent is None else _spec_shift(parent, specs)


def _validate_sibling_shift_lanes(
    siblings: list[_RegionSpec], specs: list[_RegionSpec], diagnostics: list[Diagnostic]
) -> None:
    """Report siblings that ambiguously claim overlapping worksheet rows.

    Args:
        siblings: Direct sibling specifications to compare.
        specs: All specifications used to resolve inherited shift policies.
        diagnostics: Mutable diagnostic accumulator.
    """

    for index, first in enumerate(siblings):
        for second in siblings[index + 1 :]:
            rows_overlap = not (
                first.rectangle.bottom < second.rectangle.top
                or second.rectangle.bottom < first.rectangle.top
            )
            if rows_overlap and "rows" in {
                _spec_shift(first, specs),
                _spec_shift(second, specs),
            }:
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.OVERLAPPING_ROW_SHIFTS,
                        'sibling blocks with overlapping rows must both use shift="cells"',
                        second.pair.opener.span.location,
                    )
                )


def _validate_merged_ranges(
    template: WorksheetTemplate,
    specs: list[_RegionSpec],
    diagnostics: list[Diagnostic],
) -> None:
    """Reject merges bisected by block rectangles or active cell-shift lanes.

    Args:
        template: Adapter-neutral worksheet containing merged ranges.
        specs: All structural specifications on the worksheet.
        diagnostics: Mutable diagnostic accumulator.
    """

    for merged in template.merged_ranges:
        for spec in specs:
            region = spec.rectangle
            if merged.is_disjoint(region) or region.contains(merged):
                pass
            else:
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.MERGE_CROSSES_BLOCK_BOUNDARY,
                        "merged range crosses a structural block boundary",
                        SourceLocation(
                            template.name,
                            Coordinate(merged.top, merged.left).a1,
                        ),
                    )
                )
                continue

            if _spec_shift(spec, specs) != "cells" or merged.bottom <= region.bottom:
                continue
            parent = _parent_of(spec, specs)
            if parent is not None and merged.top > parent.rectangle.bottom:
                continue
            overlaps_lane = not (merged.right < region.left or merged.left > region.right)
            contained_in_lane = region.left <= merged.left and merged.right <= region.right
            if overlaps_lane and not contained_in_lane:
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.MERGE_CROSSES_BLOCK_BOUNDARY,
                        'merged range crosses a shift="cells" lane boundary',
                        SourceLocation(
                            template.name,
                            Coordinate(merged.top, merged.left).a1,
                        ),
                    )
                )


def _make_node(
    spec: _RegionSpec,
    specs: list[_RegionSpec],
    diagnostics: list[Diagnostic],
    containing_shift: str,
) -> StructuralNode:
    """Recursively convert one linked specification into a typed AST node.

    Args:
        spec: Linked source specification to compile.
        specs: All specifications used to locate direct children.
        diagnostics: Mutable semantic diagnostic accumulator.
        containing_shift: Shift policy inherited by conditional nodes.

    Returns:
        A ``ForNode``, ``IfNode``, or explicit ``RegionNode``.

    Raises:
        TypeError: If the specification contains an unexpected opener type.
    """

    child_specs = _children_of(spec, specs)
    _validate_sibling_shift_lanes(child_specs, specs, diagnostics)
    directive = spec.pair.opener.directive
    child_containing_shift = (
        directive.shift
        if isinstance(directive, (ForDirective, RegionDirective))
        else containing_shift
    )
    child_nodes = tuple(
        _make_node(child, specs, diagnostics, child_containing_shift)
        for child in sorted(child_specs, key=lambda item: (item.rectangle.top, item.rectangle.left))
    )
    if isinstance(directive, ForDirective):
        return ForNode(
            spec.rectangle,
            child_nodes,
            directive.variable,
            directive.iterable,
            directive.direction,
            directive.shift,
            spec.pair.opener.span,
        )
    if isinstance(directive, RegionDirective):
        return RegionNode(
            spec.rectangle,
            child_nodes,
            directive.direction,
            directive.shift,
            spec.pair.opener.span,
        )
    if not isinstance(directive, IfDirective):
        raise TypeError(f"unexpected opener: {type(directive).__name__}")
    if spec.else_marker is None:
        true_rectangle = spec.rectangle
        false_rectangle = None
    else:
        true_rectangle = Rectangle.between(spec.pair.opener.coordinate, spec.else_marker.coordinate)
        false_top = spec.else_marker.coordinate.row + 1
        if false_top > spec.rectangle.bottom:
            diagnostics.append(
                Diagnostic(
                    DiagnosticCode.INVALID_BLOCK_GEOMETRY,
                    "else branch must contain at least one row",
                    spec.else_marker.span.location,
                )
            )
            false_top = spec.rectangle.bottom
        false_rectangle = Rectangle(
            false_top,
            spec.rectangle.left,
            spec.rectangle.bottom,
            spec.rectangle.right,
        )
    for child in child_specs:
        in_true = true_rectangle.contains(child.rectangle)
        in_false = false_rectangle is not None and false_rectangle.contains(child.rectangle)
        if not in_true and not in_false:
            diagnostics.append(
                Diagnostic(
                    DiagnosticCode.INVALID_BLOCK_GEOMETRY,
                    "nested block crosses a conditional branch boundary",
                    child.pair.opener.span.location,
                )
            )
    return IfNode(
        spec.rectangle,
        child_nodes,
        directive.condition,
        true_rectangle,
        false_rectangle,
        spec.pair.opener.span,
        containing_shift,
    )


def _compile_cell(
    template: WorksheetTemplate,
    coordinate: Coordinate,
    value: Any,
    diagnostics: list[Diagnostic],
) -> tuple[CellNode, list[_Marker]]:
    """Compile one raw cell into output parts and structural markers.

    Args:
        template: Owning worksheet template.
        coordinate: Source coordinate of the cell.
        value: Raw cell value.
        diagnostics: Mutable lexical and syntax diagnostic accumulator.

    Returns:
        The compiled cell node and any markers found in the cell.
    """

    if not isinstance(value, str):
        return CellNode(coordinate, (LiteralPart(value),)), []
    lexed = lex_cell(template.name, coordinate.a1, value)
    diagnostics.extend(lexed.diagnostics)
    if lexed.diagnostics:
        return CellNode(coordinate, ()), []
    parts: list[LiteralPart | ExpressionPart] = []
    markers: list[_Marker] = []
    for token_index, token in enumerate(lexed.tokens):
        if isinstance(token, TextToken):
            if token.text:
                parts.append(LiteralPart(token.text))
        elif isinstance(token, OutputToken):
            try:
                parts.append(ExpressionPart(parse_expression(token.source), token.span))
            except ExpressionSyntaxError as error:
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.INVALID_EXPRESSION,
                        str(error),
                        SourceLocation(
                            template.name,
                            coordinate.a1,
                            token.span.start + 2 + error.position,
                            token.span.end,
                        ),
                    )
                )
        else:
            try:
                directive = parse_directive(token.source)
            except (DirectiveSyntaxError, ExpressionSyntaxError) as error:
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.INVALID_DIRECTIVE,
                        str(error),
                        token.span.location,
                    )
                )
                continue
            marker = _Marker(coordinate, directive, token.span, token_index)
            markers.append(marker)
            position_error = _validate_marker_position(token_index, lexed.tokens, marker)
            if position_error is not None:
                diagnostics.append(position_error)
    return CellNode(coordinate, tuple(parts)), markers


def compile_sheet(template: WorksheetTemplate) -> CompilationResult:
    """Lex, parse, spatially link, and validate a worksheet template.

    Args:
        template: Adapter-neutral source worksheet.

    Returns:
        A compiled immutable worksheet AST or structured diagnostics.
    """

    diagnostics: list[Diagnostic] = []
    cells: dict[Coordinate, CellNode] = {}
    markers: list[_Marker] = []
    for coordinate, value in sorted(template.cells.items()):
        cell, cell_markers = _compile_cell(template, coordinate, value, diagnostics)
        cells[coordinate] = cell
        markers.extend(cell_markers)
    if diagnostics:
        return CompilationResult(None, tuple(diagnostics))

    openers = [marker for marker in markers if marker.is_opener]
    closers = [marker for marker in markers if marker.is_closer]
    else_markers = [marker for marker in markers if isinstance(marker.directive, ElseDirective)]
    pairs, pairing_error = _pair_markers(openers, closers)
    if pairing_error is not None:
        return CompilationResult(None, (pairing_error,))
    assert pairs is not None
    specs, else_error = _associate_else_markers(pairs, else_markers)
    if else_error is not None:
        return CompilationResult(None, (else_error,))
    assert specs is not None

    top_level_specs = _children_of(None, specs)
    _validate_sibling_shift_lanes(top_level_specs, specs, diagnostics)
    _validate_merged_ranges(template, specs, diagnostics)
    children = tuple(
        _make_node(spec, specs, diagnostics, "rows")
        for spec in sorted(
            top_level_specs,
            key=lambda item: (item.rectangle.top, item.rectangle.left),
        )
    )
    if diagnostics:
        return CompilationResult(None, tuple(diagnostics))
    rectangle = Rectangle(1, 1, max(1, template.max_row), max(1, template.max_column))
    return CompilationResult(CompiledSheet(template, rectangle, cells, children), ())
