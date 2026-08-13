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
    ForDirective,
    IfDirective,
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
        return isinstance(self.directive, (ForDirective, IfDirective))

    @property
    def is_closer(self) -> bool:
        return isinstance(self.directive, (EndForDirective, EndIfDirective))


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
        return self.pair.rectangle


def _marker_kinds_match(opener: _Marker, closer: _Marker) -> bool:
    return (
        isinstance(opener.directive, ForDirective) and isinstance(closer.directive, EndForDirective)
    ) or (
        isinstance(opener.directive, IfDirective) and isinstance(closer.directive, EndIfDirective)
    )


def _can_pair(opener: _Marker, closer: _Marker) -> bool:
    if not _marker_kinds_match(opener, closer):
        return False
    start = opener.coordinate
    end = closer.coordinate
    if start.row > end.row or start.column > end.column:
        return False
    return not (start == end and opener.order >= closer.order)


def _compatible_rectangles(rectangles: list[Rectangle]) -> tuple[bool, bool]:
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
    candidates = [
        possible
        for possible in specs
        if possible is not spec and possible.rectangle.contains(spec.rectangle, strict=True)
    ]
    if not candidates:
        return None
    return min(candidates, key=lambda item: item.rectangle.area)


def _children_of(spec: _RegionSpec | None, specs: list[_RegionSpec]) -> list[_RegionSpec]:
    return [candidate for candidate in specs if _parent_of(candidate, specs) is spec]


def _spec_shift(spec: _RegionSpec) -> str:
    directive = spec.pair.opener.directive
    return directive.shift if isinstance(directive, ForDirective) else "rows"


def _validate_sibling_shift_lanes(
    siblings: list[_RegionSpec], diagnostics: list[Diagnostic]
) -> None:
    for index, first in enumerate(siblings):
        for second in siblings[index + 1 :]:
            rows_overlap = not (
                first.rectangle.bottom < second.rectangle.top
                or second.rectangle.bottom < first.rectangle.top
            )
            if rows_overlap and "rows" in {_spec_shift(first), _spec_shift(second)}:
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.OVERLAPPING_ROW_SHIFTS,
                        'sibling blocks with overlapping rows must both use shift="cells"',
                        second.pair.opener.span.location,
                    )
                )


def _make_node(
    spec: _RegionSpec,
    specs: list[_RegionSpec],
    diagnostics: list[Diagnostic],
    containing_shift: str,
) -> RegionNode:
    child_specs = _children_of(spec, specs)
    _validate_sibling_shift_lanes(child_specs, diagnostics)
    directive = spec.pair.opener.directive
    child_containing_shift = (
        directive.shift if isinstance(directive, ForDirective) else containing_shift
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
    """Lex, parse, spatially link, validate, and return an immutable worksheet AST."""

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
    _validate_sibling_shift_lanes(top_level_specs, diagnostics)
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
