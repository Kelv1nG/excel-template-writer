"""Parser for structural template directives."""

from __future__ import annotations

import ast as python_ast
import re
from dataclasses import dataclass

from excel_template_writer.expressions import Expression, compile_expression


class DirectiveSyntaxError(ValueError):
    pass


class Directive:
    pass


@dataclass(frozen=True)
class ForDirective(Directive):
    variable: str
    iterable: Expression
    direction: str = "down"
    shift: str = "rows"


@dataclass(frozen=True)
class EndForDirective(Directive):
    pass


@dataclass(frozen=True)
class IfDirective(Directive):
    condition: Expression


@dataclass(frozen=True)
class ElseDirective(Directive):
    pass


@dataclass(frozen=True)
class EndIfDirective(Directive):
    pass


@dataclass(frozen=True)
class RegionDirective(Directive):
    direction: str = "down"
    shift: str = "rows"


@dataclass(frozen=True)
class EndRegionDirective(Directive):
    pass


_FOR_HEADER = re.compile(r"for\s+([A-Za-z_][A-Za-z0-9_]*)\s+in\s+(.+)", re.DOTALL)
_OPTION = re.compile(r"\s*(direction|shift)\s*=\s*('(?:\\.|[^'\\])*'|\"(?:\\.|[^\"\\])*\")")


def _find_option_start(source: str) -> int | None:
    """Locate the first top-level layout option following a loop expression.

    Args:
        source: Text following the loop variable and ``in`` keyword.

    Returns:
        The character offset of the first option, or ``None`` when absent.
    """

    quote: str | None = None
    escaped = False
    depth = 0
    for index, character in enumerate(source):
        if quote is not None:
            if escaped:
                escaped = False
            elif character == "\\":
                escaped = True
            elif character == quote:
                quote = None
            continue
        if character in {'"', "'"}:
            quote = character
        elif character in "([":
            depth += 1
        elif character in ")]":
            depth -= 1
        elif depth == 0 and character.isspace():
            tail = source[index:]
            if re.match(r"\s*(?:direction|shift)\s*=", tail):
                return index
    return None


def _parse_options(source: str, *, owner: str) -> dict[str, str]:
    """Parse quoted ``direction`` and ``shift`` assignments.

    Args:
        source: Option text to parse.
        owner: Human-readable directive kind used in error messages.

    Returns:
        Parsed option names and string values.

    Raises:
        DirectiveSyntaxError: If an option is malformed, unknown, or duplicated.
    """

    options: dict[str, str] = {}
    position = 0
    while position < len(source):
        option = _OPTION.match(source, position)
        if option is None:
            raise DirectiveSyntaxError(f"invalid {owner} option near {source[position:]!r}")
        name, literal = option.groups()
        if name in options:
            raise DirectiveSyntaxError(f"duplicate {owner} option: {name}")
        options[name] = python_ast.literal_eval(literal)
        position = option.end()
    return options


def _validate_layout_options(options: dict[str, str]) -> tuple[str, str]:
    """Apply defaults and validate currently supported layout option values.

    Args:
        options: Parsed directive options.

    Returns:
        The normalized ``(direction, shift)`` pair.

    Raises:
        DirectiveSyntaxError: If the direction or shift mode is unsupported.
    """

    direction = options.get("direction", "down")
    shift = options.get("shift", "rows")
    if direction != "down":
        raise DirectiveSyntaxError('direction must be "down"')
    if shift not in {"rows", "cells"}:
        raise DirectiveSyntaxError('shift must be "rows" or "cells"')
    return direction, shift


def _parse_for(source: str) -> ForDirective:
    """Parse a complete ``for`` opening directive.

    Args:
        source: Normalized directive source beginning with ``for``.

    Returns:
        A typed loop marker containing its expression and layout options.

    Raises:
        DirectiveSyntaxError: If the loop header, variable, or options are invalid.
        ExpressionSyntaxError: If the collection expression is invalid.
    """

    match = _FOR_HEADER.fullmatch(source)
    if match is None:
        raise DirectiveSyntaxError("expected 'for <name> in <expression>'")
    variable, remainder = match.groups()
    if variable.startswith("_"):
        raise DirectiveSyntaxError("loop variable names cannot be private")
    option_start = _find_option_start(remainder)
    expression_source = remainder if option_start is None else remainder[:option_start]
    option_source = "" if option_start is None else remainder[option_start:]
    iterable = compile_expression(expression_source.strip())
    direction, shift = _validate_layout_options(_parse_options(option_source, owner="loop"))
    return ForDirective(variable, iterable, direction, shift)


def _parse_region(source: str) -> RegionDirective:
    """Parse layout options from a ``region`` opening directive.

    Args:
        source: Text following the ``region`` keyword.

    Returns:
        A typed explicit-region marker.

    Raises:
        DirectiveSyntaxError: If a region option is invalid.
    """

    direction, shift = _validate_layout_options(_parse_options(source, owner="region"))
    return RegionDirective(direction, shift)


def parse_directive(source: str) -> Directive:
    """Parse directive source into one typed structural marker.

    Args:
        source: Text between ``{%`` and ``%}`` delimiters.

    Returns:
        The parsed directive marker.

    Raises:
        DirectiveSyntaxError: If the directive is unknown or malformed.
        ExpressionSyntaxError: If a directive expression is invalid.
    """

    normalized = source.strip()
    if normalized.startswith("for "):
        return _parse_for(normalized)
    if normalized == "endfor":
        return EndForDirective()
    if normalized.startswith("if "):
        return IfDirective(compile_expression(normalized[3:].strip()))
    if normalized == "else":
        return ElseDirective()
    if normalized == "endif":
        return EndIfDirective()
    if normalized == "region" or normalized.startswith("region "):
        return _parse_region(normalized[len("region") :])
    if normalized == "endregion":
        return EndRegionDirective()
    raise DirectiveSyntaxError(f"unknown directive: {normalized or '<empty>'}")
