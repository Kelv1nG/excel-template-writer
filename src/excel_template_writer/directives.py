"""Parser for structural template directives."""

from __future__ import annotations

import ast as python_ast
import re
from dataclasses import dataclass

from excel_template_writer.expressions import Expression, parse_expression


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


_FOR_HEADER = re.compile(r"for\s+([A-Za-z_][A-Za-z0-9_]*)\s+in\s+(.+)", re.DOTALL)
_OPTION = re.compile(r"\s*(direction|shift)\s*=\s*('(?:\\.|[^'\\])*'|\"(?:\\.|[^\"\\])*\")")


def _find_option_start(source: str) -> int | None:
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


def _parse_for(source: str) -> ForDirective:
    match = _FOR_HEADER.fullmatch(source)
    if match is None:
        raise DirectiveSyntaxError("expected 'for <name> in <expression>'")
    variable, remainder = match.groups()
    if variable.startswith("_"):
        raise DirectiveSyntaxError("loop variable names cannot be private")
    option_start = _find_option_start(remainder)
    expression_source = remainder if option_start is None else remainder[:option_start]
    option_source = "" if option_start is None else remainder[option_start:]
    iterable = parse_expression(expression_source.strip())
    options: dict[str, str] = {}
    position = 0
    while position < len(option_source):
        option = _OPTION.match(option_source, position)
        if option is None:
            raise DirectiveSyntaxError(f"invalid loop option near {option_source[position:]!r}")
        name, literal = option.groups()
        if name in options:
            raise DirectiveSyntaxError(f"duplicate loop option: {name}")
        options[name] = python_ast.literal_eval(literal)
        position = option.end()
    direction = options.get("direction", "down")
    shift = options.get("shift", "rows")
    if direction != "down":
        raise DirectiveSyntaxError('Phase 0 supports only direction="down"')
    if shift not in {"rows", "cells"}:
        raise DirectiveSyntaxError('shift must be "rows" or "cells"')
    return ForDirective(variable, iterable, direction, shift)


def parse_directive(source: str) -> Directive:
    normalized = source.strip()
    if normalized.startswith("for "):
        return _parse_for(normalized)
    if normalized == "endfor":
        return EndForDirective()
    if normalized.startswith("if "):
        return IfDirective(parse_expression(normalized[3:].strip()))
    if normalized == "else":
        return ElseDirective()
    if normalized == "endif":
        return EndIfDirective()
    raise DirectiveSyntaxError(f"unknown directive: {normalized or '<empty>'}")
