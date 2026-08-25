"""Compiled, deterministic date-to-text formatting for template filters."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import Literal, cast


class DateFormatSyntaxError(ValueError):
    """A date format cannot be compiled under the supported language subset."""


class DateFormatValueError(ValueError):
    """A runtime value cannot be consumed by a compiled date format."""


@dataclass(frozen=True)
class DateField:
    """One date field in a compiled format."""

    symbol: Literal["y", "m", "d"]
    width: int


@dataclass(frozen=True)
class DateLiteral:
    """Literal text in a compiled date format."""

    text: str


type DateFormatPart = DateField | DateLiteral


@dataclass(frozen=True)
class DateFormat:
    """Immutable date format produced during expression compilation."""

    source: str
    parts: tuple[DateFormatPart, ...]


_VALID_WIDTHS: dict[str, frozenset[int]] = {
    "y": frozenset({2, 4}),
    "m": frozenset({1, 2, 3, 4}),
    "d": frozenset({1, 2, 3, 4}),
}
_UNSUPPORTED_UNQUOTED = frozenset("{}%[];*_@#?")
_MONTH_NAMES = (
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)
_WEEKDAY_NAMES = (
    "Monday",
    "Tuesday",
    "Wednesday",
    "Thursday",
    "Friday",
    "Saturday",
    "Sunday",
)


def _append_literal(parts: list[DateFormatPart], text: str) -> None:
    """Append text while coalescing adjacent literal format nodes.

    Args:
        parts: Mutable compiled-part accumulator.
        text: Literal text to append; an empty value is ignored.
    """

    if not text:
        return
    if parts and isinstance(parts[-1], DateLiteral):
        previous = parts[-1]
        parts[-1] = DateLiteral(previous.text + text)
    else:
        parts.append(DateLiteral(text))


def parse_date_format(source: str) -> DateFormat:
    """Compile a constrained, case-insensitive Excel-like date format.

    Args:
        source: Author-supplied format literal from the ``date`` filter.

    Returns:
        An immutable format AST ready for deterministic evaluation.

    Raises:
        DateFormatSyntaxError: If the format is empty, malformed, or unsupported.
    """

    if not source:
        raise DateFormatSyntaxError("date format cannot be empty")

    parts: list[DateFormatPart] = []
    saw_field = False
    position = 0
    while position < len(source):
        character = source[position]
        if character == '"':
            closing = source.find('"', position + 1)
            if closing < 0:
                raise DateFormatSyntaxError("date format contains an unterminated quoted literal")
            _append_literal(parts, source[position + 1 : closing])
            position = closing + 1
            continue
        if character == "\\":
            if position + 1 == len(source):
                raise DateFormatSyntaxError("date format ends with a dangling escape")
            _append_literal(parts, source[position + 1])
            position += 2
            continue
        if character in _UNSUPPORTED_UNQUOTED:
            raise DateFormatSyntaxError(
                f"unsupported unquoted character {character!r} in date format"
            )
        if character.isalpha():
            symbol = character.lower()
            end = position + 1
            while end < len(source) and source[end].lower() == symbol:
                end += 1
            width = end - position
            if symbol not in _VALID_WIDTHS:
                raise DateFormatSyntaxError(f"unsupported date field {source[position:end]!r}")
            if width not in _VALID_WIDTHS[symbol]:
                raise DateFormatSyntaxError(
                    f"unsupported width for date field {source[position:end]!r}"
                )
            parts.append(DateField(cast(Literal["y", "m", "d"], symbol), width))
            saw_field = True
            position = end
            continue
        _append_literal(parts, character)
        position += 1

    if not saw_field:
        raise DateFormatSyntaxError("date format must contain at least one date field")
    return DateFormat(source, tuple(parts))


def _format_field(value: date | datetime, field: DateField) -> str:
    """Render one compiled field from a native temporal value.

    Args:
        value: Canonical date or datetime being formatted.
        field: Compiled field symbol and width.

    Returns:
        Deterministic text for the requested calendar field.
    """

    if field.symbol == "y":
        return f"{value.year % 100:02d}" if field.width == 2 else f"{value.year:04d}"
    if field.symbol == "m":
        if field.width == 1:
            return str(value.month)
        if field.width == 2:
            return f"{value.month:02d}"
        name = _MONTH_NAMES[value.month - 1]
        return name[:3] if field.width == 3 else name
    if field.width == 1:
        return str(value.day)
    if field.width == 2:
        return f"{value.day:02d}"
    name = _WEEKDAY_NAMES[value.weekday()]
    return name[:3] if field.width == 3 else name


def format_date(value: object, date_format: DateFormat) -> str:
    """Format a canonical date or datetime as text.

    Args:
        value: Runtime value supplied to the ``date`` filter.
        date_format: Previously compiled engine-owned format AST.

    Returns:
        Text assembled from literal and calendar-field parts.

    Raises:
        DateFormatValueError: If ``value`` is not a date or datetime.
    """

    if not isinstance(value, date):
        type_name = "null" if value is None else type(value).__name__
        raise DateFormatValueError(
            f"date filter requires a date or datetime value; received {type_name}"
        )
    return "".join(
        part.text if isinstance(part, DateLiteral) else _format_field(value, part)
        for part in date_format.parts
    )
