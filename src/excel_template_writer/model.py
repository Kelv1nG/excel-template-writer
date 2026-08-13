"""Pure worksheet geometry used by the compiler and layout planner."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from types import MappingProxyType
from typing import Any


@dataclass(frozen=True, order=True)
class Coordinate:
    """One-based worksheet coordinate."""

    row: int
    column: int

    def __post_init__(self) -> None:
        if self.row < 1 or self.column < 1:
            raise ValueError("worksheet coordinates are one-based")

    @property
    def a1(self) -> str:
        column = self.column
        letters = ""
        while column:
            column, remainder = divmod(column - 1, 26)
            letters = chr(65 + remainder) + letters
        return f"{letters}{self.row}"

    @classmethod
    def from_a1(cls, value: str) -> Coordinate:
        letters = ""
        digits = ""
        for character in value.upper():
            if character.isalpha() and not digits:
                letters += character
            elif character.isdigit():
                digits += character
            else:
                raise ValueError(f"invalid A1 coordinate: {value!r}")
        if not letters or not digits:
            raise ValueError(f"invalid A1 coordinate: {value!r}")
        column = 0
        for character in letters:
            column = column * 26 + ord(character) - 64
        return cls(row=int(digits), column=column)


@dataclass(frozen=True)
class Rectangle:
    top: int
    left: int
    bottom: int
    right: int

    def __post_init__(self) -> None:
        if min(self.top, self.left) < 1 or self.bottom < self.top or self.right < self.left:
            raise ValueError("invalid worksheet rectangle")

    @classmethod
    def between(cls, start: Coordinate, end: Coordinate) -> Rectangle:
        return cls(start.row, start.column, end.row, end.column)

    @property
    def height(self) -> int:
        return self.bottom - self.top + 1

    @property
    def width(self) -> int:
        return self.right - self.left + 1

    @property
    def area(self) -> int:
        return self.height * self.width

    def contains_coordinate(self, coordinate: Coordinate) -> bool:
        return (
            self.top <= coordinate.row <= self.bottom
            and self.left <= coordinate.column <= self.right
        )

    def contains(self, other: Rectangle, *, strict: bool = False) -> bool:
        contained = (
            self.top <= other.top
            and self.left <= other.left
            and self.bottom >= other.bottom
            and self.right >= other.right
        )
        return contained and (not strict or self != other)

    def is_disjoint(self, other: Rectangle) -> bool:
        return (
            self.bottom < other.top
            or other.bottom < self.top
            or self.right < other.left
            or other.right < self.left
        )


@dataclass(frozen=True)
class WorksheetTemplate:
    """Adapter-neutral worksheet values keyed by one-based coordinates."""

    name: str
    cells: Mapping[Coordinate, Any]

    def __post_init__(self) -> None:
        object.__setattr__(self, "cells", MappingProxyType(dict(self.cells)))

    @classmethod
    def from_rows(cls, name: str, rows: Sequence[Sequence[Any]]) -> WorksheetTemplate:
        cells = {
            Coordinate(row_index, column_index): value
            for row_index, row in enumerate(rows, start=1)
            for column_index, value in enumerate(row, start=1)
            if value is not None
        }
        return cls(name=name, cells=cells)

    @classmethod
    def from_cells(
        cls,
        name: str,
        cells: Mapping[str | Coordinate, Any],
    ) -> WorksheetTemplate:
        normalized: dict[Coordinate, Any] = {}
        for key, value in cells.items():
            coordinate = Coordinate.from_a1(key) if isinstance(key, str) else key
            normalized[coordinate] = value
        return cls(name=name, cells=normalized)

    @property
    def max_row(self) -> int:
        return max((coordinate.row for coordinate in self.cells), default=0)

    @property
    def max_column(self) -> int:
        return max((coordinate.column for coordinate in self.cells), default=0)
