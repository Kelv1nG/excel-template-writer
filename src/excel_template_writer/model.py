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
        """Enforce one-based worksheet coordinates.

        Raises:
            ValueError: If either coordinate component is below one.
        """

        if self.row < 1 or self.column < 1:
            raise ValueError("worksheet coordinates are one-based")

    @property
    def a1(self) -> str:
        """Return the coordinate in Excel A1 notation.

        Returns:
            The one-based column letters and row number, such as ``B4``.
        """

        column = self.column
        letters = ""
        while column:
            column, remainder = divmod(column - 1, 26)
            letters = chr(65 + remainder) + letters
        return f"{letters}{self.row}"

    @classmethod
    def from_a1(cls, value: str) -> Coordinate:
        """Parse an Excel A1 address into a coordinate.

        Args:
            value: Address containing column letters followed by a row number.

        Returns:
            The parsed one-based coordinate.

        Raises:
            ValueError: If the address is not valid A1 notation.
        """

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
        """Enforce an ordered, positive, one-based rectangle.

        Raises:
            ValueError: If an edge is invalid or the rectangle is inverted.
        """

        if min(self.top, self.left) < 1 or self.bottom < self.top or self.right < self.left:
            raise ValueError("invalid worksheet rectangle")

    @classmethod
    def between(cls, start: Coordinate, end: Coordinate) -> Rectangle:
        """Create a rectangle from its top-left and bottom-right coordinates.

        Args:
            start: Inclusive top-left coordinate.
            end: Inclusive bottom-right coordinate.

        Returns:
            The inclusive rectangle between the two coordinates.
        """

        return cls(start.row, start.column, end.row, end.column)

    @property
    def height(self) -> int:
        """Return the inclusive row count."""

        return self.bottom - self.top + 1

    @property
    def width(self) -> int:
        """Return the inclusive column count."""

        return self.right - self.left + 1

    @property
    def area(self) -> int:
        """Return the number of cells in the rectangle."""

        return self.height * self.width

    def contains_coordinate(self, coordinate: Coordinate) -> bool:
        """Return whether a coordinate lies inside the inclusive rectangle.

        Args:
            coordinate: Coordinate to test.

        Returns:
            ``True`` when the coordinate is inside or on the boundary.
        """

        return (
            self.top <= coordinate.row <= self.bottom
            and self.left <= coordinate.column <= self.right
        )

    def contains(self, other: Rectangle, *, strict: bool = False) -> bool:
        """Return whether another rectangle is contained by this rectangle.

        Args:
            other: Rectangle to test.
            strict: Require the rectangles to differ when ``True``.

        Returns:
            ``True`` when every edge of ``other`` is within this rectangle.
        """

        contained = (
            self.top <= other.top
            and self.left <= other.left
            and self.bottom >= other.bottom
            and self.right >= other.right
        )
        return contained and (not strict or self != other)

    def is_disjoint(self, other: Rectangle) -> bool:
        """Return whether two rectangles share no cells.

        Args:
            other: Rectangle to compare.

        Returns:
            ``True`` when the rectangles do not intersect.
        """

        return (
            self.bottom < other.top
            or other.bottom < self.top
            or self.right < other.left
            or other.right < self.left
        )

    def intersects(self, other: Rectangle) -> bool:
        """Return whether two rectangles share at least one cell.

        Args:
            other: Rectangle to compare.

        Returns:
            ``True`` when the rectangles intersect.
        """

        return not self.is_disjoint(other)

    def translated(self, *, rows: int = 0, columns: int = 0) -> Rectangle:
        """Return a copy moved by signed row and column offsets.

        Args:
            rows: Signed row offset.
            columns: Signed column offset.

        Returns:
            A rectangle with every edge moved by the supplied offsets.
        """

        return Rectangle(
            self.top + rows,
            self.left + columns,
            self.bottom + rows,
            self.right + columns,
        )


@dataclass(frozen=True)
class WorksheetTemplate:
    """Adapter-neutral worksheet values keyed by one-based coordinates."""

    name: str
    cells: Mapping[Coordinate, Any]
    merged_ranges: tuple[Rectangle, ...] = ()

    def __post_init__(self) -> None:
        """Detach mutable cell mappings and normalize merged ranges to a tuple."""

        object.__setattr__(self, "cells", MappingProxyType(dict(self.cells)))
        object.__setattr__(self, "merged_ranges", tuple(self.merged_ranges))

    @classmethod
    def from_rows(cls, name: str, rows: Sequence[Sequence[Any]]) -> WorksheetTemplate:
        """Create a sparse worksheet template from row-major values.

        Args:
            name: Worksheet name.
            rows: Row-major cell values; ``None`` entries are omitted.

        Returns:
            An immutable adapter-neutral worksheet template.
        """

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
        """Create a worksheet template from A1 or coordinate keys.

        Args:
            name: Worksheet name.
            cells: Mapping of A1 strings or coordinates to raw cell values.

        Returns:
            An immutable adapter-neutral worksheet template.
        """

        normalized: dict[Coordinate, Any] = {}
        for key, value in cells.items():
            coordinate = Coordinate.from_a1(key) if isinstance(key, str) else key
            normalized[coordinate] = value
        return cls(name=name, cells=normalized)

    @property
    def max_row(self) -> int:
        """Return the greatest material source row, or zero when empty."""

        return max((coordinate.row for coordinate in self.cells), default=0)

    @property
    def max_column(self) -> int:
        """Return the greatest material source column, or zero when empty."""

        return max((coordinate.column for coordinate in self.cells), default=0)
