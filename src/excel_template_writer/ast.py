"""Typed abstract syntax tree for a compiled worksheet template."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from types import MappingProxyType
from typing import Any

from excel_template_writer.expressions import Expression
from excel_template_writer.model import Coordinate, Rectangle, WorksheetTemplate
from excel_template_writer.syntax import SourceSpan


class CellPart:
    pass


@dataclass(frozen=True)
class LiteralPart(CellPart):
    value: Any


@dataclass(frozen=True)
class ExpressionPart(CellPart):
    expression: Expression
    span: SourceSpan


@dataclass(frozen=True)
class CellNode:
    coordinate: Coordinate
    parts: tuple[CellPart, ...]


class RegionNode:
    rectangle: Rectangle
    children: tuple[RegionNode, ...]
    shift: str
    span: SourceSpan


@dataclass(frozen=True)
class ForNode(RegionNode):
    rectangle: Rectangle
    children: tuple[RegionNode, ...]
    variable: str
    iterable: Expression
    direction: str
    shift: str
    span: SourceSpan


@dataclass(frozen=True)
class IfNode(RegionNode):
    rectangle: Rectangle
    children: tuple[RegionNode, ...]
    condition: Expression
    true_rectangle: Rectangle
    false_rectangle: Rectangle | None
    span: SourceSpan
    shift: str = "rows"


@dataclass(frozen=True)
class CompiledSheet:
    template: WorksheetTemplate
    rectangle: Rectangle
    cells: Mapping[Coordinate, CellNode]
    children: tuple[RegionNode, ...]

    def __post_init__(self) -> None:
        object.__setattr__(self, "cells", MappingProxyType(dict(self.cells)))
