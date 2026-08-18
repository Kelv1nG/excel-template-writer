"""Immutable snapshots owned by the openpyxl integration boundary."""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from types import MappingProxyType
from typing import Any

from excel_template_writer.model import Coordinate, WorksheetTemplate


@dataclass(frozen=True)
class CellPresentation:
    font: Any
    fill: Any
    border: Any
    alignment: Any
    number_format: str
    protection: Any
    quote_prefix: bool
    hyperlink: Any | None
    comment: Any | None
    is_formula: bool


@dataclass(frozen=True)
class DimensionPresentation:
    hidden: bool
    outline_level: int
    collapsed: bool
    font: Any
    fill: Any
    border: Any
    alignment: Any
    number_format: str
    protection: Any


@dataclass(frozen=True)
class RowPresentation(DimensionPresentation):
    height: float | None
    thick_top: bool
    thick_bottom: bool


@dataclass(frozen=True)
class ColumnPresentation(DimensionPresentation):
    width: float | None
    best_fit: bool


@dataclass(frozen=True)
class SheetSnapshot:
    template: WorksheetTemplate
    cells: Mapping[Coordinate, CellPresentation]
    rows: Mapping[int, RowPresentation]
    columns: Mapping[int, ColumnPresentation]
    freeze_panes: str | None
    show_grid_lines: bool | None
    tab_color: Any | None
    default_row_height: float | None
    default_column_width: float | None
    conditional_formatting: Any
    data_validations: Any
    auto_filter: Any
    formula_cells: tuple[Coordinate, ...]
    hyperlink_cells: tuple[Coordinate, ...]
    comment_cells: tuple[Coordinate, ...]
    has_conditional_formatting: bool
    has_data_validations: bool
    has_tables: bool
    has_drawings: bool

    def __post_init__(self) -> None:
        object.__setattr__(self, "cells", MappingProxyType(dict(self.cells)))
        object.__setattr__(self, "rows", MappingProxyType(dict(self.rows)))
        object.__setattr__(self, "columns", MappingProxyType(dict(self.columns)))


@dataclass(frozen=True)
class WorkbookSnapshot:
    sheets: tuple[SheetSnapshot, ...]
    properties: Any
    loaded_theme: bytes | None
