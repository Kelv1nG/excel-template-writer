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
class ChartSnapshot:
    chart: Any
    chart_type: str
    anchor_coordinates: tuple[Coordinate, ...]
    references: tuple[str, ...]
    has_supported_type: bool
    has_supported_anchor: bool
    is_combined: bool
    is_pivot: bool


@dataclass(frozen=True)
class PlannedChart:
    anchor_coordinates: tuple[Coordinate, ...]


@dataclass(frozen=True)
class SheetFeaturePlan:
    charts: tuple[PlannedChart, ...]

    def __post_init__(self) -> None:
        """Normalize planned chart collections to immutable tuples."""

        object.__setattr__(self, "charts", tuple(self.charts))


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
    charts: tuple[ChartSnapshot, ...]
    synthetic_chart_anchor_cells: frozenset[Coordinate]
    has_unsupported_drawings: bool

    def __post_init__(self) -> None:
        """Detach mutable presentation, row, and column mappings."""

        object.__setattr__(self, "cells", MappingProxyType(dict(self.cells)))
        object.__setattr__(self, "rows", MappingProxyType(dict(self.rows)))
        object.__setattr__(self, "columns", MappingProxyType(dict(self.columns)))
        object.__setattr__(self, "charts", tuple(self.charts))
        object.__setattr__(
            self,
            "synthetic_chart_anchor_cells",
            frozenset(self.synthetic_chart_anchor_cells),
        )


@dataclass(frozen=True)
class WorkbookSnapshot:
    sheets: tuple[SheetSnapshot, ...]
    chartsheets: tuple[str, ...]
    properties: Any
    loaded_theme: bytes | None
