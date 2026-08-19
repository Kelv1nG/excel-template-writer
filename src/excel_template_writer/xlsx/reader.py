"""Read an XLSX workbook into adapter-owned immutable snapshots."""

from __future__ import annotations

from collections.abc import Iterable
from copy import copy
from pathlib import Path
from typing import Any, cast

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.worksheet import Worksheet

from excel_template_writer.model import Coordinate, Rectangle, WorksheetTemplate
from excel_template_writer.xlsx.model import (
    CellPresentation,
    ColumnPresentation,
    DimensionPresentation,
    RowPresentation,
    SheetSnapshot,
    WorkbookSnapshot,
)


def _dimension_presentation(dimension: Any) -> DimensionPresentation:
    """Detach formatting shared by an openpyxl row or column dimension.

    Args:
        dimension: Openpyxl row or column dimension.

    Returns:
        Adapter-owned immutable presentation data.
    """

    return DimensionPresentation(
        hidden=bool(dimension.hidden),
        outline_level=int(dimension.outlineLevel or 0),
        collapsed=bool(dimension.collapsed),
        font=copy(dimension.font),
        fill=copy(dimension.fill),
        border=copy(dimension.border),
        alignment=copy(dimension.alignment),
        number_format=dimension.number_format,
        protection=copy(dimension.protection),
    )


def _read_rows(sheet: Worksheet) -> dict[int, RowPresentation]:
    """Snapshot explicitly configured worksheet row dimensions.

    Args:
        sheet: Openpyxl worksheet to read.

    Returns:
        Row numbers mapped to detached presentation data.
    """

    rows: dict[int, RowPresentation] = {}
    for index, dimension in sheet.row_dimensions.items():
        common = _dimension_presentation(dimension)
        rows[index] = RowPresentation(
            **vars(common),
            height=dimension.height,
            thick_top=bool(dimension.thickTop),
            thick_bottom=bool(dimension.thickBot),
        )
    return rows


def _read_columns(sheet: Worksheet) -> dict[int, ColumnPresentation]:
    """Expand and snapshot explicitly configured column dimensions.

    Args:
        sheet: Openpyxl worksheet to read.

    Returns:
        Column indexes mapped to detached presentation data.
    """

    columns: dict[int, ColumnPresentation] = {}
    for key, dimension in sheet.column_dimensions.items():
        common = _dimension_presentation(dimension)
        minimum = dimension.min or column_index_from_string(key)
        maximum = dimension.max or minimum
        for index in range(minimum, maximum + 1):
            columns[index] = ColumnPresentation(
                **vars(common),
                width=dimension.width,
                best_fit=bool(dimension.bestFit),
            )
    return columns


def _read_merges(sheet: Worksheet) -> tuple[Rectangle, ...]:
    """Convert openpyxl merged ranges into pure inclusive rectangles.

    Args:
        sheet: Openpyxl worksheet to inspect.

    Returns:
        Immutable merged rectangles in worksheet order.
    """

    ranges = cast(Iterable[CellRange], sheet.merged_cells.ranges)
    return tuple(
        Rectangle(
            cast(int, merged.min_row),
            cast(int, merged.min_col),
            cast(int, merged.max_row),
            cast(int, merged.max_col),
        )
        for merged in ranges
    )


def _optional_float(value: Any) -> float | None:
    """Convert an optional numeric openpyxl property to ``float``.

    Args:
        value: Optional numeric property.

    Returns:
        ``None`` or the equivalent float value.
    """

    return None if value is None else float(value)


def _merge_coordinates(merged_ranges: tuple[Rectangle, ...]) -> set[Coordinate]:
    """Enumerate every coordinate occupied by merged ranges.

    Args:
        merged_ranges: Inclusive merged rectangles.

    Returns:
        Coordinates belonging to at least one merge.
    """

    return {
        Coordinate(row, column)
        for merged in merged_ranges
        for row in range(merged.top, merged.bottom + 1)
        for column in range(merged.left, merged.right + 1)
    }


def _read_sheet(sheet: Worksheet) -> SheetSnapshot:
    """Detach supported values, presentation, dimensions, and feature flags.

    Args:
        sheet: Openpyxl worksheet to snapshot.

    Returns:
        Immutable adapter-owned worksheet state.
    """

    merged_ranges = _read_merges(sheet)
    merge_coordinates = _merge_coordinates(merged_ranges)
    values: dict[Coordinate, Any] = {}
    presentations: dict[Coordinate, CellPresentation] = {}
    formula_cells: list[Coordinate] = []
    hyperlink_cells: list[Coordinate] = []
    comment_cells: list[Coordinate] = []

    for row in sheet.iter_rows():
        for cell in row:
            coordinate = Coordinate(cell.row, cell.column)
            is_material = (
                cell.value is not None
                or cell.has_style
                or getattr(cell, "hyperlink", None) is not None
                or getattr(cell, "comment", None) is not None
                or coordinate in merge_coordinates
            )
            if not is_material:
                continue
            value = cell.value if not isinstance(cell, MergedCell) else None
            is_formula = getattr(cell, "data_type", None) == "f"
            hyperlink = copy(getattr(cell, "hyperlink", None))
            comment = copy(getattr(cell, "comment", None))
            values[coordinate] = value
            presentations[coordinate] = CellPresentation(
                font=copy(cell.font),
                fill=copy(cell.fill),
                border=copy(cell.border),
                alignment=copy(cell.alignment),
                number_format=cell.number_format,
                protection=copy(cell.protection),
                quote_prefix=bool(getattr(cell, "quotePrefix", False)),
                hyperlink=hyperlink,
                comment=comment,
                is_formula=is_formula,
            )
            if is_formula:
                formula_cells.append(coordinate)
            if hyperlink is not None:
                hyperlink_cells.append(coordinate)
            if comment is not None:
                comment_cells.append(coordinate)

    freeze_panes = sheet.freeze_panes
    if freeze_panes is not None and not isinstance(freeze_panes, str):
        freeze_panes = freeze_panes.coordinate
    return SheetSnapshot(
        template=WorksheetTemplate(sheet.title, values, merged_ranges),
        cells=presentations,
        rows=_read_rows(sheet),
        columns=_read_columns(sheet),
        freeze_panes=freeze_panes,
        show_grid_lines=sheet.sheet_view.showGridLines,
        tab_color=copy(sheet.sheet_properties.tabColor),
        default_row_height=_optional_float(sheet.sheet_format.defaultRowHeight),
        default_column_width=_optional_float(sheet.sheet_format.defaultColWidth),
        conditional_formatting=copy(sheet.conditional_formatting),
        data_validations=copy(sheet.data_validations),
        auto_filter=copy(sheet.auto_filter),
        formula_cells=tuple(formula_cells),
        hyperlink_cells=tuple(hyperlink_cells),
        comment_cells=tuple(comment_cells),
        has_conditional_formatting=bool(len(sheet.conditional_formatting)),
        has_data_validations=bool(sheet.data_validations.count),
        has_tables=bool(sheet.tables),
        has_drawings=bool(sheet._charts or sheet._images),
    )


def read_workbook(path: str | Path) -> WorkbookSnapshot:
    """Load a non-macro XLSX workbook and detach supported state.

    Args:
        path: Input ``.xlsx`` path.

    Returns:
        Immutable workbook snapshot containing no live openpyxl objects.
    """

    workbook = load_workbook(path, data_only=False, keep_links=False)
    try:
        return WorkbookSnapshot(
            sheets=tuple(_read_sheet(sheet) for sheet in workbook.worksheets),
            properties=copy(workbook.properties),
            loaded_theme=workbook.loaded_theme,
        )
    finally:
        workbook.close()
