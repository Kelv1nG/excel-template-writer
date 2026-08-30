"""Read an XLSX workbook into adapter-owned immutable snapshots."""

from __future__ import annotations

import posixpath
from collections.abc import Iterable
from copy import copy, deepcopy
from pathlib import Path
from typing import Any, cast
from xml.etree import ElementTree
from zipfile import ZipFile

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.chart import AreaChart, BarChart, LineChart, PieChart, ScatterChart
from openpyxl.drawing.spreadsheet_drawing import AbsoluteAnchor, OneCellAnchor, TwoCellAnchor
from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.worksheet import Worksheet

from excel_template_writer.model import Coordinate, Rectangle, WorksheetTemplate
from excel_template_writer.xlsx.model import (
    CellPresentation,
    ChartSnapshot,
    ColumnPresentation,
    DimensionPresentation,
    RowPresentation,
    SheetSnapshot,
    WorkbookSnapshot,
)

_CHART_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/chart"
_DRAWING_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
_DRAWING_RELATIONSHIP = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
)
_CHART_RELATIONSHIP = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
_OFFICE_RELATIONSHIP_NAMESPACE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
_PACKAGE_RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/package/2006/relationships"
_SUPPORTED_CHART_TYPES = (AreaChart, BarChart, LineChart, PieChart, ScatterChart)
_SUPPORTED_ANCHOR_TYPES = (AbsoluteAnchor, OneCellAnchor, TwoCellAnchor)


def _qualified_name(namespace: str, local_name: str) -> str:
    """Return one ElementTree expanded XML name.

    Args:
        namespace: XML namespace URI.
        local_name: Local element name.

    Returns:
        ElementTree expanded name in ``{namespace}local`` form.
    """

    return f"{{{namespace}}}{local_name}"


def _normalized_part_target(source_part: str, target: str) -> str:
    """Resolve one OOXML relationship target to a package-relative path.

    Args:
        source_part: Package-relative path of the relationship owner.
        target: Absolute or relative relationship target.

    Returns:
        Normalized package-relative target without a leading slash.
    """

    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source_part), target))


def _drawing_chart_relationships(
    archive: ZipFile,
    drawing_name: str,
) -> dict[str, str]:
    """Resolve chart relationship IDs owned by one drawing part.

    Args:
        archive: Open XLSX ZIP archive.
        drawing_name: Package-relative drawing-part path.

    Returns:
        Relationship IDs mapped to normalized chart-part paths.
    """

    directory, filename = posixpath.split(drawing_name)
    relationships_name = posixpath.join(directory, "_rels", f"{filename}.rels")
    if relationships_name not in archive.namelist():
        return {}
    root = ElementTree.fromstring(archive.read(relationships_name))
    relationships: dict[str, str] = {}
    for element in root:
        if (
            element.tag != _qualified_name(_PACKAGE_RELATIONSHIP_NAMESPACE, "Relationship")
            or element.get("Type") != _CHART_RELATIONSHIP
        ):
            continue
        relationship_id = element.get("Id")
        target = element.get("Target")
        if relationship_id is not None and target is not None:
            relationships[relationship_id] = _normalized_part_target(drawing_name, target)
    return relationships


def _drawing_profiles(path: Path) -> dict[str, tuple[bool, tuple[str, ...]]]:
    """Inspect drawing parts before openpyxl can discard unsupported objects.

    Args:
        path: XLSX package to inspect.

    Returns:
        Drawing-part paths mapped to ``(chart_only, chart_part_paths)``.
    """

    profiles: dict[str, tuple[bool, tuple[str, ...]]] = {}
    with ZipFile(path) as archive:
        drawing_names = (
            name
            for name in archive.namelist()
            if name.startswith("xl/drawings/drawing") and name.endswith(".xml")
        )
        for name in drawing_names:
            root = ElementTree.fromstring(archive.read(name))
            chart_relationships = _drawing_chart_relationships(archive, name)
            chart_parts: list[str] = []
            chart_only = root.tag == _qualified_name(_DRAWING_NAMESPACE, "wsDr")
            anchors = list(root)
            if not anchors:
                chart_only = False
            for anchor in anchors:
                if anchor.tag not in {
                    _qualified_name(_DRAWING_NAMESPACE, "absoluteAnchor"),
                    _qualified_name(_DRAWING_NAMESPACE, "oneCellAnchor"),
                    _qualified_name(_DRAWING_NAMESPACE, "twoCellAnchor"),
                }:
                    chart_only = False
                    continue
                frame_parts: list[str] = []
                for child in anchor:
                    if child.tag in {
                        _qualified_name(_DRAWING_NAMESPACE, "clientData"),
                        _qualified_name(_DRAWING_NAMESPACE, "ext"),
                        _qualified_name(_DRAWING_NAMESPACE, "from"),
                        _qualified_name(_DRAWING_NAMESPACE, "pos"),
                        _qualified_name(_DRAWING_NAMESPACE, "to"),
                    }:
                        continue
                    if child.tag != _qualified_name(_DRAWING_NAMESPACE, "graphicFrame"):
                        chart_only = False
                        continue
                    chart_elements = [
                        descendant
                        for descendant in child.iter()
                        if descendant.tag == _qualified_name(_CHART_NAMESPACE, "chart")
                    ]
                    if len(chart_elements) != 1:
                        chart_only = False
                        continue
                    relationship_id = chart_elements[0].get(
                        _qualified_name(_OFFICE_RELATIONSHIP_NAMESPACE, "id")
                    )
                    chart_part = chart_relationships.get(relationship_id or "")
                    if chart_part is None:
                        chart_only = False
                        continue
                    frame_parts.append(chart_part)
                if len(frame_parts) != 1:
                    chart_only = False
                chart_parts.extend(frame_parts)
            profiles[name] = chart_only, tuple(chart_parts)
    return profiles


def _chart_space_properties(path: Path) -> dict[str, tuple[Any, Any]]:
    """Read chart-space properties omitted by openpyxl's chart reader.

    Args:
        path: XLSX package containing chart XML parts.

    Returns:
        Chart-part paths mapped to style and rounded-corner values.
    """

    properties: dict[str, tuple[Any, Any]] = {}
    with ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.startswith("xl/charts/chart") or not name.endswith(".xml"):
                continue
            root = ElementTree.fromstring(archive.read(name))
            style = root.find(_qualified_name(_CHART_NAMESPACE, "style"))
            rounded = root.find(_qualified_name(_CHART_NAMESPACE, "roundedCorners"))
            style_value = None if style is None else int(style.get("val", "0"))
            rounded_value = (
                None if rounded is None else rounded.get("val", "0").lower() in {"1", "true"}
            )
            properties[name] = (style_value, rounded_value)
    return properties


def _sheet_has_unsupported_drawings(
    sheet: Worksheet,
    profiles: dict[str, tuple[bool, tuple[str, ...]]],
) -> bool:
    """Return whether a sheet drawing part contains anything except parsed charts.

    Args:
        sheet: Loaded worksheet whose drawing relationships are inspected.
        profiles: Package drawing-part classifications keyed by normalized path.

    Returns:
        ``True`` when any drawing content cannot be preserved by the chart profile.
    """

    relationships = [
        relationship for relationship in sheet._rels if relationship.Type == _DRAWING_RELATIONSHIP
    ]
    if not relationships:
        return bool(sheet._images)
    chart_count = 0
    for relationship in relationships:
        target = relationship.Target.lstrip("/")
        profile = profiles.get(target)
        if profile is None or not profile[0]:
            return True
        chart_count += len(profile[1])
    return chart_count != len(sheet._charts)


def _sheet_chart_parts(
    sheet: Worksheet,
    profiles: dict[str, tuple[bool, tuple[str, ...]]],
) -> tuple[str, ...]:
    """Return chart-part paths in the same drawing order as loaded charts.

    Args:
        sheet: Loaded worksheet whose drawing relationships are inspected.
        profiles: Package drawing-part classifications and chart targets.

    Returns:
        Normalized chart-part paths in worksheet drawing order.
    """

    parts: list[str] = []
    for relationship in sheet._rels:
        if relationship.Type != _DRAWING_RELATIONSHIP:
            continue
        profile = profiles.get(relationship.Target.lstrip("/"))
        if profile is not None:
            parts.extend(profile[1])
    return tuple(parts)


def _anchor_coordinates(anchor: Any) -> tuple[Coordinate, ...]:
    """Return the cell markers used by one supported chart anchor.

    Args:
        anchor: Openpyxl absolute, one-cell, two-cell, or unsupported anchor.

    Returns:
        One-based worksheet coordinates for every cell-based marker.
    """

    if isinstance(anchor, AbsoluteAnchor):
        return ()
    if isinstance(anchor, OneCellAnchor):
        markers = (anchor._from,)
    elif isinstance(anchor, TwoCellAnchor):
        markers = (anchor._from, anchor.to)
    else:
        return ()
    return tuple(Coordinate(marker.row + 1, marker.col + 1) for marker in markers)


def _chart_references(chart: Any) -> tuple[str, ...]:
    """Serialize one detached chart and collect every chart-formula reference.

    Args:
        chart: Detached openpyxl chart object.

    Returns:
        Formula strings in serialized chart order.
    """

    tree = deepcopy(chart)._write()
    return tuple(
        element.text
        for element in tree.iter()
        if element.tag.rsplit("}", 1)[-1] == "f" and element.text
    )


def _read_charts(
    sheet: Worksheet,
    chart_parts: tuple[str, ...],
    chart_space_properties: dict[str, tuple[Any, Any]],
) -> tuple[ChartSnapshot, ...]:
    """Detach supported and unsupported chart metadata for later validation.

    Args:
        sheet: Loaded worksheet containing zero or more charts.
        chart_parts: Actual chart-part paths in drawing order.
        chart_space_properties: Package-level chart properties omitted during load.

    Returns:
        Detached chart snapshots in worksheet drawing order.
    """

    snapshots: list[ChartSnapshot] = []
    for index, source in enumerate(sheet._charts):
        chart = deepcopy(source)
        chart_part = chart_parts[index] if index < len(chart_parts) else None
        properties = chart_space_properties.get(chart_part or "")
        if properties is not None:
            chart.style, chart.roundedCorners = properties
        snapshots.append(
            ChartSnapshot(
                chart=chart,
                chart_type=type(chart).__name__,
                anchor_coordinates=_anchor_coordinates(chart.anchor),
                references=_chart_references(chart),
                has_supported_type=type(chart) in _SUPPORTED_CHART_TYPES,
                has_supported_anchor=type(chart.anchor) in _SUPPORTED_ANCHOR_TYPES,
                is_combined=len(chart._charts) != 1,
                is_pivot=chart.pivotSource is not None,
            )
        )
    return tuple(snapshots)


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


def _cell_presentation(cell: Any) -> CellPresentation:
    """Detach the writer-owned presentation of one worksheet cell.

    Args:
        cell: Openpyxl cell or merged-cell object.

    Returns:
        Immutable presentation snapshot used by the workbook writer.
    """

    return CellPresentation(
        font=copy(cell.font),
        fill=copy(cell.fill),
        border=copy(cell.border),
        alignment=copy(cell.alignment),
        number_format=cell.number_format,
        protection=copy(cell.protection),
        quote_prefix=bool(getattr(cell, "quotePrefix", False)),
        hyperlink=copy(getattr(cell, "hyperlink", None)),
        comment=copy(getattr(cell, "comment", None)),
        is_formula=getattr(cell, "data_type", None) == "f",
    )


def _read_sheet(
    sheet: Worksheet,
    drawing_profiles: dict[str, tuple[bool, tuple[str, ...]]],
    chart_space_properties: dict[str, tuple[Any, Any]],
) -> SheetSnapshot:
    """Detach supported values, presentation, dimensions, and feature flags.

    Args:
        sheet: Openpyxl worksheet to snapshot.
        drawing_profiles: Package drawing-part classifications.
        chart_space_properties: Chart properties omitted by openpyxl's reader.

    Returns:
        Immutable adapter-owned worksheet state.
    """

    charts = _read_charts(
        sheet,
        _sheet_chart_parts(sheet, drawing_profiles),
        chart_space_properties,
    )
    chart_anchor_coordinates = {
        coordinate for chart in charts for coordinate in chart.anchor_coordinates
    }
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
            values[coordinate] = value
            presentations[coordinate] = _cell_presentation(cell)
            if is_formula:
                formula_cells.append(coordinate)
            if presentations[coordinate].hyperlink is not None:
                hyperlink_cells.append(coordinate)
            if presentations[coordinate].comment is not None:
                comment_cells.append(coordinate)

    synthetic_chart_anchor_cells: set[Coordinate] = set()
    for coordinate in chart_anchor_coordinates.difference(values):
        cell = sheet.cell(coordinate.row, coordinate.column)
        values[coordinate] = None
        presentations[coordinate] = _cell_presentation(cell)
        synthetic_chart_anchor_cells.add(coordinate)

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
        charts=charts,
        synthetic_chart_anchor_cells=frozenset(synthetic_chart_anchor_cells),
        has_unsupported_drawings=_sheet_has_unsupported_drawings(sheet, drawing_profiles),
    )


def read_workbook(path: str | Path) -> WorkbookSnapshot:
    """Load a non-macro XLSX workbook and detach supported state.

    Args:
        path: Input ``.xlsx`` path.

    Returns:
        Immutable workbook snapshot containing no live openpyxl objects.
    """

    source_path = Path(path)
    drawing_profiles = _drawing_profiles(source_path)
    chart_space_properties = _chart_space_properties(source_path)
    workbook = load_workbook(source_path, data_only=False, keep_links=False)
    try:
        return WorkbookSnapshot(
            sheets=tuple(
                _read_sheet(sheet, drawing_profiles, chart_space_properties)
                for sheet in workbook.worksheets
            ),
            chartsheets=tuple(sheet.title for sheet in workbook.chartsheets),
            properties=copy(workbook.properties),
            loaded_theme=workbook.loaded_theme,
        )
    finally:
        workbook.close()
