from __future__ import annotations

from dataclasses import replace

import pytest

from excel_template_writer.compiler import compile_sheet
from excel_template_writer.diagnostics import DiagnosticCode
from excel_template_writer.limits import (
    XLSX_MAX_CELL_TEXT_LENGTH,
    XLSX_MAX_COLUMNS,
    XLSX_MAX_ROWS,
    ResourceLimits,
)
from excel_template_writer.model import WorksheetTemplate
from excel_template_writer.render import render_sheet
from excel_template_writer.values import normalize_context


def test_default_limits_match_the_public_small_workbook_policy() -> None:
    limits = ResourceLimits()

    assert (
        limits.max_context_depth,
        limits.max_context_nodes,
        limits.max_container_items,
        limits.max_input_string_length,
    ) == (64, 1_000_000, 100_000, 1_000_000)
    assert (
        limits.max_repeat_iterations_per_sheet,
        limits.max_planned_cells_per_sheet,
        limits.max_planned_cells_per_workbook,
    ) == (100_000, 500_000, 1_000_000)
    assert (
        limits.max_output_rows_per_sheet,
        limits.max_output_columns_per_sheet,
        limits.max_worksheets,
    ) == (250_000, 4_096, 100)
    assert (
        limits.max_xlsx_file_bytes,
        limits.max_xlsx_uncompressed_bytes,
        limits.max_xlsx_archive_members,
    ) == (50 * 1024 * 1024, 250 * 1024 * 1024, 10_000)
    assert (XLSX_MAX_ROWS, XLSX_MAX_COLUMNS, XLSX_MAX_CELL_TEXT_LENGTH) == (
        1_048_576,
        16_384,
        32_767,
    )


def test_resource_limits_require_positive_integer_values() -> None:
    with pytest.raises(ValueError, match="max_context_depth"):
        ResourceLimits(max_context_depth=0)
    with pytest.raises(ValueError, match="max_context_nodes"):
        ResourceLimits(max_context_nodes=True)


def test_normalization_records_statistics_and_reuses_the_snapshot() -> None:
    normalized = normalize_context({"rows": [{"name": "A"}]}).require()

    assert normalized.statistics.nodes == 4
    assert normalized.statistics.maximum_depth == 3
    assert normalized.statistics.maximum_container_items == 1
    assert normalized.statistics.maximum_string_length == 1
    assert normalize_context(normalized).require() is normalized

    rejected = normalize_context(
        normalized,
        limits=ResourceLimits(max_context_depth=2),
    )
    assert rejected.context is None
    assert rejected.diagnostics[0].code is DiagnosticCode.CONTEXT_RESOURCE_LIMIT_EXCEEDED


@pytest.mark.parametrize(
    ("context", "limits", "path"),
    [
        ({"rows": [1, 2]}, ResourceLimits(max_container_items=1), "context.rows"),
        ({"value": "long"}, ResourceLimits(max_input_string_length=3), "context.value"),
        (
            {"record": {"nested": {"value": 1}}},
            ResourceLimits(max_context_depth=2),
            "context.record.nested.value",
        ),
        ({"first": 1, "second": 2}, ResourceLimits(max_context_nodes=2), "context.second"),
    ],
)
def test_normalization_limits_fail_at_the_first_breaching_path(
    context: object,
    limits: ResourceLimits,
    path: str,
) -> None:
    result = normalize_context(context, limits=limits)

    assert result.context is None
    assert len(result.diagnostics) == 1
    assert result.diagnostics[0].code is DiagnosticCode.CONTEXT_RESOURCE_LIMIT_EXCEEDED
    assert str(result.diagnostics[0].location) == path


def test_repeat_and_planned_cell_limits_fail_before_returning_a_plan() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [["{% for item in items %}{{ item }}", "Static{% endfor %}"]],
    )
    compiled = compile_sheet(template).require()

    repeat_result = render_sheet(
        compiled,
        {"items": [1, 2]},
        limits=ResourceLimits(max_repeat_iterations_per_sheet=1),
    )
    cell_result = render_sheet(
        compiled,
        {"items": [1, 2]},
        limits=ResourceLimits(max_planned_cells_per_sheet=3),
    )

    assert repeat_result.plan is None
    assert repeat_result.diagnostics[0].code is DiagnosticCode.RENDER_RESOURCE_LIMIT_EXCEEDED
    assert cell_result.plan is None
    assert cell_result.diagnostics[0].code is DiagnosticCode.RENDER_RESOURCE_LIMIT_EXCEEDED


def test_output_dimensions_and_excel_text_have_independent_hard_limits() -> None:
    oversized_grid = WorksheetTemplate.from_cells(
        "Report",
        {f"A{XLSX_MAX_ROWS + 1}": "outside Excel"},
    )
    grid_result = render_sheet(
        compile_sheet(oversized_grid).require(),
        {},
        limits=ResourceLimits(max_output_rows_per_sheet=XLSX_MAX_ROWS + 1),
    )
    text_template = WorksheetTemplate.from_rows("Report", [["{{ value }}"]])
    text_result = render_sheet(
        compile_sheet(text_template).require(),
        {"value": "x" * (XLSX_MAX_CELL_TEXT_LENGTH + 1)},
    )

    assert grid_result.plan is None
    assert grid_result.diagnostics[0].code is DiagnosticCode.XLSX_GRID_LIMIT_EXCEEDED
    assert text_result.plan is None
    assert text_result.diagnostics[0].code is DiagnosticCode.CELL_TEXT_LIMIT_EXCEEDED


def test_soft_output_dimension_limit_is_configurable() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [["{% for item in items %}{{ item }}{% endfor %}"]],
    )
    result = render_sheet(
        compile_sheet(template).require(),
        {"items": [1, 2]},
        limits=replace(ResourceLimits(), max_output_rows_per_sheet=1),
    )

    assert result.plan is None
    assert result.diagnostics[0].code is DiagnosticCode.RENDER_RESOURCE_LIMIT_EXCEEDED
