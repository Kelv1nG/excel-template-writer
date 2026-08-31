from datetime import date

import pytest

from excel_template_writer.compiler import compile_sheet
from excel_template_writer.diagnostics import DiagnosticCode
from excel_template_writer.model import Coordinate, Rectangle, WorksheetTemplate
from excel_template_writer.render import RenderPlan, render_sheet
from excel_template_writer.values import TypeAdapter


def _values_by_coordinate(plan: RenderPlan) -> dict[str, object]:
    return {cell.coordinate.a1: cell.value for cell in plan.cells}


def test_vertical_repeat_inserts_rows_and_moves_content_below() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            ["Report", None],
            ["{% for row in rows %}{{ row.name }}", "{{ row.amount }}{% endfor %}"],
            ["Total", "{{ total }}"],
        ],
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(
        compiled,
        {"rows": [{"name": "A", "amount": 10}, {"name": "B", "amount": 20}], "total": 30},
    ).require()

    assert _values_by_coordinate(plan) == {
        "A1": "Report",
        "A2": "A",
        "B2": 10,
        "A3": "B",
        "B3": 20,
        "A4": "Total",
        "B4": 30,
    }


def test_render_sheet_normalizes_caller_supplied_adapter_values() -> None:
    class Rows:
        def __init__(self) -> None:
            self.values = [{"name": "A"}, {"name": "B"}]

    template = WorksheetTemplate.from_rows(
        "Report",
        [["{% for row in rows %}{{ row.name }}{% endfor %}"]],
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(
        compiled,
        {"rows": Rows()},
        adapters=(TypeAdapter(Rows, lambda value: value.values),),
    ).require()

    assert _values_by_coordinate(plan) == {"A1": "A", "A2": "B"}


def test_empty_collection_keeps_one_blank_formatted_instance() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [["{% for row in rows %}{{ row.name }}", "Fixed{% endfor %}"]],
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(compiled, {"rows": []}).require()

    assert _values_by_coordinate(plan) == {"A1": None, "B1": "Fixed"}
    assert plan.height == 1


def test_cell_shift_grows_only_the_repeated_lane() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            ['{% for item in items shift="cells" %}{{ item }}{% endfor %}', "Beside"],
            ["After", "Still here"],
        ],
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(compiled, {"items": ["A", "B"]}).require()

    assert _values_by_coordinate(plan) == {
        "A1": "A",
        "B1": "Beside",
        "A2": "B",
        "B2": "Still here",
        "A3": "After",
    }


def test_cell_shift_region_uses_tallest_side_by_side_lane_and_exact_column_band() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {
            "A1": '{% region shift="cells" %}',
            "A2": '{% for item in left shift="cells" %}{{ item }}',
            "C2": "{% endfor %}",
            "D2": '{% for item in middle shift="cells" %}{{ item }}',
            "F2": "{% endfor %}",
            "G2": '{% for item in right shift="cells" %}{{ item }}',
            "I2": "{% endfor %}",
            "J2": "{% endregion %}",
            "A10": "Moves with region band",
            "K10": "Stays outside region band",
        },
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(
        compiled,
        {
            "left": ["L1", "L2", "L3"],
            "middle": ["M1", "M2", "M3", "M4", "M5"],
            "right": ["R1", "R2"],
        },
    ).require()

    values = _values_by_coordinate(plan)
    assert {coordinate: values[coordinate] for coordinate in ("A2", "A3", "A4")} == {
        "A2": "L1",
        "A3": "L2",
        "A4": "L3",
    }
    assert {coordinate: values[coordinate] for coordinate in ("D2", "D3", "D4", "D5", "D6")} == {
        "D2": "M1",
        "D3": "M2",
        "D4": "M3",
        "D5": "M4",
        "D6": "M5",
    }
    assert values["A14"] == "Moves with region band"
    assert values["K10"] == "Stays outside region band"
    assert "A10" not in values
    assert plan.height == 14


def test_row_shift_region_moves_complete_rows_by_tallest_child_growth() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {
            "A1": "{% region %}",
            "A2": '{% for item in left shift="cells" %}{{ item }}{% endfor %}',
            "D2": '{% for item in right shift="cells" %}{{ item }}{% endfor %}',
            "J2": "{% endregion %}",
            "A10": "Left footer",
            "K10": "Right footer",
        },
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(
        compiled,
        {"left": ["L1", "L2", "L3"], "right": ["R1", "R2", "R3", "R4", "R5"]},
    ).require()

    values = _values_by_coordinate(plan)
    assert values["A14"] == "Left footer"
    assert values["K14"] == "Right footer"
    assert "K10" not in values


def test_region_reserved_height_absorbs_child_growth_before_external_shift() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {
            "A1": '{% region shift="cells" %}',
            "A2": '{% for item in items shift="cells" %}{{ item }}{% endfor %}',
            "J10": "{% endregion %}",
            "A20": "Footer",
        },
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(compiled, {"items": [1, 2, 3, 4, 5]}).require()

    values = _values_by_coordinate(plan)
    assert values["A20"] == "Footer"
    assert "A24" not in values


def test_nested_regions_are_measured_from_the_inside_out() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {
            "A1": '{% region shift="cells" %}',
            "A2": '{% region shift="cells" %}',
            "A3": '{% for item in items shift="cells" %}{{ item }}{% endfor %}',
            "F3": "{% endregion %}",
            "J3": "{% endregion %}",
            "A10": "Outer footer",
            "K10": "Adjacent footer",
        },
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(compiled, {"items": [1, 2, 3, 4, 5]}).require()

    values = _values_by_coordinate(plan)
    assert [values[f"A{row}"] for row in range(3, 8)] == [1, 2, 3, 4, 5]
    assert values["A14"] == "Outer footer"
    assert values["K10"] == "Adjacent footer"


def test_condition_selects_and_compacts_the_matching_branch() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            ["{% if taxable %}Tax", "{{ amount }}{% else %}"],
            ["Not taxable", "{% endif %}"],
            ["Footer", None],
        ],
    )
    compiled = compile_sheet(template).require()

    true_plan = render_sheet(compiled, {"taxable": True, "amount": 25}).require()
    false_plan = render_sheet(compiled, {"taxable": False, "amount": 25}).require()

    assert _values_by_coordinate(true_plan) == {"A1": "Tax", "B1": 25, "A2": "Footer"}
    assert _values_by_coordinate(false_plan) == {"A1": "Not taxable", "A2": "Footer"}


def test_scalar_cell_keeps_excel_compatible_native_value() -> None:
    template = WorksheetTemplate.from_rows("Report", [["{{ issued_on }}"]])
    compiled = compile_sheet(template).require()
    value = date(2026, 8, 13)

    plan = render_sheet(compiled, {"issued_on": value}).require()

    assert plan.cells[0].value is value


def test_date_filter_formats_sole_and_mixed_expressions_as_text() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            ['{{ report_date | date("YYYY-mm") }}'],
            ['For the month ending {{ report_date | date("dd mmmm yyyy") }}'],
        ],
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(compiled, {"report_date": date(2026, 8, 31)}).require()

    assert _values_by_coordinate(plan) == {
        "A1": "2026-08",
        "A2": "For the month ending 31 August 2026",
    }


def test_date_filter_reports_null_as_a_type_mismatch_at_its_cell() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {"D4": '{{ report_date | date("yyyy-mm") }}'},
    )
    compiled = compile_sheet(template).require()

    result = render_sheet(compiled, {"report_date": None})

    assert result.plan is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.FILTER_TYPE_MISMATCH
    ]
    assert str(result.diagnostics[0].location) == "Report!D4:0"


def test_date_filter_preserves_an_empty_repeat_placeholder() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [['{% for row in rows %}{{ row.when | date("yyyy-mm") }}{% endfor %}']],
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(compiled, {"rows": []}).require()

    assert _values_by_coordinate(plan) == {"A1": None}


def test_sum_filter_produces_scalar_totals_without_layout_changes() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [["{{ amounts | sum }}", '{{ rows | sum("amount") }}']],
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(
        compiled,
        {
            "amounts": [1, None, 2.5],
            "rows": [{"amount": 10}, {"amount": None}, {"amount": 5}],
        },
    ).require()

    assert _values_by_coordinate(plan) == {"A1": 3.5, "B1": 15}


def test_sum_filter_reports_missing_columns_and_type_mismatches_at_the_output_cell() -> None:
    missing_template = WorksheetTemplate.from_cells(
        "Report",
        {"D4": '{{ rows | sum("amount") }}'},
    )
    invalid_template = WorksheetTemplate.from_cells(
        "Report",
        {"E5": "{{ values | sum }}"},
    )

    missing_result = render_sheet(
        compile_sheet(missing_template).require(),
        {"rows": [{"amount": 1}, {}]},
    )
    invalid_result = render_sheet(
        compile_sheet(invalid_template).require(),
        {"values": [1, True]},
    )

    assert [diagnostic.code for diagnostic in missing_result.diagnostics] == [
        DiagnosticCode.MISSING_VALUE
    ]
    assert "rows[1].amount" in missing_result.diagnostics[0].message
    assert str(missing_result.diagnostics[0].location) == "Report!D4:0"
    assert [diagnostic.code for diagnostic in invalid_result.diagnostics] == [
        DiagnosticCode.FILTER_TYPE_MISMATCH
    ]
    assert str(invalid_result.diagnostics[0].location) == "Report!E5:0"


def test_aggregate_filters_and_arithmetic_produce_scalar_values_without_layout_changes() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            [
                '{{ rows | min("amount") }}',
                '{{ rows | max("amount") }}',
                "{{ rows | count }}",
                '{{ (rows | max("amount")) - (rows | min("amount")) }}',
                '{{ empty_rows | max("amount") }}',
            ]
        ],
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(
        compiled,
        {
            "rows": [{"amount": 10}, {"amount": None}, {"amount": 25}],
            "empty_rows": [],
        },
    ).require()

    assert _values_by_coordinate(plan) == {
        "A1": 10,
        "B1": 25,
        "C1": 3,
        "D1": 15,
        "E1": None,
    }


def test_numeric_arithmetic_uses_repeat_lexical_scope() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [["{% for line in lines %}{{ line.quantity * line.price }}{% endfor %}"]],
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(
        compiled,
        {
            "lines": [
                {"quantity": 2, "price": 5},
                {"quantity": 3, "price": 4},
            ]
        },
    ).require()

    assert _values_by_coordinate(plan) == {"A1": 10, "A2": 12}


@pytest.mark.parametrize(
    ("expression", "context", "code"),
    [
        (
            "left + right",
            {"left": "1", "right": 2},
            DiagnosticCode.ARITHMETIC_TYPE_MISMATCH,
        ),
        (
            "left / right",
            {"left": 1, "right": 0},
            DiagnosticCode.DIVISION_BY_ZERO,
        ),
        (
            "left * right",
            {"left": 1e308, "right": 10.0},
            DiagnosticCode.NON_FINITE_EXPRESSION_NUMBER,
        ),
    ],
)
def test_arithmetic_failures_report_stable_diagnostics_at_the_output_cell(
    expression: str,
    context: dict[str, object],
    code: DiagnosticCode,
) -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {"F6": f"{{{{ {expression} }}}}"},
    )
    compiled = compile_sheet(template).require()

    result = render_sheet(compiled, context)

    assert result.plan is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [code]
    assert str(result.diagnostics[0].location) == "Report!F6:0"


def test_nested_repeats_are_evaluated_from_the_ast_scope_tree() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            ["{% for group in groups %}{{ group.name }}", None],
            [
                '{% for item in group.items shift="cells" %}{{ item }}{% endfor %}',
                None,
            ],
            ["--", "{% endfor %}"],
        ],
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(
        compiled,
        {
            "groups": [
                {"name": "First", "items": ["A", "B"]},
                {"name": "Second", "items": ["C"]},
            ]
        },
    ).require()

    assert _values_by_coordinate(plan) == {
        "A1": "First",
        "A2": "A",
        "A3": "B",
        "A4": "--",
        "A5": "Second",
        "A6": "C",
        "A7": "--",
    }


def test_condition_inherits_cell_shift_isolation_from_its_container() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            ['{% for record in records shift="cells" %}{{ record.name }}', None],
            ["{% if record.show %}Shown{% endif %}", "Keeps row"],
            ["Tail", "{% endfor %}"],
        ],
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(
        compiled,
        {"records": [{"name": "Entry", "show": False}]},
    ).require()

    assert _values_by_coordinate(plan) == {
        "A1": "Entry",
        "A2": "Tail",
        "B2": "Keeps row",
    }


def test_render_plan_carries_explicit_row_and_merge_provenance() -> None:
    template = WorksheetTemplate(
        "Report",
        {
            Coordinate(1, 1): "{% for card in cards %}{{ card.title }}",
            Coordinate(1, 2): None,
            Coordinate(2, 1): "Owner",
            Coordinate(2, 2): "{{ card.owner }}{% endfor %}",
        },
        (Rectangle(1, 1, 1, 2),),
    )
    compiled = compile_sheet(template).require()

    plan = render_sheet(
        compiled,
        {"cards": [{"title": "First", "owner": "Mina"}, {"title": "Second", "owner": "Jo"}]},
    ).require()

    assert [(row.destination_row, row.source_row) for row in plan.rows] == [
        (1, 1),
        (2, 2),
        (3, 1),
        (4, 2),
    ]
    assert [merge.rectangle for merge in plan.merges] == [
        Rectangle(1, 1, 1, 2),
        Rectangle(3, 1, 3, 2),
    ]
    assert all(merge.source_rectangle == Rectangle(1, 1, 1, 2) for merge in plan.merges)


def test_render_rejects_noncanonical_context_before_evaluation() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [["{% for item in items %}{{ item }}{% endfor %}"]],
    )
    compiled = compile_sheet(template).require()

    result = render_sheet(compiled, {"items": {"A", "B"}})

    assert result.plan is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.UNORDERED_CONTEXT_COLLECTION
    ]
    assert str(result.diagnostics[0].location) == "context.items"


def test_repeat_requires_an_ordered_collection_including_for_null() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [["{% for item in items %}{{ item }}{% endfor %}"]],
    )
    compiled = compile_sheet(template).require()

    null_result = render_sheet(compiled, {"items": None})
    record_result = render_sheet(compiled, {"items": {"first": "A"}})
    tuple_result = render_sheet(compiled, {"items": ("A", "B")}).require()

    assert [diagnostic.code for diagnostic in null_result.diagnostics] == [
        DiagnosticCode.EXPECTED_COLLECTION
    ]
    assert [diagnostic.code for diagnostic in record_result.diagnostics] == [
        DiagnosticCode.EXPECTED_COLLECTION
    ]
    assert _values_by_coordinate(tuple_result) == {"A1": "A", "A2": "B"}


def test_records_and_ordered_collections_are_not_scalar_cell_values() -> None:
    template = WorksheetTemplate.from_rows("Report", [["{{ value }}"]])
    compiled = compile_sheet(template).require()

    record_result = render_sheet(compiled, {"value": {"name": "A"}})
    collection_result = render_sheet(compiled, {"value": ["A", "B"]})

    assert [diagnostic.code for diagnostic in record_result.diagnostics] == [
        DiagnosticCode.COLLECTION_IN_SCALAR_CELL
    ]
    assert [diagnostic.code for diagnostic in collection_result.diagnostics] == [
        DiagnosticCode.COLLECTION_IN_SCALAR_CELL
    ]
