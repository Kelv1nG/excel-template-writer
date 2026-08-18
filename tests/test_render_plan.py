from datetime import date

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
