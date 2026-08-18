"""Show the canonical render-context contract without requiring an XLSX file.

Run from the repository root with:

    uv run python scratch/value_model_example.py
"""

from __future__ import annotations

from datetime import date
from decimal import Decimal

from excel_template_writer import (
    TypeAdapter,
    WorksheetTemplate,
    compile_sheet,
    normalize_context,
    render_sheet,
    validate_context,
)

CONTEXT = {
    "report": {
        "title": "August revenue",
        "issued_on": date(2026, 8, 19),
        "total": Decimal("2675.50"),
    },
    "regions": ["North", "South", "Central"],
    "lines": [
        {"description": "Consulting", "quantity": 2, "amount": Decimal("1500.00")},
        {"description": "Support", "quantity": 3, "amount": Decimal("1175.50")},
    ],
}


def _show_render(name: str, template: WorksheetTemplate) -> None:
    plan = render_sheet(compile_sheet(template).require(), CONTEXT).require()
    values = {cell.coordinate.a1: cell.value for cell in plan.cells}
    print(f"{name}: {values}")


def _show_valid_examples() -> None:
    assert validate_context(CONTEXT) == ()

    _show_render(
        "single values and record access",
        WorksheetTemplate.from_rows(
            "Scalars",
            [["{{ report.title }}", "{{ report.issued_on }}", "{{ report.total }}"]],
        ),
    )
    _show_render(
        "list of scalars",
        WorksheetTemplate.from_rows(
            "List",
            [["{% for region in regions %}{{ region }}{% endfor %}"]],
        ),
    )
    _show_render(
        "table-shaped list of records",
        WorksheetTemplate.from_rows(
            "Rows",
            [
                [
                    "{% for line in lines %}{{ line.description }}",
                    "{{ line.quantity }}",
                    "{{ line.amount }}{% endfor %}",
                ]
            ],
        ),
    )


class DataFrameLike:
    """Small stand-in for a pandas, Polars, Arrow, or DuckDB result."""

    def __init__(self, rows: list[dict[str, object]]) -> None:
        self.rows = rows


def _show_immutable_normalization() -> None:
    raw = {"report": {"title": "Original"}, "regions": ["North", "South"]}
    normalized = normalize_context(raw).require()
    raw["report"]["title"] = "Changed after normalization"
    raw["regions"].append("Central")
    print(f"immutable snapshot: {normalized}")


def _show_caller_supplied_adapter() -> None:
    frame = DataFrameLike(
        [
            {"description": "Implementation", "amount": Decimal("900.00")},
            {"description": "Support", "amount": Decimal("100.00")},
        ]
    )
    adapter = TypeAdapter(DataFrameLike, lambda value: value.rows, name="dataframe-like")
    template = WorksheetTemplate.from_rows(
        "Adapted rows",
        [
            [
                "{% for line in lines %}{{ line.description }}",
                "{{ line.amount }}{% endfor %}",
            ]
        ],
    )
    plan = render_sheet(
        compile_sheet(template).require(),
        {"lines": frame},
        adapters=(adapter,),
    ).require()
    values = {cell.coordinate.a1: cell.value for cell in plan.cells}
    print(f"caller-supplied adapter: {values}")


def _show_rejections() -> None:
    invalid_context = {
        "unordered": {"North", "South"},
        "nonfinite": float("inf"),
        "unadapted_table": DataFrameLike([]),
    }
    print("rejected values:")
    for diagnostic in validate_context(invalid_context):
        print(f"  {diagnostic}")


def main() -> None:
    print("No TypedValue wrappers are used; runtime categories validate permitted operations.")
    _show_valid_examples()
    _show_immutable_normalization()
    _show_caller_supplied_adapter()
    _show_rejections()


if __name__ == "__main__":
    main()
