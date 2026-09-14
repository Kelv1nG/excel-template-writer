import pytest

from excel_template_writer import MissingOutputPolicy
from excel_template_writer.compiler import compile_sheet
from excel_template_writer.diagnostics import (
    Diagnostic,
    DiagnosticCode,
    DiagnosticSeverity,
    SourceLocation,
    TemplateRenderError,
)
from excel_template_writer.model import WorksheetTemplate
from excel_template_writer.render import RenderPlan, render_sheet


def _values_by_coordinate(plan: RenderPlan) -> dict[str, object]:
    return {cell.coordinate.a1: cell.value for cell in plan.cells}


def test_missing_output_policy_is_a_public_string_enum() -> None:
    assert [policy.value for policy in MissingOutputPolicy] == [
        "error",
        "blank",
        "preserve",
    ]
    assert MissingOutputPolicy("blank") is MissingOutputPolicy.BLANK


def test_diagnostic_severity_defaults_to_error() -> None:
    diagnostic = Diagnostic(
        DiagnosticCode.MISSING_VALUE,
        "missing value: absent",
        SourceLocation("Report", "A1"),
    )

    assert diagnostic.severity is DiagnosticSeverity.ERROR


@pytest.mark.parametrize("missing_output", [MissingOutputPolicy.ERROR, "error"])
def test_missing_output_is_fatal_by_default_and_in_error_mode(
    missing_output: MissingOutputPolicy | str,
) -> None:
    template = WorksheetTemplate.from_rows("Report", [["{{ absent }}"]])
    compiled = compile_sheet(template).require()

    default_result = render_sheet(compiled, {})
    explicit_result = render_sheet(compiled, {}, missing_output=missing_output)

    for result in (default_result, explicit_result):
        assert result.plan is None
        assert [diagnostic.code for diagnostic in result.diagnostics] == [
            DiagnosticCode.MISSING_VALUE
        ]
        assert result.diagnostics[0].severity is DiagnosticSeverity.ERROR
        with pytest.raises(TemplateRenderError):
            result.require()


@pytest.mark.parametrize(
    ("missing_output", "expected"),
    [
        (MissingOutputPolicy.BLANK, None),
        ("blank", None),
        (MissingOutputPolicy.PRESERVE, "{{   customer.phone\t }}"),
        ("preserve", "{{   customer.phone\t }}"),
    ],
)
def test_tolerant_modes_render_a_sole_missing_output_with_a_warning(
    missing_output: MissingOutputPolicy | str,
    expected: object,
) -> None:
    source = "{{   customer.phone\t }}"
    template = WorksheetTemplate.from_rows("Report", [[source]])
    compiled = compile_sheet(template).require()

    result = render_sheet(
        compiled,
        {"customer": {}},
        missing_output=missing_output,
    )

    assert _values_by_coordinate(result.require()) == {"A1": expected}
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.MISSING_VALUE_RENDERED
    ]
    assert result.diagnostics[0].severity is DiagnosticSeverity.WARNING
    assert result.diagnostics[0].message == "missing value: customer.phone"
    assert str(result.diagnostics[0].location) == "Report!A1:0"


@pytest.mark.parametrize(
    ("missing_output", "expected"),
    [
        (MissingOutputPolicy.BLANK, "Phone: ; email: known@example.test"),
        (
            MissingOutputPolicy.PRESERVE,
            "Phone: {{ customer.phone }}; email: known@example.test",
        ),
    ],
)
def test_tolerant_modes_replace_only_the_missing_segment_in_mixed_text(
    missing_output: MissingOutputPolicy,
    expected: str,
) -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [["Phone: {{ customer.phone }}; email: {{ customer.email }}"]],
    )
    compiled = compile_sheet(template).require()

    result = render_sheet(
        compiled,
        {"customer": {"email": "known@example.test"}},
        missing_output=missing_output,
    )

    assert _values_by_coordinate(result.require()) == {"A1": expected}
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.MISSING_VALUE_RENDERED
    ]


def test_preserve_keeps_the_complete_compound_output_tag() -> None:
    source = "{{   missing.amount + 1   }}"
    template = WorksheetTemplate.from_rows("Report", [[source]])
    compiled = compile_sheet(template).require()

    result = render_sheet(
        compiled,
        {},
        missing_output=MissingOutputPolicy.PRESERVE,
    )

    assert _values_by_coordinate(result.require()) == {"A1": source}
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.MISSING_VALUE_RENDERED
    ]


def test_preserve_warns_for_each_missing_nested_value_in_a_repeat() -> None:
    output_tag = "{{ row.customer.phone }}"
    template = WorksheetTemplate.from_rows(
        "Report",
        [[f"{{% for row in rows %}}{output_tag}{{% endfor %}}"]],
    )
    compiled = compile_sheet(template).require()

    result = render_sheet(
        compiled,
        {
            "rows": [
                {"customer": {"phone": "123"}},
                {"customer": {}},
                {"customer": {}},
            ]
        },
        missing_output=MissingOutputPolicy.PRESERVE,
    )

    assert _values_by_coordinate(result.require()) == {
        "A1": "123",
        "A2": output_tag,
        "A3": output_tag,
    }
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.MISSING_VALUE_RENDERED,
        DiagnosticCode.MISSING_VALUE_RENDERED,
    ]
    assert all(
        diagnostic.severity is DiagnosticSeverity.WARNING for diagnostic in result.diagnostics
    )


@pytest.mark.parametrize(
    "missing_output",
    [MissingOutputPolicy.BLANK, MissingOutputPolicy.PRESERVE],
)
def test_default_filter_and_present_null_do_not_warn(
    missing_output: MissingOutputPolicy,
) -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [['{{ absent | default("fallback") }}', "{{ present }}"]],
    )
    compiled = compile_sheet(template).require()

    result = render_sheet(
        compiled,
        {"present": None},
        missing_output=missing_output,
    )

    assert _values_by_coordinate(result.require()) == {
        "A1": "fallback",
        "B1": None,
    }
    assert result.diagnostics == ()


@pytest.mark.parametrize(
    "missing_output",
    [MissingOutputPolicy.BLANK, MissingOutputPolicy.PRESERVE],
)
def test_empty_repeat_item_suppression_remains_warning_free(
    missing_output: MissingOutputPolicy,
) -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [["{% for row in rows %}{{ row.name }}{% endfor %}"]],
    )
    compiled = compile_sheet(template).require()

    result = render_sheet(
        compiled,
        {"rows": []},
        missing_output=missing_output,
    )

    assert _values_by_coordinate(result.require()) == {"A1": None}
    assert result.diagnostics == ()


def test_tolerated_missing_output_does_not_hide_an_independent_fatal_error() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [["{{ missing }}", "{{ numerator / denominator }}"]],
    )
    compiled = compile_sheet(template).require()

    result = render_sheet(
        compiled,
        {"numerator": 1, "denominator": 0},
        missing_output=MissingOutputPolicy.BLANK,
    )

    assert result.plan is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.MISSING_VALUE_RENDERED,
        DiagnosticCode.DIVISION_BY_ZERO,
    ]
    assert [diagnostic.severity for diagnostic in result.diagnostics] == [
        DiagnosticSeverity.WARNING,
        DiagnosticSeverity.ERROR,
    ]


@pytest.mark.parametrize(
    "source",
    [
        "{% for item in missing_items %}{{ item }}{% endfor %}",
        "{% if missing_condition %}Shown{% endif %}",
    ],
)
@pytest.mark.parametrize(
    "missing_output",
    [MissingOutputPolicy.BLANK, MissingOutputPolicy.PRESERVE],
)
def test_missing_structural_expression_remains_fatal(
    source: str,
    missing_output: MissingOutputPolicy,
) -> None:
    template = WorksheetTemplate.from_rows("Report", [[source]])
    compiled = compile_sheet(template).require()

    result = render_sheet(compiled, {}, missing_output=missing_output)

    assert result.plan is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [DiagnosticCode.MISSING_VALUE]
    assert result.diagnostics[0].severity is DiagnosticSeverity.ERROR


def test_rejects_an_unknown_missing_output_policy() -> None:
    template = WorksheetTemplate.from_rows("Report", [["Static"]])
    compiled = compile_sheet(template).require()

    with pytest.raises(ValueError, match="missing_output"):
        render_sheet(compiled, {}, missing_output="silent")
