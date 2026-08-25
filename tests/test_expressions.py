from datetime import date, datetime

import pytest

from excel_template_writer.expressions import (
    DateFormatExpression,
    ExpressionEvaluationError,
    ExpressionSyntaxError,
    FilterTypeError,
    FilterValidationError,
    compile_expression,
    evaluate_expression,
    parse_expression,
)


def test_expression_parser_supports_access_boolean_logic_and_comparisons() -> None:
    expression = parse_expression("invoice.customer.active and invoice.total >= 100")

    assert (
        evaluate_expression(
            expression,
            {"invoice": {"customer": {"active": True}, "total": 125}},
        )
        is True
    )


def test_single_expression_preserves_native_types() -> None:
    expression = parse_expression("invoice.issued_on")
    issued_on = date(2026, 8, 13)

    assert evaluate_expression(expression, {"invoice": {"issued_on": issued_on}}) is issued_on


def test_default_filter_handles_missing_values_explicitly() -> None:
    expression = parse_expression('invoice.reference | default("-")')

    assert evaluate_expression(expression, {"invoice": {}}) == "-"


def test_access_to_private_mapping_members_is_forbidden() -> None:
    expression = parse_expression("invoice._internal")

    with pytest.raises(ExpressionEvaluationError, match="private"):
        evaluate_expression(expression, {"invoice": {"_internal": "secret"}})


def test_function_calls_are_not_part_of_the_expression_language() -> None:
    with pytest.raises(ExpressionSyntaxError):
        parse_expression("invoice.calculate()")


@pytest.mark.parametrize(
    ("value", "expected"),
    [
        (date(2026, 8, 31), "31 August 2026"),
        (datetime(2026, 8, 31, 23, 45), "31 August 2026"),
    ],
)
def test_date_filter_compiles_its_literal_format_and_returns_text(
    value: date | datetime,
    expected: str,
) -> None:
    expression = compile_expression('report_date | date("dd mmmm yyyy")')

    assert isinstance(expression, DateFormatExpression)
    assert evaluate_expression(expression, {"report_date": value}) == expected


@pytest.mark.parametrize("value", [None, "2026-08-31", 46265])
def test_date_filter_rejects_non_temporal_values(value: object) -> None:
    expression = compile_expression('report_date | date("yyyy-mm")')

    with pytest.raises(FilterTypeError, match="requires a date or datetime"):
        evaluate_expression(expression, {"report_date": value})


@pytest.mark.parametrize(
    ("source", "message"),
    [
        ("value | unknown", "unknown filter"),
        ("value | date", "expects 1 argument"),
        ('value | date("yyyy", "mm")', "expects 1 argument"),
        ("value | date(format_name)", "arguments must be literals"),
        ("value | date(2026)", "format must be a string literal"),
        ("value | upper(1)", "expects 0 argument"),
        ("values | join(1)", "separator must be a string literal"),
    ],
)
def test_expression_compilation_rejects_invalid_filter_contracts(
    source: str,
    message: str,
) -> None:
    with pytest.raises(FilterValidationError, match=message):
        compile_expression(source)
