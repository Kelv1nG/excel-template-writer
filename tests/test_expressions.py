from datetime import date, datetime
from decimal import Decimal

import pytest

from excel_template_writer.expressions import (
    DateFormatExpression,
    ExpressionEvaluationError,
    ExpressionSyntaxError,
    FilterTypeError,
    FilterValidationError,
    MissingValueError,
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
    ("source", "scope", "expected"),
    [
        ("values | sum", {"values": []}, 0),
        ("values | sum", {"values": [None, None]}, 0),
        ("values | sum", {"values": [None, 1, 2]}, 3),
        ("values | sum", {"values": [1, 2.5, None]}, 3.5),
        (
            "values | sum",
            {"values": [Decimal("1.25"), 2, None]},
            Decimal("3.25"),
        ),
        (
            'rows | sum("amount")',
            {"rows": [{"amount": 10}, {"amount": None}, {"amount": 5}]},
            15,
        ),
    ],
)
def test_sum_filter_reduces_numeric_collections_and_record_columns(
    source: str,
    scope: dict[str, object],
    expected: int | float | Decimal,
) -> None:
    expression = compile_expression(source)

    result = evaluate_expression(expression, scope)

    assert result == expected
    assert type(result) is type(expected)


def test_sum_filter_reports_a_missing_record_column_with_its_item_path() -> None:
    expression = compile_expression('rows | sum("amount")')

    with pytest.raises(MissingValueError, match=r"rows\[1\]\.amount") as captured:
        evaluate_expression(expression, {"rows": [{"amount": 10}, {}]})

    assert captured.value.root == "rows"
    assert captured.value.path == "rows[1].amount"


@pytest.mark.parametrize(
    ("source", "scope", "message"),
    [
        ("value | sum", {"value": "123"}, "requires an ordered collection"),
        ("values | sum", {"values": [1, True]}, r"values\[1\] is bool"),
        (
            'rows | sum("amount")',
            {"rows": [{"amount": 1}, "not a record"]},
            "requires record items",
        ),
        (
            'rows | sum("amount")',
            {"rows": [{"amount": "1"}]},
            "requires numeric values",
        ),
        (
            "values | sum",
            {"values": [Decimal("1.0"), 2.0]},
            "cannot mix floating-point and decimal values",
        ),
    ],
)
def test_sum_filter_rejects_invalid_runtime_inputs(
    source: str,
    scope: dict[str, object],
    message: str,
) -> None:
    expression = compile_expression(source)

    with pytest.raises(FilterTypeError, match=message):
        evaluate_expression(expression, scope)


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
        ('values | sum("amount", "tax")', "expects 0 or 1 argument"),
        ("values | sum(column_name)", "arguments must be literals"),
        ("values | sum(1)", "column must be a string literal"),
        ('values | sum("_private")', "must not begin with '_'"),
    ],
)
def test_expression_compilation_rejects_invalid_filter_contracts(
    source: str,
    message: str,
) -> None:
    with pytest.raises(FilterValidationError, match=message):
        compile_expression(source)
