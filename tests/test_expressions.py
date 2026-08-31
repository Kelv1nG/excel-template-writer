from datetime import date, datetime
from decimal import Decimal

import pytest

from excel_template_writer.expressions import (
    ArithmeticTypeError,
    DateFormatExpression,
    DivisionByZeroError,
    ExpressionEvaluationError,
    ExpressionSyntaxError,
    FilterTypeError,
    FilterValidationError,
    MissingValueError,
    NonFiniteExpressionNumberError,
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
    ("source", "expected"),
    [
        ("2 + 3 * 4", 14),
        ("(2 + 3) * 4", 20),
        ("20 / 5 - 1", 3.0),
        ("-2 * +3", -6),
        ("2 + 3 * 4 == 14 and not false", True),
    ],
)
def test_numeric_arithmetic_uses_standard_precedence(source: str, expected: object) -> None:
    expression = compile_expression(source)

    assert evaluate_expression(expression, {}) == expected


def test_filtered_arithmetic_operands_are_explicitly_parenthesized() -> None:
    expression = compile_expression(
        '(rows | max("high")) - (rows | min("low"))',
    )
    rows = [
        {"high": 12, "low": 4},
        {"high": 18, "low": 7},
    ]

    assert evaluate_expression(expression, {"rows": rows}) == 14

    with pytest.raises(ExpressionSyntaxError, match="unexpected token"):
        compile_expression('rows | max("high") - 1')


@pytest.mark.parametrize(
    ("source", "scope", "expected"),
    [
        ("left + right", {"left": 2, "right": 3}, 5),
        ("left - right", {"left": 2, "right": 3.5}, -1.5),
        ("left * right", {"left": Decimal("2.5"), "right": 3}, Decimal("7.5")),
        ("left / right", {"left": 5, "right": 2}, 2.5),
        (
            "left / right",
            {"left": Decimal("5"), "right": 2},
            Decimal("2.5"),
        ),
        ("left + right", {"left": None, "right": 3}, None),
        ("-value", {"value": None}, None),
    ],
)
def test_numeric_arithmetic_preserves_types_and_propagates_null(
    source: str,
    scope: dict[str, object],
    expected: object,
) -> None:
    result = evaluate_expression(compile_expression(source), scope)

    assert result == expected
    assert type(result) is type(expected)


@pytest.mark.parametrize(
    ("source", "scope", "message"),
    [
        ("left + right", {"left": True, "right": 1}, "left operand is bool"),
        ("left * right", {"left": "2", "right": 3}, "left operand is str"),
        (
            "left - right",
            {"left": Decimal("2"), "right": 1.5},
            "cannot mix floating-point and decimal",
        ),
        ("-value", {"value": "2"}, "unary operand is str"),
    ],
)
def test_numeric_arithmetic_rejects_non_numeric_and_mixed_operands(
    source: str,
    scope: dict[str, object],
    message: str,
) -> None:
    with pytest.raises(ArithmeticTypeError, match=message):
        evaluate_expression(compile_expression(source), scope)


def test_numeric_arithmetic_rejects_zero_division_and_non_finite_results() -> None:
    with pytest.raises(DivisionByZeroError, match="division by zero"):
        evaluate_expression(compile_expression("left / right"), {"left": 1, "right": 0})

    with pytest.raises(NonFiniteExpressionNumberError, match="non-finite"):
        evaluate_expression(
            compile_expression("left * right"),
            {"left": 1e308, "right": 10.0},
        )


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
    ("source", "scope", "expected"),
    [
        ("values | min", {"values": []}, None),
        ("values | max", {"values": [None, None]}, None),
        ("values | min", {"values": [3, None, 1.5]}, 1.5),
        ("values | max", {"values": [Decimal("3.5"), 4]}, Decimal("4")),
        (
            'rows | min("amount")',
            {"rows": [{"amount": 10}, {"amount": None}, {"amount": 5}]},
            5,
        ),
        (
            'rows | max("amount")',
            {"rows": [{"amount": 10}, {"amount": None}, {"amount": 5}]},
            10,
        ),
        ("values | count", {"values": [1, None, "x"]}, 3),
        (
            'rows | count("value")',
            {"rows": [{"value": 1}, {"value": None}, {"value": "x"}]},
            2,
        ),
    ],
)
def test_collection_aggregate_filters_have_explicit_empty_and_null_semantics(
    source: str,
    scope: dict[str, object],
    expected: object,
) -> None:
    result = evaluate_expression(compile_expression(source), scope)

    assert result == expected
    assert type(result) is type(expected)


@pytest.mark.parametrize("filter_name", ["min", "max", "count"])
def test_column_aggregates_report_missing_values_with_item_paths(filter_name: str) -> None:
    expression = compile_expression(f'rows | {filter_name}("amount")')

    with pytest.raises(MissingValueError, match=r"rows\[1\]\.amount"):
        evaluate_expression(expression, {"rows": [{"amount": 10}, {}]})


@pytest.mark.parametrize(
    ("source", "scope", "message"),
    [
        ("value | min", {"value": "123"}, "requires an ordered collection"),
        ("values | max", {"values": [1, True]}, r"values\[1\] is bool"),
        (
            'rows | min("amount")',
            {"rows": [{"amount": "1"}]},
            "requires numeric values",
        ),
        (
            "values | max",
            {"values": [1.0, Decimal("2")]},
            "cannot mix floating-point and decimal values",
        ),
        (
            'rows | count("amount")',
            {"rows": [{"amount": 1}, "not a record"]},
            "requires record items",
        ),
    ],
)
def test_collection_aggregate_filters_reject_invalid_runtime_inputs(
    source: str,
    scope: dict[str, object],
    message: str,
) -> None:
    with pytest.raises(FilterTypeError, match=message):
        evaluate_expression(compile_expression(source), scope)


def test_sum_filter_rejects_a_non_finite_derived_total() -> None:
    expression = compile_expression("values | sum")

    with pytest.raises(NonFiniteExpressionNumberError, match="non-finite"):
        evaluate_expression(expression, {"values": [1e308, 1e308]})


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
        ('values | min("amount", "tax")', "expects 0 or 1 argument"),
        ("values | max(1)", "column must be a string literal"),
        ('values | count("_private")', "must not begin with '_'"),
    ],
)
def test_expression_compilation_rejects_invalid_filter_contracts(
    source: str,
    message: str,
) -> None:
    with pytest.raises(FilterValidationError, match=message):
        compile_expression(source)
