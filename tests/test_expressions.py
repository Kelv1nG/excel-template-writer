from datetime import date

import pytest

from excel_template_writer.expressions import (
    ExpressionEvaluationError,
    ExpressionSyntaxError,
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
