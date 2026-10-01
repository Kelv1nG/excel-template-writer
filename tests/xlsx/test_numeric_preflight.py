import math
import sys
from dataclasses import replace
from datetime import date, datetime, time
from decimal import Decimal, localcontext

import pytest
from hypothesis import given, settings
from hypothesis import strategies as st

from excel_template_writer.diagnostics import DiagnosticCode, SourceLocation, TemplateRenderError
from excel_template_writer.model import Coordinate, Rectangle
from excel_template_writer.render import PlannedCell, PlannedMerge, PlannedRow, RenderPlan
from excel_template_writer.xlsx.numeric import preflight_numeric_plans


def _plan(*values: object, sheet: str = "Report") -> RenderPlan:
    return RenderPlan(
        sheet=sheet,
        cells=tuple(
            PlannedCell(Coordinate(index, 1), value, Coordinate(1, 1), (index,))
            for index, value in enumerate(values, start=1)
        ),
        rows=(PlannedRow(1, 1),),
        merges=(),
        height=len(values),
        width=1,
    )


def _assert_accepted(value: Decimal) -> None:
    source_tuple = value.as_tuple()
    plan = _plan(value)
    result = preflight_numeric_plans((plan,))
    adapted = result.require()[0]
    assert result.diagnostics == ()
    assert type(adapted.cells[0].value) is float
    assert Decimal(repr(adapted.cells[0].value)) == value
    assert plan.cells[0].value is value
    assert plan.cells[0].value.as_tuple() == source_tuple


def test_numeric_preflight_converts_supported_decimals_in_adapter_copies_only() -> None:
    source = Decimal("12.5000")
    plan = _plan(source)

    adapted = preflight_numeric_plans((plan,)).require()[0]

    assert adapted is not plan
    assert adapted.cells[0] is not plan.cells[0]
    assert type(adapted.cells[0].value) is float
    assert Decimal(repr(adapted.cells[0].value)) == source
    assert plan.cells[0].value is source
    assert source.as_tuple() == Decimal("12.5000").as_tuple()


def test_numeric_preflight_preserves_all_planned_cell_provenance() -> None:
    cell = PlannedCell(
        Coordinate(4, 3), Decimal("0.1"), Coordinate(2, 1), (2, 3), Coordinate(4, 2), True
    )
    merge = PlannedMerge(Rectangle(4, 3, 4, 4), Rectangle(2, 1, 2, 2), (2, 3))
    plan = RenderPlan("Report", (cell,), (PlannedRow(4, 2, (2, 3)),), (merge,), 4, 4)

    adapted = preflight_numeric_plans((plan,)).require()[0]

    assert adapted == replace(plan, cells=(replace(cell, value=0.1),))
    assert adapted.rows is plan.rows
    assert adapted.merges is plan.merges


def test_numeric_preflight_passes_non_decimal_values_through_unchanged() -> None:
    values = (
        None,
        True,
        12,
        1.0,
        5e-324,
        "12.5",
        "=1+2",
        date(2026, 1, 1),
        datetime(2026, 1, 1, 12),
        time(12),
    )
    plan = _plan(*values)

    adapted = preflight_numeric_plans((plan,)).require()[0]

    assert all(new is old for new, old in zip(adapted.cells, plan.cells, strict=True))
    assert all(cell.value is value for cell, value in zip(adapted.cells, values, strict=True))
    assert preflight_numeric_plans(()).require() == ()


@pytest.mark.parametrize("precision", range(1, 16))
@pytest.mark.parametrize("sign", (0, 1))
@pytest.mark.parametrize("exponent", (-20, -3, 0, 20))
@pytest.mark.parametrize("trailing_zeros", (0, 4))
def test_numeric_preflight_accepts_each_precision_from_one_through_fifteen(
    precision: int, sign: int, exponent: int, trailing_zeros: int
) -> None:
    digits = tuple(map(int, "314159265358979"[:precision])) + (0,) * trailing_zeros
    _assert_accepted(Decimal((sign, digits, exponent - trailing_zeros)))


@pytest.mark.parametrize(
    "text",
    (
        "12.5",
        "12.50",
        "12.5000",
        "1.25E+1",
        "123456789012345",
        "-123456789012345",
        "999999999999999",
        "-999999999999999",
    ),
)
def test_numeric_preflight_treats_scale_and_trailing_zeros_as_presentation(text: str) -> None:
    _assert_accepted(Decimal(text))


def test_numeric_preflight_precision_is_independent_of_decimal_context() -> None:
    with localcontext() as context:
        context.prec = 2
        _assert_accepted(Decimal("123456789012345.000000"))
        result = preflight_numeric_plans((_plan(Decimal("1234567890123456.000000")),))
    assert [item.code for item in result.diagnostics] == [
        DiagnosticCode.XLSX_DECIMAL_PRECISION_EXCEEDED
    ]


@pytest.mark.parametrize(
    "text",
    (
        "0",
        "-0",
        "0E-1000",
        "-0E+1000",
        "2.22507385850721E-308",
        "-2.22507385850721E-308",
        "1.79769313486231E+308",
        "-1.79769313486231E+308",
    ),
)
def test_numeric_preflight_accepts_zero_and_normal_range_boundaries(text: str) -> None:
    _assert_accepted(Decimal(text))


@pytest.mark.parametrize(
    "text", ("1234567890123456", "-1234567890123456", "1.234567890123456", "-1.234567890123456")
)
def test_numeric_preflight_rejects_sixteen_significant_digits(text: str) -> None:
    result = preflight_numeric_plans((_plan(Decimal(text)),))
    assert result.plans is None
    assert [item.code for item in result.diagnostics] == [
        DiagnosticCode.XLSX_DECIMAL_PRECISION_EXCEEDED
    ]


@pytest.mark.parametrize("sign", ("", "-"))
@pytest.mark.parametrize(
    "text",
    ("2.22507385850720E-308", "1E-308", "1E-323", "1E-324", "1.79769313486232E+308", "1E+309"),
)
def test_numeric_preflight_rejects_overflow_underflow_and_subnormal_values(
    sign: str, text: str
) -> None:
    result = preflight_numeric_plans((_plan(Decimal(sign + text)),))
    assert result.plans is None
    assert [item.code for item in result.diagnostics] == [DiagnosticCode.XLSX_DECIMAL_OUT_OF_RANGE]


class _InexactDecimal(Decimal):
    def __float__(self) -> float:
        return 0.2


def test_numeric_preflight_reports_inexact_decimal_conversion() -> None:
    result = preflight_numeric_plans((_plan(_InexactDecimal("0.1")),))
    assert result.plans is None
    assert [item.code for item in result.diagnostics] == [DiagnosticCode.XLSX_DECIMAL_INEXACT]


def test_numeric_preflight_applies_precision_range_inexact_precedence() -> None:
    result = preflight_numeric_plans(
        (_plan(Decimal("1234567890123456E+309"), Decimal("1E-324"), _InexactDecimal("0.1")),)
    )
    assert [item.code for item in result.diagnostics] == [
        DiagnosticCode.XLSX_DECIMAL_PRECISION_EXCEEDED,
        DiagnosticCode.XLSX_DECIMAL_OUT_OF_RANGE,
        DiagnosticCode.XLSX_DECIMAL_INEXACT,
    ]


def test_numeric_preflight_aggregates_each_rendered_destination() -> None:
    value = Decimal("1234567890123456")
    plans = (_plan(value, value, sheet="Second"), _plan(value, sheet="First"))

    result = preflight_numeric_plans(plans)

    assert [item.location for item in result.diagnostics] == [
        SourceLocation("Second", "A1"),
        SourceLocation("Second", "A2"),
        SourceLocation("First", "A1"),
    ]
    assert all(
        item.code == DiagnosticCode.XLSX_DECIMAL_PRECISION_EXCEEDED for item in result.diagnostics
    )


def test_numeric_preflight_returns_no_plans_when_any_decimal_fails() -> None:
    valid = _plan(Decimal("12.5"))
    invalid = _plan(Decimal("1234567890123456"))

    result = preflight_numeric_plans((valid, invalid))

    assert result.plans is None
    with pytest.raises(TemplateRenderError) as caught:
        result.require()
    assert caught.value.diagnostics == result.diagnostics
    assert valid.cells[0].value == Decimal("12.5")
    assert type(valid.cells[0].value) is Decimal


@st.composite
def _decimal_tuples(draw: st.DrawFn) -> Decimal:
    digits = draw(st.lists(st.integers(0, 9), min_size=1, max_size=20))
    zeros = draw(st.integers(0, 6))
    exponent = draw(
        st.sampled_from((-350, -324, -323, -308, -307, -20, -3, 0, 20, 294, 308, 309, 350))
    )
    return Decimal((draw(st.integers(0, 1)), tuple(digits) + (0,) * zeros, exponent - zeros))


def _public_predicate(value: Decimal) -> bool:
    coefficient = "".join(map(str, value.as_tuple().digits)).rstrip("0")
    return (
        value.is_finite()
        and len(coefficient) <= 15
        and math.isfinite(float(value))
        and (value == 0 or abs(float(value)) >= sys.float_info.min)
        and Decimal(repr(float(value))) == value
    )


@settings(max_examples=300)
@given(_decimal_tuples())
def test_generated_decimal_classification_matches_public_predicate(value: Decimal) -> None:
    original = value.as_tuple()
    result = preflight_numeric_plans((_plan(value),))
    assert (result.plans is not None) == _public_predicate(value)
    assert value.as_tuple() == original
    if result.plans is not None:
        assert type(result.require()[0].cells[0].value) is float
        assert Decimal(repr(result.require()[0].cells[0].value)) == value
    else:
        assert len(result.diagnostics) == 1
        assert result.diagnostics[0].location == SourceLocation("Report", "A1")
