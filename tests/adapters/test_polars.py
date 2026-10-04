from __future__ import annotations

from collections.abc import Mapping
from datetime import UTC, date, datetime, time, timedelta
from decimal import Decimal
from importlib import import_module
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

import pytest
from openpyxl import Workbook, load_workbook

from excel_template_writer.diagnostics import (
    ContextLocation,
    DiagnosticCode,
    SourceLocation,
    TemplateRenderError,
)
from excel_template_writer.limits import ResourceLimits
from excel_template_writer.values import CanonicalValue, NormalizationResult, normalize_context
from excel_template_writer.xlsx import render_workbook

pl = pytest.importorskip("polars")
polars_adapters = import_module("excel_template_writer.adapters.polars").polars_adapters


def _codes_by_path(result: NormalizationResult) -> dict[str, DiagnosticCode]:
    return {
        diagnostic.location.path: diagnostic.code
        for diagnostic in result.diagnostics
        if isinstance(diagnostic.location, ContextLocation)
    }


def test_dataframe_becomes_ordered_canonical_records() -> None:
    frame = pl.DataFrame(
        {
            "name": ["Beta", "Alpha"],
            "amount": [Decimal("12.50"), Decimal("3.75")],
            "issued_on": [date(2026, 8, 19), date(2026, 8, 20)],
            "created_at": [datetime(2026, 8, 19, 9), datetime(2026, 8, 20, 10)],
            "optional": [None, "present"],
        },
        strict=False,
    )

    normalized = normalize_context(
        {"rows": frame},
        adapters=polars_adapters(),
    ).require()

    assert normalized["rows"] == (
        {
            "name": "Beta",
            "amount": Decimal("12.50"),
            "issued_on": date(2026, 8, 19),
            "created_at": datetime(2026, 8, 19, 9),
            "optional": None,
        },
        {
            "name": "Alpha",
            "amount": Decimal("3.75"),
            "issued_on": date(2026, 8, 20),
            "created_at": datetime(2026, 8, 20, 10),
            "optional": "present",
        },
    )


def test_null_and_nested_float_nan_become_canonical_null() -> None:
    frame = pl.DataFrame(
        {
            "metrics": [
                {"values": [1.0, float("nan")]},
                {"values": [None, 2.0]},
            ]
        }
    )

    normalized = normalize_context(
        {"rows": frame},
        adapters=polars_adapters(),
    ).require()

    assert normalized["rows"] == (
        {"metrics": {"values": (1.0, None)}},
        {"metrics": {"values": (None, 2.0)}},
    )


def test_empty_dataframe_becomes_an_empty_collection() -> None:
    frame = pl.DataFrame(schema={"name": pl.String, "amount": pl.Int64})

    normalized = normalize_context(
        {"rows": frame},
        adapters=polars_adapters(),
    ).require()

    assert normalized["rows"] == ()


def test_adapter_is_explicit_and_does_not_collect_lazy_frames() -> None:
    frame = pl.DataFrame({"name": ["Alpha"]})

    without_adapter = normalize_context({"rows": frame})
    lazy = normalize_context({"rows": frame.lazy()}, adapters=polars_adapters())

    assert _codes_by_path(without_adapter) == {
        "context.rows": DiagnosticCode.UNSUPPORTED_CONTEXT_VALUE
    }
    assert _codes_by_path(lazy) == {"context.rows": DiagnosticCode.UNSUPPORTED_CONTEXT_VALUE}


def test_canonical_validation_reports_unsupported_materialized_values() -> None:
    frame = pl.DataFrame(
        {
            "infinite": [float("inf")],
            "duration": [timedelta(days=1)],
            "binary": [b"opaque"],
        }
    )

    result = normalize_context({"rows": frame}, adapters=polars_adapters())

    assert _codes_by_path(result) == {
        "context.rows[0].infinite": DiagnosticCode.NON_FINITE_CONTEXT_NUMBER,
        "context.rows[0].duration": DiagnosticCode.UNSUPPORTED_CONTEXT_VALUE,
        "context.rows[0].binary": DiagnosticCode.UNSUPPORTED_CONTEXT_VALUE,
    }


@pytest.mark.parametrize(
    "series",
    [
        pl.Series("when", [datetime(2026, 8, 19, 9)], dtype=pl.Datetime("ns")),
        pl.Series("when", [time(9, 30)], dtype=pl.Time),
        pl.Series(
            "when",
            [datetime(2026, 8, 19, 9, tzinfo=UTC)],
            dtype=pl.Datetime("us", "UTC"),
        ),
    ],
    ids=["nanosecond-datetime", "time", "timezone-aware-datetime"],
)
def test_temporal_values_that_cannot_be_preserved_are_rejected_before_conversion(
    series: object,
) -> None:
    frame = pl.DataFrame([series])

    result = normalize_context({"rows": frame}, adapters=polars_adapters())

    assert _codes_by_path(result) == {"context.rows": DiagnosticCode.VALUE_ADAPTER_FAILED}


def test_resource_limits_apply_after_dataframe_conversion() -> None:
    frame = pl.DataFrame({"name": ["Alpha", "Beta"]})
    limits = ResourceLimits(max_container_items=1)

    result = normalize_context(
        {"rows": frame},
        adapters=polars_adapters(),
        limits=limits,
    )

    assert _codes_by_path(result) == {
        "context.rows": DiagnosticCode.CONTEXT_RESOURCE_LIMIT_EXCEEDED
    }


def test_dataframe_renders_through_the_xlsx_entrypoint(tmp_path: Path) -> None:
    template_path = tmp_path / "polars-template.xlsx"
    output_path = tmp_path / "polars-output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "{% for row in rows %}{{ row.name }}"
    sheet["B1"] = "{{ row.amount }}{% endfor %}"
    workbook.save(template_path)
    workbook.close()
    frame = pl.DataFrame({"name": ["Alpha", "Beta"], "amount": [10, 20]})

    render_workbook(
        template_path,
        output_path,
        {"rows": frame},
        adapters=polars_adapters(),
    )

    rendered = load_workbook(output_path)
    try:
        sheet = rendered.active
        assert sheet["A1"].value == "Alpha"
        assert sheet["B1"].value == 10
        assert sheet["A2"].value == "Beta"
        assert sheet["B2"].value == 20
    finally:
        rendered.close()


def _result_template(path: Path) -> Path:
    workbook = Workbook()
    try:
        sheet = workbook.active
        sheet.title = "Results"
        sheet["A1"] = "{% for row in rows %}{{ row.result }}{% endfor %}"
        sheet["A1"].number_format = "0.0000"
        workbook.save(path)
    finally:
        workbook.close()
    return path


def _assert_result_output(path: Path, values: list[float]) -> None:
    with ZipFile(path) as archive:
        root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    xml_cells = {cell.attrib["r"]: cell for cell in root.findall(".//{*}c")}
    reopened = load_workbook(path)
    try:
        for index, value in enumerate(values, 1):
            assert type(value) is float
            coordinate = f"A{index}"
            xml_cell = xml_cells[coordinate]
            assert xml_cell.get("t") == "n"
            token = xml_cell.find("{*}v")
            assert token is not None and token.text == repr(value)
            cell = reopened["Results"][coordinate]
            assert cell.data_type == "n"
            assert type(cell.value) is float
            assert cell.value.hex() == value.hex()
            assert cell.number_format == "0.0000"
    finally:
        reopened.close()


def _canonical_results(frame: object) -> list[CanonicalValue]:
    normalized = normalize_context({"rows": frame}, adapters=polars_adapters()).require()
    rows = normalized["rows"]
    assert isinstance(rows, tuple)
    results: list[CanonicalValue] = []
    for row in rows:
        assert isinstance(row, Mapping)
        results.append(row["result"])
    return results


def test_polars_float_operation_result_round_trips_bit_exactly(tmp_path: Path) -> None:
    frame = pl.DataFrame({"left": [0.1], "right": [0.2]}).with_columns(
        (pl.col("left") + pl.col("right")).alias("result")
    )
    scalar = frame["result"][0]
    assert frame.schema["result"] == pl.Float64
    assert type(scalar) is float
    assert scalar.hex() == (0.30000000000000004).hex()
    adapted = _canonical_results(frame)[0]
    assert type(adapted) is float
    assert adapted.hex() == scalar.hex()
    template = _result_template(tmp_path / "template.xlsx")
    output = tmp_path / "output.xlsx"

    render_workbook(template, output, {"rows": frame}, adapters=polars_adapters())

    _assert_result_output(output, [scalar])


def test_polars_decimal_operations_cover_one_through_fifteen_digits(tmp_path: Path) -> None:
    expected = [Decimal("314159265358979"[:digits]) for digits in range(1, 16)]
    frame = pl.DataFrame(
        {
            "left": [value - Decimal(1) for value in expected],
            "right": [Decimal(1)] * 15,
        },
        schema={"left": pl.Decimal(precision=18, scale=0), "right": pl.Decimal(18, 0)},
    ).with_columns((pl.col("left") + pl.col("right")).alias("result"))
    assert frame.schema["result"].is_decimal()
    assert frame.schema["result"].scale == 0
    scalars = frame["result"].to_list()
    assert all(type(scalar) is Decimal for scalar in scalars)
    assert scalars == expected
    adapted = _canonical_results(frame)
    assert all(type(scalar) is Decimal for scalar in adapted)
    assert adapted == expected
    template = _result_template(tmp_path / "template.xlsx")
    output = tmp_path / "output.xlsx"

    render_workbook(template, output, {"rows": frame}, adapters=polars_adapters())

    _assert_result_output(output, [float(value) for value in expected])


def test_polars_sixteen_digit_decimal_result_is_rejected_at_destination_cell(
    tmp_path: Path,
) -> None:
    frame = pl.DataFrame(
        {"left": [Decimal("1234567890123455")], "right": [Decimal(1)]},
        schema={"left": pl.Decimal(18, 0), "right": pl.Decimal(18, 0)},
    ).with_columns((pl.col("left") + pl.col("right")).alias("result"))
    assert frame.schema["result"].is_decimal()
    scalar = frame["result"][0]
    assert type(scalar) is Decimal
    assert scalar == Decimal("1234567890123456")
    assert _canonical_results(frame) == [scalar]
    template = _result_template(tmp_path / "template.xlsx")
    output = tmp_path / "absent" / "output.xlsx"

    with pytest.raises(TemplateRenderError) as caught:
        render_workbook(template, output, {"rows": frame}, adapters=polars_adapters())

    assert [(item.code, item.location) for item in caught.value.diagnostics] == [
        (DiagnosticCode.XLSX_DECIMAL_PRECISION_EXCEEDED, SourceLocation("Results", "A1"))
    ]
    assert not output.parent.exists()


@pytest.mark.parametrize(
    ("left", "right", "operation", "expected"),
    [
        (0.5, 0.25, "add", 0.75),
        (1.5, 0.25, "subtract", 1.25),
        (1.125, 2.0, "multiply", 2.25),
        (5.0, 2.0, "divide", 2.5),
    ],
    ids=["0.5+0.25", "1.5-0.25", "1.125*2", "5/2"],
)
def test_polars_exact_binary_fraction_operations_round_trip(
    tmp_path: Path, left: float, right: float, operation: str, expected: float
) -> None:
    assert type(expected) is float
    expressions = {
        "add": pl.col("left") + pl.col("right"),
        "subtract": pl.col("left") - pl.col("right"),
        "multiply": pl.col("left") * pl.col("right"),
        "divide": pl.col("left") / pl.col("right"),
    }
    frame = pl.DataFrame({"left": [left], "right": [right]}).with_columns(
        expressions[operation].alias("result")
    )
    assert frame.schema["result"] == pl.Float64
    scalar = frame["result"][0]
    assert type(scalar) is float and scalar.hex() == expected.hex()
    adapted = _canonical_results(frame)[0]
    assert type(adapted) is float and adapted.hex() == expected.hex()
    template = _result_template(tmp_path / "template.xlsx")
    output = tmp_path / "output.xlsx"

    render_workbook(template, output, {"rows": frame}, adapters=polars_adapters())

    _assert_result_output(output, [expected])


@pytest.mark.parametrize(
    ("left", "right", "operation", "coordinate", "formula"),
    [
        (0.5, 0.25, "add", "B19", "=0.5+0.25"),
        (1.5, 0.25, "subtract", "B20", "=1.5-0.25"),
        (1.125, 2.0, "multiply", "B21", "=1.125*2"),
        (5.0, 2.0, "divide", "B22", "=5/2"),
    ],
)
def test_polars_exact_binary_operations_match_frozen_excel_reference(
    tmp_path: Path, left: float, right: float, operation: str, coordinate: str, formula: str
) -> None:
    reference_path = (
        Path(__file__).parents[1] / "fixtures/numeric_fidelity/excel_numeric_reference.xlsx"
    )
    assert reference_path.is_file(), "Frozen desktop-Excel workbook is missing"
    formulas = load_workbook(reference_path, data_only=False)
    reference = load_workbook(reference_path, data_only=True)
    try:
        assert formulas["NumericReference"][coordinate].value == formula
        expected = reference["NumericReference"][coordinate].value
        assert type(expected) is float
        expressions = {
            "add": pl.col("left") + pl.col("right"),
            "subtract": pl.col("left") - pl.col("right"),
            "multiply": pl.col("left") * pl.col("right"),
            "divide": pl.col("left") / pl.col("right"),
        }
        frame = pl.DataFrame({"left": [left], "right": [right]}).with_columns(
            expressions[operation].alias("result")
        )
        assert frame.schema["result"] == pl.Float64
        scalar = frame["result"][0]
        assert type(scalar) is float and scalar.hex() == expected.hex()
        adapted = _canonical_results(frame)[0]
        assert type(adapted) is float and adapted.hex() == expected.hex()
        template = _result_template(tmp_path / "template.xlsx")
        output = tmp_path / "output.xlsx"

        render_workbook(template, output, {"rows": frame}, adapters=polars_adapters())

        _assert_result_output(output, [expected])
    finally:
        formulas.close()
        reference.close()
