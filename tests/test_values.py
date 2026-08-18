from __future__ import annotations

from datetime import UTC, date, datetime, time
from decimal import Decimal
from typing import Any, cast

import pytest

from excel_template_writer.diagnostics import (
    ContextLocation,
    DiagnosticCode,
    TemplateRenderError,
)
from excel_template_writer.values import TypeAdapter, normalize_context, validate_context


def _codes_by_path(context: object) -> dict[str, DiagnosticCode]:
    diagnostics = validate_context(context)
    return {
        diagnostic.location.path: diagnostic.code
        for diagnostic in diagnostics
        if isinstance(diagnostic.location, ContextLocation)
    }


def test_accepts_complete_canonical_value_tree() -> None:
    shared_record = {"name": "Shared"}
    context = {
        "nothing": None,
        "label": "Report",
        "approved": True,
        "count": 3,
        "ratio": 0.25,
        "amount": Decimal("12.50"),
        "issued_on": date(2026, 8, 13),
        "created_at": datetime(2026, 8, 13, 9, 30),
        "cutoff": time(17, 0),
        "record": {"name": "Acme", "tags": ["North", "Renewal"]},
        "rows": ({"name": "A", "amount": 10}, {"name": "B", "amount": 20}),
        "shared_a": shared_record,
        "shared_b": shared_record,
    }

    assert validate_context(context) == ()


def test_reports_every_invalid_value_with_its_context_path() -> None:
    context = {
        "rows": [{"amount": float("inf")}, {"amount": Decimal("NaN")}],
        "labels": {"North", "South"},
        "created_at": datetime(2026, 8, 13, tzinfo=UTC),
        "record": {7: "non-string key"},
        "opaque": object(),
        "stream": (item for item in range(2)),
    }

    assert _codes_by_path(context) == {
        "context.rows[0].amount": DiagnosticCode.NON_FINITE_CONTEXT_NUMBER,
        "context.rows[1].amount": DiagnosticCode.NON_FINITE_CONTEXT_NUMBER,
        "context.labels": DiagnosticCode.UNORDERED_CONTEXT_COLLECTION,
        "context.created_at": DiagnosticCode.TIMEZONE_AWARE_CONTEXT_VALUE,
        "context.record[7]": DiagnosticCode.CONTEXT_KEY_MUST_BE_STRING,
        "context.opaque": DiagnosticCode.UNSUPPORTED_CONTEXT_VALUE,
        "context.stream": DiagnosticCode.UNSUPPORTED_CONTEXT_VALUE,
    }


def test_rejects_a_non_mapping_context_root() -> None:
    diagnostics = validate_context(["not", "a", "mapping"])

    assert len(diagnostics) == 1
    assert diagnostics[0].code is DiagnosticCode.CONTEXT_MUST_BE_MAPPING
    assert str(diagnostics[0].location) == "context"


def test_rejects_cycles_but_allows_shared_subtrees() -> None:
    cyclic: dict[str, object] = {}
    cyclic["self"] = cyclic

    diagnostics = validate_context({"cycle": cyclic})

    assert len(diagnostics) == 1
    assert diagnostics[0].code is DiagnosticCode.CYCLIC_CONTEXT_VALUE
    assert str(diagnostics[0].location) == "context.cycle.self"

    shared = {"value": 1}
    assert validate_context({"first": shared, "second": shared}) == ()


def test_normalizes_to_an_immutable_snapshot() -> None:
    raw = {
        "rows": [{"name": "Alpha", "tags": ["new", "priority"]}],
        "issued_on": date(2026, 8, 19),
    }

    normalized = normalize_context(raw).require()
    raw["rows"][0]["name"] = "Changed"
    raw["rows"].append({"name": "Beta", "tags": []})

    rows = normalized["rows"]
    assert isinstance(rows, tuple)
    assert rows == ({"name": "Alpha", "tags": ("new", "priority")},)
    with pytest.raises(TypeError):
        cast(Any, normalized)["extra"] = 1
    with pytest.raises(TypeError):
        cast(Any, rows[0])["name"] = "Changed"


def test_failed_normalization_has_no_partial_context() -> None:
    result = normalize_context({"first": object(), "second": {1, 2}})

    assert result.context is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.UNSUPPORTED_CONTEXT_VALUE,
        DiagnosticCode.UNORDERED_CONTEXT_COLLECTION,
    ]
    with pytest.raises(TemplateRenderError) as caught:
        result.require()
    assert caught.value.diagnostics == result.diagnostics


class _Frame:
    def __init__(self, rows: list[dict[str, object]]) -> None:
        self.rows = rows


class _CellValue:
    def __init__(self, value: object) -> None:
        self.value = value


class _Tree:
    def __init__(self, name: str, children: list[_Tree]) -> None:
        self.name = name
        self.children = children


def test_adapter_output_is_recursively_normalized() -> None:
    frame = _Frame([{"name": "Alpha", "values": _CellValue([1, 2])}])

    normalized = normalize_context(
        {"rows": frame},
        adapters=(
            TypeAdapter(_Frame, lambda value: value.rows),
            TypeAdapter(_CellValue, lambda value: value.value),
        ),
    ).require()

    assert normalized["rows"] == ({"name": "Alpha", "values": (1, 2)},)

    tree = normalize_context(
        {"root": _Tree("parent", [_Tree("child", [])])},
        adapters=(
            TypeAdapter(
                _Tree,
                lambda value: {"name": value.name, "children": value.children},
            ),
        ),
    ).require()
    assert tree["root"] == {
        "name": "parent",
        "children": ({"name": "child", "children": ()},),
    }


def test_most_specific_adapter_wins_independent_of_registration_order() -> None:
    class Base:
        pass

    class Child(Base):
        pass

    adapters = (
        TypeAdapter(Base, lambda _: "base"),
        TypeAdapter(Child, lambda _: "child"),
    )

    normalized = normalize_context({"value": Child()}, adapters=adapters).require()

    assert normalized["value"] == "child"


def test_canonical_values_are_not_intercepted_by_adapters() -> None:
    calls = 0

    def convert(_: object) -> str:
        nonlocal calls
        calls += 1
        return "intercepted"

    normalized = normalize_context(
        {"text": "kept", "rows": [1, 2]},
        adapters=(TypeAdapter(object, convert),),
    ).require()

    assert dict(normalized) == {"text": "kept", "rows": (1, 2)}
    assert calls == 0


def test_duplicate_and_unrelated_adapter_matches_are_explicit_errors() -> None:
    duplicate = normalize_context(
        {"value": _Frame([])},
        adapters=(
            TypeAdapter(_Frame, lambda _: []),
            TypeAdapter(_Frame, lambda _: []),
        ),
    )
    assert [diagnostic.code for diagnostic in duplicate.diagnostics] == [
        DiagnosticCode.DUPLICATE_VALUE_ADAPTER
    ]

    class Left:
        pass

    class Right:
        pass

    class Both(Left, Right):
        pass

    ambiguous = normalize_context(
        {"value": Both()},
        adapters=(
            TypeAdapter(Left, lambda _: "left"),
            TypeAdapter(Right, lambda _: "right"),
        ),
    )
    assert [diagnostic.code for diagnostic in ambiguous.diagnostics] == [
        DiagnosticCode.AMBIGUOUS_VALUE_ADAPTER
    ]
    assert str(ambiguous.diagnostics[0].location) == "context.value"


def test_adapter_failure_invalid_output_and_cycle_keep_the_value_path() -> None:
    def fail(_: _Frame) -> object:
        raise RuntimeError("conversion unavailable")

    failed = normalize_context(
        {"frame": _Frame([])},
        adapters=(TypeAdapter(_Frame, fail, name="frame"),),
    )
    assert [diagnostic.code for diagnostic in failed.diagnostics] == [
        DiagnosticCode.VALUE_ADAPTER_FAILED
    ]
    assert str(failed.diagnostics[0].location) == "context.frame"

    invalid = normalize_context(
        {"frame": _Frame([])},
        adapters=(TypeAdapter(_Frame, lambda _: {float("nan")}),),
    )
    assert [diagnostic.code for diagnostic in invalid.diagnostics] == [
        DiagnosticCode.UNORDERED_CONTEXT_COLLECTION
    ]
    assert str(invalid.diagnostics[0].location) == "context.frame"

    cyclic = normalize_context(
        {"frame": _Frame([])},
        adapters=(TypeAdapter(_Frame, lambda _: _Frame([])),),
    )
    assert [diagnostic.code for diagnostic in cyclic.diagnostics] == [
        DiagnosticCode.VALUE_ADAPTER_CYCLE
    ]
    assert str(cyclic.diagnostics[0].location) == "context.frame"

    class First:
        other: object

    class Second:
        other: object

    indirect = normalize_context(
        {"value": First()},
        adapters=(
            TypeAdapter(First, lambda _: Second()),
            TypeAdapter(Second, lambda _: First()),
        ),
    )
    assert [diagnostic.code for diagnostic in indirect.diagnostics] == [
        DiagnosticCode.VALUE_ADAPTER_CYCLE
    ]
