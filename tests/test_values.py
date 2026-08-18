from __future__ import annotations

from datetime import UTC, date, datetime, time
from decimal import Decimal

from excel_template_writer.diagnostics import ContextLocation, DiagnosticCode
from excel_template_writer.values import validate_context


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
