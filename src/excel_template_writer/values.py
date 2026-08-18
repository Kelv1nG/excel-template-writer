"""Canonical render-context values and path-aware validation."""

from __future__ import annotations

import math
from collections.abc import Mapping
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal

from excel_template_writer.diagnostics import ContextLocation, Diagnostic, DiagnosticCode

type ScalarValue = str | bool | int | float | Decimal | date | datetime | time | None
type CanonicalValue = (
    ScalarValue | Mapping[str, CanonicalValue] | list[CanonicalValue] | tuple[CanonicalValue, ...]
)
type CanonicalContext = Mapping[str, CanonicalValue]


@dataclass(frozen=True)
class _PendingValue:
    value: object
    path: str
    ancestors: frozenset[int]


def _diagnostic(code: DiagnosticCode, message: str, path: str) -> Diagnostic:
    return Diagnostic(code, message, ContextLocation(path))


def _type_name(value: object) -> str:
    value_type = type(value)
    if value_type.__module__ == "builtins":
        return value_type.__qualname__
    return f"{value_type.__module__}.{value_type.__qualname__}"


def _key_path(parent: str, key: object) -> str:
    if isinstance(key, str) and key.isidentifier():
        return f"{parent}.{key}"
    return f"{parent}[{key!r}]"


def _is_timezone_aware(value: datetime | time) -> bool:
    try:
        return value.utcoffset() is not None
    except (OverflowError, ValueError):
        return True


def _container_children(
    value: Mapping[object, object] | list[object] | tuple[object, ...],
    path: str,
) -> tuple[tuple[str, object], ...]:
    if isinstance(value, Mapping):
        return tuple((_key_path(path, key), item) for key, item in value.items())
    return tuple((f"{path}[{index}]", item) for index, item in enumerate(value))


def validate_context(context: object) -> tuple[Diagnostic, ...]:
    """Validate a complete value tree without adapting or mutating caller-owned values."""

    if not isinstance(context, Mapping):
        return (
            _diagnostic(
                DiagnosticCode.CONTEXT_MUST_BE_MAPPING,
                f"render context must be a mapping, not {_type_name(context)}",
                "context",
            ),
        )

    diagnostics: list[Diagnostic] = []
    pending = [_PendingValue(context, "context", frozenset())]
    while pending:
        current = pending.pop()
        value = current.value

        if value is None or isinstance(value, (str, bool, int)):
            continue
        if isinstance(value, float):
            if not math.isfinite(value):
                diagnostics.append(
                    _diagnostic(
                        DiagnosticCode.NON_FINITE_CONTEXT_NUMBER,
                        "floating-point values must be finite",
                        current.path,
                    )
                )
            continue
        if isinstance(value, Decimal):
            if not value.is_finite():
                diagnostics.append(
                    _diagnostic(
                        DiagnosticCode.NON_FINITE_CONTEXT_NUMBER,
                        "decimal values must be finite",
                        current.path,
                    )
                )
            continue
        if isinstance(value, datetime):
            if _is_timezone_aware(value):
                diagnostics.append(
                    _diagnostic(
                        DiagnosticCode.TIMEZONE_AWARE_CONTEXT_VALUE,
                        "datetime values must not carry timezone information",
                        current.path,
                    )
                )
            continue
        if isinstance(value, time):
            if _is_timezone_aware(value):
                diagnostics.append(
                    _diagnostic(
                        DiagnosticCode.TIMEZONE_AWARE_CONTEXT_VALUE,
                        "time values must not carry timezone information",
                        current.path,
                    )
                )
            continue
        if isinstance(value, date):
            continue
        if isinstance(value, (set, frozenset)):
            diagnostics.append(
                _diagnostic(
                    DiagnosticCode.UNORDERED_CONTEXT_COLLECTION,
                    "sets and frozensets are not canonical ordered collections",
                    current.path,
                )
            )
            continue
        if isinstance(value, (Mapping, list, tuple)):
            identity = id(value)
            if identity in current.ancestors:
                diagnostics.append(
                    _diagnostic(
                        DiagnosticCode.CYCLIC_CONTEXT_VALUE,
                        "canonical context values must form an acyclic tree",
                        current.path,
                    )
                )
                continue
            next_ancestors = current.ancestors | {identity}
            if isinstance(value, Mapping):
                for key in value:
                    if not isinstance(key, str):
                        diagnostics.append(
                            _diagnostic(
                                DiagnosticCode.CONTEXT_KEY_MUST_BE_STRING,
                                f"record keys must be strings, not {_type_name(key)}",
                                _key_path(current.path, key),
                            )
                        )
            children = _container_children(value, current.path)
            pending.extend(
                _PendingValue(item, child_path, next_ancestors)
                for child_path, item in reversed(children)
            )
            continue

        diagnostics.append(
            _diagnostic(
                DiagnosticCode.UNSUPPORTED_CONTEXT_VALUE,
                f"unsupported canonical value type: {_type_name(value)}",
                current.path,
            )
        )

    return tuple(diagnostics)


def is_collection_value(value: object) -> bool:
    """Return whether a canonical value is non-scalar."""

    return isinstance(value, (Mapping, list, tuple))


def is_ordered_collection(value: object) -> bool:
    """Return whether a canonical value may drive a repeat."""

    return isinstance(value, (list, tuple))


__all__ = [
    "CanonicalContext",
    "CanonicalValue",
    "ScalarValue",
    "is_collection_value",
    "is_ordered_collection",
    "validate_context",
]
