"""Canonical render-context normalization and caller-supplied value adapters."""

from __future__ import annotations

import math
from collections.abc import Callable, Iterable, Iterator, Mapping
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
from types import MappingProxyType
from typing import Any, cast

from excel_template_writer.diagnostics import (
    ContextLocation,
    Diagnostic,
    DiagnosticCode,
    TemplateRenderError,
)
from excel_template_writer.limits import DEFAULT_RESOURCE_LIMITS, ResourceLimits

type ScalarValue = str | bool | int | float | Decimal | date | datetime | time | None
type InputValue = ScalarValue | Mapping[str, InputValue] | list[InputValue] | tuple[InputValue, ...]
type CanonicalValue = ScalarValue | Mapping[str, CanonicalValue] | tuple[CanonicalValue, ...]
type CanonicalContext = Mapping[str, CanonicalValue]


@dataclass(frozen=True, slots=True)
class ContextStatistics:
    """Summary measurements retained with an immutable normalized context."""

    nodes: int
    maximum_depth: int
    maximum_container_items: int
    maximum_string_length: int


class NormalizedContext(Mapping[str, CanonicalValue]):
    """An immutable context produced by :func:`normalize_context`."""

    __slots__ = ("_data", "_statistics")
    _data: Mapping[str, CanonicalValue]
    _statistics: ContextStatistics

    def __init__(self) -> None:
        raise TypeError("NormalizedContext values are created by normalize_context()")

    @classmethod
    def _create(
        cls,
        values: Mapping[str, CanonicalValue],
        statistics: ContextStatistics,
    ) -> NormalizedContext:
        instance = object.__new__(cls)
        object.__setattr__(instance, "_data", MappingProxyType(dict(values)))
        object.__setattr__(instance, "_statistics", statistics)
        return instance

    @property
    def statistics(self) -> ContextStatistics:
        return self._statistics

    def __getitem__(self, key: str) -> CanonicalValue:
        return self._data[key]

    def __iter__(self) -> Iterator[str]:
        return iter(self._data)

    def __len__(self) -> int:
        return len(self._data)

    def __repr__(self) -> str:
        return f"NormalizedContext({dict(self._data)!r})"

    def __setattr__(self, name: str, value: object) -> None:
        raise TypeError("NormalizedContext is immutable")


@dataclass(frozen=True, slots=True)
class NormalizationResult:
    """The all-or-nothing result of context normalization."""

    context: NormalizedContext | None
    diagnostics: tuple[Diagnostic, ...]

    def require(self) -> NormalizedContext:
        """Return the normalized context or raise all collected diagnostics."""

        if self.context is None:
            raise TemplateRenderError(self.diagnostics)
        return self.context


@dataclass(frozen=True, slots=True)
class TypeAdapter[T]:
    """Convert one caller-owned runtime type into canonical input values."""

    source_type: type[T]
    converter: Callable[[T], object]
    name: str | None = None

    def __post_init__(self) -> None:
        if not isinstance(self.source_type, type):
            raise TypeError("adapter source_type must be a runtime type")
        if not callable(self.converter):
            raise TypeError("adapter converter must be callable")
        if self.name is not None and not self.name:
            raise ValueError("adapter name must not be empty")

    @property
    def display_name(self) -> str:
        return self.name or _type_name(self.source_type)

    def convert(self, value: object) -> object:
        return self.converter(cast(T, value))


_UNSET = object()
_INVALID = object()


@dataclass(slots=True)
class _Slot:
    value: object = _UNSET


@dataclass(frozen=True, slots=True)
class _Visit:
    value: object
    path: str
    depth: int
    ancestors: tuple[object, ...]
    adapter_chain: frozenset[type[object]]
    target: _Slot


@dataclass(frozen=True, slots=True)
class _BuildMapping:
    entries: tuple[tuple[str, _Slot], ...]
    has_invalid_key: bool
    target: _Slot


@dataclass(frozen=True, slots=True)
class _BuildSequence:
    entries: tuple[_Slot, ...]
    target: _Slot


type _Action = _Visit | _BuildMapping | _BuildSequence


@dataclass(slots=True)
class _NormalizationState:
    limits: ResourceLimits
    nodes: int = 0
    maximum_depth: int = 0
    maximum_container_items: int = 0
    maximum_string_length: int = 0
    limit_exceeded: bool = False

    @property
    def statistics(self) -> ContextStatistics:
        return ContextStatistics(
            self.nodes,
            self.maximum_depth,
            self.maximum_container_items,
            self.maximum_string_length,
        )


def _diagnostic(code: DiagnosticCode, message: str, path: str) -> Diagnostic:
    return Diagnostic(code, message, ContextLocation(path))


def _type_name(value_or_type: object) -> str:
    value_type = value_or_type if isinstance(value_or_type, type) else type(value_or_type)
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


def _is_ancestor(value: object, ancestors: tuple[object, ...]) -> bool:
    return any(value is ancestor for ancestor in ancestors)


def _is_canonical_runtime_type(value: object) -> bool:
    return value is None or isinstance(
        value,
        (str, bool, int, float, Decimal, date, time, Mapping, list, tuple),
    )


def _record_resource_use(
    action: _Visit,
    diagnostics: list[Diagnostic],
    state: _NormalizationState,
) -> bool:
    value = action.value
    state.nodes += 1
    state.maximum_depth = max(state.maximum_depth, action.depth)
    if state.nodes > state.limits.max_context_nodes:
        message = f"canonical context exceeds max_context_nodes={state.limits.max_context_nodes:,}"
    elif action.depth > state.limits.max_context_depth:
        message = f"canonical context exceeds max_context_depth={state.limits.max_context_depth:,}"
    elif isinstance(value, str):
        size = len(value)
        state.maximum_string_length = max(state.maximum_string_length, size)
        if size <= state.limits.max_input_string_length:
            return False
        message = (
            f"input string exceeds max_input_string_length={state.limits.max_input_string_length:,}"
        )
    elif isinstance(value, (Mapping, list, tuple)):
        size = len(value)
        state.maximum_container_items = max(state.maximum_container_items, size)
        if size <= state.limits.max_container_items:
            return False
        message = f"container exceeds max_container_items={state.limits.max_container_items:,}"
    else:
        return False

    diagnostics.append(
        _diagnostic(
            DiagnosticCode.CONTEXT_RESOURCE_LIMIT_EXCEEDED,
            message,
            action.path,
        )
    )
    state.limit_exceeded = True
    action.target.value = _INVALID
    return True


def _statistics_limit_diagnostic(
    statistics: ContextStatistics,
    limits: ResourceLimits,
) -> Diagnostic | None:
    checks = (
        (
            statistics.nodes,
            limits.max_context_nodes,
            "max_context_nodes",
        ),
        (
            statistics.maximum_depth,
            limits.max_context_depth,
            "max_context_depth",
        ),
        (
            statistics.maximum_container_items,
            limits.max_container_items,
            "max_container_items",
        ),
        (
            statistics.maximum_string_length,
            limits.max_input_string_length,
            "max_input_string_length",
        ),
    )
    for actual, allowed, name in checks:
        if actual > allowed:
            return _diagnostic(
                DiagnosticCode.CONTEXT_RESOURCE_LIMIT_EXCEEDED,
                f"normalized context measurement {actual:,} exceeds {name}={allowed:,}",
                "context",
            )
    return None


def _duplicate_adapter_diagnostics(
    adapters: tuple[TypeAdapter[Any], ...],
) -> tuple[Diagnostic, ...]:
    by_type: dict[type[object], list[TypeAdapter[Any]]] = {}
    for adapter in adapters:
        by_type.setdefault(adapter.source_type, []).append(adapter)
    return tuple(
        _diagnostic(
            DiagnosticCode.DUPLICATE_VALUE_ADAPTER,
            f"multiple value adapters are registered for {_type_name(source_type)}",
            "context",
        )
        for source_type, matches in by_type.items()
        if len(matches) > 1
    )


def _resolve_adapter(
    value: object,
    adapters: tuple[TypeAdapter[Any], ...],
) -> TypeAdapter[Any] | tuple[TypeAdapter[Any], ...] | None:
    matches = tuple(adapter for adapter in adapters if isinstance(value, adapter.source_type))
    if not matches:
        return None
    most_specific = tuple(
        candidate
        for candidate in matches
        if all(issubclass(candidate.source_type, other.source_type) for other in matches)
    )
    if len(most_specific) == 1:
        return most_specific[0]
    return matches


def _visit_mapping(
    action: _Visit,
    value: Mapping[object, object],
    actions: list[_Action],
    diagnostics: list[Diagnostic],
) -> None:
    if _is_ancestor(value, action.ancestors):
        diagnostics.append(
            _diagnostic(
                DiagnosticCode.CYCLIC_CONTEXT_VALUE,
                "canonical context values must form an acyclic tree",
                action.path,
            )
        )
        action.target.value = _INVALID
        return

    next_ancestors = (*action.ancestors, value)
    visits: list[_Visit] = []
    entries: list[tuple[str, _Slot]] = []
    has_invalid_key = False
    for key, item in value.items():
        child = _Slot()
        visits.append(
            _Visit(
                item,
                _key_path(action.path, key),
                action.depth + 1,
                next_ancestors,
                frozenset(),
                child,
            )
        )
        if isinstance(key, str):
            entries.append((key, child))
        else:
            has_invalid_key = True
            diagnostics.append(
                _diagnostic(
                    DiagnosticCode.CONTEXT_KEY_MUST_BE_STRING,
                    f"record keys must be strings, not {_type_name(key)}",
                    _key_path(action.path, key),
                )
            )

    actions.append(_BuildMapping(tuple(entries), has_invalid_key, action.target))
    actions.extend(reversed(visits))


def _visit_sequence(
    action: _Visit,
    value: list[object] | tuple[object, ...],
    actions: list[_Action],
    diagnostics: list[Diagnostic],
) -> None:
    if _is_ancestor(value, action.ancestors):
        diagnostics.append(
            _diagnostic(
                DiagnosticCode.CYCLIC_CONTEXT_VALUE,
                "canonical context values must form an acyclic tree",
                action.path,
            )
        )
        action.target.value = _INVALID
        return

    next_ancestors = (*action.ancestors, value)
    entries = tuple(_Slot() for _ in value)
    actions.append(_BuildSequence(entries, action.target))
    actions.extend(
        _Visit(
            item,
            f"{action.path}[{index}]",
            action.depth + 1,
            next_ancestors,
            frozenset(),
            entries[index],
        )
        for index, item in reversed(tuple(enumerate(value)))
    )


def _build_mapping(action: _BuildMapping) -> None:
    if action.has_invalid_key or any(entry.value is _INVALID for _, entry in action.entries):
        action.target.value = _INVALID
        return
    action.target.value = MappingProxyType(
        {key: cast(CanonicalValue, entry.value) for key, entry in action.entries}
    )


def _build_sequence(action: _BuildSequence) -> None:
    if any(entry.value is _INVALID for entry in action.entries):
        action.target.value = _INVALID
        return
    action.target.value = tuple(cast(CanonicalValue, entry.value) for entry in action.entries)


def _visit_value(
    action: _Visit,
    actions: list[_Action],
    diagnostics: list[Diagnostic],
    adapters: tuple[TypeAdapter[Any], ...],
    state: _NormalizationState,
) -> None:
    value = action.value

    if _is_canonical_runtime_type(value) and _record_resource_use(action, diagnostics, state):
        return

    if value is None or isinstance(value, (str, bool, int)):
        action.target.value = value
        return
    if isinstance(value, float):
        if math.isfinite(value):
            action.target.value = value
        else:
            diagnostics.append(
                _diagnostic(
                    DiagnosticCode.NON_FINITE_CONTEXT_NUMBER,
                    "floating-point values must be finite",
                    action.path,
                )
            )
            action.target.value = _INVALID
        return
    if isinstance(value, Decimal):
        if value.is_finite():
            action.target.value = value
        else:
            diagnostics.append(
                _diagnostic(
                    DiagnosticCode.NON_FINITE_CONTEXT_NUMBER,
                    "decimal values must be finite",
                    action.path,
                )
            )
            action.target.value = _INVALID
        return
    if isinstance(value, datetime):
        if _is_timezone_aware(value):
            diagnostics.append(
                _diagnostic(
                    DiagnosticCode.TIMEZONE_AWARE_CONTEXT_VALUE,
                    "datetime values must not carry timezone information",
                    action.path,
                )
            )
            action.target.value = _INVALID
        else:
            action.target.value = value
        return
    if isinstance(value, time):
        if _is_timezone_aware(value):
            diagnostics.append(
                _diagnostic(
                    DiagnosticCode.TIMEZONE_AWARE_CONTEXT_VALUE,
                    "time values must not carry timezone information",
                    action.path,
                )
            )
            action.target.value = _INVALID
        else:
            action.target.value = value
        return
    if isinstance(value, date):
        action.target.value = value
        return
    if isinstance(value, (set, frozenset)):
        diagnostics.append(
            _diagnostic(
                DiagnosticCode.UNORDERED_CONTEXT_COLLECTION,
                "sets and frozensets are not canonical ordered collections",
                action.path,
            )
        )
        action.target.value = _INVALID
        return
    if isinstance(value, Mapping):
        _visit_mapping(action, value, actions, diagnostics)
        return
    if isinstance(value, (list, tuple)):
        _visit_sequence(action, value, actions, diagnostics)
        return

    adapter = _resolve_adapter(value, adapters)
    if adapter is None:
        diagnostics.append(
            _diagnostic(
                DiagnosticCode.UNSUPPORTED_CONTEXT_VALUE,
                f"unsupported canonical value type: {_type_name(value)}",
                action.path,
            )
        )
        action.target.value = _INVALID
        return
    if isinstance(adapter, tuple):
        names = ", ".join(sorted(match.display_name for match in adapter))
        diagnostics.append(
            _diagnostic(
                DiagnosticCode.AMBIGUOUS_VALUE_ADAPTER,
                f"multiple unrelated value adapters match {_type_name(value)}: {names}",
                action.path,
            )
        )
        action.target.value = _INVALID
        return

    if _is_ancestor(value, action.ancestors) or adapter.source_type in action.adapter_chain:
        diagnostics.append(
            _diagnostic(
                DiagnosticCode.VALUE_ADAPTER_CYCLE,
                f"value adapter conversion formed a cycle for {_type_name(value)}",
                action.path,
            )
        )
        action.target.value = _INVALID
        return
    try:
        converted = adapter.convert(value)
    except Exception as error:
        detail = str(error)
        suffix = f": {detail}" if detail else ""
        diagnostics.append(
            _diagnostic(
                DiagnosticCode.VALUE_ADAPTER_FAILED,
                f"value adapter {adapter.display_name} failed with {type(error).__name__}{suffix}",
                action.path,
            )
        )
        action.target.value = _INVALID
        return

    actions.append(
        _Visit(
            converted,
            action.path,
            action.depth,
            (*action.ancestors, value),
            action.adapter_chain | {adapter.source_type},
            action.target,
        )
    )


def normalize_context(
    context: object,
    *,
    adapters: Iterable[TypeAdapter[Any]] = (),
    limits: ResourceLimits = DEFAULT_RESOURCE_LIMITS,
) -> NormalizationResult:
    """Adapt and immutably normalize a complete render context."""

    configured_adapters = tuple(adapters)
    diagnostics = list(_duplicate_adapter_diagnostics(configured_adapters))
    if not isinstance(context, Mapping):
        diagnostics.append(
            _diagnostic(
                DiagnosticCode.CONTEXT_MUST_BE_MAPPING,
                f"render context must be a mapping, not {_type_name(context)}",
                "context",
            )
        )
        return NormalizationResult(None, tuple(diagnostics))
    if diagnostics:
        return NormalizationResult(None, tuple(diagnostics))
    if isinstance(context, NormalizedContext):
        limit_diagnostic = _statistics_limit_diagnostic(context.statistics, limits)
        if limit_diagnostic is not None:
            return NormalizationResult(None, (limit_diagnostic,))
        return NormalizationResult(context, ())

    root = _Slot()
    state = _NormalizationState(limits)
    actions: list[_Action] = [_Visit(context, "context", 0, (), frozenset(), root)]
    while actions:
        action = actions.pop()
        if isinstance(action, _Visit):
            _visit_value(action, actions, diagnostics, configured_adapters, state)
            if state.limit_exceeded:
                break
        elif isinstance(action, _BuildMapping):
            _build_mapping(action)
        else:
            _build_sequence(action)

    if diagnostics:
        return NormalizationResult(None, tuple(diagnostics))
    if root.value is _UNSET or root.value is _INVALID or not isinstance(root.value, Mapping):
        raise AssertionError("valid mapping normalization did not produce a mapping")
    return NormalizationResult(
        NormalizedContext._create(
            cast(Mapping[str, CanonicalValue], root.value),
            state.statistics,
        ),
        (),
    )


def validate_context(
    context: object,
    *,
    limits: ResourceLimits = DEFAULT_RESOURCE_LIMITS,
) -> tuple[Diagnostic, ...]:
    """Validate a value tree using the canonical no-adapter boundary."""

    return normalize_context(context, limits=limits).diagnostics


def is_collection_value(value: object) -> bool:
    """Return whether a canonical value is non-scalar."""

    return isinstance(value, (Mapping, list, tuple))


def is_ordered_collection(value: object) -> bool:
    """Return whether a canonical value may drive a repeat."""

    return isinstance(value, (list, tuple))


__all__ = [
    "CanonicalContext",
    "CanonicalValue",
    "ContextStatistics",
    "InputValue",
    "NormalizationResult",
    "NormalizedContext",
    "ScalarValue",
    "TypeAdapter",
    "is_collection_value",
    "is_ordered_collection",
    "normalize_context",
    "validate_context",
]
