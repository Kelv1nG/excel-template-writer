"""Caller-scoped adapters for eager Polars data frames."""

from __future__ import annotations

import math
from collections.abc import Mapping
from typing import cast

import polars as pl

from excel_template_writer.values import TypeAdapter


def _contains_lossy_python_temporal(dtype: object) -> bool:
    """Return whether a Polars dtype loses temporal precision in Python.

    Args:
        dtype: Polars dtype, possibly nested in a list, array, or struct.

    Returns:
        ``True`` for nanosecond datetimes, times, or nested occurrences.
    """

    if isinstance(dtype, pl.Datetime):
        return dtype.time_unit == "ns"
    if dtype == pl.Time:
        return True
    if isinstance(dtype, (pl.List, pl.Array)):
        return _contains_lossy_python_temporal(dtype.inner)
    if isinstance(dtype, pl.Struct):
        return any(_contains_lossy_python_temporal(field.dtype) for field in dtype.fields)
    return False


def _contains_timezone_aware_datetime(dtype: object) -> bool:
    """Return whether a Polars dtype contains a timezone-aware datetime.

    Args:
        dtype: Polars dtype, possibly nested in a list, array, or struct.

    Returns:
        ``True`` when any contained datetime has a timezone.
    """

    if isinstance(dtype, pl.Datetime):
        return dtype.time_zone is not None
    if isinstance(dtype, (pl.List, pl.Array)):
        return _contains_timezone_aware_datetime(dtype.inner)
    if isinstance(dtype, pl.Struct):
        return any(_contains_timezone_aware_datetime(field.dtype) for field in dtype.fields)
    return False


def _replace_float_nans(value: object) -> object:
    """Copy materialized containers while replacing float NaNs with nulls.

    Args:
        value: Scalar or nested mapping/list/tuple materialized by Polars.

    Returns:
        A detached structure in which every float NaN is ``None``.
    """

    root_key = object()
    root: dict[object, object] = {root_key: None}
    stack: list[tuple[object, dict[object, object] | list[object], object]] = [
        (value, root, root_key)
    ]

    while stack:
        current, destination, key = stack.pop()

        if isinstance(current, float) and math.isnan(current):
            converted: object = None
        elif isinstance(current, Mapping):
            converted = {}
            stack.extend(
                (item, converted, item_key) for item_key, item in reversed(tuple(current.items()))
            )
        elif isinstance(current, (list, tuple)):
            converted = [None] * len(current)
            stack.extend(
                (item, converted, index) for index, item in reversed(tuple(enumerate(current)))
            )
        else:
            converted = current

        if isinstance(destination, list):
            destination[cast(int, key)] = converted
        else:
            destination[key] = converted

    return root[root_key]


def _convert_dataframe(frame: pl.DataFrame) -> object:
    """Convert an eager Polars frame into ordered canonical-ready records.

    Args:
        frame: Eager Polars data frame to materialize.

    Returns:
        A list of row dictionaries with NaNs converted to ``None``.

    Raises:
        ValueError: If columns are invalid or temporal values cannot be preserved.
    """

    columns = frame.columns
    if any(not isinstance(column, str) for column in columns):
        raise ValueError("Polars DataFrame columns must all be strings")
    if len(columns) != len(set(columns)):
        raise ValueError("Polars DataFrame column names must be unique")

    lossy_columns = [
        name for name, dtype in frame.schema.items() if _contains_lossy_python_temporal(dtype)
    ]
    if lossy_columns:
        names = ", ".join(repr(name) for name in lossy_columns)
        raise ValueError(
            "Polars nanosecond Datetime and Time values cannot be converted to Python "
            f"without precision loss; cast these columns first: {names}"
        )

    timezone_columns = [
        name for name, dtype in frame.schema.items() if _contains_timezone_aware_datetime(dtype)
    ]
    if timezone_columns:
        names = ", ".join(repr(name) for name in timezone_columns)
        raise ValueError(
            "timezone-aware Polars Datetime values are outside the canonical value model; "
            f"convert these columns to an explicit supported representation first: {names}"
        )

    return _replace_float_nans(frame.to_dicts())


def polars_adapters() -> tuple[TypeAdapter[pl.DataFrame], ...]:
    """Return explicit adapters for supported Polars runtime types.

    Returns:
        A one-item tuple containing the eager ``polars.DataFrame`` adapter.
    """

    return (TypeAdapter(pl.DataFrame, _convert_dataframe, name="polars.DataFrame"),)


__all__ = ["polars_adapters"]
