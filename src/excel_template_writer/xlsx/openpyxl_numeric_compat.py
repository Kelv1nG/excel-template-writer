"""Guarded, in-memory numeric serialization for the certified OpenPyXL release.

This process-global hook also affects independent OpenPyXL users. Activation and
its self-test are synchronized; subsequent workbook saves never take this lock.
"""

from __future__ import annotations

import math
from collections.abc import Callable
from io import BytesIO
from threading import Lock
from types import FunctionType

import openpyxl
import openpyxl.cell._writer as cell_writer
from openpyxl.compat.strings import safe_string as stock_formatter

from excel_template_writer.diagnostics import Diagnostic, DiagnosticCode, SourceLocation

_CERTIFIED_VERSION = "3.1.5"
_OWNER = "excel_template_writer.xlsx.openpyxl_numeric_compat"
_WRAPPER = "finite-builtin-float-repr-v1"
_TARGET = "openpyxl.cell._writer.safe_string"


class _OwnedNumericFormatter:
    """Retain the original formatter and stable reload-recognizable ownership."""

    owner = _OWNER
    wrapper = _WRAPPER
    target = _TARGET

    def __init__(self, original: Callable):
        """Capture the formatter used for every value outside the finite-float contract.

        Args:
            original: Unmodified OpenPyXL formatter to delegate to.
        """
        self.original = original

    def __call__(self, value: object):
        """Format a built-in finite float safely, delegating all other values.

        Args:
            value: OpenPyXL cell-writer value, including converted temporal serials.
        """
        if type(value) is float and math.isfinite(value):
            return repr(value)
        return self.original(value)


# A module reload must not reset a process-lifetime failure or replace its lock.
if "_lock" not in globals():
    _lock = Lock()
    _state = "UNINSTALLED"
    _active_wrapper: _OwnedNumericFormatter | None = None
    _failure: Diagnostic | None = None


def _owned_hook(hook: object) -> bool:
    """Recognize an earlier module incarnation without accepting foreign wrappers.

    Args:
        hook: Current cell-writer numeric formatter to inspect.
    """
    hook_type = type(hook)
    call = vars(hook_type).get("__call__")
    return (
        hook_type.__module__ == __name__
        and hook_type.__name__ == _OwnedNumericFormatter.__name__
        and getattr(hook, "owner", None) == _OWNER
        and getattr(hook, "wrapper", None) == _WRAPPER
        and getattr(hook, "target", None) == _TARGET
        and getattr(hook, "original", None) is stock_formatter
        and isinstance(call, FunctionType)
        and call.__code__ == _OwnedNumericFormatter.__call__.__code__
    )


def _validate_writer_hooks() -> None:
    """Reject cell writers that do not resolve the expected module-global hook."""
    for name in ("etree_write_cell", "lxml_write_cell"):
        function = getattr(cell_writer, name, None)
        if (
            not isinstance(function, FunctionType)
            or function.__module__ != cell_writer.__name__
            or function.__globals__ is not vars(cell_writer)
            or "safe_string" not in function.__code__.co_names
        ):
            raise RuntimeError(f"cell-writer hook shape check failed for {name}")


def _run_self_test() -> None:
    """Save and reopen the known formatter regression without the public renderer."""
    value = 100000.00000000001
    workbook = openpyxl.Workbook()
    stream = BytesIO()
    try:
        workbook.active["A1"] = value
        workbook.save(stream)
    finally:
        workbook.close()
    reopened = openpyxl.load_workbook(BytesIO(stream.getvalue()))
    try:
        actual = reopened.active["A1"].value
        if type(actual) is not float or actual.hex() != value.hex():
            raise RuntimeError("numeric save-and-reopen self-test did not preserve float bits")
    finally:
        reopened.close()


def _fail(code: DiagnosticCode, message: str) -> Diagnostic:
    """Cache a workbook-local failure for the remainder of the process.

    Args:
        code: Version or hook compatibility diagnostic category.
        message: Underlying check failure to retain on later calls.
    """
    global _state, _failure
    _failure = Diagnostic(code, message, SourceLocation("<workbook>", "A1"))
    _state = "FAILED"
    return _failure


def ensure_openpyxl_numeric_compatibility() -> Diagnostic | None:
    """Install the certified hook once, or return the process's sticky failure."""
    global _state, _active_wrapper
    if _state == "INSTALLED" and getattr(cell_writer, "safe_string", None) is _active_wrapper:
        return None
    with _lock:
        if _state == "FAILED":
            return _failure
        if _state == "INSTALLED":
            if getattr(cell_writer, "safe_string", None) is _active_wrapper:
                return None
            return _fail(
                DiagnosticCode.XLSX_OPENPYXL_COMPAT_FAILED,
                "numeric formatter ownership check failed: active hook was replaced",
            )
        _state = "INSTALLING"
        if openpyxl.__version__ != _CERTIFIED_VERSION:
            return _fail(
                DiagnosticCode.XLSX_OPENPYXL_COMPAT_UNVERIFIED,
                "OpenPyXL version is not certified for numeric serialization",
            )
        phase = "hook validation"
        try:
            _validate_writer_hooks()
            hook = getattr(cell_writer, "safe_string", None)
            if hook is stock_formatter:
                _active_wrapper = _OwnedNumericFormatter(stock_formatter)
                vars(cell_writer)["safe_string"] = _active_wrapper
            elif _owned_hook(hook):
                _active_wrapper = hook
            else:
                raise RuntimeError("numeric formatter hook is missing, foreign, or malformed")
            phase = "numeric self-test"
            _run_self_test()
            if cell_writer.safe_string is not _active_wrapper:
                raise RuntimeError("numeric formatter ownership changed during self-test")
        except Exception as exc:
            if (
                _active_wrapper is not None
                and getattr(cell_writer, "safe_string", None) is _active_wrapper
            ):
                cell_writer.safe_string = _active_wrapper.original
            return _fail(
                DiagnosticCode.XLSX_OPENPYXL_COMPAT_FAILED,
                f"OpenPyXL {phase} failed: {type(exc).__name__}: {exc}",
            )
        _state = "INSTALLED"
        return None
