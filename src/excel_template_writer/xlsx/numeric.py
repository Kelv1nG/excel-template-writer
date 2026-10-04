"""Pure numeric adaptation of completed render plans at the XLSX boundary."""

from __future__ import annotations

import math
import sys
from dataclasses import dataclass, replace
from decimal import Decimal

from excel_template_writer.diagnostics import (
    Diagnostic,
    DiagnosticCode,
    SourceLocation,
    TemplateRenderError,
)
from excel_template_writer.render import PlannedCell, RenderPlan


@dataclass(frozen=True)
class NumericPreflightResult:
    plans: tuple[RenderPlan, ...] | None
    diagnostics: tuple[Diagnostic, ...]

    def require(self) -> tuple[RenderPlan, ...]:
        """Return adapted plans or raise all numeric diagnostics.

        Returns:
            Immutable adapter-local plans with supported Decimal cells converted.

        Raises:
            TemplateRenderError: If any Decimal cell is unsupported.
        """

        if self.plans is None:
            raise TemplateRenderError(self.diagnostics)
        return self.plans


def preflight_numeric_plans(plans: tuple[RenderPlan, ...]) -> NumericPreflightResult:
    """Validate final Decimal cells and copy plans without changing the layout IR.

    Args:
        plans: Complete plans in deterministic worksheet and cell order. Values
            have already passed canonical and expression non-finite validation.

    Returns:
        Adapted plans, or all unsupported Decimal destination diagnostics and no plans.
    """

    adapted_plans: list[RenderPlan] = []
    diagnostics: list[Diagnostic] = []
    for plan in plans:
        cells: list[PlannedCell] = []
        for cell in plan.cells:
            value = cell.value
            if not isinstance(value, Decimal):
                cells.append(cell)
                continue

            digits = value.as_tuple().digits
            precision = len(digits)
            while precision > 1 and digits[precision - 1] == 0:
                precision -= 1
            location = SourceLocation(plan.sheet, cell.coordinate.a1)
            if precision > 15:
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.XLSX_DECIMAL_PRECISION_EXCEEDED,
                        "Decimal numeric cell exceeds 15 significant digits",
                        location,
                    )
                )
                continue

            converted = float(value)
            if not math.isfinite(converted) or (value != 0 and abs(converted) < sys.float_info.min):
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.XLSX_DECIMAL_OUT_OF_RANGE,
                        "Decimal numeric cell overflows, underflows, or becomes subnormal",
                        location,
                    )
                )
                continue
            if Decimal(repr(converted)) != value:
                diagnostics.append(
                    Diagnostic(
                        DiagnosticCode.XLSX_DECIMAL_INEXACT,
                        "Decimal numeric cell differs from its round-trip-safe float token",
                        location,
                    )
                )
                continue
            cells.append(replace(cell, value=converted))
        adapted_plans.append(replace(plan, cells=tuple(cells)))

    if diagnostics:
        return NumericPreflightResult(None, tuple(diagnostics))
    return NumericPreflightResult(tuple(adapted_plans), ())


__all__ = ["NumericPreflightResult", "preflight_numeric_plans"]
