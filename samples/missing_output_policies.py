"""Generate maintained samples for missing output-expression policies."""

from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook

from excel_template_writer.diagnostics import (
    Diagnostic,
    DiagnosticCode,
    DiagnosticSeverity,
    TemplateRenderError,
)
from excel_template_writer.xlsx import render_workbook
from samples._common import (
    LIGHT_BLUE,
    LIGHT_GOLD,
    LIGHT_GREEN,
    SAMPLES_DIR,
    atomic_save,
    paint,
    prepare_sheet,
    sample_paths,
)

TEMPLATE_PATH, OUTPUT_PATH = sample_paths("missing_output_policies")
EXACT_MISSING_TAG = "{{  customer.email  }}"
MIXED_MISSING_VALUE = f"Contact: {EXACT_MISSING_TAG}"


def build_template(path: Path = TEMPLATE_PATH) -> Path:
    """Build a template showing strict, blank, and preserve missing-output policies.

    Args:
        path: Destination path for the authored template workbook.

    Returns:
        The verified template path.
    """

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Missing outputs"
    prepare_sheet(
        sheet,
        "Missing output policies",
        "Strict mode rejects absent output values; blank and preserve modes publish warnings.",
        widths=(28, 42, 54),
    )
    for column, label in enumerate(
        ("Case", "Template value", "Expected preserve-mode result"),
        start=1,
    ):
        sheet.cell(4, column, label)
    paint(sheet, "A4:C4", fill=LIGHT_BLUE, bold=True, horizontal="center")
    rows = (
        ("Provided value", "{{ customer.name }}", "Acme Industries"),
        ("Missing sole expression", EXACT_MISSING_TAG, "Exact authored tag, including spaces"),
        ("Missing in mixed text", MIXED_MISSING_VALUE, "Literal prefix plus exact authored tag"),
        (
            "Explicit default",
            '{{ customer.phone | default("Not supplied") }}',
            "Not supplied; no warning",
        ),
        ("Present null", "{{ customer.middle_name }}", "Blank; null is not missing"),
    )
    for row_number, values in enumerate(rows, start=5):
        for column, value in enumerate(values, start=1):
            sheet.cell(row_number, column, value)
        paint(
            sheet,
            f"A{row_number}:C{row_number}",
            fill=LIGHT_GOLD if row_number % 2 else "FFFFFFFF",
        )
    paint(sheet, "B6:B7", fill=LIGHT_GREEN)
    sheet.row_dimensions[7].height = 30
    result = atomic_save(workbook, path)
    workbook.close()
    return result


def _context() -> dict[str, object]:
    """Return deliberately incomplete input shared by all three policy renders."""

    return {
        "customer": {
            "name": "Acme Industries",
            "middle_name": None,
        }
    }


def _assert_missing_diagnostics(
    diagnostics: tuple[Diagnostic, ...],
    *,
    code: DiagnosticCode,
    severity: DiagnosticSeverity,
) -> None:
    """Assert the two missing-email occurrences retain locations and severity.

    Args:
        diagnostics: Render diagnostics to verify.
        code: Expected stable missing-value diagnostic code.
        severity: Expected diagnostic severity for the active policy.
    """

    assert [diagnostic.code for diagnostic in diagnostics] == [code, code]
    assert all(diagnostic.severity is severity for diagnostic in diagnostics)
    assert [str(diagnostic.location) for diagnostic in diagnostics] == [
        "Missing outputs!B6:0",
        "Missing outputs!B7:9",
    ]
    assert all("customer.email" in diagnostic.message for diagnostic in diagnostics)


def _assert_common_values(path: Path) -> None:
    """Reopen one rendered workbook and verify non-missing value behavior.

    Args:
        path: Rendered workbook path to reopen.
    """

    workbook = load_workbook(path, data_only=False)
    try:
        sheet = workbook["Missing outputs"]
        assert sheet["B5"].value == "Acme Industries"
        assert sheet["B8"].value == "Not supplied"
        assert sheet["B9"].value is None
    finally:
        workbook.close()


def _verify_strict_mode(template_path: Path, output_path: Path) -> None:
    """Verify strict mode reports errors and publishes no workbook.

    Args:
        template_path: Authored workbook containing missing output expressions.
        output_path: Temporary path that must remain unpublished.
    """

    try:
        render_workbook(
            template_path,
            output_path,
            _context(),
            missing_output="error",
        )
    except TemplateRenderError as error:
        _assert_missing_diagnostics(
            error.diagnostics,
            code=DiagnosticCode.MISSING_VALUE,
            severity=DiagnosticSeverity.ERROR,
        )
    else:
        raise AssertionError("strict missing-output mode must reject the render")
    assert not output_path.exists()


def _verify_blank_mode(template_path: Path, output_path: Path) -> None:
    """Verify blank mode publishes formatted blanks and warning diagnostics.

    Args:
        template_path: Authored workbook containing missing output expressions.
        output_path: Temporary blank-mode output path.
    """

    result = render_workbook(
        template_path,
        output_path,
        _context(),
        missing_output="blank",
    )
    _assert_missing_diagnostics(
        result.diagnostics,
        code=DiagnosticCode.MISSING_VALUE_RENDERED,
        severity=DiagnosticSeverity.WARNING,
    )
    _assert_common_values(output_path)
    workbook = load_workbook(output_path, data_only=False)
    try:
        sheet = workbook["Missing outputs"]
        assert sheet["B6"].value is None
        assert sheet["B6"].fill.fgColor.rgb == LIGHT_GREEN
        assert sheet["B7"].value == "Contact: "
    finally:
        workbook.close()


def render_sample(
    template_path: Path = TEMPLATE_PATH,
    output_path: Path = OUTPUT_PATH,
) -> Path:
    """Render and verify strict, blank, and preserve policies.

    The committed output workbook uses preserve mode so its unresolved expression
    tags remain directly inspectable. Strict and blank outputs are temporary.

    Args:
        template_path: Authored sample template path.
        output_path: Separate preserve-mode output workbook path.

    Returns:
        The verified preserve-mode output path.
    """

    with TemporaryDirectory() as temporary_directory:
        temporary_path = Path(temporary_directory)
        _verify_strict_mode(template_path, temporary_path / "strict.xlsx")
        _verify_blank_mode(template_path, temporary_path / "blank.xlsx")

    result = render_workbook(
        template_path,
        output_path,
        _context(),
        missing_output="preserve",
    )
    _assert_missing_diagnostics(
        result.diagnostics,
        code=DiagnosticCode.MISSING_VALUE_RENDERED,
        severity=DiagnosticSeverity.WARNING,
    )
    _assert_common_values(output_path)
    workbook = load_workbook(output_path, data_only=False)
    try:
        sheet = workbook["Missing outputs"]
        assert sheet["B6"].value == EXACT_MISSING_TAG
        assert sheet["B6"].fill.fgColor.rgb == LIGHT_GREEN
        assert sheet["B7"].value == MIXED_MISSING_VALUE
    finally:
        workbook.close()
    return output_path


def main() -> None:
    """Generate the missing-output template and preserve-mode output workbooks."""

    template = build_template()
    output = render_sample(template)
    print(f"Template: {template.relative_to(SAMPLES_DIR.parent)}")
    print(f"Output:   {output.relative_to(SAMPLES_DIR.parent)}")


if __name__ == "__main__":
    main()
