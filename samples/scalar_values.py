"""Generate the maintained scalar values and expressions sample workbooks."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_template_writer.xlsx import render_workbook
from samples._common import (
    LIGHT_BLUE,
    LIGHT_GOLD,
    SAMPLES_DIR,
    assert_no_template_tags,
    atomic_save,
    paint,
    prepare_sheet,
    sample_paths,
)

TEMPLATE_PATH, OUTPUT_PATH = sample_paths("scalar_values")


def build_template(path: Path = TEMPLATE_PATH) -> Path:
    """Build a template demonstrating typed cells, mixed text, and filters.

    Args:
        path: Destination path for the authored template workbook.

    Returns:
        The verified template path.
    """

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Scalar values"
    prepare_sheet(
        sheet,
        "Scalar values and expressions",
        "Single-expression cells preserve types; mixed content becomes text; filters are explicit.",
        widths=(22, 34, 25),
    )
    sheet.append(["Case", "Template value", "Expected behavior"])
    paint(sheet, "A4:C4", fill=LIGHT_BLUE, bold=True, horizontal="center")
    rows = (
        ("Text", "{{ customer.name }}", "native text"),
        ("Integer", "{{ invoice.quantity }}", "native number"),
        ("Decimal", "{{ invoice.amount }}", "currency-formatted number"),
        ("Date", "{{ invoice.issued_on }}", "native date"),
        ("Boolean", "{{ invoice.approved }}", "native boolean"),
        ("Mixed text", "Invoice {{ invoice.number }}", "always text"),
        ("Upper filter", "{{ customer.name | upper }}", "uppercase text"),
        ("Join filter", '{{ labels | join(" / ") }}', "one text value"),
        ("Default filter", '{{ missing | default("not supplied") }}', "explicit fallback"),
    )
    for row_number, values in enumerate(rows, start=5):
        for column, value in enumerate(values, start=1):
            sheet.cell(row_number, column, value)
        paint(
            sheet, f"A{row_number}:C{row_number}", fill=LIGHT_GOLD if row_number % 2 else "FFFFFFFF"
        )
    sheet["B7"].number_format = '$#,##0.00;[Red]-$#,##0.00;"-"'
    sheet["B8"].number_format = "yyyy-mm-dd"
    result = atomic_save(workbook, path)
    workbook.close()
    return result


def render_sample(
    template_path: Path = TEMPLATE_PATH,
    output_path: Path = OUTPUT_PATH,
) -> Path:
    """Render and verify the scalar sample workbook.

    Args:
        template_path: Authored sample template path.
        output_path: Separate rendered workbook path.

    Returns:
        The verified rendered workbook path.
    """

    render_workbook(
        template_path,
        output_path,
        {
            "customer": {"name": "Acme Industries"},
            "invoice": {
                "number": "INV-1042",
                "quantity": 3,
                "amount": 1250.75,
                "issued_on": date(2026, 8, 19),
                "approved": True,
            },
            "labels": ["priority", "renewal"],
        },
    )
    assert_no_template_tags(output_path)
    workbook = load_workbook(output_path)
    try:
        sheet = workbook["Scalar values"]
        assert sheet["B5"].value == "Acme Industries"
        assert sheet["B6"].value == 3
        assert sheet["B7"].value == 1250.75
        assert sheet["B8"].data_type == "d"
        assert sheet["B9"].value is True
        assert sheet["B10"].value == "Invoice INV-1042"
        assert sheet["B13"].value == "not supplied"
    finally:
        workbook.close()
    return output_path


def main() -> None:
    """Generate both scalar sample workbooks and print their paths."""

    template = build_template()
    output = render_sample(template)
    print(f"Template: {template.relative_to(SAMPLES_DIR.parent)}")
    print(f"Output:   {output.relative_to(SAMPLES_DIR.parent)}")


if __name__ == "__main__":
    main()
