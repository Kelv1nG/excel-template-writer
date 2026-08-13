"""Generate and render a visual workbook showcasing the current interpreter.

Run from the repository root with:

    uv run python scratch/demo.py

This is deliberately a throwaway adapter, not the production XLSX integration.
It compiles and plans every sheet before constructing a separate output workbook.
"""

from __future__ import annotations

import os
import tempfile
from collections.abc import Iterable
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, cast

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook as WorkbookType
from openpyxl.worksheet.cell_range import CellRange
from openpyxl.worksheet.worksheet import Worksheet

from excel_template_writer import Coordinate, WorksheetTemplate, compile_sheet, render_sheet
from excel_template_writer.render import RenderPlan

SCRATCH_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = SCRATCH_DIR / "demo_template.xlsx"
OUTPUT_PATH = SCRATCH_DIR / "demo_output.xlsx"

NAVY = "17365D"
LIGHT_BLUE = "EAF3F8"
GREEN = "E2F0D9"
GOLD = "FFF2CC"
GRAY = "E7E6E6"
WHITE = "FFFFFF"
THIN_GRAY = Side(style="thin", color="A6A6A6")


DEMO_CONTEXT: dict[str, Any] = {
    "report": {
        "name": "Quarterly revenue",
        "issued_on": date(2026, 8, 13),
        "approved": True,
        "total": 12_345.67,
    },
    "customer": {"name": "Acme Industries"},
    "labels": ["North", "Enterprise", "Renewal"],
    "lines": [
        {"description": "Consulting", "quantity": 2, "unit_price": 1_500.0},
        {"description": "Implementation", "quantity": 1, "unit_price": 4_250.0},
        {"description": "Support", "quantity": 6, "unit_price": 350.0},
    ],
    "cities": ["Singapore", "Tokyo", "Sydney"],
    "left_items": ["Left A", "Left B", "Left C"],
    "right_items": ["Right 1", "Right 2"],
    "empty_items": [],
    "cards": [
        {"title": "Launch", "owner": "Mina", "status": "On track", "budget": 75_000},
        {"title": "Renewal", "owner": "Jules", "status": "At risk", "budget": 42_500},
    ],
    "groups": [
        {"name": "Hardware", "items": ["Laptop", "Monitor"], "total": 2},
        {"name": "Software", "items": ["Editor", "Database", "Support"], "total": 3},
    ],
    "account": {"active": True, "overdue": False, "vip": False},
}


@dataclass(frozen=True)
class PlannedSheet:
    source: Worksheet
    plan: RenderPlan


def _atomic_save(workbook: WorkbookType, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    handle, temporary_name = tempfile.mkstemp(suffix=".xlsx", dir=path.parent)
    os.close(handle)
    temporary_path = Path(temporary_name)
    try:
        workbook.save(temporary_path)
        load_workbook(temporary_path, read_only=True, data_only=False).close()
        os.replace(temporary_path, path)
    finally:
        temporary_path.unlink(missing_ok=True)


def _new_sheet(workbook: WorkbookType, title: str, description: str) -> Worksheet:
    sheet = workbook.create_sheet(title)
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "A3"
    sheet.merge_cells("A1:D1")
    sheet["A1"] = title
    sheet["A1"].font = Font(name="Aptos Display", size=18, bold=True, color=WHITE)
    sheet["A1"].fill = PatternFill("solid", fgColor=NAVY)
    sheet["A1"].alignment = Alignment(vertical="center")
    sheet.row_dimensions[1].height = 30
    sheet.merge_cells("A2:D2")
    sheet["A2"] = description
    sheet["A2"].font = Font(name="Aptos", size=10, italic=True, color="595959")
    sheet["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    sheet.row_dimensions[2].height = 32
    for column, width in enumerate((28, 19, 19, 21), start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    return sheet


def _style_header(sheet: Worksheet, row: int, labels: tuple[str, ...]) -> None:
    for column, label in enumerate(labels, start=1):
        cell = sheet.cell(row, column, label)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(vertical="center")
        cell.border = Border(bottom=THIN_GRAY)
    sheet.row_dimensions[row].height = 22


def _style_block(
    sheet: Worksheet,
    min_row: int,
    max_row: int,
    *,
    fill: str = LIGHT_BLUE,
) -> None:
    for row in sheet.iter_rows(min_row=min_row, max_row=max_row, min_col=1, max_col=4):
        for cell in row:
            cell.fill = PatternFill("solid", fgColor=fill)
            cell.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    for row in range(min_row, max_row + 1):
        sheet.row_dimensions[row].height = 22


def _build_scalars(workbook: WorkbookType) -> None:
    sheet = _new_sheet(
        workbook,
        "Scalars",
        "Native values, mixed text, missing-value defaults, indexing, comparisons, and filters.",
    )
    _style_header(sheet, 3, ("Capability", "Template value", "Notes", "Expected type"))
    rows = [
        ("Plain string", "{{ report.name }}", "Sole expression", "string"),
        ("Mixed text", "Prepared for {{ customer.name }}", "Literal + expression", "string"),
        ("Date", "{{ report.issued_on }}", "Number format is copied", "date"),
        ("Boolean", "{{ report.approved }}", "Native Excel boolean", "boolean"),
        ("Number", "{{ report.total }}", "Native numeric value", "float"),
        (
            "Missing default",
            '{{ report.reference | default("-") }}',
            "Missing is handled",
            "string",
        ),
        ("Upper filter", "{{ customer.name | upper }}", "Pure text filter", "string"),
        ("Lower filter", "{{ customer.name | lower }}", "Pure text filter", "string"),
        ("Join filter", '{{ labels | join(" / ") }}', "Collection becomes text", "string"),
        ("Index lookup", "{{ labels[0] }}", "Zero-based indexing", "string"),
        (
            "Boolean expression",
            "{{ report.approved and report.total >= 100 }}",
            "Comparison + and",
            "boolean",
        ),
    ]
    for row_index, values in enumerate(rows, start=4):
        for column, value in enumerate(values, start=1):
            sheet.cell(row_index, column, value)
        _style_block(sheet, row_index, row_index, fill=WHITE if row_index % 2 else LIGHT_BLUE)
    sheet["B6"].number_format = "yyyy-mm-dd"
    sheet["B8"].number_format = "$#,##0.00;[Red]-$#,##0.00"
    sheet.auto_filter.ref = f"A3:D{3 + len(rows)}"


def _build_row_repeat(workbook: WorkbookType) -> None:
    sheet = _new_sheet(
        workbook,
        "Row Repeat",
        "A styled rectangular table body repeats down and inserts complete rows.",
    )
    _style_header(sheet, 3, ("Description", "Quantity", "Unit price", "Type"))
    sheet["A4"] = "{% for line in lines %}{{ line.description }}"
    sheet["B4"] = "{{ line.quantity }}"
    sheet["C4"] = "{{ line.unit_price }}"
    sheet["D4"] = "Line{% endfor %}"
    _style_block(sheet, 4, 4)
    sheet["C4"].number_format = "$#,##0.00"
    sheet["A5"] = "This footer moves below every rendered row."
    sheet.merge_cells("A5:D5")
    _style_block(sheet, 5, 5, fill=GOLD)


def _build_list(workbook: WorkbookType) -> None:
    sheet = _new_sheet(
        workbook,
        "One-cell List",
        "Lists use the same repeat primitive; this block is exactly one cell.",
    )
    _style_header(sheet, 3, ("City", "Static neighbor", "", ""))
    sheet["A4"] = "{% for city in cities %}{{ city }}{% endfor %}"
    sheet["B4"] = "Moves only if below the block"
    _style_block(sheet, 4, 4)
    sheet["A5"] = "End of list"
    _style_block(sheet, 5, 5, fill=GOLD)


def _build_cell_shift(workbook: WorkbookType) -> None:
    sheet = _new_sheet(
        workbook,
        "Cell Shift",
        "Two side-by-side lists grow independently because both explicitly shift cells.",
    )
    _style_header(sheet, 3, ("Left list", "Stationary", "Right list", "Stationary"))
    sheet["A4"] = '{% for item in left_items shift="cells" %}{{ item }}{% endfor %}'
    sheet["B4"] = "Beside left"
    sheet["C4"] = '{% for item in right_items shift="cells" %}{{ item }}{% endfor %}'
    sheet["D4"] = "Beside right"
    _style_block(sheet, 4, 4)
    sheet["A5"] = "After left lane"
    sheet["B5"] = "Does not move"
    sheet["C5"] = "After right lane"
    sheet["D5"] = "Does not move"
    _style_block(sheet, 5, 5, fill=GOLD)


def _build_empty_repeat(workbook: WorkbookType) -> None:
    sheet = _new_sheet(
        workbook,
        "Empty Repeat",
        "An empty collection keeps one formatted body row; item-dependent values become blank.",
    )
    _style_header(sheet, 3, ("Description", "Quantity", "State", "Amount"))
    sheet["A4"] = "{% for item in empty_items %}{{ item.description }}"
    sheet["B4"] = "{{ item.quantity }}"
    sheet["C4"] = "Formatted placeholder retained"
    sheet["D4"] = "{{ item.amount }}{% endfor %}"
    _style_block(sheet, 4, 4, fill=GREEN)
    sheet["D4"].number_format = "$#,##0.00"
    sheet["A5"] = "The footer stays directly after the single placeholder row."
    sheet.merge_cells("A5:D5")
    _style_block(sheet, 5, 5, fill=GOLD)


def _build_cards(workbook: WorkbookType) -> None:
    sheet = _new_sheet(
        workbook,
        "Repeated Cards",
        "A two-row styled block repeats. Its contained title merge is copied for each instance.",
    )
    sheet.merge_cells("A4:D4")
    sheet["A4"] = "{% for card in cards %}{{ card.title }}"
    sheet["A5"] = "Owner"
    sheet["B5"] = "{{ card.owner }}"
    sheet["C5"] = "{{ card.status }}"
    sheet["D5"] = "{{ card.budget }}{% endfor %}"
    _style_block(sheet, 4, 5)
    sheet["A4"].font = Font(size=14, bold=True, color=WHITE)
    sheet["A4"].fill = PatternFill("solid", fgColor=NAVY)
    sheet.row_dimensions[4].height = 26
    sheet["D5"].number_format = "$#,##0"
    sheet["A6"] = "End of cards"
    sheet.merge_cells("A6:D6")
    _style_block(sheet, 6, 6, fill=GOLD)


def _build_nested(workbook: WorkbookType) -> None:
    sheet = _new_sheet(
        workbook,
        "Nested Repeats",
        "An outer group block contains an inner item block and receives its measured height.",
    )
    _style_header(sheet, 3, ("Group / item", "Kind", "", "Count"))
    sheet["A4"] = "{% for group in groups %}{{ group.name }}"
    sheet["B4"] = "Group"
    sheet["A5"] = "{% for item in group.items %}{{ item }}"
    sheet["B5"] = "Item{% endfor %}"
    sheet["A6"] = "Group end"
    sheet["D6"] = "{{ group.total }}{% endfor %}"
    _style_block(sheet, 4, 6)
    sheet["A4"].font = Font(bold=True, color=WHITE)
    sheet["A4"].fill = PatternFill("solid", fgColor=NAVY)
    sheet["B4"].font = Font(bold=True, color=WHITE)
    sheet["B4"].fill = PatternFill("solid", fgColor=NAVY)
    sheet["A7"] = "End of nested groups"
    sheet.merge_cells("A7:D7")
    _style_block(sheet, 7, 7, fill=GOLD)


def _build_conditions(workbook: WorkbookType) -> None:
    sheet = _new_sheet(
        workbook,
        "Conditions",
        "Stacked equal-width branches compact vertically; the last block has no else branch.",
    )
    sheet["A4"] = "{% if account.active %}Account"
    sheet["D4"] = "ACTIVE{% else %}"
    sheet["A5"] = "Account"
    sheet["D5"] = "INACTIVE{% endif %}"
    sheet["A6"] = "{% if account.overdue %}Payment"
    sheet["D6"] = "OVERDUE{% else %}"
    sheet["A7"] = "Payment"
    sheet["D7"] = "CURRENT{% endif %}"
    sheet["A8"] = "{% if account.vip %}VIP-only row"
    sheet["D8"] = "Visible{% endif %}"
    for row in range(4, 9):
        _style_block(sheet, row, row, fill=GREEN if row % 2 == 0 else GRAY)
    sheet["A9"] = "Footer compacts upward after branch selection."
    sheet.merge_cells("A9:D9")
    _style_block(sheet, 9, 9, fill=GOLD)


def _build_notes(workbook: WorkbookType) -> None:
    sheet = _new_sheet(
        workbook,
        "Current Boundaries",
        "Capabilities intentionally not simulated by this Phase 0 demonstration.",
    )
    _style_header(sheet, 3, ("Not demonstrated", "Reason", "Planned boundary", ""))
    rows = [
        ("Horizontal repeats", "Only direction=down is executable", "Later language phase", ""),
        ("Excel formulas", "Copy/translation policy is unresolved", "Dedicated formula pass", ""),
        ("Native Excel Tables", "Table objects are not the current table model", "Future", ""),
        ("Images and drawings", "No preservation contract yet", "Future", ""),
        ("Production XLSX adapter", "This file uses a scratch adapter", "Next milestone", ""),
    ]
    for row_index, values in enumerate(rows, start=4):
        for column, value in enumerate(values, start=1):
            sheet.cell(row_index, column, value)
        _style_block(sheet, row_index, row_index, fill=WHITE if row_index % 2 else GRAY)


def build_demo_template(path: Path = TEMPLATE_PATH) -> Path:
    """Create the authored workbook containing visible template tags."""

    workbook = Workbook()
    workbook.remove(workbook.active)
    workbook.properties.title = "Excel Template Writer capability demonstration"
    workbook.properties.subject = "Authored template with visible spatial tags"
    workbook.properties.creator = "excel-template-writer scratch demo"
    _build_scalars(workbook)
    _build_row_repeat(workbook)
    _build_list(workbook)
    _build_cell_shift(workbook)
    _build_empty_repeat(workbook)
    _build_cards(workbook)
    _build_nested(workbook)
    _build_conditions(workbook)
    _build_notes(workbook)
    _atomic_save(workbook, path)
    return path


def _read_sheet(source: Worksheet) -> WorksheetTemplate:
    formulas = [
        cell.coordinate for row in source.iter_rows() for cell in row if cell.data_type == "f"
    ]
    if formulas:
        raise ValueError(
            f"{source.title}: formulas are outside this demo's supported boundary: {formulas}"
        )
    cells = {
        Coordinate(cell.row, cell.column): cell.value
        for row in source.iter_rows()
        for cell in row
        if cell.value is not None
    }
    return WorksheetTemplate(source.title, cells)


def _compile_and_plan(source_workbook: WorkbookType, context: dict[str, Any]) -> list[PlannedSheet]:
    planned: list[PlannedSheet] = []
    for source in source_workbook.worksheets:
        compiled = compile_sheet(_read_sheet(source)).require()
        plan = render_sheet(compiled, context).require()
        planned.append(PlannedSheet(source, plan))
    return planned


def _copy_cell(source: Cell, destination: Cell, value: Any) -> None:
    destination.value = value
    if isinstance(value, str) and value.startswith("="):
        destination.data_type = "s"
    if source.has_style:
        destination.font = copy(source.font)
        destination.fill = copy(source.fill)
        destination.border = copy(source.border)
        destination.alignment = copy(source.alignment)
        destination.protection = copy(source.protection)
        destination.number_format = source.number_format
    if source.hyperlink is not None:
        destination.hyperlink = copy(source.hyperlink)


def _copy_dimensions(source: Worksheet, destination: Worksheet, plan: RenderPlan) -> None:
    for column in range(1, plan.width + 1):
        letter = get_column_letter(column)
        source_dimension = source.column_dimensions[letter]
        destination_dimension = destination.column_dimensions[letter]
        destination_dimension.width = source_dimension.width
        destination_dimension.hidden = source_dimension.hidden

    row_sources: dict[int, int] = {}
    for planned_cell in plan.cells:
        row_sources.setdefault(planned_cell.coordinate.row, planned_cell.source_coordinate.row)
    for destination_row, source_row in row_sources.items():
        source_dimension = source.row_dimensions[source_row]
        destination_dimension = destination.row_dimensions[destination_row]
        destination_dimension.height = source_dimension.height
        destination_dimension.hidden = source_dimension.hidden


def _copy_merges(source: Worksheet, destination: Worksheet, plan: RenderPlan) -> None:
    merged_ranges = cast(Iterable[CellRange], source.merged_cells.ranges)
    for merged in merged_ranges:
        min_row = cast(int, merged.min_row)
        min_column = cast(int, merged.min_col)
        max_row = cast(int, merged.max_row)
        max_column = cast(int, merged.max_col)
        top_left = Coordinate(min_row, min_column)
        occurrences = [cell.coordinate for cell in plan.cells if cell.source_coordinate == top_left]
        for occurrence in occurrences:
            destination.merge_cells(
                start_row=occurrence.row,
                start_column=occurrence.column,
                end_row=occurrence.row + max_row - min_row,
                end_column=occurrence.column + max_column - min_column,
            )


def _write_plans(planned_sheets: list[PlannedSheet], path: Path) -> None:
    output = Workbook()
    output.remove(output.active)
    output.properties.title = "Rendered Excel Template Writer demonstration"
    output.properties.subject = "Output produced from a validated render plan"
    output.properties.creator = "excel-template-writer scratch demo"

    for planned_sheet in planned_sheets:
        source = planned_sheet.source
        plan = planned_sheet.plan
        destination = output.create_sheet(source.title)
        destination.sheet_view.showGridLines = source.sheet_view.showGridLines
        destination.freeze_panes = source.freeze_panes
        if source.sheet_properties.tabColor is not None:
            destination.sheet_properties.tabColor = copy(source.sheet_properties.tabColor)
        _copy_dimensions(source, destination, plan)
        for planned_cell in plan.cells:
            source_cell = source[planned_cell.source_coordinate.a1]
            destination_cell = destination[planned_cell.coordinate.a1]
            _copy_cell(source_cell, destination_cell, planned_cell.value)
        _copy_merges(source, destination, plan)

    _atomic_save(output, path)


def _validate_output(template_path: Path, output_path: Path) -> None:
    template = load_workbook(template_path, data_only=False)
    output = load_workbook(output_path, data_only=False)
    try:
        assert template.sheetnames == output.sheetnames
        assert any(
            "{%" in cell.value or "{{" in cell.value
            for sheet in template.worksheets
            for row in sheet.iter_rows()
            for cell in row
            if isinstance(cell.value, str)
        )
        for sheet in output.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str):
                        assert "{%" not in cell.value and "{{" not in cell.value

        rendered_date = output["Scalars"]["B6"].value
        assert isinstance(rendered_date, (date, datetime))
        normalized_date = (
            rendered_date.date() if isinstance(rendered_date, datetime) else rendered_date
        )
        assert normalized_date == date(2026, 8, 13)
        assert output["Scalars"]["B7"].value is True
        assert output["Scalars"]["B8"].value == 12_345.67
        assert output["Row Repeat"]["A4"].value == "Consulting"
        assert output["Row Repeat"]["A6"].value == "Support"
        assert output["Row Repeat"]["A7"].value.startswith("This footer")
        assert output["Cell Shift"]["A6"].value == "Left C"
        assert output["Cell Shift"]["B5"].value == "Does not move"
        assert output["Empty Repeat"]["A4"].value is None
        assert output["Empty Repeat"]["C4"].value == "Formatted placeholder retained"
        assert "A4:D4" in output["Repeated Cards"].merged_cells
        assert "A6:D6" in output["Repeated Cards"].merged_cells
        assert output["Nested Repeats"]["A5"].value == "Laptop"
        assert output["Conditions"]["D4"].value == "ACTIVE"
        assert output["Conditions"]["D5"].value == "CURRENT"
        assert output["Conditions"]["A6"].value.startswith("Footer")
        assert output["Row Repeat"]["A4"].fill.fgColor.rgb.endswith(LIGHT_BLUE)
    finally:
        template.close()
        output.close()


def render_demo(
    template_path: Path = TEMPLATE_PATH,
    output_path: Path = OUTPUT_PATH,
    context: dict[str, Any] | None = None,
) -> Path:
    """Compile all sheets, then write and reload a separate output workbook."""

    source = load_workbook(template_path, data_only=False, keep_links=False)
    try:
        planned = _compile_and_plan(source, DEMO_CONTEXT if context is None else context)
        _write_plans(planned, output_path)
    finally:
        source.close()
    _validate_output(template_path, output_path)
    return output_path


def main() -> None:
    template_path = build_demo_template()
    output_path = render_demo(template_path)
    print(f"Template: {template_path}")
    print(f"Rendered: {output_path}")
    print("Validated: tags removed, native types retained, styles and demo merges preserved.")


if __name__ == "__main__":
    main()
