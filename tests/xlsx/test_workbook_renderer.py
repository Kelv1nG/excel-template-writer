from __future__ import annotations

from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

from excel_template_writer.diagnostics import (
    DiagnosticCode,
    TemplateCompilationError,
    TemplateRenderError,
)
from excel_template_writer.xlsx import render_workbook


def _save(workbook: Workbook, path: Path) -> Path:
    workbook.save(path)
    return path


def test_render_workbook_copies_values_direct_styles_styled_blanks_and_dimensions(
    tmp_path: Path,
) -> None:
    template_path = tmp_path / "template.xlsx"
    output_path = tmp_path / "output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Report"
    sheet.column_dimensions["A"].width = 26
    sheet.column_dimensions["B"].width = 18
    sheet.column_dimensions["B"].hidden = True
    sheet.row_dimensions[2].height = 31
    sheet.row_dimensions[2].outlineLevel = 1
    sheet["A1"] = "Description"
    sheet["B1"] = "Amount"
    sheet["A2"] = "{% for line in lines %}{{ line.description }}"
    sheet["B2"] = "{{ line.amount }}{% endfor %}"
    sheet["A3"] = None
    sheet["B3"] = "Footer"

    thin = Side(style="thin", color="FF445566")
    for cell in sheet[2]:
        cell.font = Font(name="Aptos", bold=True, color="FF17365D")
        cell.fill = PatternFill("solid", fgColor="FFEAF3F8")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.protection = Protection(locked=False, hidden=True)
    sheet["B2"].number_format = "$#,##0.00;[Red]-$#,##0.00"

    # This cell has no value but is materially part of the workbook presentation.
    sheet["A3"].fill = PatternFill("solid", fgColor="FFFFF2CC")
    sheet["A3"].border = Border(bottom=thin)
    _save(workbook, template_path)

    result = render_workbook(
        template_path,
        output_path,
        {
            "lines": [
                {"description": "Consulting", "amount": 1250},
                {"description": "Support", "amount": 350},
            ]
        },
    )

    assert result.output_path == output_path
    assert result.diagnostics == ()
    template = load_workbook(template_path)
    rendered = load_workbook(output_path)
    try:
        source = template["Report"]
        target = rendered["Report"]
        assert source["A2"].value.startswith("{% for")
        assert target["A2"].value == "Consulting"
        assert target["A3"].value == "Support"
        assert target["B2"].value == 1250
        assert target["B3"].value == 350
        assert target["B4"].value == "Footer"
        assert target["A4"].value is None
        assert target["A4"].fill.fgColor.rgb == "FFFFF2CC"
        assert target["A4"].border.bottom.style == "thin"

        for coordinate in ("A2", "A3"):
            cell = target[coordinate]
            assert cell.font.bold is True
            assert cell.font.color.rgb == "FF17365D"
            assert cell.fill.fgColor.rgb == "FFEAF3F8"
            assert cell.border.left.style == "thin"
            assert cell.alignment.horizontal == "center"
            assert cell.alignment.wrap_text is True
            assert cell.protection.locked is False
            assert cell.protection.hidden is True
        assert target["B2"].number_format == "$#,##0.00;[Red]-$#,##0.00"
        assert target["B3"].number_format == "$#,##0.00;[Red]-$#,##0.00"
        assert target.row_dimensions[2].height == 31
        assert target.row_dimensions[3].height == 31
        assert target.row_dimensions[2].outlineLevel == 1
        assert target.row_dimensions[3].outlineLevel == 1
        assert target.column_dimensions["A"].width == 26
        assert target.column_dimensions["B"].width == 18
        assert target.column_dimensions["B"].hidden is True
    finally:
        template.close()
        rendered.close()


def test_render_workbook_repeats_contained_merges_and_row_heights(tmp_path: Path) -> None:
    template_path = tmp_path / "cards.xlsx"
    output_path = tmp_path / "cards-output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Cards"
    sheet.merge_cells("A1:C1")
    sheet["A1"] = "{% for card in cards %}{{ card.title }}"
    sheet["A2"] = "Owner"
    sheet["B2"] = "{{ card.owner }}"
    sheet["C2"] = "{{ card.amount }}{% endfor %}"
    sheet.row_dimensions[1].height = 28
    sheet.row_dimensions[2].height = 19
    sheet["A3"] = "After cards"
    sheet["A1"].fill = PatternFill("solid", fgColor="FF17365D")
    sheet["A1"].font = Font(bold=True, color="FFFFFFFF")
    sheet["C2"].number_format = "$#,##0"
    _save(workbook, template_path)

    render_workbook(
        template_path,
        output_path,
        {
            "cards": [
                {"title": "Alpha", "owner": "Mina", "amount": 100},
                {"title": "Beta", "owner": "Jules", "amount": 200},
            ]
        },
    )

    rendered = load_workbook(output_path)
    try:
        target = rendered["Cards"]
        assert {str(item) for item in target.merged_cells.ranges} == {"A1:C1", "A3:C3"}
        assert target["A1"].value == "Alpha"
        assert target["A3"].value == "Beta"
        assert target["A5"].value == "After cards"
        assert [target.row_dimensions[row].height for row in range(1, 5)] == [28, 19, 28, 19]
        assert target["A3"].fill.fgColor.rgb == "FF17365D"
        assert target["C4"].number_format == "$#,##0"
    finally:
        rendered.close()


def test_render_workbook_preserves_native_value_types(tmp_path: Path) -> None:
    template_path = tmp_path / "types.xlsx"
    output_path = tmp_path / "types-output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "{{ issued_on }}"
    sheet["A1"].number_format = "yyyy-mm-dd"
    sheet["B1"] = "{{ approved }}"
    sheet["C1"] = "{{ total }}"
    sheet["D1"] = "{{ exact_amount }}"
    sheet["D1"].number_format = "$#,##0.00"
    sheet["E1"] = "{{ cutoff }}"
    sheet["E1"].number_format = "hh:mm:ss"
    _save(workbook, template_path)

    render_workbook(
        template_path,
        output_path,
        {
            "issued_on": date(2026, 8, 13),
            "approved": True,
            "total": 42.5,
            "exact_amount": Decimal("19.75"),
            "cutoff": time(17, 30),
        },
    )

    rendered = load_workbook(output_path)
    try:
        sheet = rendered.active
        issued_on = sheet["A1"].value
        assert isinstance(issued_on, (date, datetime))
        assert sheet["A1"].data_type == "d"
        assert sheet["B1"].value is True
        assert sheet["B1"].data_type == "b"
        assert sheet["C1"].value == 42.5
        assert sheet["C1"].data_type == "n"
        assert sheet["D1"].value == 19.75
        assert sheet["D1"].data_type == "n"
        assert sheet["D1"].number_format == "$#,##0.00"
        assert sheet["E1"].value == time(17, 30)
        assert sheet["E1"].data_type == "d"
        assert sheet["E1"].number_format == "hh:mm:ss"
    finally:
        rendered.close()


def test_repeats_blank_cells_that_carry_formatting(tmp_path: Path) -> None:
    template_path = tmp_path / "styled-blanks.xlsx"
    output_path = tmp_path / "styled-blanks-output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "{% for item in items %}{{ item }}"
    sheet["B1"] = None
    sheet["B1"].fill = PatternFill("solid", fgColor="FFFFF2CC")
    sheet["B1"].border = Border(bottom=Side(style="double", color="FF000000"))
    sheet["C1"] = "{% endfor %}"
    _save(workbook, template_path)

    render_workbook(template_path, output_path, {"items": ["A", "B"]})

    rendered = load_workbook(output_path)
    try:
        sheet = rendered.active
        for coordinate in ("B1", "B2"):
            assert sheet[coordinate].value is None
            assert sheet[coordinate].fill.fgColor.rgb == "FFFFF2CC"
            assert sheet[coordinate].border.bottom.style == "double"
    finally:
        rendered.close()


@pytest.mark.parametrize("feature", ["formula", "conditional-formatting"])
def test_rejects_affected_unsupported_features(tmp_path: Path, feature: str) -> None:
    template_path = tmp_path / f"{feature}.xlsx"
    output_path = tmp_path / "should-not-exist.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "{% for item in items %}{{ item }}{% endfor %}"
    if feature == "formula":
        sheet["B2"] = "=SUM(A1:A1)"
        expected = DiagnosticCode.FORMULA_REQUIRES_UNSUPPORTED_TRANSFORM
    else:
        sheet.conditional_formatting.add(
            "A1:A3",
            CellIsRule(operator="greaterThan", formula=["0"]),
        )
        expected = DiagnosticCode.CONDITIONAL_FORMATTING_REQUIRES_UNSUPPORTED_TRANSFORM
    _save(workbook, template_path)

    with pytest.raises(TemplateCompilationError) as caught:
        render_workbook(template_path, output_path, {"items": [1, 2]})

    assert expected in {diagnostic.code for diagnostic in caught.value.diagnostics}
    assert not output_path.exists()


def test_rejects_merge_that_crosses_a_repeat_boundary(tmp_path: Path) -> None:
    template_path = tmp_path / "crossing-merge.xlsx"
    output_path = tmp_path / "should-not-exist.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "{% for item in items %}{{ item }}{% endfor %}"
    sheet.merge_cells("A1:B1")
    _save(workbook, template_path)

    with pytest.raises(TemplateCompilationError) as caught:
        render_workbook(template_path, output_path, {"items": [1, 2]})

    assert DiagnosticCode.MERGE_CROSSES_BLOCK_BOUNDARY in {
        diagnostic.code for diagnostic in caught.value.diagnostics
    }
    assert not output_path.exists()


def test_moves_static_merge_below_row_expansion(tmp_path: Path) -> None:
    template_path = tmp_path / "moving-merge.xlsx"
    output_path = tmp_path / "moving-merge-output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "{% for item in items %}{{ item }}{% endfor %}"
    sheet.merge_cells("A3:B3")
    sheet["A3"] = "Merged footer"
    _save(workbook, template_path)

    render_workbook(template_path, output_path, {"items": ["A", "B"]})

    rendered = load_workbook(output_path)
    try:
        sheet = rendered.active
        assert {str(item) for item in sheet.merged_cells.ranges} == {"A4:B4"}
        assert sheet["A4"].value == "Merged footer"
    finally:
        rendered.close()


def test_preserves_unaffected_formula_and_conditional_formatting(tmp_path: Path) -> None:
    template_path = tmp_path / "static-features.xlsx"
    output_path = tmp_path / "static-features-output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "=1+1"
    sheet["B1"] = "{{ label }}"
    sheet.conditional_formatting.add(
        "A1:A2",
        CellIsRule(operator="greaterThan", formula=["0"]),
    )
    _save(workbook, template_path)

    render_workbook(template_path, output_path, {"label": "Static layout"})

    rendered = load_workbook(output_path, data_only=False)
    try:
        sheet = rendered.active
        assert sheet["A1"].value == "=1+1"
        assert sheet["A1"].data_type == "f"
        assert sheet["B1"].value == "Static layout"
        assert len(sheet.conditional_formatting) == 1
    finally:
        rendered.close()


def test_runtime_string_that_starts_with_equals_remains_text(tmp_path: Path) -> None:
    template_path = tmp_path / "formula-injection.xlsx"
    output_path = tmp_path / "formula-injection-output.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "{{ supplied }}"
    _save(workbook, template_path)

    render_workbook(template_path, output_path, {"supplied": '=HYPERLINK("bad")'})

    rendered = load_workbook(output_path, data_only=False)
    try:
        assert rendered.active["A1"].value == '=HYPERLINK("bad")'
        assert rendered.active["A1"].data_type == "s"
    finally:
        rendered.close()


def test_rejects_custom_row_height_repeated_by_cell_shift(tmp_path: Path) -> None:
    template_path = tmp_path / "cell-shift-height.xlsx"
    output_path = tmp_path / "should-not-exist.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = '{% for item in items shift="cells" %}{{ item }}{% endfor %}'
    sheet.row_dimensions[1].height = 28
    _save(workbook, template_path)

    with pytest.raises(TemplateCompilationError) as caught:
        render_workbook(template_path, output_path, {"items": ["A", "B"]})

    assert DiagnosticCode.CELL_SHIFT_WITH_CUSTOM_ROW_HEIGHT in {
        diagnostic.code for diagnostic in caught.value.diagnostics
    }
    assert not output_path.exists()


def test_allows_custom_height_repeated_by_an_outer_row_shift(tmp_path: Path) -> None:
    template_path = tmp_path / "nested-height.xlsx"
    output_path = tmp_path / "nested-height-output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "{% for group in groups %}{{ group.name }}"
    sheet["A2"] = '{% for item in group.items shift="cells" %}{{ item }}{% endfor %}'
    sheet["B3"] = "{% endfor %}"
    sheet.row_dimensions[2].height = 28
    _save(workbook, template_path)

    render_workbook(
        template_path,
        output_path,
        {
            "groups": [
                {"name": "First", "items": ["A"]},
                {"name": "Second", "items": ["B"]},
            ]
        },
    )

    rendered = load_workbook(output_path)
    try:
        sheet = rendered.active
        assert sheet.row_dimensions[2].height == 28
        assert sheet.row_dimensions[5].height == 28
    finally:
        rendered.close()


def test_rejects_data_validation_on_a_transformed_sheet(tmp_path: Path) -> None:
    template_path = tmp_path / "validation.xlsx"
    output_path = tmp_path / "should-not-exist.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "{% for item in items %}{{ item }}{% endfor %}"
    validation = DataValidation(type="list", formula1='"Open,Closed"')
    validation.add("A1:A3")
    sheet.add_data_validation(validation)
    _save(workbook, template_path)

    with pytest.raises(TemplateCompilationError) as caught:
        render_workbook(template_path, output_path, {"items": ["A", "B"]})

    assert DiagnosticCode.DATA_VALIDATION_REQUIRES_UNSUPPORTED_TRANSFORM in {
        diagnostic.code for diagnostic in caught.value.diagnostics
    }


def test_rejects_native_excel_table_even_without_layout_changes(tmp_path: Path) -> None:
    template_path = tmp_path / "table.xlsx"
    output_path = tmp_path / "should-not-exist.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["Name"])
    sheet.append(["Alpha"])
    sheet.add_table(Table(displayName="ItemsTable", ref="A1:A2"))
    _save(workbook, template_path)

    with pytest.raises(TemplateCompilationError) as caught:
        render_workbook(template_path, output_path, {})

    assert DiagnosticCode.TABLE_REQUIRES_UNSUPPORTED_TRANSFORM in {
        diagnostic.code for diagnostic in caught.value.diagnostics
    }


def test_preserves_static_hyperlink_and_comment(tmp_path: Path) -> None:
    template_path = tmp_path / "links.xlsx"
    output_path = tmp_path / "links-output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "OpenAI"
    sheet["A1"].hyperlink = "https://openai.com"
    sheet["B1"] = "Reviewed"
    sheet["B1"].comment = Comment("Approved", "Reviewer")
    _save(workbook, template_path)

    render_workbook(template_path, output_path, {})

    rendered = load_workbook(output_path)
    try:
        sheet = rendered.active
        assert sheet["A1"].hyperlink.target == "https://openai.com"
        assert sheet["B1"].comment.text == "Approved"
        assert sheet["B1"].comment.author == "Reviewer"
    finally:
        rendered.close()


def test_rejects_input_path_as_output_path(tmp_path: Path) -> None:
    template_path = tmp_path / "same.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "{{ value }}"
    _save(workbook, template_path)

    with pytest.raises(TemplateCompilationError) as caught:
        render_workbook(template_path, template_path, {"value": "Rendered"})

    assert DiagnosticCode.INPUT_OUTPUT_PATH_CONFLICT in {
        diagnostic.code for diagnostic in caught.value.diagnostics
    }
    unchanged = load_workbook(template_path)
    try:
        assert unchanged.active["A1"].value == "{{ value }}"
    finally:
        unchanged.close()


def test_rejects_non_xlsx_paths_before_loading(tmp_path: Path) -> None:
    with pytest.raises(TemplateCompilationError) as caught:
        render_workbook(tmp_path / "template.xlsm", tmp_path / "output.xlsx", {})

    assert DiagnosticCode.UNSUPPORTED_WORKBOOK_FORMAT in {
        diagnostic.code for diagnostic in caught.value.diagnostics
    }


def test_rejects_noncanonical_context_before_writing(tmp_path: Path) -> None:
    template_path = tmp_path / "template.xlsx"
    output_path = tmp_path / "should-not-exist.xlsx"
    workbook = Workbook()
    workbook.active["A1"] = "{% for item in items %}{{ item }}{% endfor %}"
    _save(workbook, template_path)

    with pytest.raises(TemplateRenderError) as caught:
        render_workbook(template_path, output_path, {"items": {"A", "B"}})

    assert [diagnostic.code for diagnostic in caught.value.diagnostics] == [
        DiagnosticCode.UNORDERED_CONTEXT_COLLECTION
    ]
    assert str(caught.value.diagnostics[0].location) == "context.items"
    assert not output_path.exists()
