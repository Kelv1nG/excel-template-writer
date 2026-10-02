from __future__ import annotations

from decimal import Decimal
from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples"
SAMPLE_STEMS = (
    "scalar_values",
    "missing_output_policies",
    "repeated_blocks",
    "conditions_and_nesting",
    "cell_shift_lanes",
    "fixed_layout_fill",
    "default_cell_style",
    "fixed_range_charts",
    "template_images",
    "template_text_shapes",
    "regions",
    "polars_dataframe",
)


@pytest.mark.parametrize("stem", SAMPLE_STEMS)
def test_committed_sample_workbook_pairs_are_valid_and_rendered(stem: str) -> None:
    template_path = SAMPLES / f"{stem}_template.xlsx"
    output_path = SAMPLES / f"{stem}_output.xlsx"
    assert template_path.is_file()
    assert output_path.is_file()

    template = load_workbook(template_path, read_only=True, data_only=False)
    output = load_workbook(output_path, read_only=True, data_only=False)
    try:
        assert template.sheetnames == output.sheetnames
        template_text = [
            cell.value
            for sheet in template.worksheets
            for row in sheet.iter_rows()
            for cell in row
            if isinstance(cell.value, str)
        ]
        output_text = [
            cell.value
            for sheet in output.worksheets
            for row in sheet.iter_rows()
            for cell in row
            if isinstance(cell.value, str)
        ]
        assert any("{%" in value or "{{" in value for value in template_text)
        assert all("{%" not in value for value in output_text)
        if stem == "missing_output_policies":
            assert "{{  customer.email  }}" in output_text
            assert "Contact: {{  customer.email  }}" in output_text
        else:
            assert all("{{" not in value for value in output_text)
    finally:
        template.close()
        output.close()


def test_sample_catalog_names_every_committed_pair() -> None:
    catalog = (SAMPLES / "README.md").read_text(encoding="utf-8")

    for stem in SAMPLE_STEMS:
        assert f"{stem}_template.xlsx" in catalog
        assert f"{stem}_output.xlsx" in catalog


def test_scalar_values_sample_demonstrates_numeric_fidelity() -> None:
    path = SAMPLES / "scalar_values_output.xlsx"
    workbook = load_workbook(path)
    try:
        sheet = workbook["Scalar values"]
        cases = {row[0].value: row[1] for row in sheet.iter_rows(min_row=5)}
        float_cell = cases["Float bit fidelity"]
        assert float_cell.data_type == "n" and type(float_cell.value) is float
        assert float_cell.value.hex() == (100000.00000000001).hex()
        assert float_cell.number_format == "0.00000000000"
        with ZipFile(path) as archive:
            root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        xml_cells = {cell.attrib["r"]: cell for cell in root.findall(".//{*}c")}
        for label, amount, number_format in (
            ("Decimal", "1250.75", '$#,##0.00;[Red]-$#,##0.00;"-"'),
            ("15-digit financial Decimal", "1234567890123.45", "#,##0.00"),
            ("Decimal display scale", "12.50", "0.00"),
        ):
            cell = cases[label]
            assert cell.data_type == "n" and type(cell.value) is float
            assert Decimal(repr(cell.value)) == Decimal(amount)
            assert cell.number_format == number_format
            xml_cell = xml_cells[cell.coordinate]
            assert xml_cell.get("t") == "n"
            token = xml_cell.find("{*}v")
            assert token is not None and token.text is not None
            assert Decimal(token.text) == Decimal(amount)
    finally:
        workbook.close()


def test_polars_sample_demonstrates_computed_numeric_writeback() -> None:
    path = SAMPLES / "polars_dataframe_output.xlsx"
    workbook = load_workbook(path)
    try:
        sheet = workbook["Polars DataFrame"]
        columns = {cell.value: cell.column for cell in sheet[4]}
        rows = {row[0].value: row[0].row for row in sheet.iter_rows(min_row=5)}
        cell = sheet.cell(rows["Consulting"], columns["Computed amount"])
        assert cell.data_type == "n" and type(cell.value) is float
        assert Decimal(repr(cell.value)) == Decimal("1312.50")
        assert cell.number_format == '$#,##0.00;[Red]-$#,##0.00;"-"'
        with ZipFile(path) as archive:
            root = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
        xml_cell = root.find(f".//{{*}}c[@r='{cell.coordinate}']")
        assert xml_cell is not None and xml_cell.get("t") == "n"
        token = xml_cell.find("{*}v")
        assert token is not None and token.text is not None
        assert Decimal(token.text) == Decimal("1312.50")
    finally:
        workbook.close()
