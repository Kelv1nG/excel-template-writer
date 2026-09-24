import pytest

from excel_template_writer.ast import ForNode, IfNode, RegionNode
from excel_template_writer.compiler import compile_sheet
from excel_template_writer.diagnostics import DiagnosticCode
from excel_template_writer.model import Coordinate, Rectangle, WorksheetTemplate


def test_links_opposite_corners_into_a_rectangular_for_node() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            ["Description", "Amount"],
            ["{% for row in rows %}{{ row.name }}", "{{ row.amount }}{% endfor %}"],
        ],
    )

    compiled = compile_sheet(template).require()

    assert len(compiled.children) == 1
    node = compiled.children[0]
    assert isinstance(node, ForNode)
    assert node.rectangle == Rectangle(top=2, left=1, bottom=2, right=2)
    assert node.variable == "row"
    assert node.direction == "down"
    assert node.shift == "rows"


def test_accepts_fixed_no_shift_for_loop() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [['{% for row in rows shift="none" %}{{ row.name }}{% endfor %}']],
    )

    node = compile_sheet(template).require().children[0]

    assert isinstance(node, ForNode)
    assert node.shift == "none"


def test_rejects_structural_children_inside_fixed_no_shift_loop() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            ['{% for item in items shift="none" %}{{ item.name }}', None],
            ["{% if item.show %}Shown{% endif %}", "{% endfor %}"],
        ],
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.INVALID_BLOCK_GEOMETRY
    ]
    assert 'shift="none" repeat cannot contain structural blocks' in result.diagnostics[0].message


def test_links_stacked_equal_width_if_else_branches() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            ["{% if invoice.taxable %}Tax", "{{ invoice.tax }}{% else %}"],
            ["Not taxable", "{% endif %}"],
        ],
    )

    compiled = compile_sheet(template).require()
    node = compiled.children[0]

    assert isinstance(node, IfNode)
    assert node.rectangle == Rectangle(1, 1, 2, 2)
    assert node.true_rectangle == Rectangle(1, 1, 1, 2)
    assert node.false_rectangle == Rectangle(2, 1, 2, 2)


def test_links_explicit_region_and_its_nested_blocks() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {
            "A1": '{% region direction="down" shift="cells" %}',
            "A2": '{% for item in items shift="cells" %}{{ item }}',
            "C2": "{% endfor %}",
            "J2": "{% endregion %}",
        },
    )

    compiled = compile_sheet(template).require()

    assert len(compiled.children) == 1
    region = compiled.children[0]
    assert isinstance(region, RegionNode)
    assert region.rectangle == Rectangle(1, 1, 2, 10)
    assert region.direction == "down"
    assert region.shift == "cells"
    assert len(region.children) == 1
    assert isinstance(region.children[0], ForNode)
    assert region.children[0].rectangle == Rectangle(2, 1, 2, 3)


def test_region_defaults_to_downward_row_shifting() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {"A1": "{% region %}", "B2": "{% endregion %}"},
    )

    region = compile_sheet(template).require().children[0]

    assert isinstance(region, RegionNode)
    assert region.direction == "down"
    assert region.shift == "rows"


def test_rejects_unsupported_region_direction() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {"A1": '{% region direction="right" %}', "B2": "{% endregion %}"},
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.INVALID_DIRECTIVE
    ]
    assert 'direction must be "down"' in result.diagnostics[0].message


@pytest.mark.parametrize(
    "options",
    [
        'direction="down" direction="down"',
        'shift="columns"',
        'shift="none"',
        'unknown="value"',
    ],
)
def test_rejects_invalid_region_options(options: str) -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {"A1": f"{{% region {options} %}}", "B2": "{% endregion %}"},
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.INVALID_DIRECTIVE
    ]


def test_rejects_merge_crossing_a_cell_shift_region_lane() -> None:
    template = WorksheetTemplate(
        "Report",
        {
            Coordinate(1, 1): '{% region shift="cells" %}',
            Coordinate(2, 10): "{% endregion %}",
            Coordinate(10, 10): "Crossing merge",
        },
        (Rectangle(10, 10, 10, 11),),
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert DiagnosticCode.MERGE_CROSSES_BLOCK_BOUNDARY in {
        diagnostic.code for diagnostic in result.diagnostics
    }


def test_nested_cell_shift_lane_does_not_claim_merges_below_its_region() -> None:
    template = WorksheetTemplate(
        "Report",
        {
            Coordinate(1, 1): '{% region shift="cells" %}',
            Coordinate(2, 1): '{% for item in items shift="cells" %}{{ item }}',
            Coordinate(2, 3): "{% endfor %}",
            Coordinate(2, 10): "{% endregion %}",
            Coordinate(10, 1): "Region footer",
        },
        (Rectangle(10, 1, 10, 10),),
    )

    compiled = compile_sheet(template).require()

    assert isinstance(compiled.children[0], RegionNode)


def test_rejects_partially_overlapping_regions() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            ["{% for row in rows %}", "{% if enabled %}", None],
            [None, "{% endfor %}", "{% endif %}"],
        ],
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert DiagnosticCode.PARTIAL_BLOCK_OVERLAP in {item.code for item in result.diagnostics}


def test_reports_ambiguous_else_markers_instead_of_guessing() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            ["{% if enabled %}", "{% else %}"],
            [None, "{% else %}"],
            [None, "{% endif %}"],
        ],
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert DiagnosticCode.AMBIGUOUS_BLOCK_PAIRING in {item.code for item in result.diagnostics}


def test_rejects_side_by_side_siblings_that_both_claim_row_insertion() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            [
                "{% for left in left_items %}{{ left }}{% endfor %}",
                "{% for right in right_items %}{{ right }}{% endfor %}",
            ]
        ],
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert DiagnosticCode.OVERLAPPING_ROW_SHIFTS in {item.code for item in result.diagnostics}


def test_allows_side_by_side_none_and_cell_shift_siblings() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            [
                '{% for left in left_items shift="none" %}{{ left }}{% endfor %}',
                None,
                '{% for right in right_items shift="cells" %}{{ right }}{% endfor %}',
            ]
        ],
    )

    compiled = compile_sheet(template).require()

    assert [node.shift for node in compiled.children if isinstance(node, ForNode)] == [
        "none",
        "cells",
    ]


def test_rejects_side_by_side_row_and_none_shift_siblings() -> None:
    template = WorksheetTemplate.from_rows(
        "Report",
        [
            [
                "{% for left in left_items %}{{ left }}{% endfor %}",
                '{% for right in right_items shift="none" %}{{ right }}{% endfor %}',
            ]
        ],
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.OVERLAPPING_ROW_SHIFTS
    ]
    assert "must not use whole-row shifting" in result.diagnostics[0].message


def test_reports_invalid_date_format_at_the_output_cell() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {"C7": '{{ report_date | date("{yyyy}-{mm}") }}'},
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.INVALID_DATE_FORMAT
    ]
    assert str(result.diagnostics[0].location) == "Report!C7:0"


def test_reports_unknown_filter_at_the_output_cell() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {"B3": "{{ value | invented }}"},
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [DiagnosticCode.INVALID_FILTER]
    assert str(result.diagnostics[0].location) == "Report!B3:0"


def test_reports_invalid_sum_column_at_the_output_cell() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {"E6": "{{ rows | sum(1) }}"},
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [DiagnosticCode.INVALID_FILTER]
    assert str(result.diagnostics[0].location) == "Report!E6:0"


def test_reports_invalid_arithmetic_syntax_at_the_output_cell() -> None:
    template = WorksheetTemplate.from_cells(
        "Report",
        {"F7": "{{ subtotal + }}"},
    )

    result = compile_sheet(template)

    assert result.compiled is None
    assert [diagnostic.code for diagnostic in result.diagnostics] == [
        DiagnosticCode.INVALID_EXPRESSION
    ]
    assert str(result.diagnostics[0].location) == "Report!F7:12"
