from excel_template_writer.ast import ForNode, IfNode
from excel_template_writer.compiler import compile_sheet
from excel_template_writer.diagnostics import DiagnosticCode
from excel_template_writer.model import Rectangle, WorksheetTemplate


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
