from excel_template_writer.diagnostics import DiagnosticCode
from excel_template_writer.syntax import DirectiveToken, OutputToken, TextToken, lex_cell


def test_lexes_mixed_cell_content_and_preserves_offsets() -> None:
    result = lex_cell("Sheet1", "B4", "Invoice {{ invoice.number }}!")

    assert result.diagnostics == ()
    assert [type(token) for token in result.tokens] == [TextToken, OutputToken, TextToken]
    output = result.tokens[1]
    assert isinstance(output, OutputToken)
    assert output.source == "invoice.number"
    assert (output.span.start, output.span.end) == (8, 28)


def test_lexes_directives_without_treating_them_as_output() -> None:
    result = lex_cell("Sheet1", "A2", "{% for row in rows %}{{ row.name }}{% endfor %}")

    assert result.diagnostics == ()
    assert [type(token) for token in result.tokens] == [
        DirectiveToken,
        OutputToken,
        DirectiveToken,
    ]


def test_reports_unterminated_tag_at_the_cell() -> None:
    result = lex_cell("Sheet1", "C7", "{{ invoice.number")

    assert result.tokens == ()
    assert len(result.diagnostics) == 1
    diagnostic = result.diagnostics[0]
    assert diagnostic.code is DiagnosticCode.UNTERMINATED_EXPRESSION
    assert diagnostic.location.sheet == "Sheet1"
    assert diagnostic.location.cell == "C7"
