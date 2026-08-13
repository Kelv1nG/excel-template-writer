"""Cell-level lexer for literal text, output tags, and directive tags."""

from __future__ import annotations

from dataclasses import dataclass

from excel_template_writer.diagnostics import Diagnostic, DiagnosticCode, SourceLocation


@dataclass(frozen=True)
class SourceSpan:
    sheet: str
    cell: str
    start: int
    end: int

    @property
    def location(self) -> SourceLocation:
        return SourceLocation(self.sheet, self.cell, self.start, self.end)


@dataclass(frozen=True)
class TextToken:
    text: str
    span: SourceSpan


@dataclass(frozen=True)
class OutputToken:
    source: str
    span: SourceSpan


@dataclass(frozen=True)
class DirectiveToken:
    source: str
    span: SourceSpan


CellToken = TextToken | OutputToken | DirectiveToken


@dataclass(frozen=True)
class CellLexResult:
    tokens: tuple[CellToken, ...]
    diagnostics: tuple[Diagnostic, ...]


def lex_cell(sheet: str, cell: str, value: str) -> CellLexResult:
    tokens: list[CellToken] = []
    position = 0
    while position < len(value):
        output_start = value.find("{{", position)
        directive_start = value.find("{%", position)
        starts = [start for start in (output_start, directive_start) if start >= 0]
        if not starts:
            if position < len(value):
                tokens.append(
                    TextToken(value[position:], SourceSpan(sheet, cell, position, len(value)))
                )
            break
        start = min(starts)
        if start > position:
            tokens.append(
                TextToken(value[position:start], SourceSpan(sheet, cell, position, start))
            )
        is_output = start == output_start
        closing = "}}" if is_output else "%}"
        end_start = value.find(closing, start + 2)
        if end_start < 0:
            code = (
                DiagnosticCode.UNTERMINATED_EXPRESSION
                if is_output
                else DiagnosticCode.UNTERMINATED_DIRECTIVE
            )
            diagnostic = Diagnostic(
                code,
                f"unterminated {'expression' if is_output else 'directive'} tag",
                SourceLocation(sheet, cell, start, len(value)),
            )
            return CellLexResult((), (diagnostic,))
        end = end_start + 2
        source = value[start + 2 : end_start].strip()
        span = SourceSpan(sheet, cell, start, end)
        if is_output:
            if not source:
                diagnostic = Diagnostic(
                    DiagnosticCode.EMPTY_EXPRESSION,
                    "output expression cannot be empty",
                    span.location,
                )
                return CellLexResult((), (diagnostic,))
            tokens.append(OutputToken(source, span))
        else:
            tokens.append(DirectiveToken(source, span))
        position = end
    return CellLexResult(tuple(tokens), ())
