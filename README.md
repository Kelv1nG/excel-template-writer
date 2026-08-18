# Excel Template Writer

A spatial, declarative template interpreter for generating `.xlsx` workbooks.

The current implementation compiles worksheets into a typed spatial AST, evaluates them into an
adapter-neutral render plan, and applies the completed plan to a new `.xlsx` workbook through an
`openpyxl` adapter. Direct cell formatting, formatted blanks, row and column dimensions, and merged
ranges follow their source template cells.

The language uses Jinja-like tags inside ordinary cells:

```text
A2: {% for row in rows %}{{ row.description }}
B2: {{ row.amount }}{% endfor %}
```

The two directive cells are opposite corners of one rectangular repeat body. Vertical growth inserts whole rows by default; `shift="cells"` isolates growth to the block's columns. Rectangular `if`/`else` blocks use the same spatial model.

## Current API

```python
from excel_template_writer import WorksheetTemplate, compile_sheet, render_sheet

template = WorksheetTemplate.from_rows(
    "Report",
    [
        ["Description", "Amount"],
        ["{% for row in rows %}{{ row.name }}", "{{ row.amount }}{% endfor %}"],
        ["Total", "{{ total }}"],
    ],
)

compiled = compile_sheet(template).require()
plan = render_sheet(
    compiled,
    {"rows": [{"name": "Service", "amount": 125}], "total": 125},
).require()
```

Compilation and rendering return structured diagnostics; `.require()` raises an exception only when a caller prefers exception-based control flow.

Render contexts use ordinary Python values—no `TypedValue` wrapper is required. Supported values
are null, strings, booleans, integers, finite floats and decimals, timezone-naive dates/times,
string-keyed mappings, and ordered lists or tuples. The complete context is validated before
evaluation, with errors reported at paths such as `context.rows[2].amount`.

```python
from excel_template_writer import validate_context

context = {
    "title": "Revenue",
    "regions": ["North", "South"],
    "rows": [{"description": "Service", "amount": 125}],
}
assert validate_context(context) == ()
```

A list of records is table-shaped input, not a special table type. Template directives and their
rectangles decide whether those records render as rows, cards, or another layout. DataFrames and
other library-specific objects require a platform adapter; those adapters are not implemented yet.

Render a workbook to a separate output path with:

```python
from excel_template_writer.xlsx import render_workbook

render_workbook(
    "template.xlsx",
    "rendered.xlsx",
    {"rows": [{"name": "Service", "amount": 125}], "total": 125},
)
```

All worksheets are compiled, planned, and validated before the output workbook is written. The
input file is never overwritten, and the serialized output is reopened before it is published.

## Development

The repository targets CPython 3.12 and uses `uv` exclusively:

```powershell
uv sync --all-groups
uv run pytest
uv run ruff check src tests
uv run ruff format --check src tests
uv run ty check
```

Documentation:

- [`SPEC.md`](SPEC.md) — normative product and language design;
- [`docs/directives.md`](docs/directives.md) — current author-facing syntax and directives;
- [`docs/explained.md`](docs/explained.md) — holistic compiler, AST, renderer, and adapter architecture;
- [`docs/DEVELOPMENT.md`](docs/DEVELOPMENT.md) — contributor workflow.

Affected formulas, conditional formatting, data validation, native Excel Tables, drawings, and
other unsupported coordinate-dependent features are rejected explicitly rather than silently
damaged. See the specification and architecture guide for the current boundary.
