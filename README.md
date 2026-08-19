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
string-keyed mappings, and ordered lists or tuples. Before evaluation, `normalize_context()` makes
an immutable snapshot in which records are read-only mappings and lists become tuples. Errors are
aggregated at paths such as `context.rows[2].amount`; a failed result contains no partial context.

```python
from excel_template_writer import normalize_context

context = {
    "title": "Revenue",
    "regions": ["North", "South"],
    "rows": [{"description": "Service", "amount": 125}],
}
normalized = normalize_context(context).require()
```

A list of records is table-shaped input, not a special table type. Template directives and their
rectangles decide whether those records render as rows, cards, or another layout. Polars support is
an optional extra, so the base package does not import or install Polars:

```python
import polars as pl

from excel_template_writer import normalize_context
from excel_template_writer.adapters.polars import polars_adapters

frame = pl.DataFrame(
    {"description": ["Service", "Support"], "amount": [125, 75]}
)
normalized = normalize_context(
    {"rows": frame},
    adapters=polars_adapters(),
).require()
```

Install the integration in a consuming project with `uv add "excel-template-writer[polars]"`, or
run this repository with `uv run --extra polars ...`. The adapter accepts eager `DataFrame` values,
preserves row order, introduces no index, and converts null and nested float NaN values to canonical
nulls. It deliberately does not collect a `LazyFrame`. Nanosecond/timezone-bearing temporal values
that Python cannot represent without loss are rejected before materialization.

Adapters remain scoped to one call. Canonical values always keep their built-in meaning, and
adapter output is recursively normalized. Other integrations can use caller-supplied `TypeAdapter`
instances; pandas, Arrow, and DuckDB adapters are not bundled yet.

Every operation also uses an immutable `ResourceLimits` policy. The defaults are permissive for
small business workbooks but bound context depth and size, repeat work, planned cells, worksheet
dimensions, sheet count, and XLSX package size. Callers can override the configurable ceilings;
Excel's absolute grid and cell-text limits always apply.

```python
from excel_template_writer import ResourceLimits

limits = ResourceLimits(max_repeat_iterations_per_sheet=10_000)
plan = render_sheet(compiled, context, limits=limits).require()
```

Render a workbook to a separate output path with:

```python
from excel_template_writer.xlsx import render_workbook

render_workbook(
    "template.xlsx",
    "rendered.xlsx",
    {"rows": [{"name": "Service", "amount": 125}], "total": 125},
    limits=limits,
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
