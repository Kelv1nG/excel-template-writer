# Excel Template Writer

A spatial, declarative template interpreter for generating `.xlsx` workbooks.

The project is in its executable-design phase. The current implementation compiles an in-memory worksheet into a typed spatial AST and evaluates it into an adapter-neutral destination-cell plan. It deliberately does not mutate an Excel workbook yet.

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

The next implementation boundary is the `openpyxl` adapter for reading and writing `.xlsx` files; it must consume the render plan rather than bypassing the compiler.
