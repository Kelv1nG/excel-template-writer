# Capability demo

Run the throwaway demonstration from the repository root:

```powershell
uv run python scratch/demo.py
```

It recreates two files:

- `demo_template.xlsx` — the authored workbook with visible `{% ... %}` and `{{ ... }}` tags;
- `demo_output.xlsx` — the rendered workbook.

The workbook uses one sheet per behavior so the source geometry remains easy to inspect:

- typed scalar cells, mixed text, lookups, comparisons, and implemented filters;
- a normal table-style row repeat using whole-row shifting;
- a one-cell list;
- two independently expanding `shift="cells"` lists;
- the formatted placeholder retained for an empty collection;
- a two-row repeated card with a contained merged title;
- nested repeats;
- true, false, and omitted conditional branches;
- a summary of capabilities that are not implemented yet.

The script is intentionally not the production workbook adapter. It uses the real compiler and render planner, applies the resulting source-to-destination mapping to a new workbook, and reopens the saved file to verify the demonstrated values, types, styles, and merges.
