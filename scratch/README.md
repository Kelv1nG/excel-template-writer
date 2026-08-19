# Visual formatting acceptance demo

Run the demo from the repository root:

```powershell
uv run python scratch/demo.py
```

The command recreates two files:

- `demo_template.xlsx` — the authored workbook with visible `{% ... %}` and `{{ ... }}` tags;
- `demo_output.xlsx` — the rendered workbook produced by the production XLSX adapter.

Open the two workbooks side by side and begin on the `START HERE` sheet. The numbered sheets
isolate the current presentation and layout contracts:

1. a styled rectangular table body repeated as complete rows;
2. multi-row cards with repeated merged headers;
3. the single formatted placeholder retained by an empty repeat;
4. independently growing `shift="cells"` lanes;
5. differently styled conditional branches that compact after selection;
6. native date, number, boolean, and string values with template-owned formats;
7. nested repeats whose row styles and heights follow measured group sizes.

The intentionally blank gold cells make blank-cell style copying easy to see. Gold footer bands
show where content below a block lands after expansion or compaction.

The script does not contain a second renderer. It creates the fixture, calls
`excel_template_writer.xlsx.render_workbook`, reloads both files with `openpyxl`, and asserts the
demonstrated values, types, direct styles, row heights, column widths, styled blanks, and merged
ranges.

## Canonical value model example

Run the smaller, workbook-free value example with:

```powershell
uv run python scratch/value_model_example.py
```

It demonstrates ordinary scalar values, a record, a list of scalars, table-shaped data as a list
of records, the immutable normalized snapshot, and a caller-supplied adapter for a DataFrame-like
object. It also demonstrates an intentionally lowered `ResourceLimits` ceiling and prints the
path-aware diagnostics produced for an unordered set, a non-finite number, and that same object
when no adapter is supplied. The example intentionally uses no `TypedValue` wrapper: runtime value
categories control permitted operations, while the template controls layout.

## Polars adapter example

Run the optional Polars example with:

```powershell
uv run --extra polars python scratch/polars_example.py
```

It creates `polars_template.xlsx` and `polars_output.xlsx` in `scratch/`. The template contains a
styled rectangular repeat, while the input rows come directly from an eager `polars.DataFrame`.
The script prints the immutable canonical records, renders through the production XLSX entrypoint,
and verifies row order, native numeric/temporal values, and NaN-to-blank conversion after save/reload.
The generated workbooks are disposable and are not source fixtures.
