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
