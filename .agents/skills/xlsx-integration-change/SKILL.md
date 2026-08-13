---
name: xlsx-integration-change
description: Design, implement, or review changes involving openpyxl, XLSX loading or writing, workbook adapters, cell types, styles, dimensions, merged cells, formulas, package preservation, or binary workbook fixtures. Use whenever a change crosses between the interpreter's render plan and an actual `.xlsx` file.
---

# Xlsx Integration Change

Change workbook input or output without leaking `openpyxl` behavior into language semantics or silently damaging unsupported workbook features.

## Workflow

1. Read the relevant workbook, layout, compatibility, and testing sections of `SPEC.md`.
2. Identify whether the change belongs to the workbook reader, immutable workbook model, render-plan adapter, or writer. Keep it out of the parser, evaluator, and pure layout planner.
3. Define the supported behavior and rejection boundary before editing code.
4. Create the smallest `.xlsx` fixture that demonstrates the behavior. Avoid large business workbooks and unrelated features.
5. Implement the adapter change without making layout decisions in the writer.
6. Save and reopen the result with `openpyxl`, then assert semantic properties:
   - cell values and Python/Excel types;
   - styles and number formats;
   - row heights and column widths;
   - merged ranges;
   - expected sheet and workbook metadata.
7. Inspect selected OOXML parts only when public workbook assertions cannot prove correctness.
8. Run applicable checks through `uv run` and report unsupported or unverified workbook features.

## Workbook rules

- Support `.xlsx` only.
- Never overwrite the input template.
- Parse, validate, and plan before mutating a workbook.
- Copy contained merged ranges per repeated instance; reject boundaries or shift lanes that bisect a merge.
- Reject a formula when a transformation would copy, move, or require translating it in the first release.
- Reject or warn about unsupported workbook objects according to `SPEC.md`; never claim silent preservation.
- Do not compare entire binary workbooks as golden files. ZIP metadata and package ordering can change without semantic differences.
- Do not edit binary fixtures manually. Generate them with a documented fixture builder or Excel and keep each fixture narrowly scoped.

## Definition of done

- Adapter boundaries remain intact.
- The fixture proves the requested behavior and relevant failure behavior.
- Save/reload verification succeeds.
- No unsupported feature is silently lost or reinterpreted.
- Diagnostics identify the worksheet and affected cell or rectangle.
