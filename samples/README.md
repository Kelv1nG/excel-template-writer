# Maintained samples

This directory is the executable catalog of currently supported user-facing features. Every sample
contains Python source that builds an authored template, renders it through the production XLSX
entrypoint, reopens the result, and asserts the behavior it demonstrates. The generated template
and output workbooks are committed beside the source so they can be opened immediately in Excel.

## Run all samples

From the repository root:

```powershell
uv run --all-extras python -m samples.generate_all
```

The base package can run every example except Polars:

```powershell
uv run python -m samples.generate_all
```

When the optional dependency is absent, the generator prints one Polars skip message and continues.

## Sample catalog

| Python module | Generated workbooks | Current features demonstrated |
| --- | --- | --- |
| `samples.scalar_values` | `scalar_values_template.xlsx`, `scalar_values_output.xlsx` | Native scalar cells, exact float round trips, supported 15-digit Decimal amounts and authored decimal display scale, mixed text, mapping access, native dates versus textual `date`, numeric `sum`/`min`/`max`, record and non-null `count`, basic arithmetic with precedence and unary signs, plus `upper`, `join`, and `default` filters |
| `samples.missing_output_policies` | `missing_output_policies_template.xlsx`, `missing_output_policies_output.xlsx` | Strict `error`, tolerant `blank`, and exact-tag `preserve` handling for missing output values; `W1301` warnings, mixed text, `default` precedence, present nulls, and formatted blanks |
| `samples.repeated_blocks` | `repeated_blocks_template.xlsx`, `repeated_blocks_output.xlsx` | One-cell lists, styled rectangular table rows, row shifting, formatted blanks, directive-only cell fill/border preservation, merged footers, and empty-repeat placeholders |
| `samples.default_cell_style` | `default_cell_style_template.xlsx`, `default_cell_style_output.xlsx` | Workbook default white fill for absent blank cells, per-sheet 80% zoom preservation, static text above three expanding tables, direct white-fill and explicit no-fill controls, and directive-only fill/border preservation |
| `samples.conditions_and_nesting` | `conditions_and_nesting_template.xlsx`, `conditions_and_nesting_output.xlsx` | `if`/`else`, no-`else` conditions, boolean expressions, nested repeats, lexical scope, and bottom-up measurement |
| `samples.cell_shift_lanes` | `cell_shift_lanes_template.xlsx`, `cell_shift_lanes_output.xlsx` | Side-by-side `shift="cells"` repeats with independently moving lanes and stationary neighboring cells |
| `samples.fixed_layout_fill` | `fixed_layout_fill_template.xlsx`, `fixed_layout_fill_output.xlsx` | `shift="none"` filling of prepared alternating rows while destination fills, number formats, protection, row heights, neighboring labels, and footers remain fixed; overflow fails atomically |
| `samples.fixed_range_charts` | `fixed_range_charts_template.xlsx`, `fixed_range_charts_output.xlsx` | Fixed nine-row chart references across a twelve-row repeat, including a stationary side chart plus charts pushed downward by whole-row and cell-lane expansion |
| `samples.template_images` | `template_images_template.xlsx`, `template_images_output.xlsx` | Embedded PNG byte preservation plus stationary and downward-moving pictures under whole-row and isolated cell-lane expansion |
| `samples.template_text_shapes` | `template_text_shapes_template.xlsx`, `template_text_shapes_output.xlsx` | Editable styled text boxes, callouts, and arrows; literal tag-like shape text; and stationary or downward-moving shapes under whole-row and isolated cell-lane expansion |
| `samples.regions` | `regions_template.xlsx`, `regions_output.xlsx` | Explicit vertical regions, `shift="cells"`, `shift="rows"`, tallest-lane measurement, reserved source height, exact column bands, and nested regions |
| `samples.polars_dataframe` | `polars_dataframe_template.xlsx`, `polars_dataframe_output.xlsx` | Explicit eager-Polars adapter, Decimal result column computed before rendering, numeric writeback with authored currency format, row-order preservation, typed values, and null/NaN normalization |

Run one sample independently with, for example:

```powershell
uv run python -m samples.regions
uv run python -m samples.missing_output_policies
uv run python -m samples.default_cell_style
uv run --extra polars python -m samples.polars_dataframe
```

The scalar sample labels the formerly lossy float `100000.00000000001`, financial
`Decimal("1234567890123.45")`, and `Decimal("12.50")` displayed with format `0.00`. The Polars sample
computes `computed_amount = amount * Decimal("1.05")` in Polars, checks its Decimal dtype/scalar,
and renders that result as a numeric cell. Both builders assert reopened amounts, types, and
formats. These numeric examples use non-temporal formats: the pinned OpenPyXL reopen-as-float
guarantee requires formats that it does not interpret as dates, times, or durations. With temporal
formats, exact numeric XML tokens and authored formats remain preserved, but the reader may return
temporal Python values or `#VALUE!` for out-of-range serials; no new temporal fidelity is promised.
The guarantee does not cover desktop Excel resaving the workbooks. Decimal scale is presentation,
and business rounding belongs to the caller. Regenerate just these pairs with:

```powershell
uv run --all-extras python -m samples.scalar_values
uv run --all-extras python -m samples.polars_dataframe
```

The committed `missing_output_policies_output.xlsx` workbook uses `preserve` so unresolved tags and
their exact authored spacing remain visible. Its generator also renders temporary `error` and
`blank` variants to verify that strict mode publishes no workbook, while blank mode publishes a
formatted blank plus nonfatal `W1301` diagnostics.

The `default_cell_style` pair keeps gridlines enabled and the worksheet zoom at 80%. Its blank band
`B6:H11` has no individual cell records and inherits white from the workbook default cell format
(`cellXfs[0]`) and built-in Normal style (`cellStyleXfs`). The generator checks both saved OOXML
definitions and the absent coordinates, then verifies zoom, direct white-fill and explicit no-fill
controls, plus all 3/2/4 rendered table rows after reload.

## Required sample coverage for new features

Whenever a user-visible template-language, layout, value-adapter, formatting, or XLSX behavior is
added or changed, the same change must add or update supporting material under `samples/`.

At minimum, that supporting material must include:

- executable Python that builds the template and renders it through the production API;
- an authored `*_template.xlsx` workbook with visible tags;
- the matching `*_output.xlsx` workbook;
- save/reload assertions proving the relevant values, geometry, types, and presentation; and
- an entry in the catalog above.

Samples document the current implementation. They do not define semantics independently: `SPEC.md`
remains normative, and tests remain the executable correctness contract.
