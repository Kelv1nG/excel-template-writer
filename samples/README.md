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
| `samples.scalar_values` | `scalar_values_template.xlsx`, `scalar_values_output.xlsx` | Native scalar cells, mixed text, mapping access, native dates versus textual `date`, numeric `sum`/`min`/`max`, record and non-null `count`, basic arithmetic with precedence and unary signs, plus `upper`, `join`, and `default` filters |
| `samples.repeated_blocks` | `repeated_blocks_template.xlsx`, `repeated_blocks_output.xlsx` | One-cell lists, styled rectangular table rows, row shifting, formatted blanks, directive-only cell fill/border preservation, merged footers, and empty-repeat placeholders |
| `samples.conditions_and_nesting` | `conditions_and_nesting_template.xlsx`, `conditions_and_nesting_output.xlsx` | `if`/`else`, no-`else` conditions, boolean expressions, nested repeats, lexical scope, and bottom-up measurement |
| `samples.cell_shift_lanes` | `cell_shift_lanes_template.xlsx`, `cell_shift_lanes_output.xlsx` | Side-by-side `shift="cells"` repeats with independently moving lanes and stationary neighboring cells |
| `samples.fixed_range_charts` | `fixed_range_charts_template.xlsx`, `fixed_range_charts_output.xlsx` | Fixed nine-row chart references across a twelve-row repeat, including a stationary side chart plus charts pushed downward by whole-row and cell-lane expansion |
| `samples.template_images` | `template_images_template.xlsx`, `template_images_output.xlsx` | Embedded PNG byte preservation plus stationary and downward-moving pictures under whole-row and isolated cell-lane expansion |
| `samples.template_text_shapes` | `template_text_shapes_template.xlsx`, `template_text_shapes_output.xlsx` | Editable styled text boxes, callouts, and arrows; literal tag-like shape text; and stationary or downward-moving shapes under whole-row and isolated cell-lane expansion |
| `samples.regions` | `regions_template.xlsx`, `regions_output.xlsx` | Explicit vertical regions, `shift="cells"`, `shift="rows"`, tallest-lane measurement, reserved source height, exact column bands, and nested regions |
| `samples.polars_dataframe` | `polars_dataframe_template.xlsx`, `polars_dataframe_output.xlsx` | Explicit eager-Polars adapter, row-order preservation, typed values, and null/NaN normalization |

Run one sample independently with, for example:

```powershell
uv run python -m samples.regions
uv run --extra polars python -m samples.polars_dataframe
```

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
