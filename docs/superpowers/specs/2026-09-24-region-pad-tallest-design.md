# Tallest-Lane Region Padding Design

**Status:** Proposed for review

**Date:** 2026-09-24

## Context

A `region` already lets independent `shift="cells"` loops grow side by side. The region measures
the tallest completed child, so content after the region moves far enough for the tallest table.
Shorter loops, however, render only their real iterations. They do not create styled blank cells
through the remaining height, so lane borders and fills stop early.

Template authors can work around this by preprocessing the data into one joined rectangular
collection or by adding a separate frame loop. Both approaches obscure the fact that the three
tables are independent. The language needs one explicit layout option that says the region should
equalize the presentation of compatible sibling loops without changing their data.

## Goals

- Keep sibling tables independent: no renderer-side sorting, joining, grouping, or reconciliation.
- Extend the direct formatting of shorter lanes to the tallest sibling lane.
- Make the behavior explicit at the region that owns the relationship.
- Preserve current output for every existing template.
- Keep padding out of expression scope and loop semantics.
- Reject layouts whose intended padding is ambiguous.
- Carry enough information in the render plan for the XLSX adapter to copy presentation without
  accidentally copying values or coordinate-bound workbook objects.

## Non-goals

- Inferring a frame from blank cells, neighboring borders, Excel Tables, or worksheet used ranges.
- Recomputing contextual first-row, last-row, outer-border, or alternating-row styles.
- Padding arbitrary nested/dynamic blocks in the first version.
- Making unequal source-body heights align by partial blocks.
- Adding joins, keyed merging, filler records, or other input preprocessing to the renderer.
- Copying formulas, hyperlinks, comments, drawings, data validation, or conditional formatting into
  padded cells.

## User-facing syntax

`region` gains a `pad` option:

```text
{% region direction="down" shift="cells" pad="tallest" %}
```

Allowed values are:

| Value | Meaning |
| --- | --- |
| `none` | Default. Preserve the current behavior; shorter lanes end after their own iterations. |
| `tallest` | Add format-only body instances to compatible shorter sibling loops until all have the tallest sibling's effective iteration count. |

The option belongs only to `region`. A `pad` option on `for`, `if`, or another directive is invalid.
The region's existing `shift` option remains independent: it still decides whether completed
growth outside the source region shifts full rows or only the region's column band. Both
`shift="rows"` and `shift="cells"` remain legal on the region itself. Participating child loops
must use `shift="cells"`.

### Authored example

The following one-row loop bodies define three independent tables in `A:C`, `D:F`, and `G:I`:

```text
A4: {% region shift="cells" pad="tallest" %}

A6: {% for row in table1 shift="cells" %}{{ row.a }}
B6: {{ row.b }}
C6: {{ row.c }}{% endfor %}

D6: {% for row in table2 shift="cells" %}{{ row.d }}
E6: {{ row.e }}
F6: {{ row.f }}{% endfor %}

G6: {% for row in table3 shift="cells" %}{{ row.g }}
H6: {{ row.h }}
I6: {{ row.i }}{% endfor %}

I7: {% endregion %}
```

If the collection lengths are 3, 5, and 2, the loops still evaluate exactly 3, 5, and 2 items.
The first lane receives two format-only body instances and the third receives three. The footer
row at source row 7 moves to the common row after the fifth body instance. Authors place the side
borders on each loop body's source cells and the bottom border on that footer row.

## Eligibility and validation

`pad="tallest"` deliberately supports one narrow, deterministic geometry in its first version.
The compiler validates it before evaluation or workbook mutation.

A padded region is valid only when:

1. It has at least two direct structural children.
2. Every direct structural child is a `ForNode`; ordinary static cells in the region are allowed.
3. Every participating loop uses `direction="down"` and `shift="cells"`.
4. The loop rectangles are side by side: they are column-disjoint and have the same source top row
   and source height. Existing sibling-disjointness validation already forbids overlap.
5. Each participating loop is a leaf structural node. Its body cannot contain a nested `for`, `if`,
   or `region` in this version.
6. Existing merge rules hold. A merge may be wholly contained in a loop body, but no merge may
   cross a loop, shift-lane, or region boundary.

Static headers, labels, gaps, and footers can remain in the padded region. They are not padding
sources. Static cells below a loop move only when their columns fall inside that loop's normal
cell-shift band; cells in gap columns retain the existing cell-shift behavior and do not move merely
because the region is padded. A continuous footer should therefore lie inside the participating
loop bands, as it does when adjacent tables tile the footer's columns. A padded region may be nested
inside an ordinary outer construct and is measured from the inside out. A loop participating in an
outer padded region cannot contain another structural block, including another padded region.
Separate padded regions own and validate their own sibling loops.

The strict shape is intentional. Equal top rows and equal body heights make the target a whole
number of source-body instances in every lane. Leaf loops ensure that an instance has fixed geometry
and that measuring it does not require inventing a scope for absent data.

### Diagnostics

Directive syntax continues to use `E1102 INVALID_DIRECTIVE`:

- an unknown `pad` value reports `pad must be "none" or "tallest"`;
- a duplicate reports `duplicate region option: pad`;
- `pad` on a loop remains an invalid loop option.

Unsupported padded geometry uses a new stable diagnostic:

```text
E1403 INVALID_REGION_PADDING
```

Region-wide failures are located at the region opener. A failure attributable to one child is
located at that child's opener. Representative messages are:

- `pad="tallest" requires two or more direct side-by-side for blocks`
- `a padded region may contain only direct for blocks`
- `padded for blocks must use direction="down" and shift="cells"`
- `padded for blocks must begin on the same row and have the same source height`
- `padded for blocks cannot contain nested structural blocks`

Diagnostics are emitted in source order and use the existing compiler aggregation behavior. The
compiler should avoid redundant follow-on messages when an invalid child kind already explains why
there are not two eligible lanes.

## Runtime semantics

### Effective count and target

For each participating loop, the renderer evaluates its collection once in the loop's normal parent
scope. Participating collection expressions are evaluated in source-coordinate order so diagnostics
remain deterministic. It preserves each collection's item order exactly. Evaluation order is
separate from the existing reverse-coordinate placement pass used to apply cell shifts safely.

The loop's **effective count** is:

```text
max(1, collection length)
```

This matches the existing empty-collection rule: an empty collection retains one ordinary source
instance. The region target is the maximum effective count of its participating loops. The source
region's reserved height is not a padding target and never creates padding by itself.

Collection evaluation and type failures keep their existing diagnostics. No successful plan or
workbook is returned after such a failure. Planning may use the existing one-instance fallback to
finish deterministic diagnostic collection, but it must not reinterpret a failed collection as
valid empty input.

### Real and empty instances

Each loop first renders exactly as it does today:

- a non-empty collection renders one evaluated body per item;
- an empty collection renders its one ordinary empty placeholder;
- real items receive their normal item scope and instance path;
- the empty placeholder keeps its current missing-root behavior and retains static literal content.

No collection is extended and no synthetic item enters expression scope. Loop order, length, and
future loop metadata continue to describe only the real collection. The ordinary empty placeholder
remains the one existing exception defined by the empty-collection contract.

### Format-only instances

After normal loop rendering, a shorter lane receives:

```text
target - effective_count
```

format-only copies with the full geometry of its exact source body rectangle. A multi-row body is
copied as a complete body each time. The plan remains sparse: each copy emits only source material
cells plus explicitly contained merges. Untouched default blank coordinates do not become
`PlannedCell` entries and require no presentation lookup. Every emitted padded cell:

- has a blank value;
- retains its source coordinate for presentation lookup;
- copies font, fill, border, alignment, number format, protection, and the adapter's existing
  quote-prefix presentation state;
- does not evaluate output expressions;
- does not retain literal text or directive text;
- does not copy a formula, hyperlink, comment, drawing anchor, data validation, conditional
  formatting, or another coordinate-bound workbook object.

A merge wholly contained in the source body is reproduced for each format-only instance, with a
blank top-left value. Existing collision and merge-boundary validation still applies.

Padding rows do not map to source row dimensions. In particular, `shift="cells"` padding cannot
copy a custom row height, hidden state, outline level, collapsed state, or row-level style into only
one column lane because those properties are worksheet-wide. Padded cells use the worksheet's
existing/default row geometry. Column dimensions are unaffected by vertical padding.

Static cells below each loop and inside its column band are displaced by the loop's completed padded
height under the existing cell-shift rules. Because every participating body has the same source row
span and target count, lane-local footer cells arrive at the same destination rows. Static cells in
columns not owned by a participating loop are not moved by padding. The region's completed height
remains the maximum completed child edge, so padding does not add height beyond the tallest
real/empty lane.

### Resource limits

Format-only instances are not loop iterations and do not increment
`max_repeat_iterations_per_sheet`. The tallest real collection already bounds the target count.
Padded material cells count toward the planned-cell limit, and padded cells and merges participate
in ordinary collision validation. Padding does not create a row extent beyond the already measured
tallest lane, so it has no separate rendered-row or XLSX-grid charge. These rules bound the added
material without inventing data iterations.

## AST and render-plan representation

The language front end adds an explicit normalized value:

```python
RegionDirective.pad: str = "none"
RegionNode.pad: str = "none"
```

Using `"none"` instead of `None` matches the existing normalized `direction` and `shift` fields and
makes the default visible in tests and debugging. `ForDirective` and `ForNode` do not gain a padding
field.

The render plan adds:

```python
PlannedCell.format_only: bool = False
```

`format_only=False` preserves the meaning of every current planned cell. A padded cell has
`format_only=True`, `value=None`, and a real `source_coordinate` used only to retrieve direct cell
presentation. `format_only=True` with a non-blank value violates the render-plan invariant and is an
internal error. Its `instance_path` remains the enclosing real scope's path and does not append a
synthetic loop index. Multiple padded cells do not need unique paths because destination coordinates
already identify material allocations.

Format-only instances do not contribute or overwrite worksheet-wide `PlannedRow` provenance.
Destination rows already mapped by the enclosing area or a real sibling retain that mapping;
genuinely new rows created by cell-shift growth have `source_row=None` under the existing planner
rule. Padding in one lane must never clear or replace row provenance shared with another lane.
Padded contained merges continue to use explicit `PlannedMerge` entries; the writer never infers
them from blank cells.

The adapter maintains two conceptual projections of planned cells:

- **content destinations** exclude `format_only` cells and are used for formulas, hyperlinks,
  comments, drawing anchors, and custom-row-height repeat analysis;
- **material-layout destinations** include `format_only` cells and are used for identity detection,
  collisions, merges, and resource accounting.

Any format-only destination makes a plan non-identity. A padded cell's source coordinate is a style
lookup key, not a claim that source content or an attached object was copied. Normal cells and
actual loop instances retain all current validation.

## Component design

### Directive parser

- Accept `pad` only while parsing `region` options.
- Normalize an omitted option to `"none"`.
- Accept only `"none"` and `"tallest"`.
- Recognize `pad` as an option-shaped token when separating a loop collection expression from its
  trailing options, but keep the loop's allowed-option set limited to `direction` and `shift`. This
  makes a misplaced `pad` an intentional invalid-loop-option diagnostic instead of accidental
  expression syntax.
- Update every changed helper's docstring, including complete `Args`, `Returns`, and `Raises`
  sections where applicable.

### Compiler and semantic validator

- Thread `RegionDirective.pad` into `RegionNode.pad`.
- Validate the strict padded-region geometry after direct children are known.
- Emit `E1403` at deterministic region/child source locations.
- Leave ordinary regions and all existing sibling-shift validation unchanged.

### Evaluator and layout planner

- Evaluate each participating collection once in source order and retain the prepared real/empty
  loop blocks; keep that evaluation pass separate from reverse-coordinate child placement.
- Compute the maximum effective count.
- Append presentation-only source-body blocks to short lanes without calling expression evaluation.
- Feed the completed child blocks through the existing cell-shift placement and collision logic.
- Account for format-only cells and completed geometry in existing resource checks.
- Do not add an ad hoc second workbook mutation pass.

### XLSX validation and writer

- Build the content and material-layout projections described above rather than globally dropping
  format-only cells from all source-to-destination consumers.
- Exclude format-only mappings from coordinate-bound feature and custom-row-height analyses while
  retaining them for identity, collision, merge, and resource behavior.
- Use `source_coordinate` to copy only the direct presentation of a format-only cell.
- Do not send a format-only cell through the normal cell-copy helper. Use a positive-allowlist
  presentation helper that copies only font, fill, border, alignment, number format, protection,
  and quote prefix. It does not copy or inspect the source value, formula state, hyperlink, comment,
  or any present or future attached-cell feature.
- Enforce the invariant that every format-only cell has a blank value.
- Apply explicit padded merges from the validated plan.
- Keep layout decisions in the planner; the writer only executes the plan.
- Do not relax current `E2105`/`E2106` handling for conditional formatting or data validation on a
  structurally transformed worksheet.

## Compatibility

- Omitted `pad` and explicit `pad="none"` are behaviorally identical to the current release.
- No existing loop syntax, scope, collection ordering, empty-placeholder behavior, or region
  measurement changes.
- `pad="tallest"` is opt-in and fails explicitly outside its supported geometry.
- Direct cell formatting remains source-authored; the feature copies it but never infers it.
- Unsupported workbook objects are still rejected or preserved under their existing contracts.

## Verification plan

### Parser and compiler tests

- An ordinary region compiles with `RegionNode.pad == "none"`.
- `pad="none"` and `pad="tallest"` parse and reach the AST.
- Unknown/duplicate values and `pad` on `for` produce `E1102` at the exact opener.
- Three equal-height, same-top, side-by-side leaf cell-shift loops compile.
- `E1403` covers fewer than two loops, a non-loop direct child, row-shift child, staggered top,
  unequal source height, and a nested structural child, with exact source locations.

### Pure render-plan tests

- Collections of lengths 3, 5, and 2 retain only their own values while short lanes receive blank
  `format_only` cells through effective count 5.
- An ordinary region with the same inputs remains ragged.
- Equal lengths add no format-only cells.
- Three empty collections keep one ordinary placeholder each and add no padding.
- An empty lane beside a non-empty lane keeps its ordinary placeholder, including current static
  literals, while only the additional bodies are completely blank.
- Equal two-row bodies pad in complete two-row units and copy safely contained merges.
- A taller reserved source region does not create padding beyond the tallest child.
- A padded inner region in an ordinary outer construct measures inside out.
- Padding cells count toward planned-cell limits but not repeat-iteration limits; padding adds no
  row extent beyond the already measured tallest lane.
- Footer cells move once to the common target boundary; adjacent columns outside a cell-shift region
  remain unchanged.

### XLSX integration tests

Save and reopen a minimal workbook and assert:

- padded cell values are blank;
- font, fill, left/right borders, alignment, number format, protection, and quote-prefix presentation
  match the corresponding source cells;
- the authored footer/bottom border appears once after the tallest lane;
- contained padded merges are present and non-overlapping;
- formulas, hyperlinks, and comments occur only at their valid normal destinations and are absent
  from padding;
- padded cells sourced from formula cells reopen with a blank value and a non-formula cell type;
- chart, image, and static-text-shape anchors in a short lane retain exactly one normal destination
  and are not copied by padding;
- conditional formatting and data validation on a padded structurally transformed sheet retain
  their existing `E2105` and `E2106` rejection behavior;
- padding alone does not copy custom row dimensions or trigger the cell-shift custom-height rule;
- columns outside the region remain unchanged.

For a padded contained merge, reload assertions cover the merged rectangle and its authored edge
borders/protection, because merging can affect subordinate-cell presentation in `openpyxl`.

Tests compare semantic workbook properties after reload, not XLSX package bytes.

## Documentation and maintained example

Implementation includes all of the following in the same change:

- normative updates to `SPEC.md` for region padding, direct formatting, row dimensions, merges,
  render-plan representation, diagnostics, and resource accounting;
- user-facing syntax, parameter descriptions, constraints, and the three-table example in
  `docs/directives.md`;
- the planner/IR boundary in `docs/explained.md`;
- complete parameter documentation in changed/new Python docstrings;
- a brief capability reference in `README.md`;
- a dedicated padded-independent-tables sheet in `samples/regions.py` using three collections of
  lengths 3, 5, and 2, distinct lane formatting, side borders, and an authored footer border;
- assertions in that sample for independent values, blank padding, copied presentation, common
  footer position, and unaffected neighboring cells;
- regenerated and committed `samples/regions_template.xlsx` and `samples/regions_output.xlsx`;
- an updated `samples/README.md` catalog entry.

The maintained sample generator, not manual binary editing, is the source of the example workbook.

## Rejected alternatives

### Implicit border extension

Inferring that visually adjacent borders form one frame makes blank styles and neighboring cells
semantic. It conflicts with the rule that regions come only from explicit markers and would be
unpredictable for fills, stripes, and interior borders.

### `align="stretch"`

This CSS-like name does not say what is stretched in a spreadsheet and could be confused with cell
alignment. `pad="tallest"` names both the operation and its target.

### `fill="tallest"`

Excel already uses *fill* for cell background formatting and fill handles, making this wording
ambiguous.

### Per-loop padding options

Repeating the same option on every loop makes agreement and ownership unclear. The relationship is
between siblings, so the containing region should declare it once.

### Renderer-side data joining or filler records

Synthetic records would change loop length/scope and entangle layout with input preprocessing. It
would also expose fake values to expressions and future loop metadata. Format-only layout padding
keeps those concerns separate.

### A post-render workbook repair pass

Adding borders after rendering would bypass the validated layout plan, duplicate geometry logic in
the XLSX writer, and make merges and unsupported workbook objects unsafe. Padding must be planned
before any workbook mutation.
