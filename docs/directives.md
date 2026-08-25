# Template language reference

This is the author-facing reference for the currently executable template language. It documents what the implementation supports now. Features proposed in [`../SPEC.md`](../SPEC.md) but not implemented are listed separately at the end.

## Tags at a glance

| Syntax | Kind | Purpose |
| --- | --- | --- |
| `{{ expression }}` | Output tag | Evaluate a value and place it in the cell. |
| `{% for item in items %}` | Opening directive | Start a rectangular repeated block. |
| `{% endfor %}` | Closing directive | Mark the bottom-right corner of a repeated block. |
| `{% if condition %}` | Opening directive | Start the true rectangle of a conditional block. |
| `{% else %}` | Branch directive | End the true rectangle and introduce the false rectangle. |
| `{% endif %}` | Closing directive | Mark the bottom-right corner of a conditional block. |
| `{% region %}` | Opening directive | Start an explicit vertical layout boundary. |
| `{% endregion %}` | Closing directive | Mark the bottom-right corner of a region. |

Output tags use `{{ ... }}`. Directives use `{% ... %}`. Output tags produce values; directives control rectangular worksheet structure and are removed from the rendered cells.

## Spatial authoring rules

Structural tags are paired by their coordinates rather than by a block name.

1. Put the opening directive in the top-left cell of the intended rectangle.
2. Put the matching closing directive in the bottom-right cell.
3. The opening directive must be the first non-whitespace template token in its cell.
4. The closing directive must be the last non-whitespace template token in its cell.
5. Both boundary cells remain part of the rendered body after directive text is removed.
6. Nested rectangles must be completely contained by their parent.
7. Sibling rectangles must be disjoint; they may not partially overlap.
8. Marker geometry must produce exactly one interpretation. The compiler reports ambiguous layouts instead of guessing.

A one-cell block is legal when the opening directive appears before the closing directive in the same cell:

```text
{% for city in cities %}{{ city }}{% endfor %}
```

There are no required `id=` attributes in the current language.

## Output expressions

An output tag evaluates an expression:

```text
{{ invoice.number }}
```

When the expression is the cell's only content, the engine preserves its native scalar type, such as a number, boolean, date, or blank.

When literal text is mixed with an expression, the result is text:

```text
Invoice {{ invoice.number }}
```

A collection cannot be written directly to a scalar cell. Use a `for` block to repeat cells, or use `join` when the desired result really is one text value.

### Excel formatting

Format placeholder cells in Excel exactly as their output should appear. Every rendered destination
inherits the source cell's font, fill, border, alignment, number format, and protection. This also
applies to blank formatted cells in a repeated rectangle.

For example, keep `{{ line.amount }}` as a numeric expression and give its template cell an Excel
currency number format. Repeated values remain numeric and each copy receives that format. The
language does not infer currency, percentages, dates, table stripes, or first/last-row borders.

Whole-row repeats also copy explicit row height and supported row properties. `shift="cells"`
cannot repeat a custom row height because Excel row height applies to the entire worksheet row; the
adapter rejects that combination rather than changing neighboring lanes.

Merged ranges fully contained in a block repeat with it. A merge crossing a block or cell-shift
lane boundary is invalid.

### Expression features

The current expression language supports:

| Feature | Example |
| --- | --- |
| Name | `invoice` |
| Mapping property | `invoice.customer.name` |
| Mapping/list index | `items[0]`, `invoice["number"]` |
| String literal | `"unknown"`, `'unknown'` |
| Number literal | `10`, `12.5` |
| Boolean literal | `true`, `false` |
| Null literal | `null`, `none` |
| Comparison | `total >= 100`, `status == "open"` |
| Boolean operators | `active and paid`, `vip or staff`, `not disabled` |
| Parentheses | `(active or vip) and paid` |
| Filter pipeline | `name \| upper` |

Property syntax reads mapping keys. It does not invoke attributes or methods on arbitrary Python objects. Platform data should be normalized to mappings and ordered collections before rendering.

### Render context values

Template authors use normal variable names; they do not declare types or use wrappers such as
`TypedValue(value, "table")`. The platform supplies a context made from:

- null, strings, booleans, integers, finite floats and decimals;
- timezone-naive dates, datetimes, and times;
- records with string keys;
- finite ordered lists or tuples containing supported values.

Runtime types determine which expression operations are legal. Template syntax determines
presentation. For example, the same list of records may be rendered by a one-row repeat as a table
body or by a multi-row repeat as cards.

The engine normalizes the complete input into an immutable snapshot before evaluating any sheet.
Records become read-only mappings and lists or tuples become tuples. This prevents caller-side
mutation from changing one render halfway through. Sets, generators, non-string record keys,
non-finite numbers, timezone-aware temporal values, cycles, arbitrary objects, and unadapted
DataFrames are rejected. Diagnostics name the input path, for example
`context.lines[2].amount`.

Pandas, Polars, Arrow, DuckDB, ORM, and similar objects belong behind explicit `TypeAdapter`
instances. An adapter converts one declared runtime type into ordinary string-keyed records and
ordered collections; its output is normalized recursively. Adapters cannot override canonical
values, and ambiguous, duplicate, failing, or cyclic conversions are explicit errors.

The optional Polars integration supplies adapters for eager `polars.DataFrame` values:

```python
from excel_template_writer.adapters.polars import polars_adapters
from excel_template_writer.xlsx import render_workbook

render_workbook(
    "template.xlsx",
    "output.xlsx",
    {"lines": frame},
    adapters=polars_adapters(),
)
```

It preserves frame row order and column names, adds no index column, and converts Polars nulls and
float NaNs to template `null`. It does not accept or collect `LazyFrame` values. Nanosecond
`Datetime`, `Time`, and timezone-bearing `Datetime` columns are rejected because materializing them
as Python values would lose precision or violate the canonical temporal model. Templates still use
ordinary expressions such as `line.amount`; the adapter introduces no Polars-specific directive.

### Resource limits

The engine applies permissive default safety limits before and during rendering. They cover context
depth and size, individual collection length, repeat iterations, planned cells, worksheet
dimensions, workbook sheet count, and compressed/uncompressed XLSX package size. Platform code may
override these through `ResourceLimits`; there is no template directive that can weaken them.

Excel's absolute limits of 1,048,576 rows, 16,384 columns, and 32,767 characters in one text cell
always apply. Limit failures stop immediately and produce no context, plan, or output workbook.
Wall-clock cancellation belongs to the hosting platform.

Private names or keys beginning with `_`, imports, assignment, comprehensions, lambdas, and arbitrary function or method calls are prohibited.

### Implemented filters

| Filter | Meaning | Example |
| --- | --- | --- |
| `default(value)` | Return the fallback only when the input path is missing. A present `null` remains null. | `phone \| default("-")` |
| `string` | Convert a value to text; null becomes an empty string. | `reference \| string` |
| `upper` | Convert to uppercase text. | `customer.name \| upper` |
| `lower` | Convert to lowercase text. | `customer.name \| lower` |
| `join(separator)` | Join collection items as text. The default separator is `", "`. | `labels \| join(" / ")` |
| `date(format)` | Format a native date or datetime as deterministic text. | `report_date \| date("dd mmmm yyyy")` |

Filter names, argument counts, and literal argument types are checked while compiling the template.
Unknown filters and dynamic arguments such as `date(format_variable)` are errors. The filters
`datetime`, `number`, and `money` remain proposed and are not executable yet.

#### `date(format)`

Use `date` when a native date must become text, especially inside a sentence:

```text
{{ report_date | date("yyyy-mm") }}
For the month ending {{ report_date | date("dd mmmm yyyy") }}
```

The case-insensitive format codes are:

| Code | Output for 2026-08-31 |
| --- | --- |
| `yyyy`, `yy` | `2026`, `26` |
| `mmmm`, `mmm` | `August`, `Aug` |
| `mm`, `m` | `08`, `8` |
| `dddd`, `ddd` | `Monday`, `Mon` |
| `dd`, `d` | `31`, `31` |

Spaces and ordinary punctuation are literal. For literal words inside the format, use an
Excel-style double-quoted section while quoting the complete expression argument with single
quotes:

```text
{{ report_date | date('dd "of" mmmm yyyy') }}
```

A backslash escapes the next format character. The engine rejects empty formats, unknown letters,
unsupported field widths, braces such as `{yyyy}`, Python `%Y` codes, unterminated quoted text, and
advanced Excel number-format constructs. Month and weekday names are English in the current
release.

The input must be a native `date` or `datetime` value. A datetime contributes only its calendar
date. Strings, numbers, and `null` are rejected instead of being parsed or guessed. Missing values
remain missing-value errors.

`date` always returns text, including when it is the only expression in a cell. To preserve a
native Excel date, keep the expression unfiltered and format the placeholder cell in Excel:

```text
{{ report_date }}
```

For example, the Excel custom number format `"For the month ending "dd mmmm yyyy` displays a full
label while preserving the underlying date value.

## `for` and `endfor`

### Syntax

```text
{% for <variable> in <collection-expression> [direction="down"] [shift="rows"|"cells"] %}
...
{% endfor %}
```

The loop variable must be a normal identifier and cannot begin with `_`. Options must use quoted string values. Each option may appear at most once.

`direction="down"` is the only implemented direction and is the default. `direction="right"` is not accepted yet.

`shift="rows"` is the default. `shift="cells"` is available when growth must be isolated to the block's columns.

### One-row rectangular repeat

For a block spanning `A4:D4`:

```text
A4: {% for line in lines %}{{ line.description }}
B4: {{ line.quantity }}
C4: {{ line.unit_price }}
D4: {{ line.total }}{% endfor %}
```

If `lines` contains three records, the entire `A4:D4` body is rendered three times at rows 4, 5, and 6. Static content below it moves down because the default shift policy is `rows`.

The data must already be in display order. The renderer does not sort or match it to headers.

### Multi-row repeat

The opening and closing cells can define a taller rectangle:

```text
A4: {% for card in cards %}{{ card.title }}
...
D5: {{ card.total }}{% endfor %}
```

This defines `A4:D5`, a two-row body. Each item contributes a complete two-row instance, plus any height added by nested regions.

### Loop scope

Inside the body, the loop variable is available to output expressions, conditions, and nested loops:

```text
{% for group in groups %}{{ group.name }}
{% for item in group.items %}{{ item }}{% endfor %}
{% endfor %}
```

The parent scope remains visible inside a nested loop. Each child scope is separate; rendering does not mutate the input context.

Loop metadata such as `loop.index` or `loop.first` is not implemented yet.

### Empty collections

An empty collection keeps exactly one source-body instance:

- directive tags disappear;
- item-dependent expressions become blank;
- static literal content remains;
- the block retains one row or one multi-row body worth of layout.

For this source row:

```text
A4: {% for item in items %}{{ item.name }}
B4: Formatted placeholder{% endfor %}
```

an empty `items` collection produces a blank `A4` and the literal text `Formatted placeholder` in `B4`.

Missing values unrelated to the absent item remain errors. There is no `{% empty %}` branch yet.

### `shift="rows"`

```text
{% for item in items shift="rows" %}...{% endfor %}
```

This is the default. When the block changes height, every cell below its bottom edge moves by the same amount. Use it for ordinary report tables and vertically stacked sections.

Two sibling blocks whose source row spans overlap cannot independently use whole-row shifting. The compiler rejects that layout because both blocks would claim the same rows.

### `shift="cells"`

```text
{% for item in items shift="cells" %}...{% endfor %}
```

Only cells in columns covered by the block move. Cells in neighboring columns remain stationary.

This permits independently growing side-by-side lists, provided both siblings use `shift="cells"` and their rectangles are disjoint:

```text
A4: {% for left in left_items shift="cells" %}{{ left }}{% endfor %}
C4: {% for right in right_items shift="cells" %}{{ right }}{% endfor %}
```

Content below column A moves as the left list grows; content below column C moves independently as the right list grows.

## `region` and `endregion`

### Syntax

```text
{% region [direction="down"] [shift="rows"|"cells"] %}
...
{% endregion %}
```

A region is an explicit rectangular layout container. Its opening tag is the top-left corner and
its closing tag is the bottom-right corner. It introduces no variable and does not inspect the
cells around it to infer a boundary.

`direction="down"` is the default and only implemented direction. One region has one growth axis;
horizontal and mixed-direction regions are rejected rather than guessed.

### Measuring child layouts together

The source rectangle reserves a minimum amount of space. The renderer first completes all direct
children inside it, then exposes the completed region to the surrounding worksheet as one unit.
Side-by-side child lanes use the tallest completed bottom edge, while vertically stacked children
accumulate their actual movement. Nested regions are measured from the inside out.

For a region from `A1:J10`, child output ending on or above row 10 causes no movement outside the
region. If its completed content ends on row 14, the region has four rows of external growth.

### `shift="rows"`

This is the default. External growth inserts complete worksheet rows below the region's source
bottom. Content in every column moves together.

### `shift="cells"`

External growth moves only cells below the source bottom in the region's exact column band. For an
`A1:J10` region that grows by four rows, `A20:J20` moves to `A24:J24`, while `K20:P20` stays on row
20. This is the explicit way to say that several internal lanes form one visual section without
moving a neighboring section.

```text
A1: {% region shift="cells" %}
A2:C2: first repeated lane, using shift="cells"
D2:F2: second repeated lane, using shift="cells"
G2:I2: third repeated lane, using shift="cells"
J10: {% endregion %}
```

The region markers disappear. Static and formatted blank cells below the region move only when
they are in the affected band. A merged range may not cross the region rectangle or the edge of a
cell-shift band. Because Excel row heights apply across a complete worksheet row, cell-band
movement does not translate row dimensions independently.

## `if`, `else`, and `endif`

Conditions use vertically stacked, equal-width branch rectangles.

### With an `else` branch

For a true branch in `A4:D4` and a false branch in `A5:D5`:

```text
A4: {% if account.active %}Account
D4: ACTIVE{% else %}
A5: Account
D5: INACTIVE{% endif %}
```

The geometry is:

```text
A4 ┌──────── true branch ────────┐ D4 / else
A5 └──────── false branch ───────┘ D5 / endif
```

The `if` marker is the top-left corner of the true branch. `else` is its bottom-right corner. The false branch starts on the next row at the same left column and ends at `endif`.

Only the selected branch is rendered. The other branch contributes no cells or height, so content below compacts upward.

### Without an `else` branch

```text
A4: {% if account.vip %}VIP account
D4: Visible{% endif %}
```

The `if` and `endif` cells define one true rectangle. A truthy condition renders it. A false condition produces no block and closes its vertical space.

A one-cell condition is also legal:

```text
{% if show_note %}Important{% endif %}
```

### Conditional scope and shifting

An `if` directive does not add variables to scope. Its expression can read the surrounding context, including variables from containing loops.

Conditions do not accept `direction` or `shift` options. At the worksheet root they use row
shifting. Inside a loop or region they inherit the nearest container's shift policy, so a
condition in a `shift="cells"` lane does not move unrelated neighboring columns.

Nested blocks must be entirely inside one branch. A loop or condition cannot cross the boundary between the true and false rectangles.

## Common invalid templates

### Opening tag after output content

```text
{{ title }}{% for item in items %}
```

An opening marker must be the first non-whitespace template token in its cell.

### Content after a closing tag

```text
{% endfor %} trailing text
```

A closing marker must be the final non-whitespace template token in its cell.

### Partially overlapping rectangles

Two blocks may be nested or disjoint. They cannot overlap like offset rectangles. The compiler reports `E1205` rather than choosing which block owns the shared cells.

### Missing closing marker

Every `for` needs one compatible `endfor`, every `if` needs one compatible `endif`, and every
`region` needs one compatible `endregion`. The compiler reports `E1201` when it cannot form a
complete spatial pairing.

### Collection in a scalar cell

```text
{{ lines }}
```

This is an error. Use a loop to create cells or `join` to intentionally produce text.

### Unknown or unsupported options

Unknown directives, duplicate loop or region options, `direction="right"`, and shift modes other
than `rows` or `cells` are rejected during compilation.

## Diagnostic codes

| Code | Meaning |
| --- | --- |
| `E1001` | Unterminated output expression |
| `E1002` | Unterminated directive |
| `E1003` | Empty output expression |
| `E1101` | Invalid expression syntax |
| `E1102` | Invalid or unsupported directive syntax |
| `E1103` | Unknown directive option; reserved in the current diagnostics model |
| `E1104` | Unknown filter or invalid filter argument contract |
| `E1105` | Invalid or unsupported `date` format |
| `E1201` | Unmatched block marker |
| `E1202` | Invalid block geometry |
| `E1203` | Ambiguous block pairing |
| `E1204` | Marker is not at the required cell-token boundary |
| `E1205` | Partially overlapping blocks |
| `E1301` | Missing value during rendering |
| `E1302` | Collection used as a scalar cell value |
| `E1303` | `for` expression did not produce a supported collection |
| `E1304` | Runtime value has the wrong type for a compiled filter |
| `E1401` | Two source allocations collided at one destination cell |
| `E1402` | Sibling blocks have conflicting whole-row shift lanes |
| `E1501` | Render context root is not a mapping |
| `E1502` | Record key is not a string |
| `E1503` | Context contains an unsupported value type |
| `E1504` | Context contains a set or frozenset instead of an ordered collection |
| `E1505` | Context contains a non-finite float or decimal |
| `E1506` | Context contains a cyclic record or collection |
| `E1507` | Context contains a timezone-aware datetime or time |
| `E1510` | More than one adapter is registered for the same source type |
| `E1511` | Unrelated adapters match a value, so no unique most-specific adapter exists |
| `E1512` | A value adapter raised an exception |
| `E1513` | Adapter conversion formed a direct or indirect cycle |
| `E1601` | Canonical context exceeded a configured resource limit |
| `E1602` | Rendering exceeded a configured repeat, cell, row, or column limit |
| `E1603` | Rendered geometry exceeded an absolute XLSX grid limit |
| `E1604` | Rendered cell text exceeded Excel's absolute character limit |
| `E2104` | Merged range crosses a block or cell-shift lane boundary |
| `E2105` | Conditional formatting would require an unsupported transform |
| `E2106` | Data validation would require an unsupported transform |
| `E2107` | Native Excel Table is unsupported by the current writer |
| `E2108` | Chart, image, or drawing anchor is unsupported by the current writer |
| `E2109` | Hyperlink would be copied or moved |
| `E2110` | Comment would be copied or moved |
| `E2111` | Cell shifting would repeat a worksheet-wide custom row height |
| `E2201` | Compressed size, uncompressed size, or ZIP member count exceeded a package limit |
| `E2202` | Workbook exceeded the configured worksheet-count limit |
| `E3101` | Formula would require copying, movement, or translation |
| `E3201` | Input and output resolve to the same path |
| `E3202` | Input or output is not an `.xlsx` file |

Diagnostics include at least the worksheet and cell, and lexical diagnostics also carry character offsets.
Context diagnostics instead carry a canonical input path beginning with `context`.

## Not implemented as directives

The following syntax is not currently recognized:

- `{% empty %}`;
- `{% set %}`;
- `{% include %}`, `{% import %}`, or macros;
- `{% break %}` and `{% continue %}`;
- sheet repetition;
- image directives;
- horizontal `for` expansion.

Excel formulas are also outside the first release. Do not treat unsupported syntax as comments: an unknown directive is a compilation error.

## Working examples

The maintained [`../samples/README.md`](../samples/README.md) catalog contains focused, executable
Python examples with matching template and rendered workbooks for every current feature area.
Regenerate the complete set with:

```powershell
uv run --all-extras python -m samples.generate_all
```

The generated workbook [`../scratch/demo_template.xlsx`](../scratch/demo_template.xlsx) is a broader
acceptance playground containing visible tags for every currently demonstrated construct. Compare it
with [`../scratch/demo_output.xlsx`](../scratch/demo_output.xlsx), or regenerate both with:

```powershell
uv run python scratch/demo.py
```

The executable [`../scratch/value_model_example.py`](../scratch/value_model_example.py) demonstrates
canonical scalar, record, list, and table-shaped inputs plus representative rejection diagnostics.
