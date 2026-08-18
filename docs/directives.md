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

Sets, generators, non-string record keys, non-finite numbers, timezone-aware temporal values,
cycles, arbitrary objects, and unadapted DataFrames are rejected before evaluation. Diagnostics
name the input path, for example `context.lines[2].amount`.

Pandas, Polars, Arrow, DuckDB, ORM, and similar objects belong behind platform adapters. An adapter
must convert them into ordinary string-keyed records and ordered collections before rendering.
Adapter implementation is a later milestone; the current API rejects those objects explicitly.

Private names or keys beginning with `_`, imports, assignment, comprehensions, lambdas, and arbitrary function or method calls are prohibited.

### Implemented filters

| Filter | Meaning | Example |
| --- | --- | --- |
| `default(value)` | Return the fallback only when the input path is missing. A present `null` remains null. | `phone \| default("-")` |
| `string` | Convert a value to text; null becomes an empty string. | `reference \| string` |
| `upper` | Convert to uppercase text. | `customer.name \| upper` |
| `lower` | Convert to lowercase text. | `customer.name \| lower` |
| `join(separator)` | Join collection items as text. The default separator is `", "`. | `labels \| join(" / ")` |

The filters `date`, `datetime`, `number`, and `money` appear in the broader specification as proposed features but are not executable yet. Use workbook number formats for the current demo where possible.

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

Conditions do not accept `direction` or `shift` options. At the worksheet root they use row shifting. Inside a loop they inherit the containing loop's shift policy, so a condition in a `shift="cells"` lane does not move unrelated neighboring columns.

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

Every `for` needs one compatible `endfor`, and every `if` needs one compatible `endif`. The compiler reports `E1201` when it cannot form a complete spatial pairing.

### Collection in a scalar cell

```text
{{ lines }}
```

This is an error. Use a loop to create cells or `join` to intentionally produce text.

### Unknown or unsupported options

Unknown directives, duplicate loop options, `direction="right"`, and shift modes other than `rows` or `cells` are rejected during compilation.

## Diagnostic codes

| Code | Meaning |
| --- | --- |
| `E1001` | Unterminated output expression |
| `E1002` | Unterminated directive |
| `E1003` | Empty output expression |
| `E1101` | Invalid expression syntax |
| `E1102` | Invalid or unsupported directive syntax |
| `E1103` | Unknown directive option; reserved in the current diagnostics model |
| `E1201` | Unmatched block marker |
| `E1202` | Invalid block geometry |
| `E1203` | Ambiguous block pairing |
| `E1204` | Marker is not at the required cell-token boundary |
| `E1205` | Partially overlapping blocks |
| `E1301` | Missing value during rendering |
| `E1302` | Collection used as a scalar cell value |
| `E1303` | `for` expression did not produce a supported collection |
| `E1401` | Two source allocations collided at one destination cell |
| `E1402` | Sibling blocks have conflicting whole-row shift lanes |
| `E1501` | Render context root is not a mapping |
| `E1502` | Record key is not a string |
| `E1503` | Context contains an unsupported value type |
| `E1504` | Context contains a set or frozenset instead of an ordered collection |
| `E1505` | Context contains a non-finite float or decimal |
| `E1506` | Context contains a cyclic record or collection |
| `E1507` | Context contains a timezone-aware datetime or time |
| `E2104` | Merged range crosses a block or cell-shift lane boundary |
| `E2105` | Conditional formatting would require an unsupported transform |
| `E2106` | Data validation would require an unsupported transform |
| `E2107` | Native Excel Table is unsupported by the current writer |
| `E2108` | Chart, image, or drawing anchor is unsupported by the current writer |
| `E2109` | Hyperlink would be copied or moved |
| `E2110` | Comment would be copied or moved |
| `E2111` | Cell shifting would repeat a worksheet-wide custom row height |
| `E3101` | Formula would require copying, movement, or translation |
| `E3201` | Input and output resolve to the same path |
| `E3202` | Input or output is not an `.xlsx` file |

Diagnostics include at least the worksheet and cell, and lexical diagnostics also carry character offsets.
Context diagnostics instead carry a canonical input path beginning with `context`.

## Not implemented as directives

The following syntax is not currently recognized:

- `{% empty %}`;
- `{% region %}` or another isolation-container directive;
- `{% set %}`;
- `{% include %}`, `{% import %}`, or macros;
- `{% break %}` and `{% continue %}`;
- sheet repetition;
- image directives;
- horizontal `for` expansion.

Excel formulas are also outside the first release. Do not treat unsupported syntax as comments: an unknown directive is a compilation error.

## Working examples

The generated workbook [`../scratch/demo_template.xlsx`](../scratch/demo_template.xlsx) contains visible tags for every currently demonstrated construct. Compare it with [`../scratch/demo_output.xlsx`](../scratch/demo_output.xlsx), or regenerate both with:

```powershell
uv run python scratch/demo.py
```

The executable [`../scratch/value_model_example.py`](../scratch/value_model_example.py) demonstrates
canonical scalar, record, list, and table-shaped inputs plus representative rejection diagnostics.
