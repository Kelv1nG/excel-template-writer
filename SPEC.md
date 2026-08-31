# Excel Template Engine Specification

Status: Draft 0.2
Scope: Normative product, language, and rendering contract

## 1. Purpose

The engine renders a user-authored Excel workbook template and structured input data into a new `.xlsx` workbook.

The intended host is a larger platform in which a user or administrator designs a workbook in Excel, the platform supplies data, and the engine returns the finished workbook as bytes or a file.

The language is declarative and Jinja-like, but its execution model is spatial. A spreadsheet template is not treated as one long string. Expressions belong to cells, and structural constructs own rectangular cell regions.

## 2. Confirmed product decisions

- The only supported file format is `.xlsx`.
- Template authors are ordinary Excel users or technically capable business users.
- The authoring language must be declarative and reasonably learnable without Python knowledge.
- Required output kinds include scalar values, lists, and rectangular styled blocks.
- A “table” means a rectangular block styled as a table. It does not initially mean an Excel native Table object.
- Merged cells are required.
- Static headers and static rows are authored directly in the workbook.
- The render context is presentation-ready: the renderer does not infer fields from headers, sort collections, or perform keyed merging, joins, or reconciliation.
- Vertical repeats insert complete rows by default. Authors can select a narrower shift policy explicitly.
- An empty repeat retains one blank formatted instance of its source block.
- Conditions are required in the first usable language release.
- The first usable release supports vertical (`down`) repetition; horizontal repetition follows after the vertical model is proven.
- Conditional branches are vertically stacked, equal-width rectangles in the first usable release.
- Templates are treated as untrusted regardless of who uploads them.
- Formula copying, translation, and calculation are outside the first release.
- The implementation must use a real tokenizer, parser, AST, validation passes, and a deterministic layout engine. Direct string replacement is not an acceptable architecture.
- `openpyxl` is the proposed workbook adapter.

## 3. Goals

1. Preserve the visual design of an authored workbook while replacing template constructs with typed values.
2. Repeat one-cell, one-row, one-column, or multi-cell rectangular blocks vertically or horizontally.
3. Make movement caused by expansion explicit and predictable.
4. Detect invalid templates and layout collisions before producing a workbook.
5. Preserve Excel value types rather than converting everything to text.
6. Copy styles, dimensions, and merged ranges as part of block expansion.
7. Keep the template language safe when templates and input data are untrusted.
8. Separate language semantics from `openpyxl` so the interpreter can be tested without writing workbook files.

## 4. Non-goals for the first release

- `.xls`, `.xlsm`, `.xltx`, `.ods`, CSV, or password-protected files
- Executing arbitrary Python or importing Python modules from a template
- Calculating, copying, or translating Excel formulas
- Native Excel Table creation or resizing
- Pivot tables, pivot charts, slicers, macros, form controls, and unsupported drawing objects
- Creating charts, resizing chart data ranges, or resizing and repeating chart anchors
- Creating images from render data, repeating images, or resizing image anchors
- Creating shapes, evaluating shape text, repeating shapes, or resizing shape anchors
- A complete implementation of Jinja2
- Automatic inference of repeat regions from blank cells, formatting, or worksheet used ranges
- Automatic matching between displayed headers and input field names
- Sorting, grouping, joining, keyed merging, or reconciliation of input records
- Editing the input template in place

## 5. Design principles

### 5.1 Explicit geometry

Every structural construct has an exact source rectangle. The engine must never guess a block's width or height from neighboring blank or styled cells.

### 5.2 Familiar surface, spreadsheet-specific semantics

The delimiters resemble Jinja:

```text
{{ expression }}
{% directive %}
```

The language is nevertheless its own language. Spreadsheet direction, rectangular regions, cell types, formulas, merges, and expansion are first-class concepts.

### 5.3 Typed output

An expression that occupies an entire cell writes a typed Excel value. An expression embedded in other text produces a string.

### 5.4 Parse, validate, plan, then mutate

The engine must parse and validate the complete template and compute the output layout before changing a workbook. It must not expand rows while still discovering template syntax.

### 5.5 Fail visibly

Ambiguous geometry, overlapping regions, illegal merges, missing values, and collisions are errors by default. Silent overwriting is not permitted.

## 6. Template authoring model

Template syntax is stored in normal cell values. This keeps templates inspectable, copyable, and understandable without requiring users to find Excel's Notes interface.

### 6.1 Scalar cells

```text
{{ invoice.number }}
```

```text
Invoice {{ invoice.number }}
```

```text
{{ invoice.total | money }}
```

If the first example evaluates to a number, date, datetime, boolean, or blank, that type is written directly to Excel. The second example always produces text because it mixes literal text and an expression.

### 6.2 Rectangular repeat blocks

The proposed syntax places an opening directive in the top-left cell of a block and its matching closing directive in the bottom-right cell.

For a one-row block spanning `A6:F6`:

```text
A6: {% for line in invoice.lines direction="down" %}{{ line.description }}
B6: {{ line.quantity }}
C6: {{ line.unit_price | money }}
F6: {{ line.total | money }}{% endfor %}
```

The structural tags are removed from the output. The remaining content and formatting in `A6:F6` are cloned once per item.

For a two-row block spanning `A6:F7`, the opening tag remains in `A6` and the closing tag is placed in `F7`.

Block IDs are not part of the normal authoring syntax. The spatial linker pairs opening and closing markers using their coordinates, directive types, and the requirement that the resulting rectangles form one valid nested or disjoint region tree. If no unique valid pairing exists, the template is ambiguous and validation fails rather than guessing.

### 6.3 Lists

A list is not implicitly converted into comma-separated text or automatically spilled from a scalar expression. It is rendered using the same repeat primitive as any other block.

A vertical one-cell list uses a one-cell block with `direction="down"`. A horizontal list uses `direction="right"`.

This unifies list and table behavior:

- one-cell repeated block: list
- one-row repeated block: conventional table body
- multi-row repeated block: cards, grouped records, or formatted report sections
- one-column block repeated right: horizontal list or horizontal report

### 6.4 Why there is no `xl:each`

An earlier design used `xl:each`, inspired by Jxls commands such as `jx:each` stored in Excel Notes. The prefix distinguished workbook commands from expressions, but it adds terminology without helping the author. This specification instead proposes the familiar `{% for %}` form.

### 6.5 Headers, static rows, and prepared data

Headers are ordinary static template content. They do not participate in data binding. A displayed caption such as `Product Code` may be translated, merged, or renamed without changing an expression such as `{{ row.code }}` below it.

Static rows may appear before or after a repeat block. For example, the template may contain the literal rows `a`, `b`, and `c`, followed by a repeat block that appends additional rows.

The renderer preserves collection order exactly as supplied. Any sorting, grouping, joining, required-key insertion, or matching of input records to static rows happens in the platform's processing stage before rendering. The renderer never scans static cell values or headers to infer such behavior.

## 7. Spatial block rules

1. A block's opening tag is the first non-whitespace template token in its top-left cell.
2. Its closing tag is the last non-whitespace template token in its bottom-right cell.
3. Both boundary cells belong to the block after the directive text is removed.
4. A block must have a positive width and height.
5. Nested blocks must be fully contained by their parent.
6. Sibling blocks must be disjoint. Partial overlap is invalid.
7. A merged range must be fully inside or fully outside a block. A block boundary cannot bisect a merge.
8. Each opening marker must have exactly one spatially valid closing marker of the same directive type.
9. The set of paired rectangles must have exactly one valid nested/disjoint interpretation; ambiguity is a validation error.
10. Direction is either `down` or `right` in the initial language design.
11. The parser records source coordinates; rendering uses derived destination coordinates and never rewrites the source AST.

## 8. Expansion semantics

Direction answers where copies are placed. Shift mode answers which neighboring cells are displaced to make room.

The proposed vocabulary follows Excel's own insertion concepts.

### 8.1 `shift="rows"`

This is the default for `direction="down"`. The engine inserts complete worksheet rows for growth, so all content below the block moves.

For a collection of length `n`, the first item uses the source block and the engine inserts space for the remaining `n - 1` instances. A source block with multiple rows inserts the corresponding multiple of complete worksheet rows.

Two sibling blocks whose source row spans overlap cannot independently use `shift="rows"`. Their insertions would both claim the same worksheet rows without an isolation boundary. Such siblings must use `shift="cells"`, or be placed in an explicit region construct that defines how their completed layouts are combined. Validation rejects the ambiguous row-shift arrangement.

### 8.2 `shift="columns"`

This is the default for `direction="right"`. The engine inserts complete worksheet columns for growth, so all content to the right moves.

### 8.3 `shift="cells"`

- For `direction="down"`, only cells in columns intersecting the block are shifted down.
- For `direction="right"`, only cells in rows intersecting the block are shifted right.
- Content beside the repeated lane remains stationary.
- A collision or a partially intersected merged range is an error.

This mode supports independently growing blocks placed side by side, provided their shifted lanes do not collide with another protected region.

### 8.4 Explicit regions and isolation

An explicit rectangular `region` is a layout container and isolation boundary. It uses the same opposite-corner marker model as other structural constructs:

```text
A1:  {% region direction="down" shift="cells" %}
J10: {% endregion %}
```

The markers above define exactly `A1:J10`. Blank cells, styles, and worksheet used ranges never enlarge or shrink that source rectangle. Both markers are removed from rendered output.

`direction="down"` is the default and the only supported region direction in the first usable release. A horizontal or mixed-axis child is invalid inside a vertical region. A single region cannot grow both down and right.

The source rectangle is the region's minimum reserved allocation. Children are measured and positioned inside the region before the completed region is exposed to its parent:

- side-by-side child lanes contribute their maximum completed bottom edge, not the sum of their heights;
- vertically stacked children contribute their actual cumulative displacement;
- nested regions are measured from the inside out;
- output that fits inside the source rectangle consumes its reserved space and causes no external movement;
- when child output extends below the source bottom, the region's external growth is the completed bottom edge minus the source bottom.

The region's `shift` option controls only how that completed external growth affects cells outside the source rectangle:

- `shift="rows"` is the default and shifts complete worksheet rows below the region's source bottom;
- `shift="cells"` shifts only cells below the source bottom whose columns fall within the region's exact left-to-right column band. Cells in columns outside that band remain at their original rows.

For example, a completed `A1:J10` region that requires four additional rows moves content at `A20:J20` to `A24:J24` when `shift="cells"`, while content at `K20:P20` remains on row 20. Content already inside `A1:J10` participates in the region's internal layout; the engine does not shift the whole source rectangle as one pre-existing slab.

Children remain subject to the normal nested/disjoint rectangle rules. Partial overlaps, ambiguous marker pairing, output collisions, and merged ranges that cross a region or cell-shift boundary are errors. Row heights are worksheet-wide and cannot move independently for only part of a row; `shift="cells"` therefore moves cell content and cell formatting within its column band but does not translate worksheet row dimensions.

### 8.5 Empty collections

An empty collection retains exactly one instance of the source block.

- Structural directives are removed.
- Styles, dimensions, and merges are retained.
- Dynamic expressions that depend on the absent loop item render as blank.
- Static literal content is retained.
- The loop body is not considered to have executed, and no item variable is added to the ordinary evaluation scope.

A later explicit `empty` branch may replace this default instance with content such as “No records.”

### 8.6 Layout planning

Each AST node first reports its measured output size. Parent nodes then allocate destinations to children. This bottom-up measurement followed by top-down placement produces a coordinate transformation from source cells to output cells.

The transformation is used consistently for cells, merges, row heights, column widths, and
supported workbook objects whose individual feature contract requires coordinate transformation.
Template-authored charts, images, and static text shapes use the planned-anchor policies in
sections 13.5, 13.6, and 13.7.

## 9. Data and type system

The interpreter consumes one canonical, language-neutral value tree. The render context is its
root: a mapping from string names to canonical values. Template authors do not declare input types
and callers do not wrap ordinary values in labels such as `TypedValue(value, "table")`.

Runtime types determine which expression operations are valid. They never determine worksheet
layout. Template syntax and rectangular geometry decide whether the same ordered collection is
rendered as a one-cell list, conventional table body, multi-row cards, or another repeated block.

### 9.1 Canonical scalar values

- `null`
- string
- boolean
- integer
- finite floating-point number
- finite decimal number
- date
- timezone-naive datetime
- timezone-naive time

NaN, positive infinity, and negative infinity are not canonical numeric values. Timezone-aware
datetime and time values must be converted under a platform policy before rendering because XLSX
does not retain their timezone semantics.

### 9.2 Records and ordered collections

A record is a mapping whose keys are strings and whose values are canonical values. Property syntax
such as `customer.name` reads record keys; it does not invoke attributes on arbitrary objects.

An ordered collection is a finite Python-style list or tuple of canonical values. Collection order
is preserved exactly. Sets and frozensets are rejected because they are not a deterministic input
contract. Generators, cursors, query objects, and other arbitrary iterators are rejected because the
complete render must be validated and measured before workbook mutation.

Collections may contain scalars, records, other ordered collections, or a mixture of canonical
values. A list of records is table-shaped data, but it is not a distinct semantic `table` type. Its
template repeat rectangle controls its presentation.

The canonical tree must be acyclic. Reusing one record or collection in multiple branches is
allowed; a reference that leads back to one of its ancestors is rejected.

### 9.3 Platform adapters and tabular data

Platform-specific values are normalized before expression evaluation. Pandas and Polars data
frames, Arrow tables or batches, DuckDB results, ORM models, and similar objects are not canonical
core values. A platform-facing API may select a registered adapter from the concrete runtime type,
but the adapter must produce the canonical value tree before invoking the interpreter.

A tabular adapter must:

- preserve input row order;
- require unique string column names;
- omit an index unless the platform explicitly materializes it as a column;
- convert library-specific null, NaN, and not-a-time sentinels to `null`;
- convert library-specific scalar objects to canonical scalar values;
- reject values whose timezone, nested-object, or other semantics cannot be preserved.

Adapters are an integration concern. Their presence does not allow the core to infer headers,
sorting, grouping, layout, or a special table rendering mode.

Adapters are caller-supplied and scoped to one normalization or render operation. The core has no
mutable process-global adapter registry. Each adapter declares one source runtime type and a pure
conversion function. Canonical values are recognized before adapter lookup, so an adapter cannot
override the meaning of strings, numbers, records, lists, or tuples.

For a non-canonical value, adapter resolution selects the unique most-specific matching source
type. Exact-type registration is more specific than a base type. Duplicate registrations for the
same source type and matches through unrelated source types are errors; registration order is never
a tie-breaker.

Adapter output is recursively normalized at the same context path. It may contain other adapted
values. An adapter exception, unsupported output, direct or indirect conversion cycle, or ambiguous
adapter match produces a diagnostic and no normalized context.

#### 9.3.1 Bundled Polars adapter

The optional `polars` package extra provides `polars_adapters()`. The function returns caller-scoped
`TypeAdapter` instances and does not register mutable global state. Importing the core package does
not import or require Polars.

The first bundled adapter accepts eager `polars.DataFrame` values only. It does not accept or collect
`LazyFrame` values, execute queries, or treat `Series` as an implicit collection. Callers must make
those operations explicit before rendering.

For an accepted frame, the adapter:

- preserves frame row order and declared column order;
- requires unique string column names and emits one record per row;
- introduces no index field because Polars has no implicit row index;
- converts Polars null values and floating-point NaN values, including nested NaN values, to
  canonical `null`;
- materializes Python-native scalar values and delegates their recursive validation to canonical
  normalization at the original context path.

Polars nanosecond `Datetime` values, `Time` values, and timezone-bearing `Datetime` values are
rejected before materialization. Python temporal objects cannot preserve nanosecond precision, and
timezone-aware temporal values are outside the canonical model. Library values that have no
canonical scalar representation, including duration, binary, or object values, are rejected by the
ordinary canonical-value diagnostics after conversion. The adapter never stringifies unsupported
values to make them renderable.

### 9.4 Normalization result and immutability

`normalize_context(raw_context, adapters)` validates and converts the complete input before template
evaluation. On success it returns an immutable normalized context:

- records become read-only mappings;
- lists and tuples become tuples;
- canonical scalar objects are retained;
- shared input subtrees may be copied, and input object identity is not part of the language;
- caller-owned mappings and lists are never mutated.

Normalization aggregates independent problems where traversal can continue safely. A result with
any diagnostic contains no usable context. Rendering entry points normalize raw input automatically;
an already-normalized context can be reused across worksheets without another traversal.

### 9.5 Rejection boundary

The following are not canonical values:

- mappings with non-string keys;
- sets and frozensets;
- arbitrary iterators or generators;
- non-finite floating-point or decimal numbers;
- timezone-aware datetime or time values;
- bytes, complex numbers, functions, modules, and arbitrary class instances;
- cyclic records or collections;
- unadapted platform-specific tabular or model objects.

Canonical-context validation reports stable diagnostics with a path such as
`context.lines[2].amount`. It validates the entire supplied context, including values unused by a
particular template, before evaluation begins.

### 9.6 Cell assignment

- A sole expression preserves its canonical scalar type.
- Mixed literal and expression content is converted to text.
- `null` produces a blank cell by default.
- A collection or mapping used as a scalar is a type error unless passed through a registered
  filter that produces a scalar.
- A `for` expression requires an ordered collection; it does not iterate record keys.
- The core language has no sorting operation in the first release. The platform prepares final order.

### 9.7 Missing values

A missing name or property is an error by default. A `default` filter may handle absence intentionally:

```text
{{ customer.phone | default("—") }}
```

This distinguishes missing data from present-but-null data.

## 10. Expression language

The expression grammar is deliberately smaller than Python and Jinja2.

Proposed initial features:

- names: `invoice`
- property lookup: `invoice.customer.name`
- mapping/list indexing: `items[0]`
- literals: strings, numbers, booleans, and null
- comparisons
- boolean operators
- numeric unary operators: `+` and `-`
- numeric binary operators: `+`, `-`, `*`, and `/`
- parentheses
- filters with literal arguments
- a small set of pure registered functions, if later justified

Prohibited:

- attribute or key names beginning with `_`
- imports
- assignment
- comprehensions
- lambdas
- reflection or access to Python object internals
- arbitrary method calls
- filesystem, network, environment, or process access

The engine owns an expression AST and evaluator. Using a third-party parser internally is acceptable, but public semantics cannot depend accidentally on Python evaluation behavior.

## 11. Built-in filters

The executable set remains small:

- `default(value)`
- `string`
- `upper`
- `lower`
- `date(format)`
- `join(separator)`
- `sum` and `sum(column)`
- `min` and `min(column)`
- `max` and `max(column)`
- `count` and `count(column)`

Filter names and argument contracts are validated during compilation. Filter arguments are literals;
runtime expressions cannot dynamically select formatting rules. Unknown filters, invalid argument
counts, and arguments of the wrong literal type are template errors rather than deferred guesses.

### 11.1 Date text formatting

`date(format)` is a text-formatting filter for the cases where a date must be embedded in literal
cell content:

```text
For the month ending {{ report_date | date("dd mmmm yyyy") }}
{{ report_date | date("yyyy-mm") }}
```

The filter requires exactly one string-literal format argument. It accepts canonical `date` and
`datetime` values and always returns text. A `datetime` contributes its calendar-date fields; its
time portion is not rendered. Strings are not parsed as dates, numbers are not interpreted as Excel
date serials, and `null` is a type error. A missing input remains a missing-value error.

The date format is a case-insensitive, deliberately constrained subset of Excel date-format codes:

| Code | Meaning |
| --- | --- |
| `yyyy`, `yy` | Four- or two-digit year |
| `mmmm`, `mmm` | Full or abbreviated English month name |
| `mm`, `m` | Zero-padded or unpadded month number |
| `dddd`, `ddd` | Full or abbreviated English weekday name |
| `dd`, `d` | Zero-padded or unpadded day of month |

Spaces and ordinary punctuation are literal. Double-quoted sections and a backslash-escaped next
character introduce explicit literal text within the format. Empty formats, unsupported field
letters or field widths, unterminated quoted sections, dangling escapes, braces, Python
`strftime` percent codes, Excel multi-section formats, colors, conditions, and spacing/fill codes
are compilation errors. Month and weekday names are deterministic English text in this release;
locale selection is deferred to a separate language design.

The format is parsed into an engine-owned immutable format AST during expression compilation. It is
not passed through to Python `strftime` or interpreted by the XLSX writer. Date formatting does not
change geometry, scope, shifting, or block measurement beyond contributing its final text length.

Because `date(format)` returns text, even a sole filtered expression produces an Excel text cell and
the source cell's number format has no effect on its display. To retain an actual Excel date, authors
use an unfiltered sole expression such as `{{ report_date }}` and set the desired number format on
the template cell. The template cell remains authoritative for native Excel presentation.

The `datetime`, `number`, and `money` filters remain proposed. Their value-versus-presentation
semantics must be resolved independently before implementation; the behavior of `date` must not be
silently generalized to them.

### 11.2 Collection aggregates

The aggregate filters reduce an ordered collection to one canonical scalar. `sum`, `min`, and
`max` accept either a collection of numeric values or one literal top-level record key. Without an
argument, each collection item is a candidate value:

```text
{{ amounts | sum }}
{{ amounts | min }}
{{ amounts | max }}
```

With one string-literal argument, each collection item must be a record and the argument identifies
one top-level record key to aggregate:

```text
{{ lines | sum("amount") }}
{{ lines | min("amount") }}
{{ lines | max("amount") }}
```

`count` without an argument returns the number of collection items, including items that are
`null`. `count(column)` requires record items and returns the number of present, non-null values at
that key:

```text
{{ lines | count }}
{{ lines | count("amount") }}
```

Every column argument is a literal key, not an expression or dotted path, and may not begin with
`_`. A missing key is a missing-value error that identifies the failing collection index and key.
A non-record collection item in column mode is a filter type error. For `sum`, `min`, and `max`, a
boolean or any non-numeric selected value is also a filter type error; strings are never coerced to
numbers. `count(column)` may count a non-null value of any canonical type.

Numeric aggregates ignore present `null` values. `sum` returns integer `0` for an empty or all-null
selection. `min` and `max` return `null` because no extremum exists. `count` returns integer `0` for
an empty collection or when `count(column)` selects only nulls. A numeric result containing only
integers remains an integer; integers may be combined with either finite floats or finite decimals,
producing that wider numeric type. Mixing a float and a decimal in one aggregate is a filter type
error rather than an implicit lossy conversion.

Aggregates preserve input order and perform no sorting, grouping, predicate filtering, distinct
selection, or reconciliation. Each aggregate expression traverses its input independently and does
not mutate or cache the canonical context. It produces one scalar, so it has no effect on lexical
scope, rectangular geometry, layout measurement, shifting, or workbook-writer behavior. Traversal
is bounded by the collection limits already enforced during context normalization.

### 11.3 Numeric arithmetic

The expression language supports unary `+` and `-`, followed by multiplicative `*` and `/`, then
additive `+` and `-` in standard precedence order. Arithmetic binds more tightly than comparisons,
which bind more tightly than boolean operators. Operators at the same precedence are
left-associative. Parentheses may make any order explicit:

```text
{{ quantity * unit_price + tax }}
{{ (subtotal - discount) / installment_count }}
{{ -adjustment }}
```

The filter pipeline continues to apply to the complete expression on its left. A filtered value
used as one arithmetic operand must therefore be parenthesized:

```text
{{ (lines | max("high")) - (lines | min("low")) }}
```

Arithmetic accepts canonical integers, finite floats, and finite decimals; booleans, strings,
dates, and collections are not numbers. `+` never concatenates strings or collections. Unary signs
preserve the operand type. For binary operations, integers may combine with either floats or
decimals, producing that wider family, while floats and decimals may not mix implicitly. Division
of integer or float operands returns a float. Division involving a decimal and no float returns a
decimal.

If either operand of a binary arithmetic expression is present `null`, the result is `null`; null
is never coerced to zero. Division by zero is an evaluation error. Any operation or aggregate that
would produce a non-finite float or decimal is also an evaluation error, preserving the canonical
finite-number boundary. Arithmetic returns one scalar and has no effect on scope or layout.

## 12. Formulas

Formula copying, translation, and calculation are outside the first release. A formula must never be treated as an ordinary string and copied into a new position with potentially incorrect references.

The first-release validator rejects a formula when a structural transformation would copy it, move it, or require its references to be adjusted. Formula cells that can be proven unaffected may be preserved verbatim. If the implementation cannot prove that a formula is unaffected, it must reject the template rather than silently risk changing its meaning.

Full formula support may be designed later as a separate transformation pass over parsed Excel references. It must not use regular-expression substitution over formula text.

## 13. Styles and merged cells

### 13.1 Styles

The template cell is authoritative for presentation. Every rendered cell copies the effective
direct formatting of its source template cell. The renderer does not infer formatting from the
value, field name, headers, or neighboring cells.

Repeated and shifted cells copy:

- font
- fill
- border
- alignment
- number format
- protection

A formatted blank cell is a material template cell. It participates in source-to-destination
mapping and is copied even though its value is blank. This is required for continuous fills,
borders, alignment, protected entry areas, and merged-range presentation.

A cell containing only structural directive text remains a material blank cell after the
directive is removed. Its direct formatting is copied exactly like any other formatted blank;
removing template syntax must not expose Excel's default fill or borders.

Formatting is copied exactly per source cell and per repeated instance. The engine does not
recompute outer borders, alternating stripes, first/last-row styles, or other contextual table
effects. Such behavior requires an explicit future language feature rather than inference.

The source number format remains authoritative for rendered values. Authors format placeholder
cells as dates, currency, percentages, or other Excel formats in the template. The renderer does
not select a number format from a variable name or runtime type.

Styles should be reused/deduplicated where possible to prevent excessive workbook style records.
The first production adapter guarantees the effective cell appearance, not preservation of named
style authoring identity.

### 13.2 Row and column dimensions

For `shift="rows"`, every output row maps to a source row. Repeated rows copy explicit height,
hidden state, outline level, collapsed state, and safely representable row-level direct style.
Static rows retain the properties of their mapped source rows after movement.

Automatic/default row height remains automatic. The engine does not measure text or calculate an
auto-fit height.

`shift="cells"` cannot apply different row heights to independent column lanes because Excel row
height is worksheet-wide. A cell-shift block that would repeat a row with an explicit custom height
is rejected in the first production adapter. It may use the worksheet's existing/default row
heights instead.

Vertical rendering does not create columns. Existing column width, hidden state, outline level,
collapsed state, best-fit state, and safely representable column-level style are preserved. A later
horizontal layout phase must introduce explicit source-to-destination column mappings.

### 13.3 Merged cells

- A merge fully contained within a repeated block is reproduced for every block instance.
- A merge outside a block is transformed when its cells are shifted.
- A block boundary or shift lane may not split a merged range.
- Only the top-left cell of a merged range may receive a rendered value.
- Overlapping destination merges are fatal layout errors.

Merged ranges are explicit entries in the render plan. The workbook writer must not infer merge
copies from blank cells or from the final value grid.

### 13.4 Range-bound and unsupported workbook formatting

Conditional formatting and data validation are range-bound workbook features, not direct cell
styles. In the first production adapter they are preserved only when the sheet contains no
structural layout transformation. If a repeat or condition could require a range to move, copy,
resize, contract, or split, rendering is rejected until that feature has its own validated render
plan representation.

The same safety rule applies to other coordinate-dependent workbook objects. Native Excel Tables
and unsupported drawing objects or anchors are rejected by the first production adapter. The
narrow template-authored chart, embedded-image, and static text-shape profiles in sections 13.5,
13.6, and 13.7 are the only supported drawing exceptions.
Hyperlinks and comments are preserved only when their cell has exactly one unchanged destination;
copying or moving them is rejected until an explicit policy is implemented.

### 13.5 Template-authored worksheet charts

The first chart profile preserves an existing chart authored on an ordinary worksheet. Rendering
does not create a chart or infer categories, values, titles, series, or presentation from template
data. The chart stored in the template remains authoritative.

Chart data references are fixed output coordinates. Every chart formula must be one direct,
concrete A1 cell or rectangular range on a worksheet in the same workbook. The adapter preserves
that formula verbatim; it does not expand, contract, translate, or otherwise rewrite the range. For
example, a chart authored against `A2:A10` and `B2:B10` continues to use those nine rows when a
repeat renders more than nine items. Items after row 10 are rendered but not plotted. When fewer
than nine items are rendered, the remaining referenced cells are blank and the chart's authored
empty-cell display policy controls their presentation.

The supported first profile is limited to ordinary two-dimensional area, bar or column, line, pie,
and scatter charts with direct worksheet references. Multiple charts and multiple series are
allowed. Chart sheets, pivot charts, combined charts, three-dimensional charts, external-workbook
references, defined names, structured Excel Table references, and dynamic reference formulas are
rejected. Native Excel Tables remain unsupported independently of chart use.

Chart anchor movement follows the completed render plan while chart size remains fixed. An absolute
anchor has no cell marker and remains at its authored position. A one-cell anchor moves its marker
to that source cell's sole planned destination while retaining its offsets and extent. A two-cell
anchor may translate when both source markers have exactly one destination and both destinations
apply the same row and column delta; its marker offsets and size are retained. Different marker
deltas would resize the chart and are rejected in this profile.

An anchor marker that is removed or copied by a repeat, condition, region, or lane is a validation
error. Charts are never copied per repeat instance. A chart below a growing table moves downward
when its cell-based anchor is shifted by `shift="rows"`, or when its anchor lies in the affected
band of a `shift="cells"` layout. A chart beside a cell-shift lane stays where authored when its
anchor remains outside the lane.

Chart style, axes, legend, titles, labels, blank-cell policy, and other state representable by the
supported `openpyxl` chart object are preserved through save and reload. Chart value caches are not
calculated by the renderer; consumers render the chart from its fixed worksheet references. A
worksheet drawing part containing a connector, a shape outside the static profile in section 13.7,
an unsupported graphic frame, or another unsupported drawing object is rejected rather than
partially preserved. Supported pictures and static text shapes may share an ordered drawing part
with supported charts.

### 13.6 Template-authored worksheet images

The first image profile preserves embedded pictures already authored on an ordinary worksheet.
Rendering does not create pictures from context data, read filenames or URLs from expressions, or
repeat one picture for every block instance. PNG and JPEG media embedded inside the XLSX package
are supported. Linked or external pictures, SVG, GIF, BMP, TIFF, WMF, EMF, worksheet backgrounds,
header/footer pictures, grouped drawings, and data-driven image directives are rejected or remain
outside this profile. Static text shapes are governed separately by section 13.7.

Image movement follows the same completed source-to-destination mapping used for chart anchors. An
absolute anchor remains at its authored position. A one-cell anchor moves its marker to that source
cell's sole planned destination while preserving its cell offset and physical extent. A two-cell
anchor may translate only when both source markers have exactly one destination and both receive
the same row and column delta. Different marker deltas would resize the picture and are rejected.

An image below a growing `shift="rows"` block moves downward. Under `shift="cells"`, an image moves
only when its cell marker lies in the affected column band; an image beside that band stays where
authored. The marker coordinates, rather than the picture's visible pixel bounds, determine lane
membership. An anchor marker removed by a condition or copied by a repeat is a validation error.
Images are never copied per repeat instance.

The adapter preserves the embedded media bytes, anchor type, marker offsets, physical extent,
picture frame, cropping, rotation or flips, non-visual name and description, and lock/print flags
that are representable by the supported DrawingML model. It preserves the original ordering of
supported charts and pictures in each worksheet drawing so their stacking order is not silently
regrouped by object type or anchor type. Extra drawing relationships such as picture hyperlinks,
alternate vector representations, or other unsupported effects cause explicit rejection rather
than partial preservation.

### 13.7 Template-authored static text shapes

The first text-shape profile preserves ordinary ungrouped DrawingML shapes that contain authored
text, including text boxes and text-bearing decorative presets such as rectangles, rounded
rectangles, arrows, and callouts. The shape remains an editable Excel object. The adapter preserves
its authored rich-text runs, paragraphs, fills, outlines, geometry, rotation, margins, wrapping,
autofit or overflow settings, locks, print flags, name, description, and other inline DrawingML
state rather than converting it to cells or an image.

Shape text is opaque template content in this profile. Text such as `{{ customer.name }}` or
`{% for item in items %}` inside a shape is preserved literally and is not lexed, evaluated, or
used to define a block. Rendering never changes the text or measures it, and it does not resize a
shape to fit rendered workbook content. The shape's existing Excel wrap, autofit, clipping, and
overflow behavior remains authoritative.

Movement follows the same completed source-to-destination mapping used for chart and image
anchors. An absolute anchor remains at its authored position. A one-cell anchor moves its marker
to that source cell's sole planned destination while retaining its offsets and physical extent. A
two-cell anchor may translate only when both markers have exactly one destination and both receive
the same row and column delta. Different marker deltas would resize the shape and are rejected.
Markers removed or copied by structural rendering are also rejected; shapes are never repeated.
For `shift="cells"`, marker coordinates determine lane membership, not the shape's visible bounds.

Supported shapes may share one drawing part with supported charts and pictures. Their source
stacking order is preserved. Every supported shape must be a relationship-free `xdr:sp` with a
normal text body. Grouped shapes, connectors, WordArt or text-warp shapes, dynamic text fields,
cell-linked text, macros, shape hyperlinks, image-filled shapes, linked or external content,
SmartArt, form controls, OLE objects, and legacy VML drawings are outside this profile and cause
explicit rejection rather than partial preservation.

## 14. Conditions and advanced constructs

Conditions are required in the first usable language release and are represented in the AST from the beginning.

The first-release syntax uses vertically stacked, equal-width branches. For a true branch in `A10:F11` and an else branch in `A12:F13`:

```text
A10: {% if invoice.taxable %}Tax
F11: {{ invoice.tax }}{% else %}
A12: Not taxable
F13: {% endif %}
```

The `if` marker is the top-left corner of the true branch. The `else` marker is the bottom-right corner of the true branch. The false branch starts in the same column as the `if` marker on the next row and ends at the `endif` marker. Therefore both branches have the same width and are adjacent without an unowned row between them.

Without an `else`, the `if` and `endif` markers are the top-left and bottom-right corners of the single conditional body.

Only the selected branch contributes cells and measured height. Removing the unselected branch closes the vertical gap according to the containing shift policy. Each branch follows the same containment, merge, measurement, shift, and collision rules as a repeat body.

Potential later constructs:

- an empty branch for repeats
- grouping and subtotals
- page-break controls
- repeated print headers
- data-driven image directives
- sheet repetition

These must be new AST node types, not special-case string replacement hooks.

## 15. Compiler and interpreter architecture

The engine is organized as a front end, semantic model, layout planner, and workbook back end.

### 15.1 Workbook reader

Loads the source workbook into an immutable template model containing:

- worksheets and dimensions
- cell values and formulas
- style references
- row and column properties
- merged ranges
- supported template-authored charts, images, and static text shapes with their planned anchors
- supported workbook metadata

The reader is the only layer directly coupled to `openpyxl` input objects.

### 15.2 Lexer

Tokenizes each text cell into literal text, output expressions, opening directives, and closing directives. Every token carries a source span containing workbook, sheet, cell coordinate, and character offsets.

### 15.3 Parser

Parses tokens into expression nodes and structural markers according to a published grammar. Syntax errors include exact worksheet and cell locations.

### 15.4 Spatial linker

Pairs structural markers by directive type and geometry, derives exact rectangles, and builds a region tree. Block IDs are not part of the current authoring language. This is separate from textual parsing because spreadsheet blocks exist in two dimensions.

### 15.5 Semantic validator

Checks:

- uniquely pairable structural markers
- valid block geometry
- containment and non-overlap
- legal direction/expansion combinations
- merged-range boundaries
- name scoping and loop variables
- expression and filter validity
- supported workbook features

Validation can run without render data for structural checks and with a context/schema for name and type checks.

### 15.6 AST and intermediate representation

Representative AST nodes:

- `WorkbookNode`
- `SheetNode`
- `RegionNode`
- `CellTemplateNode`
- `TextNode`
- `OutputNode`
- `ForNode`
- `IfNode`
- `ElseNode`
- later: `ImageNode`

Structural nodes contain a source rectangle, options, scope, and child nodes. Expression nodes contain parsed expressions rather than raw strings.

Evaluation produces a render/layout IR rather than immediately writing cells. The IR records evaluated values, measured rectangles, instance paths, source-to-destination mappings, merges, and planned shifts.

### 15.7 Evaluator

Evaluates expressions against lexical scopes. A loop adds its item variable and loop metadata to a child scope without mutating its parent.

Suggested loop metadata:

```text
loop.index       # 1-based
loop.index0      # 0-based
loop.first
loop.last
loop.length
```

### 15.8 Layout planner

Measures expanded nodes, allocates destination rectangles, produces coordinate transformations, and detects collisions. It has no dependency on live `openpyxl` worksheet mutation.

### 15.9 Workbook writer

Applies a validated render plan to a copy or newly constructed workbook model, then serializes it through `openpyxl`.

The writer performs no expression parsing and makes no layout decisions.

### 15.10 Output verification

After serialization, the engine should reopen the output with `openpyxl` as a basic package-integrity check. Test environments should also open representative fixtures in Excel or LibreOffice, but LibreOffice is not required at runtime.

## 16. Diagnostics

Diagnostics are first-class values, not only exception strings.

Each diagnostic includes:

- stable error code
- severity: error or warning
- human-readable message
- worksheet name
- cell or rectangle
- related location, when applicable
- optional hint
- AST/render instance path when the error occurs during evaluation

Examples:

- `E1001 UNTERMINATED_EXPRESSION`
- `E1104 INVALID_FILTER`
- `E1105 INVALID_DATE_FORMAT`
- `E1203 AMBIGUOUS_BLOCK_PAIRING`
- `E1205 PARTIAL_BLOCK_OVERLAP`
- `E1301 MISSING_VALUE`
- `E1302 COLLECTION_IN_SCALAR_CELL`
- `E1304 FILTER_TYPE_MISMATCH`
- `E1305 ARITHMETIC_TYPE_MISMATCH`
- `E1306 DIVISION_BY_ZERO`
- `E1307 NON_FINITE_EXPRESSION_NUMBER`
- `E1401 LAYOUT_COLLISION`
- `E1402 OVERLAPPING_ROW_SHIFTS`
- `E2104 MERGE_CROSSES_BLOCK_BOUNDARY`
- `E3101 FORMULA_REQUIRES_UNSUPPORTED_TRANSFORM`

Rendering is atomic: any error prevents an output workbook from being returned as successful.

## 17. Security

- Templates and data are treated as untrusted.
- Expression evaluation is allowlist-based.
- Rendering performs no network access.
- Rendering performs no template-directed filesystem access.
- ZIP/package size, cell count, nesting depth, collection length, and string length have
  configurable limits. Wall-clock cancellation and process-level memory limits belong to the host
  platform because they cannot be enforced deterministically inside a synchronous pure planner.
- The input file is never overwritten.
- External workbook links are rejected or handled under an explicit policy.
- Formula injection is a platform concern for user-supplied strings. Plain strings beginning with `=` must remain strings unless a trusted formula construct is explicitly used.

### 17.1 Resource-limit policy

`ResourceLimits` is one immutable configuration shared by normalization, layout planning, and the
XLSX adapter. Defaults are deliberately permissive for ordinary small business workbooks while
still bounding accidental or hostile input:

| Limit | Default |
| --- | ---: |
| Canonical context nesting depth | 64 |
| Canonical context nodes | 1,000,000 |
| Items in one mapping, list, or tuple | 100,000 |
| Characters in one input string | 1,000,000 |
| Total repeat iterations per worksheet | 100,000 |
| Planned cells per worksheet | 500,000 |
| Planned cells per workbook | 1,000,000 |
| Rendered rows per worksheet | 250,000 |
| Rendered columns per worksheet | 4,096 |
| Worksheets per workbook | 100 |
| Compressed XLSX package bytes | 50 MiB |
| Uncompressed XLSX package bytes | 250 MiB |
| ZIP members in one XLSX package | 10,000 |

The root context mapping has depth zero. Context-node and container limits apply after adapter
conversion to the canonical tree. An empty-repeat placeholder counts as one repeat iteration.
Planned-cell totals include material blank cells retained for formatting.

The XLSX format's absolute bounds are separate and cannot be raised or disabled: 1,048,576 rows,
16,384 columns, and 32,767 characters in a cell text value. A lower configured resource limit is
checked first. Runtime strings are checked after expression interpolation, before the writer can
silently truncate them.

Resource-limit diagnostics are fail-fast because continuing traversal or planning would defeat the
safety boundary. Ordinary syntax, semantic, and canonical-value diagnostics continue to aggregate
where safe. Every limit failure produces no normalized context, render plan, or published workbook.

`NormalizedContext` retains immutable summary statistics. Reusing it under equal or looser limits
requires no traversal; stricter limits are checked against those statistics.

The XLSX adapter checks the source package before `openpyxl` loads it and checks the serialized
temporary output before publication. Compressed size, declared uncompressed ZIP size, member count,
and sheet count are bounded without interpreting worksheet layout.

## 18. API direction

The precise host API is intentionally deferred, but the core boundary should accept workbook bytes or a seekable binary stream plus normalized data and options, and return workbook bytes plus diagnostics and metadata.

File paths are convenience adapters, not the core abstraction.

The core operations are conceptually:

```text
normalize_context(raw_context, adapters, limits) -> NormalizationResult
compile(template) -> CompiledTemplate
validate(compiled_template, optional_schema) -> Diagnostics
render(compiled_template, normalized_context, limits, options) -> RenderResult
```

A compiled template may be cacheable if it does not retain mutable `openpyxl` objects.

## 19. Testing strategy

### 19.1 Language tests

- lexer golden tests by cell content
- parser and expression AST tests
- invalid grammar diagnostics
- block pairing and spatial containment tests
- name scope and type tests

### 19.2 Layout tests

- source-to-destination coordinate mapping
- vertical and horizontal repeats
- one-cell, one-row, one-column, and multi-row blocks
- nested blocks
- side-by-side blocks
- empty and singleton collections
- cell, row, and column expansion
- collision detection
- merged-range transforms

Most layout tests should operate without `openpyxl`.

### 19.3 Workbook integration tests

Fixture workbooks verify values, types, style IDs/semantics, dimensions, merges, formula
rejection/preservation boundaries, fixed chart references, planned chart/image/text-shape anchors,
static rich shape text, embedded media bytes and drawing order, and package integrity after
save/reload.

Where workbook XML affects correctness, tests may inspect selected OOXML parts. XML manipulation is not the primary render strategy.

### 19.4 Property and fuzz tests

- parser never crashes on arbitrary cell strings
- valid region trees never produce overlapping output allocations
- coordinate transformations preserve ordering
- render output respects configured size limits
- load/save/reload succeeds for supported fixture workbooks

## 20. Proposed implementation phases

### Phase 0: executable design model

- grammar and AST
- spatial block linker
- rectangular `if` and `else` model
- diagnostics model
- pure layout planner
- in-memory tests without workbook mutation

Phase 0 accepts a two-dimensional in-memory worksheet representation and produces a deterministic AST and destination-cell render plan. The Phase 0 core must not import `openpyxl`.

### Phase 1: scalar renderer

- `.xlsx` load/save
- scalar expressions and types
- formatting and scalar-reduction filters
- template validation
- output integrity reload

### Phase 2: vertical repeat blocks

- `direction="down"`
- `shift="rows"` and `shift="cells"`
- vertical `region` containers and explicit external shift boundaries
- styles, row heights, and merges
- empty collections and nested vertical repeats

### Phase 3: conditions

- rectangular `if` and `else`
- branch measurement and shifting
- nesting conditions and repeats

Completion of Phase 3 constitutes the first usable language release.

### Phase 4: horizontal and mixed layouts

- `direction="right"`
- `shift="columns"` and `shift="cells"`
- horizontal region containers
- mixed-direction nesting and collision rules

### Phase 5: richer reports

- empty branches
- grouping/subtotals as justified by real templates
- additional explicitly supported workbook features

## 21. Resolved region design

Vertical isolation regions use the explicit opposite-corner syntax and completed-layout behavior defined in section 8.4. Horizontal regions and mixed-direction nesting remain deferred to Phase 4; they must not be inferred from the vertical rules without a separate language-design review.

## 22. Library rationale and known constraints

`openpyxl` is the proposed adapter because the engine must load and modify an existing `.xlsx` workbook. XlsxWriter is a write-only generator and cannot use an existing workbook as a template.

`openpyxl` does not calculate formulas and does not preserve every possible OOXML object. The supported workbook-feature profile must therefore be explicit. Features outside that profile should be rejected or warned about during validation rather than silently promised.

References:

- [openpyxl tutorial and preservation warnings](https://openpyxl.readthedocs.io/en/stable/tutorial.html)
- [XlsxWriter project description and write-only constraint](https://xlsxwriter.com/)
- [Jxls rectangular area model](https://jxls.sourceforge.net/reference/how_it_works.html)
- [Jxls Excel markup](https://jxls.sourceforge.net/jxls-2.x/reference/excel_markup.html)
