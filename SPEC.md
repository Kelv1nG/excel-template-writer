# Excel Template Engine Specification

Status: Draft 0.1
Scope: Product and language design; not yet an implementation contract

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
- Pivot tables, slicers, macros, form controls, and unsupported drawing objects
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

Two sibling blocks whose source row spans overlap cannot independently use `shift="rows"`. Their insertions would both claim the same worksheet rows without an isolation boundary. Such siblings must use `shift="cells"`, or later be placed in an explicit region construct that defines how their completed layouts are combined. Validation rejects the ambiguous row-shift arrangement.

### 8.2 `shift="columns"`

This is the default for `direction="right"`. The engine inserts complete worksheet columns for growth, so all content to the right moves.

### 8.3 `shift="cells"`

- For `direction="down"`, only cells in columns intersecting the block are shifted down.
- For `direction="right"`, only cells in rows intersecting the block are shifted right.
- Content beside the repeated lane remains stationary.
- A collision or a partially intersected merged range is an error.

This mode supports independently growing blocks placed side by side, provided their shifted lanes do not collide with another protected region.

### 8.4 Explicit regions and isolation

An explicit rectangular `region` construct will act as a layout container and isolation boundary. A child is measured and positioned inside its nearest region; the completed region is then handled as one unit by its parent.

This provides a principled way to build independently growing side-by-side report sections. The exact author syntax and overflow policy for regions remain open design work; isolation must not be implemented through guessed blank areas.

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

The transformation is used consistently for cells, merges, row heights, column widths, and other supported workbook objects.

## 9. Data and type system

The render context is a mapping from string names to supported values.

### 9.1 Supported values

- null
- string
- boolean
- integer
- finite floating-point number
- decimal number
- date
- datetime
- time, if `openpyxl` round-tripping is reliable
- mapping with string keys
- ordered collection of supported values

Platform-specific objects must be normalized before expression evaluation. Pandas data frames, ORM models, query objects, and arbitrary iterators are not part of the core language contract.

Collections are ordered, and iteration preserves their supplied order. The core language has no sorting operation in the first release. The platform must prepare the final order before rendering.

### 9.2 Cell assignment

- A sole expression preserves its scalar type.
- Mixed literal and expression content is converted to text.
- `null` produces a blank cell by default.
- A collection or mapping used as a scalar is a type error unless passed through a registered formatting filter.
- NaN and infinity require an explicit policy because they are not ordinary Excel numeric values.

### 9.3 Missing values

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

## 11. Proposed built-in filters

The first set should remain small:

- `default(value)`
- `string`
- `upper`
- `lower`
- `date(format)`
- `datetime(format)`
- `number(format)`
- `money(format_or_currency)`
- `join(separator)`

Some apparent formatting filters may instead set Excel number formats while retaining numeric cell values. The filter specification must say whether a filter changes the value, the cell format, or both.

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

The same safety rule applies to other coordinate-dependent workbook objects. Native Excel Tables,
drawings, charts, images, and unsupported anchors are rejected by the first production adapter.
Hyperlinks and comments are preserved only when their cell has exactly one unchanged destination;
copying or moving them is rejected until an explicit policy is implemented.

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
- images
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
- `E1203 AMBIGUOUS_BLOCK_PAIRING`
- `E1205 PARTIAL_BLOCK_OVERLAP`
- `E1301 MISSING_VALUE`
- `E1302 COLLECTION_IN_SCALAR_CELL`
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
- ZIP/package size, cell count, nesting depth, collection length, string length, and render time must have configurable limits.
- The input file is never overwritten.
- External workbook links are rejected or handled under an explicit policy.
- Formula injection is a platform concern for user-supplied strings. Plain strings beginning with `=` must remain strings unless a trusted formula construct is explicitly used.

## 18. API direction

The precise host API is intentionally deferred, but the core boundary should accept workbook bytes or a seekable binary stream plus normalized data and options, and return workbook bytes plus diagnostics and metadata.

File paths are convenience adapters, not the core abstraction.

The core operations are conceptually:

```text
compile(template) -> CompiledTemplate
validate(compiled_template, optional_schema) -> Diagnostics
render(compiled_template, context, options) -> RenderResult
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

Fixture workbooks verify values, types, style IDs/semantics, dimensions, merges, formula rejection/preservation boundaries, and package integrity after save/reload.

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
- formatting filters
- template validation
- output integrity reload

### Phase 2: vertical repeat blocks

- `direction="down"`
- `shift="rows"` and `shift="cells"`
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
- side-by-side regions
- mixed-direction nesting and collision rules

### Phase 5: richer reports

- empty branches
- grouping/subtotals as justified by real templates
- optional images and other explicitly supported workbook features

## 21. Open design decisions

The following decisions need confirmation before this becomes version 1.0 of the language specification:

1. What is the exact author syntax and overflow behavior for isolation regions?
2. What practical workbook and collection size should the default resource limits target after benchmarks exist?

## 22. Library rationale and known constraints

`openpyxl` is the proposed adapter because the engine must load and modify an existing `.xlsx` workbook. XlsxWriter is a write-only generator and cannot use an existing workbook as a template.

`openpyxl` does not calculate formulas and does not preserve every possible OOXML object. The supported workbook-feature profile must therefore be explicit. Features outside that profile should be rejected or warned about during validation rather than silently promised.

References:

- [openpyxl tutorial and preservation warnings](https://openpyxl.readthedocs.io/en/stable/tutorial.html)
- [XlsxWriter project description and write-only constraint](https://xlsxwriter.com/)
- [Jxls rectangular area model](https://jxls.sourceforge.net/reference/how_it_works.html)
- [Jxls Excel markup](https://jxls.sourceforge.net/jxls-2.x/reference/excel_markup.html)
