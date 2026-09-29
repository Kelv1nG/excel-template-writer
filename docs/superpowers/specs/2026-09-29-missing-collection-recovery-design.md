# Missing Collection Recovery Design

**Date:** 2026-09-29

**Status:** Proposed for review

## Intent

Allow `render_sheet()` and `render_workbook()` to publish a useful result when a collection needed
by a repeat or a collection-producing output expression is missing or unavailable, while keeping
strict rendering as the default and keeping genuine template, type, data-shape, and layout errors
fatal.

The feature must not add syntax to workbook templates. Existing templates continue to use ordinary
expressions such as:

```jinja
{% for item in tables["some_table"] %}
{{ item.name }}
{% endfor %}
```

and:

```jinja
{{ tables["some_table"] | sum("some_col") }}
```

The render-scoped `MissingOutputPolicy.ERROR`, `.BLANK`, and `.PRESERVE` values govern recovery.

## Goals

- Keep `ERROR` as the default and preserve its current strict behavior.
- Let `BLANK` publish a workbook with an unavailable repeat's source rectangle blanked in place.
- Let `PRESERVE` publish a workbook with an unavailable repeat's exact source rectangle retained.
- Let `BLANK` and `PRESERVE` recover a collection filter whose receiver is canonical `null`.
- Emit a warning for every recovered repeat or output occurrence.
- Keep wrong non-null types and all independent errors fatal.
- Preserve the pipeline boundary: evaluation and layout decide behavior; the XLSX writer only
  applies a validated render plan.

## Non-goals

- No template-level policy/filter/directive syntax.
- No automatic conversion of all canonical `null` values to empty collections.
- No change to `if` conditions: missing conditions remain strict, while present null retains its
  existing falsey behavior.
- No change to ordinary empty-collection semantics.
- No inference of tables or repeat capacity from blank cells, styles, or worksheet used ranges.
- No recovery from malformed templates, invalid non-null data, arithmetic failures, collisions,
  resource limits, or unsupported workbook behavior.

## Terminology

An **unavailable collection** is either:

1. a collection expression that raises `MissingValueError` because a required name, key, property,
   or index is absent; or
2. a collection expression or collection-filter receiver that successfully evaluates to canonical
   `null` (`None`).

An empty ordered collection (`[]` or `()`) is available data with zero items. It is not an
unavailable collection and produces no recovery warning.

A present non-null value of the wrong type is invalid data, not an unavailable collection.

## Public semantics

### Policy matrix

| Construct | `ERROR` | `BLANK` | `PRESERVE` |
| --- | --- | --- | --- |
| Missing output-tag value | Existing fatal `E1301` behavior | Existing blank-tag behavior with `W1301` | Existing exact-tag behavior with `W1301` |
| Null receiver of a collection filter | Fatal filter mismatch, as today | Blank the complete output tag and warn | Preserve the complete exact output tag and warn |
| Missing repeat collection | Fatal `E1301`, as today | Blank the source rectangle in place and warn | Preserve the exact source rectangle and warn |
| Null repeat collection | Fatal `E1303`, as today | Blank the source rectangle in place and warn | Preserve the exact source rectangle and warn |
| Empty ordered repeat collection | Existing empty-placeholder behavior without a warning | Same | Same |
| Non-null non-collection repeat value | Fatal `E1303` | Fatal `E1303` | Fatal `E1303` |

### Blank repeat recovery

When a repeat collection is unavailable under `BLANK`, the repeat produces exactly one block with
the same width and height as its authored source rectangle.

- Every otherwise-supported cell value in the rectangle becomes blank (`None`), including static
  text, output tags, and structural directive text. Existing formula-support boundaries still
  apply before rendering.
- Direct cell presentation remains sourced from the corresponding authored cell. Fonts, fills,
  borders, alignment, number formats, and protection remain unchanged.
- Contained merges remain at their exact authored coordinates.
- Authored row and column presentation remains unchanged.
- The completed block has the same geometry as the source rectangle. It inserts, removes, collapses,
  or shifts no rows or cells, regardless of the repeat's authored shift mode.
- No descendant output expression or structural node is evaluated.
- Content outside the rectangle is unaffected.

For sparse rectangles, authored material cells are planned as blank cells and already-blank default
cells remain blank. The planner retains the complete rectangle's dimensions without manufacturing
presentation from neighboring cells.

### Preserve repeat recovery

When a repeat collection is unavailable under `PRESERVE`, the repeat also produces exactly one
source-sized block with identity geometry.

- Every authored cell value is retained exactly, including `{% for ... %}`, `{% endfor %}`, and
  `{{ ... }}` source text and spacing.
- Presentation, contained merges, and row/column provenance remain unchanged.
- No descendant output expression or structural node is evaluated.
- Nothing outside the rectangle is shifted.

The preserved region is an opaque diagnostic artifact. The engine does not promise that a workbook
containing a mixture of rendered content and preserved template regions can be rendered a second
time as a complete template.

### Nesting and parent placement

An unavailable outer repeat produces one blank or preserved source block and suppresses all
descendant evaluation, so it emits no secondary diagnostics from nested template expressions.

When an unavailable repeat is encountered inside a successful parent repeat, recovery occurs once
per parent instance. The recovered child block retains source-sized geometry and is translated only
by the already-valid parent placement. Each recovered occurrence emits its own warning at the child
repeat's opening tag.

### Collection-filter recovery

The collection-receiver rule applies consistently to collection-consuming filters:

- `join`
- `sum`
- `min`
- `max`
- `count`

If the receiver is missing, existing missing-output handling continues to apply. If the receiver is
canonical `null`, evaluation raises a distinct null-collection failure that the output boundary can
recover under `BLANK` or `PRESERVE`.

Recovery applies to the complete output tag, never to a partial filter pipeline. Existing empty
collection results remain unchanged: `sum` and `count` return `0`, `min` and `max` return null, and
`join` returns empty text.

The following remain fatal filter errors under every policy:

- a non-null receiver that is not an ordered collection;
- a non-record item in column mode;
- a missing column only when it is not otherwise handled by existing missing-output/default
  semantics;
- a nonnumeric selected value for numeric aggregates;
- incompatible numeric families or non-finite results.

Directly rendering canonical `null` remains an ordinary blank cell without a warning. Existing null
propagation for arithmetic and existing selected-null aggregate behavior do not change.

### Condition behavior does not change

A missing `if` condition remains fatal under every policy. A present-null condition retains its
existing falsey behavior. Recovering a missing condition by selecting, blanking, or preserving
conditional branches requires a separate design because a condition owns competing rectangles and
does not have one neutral branch result.

## Diagnostics

Strict-mode codes remain compatible:

- missing repeat collection: `E1301 MISSING_VALUE`;
- null or other non-collection repeat value: `E1303 EXPECTED_COLLECTION`;
- null or otherwise invalid collection-filter receiver: `E1304 FILTER_TYPE_MISMATCH`.

Tolerant recovery distinguishes missing from present null:

- `W1301 MISSING_VALUE_RENDERED` also covers a missing repeat collection recovered by `BLANK` or
  `PRESERVE`;
- add `W1302 NULL_COLLECTION_RENDERED` for a present-null repeat collection or collection-filter
  receiver recovered by `BLANK` or `PRESERVE`.

Messages for new collection recovery identify the action (`blanked` or `preserved`) and the affected
collection path; existing missing output-tag warning messages remain compatible. An output warning
uses the complete output token's source location. A repeat warning uses the opening directive's
source location. Repeated runtime occurrences produce repeated warnings.

Warnings do not publish a workbook if any independent error is also present. Rendering remains
atomic with respect to errors.

## Architecture

### Expression evaluation

Introduce a dedicated expression-evaluation failure for a canonical-null collection receiver. It
must be distinguishable from `FilterTypeError`, while strict rendering still maps it to `E1304`.
Collection filters validate their receiver before iterating it, preventing raw Python `TypeError`
from escaping for inputs such as `join(None)` or `join(1)`.

Do not catch generic `FilterTypeError` as recoverable. Wrong non-null values remain errors.

### Structural evaluation

Repeat control evaluation must return enough information for `ForNode` planning to distinguish:

- a valid ordered collection;
- a missing collection;
- a present-null collection;
- a present non-null value of the wrong type; and
- another expression failure.

The renderer applies the active policy only to the missing and present-null cases. It must decide
recovery before rendering descendants.

### Layout IR

Add pure planner helpers that produce either:

- an identity-sized block whose cell values are blank but whose presentation/row/merge provenance
  remains authored; or
- an identity-sized block whose cell values and provenance remain authored exactly.

Both helpers operate on the immutable template model and return ordinary layout IR. They do not
read or mutate live `openpyxl` objects. Parent placement consumes the blocks through the existing
layout interfaces.

### Workbook adapter

No layout decisions belong in the XLSX writer. The writer receives planned blank or preserved cell
values plus identity presentation coordinates and applies them normally. Workbook-level tests must
save and reopen the result to verify values, styles, dimensions, and merges.

The source workbook remains untouched and no output file is published after an independent error.

## Compatibility

- `ERROR` behavior and diagnostics remain unchanged.
- Existing missing output-tag behavior remains unchanged.
- `BLANK` and `PRESERVE` become more tolerant for missing/null repeat collections and null
  collection-filter receivers.
- Empty ordered collections retain their existing warning-free placeholder semantics.
- Wrong non-null values remain fatal.
- The public `missing_output` argument and `MissingOutputPolicy` enum remain source-compatible; their
  documented scope expands from output tags to recoverable collection-driven render output.

## Documentation and maintained sample

Update `SPEC.md` before implementation to define the new semantic boundary in the overview, empty
collection, missing-value, aggregate, evaluator, diagnostics, API, and testing sections. Update
`README.md`, `docs/directives.md`, and `docs/explained.md` consistently.

Extend `samples/missing_output_policies.py` and its matching template/output workbooks with:

- a multi-cell repeat whose collection is missing;
- a multi-cell repeat whose collection is null;
- a collection aggregate whose receiver is null;
- static text and styled/merged cells inside the repeat so blank-mode content clearing and
  presentation retention are visible;
- strict-mode non-publication, temporary blank-mode verification, and committed preserve-mode
  output verification.

Binary sample workbooks must be regenerated through the executable sample builder and verified by
save/reload assertions, never edited manually.

## Test strategy

Focused pure tests must cover:

- missing and null repeat collections under all three policies;
- exact rectangle size and zero displacement for blank/preserve recovery under `rows`, `cells`, and
  `none` shift modes;
- blanking of all cell content, including static text, while retaining presentation provenance and
  merges;
- character-for-character source text preservation for output/directive cells under `PRESERVE`;
- suppression of nested output and structural evaluation beneath a recovered repeat;
- recovery per parent-loop instance and warning cardinality/location;
- unchanged behavior for actual empty collections;
- fatal non-null collection mismatches;
- null collection receivers for every collection filter under all policies;
- fatal invalid collection members, numeric values, and independent render errors;
- no raw Python `TypeError` for invalid collection-filter receivers.

Workbook integration tests must build a minimal `.xlsx`, render under `BLANK` and `PRESERVE`, reopen
it, and assert cell values, effective styles, number formats, row heights, column widths, merges,
diagnostics, and unchanged neighboring content. Strict mode must publish no output workbook.

The maintained sample and full configured test, lint, format, and type-check gates remain required.

## Alternatives considered

### Normalize missing/null tables to empty collections before rendering

This is appropriate only when missing/null means known-empty data. It cannot preserve unresolved
template regions or emit renderer diagnostics for unavailable data, so it does not meet the full
recovery goal.

### Add a template filter or directive option

This makes optionality visible at every use site but clutters templates for nontechnical authors and
duplicates policy throughout a workbook. The selected design keeps templates unchanged.

### Add a separate repeat policy

This separates scalar and structural recovery precisely but adds another caller-facing option and
interaction matrix. The selected design uses the existing render-scoped policy with explicit
node-specific behavior and leaves conditions strict.

### Treat null as empty everywhere

This would erase the distinction between unavailable and known-empty data, suppress warnings, and
prevent `PRESERVE` from exposing unresolved template content. It is not adopted.
