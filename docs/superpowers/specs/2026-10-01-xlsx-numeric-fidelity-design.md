# XLSX Numeric Fidelity Design

**Date:** 2026-10-01

**Status:** Proposed for written review

## Intent

Prevent the XLSX writer from changing supported financial numbers during serialization. A finite
Python float or supported finite Decimal may pass through normalization, Polars, expression
evaluation, and layout unchanged today, but OpenPyXL 3.1.5 can shorten the worksheet XML token in a
way that changes the value when the workbook is reopened.

The confirmed serializer uses `%.16g` for numeric values. For example, the float
`100000.00000000001` is currently written as `100000` and reopens one binary64 step lower. The
failure is value-dependent, so an ordinary passing decimal does not establish fidelity.

This design adds an explicit numeric-output contract at the XLSX boundary. It preserves the core
language's existing numeric families and keeps workbook-specific conversion out of the evaluator
and layout planner.

## Decisions

- Support non-macro .xlsx workbooks only. This change does not add .xlsm support.
- Pin the direct dependency to OpenPyXL 3.1.5, the version whose private writer hook is certified.
- Preserve every finite built-in Python float in numeric XML, including integer-valued floats,
  both signs of zero, subnormals, and the finite binary64 extremes. Exact float type/bits after
  pinned OpenPyXL reopen require a format that it does not interpret as a date, time, or duration.
- Write a finite Decimal as an Excel numeric cell only when it is inside the documented 15-digit
  decimal domain and its adapter conversion does not change the represented decimal amount.
- Treat Decimal scale as presentation rather than stored-value identity. Values such as 12.5,
  12.50, and 1.25E+1 are the same amount; the authored cell number format controls displayed
  trailing zeros.
- Reject unsupported Decimal values before mutating or publishing a workbook. Never round them,
  stringify them, or silently fall back to a text cell.
- Use a guarded, process-lifetime OpenPyXL compatibility shim for float serialization.
- Require portable CI tests, including a frozen Excel-authored reference workbook. Live desktop
  Excel automation is optional and is not a default test or release gate.

## Goals

- Make no-op transfer of a finite float bit-preserving across template render, XLSX save, and
  OpenPyXL reopen with non-temporal number formats; preserve its numeric XML token for all formats.
- Give financial Decimal values a deterministic, documented acceptance rule through 15
  significant digits.
- Preserve numeric cell semantics so Excel formulas, sorting, filtering, and downstream readers
  continue to see numbers rather than strings.
- Detect all unsupported planned Decimal cells before creating or mutating the destination
  workbook.
- Report stable diagnostics at the exact destination worksheet and cell.
- Exercise direct Python, repeated-row, and Polars-produced values through the production render
  API.
- Compare a carefully bounded set of operation results with values calculated and saved by desktop
  Excel without requiring Excel in ordinary CI.

## Non-goals

- Byte-identical XLSX ZIP packages.
- Exact preservation after a person or desktop Excel opens, edits, recalculates, and resaves an
  arbitrary workbook.
- A new formula engine or fresh formula caches in generated workbooks.
- Generic equivalence between Polars and Excel arithmetic, reductions, reassociation, or rounding.
- Preserving Decimal object type, exponent, scale, or signed-zero identity after reopening an XLSX
  numeric cell.
- New fidelity guarantees for integers, dates, datetimes, times, timedeltas, booleans, formulas, or
  other existing cell types. They receive focused regression coverage because they share the
  writer.
- Enumerating every possible number with up to 15 digits. Coverage is systematic by precision and
  boundary, then supplemented with generated examples.
- A public write-only workbook path or .xlsm support.

## Public numeric contract

### Finite built-in floats

The contract applies when a planned numeric cell value has exact type float and is finite. The
XLSX writer emits a numeric worksheet cell whose decimal XML token preserves its binary64 value.
When the effective authored number format is not interpreted by the pinned OpenPyXL reader as a
date, time, or duration, reopening reconstructs:

- exact Python type float, including a value such as 1.0; and
- the same IEEE 754 binary64 bit pattern, measured by float.hex() or an equivalent raw-bit check.

The adapter uses repr(value) as the worksheet token. Modern Python chooses the shortest decimal
token that parses back to the same binary64 value. The adapter does not assign that token as the
cell value, because doing so would create a text cell.

The stored-token guarantee and non-temporal-format library round trip include +0.0, -0.0,
binary64 subnormal values, sys.float_info.min, and sys.float_info.max. These edge values are not a
promise about desktop Excel's behavior after open-and-save. Non-finite floats remain rejected by
the existing canonical value and expression diagnostics before XLSX preflight.

The stored-token guarantee covers every final planned cell with an exact built-in float, whether
the value was static in the source model, produced by a whole-cell expression, or repeated from
input data. Float subclasses keep their existing best-effort behavior and are outside the new
guarantee; the shim deliberately delegates them to OpenPyXL's original formatter rather than
guessing their semantics.

### Finite Decimal values

The core language continues to accept every finite Decimal as a canonical scalar. The narrower XLSX
policy applies only when a Decimal reaches a planned cell as a numeric value. An unused Decimal or
a value deliberately rendered as part of text is not reclassified by this policy.

A Decimal is supported as an XLSX numeric cell exactly when all of these conditions hold:

1. It is finite. Existing canonical validation handles this condition.
2. Its mathematical precision is at most 15 significant digits.
3. Converting it with float(value) produces a finite binary64 value.
4. The value is zero, or the converted magnitude is at least sys.float_info.min. Nonzero Decimal
   values that would become a binary64 subnormal or underflow to zero are outside the supported
   financial domain.
5. Decimal(repr(converted)) compares numerically equal to the original Decimal.

Mathematical precision is counted from Decimal.as_tuple().digits after removing coefficient
trailing zeros. Zero has precision one. This deliberately treats insignificant trailing zeros as
formatting: 12.50 has three significant digits, just like 12.5. The implementation must not use
Decimal.normalize() for this calculation because normalize() is affected by the active decimal
context.

Condition 5 means that the decimal token placed in worksheet XML represents the same decimal
amount. It does not claim that values such as 0.1 have an exact finite binary representation.
Accepted Decimal values are converted to built-in float at the XLSX boundary and then receive the
same round-trip-safe serialization as other floats. Reopening produces a float under the same
non-temporal-format condition; temporal formats receive the existing reader interpretation
described below. Decimal type, exponent, scale, and signed-zero identity are not preserved.

Validation order is deterministic: non-finite input is handled by existing validation, followed by
precision, range, and then decimal-token equality. A value that fails more than one predicate gets
the first diagnostic in that order.

### Presentation and other scalar types

Cell number_format remains authored presentation and is independent from stored numeric equality.
For example, 12.5 may display as 12.50 when the template uses format 0.00. Numeric tests compare the
stored value separately from number-format tests.

Authored number formats remain preserved for both floats and accepted Decimals. With a recognized
temporal format, pinned OpenPyXL may interpret the exact numeric token as a temporal Python value
(date, time, datetime, or timedelta), or return #VALUE! for an out-of-range serial. This is existing
reader interpretation, not token loss, and adds no temporal fidelity guarantee. The exact
reopen-as-float type/bit guarantee applies only to formats that the reader does not interpret as
dates, times, or durations. The same qualification applies to accepted Decimals.

Integers continue through OpenPyXL's original formatter and keep their current behavior; this
change adds no new large-integer guarantee. The compatibility wrapper also delegates booleans,
None, strings, formulas, Decimal objects that have not passed through preflight, and other
non-float objects to OpenPyXL's original formatter.

OpenPyXL converts some temporal cell values to float serials before calling the formatter. Those
serials will consequently use the corrected float token. Temporal regression tests use the pinned
OpenPyXL `from_excel(to_excel(value, epoch), epoch, timedelta=...)` result, including the 1900 and
1904 epochs and the ISO-date path. They do not promise identity with the original microseconds;
removing the lossy formatter can legitimately alter a half-millisecond tie by one millisecond.

## Architecture

### Pipeline placement

The production sequence remains:

1. normalize the complete context;
2. read the source workbook into immutable adapter-owned models;
3. compile, validate, evaluate, and plan every worksheet;
4. plan workbook features for every worksheet;
5. run adapter-local numeric preflight over every completed sheet plan;
6. ensure the OpenPyXL compatibility layer is active in memory and healthy;
7. create, populate, serialize, verify, and atomically publish the destination workbook.

No live destination workbook exists during numeric preflight. Any numeric diagnostic prevents the
writer from being entered, so an absent output stays absent and a pre-existing output stays
unchanged.

### Adapter-local numeric preflight

A new module inside excel_template_writer.xlsx owns the Decimal policy. It visits planned cells in
stable workbook, sheet, row, and column order and returns either:

- adapter-local copies of the sheet plans in which every supported numeric Decimal cell is replaced
  by its validated built-in float; or
- all independently discoverable numeric diagnostics, with no adapted plans.

The original layout IR remains immutable and retains its Decimal values. This proves that the core
normalizer, evaluator, and layout planner did not round financial values. The conversion is an XLSX
representation decision, not a language arithmetic decision.

The visitor checks every final occurrence. A Decimal repeated into three cells is validated at all
three destination coordinates, and a failure identifies each affected rendered cell. No logic is
added to the OpenPyXL writer's cell-application function to make first-error or partial-mutation
behavior possible.

The internal `write_workbook()` boundary invokes preflight and the compatibility guard before it
constructs an OpenPyXL `Workbook`. The public `render_workbook()` therefore cannot bypass them, and
an internal caller of the writer receives the same fail-closed behavior. Input loading occurs
earlier because it does not serialize cells; deferring process-global in-memory activation until
the write boundary avoids changing process state for a template that fails compilation or layout.

### OpenPyXL compatibility module

A named module, excel_template_writer.xlsx.openpyxl_numeric_compat, owns the only private OpenPyXL
integration. It targets openpyxl.cell._writer.safe_string because OpenPyXL 3.1.5 imports that
callable into the cell-writer module.

The owned wrapper behaves as follows:

```python
if type(value) is float and math.isfinite(value):
    return repr(value)
return original_safe_string(value)
```

The exact-type check is intentional. It prevents Decimal, NumPy scalars, booleans, or arbitrary
numeric-looking objects from receiving an invalid repr token.

In-memory activation neither installs a package nor edits OpenPyXL files on disk. Dependency
installation remains managed by uv. Activation verifies all of the following before normal
workbook serialization:

- the installed distribution version is exactly 3.1.5 and is in the local allowlist;
- the expected private module and hook exist;
- before first activation, the hook is the expected stock OpenPyXL formatter;
- an already-active wrapper carries this project's ownership/version marker and points to the
  captured original formatter; and
- an in-memory save-and-reopen self-test preserves the float.hex() value of
  100000.00000000001.

Unknown versions or private internals fail closed. There is no fallback to the known-lossy stock
formatter.

### Lifetime, ownership, and concurrency

In-memory activation is guarded by one module lock and a process-lifetime state machine:

```text
UNINSTALLED -> INSTALLING -> INSTALLED
                          -> FAILED
```

These internal state labels refer only to in-memory hook activation, not package installation.

Concurrent callers wait during INSTALLING. INSTALLED calls cheaply verify that the active hook is
still the owned wrapper. If another library replaces it, the state becomes sticky FAILED. A failed
self-test restores the captured original only when the hook still points to this project's wrapper,
then records the failure. Later calls return the cached diagnostic instead of retrying silently.

The wrapper is never activated and removed around an individual save. Toggling a process-global
hook per save would create a race. Separate render requests must still use separate OpenPyXL
Workbook instances and distinct output paths; this design does not make a shared Workbook
thread-safe.

Because the hook is process-global, unrelated OpenPyXL writers in the same process also receive the
safer finite-float representation. A host that requires simultaneous stock and corrected behavior
must isolate one writer in another process or adopt the OOXML-rewrite alternative described below.

### Version pin and upgrade rule

The direct dependency is pinned exactly to `openpyxl==3.1.5`; the compatibility allowlist contains
exactly `3.1.5`. The dependency pin, `uv.lock`, and allowlist must change together, with dependency
installation and lock updates managed by uv.

An OpenPyXL upgrade requires inspection of its numeric cell-writing path plus the full numeric
fidelity suite. The maintainer must then certify the new hook, adapt the shim, or remove the shim if
upstream serialization has become round-trip safe. A dependency update must never widen the range
without that review.

### Atomic output behavior

Decimal policy failures occur before writer mutation. Compatibility in-memory activation and its
self-test use an in-memory workbook before the destination writer is entered. Failures during the
later save continue to use the existing temporary-file, package-verification, reopen, and
atomic-replace path.

Thus every numeric or compatibility error has the same publication rule: no partial workbook is
returned, an absent destination is not created, and an existing destination is not replaced.

## Diagnostics

Add these stable codes in the existing E32xx XLSX-operation family:

| Code | Name | Location | Meaning |
| --- | --- | --- | --- |
| E3203 | XLSX_DECIMAL_PRECISION_EXCEEDED | Rendered Sheet!cell | Mathematical precision exceeds 15 digits. |
| E3204 | XLSX_DECIMAL_OUT_OF_RANGE | Rendered Sheet!cell | Conversion overflows, underflows, or would produce a nonzero binary64 subnormal. |
| E3205 | XLSX_DECIMAL_INEXACT | Rendered Sheet!cell | The round-trip-safe binary64 token is not numerically equal to the Decimal amount. |
| E3206 | XLSX_OPENPYXL_COMPAT_UNVERIFIED | <workbook>!A1 | Installed OpenPyXL version is not certified. |
| E3207 | XLSX_OPENPYXL_COMPAT_FAILED | <workbook>!A1 | The expected hook, ownership check, in-memory activation, or behavioral self-test failed. |

Finite Decimal policy errors are TemplateRenderError diagnostics because they concern planned
output values, not template syntax. Messages include the rejected value category without embedding
unstable object representations. Compatibility failures also prevent rendering and cache the
underlying reason for repeat calls and debugging.

Existing E1505 NON_FINITE_CONTEXT_NUMBER and E1307 NON_FINITE_EXPRESSION_NUMBER remain authoritative
for non-finite values. No duplicate XLSX diagnostic is added for an error already rejected by the
core.

E3207 intentionally combines private-hook shape, hook ownership/conflict, in-memory activation, and
self-test failures because they have the same caller action: stop output and repair or recertify
the pinned integration. Its stable message identifies the failed check. E3206 remains separate so
an uncertified dependency upgrade is immediately distinguishable.

## Polars and Excel comparison contract

Polars operation tests compute a result column in Polars before rendering and use a template such
as {{ row.result }}. Arithmetic inside the template would test the Python evaluator rather than
Polars and is not a substitute.

Each Polars test asserts four boundaries:

1. the Polars dtype and exact produced scalar;
2. the canonical adapter result before rendering;
3. the numeric worksheet XML token; and
4. the value and type after reopening the written workbook.

Float64 values use float.hex() or raw-bit equality. Decimal values use numeric Decimal equality
before the XLSX boundary, then the documented Decimal-to-float predicate and token equality after
that boundary.

Comparison with Excel is deliberately limited to operations for which both engines receive the
same binary inputs and explicit evaluation order. Representative exact cases include 0.5 + 0.25,
1.5 - 0.25, 1.125 * 2, and 5 / 2. The suite also preserves Polars-specific regression results such
as Float64 0.1 + 0.2, but it does not assert that every Excel build emits the same final binary64
bits for that operation.

The design makes no equivalence claim for reductions, fused or reassociated arithmetic,
near-zero correction, overflow, division by zero, or values beyond Excel's documented 15-digit
input domain.

## Test strategy

### Pure numeric-policy tests

Deterministic Decimal cases cover every precision p from 1 through 15, both signs, several scales
and exponents, and appended insignificant zeros. A corresponding 16-digit family proves the
rejection boundary. Required examples include:

- scale-equivalent 12.5, 12.50, 12.5000, and 1.25E+1;
- zero values with different signs and exponents;
- 123456789012345 and 999999999999999 as 15-digit values;
- 1234567890123456 and 1.234567890123456 as 16-digit values;
- ordinary positive and negative exponents;
- the selected lower normal and upper finite Decimal boundaries;
- nonzero subnormal, underflow-to-zero, and overflow cases; and
- 15-digit and 16-digit results produced by arithmetic rather than only literals.

Tests calculate expected classification from a small, independent expression of the five public
rules. They also verify diagnostic precedence and prove that the precision counter is independent
of the active decimal context.

Add Hypothesis in the development dependency group for two bounded property families:

- raw 64-bit patterns decoded as finite floats, with approximately 1,000-2,000 pure formatter
  examples; and
- Decimal tuples with 1-20 coefficient digits, exponents around the supported boundaries, signs,
  and appended zeros.

Deterministic examples remain the authority for boundaries. Property tests broaden the search
space; they do not replace the explicit matrix. XLSX I/O property tests use a smaller batch so the
suite remains practical.

### Compatibility-shim tests

Focused tests prove:

- the stock OpenPyXL 3.1.5 formatter loses the documented regression value;
- the wrapper emits a token that parses to the same binary64 value;
- non-float and non-finite inputs delegate to the captured original formatter;
- in-memory activation is idempotent and never wraps the wrapper again;
- version, missing-hook, foreign-hook, post-activation replacement, and self-test failures fail closed;
- FAILED is sticky and concurrent ensure calls observe one activation result;
- the successful in-memory self-test uses OpenPyXL primitives without recursively entering the
  production render API; and
- two independent workbooks can save concurrently after successful activation.

Direct compatibility tests exercise both OpenPyXL cell-writer implementations when available and
a raw write-only workbook after an explicit ensure call. That is hook coverage only; it does not
create a public write-only rendering API.

State-machine failure scenarios run in isolated subprocesses or equivalently isolated module
instances so a deliberately sticky failure cannot make the rest of the test process order
dependent.

### Production XLSX round-trip tests

Production tests construct a minimal template, call render_workbook(), inspect selected worksheet
XML with an XML parser, reopen the output, and assert semantics rather than entire ZIP bytes.

The fixed float matrix includes:

- 1.0 and -0.0 for type and signed-zero fidelity;
- 0.1234523461234556 as a value that already survived;
- 0.30000000000000004, 100000.00000000001, and 123456.78901234567 as formatter regressions;
- positive and negative cases;
- exponent-form values;
- nextafter neighbors around powers of ten and formatter boundaries;
- the minimum positive subnormal; and
- sys.float_info.min and sys.float_info.max.

For every float with a non-temporal format, the reopened value has exact type float and the same
float.hex(). Selected XML tokens must parse to the same bits and remain numeric cells regardless
of format. Float and accepted-Decimal temporal-format cases independently assert numeric tokens,
authored formats, and pinned-reader date/duration or out-of-range-error interpretation.

Accepted Decimal tests assert that:

- the pure render plan still contains the original Decimal;
- the output XML cell is numeric rather than inline text;
- Decimal(XML token) equals the source amount;
- with non-temporal formats, the reopened float has the expected float.hex(); and
- the authored number_format remains unchanged.

Rejected Decimal tests assert the exact code and rendered cell location. They cover direct cells,
repeated rows, several failures in one workbook, an absent destination, and a pre-existing
destination whose bytes remain unchanged.

Focused non-regression cases cover integers, booleans, blanks, formulas, styles, number formats,
the two workbook date epochs, ISO dates, and representative temporal half-millisecond boundaries.
Existing package, drawing, merged-cell, validation, and atomic-save suites remain part of the full
gate.

The production post-save reopen remains a package-integrity check rather than an exhaustive
per-cell numeric re-audit. Runtime confidence comes from complete Decimal preflight, the verified
serializer guard, and the focused/property tests; reopening every numeric cell and comparing it to
the plan would add a second full-workbook pass outside this change.

### Polars end-to-end tests

The optional Polars suite adds explicitly materialized result columns for Float64 and Decimal
arithmetic. It includes:

- exact shared-operation cases used by the Excel oracle;
- a Float64 result known to need the corrected serializer;
- a parameterized Decimal result for every precision from 1 through 15;
- signs and scale variants;
- a supported 15-digit boundary result; and
- a 16-digit or out-of-range result that fails at its rendered destination cell.

These tests first assert Polars' produced dtype and scalar. This prevents a changed Polars
arithmetic or casting behavior from being misdiagnosed as a writer regression.

### Frozen desktop-Excel reference

Commit one narrow, immutable .xlsx fixture that was created, fully recalculated, and saved by
desktop Excel. A neighboring manifest records:

- Excel version and build;
- locale and calculation settings;
- that Precision as displayed is disabled;
- fixture creation date and checksum;
- the authoritative cells, formulas, inputs, and comparison mode; and
- exact regeneration steps.

The fixture includes the p=1..15 Decimal-value matrix and the bounded shared operations. Portable
CI loads it with data_only=False to verify formula text and with data_only=True to read cached Excel
results. It also inspects selected raw formula and cached-value nodes so a missing cache cannot be
mistaken for a blank expected result.

Tests compare named cells and documented numeric semantics only. They never compare complete XLSX
bytes and never resave the reference through OpenPyXL. The cached results are a versioned snapshot
of the recorded Excel build, not a universal formula oracle.

Live Excel/COM testing may later recalculate a temporary copy on a controlled Windows machine and
report the Excel build. It is optional, serialized, and outside the default acceptance gate. This
change does not add pywin32 or make a licensed Excel installation a development dependency.

## Specification, documentation, and samples

SPEC.md is updated before implementation to define:

- the universal finite-float numeric-token contract and non-temporal-format reopen type/bit contract;
- signed zero, subnormal values, and finite extremes;
- the five-rule Decimal numeric-cell policy and loss of scale/type identity;
- number-format responsibility;
- adapter-local preflight and atomic failure;
- the new stable diagnostics; and
- the distinction between library round-trip fidelity and desktop Excel behavior.

The normative edits touch the canonical-value/cell-assignment sections, XLSX writer and output
verification sections, diagnostics, testing strategy, and compatibility/version policy. The XLSX
restriction remains cell-scoped and does not narrow the core finite-Decimal type.

README.md and the appropriate explanatory documentation receive a shorter user-facing numeric
output section.

Because this is a user-visible XLSX and adapter behavior change, maintained samples are updated and
regenerated through their executable builders:

- scalar_values gains a 15-digit financial Decimal with an authored financial number format; and
- polars_dataframe gains an explicitly computed numeric result column that exercises the same
  write-back boundary.

Their matching template and output .xlsx files and samples/README.md entries are regenerated and
verified through save/reload assertions and the sample-manifest tests. The existing
samples/generate_all.py entry point remains able to rebuild the complete maintained set. Binary
workbooks are never edited as ordinary source.

## Alternatives considered

### Targeted worksheet-XML rewriting

The writer could record every planned float cell and rewrite only those numeric value nodes inside
the completed package. This avoids process-global mutation but adds ZIP relationship, worksheet
mapping, and XML-rewrite complexity to every save. It is the preferred fallback if the global hook
becomes unacceptable, but not the initial implementation.

### Maintained OpenPyXL fork

A fork would isolate the source change and avoid a private runtime hook. It also requires publishing
and securing a custom distribution, tracking upstream fixes, and repeatedly rebasing a one-line
behavioral change. That maintenance burden is not justified for the controlled pinned runtime.

### Write Decimal as text

Text preserves the authored digits and scale but breaks numeric formulas, sorting, filtering, and
downstream numeric readers. It is not an implicit fallback. A template author may still deliberately
produce text through ordinary mixed-text rendering.

### Round unsupported Decimal values

Implicit rounding would hide the exact class of financial discrepancy this feature is meant to
prevent. Callers must round or quantize under an explicit business policy before rendering.

## Acceptance criteria

The implementation is complete when:

- SPEC.md contains the approved contract before runtime behavior changes;
- OpenPyXL and the lock file are pinned to 3.1.5;
- every fixed and generated supported float preserves its binary64 value in numeric XML and
  round-trips with the same float type/bits when its format is not interpreted as temporal;
- supported Decimal values through 15 significant digits remain numeric and represent the same
  decimal amount in worksheet XML;
- unsupported Decimal values fail with stable, precisely located diagnostics and no publication;
- the core render plan retains Decimal values until the XLSX preflight boundary;
- direct Python, repeats, and Polars arithmetic paths are covered;
- the frozen Excel-authored fixture passes in portable CI;
- samples and user documentation are regenerated consistently;
- the compatibility layer is version-gated, owned, idempotent, self-tested, concurrency-tested,
  and fail-closed; and
- focused tests plus the full pytest, Ruff check, Ruff format, and ty gates pass.
