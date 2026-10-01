# Numeric Fidelity Handoff for Excel Template Writer

Status: serializer design selected; numeric-domain contracts remain implementation prerequisites

Prepared: 2026-09-30

Decision updated: 2026-10-01

Source project: `excel-data-reader`
Target project: `excel-template-writer`

## Executive summary

A numeric value can pass through the reader, Python, Polars, and the template writer unchanged,
then change when OpenPyXL serializes the destination workbook.

The confirmed cause in the currently locked OpenPyXL 3.1.5 is its numeric formatter:

```python
value = "%.16g" % value
```

Some IEEE 754 binary64 values require 17 significant decimal digits to round-trip to the same
binary value. Formatting every number with 16 significant digits can therefore change the value
stored in the output workbook.

This is a confirmed writer-boundary precision risk. Whether it is a defect under the current
public contract is a specification decision: the current contract preserves numeric types but does
not yet promise bit-identical binary64 output. Automatically converting every non-integer value to
`decimal.Decimal` is not an effective fix: OpenPyXL sends `Decimal` through the same 16-digit
formatter, and arbitrary decimal values are not necessarily exactly representable as binary64
Excel numeric cells. Many finite in-range `Decimal` values can be written approximately, but that
requires an explicit conversion and rounding policy.

This mechanism can explain discrepancies of the reported magnitude, but a representative failing
production value should still be traced through the same boundaries before attributing every
observed discrepancy to it.

## Decision

Implement a small, locally maintained OpenPyXL compatibility shim inside the template writer. For
the verified OpenPyXL version, the shim will replace the cell writer's imported formatter so that
finite built-in Python `float` values use `repr(value)` when written to worksheet XML. Every other
value that reaches the hook without first becoming a built-in float will continue through
OpenPyXL's original formatter.

This is preferable here to maintaining an OpenPyXL fork because the project controls and locks its
runtime, upgrades OpenPyXL infrequently, and needs only one narrowly scoped behavior change. The
shim must nevertheless be treated as a dependency compatibility layer, not as an unrestricted or
temporary monkeypatch:

- keep it in one named module in the XLSX adapter;
- verify the exact OpenPyXL version and private hook before installing it;
- have every XLSX entry point ensure it is installed before any cells can serialize;
- make installation idempotent and run a behavioral self-test;
- fail closed on an unknown version or failed self-test;
- do not patch and restore around each workbook save.

The initial verified version is OpenPyXL 3.1.5. As of 2026-10-01, it is also the latest release
listed on [PyPI](https://pypi.org/project/openpyxl/), so there is no newer published release to
adopt as an upstream fix at the time of this decision.

## Numeric model behind the decision

Three representations are involved in a normal XLSX round-trip:

1. Excel, Python `float`, Polars `Float64`, and DuckDB `DOUBLE` calculate or hold a binary
   floating-point value.
2. An `.xlsx` worksheet serializes that value as decimal characters inside an XML `<v>` element.
3. Excel's cell number format controls the text shown in the UI.

The XML characters are a transport encoding; they do not turn the value into exact decimal
arithmetic. A sufficiently informative decimal token can identify one binary64 value uniquely, so
a correctly rounded binary64 parser such as Python/OpenPyXL can reconstruct the same bits. Python's
`repr(float_value)` produces the shortest token with that round-trip property. OpenPyXL 3.1.5's
`%.16g` sometimes removes one digit too many, which can make its parser select an adjacent binary64
value instead. Desktop Excel remains outside this parser guarantee.

This correction preserves the binary value already held by the writer. It does not make values
such as `0.1` exact in base 2, change arithmetic to decimal arithmetic, or determine what Excel
displays. It also does not override Microsoft Excel's documented 15-significant-digit behavior.
Open-and-resave testing in desktop Excel remains a separate compatibility test from an OpenPyXL
save-and-reopen test. See Microsoft's documentation on
[floating-point arithmetic in Excel](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/floating-point-arithmetic-inaccurate-result)
and [Excel specifications and limits](https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3).

## User-visible consequence

A fetch-and-copy operation can produce a destination value that differs from the source even when
the application performs no arithmetic.

Confirmed examples from an OpenPyXL save-and-reopen test:

| Input Python value | Number written in worksheet XML | Reopened value | Difference |
| --- | --- | --- | --- |
| `0.1234523461234556` | `0.1234523461234556` | `0.1234523461234556` | none |
| `0.30000000000000004` | `0.3` | `0.3` | `-5.551115123125783e-17` |
| `100000.00000000001` | `100000` | `100000` | `-1.4551915228366852e-11` |
| `123456.78901234567` | `123456.7890123457` | `123456.7890123457` | `+2.9103830456733704e-11` |

The value `0.1234523461234556` happens to survive. The problem is value-dependent, so a passing
test containing one ordinary decimal does not establish numeric fidelity.

## Confirmed pipeline behavior

The reproduced value was checked at each boundary:

| Boundary | Result |
| --- | --- |
| Excel Data Reader/OpenPyXL load | OpenPyXL parses the XML token; the reader returns that Python numeric value without further numeric transformation |
| Reader extraction model | The cell value is retained unchanged |
| Polars `Float64` materialization | Same Python float and identical `float.hex()` value |
| Template writer normalization | Finite `float` retained unchanged |
| Assignment to destination cell | Same value assigned to `destination.value` |
| `Workbook.save()` | OpenPyXL formats the value with `%.16g`; affected values change |

Relevant code locations are:

- Reader extraction: `excel-data-reader/src/excel_data_reader/reader.py`, `_cell_data`
- Polars conversion: `excel-template-writer/src/excel_template_writer/adapters/polars.py`,
  `_convert_dataframe`
- Scalar normalization: `excel-template-writer/src/excel_template_writer/values.py`, scalar
  normalization logic
- Cell assignment: `excel-template-writer/src/excel_template_writer/xlsx/writer.py`, `_apply_cell`
- Workbook save: `excel-template-writer/src/excel_template_writer/xlsx/writer.py`, `_write_sheet`
  and the final `workbook.save(temporary_path)` call
- OpenPyXL formatting: `openpyxl/compat/strings.py`, `safe_string`
- OpenPyXL numeric XML write: `openpyxl/cell/_writer.py`

The template writer currently declares `openpyxl>=3.1.5,<4`, and its lock file resolves 3.1.5. The
selected private compatibility hook requires changing the direct dependency to `openpyxl==3.1.5`
until another version is explicitly certified.

## Minimal reproduction

Run this in the template writer environment:

```python
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook, load_workbook


source = 100000.00000000001

with TemporaryDirectory() as temporary_directory:
    output_path = Path(temporary_directory) / "roundtrip.xlsx"

    workbook = Workbook()
    workbook.active["A1"] = source
    workbook.save(output_path)

    reopened = load_workbook(output_path, data_only=True)
    result = reopened.active["A1"].value

    print("source:", repr(source), source.hex())
    print("result:", repr(result))
    print("difference:", result - source)
```

Expected result with OpenPyXL 3.1.5:

```text
source: 100000.00000000001 0x1.86a0000000001p+16
result: 100000
difference: -1.4551915228366852e-11
```

For float-to-float comparisons, compare `float.hex()` as well as the printed decimal text. Decimal
display text alone cannot prove that the underlying binary values are identical.

## Proposed public contract

Update `SPEC.md` before changing behavior, as required by the template writer repository policy.
The contract should distinguish passthrough numeric fidelity from decimal-domain semantics.

Define two layers separately. The serializer guarantee applies to every finite built-in `float`
that the canonical value model accepts: the XLSX adapter must emit enough information for the
supported OpenPyXL reader to reconstruct the same binary64 value. The accepted Excel-value domain
remains a public policy enforced before serialization. The `SPEC.md` change must explicitly decide
whether subnormal values and IEEE 754 extremes are accepted, rejected, or documented as
library-round-trip-only values, with stable diagnostics for rejections. The shim must not silently
broaden that domain merely because `repr` can serialize a value.

Suggested contract for supported Python floats:

> A `float` within the documented supported range, written by a whole-cell expression, is emitted
> as an Excel numeric cell. Saving and reopening the generated workbook through the supported XLSX
> adapter must recover the same IEEE 754 binary64 value and Python `float` type. This includes an
> integer-valued float such as `1.0`. Signed zero is preserved unless the specification explicitly
> chooses and tests normalization. The writer must not introduce an undocumented numeric rounding
> step.

This is a library save-and-reopen contract. It does not promise that a value will remain unchanged
after a person edits it or Microsoft Excel recalculates and saves it. If exact behavior after an
Excel application round-trip is a requirement, test that workflow separately.

The selected `repr` behavior intentionally writes `1.0` as `1.0` rather than stock OpenPyXL's `1`.
OpenPyXL therefore reopens it as a `float` instead of an `int`. That type fidelity is part of the
proposed contract and must be recorded in `SPEC.md` because it is publicly observable.

These type and signed-zero guarantees are specific to the template writer's OpenPyXL adapter.
SpreadsheetML numeric cells and desktop Excel do not generally promise to preserve a Python
`int`/`float` distinction or the sign of zero.

Suggested contract for integers:

> Integers within the explicitly supported Excel numeric range remain numeric and equal after a
> save-and-reopen cycle. The specification must define what happens outside that range.

Suggested contract for `Decimal`:

> `Decimal` handling is explicit rather than inferred from the presence of a fractional part. The
> writer either converts it to an Excel binary64 number under a documented rounding/range policy,
> rejects a value that cannot satisfy the selected exactness or range policy with a stable
> diagnostic, or writes text only when the caller explicitly requests textual preservation.

Writing a `Decimal` as text preserves its digits but changes its Excel type and affects formulas,
sorting, filtering, and downstream readers. It must not be a silent fallback.

### Contract decisions still required

The serializer mechanism is selected, but implementation must not begin by guessing the remaining
business semantics. The owner of the template writer's public contract must resolve these in
`SPEC.md` first:

- the accepted finite-float range, including whether subnormal values and IEEE 754 extremes are
  supported for library round-trips or rejected for desktop-Excel compatibility;
- the exact integer range and behavior outside it;
- one concrete `Decimal` policy: documented binary64 conversion and rounding, rejection, or
  explicitly requested text; and
- stable diagnostics for every rejected category, for example `XLSX_FLOAT_OUT_OF_RANGE`,
  `XLSX_INTEGER_OUT_OF_RANGE`, or `XLSX_DECIMAL_INEXACT` where those meanings are adopted.

This handoff recommends preserving both signs of zero and the Python `float` type through the
OpenPyXL adapter. It does not extend that guarantee to a desktop Excel open-and-resave cycle.

## Selected implementation direction

Keep the correction inside the XLSX adapter boundary. Do not change the reader, core value model,
Polars adapter, or DuckDB mapping merely to compensate for OpenPyXL serialization. In particular,
do not convert every non-integer float to `Decimal`.

### Exact compatibility hook

In OpenPyXL 3.1.5, `openpyxl.cell._writer` imports `safe_string` into its own module namespace:

```python
from openpyxl.compat import safe_string
```

Both the etree and lxml cell-writing paths then call that module-global name when producing a
numeric `<v>` element. Replacing only `openpyxl.compat.safe_string` after this import does not
change the already-bound reference used by the cell writer. The verified hook is therefore:

```python
openpyxl.cell._writer.safe_string
```

The compatibility module should capture the original function and use logic equivalent to:

```python
def roundtrip_safe_string(value: object) -> str:
    if type(value) is float and math.isfinite(value):
        return repr(value)
    return original_safe_string(value)
```

Use `type(value) is float`, rather than applying `repr` to every numeric-looking object. For
example, `repr(Decimal("19.75"))` is `Decimal('19.75')`, which is not a valid worksheet numeric
token. NumPy scalar representations also require their existing normalization policy. Integers,
booleans, `Decimal`, NumPy values, `None`, strings, formulas, and non-finite values must continue
through existing behavior unless their own documented contract is deliberately changed.

OpenPyXL converts dates, times, and timedeltas to built-in float serials before this hook unless an
ISO representation is selected. Those serials therefore also use `repr`; the shim cannot identify
their original cell-level type at this point. OpenPyXL 3.1.5 reloads numeric temporal serials at
millisecond resolution, so values on a half-millisecond boundary can reopen one millisecond
differently after the lossy `%.16g` step is removed. For numeric temporal cells, define the expected
result as the pinned adapter's canonical
`from_excel(to_excel(source, epoch), epoch, timedelta=isinstance(source, timedelta))` result, not
as the stock 16-digit XML result and not necessarily as the source's original microseconds. Test
dates/datetimes, times, timedeltas, and the ISO-date path separately. This accepts corrected 1-ms
tie differences while preventing larger or unrelated temporal regressions.

`repr(float_value)` is preferred to unconditional `.17g`: modern Python returns the shortest
decimal token that parses back to the same binary64 value. Seventeen significant digits are
sufficient for round-tripping binary64, but always emitting all 17 can add unnecessary digits.

Do not assign `repr(value)` to `cell.value`. That would create a text cell and change formula,
sorting, filtering, and downstream-reader behavior. The override applies only while serializing an
already numeric cell to its XML `<v>` token.

### Ownership and installation

Place the shim in one explicitly named module, for example:

```text
src/excel_template_writer/xlsx/openpyxl_numeric_compat.py
```

Do not rely solely on application startup or import side effects: the template writer may be used
as a library. Every public XLSX render/write entry point should call
`ensure_openpyxl_numeric_compat()` before creating or loading a workbook and before any cell can be
serialized. This is especially important for write-only workbooks because `append()` can serialize
a row before `Workbook.save()`.

The ensure function should use one module lock and a process-lifetime state machine:

```text
UNINSTALLED -> INSTALLING -> INSTALLED
                          -> FAILED
```

Concurrent callers wait while the state is `INSTALLING`. In `INSTALLED`, each guarded entry cheaply
checks that `openpyxl.cell._writer.safe_string` is still the owned wrapper before returning. If a
third party replaced it, transition to sticky `FAILED` and raise the compatibility diagnostic.
`FAILED` re-raises the cached diagnostic; later calls must not silently retry. The self-test must
use OpenPyXL primitives directly rather than re-entering the template writer's guarded entry point.
Do not install and restore the shim around individual saves: another thread could otherwise
serialize cells while the process temporarily uses the wrong formatter.

Once installation reaches the immutable `INSTALLED` state, the shim introduces no
formatter-toggle race between two or more concurrent template writers when all of the following
are true:

- every writer in the process uses the same fidelity contract;
- every writer passes the guard before any workbook content can serialize;
- the wrapper is pure and stateless; and
- each request owns a separate OpenPyXL `Workbook` instance.

OpenPyXL workbook objects themselves should not be shared across threads. Multiple processes are
independent; each process installs and verifies its own shim. This statement concerns the shim's
state only; overall concurrent-writing support still depends on the template writer's tested
execution model, independent output paths, and OpenPyXL usage. This design is not appropriate for
a host process that simultaneously requires both patched and stock OpenPyXL serialization, or for
an unrelated OpenPyXL consumer that cannot accept the changed float-cell behavior. The entry guard
can detect replacement before a save begins, but it cannot prevent uncoordinated third-party
mutation during a save already in progress.

The hook is module-global within `openpyxl.cell._writer`, so it affects all float cell
serialization in that process, including float serials produced from dates, times, and timedeltas.
It does not replace every caller of `openpyxl.compat.safe_string`. Date epochs, `iso_dates`, normal
and write-only workbooks, and both etree and lxml writer paths therefore need explicit regression
coverage.

### Version and installation guard

The shim depends on a private OpenPyXL implementation detail and is not automatically safe merely
because it lives in local code. Make that dependency explicit:

1. Maintain an allowlist of versions verified by the template writer, initially `{ "3.1.5" }`, and
   pin the direct dependency to that version.
2. For the stock 3.1.5 precondition, verify that `openpyxl.cell._writer.safe_string is
   openpyxl.compat.safe_string`. Reject a different unowned callable rather than wrapping it.
3. Give the installed wrapper an owner/version marker and a reference to the captured original so
   a second installation recognizes the same shim instead of wrapping again. Reject an unknown
   wrapper, including one that merely imitates the marker incorrectly.
4. Run an in-memory behavioral self-test using `io.BytesIO` and `100000.00000000001`: save a
   workbook, inspect or reopen it, and confirm that the recovered float has the same `float.hex()`
   representation.
5. If installation or the self-test fails, restore the captured original once if the hook still
   points to this shim, record the sticky `FAILED` state, and stop the XLSX entry point with a
   stable diagnostic such as `XLSX_OPENPYXL_COMPAT_SELF_TEST_FAILED`.
6. Reject an unverified installed version with a separate stable diagnostic such as
   `XLSX_OPENPYXL_COMPAT_UNVERIFIED`; do not silently fall back to known-lossy output.

Keep both the direct dependency and lock file pinned to 3.1.5 until another version has been
inspected and tested. For each upgrade, inspect the cell-writer path, run the full fidelity suite,
and then either change the pin and allowlist together, adapt the shim, or remove it if upstream
serialization is round-trip safe.

### Alternatives and exit strategy

A maintained OpenPyXL fork would isolate the source change from other in-process consumers, but it
would also require publishing a custom distribution, pinning its provenance, tracking security
updates, and rebasing the patch. That overhead is not justified for the currently controlled,
stable runtime.

If global mutation later becomes unacceptable, use targeted OOXML post-processing as the fallback:
rewrite only the numeric `<v>` elements for known planned cells, using an XML parser and explicit
sheet/cell mappings. Do not use unrestricted string replacement inside the XLSX ZIP archive. This
is more isolated but materially more complex because it must preserve relationships, cell types,
macros, and every supported workbook feature.

## Required tests

Follow the template writer's XLSX integration workflow: create the smallest workbook fixture,
render it, save it, reopen it, and inspect the OOXML when necessary.

### Contract tests

Add contract-level cases for ordinary supported values such as:

```python
[
    1.0,                        # type fidelity; must reopen as float rather than int
    -0.0,                       # signed-zero policy
    0.1234523461234556,       # already survives; prevents over-correction
    0.30000000000000004,      # requires round-trip-safe output
    100000.00000000001,       # reproduced approximately 1e-11 drift
    123456.78901234567,       # reproduced approximately 2.91e-11 drift
    -100000.00000000001,
    1e20,                      # exponent-form output
]
```

As part of the `SPEC.md` update, add explicit boundary tests for the selected minimum and maximum,
both signs of zero, and subnormal values. Stock OpenPyXL 3.1.5 can serialize the maximum finite
binary64 value to an overlarge token that reopens as infinity; the `repr` shim avoids that specific
loss in an OpenPyXL round-trip. Desktop Excel's accepted range and resave behavior remain separate
constraints.

Use category and property tests rather than attempting to enumerate every binary64 value. Generate
random finite values from raw 64-bit patterns inside the supported domain and assert that parsing
the emitted token recovers the same bits. Add `math.nextafter()` neighbors around powers of ten and
other formatter boundaries, plus very small and very large supported values. Define and test an
explicit policy for values Excel does not support, including subnormals and out-of-range finite
values.

For each supported float:

1. Render it through a whole-cell template expression.
2. Save the workbook.
3. Reopen it with the supported reader.
4. Assert `type(reopened_value) is float` for a source `float`, including `1.0` and `-0.0`.
5. Assert `reopened_value.hex() == source_value.hex()`.
6. Inspect the worksheet XML for at least one regression value and assert its `<v>` token parses
   back to the original float.

### Integration coverage

Add coverage at these entry points:

- direct Python scalar context;
- Polars `Float64` context;
- a repeated-row/table render, not only one scalar cell;
- a value copied from an Excel Data Reader result, if the projects have an integration test
  boundary;
- both a newly created sheet and a rendered template sheet if they use different write paths;
- normal and write-only OpenPyXL workbook paths;
- both lxml and etree cell-writer paths where the supported environment can exercise them;
- separate workbooks saved concurrently after one successful ensure call;
- `.xlsx` and macro-preserving `.xlsm` output, if both are supported.

### Compatibility-shim tests

Add focused tests that prove:

- stock OpenPyXL 3.1.5 loses at least one documented regression value;
- the shim emits `100000.00000000001` as a numeric XML token and reloads the same binary64 value;
- installation is idempotent and does not wrap its own wrapper;
- a post-install replacement of the owned hook is detected on the next guarded entry and makes the
  failure sticky;
- an unverified OpenPyXL version fails closed with the documented diagnostic;
- a missing or changed private hook fails the guarded installation/self-test;
- the original formatter is used for every value outside the narrow finite built-in `float` case;
- the state machine blocks concurrent callers during installation and makes failure sticky;
- write-only workbooks install the shim before the first `append()`; and
- no per-save patch/unpatch race is present.

### Non-regression coverage

Confirm that the change preserves semantic behavior, while allowing the XML token for a finite
float or temporal serial to become more informative:

- integer, boolean, `None`, and formula behavior;
- numeric dates, times, and timedeltas match `from_excel(to_excel(source, epoch), epoch,
  timedelta=isinstance(source, timedelta))`, including half-millisecond boundaries;
- the 1900 and 1904 date epochs and `iso_dates` behavior;
- `Decimal` and normalized NumPy scalar behavior;
- cell styles and number formats;
- workbook relationships, macros, charts, drawings, merged cells, tables, and validations;
- atomic-save and reopen validation behavior;
- the existing Decimal behavior except where the revised specification deliberately changes it.

Run the complete commands required by the target repository before handoff.

## Acceptance criteria

The change is complete when all of the following are true:

- `SPEC.md` explicitly defines the accepted float domain, float type fidelity, signed zero,
  subnormals, integer boundaries, and Decimal output behavior.
- The direct OpenPyXL dependency and lock file are pinned to the same verified version.
- Each supported regression value survives render, save, and reopen with the expected binary64
  value.
- At least one test inspects the actual worksheet XML numeric token.
- The Polars path is covered and does not receive an unnecessary Decimal conversion.
- Display formatting is tested separately from numeric equality.
- Unsupported Decimal or integer cases produce documented behavior rather than silent loss.
- The compatibility shim is isolated to the XLSX adapter, version-gated, idempotent, installed once
  through guarded XLSX entry points, and covered by its behavioral self-test.
- Unknown OpenPyXL versions or changed private internals fail closed with stable diagnostics.
- Concurrent saves use independent workbook instances and never patch or restore serializer state
  per save.
- Write-only paths run the compatibility guard before the first row can serialize.
- Numeric temporal cells match the pinned adapter's canonical `to_excel`/`from_excel` round-trip,
  including half-millisecond boundary cases.
- All target-repository tests, lint checks, formatting checks, and type checks pass.

## Important distinctions

### Stored value versus displayed value

Excel's number format controls what a user sees. Two cells can contain identical binary values but
display different text, or display identical text while containing different binary values.
Numeric fidelity tests must compare the reopened value, while presentation tests compare the cell's
`number_format` and rendered appearance separately.

### Formula text versus calculated value

OpenPyXL does not calculate formulas. Copying formula text is different from copying a cached
calculation result, and moving formula text to another coordinate is different from Excel's
reference-aware copy/paste operation. This issue should not be mixed into the numeric serialization
fix.

### Binary64 versus exact decimal arithmetic

Excel numeric cells, Python `float`, Polars `Float64`, and DuckDB `DOUBLE` are compatible binary64
domains for ordinary transfer. `Decimal` represents a different semantic choice. Use it when the
business domain requires an explicit decimal scale and rounding policy, not merely because a value
contains a decimal point.

## Evidence and primary references

- [Python's floating-point tutorial](https://docs.python.org/3/tutorial/floatingpoint.html)
  documents binary64 approximation and the shortest `repr` representation that preserves the
  `eval(repr(x)) == x` round-trip invariant.
- [OpenPyXL 3.1.5 on PyPI](https://pypi.org/project/openpyxl/3.1.5/) identifies the exact release.
  Its immutable
  [source archive](https://files.pythonhosted.org/packages/3d/f9/88d94a75de065ea32619465d2f77b29a0469500e99012523b91cc4141cd1/openpyxl-3.1.5.tar.gz)
  has SHA-256 `cf0e3cf56142039133628b5acffe8ef0c12bc902d2aadd3e0fe5878dc08d1050`.
  Inspect `openpyxl/compat/strings.py` for `safe_string` and
  `openpyxl/cell/_writer.py` for the imported hook and both writer paths.
- Microsoft's [CellValue (`<v>`) reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.cellvalue)
  documents that a non-string cell value is expressed directly in the SpreadsheetML `<v>`
  element.
- Microsoft's
  [floating-point arithmetic guidance](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/floating-point-arithmetic-inaccurate-result)
  and [Excel limits](https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3)
  document Excel's binary arithmetic constraints and 15-significant-digit limit.
- The source project's companion [numeric-fidelity note](../../NUMERIC_FIDELITY.md) provides a
  slower, conceptual explanation for future maintenance.

## Suggested issue title

`Preserve finite float values across XLSX save and reopen`

Suggested issue summary:

> OpenPyXL 3.1.5 serializes numeric cells with 16 significant digits. Some finite binary64 values
> require 17 digits to round-trip, allowing a no-op template copy to change the stored number. Add
> an explicit numeric-fidelity contract and a guarded, version-checked compatibility shim for the
> OpenPyXL cell writer, with installation, concurrency, save/reopen, and OOXML regression tests.
