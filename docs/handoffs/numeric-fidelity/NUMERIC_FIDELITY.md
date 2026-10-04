# Numeric Fidelity Notes

This is a personal reference for reviewing Excel numeric pipelines. It is intentionally separate
from `AGENTS.md`, so it is not a standing instruction that must be considered on every task.

For the full template-writer engineering handoff, see
[`handoff/excel-template-writer/README.md`](handoff/excel-template-writer/README.md).

## Short version

- An Excel number has three relevant representations: binary floating point while Excel calculates,
  decimal characters while an `.xlsx` file serializes the number in XML, and formatted text in the
  Excel UI.
- The decimal characters in XML are a serialization format. They do not mean Excel is using exact
  decimal arithmetic internally.
- Python `float`, Polars `Float64`, and DuckDB `DOUBLE` are the natural passthrough types for an
  ordinary Excel number.
- A decimal point does not automatically mean the value should become `Decimal`.
- `Decimal` is appropriate when the business contract requires fixed decimal scale or explicit
  decimal rounding, such as currency accounting rules.
- OpenPyXL 3.1.5 can change certain floats while saving because it writes numbers with only 16
  significant digits. Some binary64 values require 17 digits to round-trip.
- Therefore, a simple read-and-copy operation can drift even when no arithmetic is performed.

## The mental model: one number through Excel

Consider entering `0.1` into a normal numeric cell.

### 1. Input

You initially enter the decimal characters:

```text
0.1
```

At this moment they are user input. Excel must decide whether they represent a number, text, a
date, or another supported cell value.

### 2. Internal numeric value

For a normal numeric cell, Excel converts the input into its supported binary floating-point
representation. Excel uses double-precision floating-point arithmetic and exposes 15 significant
decimal digits of numeric precision.

Some decimal values have an exact finite binary representation. Others do not.

| Decimal value | Exact as a binary floating-point value? | Reason |
| --- | --- | --- |
| `0.5` | Yes | `1/2` is a power-of-two fraction |
| `0.25` | Yes | `1/4` is a power-of-two fraction |
| `1.125` | Yes | `1 + 1/8` |
| `0.1` | No | Its binary fraction repeats forever |
| `0.01` | No | Its binary fraction repeats forever |
| `19.99` | No | It cannot be expressed as a finite binary fraction |

For an inexact value such as `0.1`, Excel stores the nearest supported binary value. Conceptually,
that value is approximately:

```text
0.10000000000000000555...
```

The approximation is already present before a formula uses the value.

### 3. Calculation

Excel normally calculates with the stored values, not the decimal text currently visible in the
cells. Each arithmetic operation produces a mathematical result and then fits that result into the
next supported floating-point value.

For repeated multiplication, the process is conceptually:

```text
x0 = stored approximation of a
x1 = round_to_float(x0 * stored approximation of b)
x2 = round_to_float(x1 * stored approximation of b)
x3 = round_to_float(x2 * stored approximation of b)
...
```

Therefore, the thought that repeated operations can eventually differ from ideal real-number or
exact-decimal arithmetic is correct. There are two opportunities for approximation:

1. `a` and `b` may already be approximations.
2. Every multiplication result may require another floating-point rounding step.

The error does not necessarily increase on every operation. It can grow, shrink, cancel out, or
remain too small to affect the displayed digits. But exact agreement with unlimited-precision
decimal arithmetic is not guaranteed.

Recalculating the same deterministic formula with unchanged inputs normally returns the same
result. Error accumulates when a rounded result feeds a later operation, not merely because Excel
recalculates an otherwise unchanged formula.

In an Excel support article, Microsoft demonstrates the same propagation using a VBA `Double`:
repeatedly adding `0.0001` 10,000 times can produce `0.999999999999996` instead of exactly `1`.
Excel also applies a limited correction for some addition or subtraction results at or very close
to zero, so not every small floating-point artifact is exposed in the same way.

### 4. Display in the Excel UI

The cell's number format converts the stored number into the text shown to the user. Excel may hold
an approximation of `0.1` while showing only:

```text
0.1
```

Changing the number format normally changes only the display. By default, formulas continue using
the stored value.

Excel's optional **Set precision as displayed** setting is different: it permanently changes stored
values to match the displayed precision. Microsoft warns that the original precision cannot be
recovered after enabling it and saving the workbook.

### 5. Saving an `.xlsx` file

An `.xlsx` file is a ZIP archive containing XML. XML is text, so the current numeric value must be
serialized as decimal characters, for example:

```xml
<c r="A1" t="n"><v>0.1</v></c>
```

This `<v>` content is not a `Decimal` object and does not cause decimal arithmetic. It is text used
to transport the number through the file format.

### 6. Reopening the file

Excel or OpenPyXL parses the XML characters and reconstructs an in-memory numeric value. An
IEEE-754 parser such as Python/OpenPyXL can reconstruct the same binary64 value from a
round-trip-safe decimal token. A token containing too few significant digits can reconstruct a
nearby but different value.

Preservation after opening and resaving in Microsoft Excel is a separate application-fidelity
question. Microsoft documents Excel as using 15 significant digits of numeric precision, so a
17-digit serialization fix verified with OpenPyXL must not automatically be described as a
bit-preservation guarantee through the Excel desktop application.

The complete cycle is:

```text
typed decimal characters
        ↓ parse
binary floating-point value
        ↓ calculate and round after operations
binary floating-point result
        ↓ format for the screen
displayed decimal text
        ↓ serialize when saving
decimal characters in XLSX XML
        ↓ parse when reopening
binary floating-point value
```

## Direct copying versus recalculation

A direct copy of a constant numeric cell within Excel can preserve its stored numeric value. To
reproduce its appearance as well, the number format must also be copied.

The situation changes when:

- a formula is copied and recalculated;
- a destination formula uses a different evaluation order;
- a value passes through a file writer that does not emit a round-trip-safe numeric token;
- the destination uses a different number format; or
- **Set precision as displayed** is enabled.

This project's confirmed problem is the third case: OpenPyXL 3.1.5 can shorten the serialized XML
token enough to reconstruct a different binary value. That loss is additional to Excel's ordinary
floating-point calculation behavior.

## Confirmed example

```text
Before save: 100000.00000000001
XLSX XML:    100000
After reopen: 100000
Difference: -1.4551915228366852e-11
```

The binary64 value that Python represents as `0.1234523461234556` survived the tested OpenPyXL
save/reopen unchanged. The problem affects some values, not every number containing a fractional
part.

## What to check when a value looks wrong

1. Compare the Python values with `repr(value)`.
2. For floats with non-temporal formats, compare `value.hex()` before writing and after reopening
   the saved workbook. With date/time/duration formats, check the numeric XML token separately
   from the reader's temporal interpretation.
3. Check the actual worksheet XML `<v>` token when the bits changed.
4. Check the destination cell's `number_format` when only the visible text changed.
5. Determine whether the source cell contains a constant, formula text, or a cached formula result.
6. Identify the first pipeline boundary where the value changes; do not assume the reader caused
   it.

## Type choices

| Requirement | Preferred representation |
| --- | --- |
| Preserve an ordinary Excel numeric value | Python `float` → Polars `Float64` → DuckDB `DOUBLE` |
| Exact fixed-scale business decimal | `Decimal` / Polars Decimal / DuckDB `DECIMAL(p, s)` with an explicit scale and rounding policy |
| Preserve digit text exactly, without numeric behavior | Text/string |
| Copy what the user sees | Copy the value and deliberately handle `number_format`; visible text alone is not the stored value |

Do not silently write an unrepresentable `Decimal` as text. That changes Excel formulas, sorting,
filtering, and downstream type inference.

## Polars ingestion through DuckDB to Excel

For financial reporting, a representative preprocessing path is:

```text
Polars ingestion
    → register or load the frame in DuckDB
    → perform an explicitly typed group-by or other transformation
    → materialize the final result as an eager Polars DataFrame with `.pl()`
    → normalize it with `polars_adapters()` into canonical values
    → render those values into an .xlsx template
    → inspect numeric XML and reopen the workbook
```

DuckDB remains an input-preprocessing concern. The renderer does not execute SQL, perform grouping,
or provide a bundled DuckDB adapter. The caller must materialize the final ordered records before
rendering, either as ordinary Python values or through an explicit platform adapter. A Polars frame
materialized from DuckDB with `.pl()` uses the bundled `polars_adapters()` only after the SQL
transformation is complete; do not pass a live DuckDB relation or cursor to the renderer.

Use explicit numeric types at the DuckDB boundary:

- use `DOUBLE` only when the intended domain is binary64 passthrough;
- use `DECIMAL(p, s)` for fixed-scale financial amounts, with `p`, `s`, and the rounding rule stated
  by the business contract; and
- cast the final aggregate deliberately instead of relying on DuckDB's inferred aggregate width or
  scale.

Assert the actual registered and result types rather than relying on a bare `DECIMAL` declaration or
implicit casts. Division involving a DuckDB Decimal produces approximate floating-point output, so
division and any other type-changing transformation need an explicit business rounding and cast back
to the intended `DECIMAL(p, s)` before rendering.

A deterministic financial group-by test should use a query shaped like this:

```sql
SELECT
    reporting_key,
    CAST(
        ROUND(
            COALESCE(SUM(CAST(amount AS DECIMAL(15, 2))), 0),
            2
        ) AS DECIMAL(15, 2)
    ) AS reported_amount
FROM ingested_polars
GROUP BY reporting_key
ORDER BY reporting_key;
```

The precision and scale in this example are illustrative. Choose them from the reporting contract,
and ensure every final amount also satisfies the writer's supported-Decimal predicate. `ORDER BY`
is required because SQL group output order is otherwise unspecified. A final result `ORDER BY`
does not define the input order of an aggregate. If a diagnostic test intentionally aggregates
`DOUBLE` values and compares exact bits, give every source row a stable ordinal and use DuckDB's
aggregate `ORDER BY` syntax, such as `SUM(amount ORDER BY source_ordinal)`, while keeping the same
source order in the Excel reference. Prefer `DECIMAL` plus business-scale comparison for financial
tests because floating-point `SUM` is order-sensitive. `COALESCE` is intentional when the
Excel oracle treats a group containing only blanks as zero: DuckDB `SUM` over only `NULL` values is
`NULL`, while an Excel `SUMIFS`-style calculation commonly returns zero. If null and zero have
different business meanings, preserve the null instead.

### Group-by comparison protocol

Test each boundary independently so a DuckDB change is not mistaken for an XLSX writer regression:

1. Assert the ingested Polars schema and source scalars, including the distinction among null,
   floating-point NaN, and numeric zero.
2. Assert DuckDB's registered column types, the final query schema, group keys, and ordered result
   values. For `DECIMAL`, compare exact decimal values and the declared result scale.
3. Assert the eager Polars schema, row order, and scalar values returned by `.pl()`, then assert the
   canonical values produced by `polars_adapters()`. The template should only select the
   already-computed result, for example `{{ row.reported_amount }}`.
4. Render through the production XLSX path, assert that the cell is numeric, and compare the selected
   worksheet XML `<v>` token with the supported source amount.
5. Reopen the output with the pinned OpenPyXL version and apply the writer's documented
   non-temporal-format assertions.
6. For an Excel comparison, use an immutable workbook that desktop Excel fully recalculated and
   saved. Give Excel the same source rows and reporting keys, use an explicit `SUMIFS` plus `ROUND`
   formula (or another documented equivalent), and record the Excel build and calculation settings.

Cover all significant-digit counts from 1 through 15, several groups, duplicate keys, positive and
negative amounts, cancellation, zero, null-only groups, values on both sides of a rounding boundary,
and a supported 15-digit aggregate. Include a 16-digit or otherwise unsupported final Decimal as a
negative control and assert the writer diagnostic at the rendered destination cell.

DuckDB `DECIMAL` aggregation and Excel formula calculation do not use the same arithmetic domain:
DuckDB calculates the selected fixed-scale decimal operation exactly, while Excel calculates with
binary64 values and exposes at most 15 significant digits. Compare their results at the explicitly
declared reporting scale after the same business rounding rule. Do not use this test to claim raw-bit
equivalence, arbitrary `DOUBLE` reduction equivalence, identical reassociation, or general parity
among DuckDB, Polars, and Excel. The writer guarantee begins after the final value is materialized.

### Local group-by characterization

A local characterization run on 2026-10-05 exercised the complete path with DuckDB 1.5.5, Polars
1.43.2, desktop Excel 16.0 build 19127.0, and the production XLSX writer. This is recorded evidence,
not a default-CI guarantee or a pinned DuckDB dependency:

- all 15 quarter-unit aggregate cases spanning 1 through 15 significant digits matched between
  DuckDB `DECIMAL`, Excel `SUMIFS`, the written numeric XML token, and the reopened workbook;
- the writer preserved all 19 DuckDB Decimal aggregate results, and Polars materialization matched
  DuckDB in all 19;
- unrounded Excel results matched exact DuckDB Decimal arithmetic in 16 of 19 cases. Expected
  binary64 tails appeared for `0.1 + 0.2`, ten additions of `0.1`, and one 15-digit fractional sum;
- seven small `DOUBLE` aggregates matched DuckDB and Excel bit-for-bit, and the writer preserved all
  seven, while a separate native Polars group-by produced different bits in three cases; and
- a financial-scale comparison using DuckDB `ROUND(SUM(amount), 2)` and Excel
  `ROUND(SUMIFS(...), 2)` matched in all eight cases, including tenths, mixed cents, a 15-digit
  fractional input, positive and negative half-cent cases, and values immediately below and above a
  cent-rounding boundary. The XLSX token and reopened value also matched every rounded result.

These cases support preservation of the value DuckDB produced and selected shared business-rounded
results. They do not establish generic aggregate parity, especially for larger or reordered
floating-point groups, different DuckDB execution plans, other Excel builds, or formulas other than
the recorded `SUMIFS` cases.

Relevant DuckDB references:

- [Python and Polars integration](https://duckdb.org/docs/stable/clients/python/overview)
- [Aggregate functions and null handling](https://duckdb.org/docs/stable/sql/functions/aggregates)
- [Order-preservation rules](https://duckdb.org/docs/stable/sql/dialect/order_preservation)
- [Numeric and fixed-point types](https://duckdb.org/docs/stable/sql/data_types/numeric)

## Writer reminder

For a no-op copy with a format that the reader does not interpret as a date, time, or duration,
the desired invariant is:

```python
assert type(reopened_value) in (int, float)
assert float(reopened_value).hex() == source_value.hex()
```

Use a save-and-reopen integration test. Checking only the in-memory destination cell before
`Workbook.save()` will miss the confirmed OpenPyXL serialization loss. Test numeric type separately
if the public contract distinguishes an integer result from an equal floating-point result.

The current template-writer contract preserves finite built-in float bits and accepted Decimal
amounts in numeric XML for every authored format. Pinned OpenPyXL can interpret temporal formats
as temporal Python values or `#VALUE!` for out-of-range serials; this is not token loss and adds no
temporal fidelity guarantee. Accepted Decimals reopen as floats only under the same
non-temporal-format condition. Desktop Excel resave fidelity remains outside the guarantee.

Keep numeric serialization fixes at the XLSX writer boundary. Avoid changing the reader or forcing
all fractional values to `Decimal` merely to compensate for a writer dependency.

## Separate concerns

- **Value fidelity:** Did the numeric value change?
- **Display fidelity:** Is the number format the same?
- **Formula fidelity:** Was formula text, a cached result, or a recalculated result intended?
- **Decimal semantics:** Does the business domain require a particular scale and rounding mode?
- **Excel application fidelity:** Must the value also survive opening, recalculating, editing, and
  resaving in Microsoft Excel?

Answer these separately. A single comparison of displayed cell text cannot distinguish them.

## Practical rules

- Use ordinary binary floating-point types when the goal is to carry normal Excel numbers through
  Python, Polars, and DuckDB without changing their numeric domain.
- Use `ROUND` at a deliberate business boundary when the requirement is expressed in decimal
  places, such as cents or a published reporting scale.
- Use `Decimal` and DuckDB `DECIMAL(p, s)` when the business contract requires exact decimal scale
  and rounding. Decide explicitly how that value will be represented when written back to Excel.
- Store identifiers such as credit-card numbers as text when all digits must be preserved and the
  value is not meant for arithmetic.
- Do not rely on the visible cell text to test numeric equality.
- Avoid **Set precision as displayed** unless permanently discarding hidden precision is explicitly
  intended.

## Official Microsoft references

- [Floating-point arithmetic may give inaccurate results in Excel](https://learn.microsoft.com/en-us/troubleshoot/microsoft-365-apps/excel/floating-point-arithmetic-inaccurate-result)
  explains Excel's IEEE 754 arithmetic, repeating binary fractions, 15-digit precision, propagation
  through repeated calculations, and Excel's limited near-zero correction.
- [Change formula recalculation, iteration, or precision in Excel](https://support.microsoft.com/en-us/excel/change-formula-recalculation-iteration-or-precision-in-excel)
  explains that Excel normally calculates using stored rather than displayed values and documents
  **Set precision as displayed**.
- [Excel specifications and limits](https://support.microsoft.com/en-us/excel/excel-specifications-and-limits)
  lists Excel's 15-digit number-precision limit.
- [Format numbers as text](https://support.microsoft.com/en-us/excel/format-numbers-as-text)
  explains why identifiers with more than 15 digits should be stored as text rather than numbers.
