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
