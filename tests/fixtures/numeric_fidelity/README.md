# Frozen desktop-Excel numeric reference

`excel_numeric_reference.xlsx` is authored, fully recalculated, and saved by installed Microsoft
Excel through the adjacent PowerShell COM builder. The JSON manifest records the actual Excel
version/build and executable file version, culture/country/decimal separator, UTC creation time,
calculation settings, authoritative named cells, comparison modes, and final file SHA-256.

`NumericReference!B2:B16` contains named `precision_01` through `precision_15` literal formulas
`=0.3`, `=0.31`, …, `=0.314159265358979`. These final amounts provide decimal-transport evidence
for significant precision 1–15; they do not establish Excel/Polars decimal arithmetic parity.
`B19:B22` contains `exact_add` (`=0.5+0.25`), `exact_subtract` (`=1.5-0.25`), `exact_multiply`
(`=1.125*2`), and `exact_divide` (`=5/2`). Only those four cases establish operation parity:
identical exact binary-fraction inputs and a single explicit operation, compared by binary64 bits.

Regenerate from the repository root on Windows with licensed desktop Excel installed:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/fixtures/numeric_fidelity/build_excel_reference.ps1
```

The builder launches a hidden Excel instance, disables alerts, events and macros, selects automatic
calculation, disables iteration and Precision as displayed, calls `CalculateFullRebuild`, and saves
macro-free format 51 to a temporary `.xlsx` before replacing the reference. It closes its workbook,
quits Excel, and releases COM objects in `finally`, including failure paths. Review the workbook
and regenerated manifest together; timestamps, Excel build and checksum may legitimately change.
Never invoke the builder in default CI. No live Excel or `pywin32` dependency is required by pytest.

Validate portably from the repository root:

```powershell
uv run --all-extras pytest tests/xlsx/test_excel_numeric_reference.py tests/adapters/test_polars.py -q --basetemp=.pytest_cache/excel-reference
uv run --all-extras pytest --basetemp=.pytest_cache/full-suite
uv run --all-extras ruff check src tests samples scratch
uv run --all-extras ruff format --check src tests samples scratch
uv run --all-extras ty check
git diff --check
```

Tests read formula text with `data_only=False`, cached numbers with `data_only=True`, and selected
OOXML formula/cache nodes. The checksum detects accidental fixture replacement; tests compare
semantic numeric content rather than treating workbook bytes as a universal golden output.
Excel can write a longer cache token (for example `0.31409999999999999` for `=0.3141`); tests
compare the decoded numeric amount and selected binary64 identities, not literal XML spelling.

Never resave this reference with OpenPyXL: OpenPyXL does not calculate formulas and a save can
discard the Excel-authored cached results. Use OpenPyXL only to read it; regenerate exclusively
with the documented Excel builder. Do not hand-edit its ZIP/XML or convert it to `.xlsm`.

This reference is evidence from the recorded Excel build and locale, not a universal formula
oracle. It makes no claim about generic arithmetic equivalence, reductions, reassociation,
non-exact decimal operations, subnormal/extreme-number behavior in Excel, or preservation after
Excel opens and resaves renderer output. The renderer guarantee remains the pinned OpenPyXL
save-and-reopen contract in `SPEC.md`.
