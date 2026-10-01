# XLSX Numeric Fidelity Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development
> (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use
> checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve finite built-in floats bit-for-bit through XLSX save/reopen and write supported
financial Decimals as numeric cells without silent rounding.

**Architecture:** Keep the core render plan unchanged, then perform pure Decimal validation and
adaptation at the XLSX writer boundary. Activate one version-checked, process-lifetime OpenPyXL
formatter wrapper before constructing the destination workbook; its lock protects activation only,
not workbook saves.

**Tech Stack:** CPython 3.12, uv, OpenPyXL 3.1.5, pytest, Hypothesis, optional Polars 1.x,
SpreadsheetML XML, and a frozen desktop-Excel reference workbook.

**Spec:** `docs/superpowers/specs/2026-10-01-xlsx-numeric-fidelity-design.md`

## Global Constraints

- Update normative `SPEC.md` before changing runtime behavior.
- Support `.xlsx` only; do not add `.xlsm` or public write-only rendering.
- Pin `openpyxl==3.1.5` and keep the compatibility allowlist identical to that pin.
- Preserve every finite value whose exact type is built-in `float` through the pinned OpenPyXL
  save/reopen path, including signed zero, subnormals, and finite extremes.
- Keep every finite Decimal canonical in the core; apply the 15-significant-digit policy only to
  final numeric planned cells.
- Never round, stringify, or publish partial output for an unsupported Decimal.
- Keep the interpreter, evaluator, and layout planner free of OpenPyXL imports and workbook-specific
  conversion.
- The process-global activation lock must not surround workbook saves; concurrent writes use
  separate Workbook instances and output paths.
- Preserve the existing untracked `outputs/` directory and the three existing `scratch/*.py` files.
- Build binary fixtures and samples through documented tools; never hand-edit `.xlsx` files.

## Review Focus

- A Decimal with more than 15 coefficient digits but only trailing insignificant zeros must be
  accepted; Task 2 pins this with
  `test_numeric_preflight_treats_scale_and_trailing_zeros_as_presentation`.
- A float subclass must delegate to stock OpenPyXL behavior rather than receiving the built-in-float
  guarantee; Task 3 pins this in `test_wrapper_delegates_every_non_builtin_finite_float`.
- Concurrent first use must run one activation/self-test, while later independent saves remain
  concurrent; Task 3 pins both behaviors in subprocess tests.
- Temporal serials may move at a half-millisecond tie but must match pinned OpenPyXL conversion for
  both numeric epochs, while ISO-date mode must continue to bypass numeric formatting; Task 3 pins
  both paths.
- Repeated invalid Decimals must report every rendered destination and leave a missing or existing
  output untouched; Task 4 pins destination ordering and both atomicity cases.

---

### Task 1: Lock the normative contract and dependencies

**Files:**
- Modify: `SPEC.md:326-340,451-459,1023-1074,1185-1200,1256-1265`
- Modify: `docs/superpowers/specs/2026-10-01-xlsx-numeric-fidelity-design.md`
- Modify: `pyproject.toml`
- Modify: `uv.lock`

**Interfaces:**
- Consumes: approved design at
  `docs/superpowers/specs/2026-10-01-xlsx-numeric-fidelity-design.md`.
- Produces: normative float/Decimal/XLSX contract, reserved E3203-E3207 codes, exact dependency pin,
  and Hypothesis test dependency used by Tasks 2-4.

- [ ] **Step 1: Update `SPEC.md` before runtime code**

Document all of these exact decisions:

- exact built-in finite floats use a round-trip-safe numeric XML token and reopen as `float` with
  identical `float.hex()`;
- the guarantee is a pinned-library round trip, not desktop-Excel resave behavior;
- Decimal numeric cells use the five predicates from the design, in precision/range/token-equality
  order;
- mathematical precision strips coefficient trailing zeros without `Decimal.normalize()`;
- accepted Decimal values become built-in floats and lose type, exponent, scale, and signed-zero
  identity;
- integers and temporal values receive no new public fidelity guarantee;
- all sheet plans finish before numeric preflight, which precedes workbook creation;
- E3203-E3207 meanings, locations, and atomic publication behavior;
- portable Excel-reference tests versus optional live Excel; and
- the exact OpenPyXL pin/upgrade certification rule.

- [ ] **Step 2: Clarify the design's activation terminology**

Replace ambiguous uses of “installation” for the runtime hook with “in-memory activation” and state
that it neither installs a package nor edits OpenPyXL files on disk. Preserve “dependency
installation” where it refers to uv.

- [ ] **Step 3: Pin OpenPyXL and add Hypothesis through uv**

Run:

```powershell
uv add "openpyxl==3.1.5"
uv add --dev "hypothesis>=6,<7"
```

Expected: `pyproject.toml` contains the exact OpenPyXL pin and Hypothesis dev range; `uv.lock` is
updated without changing the CPython 3.12 or Polars contracts.

- [ ] **Step 4: Verify the contract and lock**

Run:

```powershell
uv lock --check
uv run python -c "from importlib.metadata import version; assert version('openpyxl') == '3.1.5'"
uv run pytest tests/test_documentation.py tests/test_architecture.py -q
```

Expected: lock check succeeds, version assertion succeeds, and the focused tests pass.

- [ ] **Step 5: Commit**

```powershell
git add SPEC.md docs/superpowers/specs/2026-10-01-xlsx-numeric-fidelity-design.md pyproject.toml uv.lock
git commit -m "docs: specify xlsx numeric fidelity contract"
```

---

### Task 2: Add pure Decimal numeric preflight

**Files:**
- Create: `src/excel_template_writer/xlsx/numeric.py`
- Create: `tests/xlsx/test_numeric_preflight.py`
- Modify: `src/excel_template_writer/diagnostics.py`
- Modify: `tests/test_render_plan.py`
- Modify: `tests/test_values.py`

**Interfaces:**
- Consumes: frozen `RenderPlan` and `PlannedCell` dataclasses plus E3203-E3205 from Task 1's
  normative contract.
- Produces:
  `NumericPreflightResult(plans: tuple[RenderPlan, ...] | None, diagnostics: tuple[Diagnostic, ...])`,
  `NumericPreflightResult.require() -> tuple[RenderPlan, ...]`, and
  `preflight_numeric_plans(plans: tuple[RenderPlan, ...]) -> NumericPreflightResult`.

- [ ] **Step 1: Write the failing pure-policy and boundary tests**

Add these tests to `tests/xlsx/test_numeric_preflight.py`:

- `test_numeric_preflight_converts_supported_decimals_in_adapter_copies_only`;
- `test_numeric_preflight_preserves_all_planned_cell_provenance`;
- `test_numeric_preflight_passes_non_decimal_values_through_unchanged`;
- `test_numeric_preflight_accepts_each_precision_from_one_through_fifteen` using prefixes of
  `314159265358979`, both signs, four exponent placements, and appended zeros;
- `test_numeric_preflight_treats_scale_and_trailing_zeros_as_presentation`;
- `test_numeric_preflight_precision_is_independent_of_decimal_context`;
- `test_numeric_preflight_accepts_zero_and_normal_range_boundaries`;
- `test_numeric_preflight_rejects_sixteen_significant_digits`;
- `test_numeric_preflight_rejects_overflow_underflow_and_subnormal_values`;
- `test_numeric_preflight_reports_inexact_decimal_conversion` using a Decimal subclass that
  overrides `__float__`; do not invent a built-in Decimal witness;
- `test_numeric_preflight_applies_precision_range_inexact_precedence`;
- `test_numeric_preflight_aggregates_each_rendered_destination`;
- `test_numeric_preflight_returns_no_plans_when_any_decimal_fails`; and
- `test_generated_decimal_classification_matches_public_predicate` using a custom Hypothesis
  Decimal-tuple strategy with 1-20 digits and boundary exponents.

The deterministic range matrix accepts `±2.22507385850721E-308` and
`±1.79769313486231E+308`; it rejects `±2.22507385850720E-308`, `±1E-308`,
`±1E-323`, `±1E-324`, `±1.79769313486232E+308`, and `±1E+309`. Precision tests accept
`±123456789012345` and `±999999999999999`, then reject
`±1234567890123456` and `±1.234567890123456`.

The essential accepted assertion is:

```python
result = preflight_numeric_plans((plan,))
adapted = result.require()[0]
assert type(adapted.cells[0].value) is float
assert Decimal(repr(adapted.cells[0].value)) == source_decimal
assert plan.cells[0].value.as_tuple() == source_decimal.as_tuple()
```

Add `test_render_plan_retains_decimal_unchanged_until_xlsx_preflight` with a 16-digit Decimal and
`test_normalization_keeps_xlsx_unsupported_finite_numeric_values_canonical` with that Decimal plus
a finite float subnormal.

- [ ] **Step 2: Run the tests to verify RED**

Run:

```powershell
uv run pytest tests/xlsx/test_numeric_preflight.py tests/test_render_plan.py tests/test_values.py -q
```

Expected: FAIL because `xlsx.numeric` and E3203-E3205 do not exist.

- [ ] **Step 3: Implement the pure preflight**

Add these exact diagnostic members:

```python
XLSX_DECIMAL_PRECISION_EXCEEDED = "E3203"
XLSX_DECIMAL_OUT_OF_RANGE = "E3204"
XLSX_DECIMAL_INEXACT = "E3205"
```

Implement `NumericPreflightResult` and `preflight_numeric_plans` without importing OpenPyXL.
Iterate plans and cells in their existing deterministic order. For each Decimal:

1. count coefficient digits after stripping trailing zeros, with zero precision equal to one;
2. reject precision above 15;
3. convert with `float(value)` and reject nonfinite, zero-underflow, or nonzero magnitude below
   `sys.float_info.min`;
4. reject when `Decimal(repr(converted)) != value`; and
5. otherwise use `dataclasses.replace` on the cell and plan.

Use `SourceLocation(plan.sheet, cell.coordinate.a1)`. Aggregate one diagnostic per invalid
destination. If any fail, return no plans. `require()` raises `TemplateRenderError`.

- [ ] **Step 4: Run focused tests to verify GREEN**

Run:

```powershell
uv run pytest tests/xlsx/test_numeric_preflight.py -q
uv run pytest tests/test_render_plan.py tests/test_values.py -k "decimal or numeric" -q
```

Expected: all selected tests pass.

- [ ] **Step 5: Commit**

```powershell
git add src/excel_template_writer/diagnostics.py src/excel_template_writer/xlsx/numeric.py tests/xlsx/test_numeric_preflight.py tests/test_render_plan.py tests/test_values.py
git commit -m "feat: validate decimal xlsx numeric cells"
```

---

### Task 3: Add guarded OpenPyXL numeric activation

**Files:**
- Create: `src/excel_template_writer/xlsx/openpyxl_numeric_compat.py`
- Create: `tests/xlsx/test_openpyxl_numeric_compat.py`
- Modify: `src/excel_template_writer/diagnostics.py`

**Interfaces:**
- Consumes: pinned OpenPyXL 3.1.5 and the shared diagnostic model.
- Produces:
  `ensure_openpyxl_numeric_compatibility() -> Diagnostic | None` and E3206/E3207 for Task 4.

- [ ] **Step 1: Write failing formatter and activation tests**

Direct, non-activating tests must prove:

- stock OpenPyXL writes `100000.00000000001` as `100000`;
- the owned wrapper emits `repr(value)` for fixed values and 1,000-2,000 generated finite raw-bit
  binary64 values whose parsed `float.hex()` is identical;
- `test_wrapper_delegates_every_non_builtin_finite_float` proves integers, booleans, Decimal, None,
  strings, datetimes, nonfinite floats, and a float subclass delegate to a spy original formatter;
- both OpenPyXL cell-writer functions resolve the replaced module-global hook; and
- `test_temporal_serials_match_pinned_openpyxl_conversion` proves numeric temporal serials match
  `from_excel(to_excel(value, epoch), epoch, timedelta=...)` for the 1900/1904 epochs and
  half-millisecond boundaries, while `iso_dates=True` continues to emit and reload ISO temporal
  values without numeric formatter involvement.

Fresh-subprocess tests must prove:

- first activation, idempotence, and module-reload ownership recognition;
- raw normal and write-only workbook round trips;
- version mismatch returns E3206 at `<workbook>!A1`;
- missing/foreign/malformed hooks and self-test failure return E3207;
- failed self-test restores stock only while the project still owns the hook;
- post-activation replacement becomes sticky failure without overwriting the foreign hook;
- `test_concurrent_activation_runs_one_self_test` has eight callers wait on one gated self-test and
  receive success; and
- `test_activated_process_saves_independent_workbooks_concurrently` proves two saves overlap after
  activation.

- [ ] **Step 2: Run the compatibility tests to verify RED**

Run:

```powershell
uv run --all-extras pytest tests/xlsx/test_openpyxl_numeric_compat.py -q
```

Expected: FAIL because the compatibility module and E3206-E3207 do not exist.

- [ ] **Step 3: Implement the owned wrapper and state machine**

Add:

```python
XLSX_OPENPYXL_COMPAT_UNVERIFIED = "E3206"
XLSX_OPENPYXL_COMPAT_FAILED = "E3207"
```

Implement a callable owned wrapper retaining its original formatter and stable owner/wrapper/target
metadata. It returns `repr(value)` only for `type(value) is float and math.isfinite(value)`.

Use an exact `"3.1.5"` certification constant, `UNINSTALLED -> INSTALLING -> INSTALLED/FAILED`
state, one lock, one active-wrapper reference, and one cached failure Diagnostic. Hold the lock only
through activation/self-test. On the installed fast path, verify hook identity and return without
holding a lock during any save.

The in-memory self-test writes and reopens `100000.00000000001` using OpenPyXL primitives, never the
public renderer. Restore the original only if the hook still points to the owned wrapper. Unknown
versions return E3206; all hook/ownership/self-test failures return E3207.

- [ ] **Step 4: Run compatibility tests to verify GREEN**

Run:

```powershell
uv run --all-extras pytest tests/xlsx/test_openpyxl_numeric_compat.py -q
```

Expected: all tests pass without leaving the parent pytest process activated or failed.

- [ ] **Step 5: Commit**

```powershell
git add src/excel_template_writer/diagnostics.py src/excel_template_writer/xlsx/openpyxl_numeric_compat.py tests/xlsx/test_openpyxl_numeric_compat.py
git commit -m "fix: activate round-trip-safe openpyxl numeric writing"
```

---

### Task 4: Integrate the writer and prove production fidelity

**Files:**
- Create: `tests/xlsx/test_numeric_fidelity.py`
- Modify: `tests/adapters/test_polars.py`
- Modify: `src/excel_template_writer/xlsx/writer.py`

**Interfaces:**
- Consumes: `preflight_numeric_plans(...)` from Task 2 and
  `ensure_openpyxl_numeric_compatibility()` from Task 3.
- Produces: the unchanged `write_workbook(...) -> Path` interface with preflight and activation
  enforced before destination mutation.

- [ ] **Step 1: Write failing production XLSX tests**

Add production-API tests for:

- `test_render_workbook_round_trips_fixed_float_matrix_bit_exactly`: `1.0`, both zeros, known
  formatter regressions, signs,
  exponent forms, `nextafter` neighbors, minimum subnormal, `sys.float_info.min`, and
  `sys.float_info.max`;
- `test_float_xml_tokens_use_repr_and_remain_numeric`: selected worksheet `<v>` tokens;
- `test_render_workbook_round_trips_generated_finite_float_batches`: about 30 Hypothesis XLSX
  examples;
- `test_supported_decimal_is_numeric_in_xml_and_preserves_number_format`: accepted matrices through
  15 digits, reopened float bits, and formatting;
- `test_decimal_scale_variants_have_equal_amounts_and_authored_formats`;
- `test_decimal_policy_does_not_reject_unused_or_mixed_text_values`;
- `test_decimal_preflight_runs_before_output_directory_creation`;
- `test_decimal_preflight_preserves_existing_output_bytes`;
- `test_repeated_invalid_decimals_report_each_destination_coordinate`; and
- `test_openpyxl_compatibility_failure_is_atomic_before_workbook_creation`, with a monkeypatched
  compatibility Diagnostic and `Workbook()` sentinel.

Add Polars tests that compute an explicit `result` column before rendering:

- `test_polars_float_operation_result_round_trips_bit_exactly` for `0.1 + 0.2`, asserting Polars
  dtype/scalar, adapter bits, XML token, and reopened bits;
- `test_polars_decimal_operations_cover_one_through_fifteen_digits` using exact Decimal addition
  for every precision and asserting dtype/scalar before rendering;
- `test_polars_sixteen_digit_decimal_result_is_rejected_at_destination_cell`; and
- parameterized exact binary-fraction operations `0.5 + 0.25`, `1.5 - 0.25`,
  `1.125 * 2`, and `5 / 2` for later comparison with the frozen Excel reference.

- [ ] **Step 2: Run production tests to verify RED**

Run:

```powershell
uv run --all-extras pytest tests/xlsx/test_numeric_fidelity.py tests/adapters/test_polars.py -q
```

Expected: FAIL because `write_workbook` neither invokes preflight nor activates the formatter;
known floats reopen with changed bits and invalid Decimals publish output.

- [ ] **Step 3: Integrate both boundaries at the top of `write_workbook`**

Before `Path(output_path)`, parent-directory creation, or `Workbook()`:

```python
adapted_plans = preflight_numeric_plans(plans).require()
compatibility_diagnostic = ensure_openpyxl_numeric_compatibility()
if compatibility_diagnostic is not None:
    raise TemplateRenderError((compatibility_diagnostic,))
```

Use `adapted_plans` in the existing sheet/plan/feature-plan zip. Keep all expression, layout,
presentation, package verification, temporary-file, reopen, and atomic-replace behavior unchanged.
Document the new `TemplateRenderError` causes in the function docstring.

- [ ] **Step 4: Run production and regression tests to verify GREEN**

Run:

```powershell
uv run --all-extras pytest tests/xlsx/test_numeric_fidelity.py tests/adapters/test_polars.py -q
uv run --all-extras pytest tests/xlsx/test_workbook_renderer.py tests/xlsx/test_default_style.py -q
```

Expected: all tests pass. Existing integer, boolean, formula, style, merge, drawing, validation, and
atomic-save tests remain green.

- [ ] **Step 5: Commit**

```powershell
git add src/excel_template_writer/xlsx/writer.py tests/xlsx/test_numeric_fidelity.py tests/adapters/test_polars.py
git commit -m "feat: preserve numeric fidelity in xlsx output"
```

---

### Task 5: Add the portable frozen Excel reference

**Files:**
- Create: `tests/fixtures/numeric_fidelity/build_excel_reference.ps1`
- Create: `tests/fixtures/numeric_fidelity/excel_numeric_reference.xlsx`
- Create: `tests/fixtures/numeric_fidelity/excel_numeric_reference.json`
- Create: `tests/fixtures/numeric_fidelity/README.md`
- Create: `tests/xlsx/test_excel_numeric_reference.py`
- Modify: `tests/adapters/test_polars.py`

**Interfaces:**
- Consumes: production renderer from Task 4 and desktop Excel only while regenerating the fixture.
- Produces: immutable Excel-calculated evidence consumed portably by pytest without Excel installed.

- [ ] **Step 1: Write failing fixture/reference tests**

Add:

- `test_excel_reference_fixture_matches_recorded_provenance_and_checksum`;
- `test_excel_reference_contains_documented_formulas_and_cached_values`;
- `test_excel_reference_selected_xml_nodes_have_formula_and_cache`;
- `test_rendered_decimal_precision_matrix_matches_frozen_excel_results`; and
- `test_polars_exact_binary_operations_match_frozen_excel_reference`.

The 1-15 precision cells use formula literals for the final amounts
`0.3` through `0.314159265358979`. They are a decimal-transport oracle, not a claim that Excel and
Polars performed identical decimal arithmetic. Only the four exact binary-fraction formulas are
operation-parity cases.

Run:

```powershell
uv run --all-extras pytest tests/xlsx/test_excel_numeric_reference.py -q
```

Expected: FAIL because the fixture and manifest do not exist.

- [ ] **Step 2: Add the documented hidden-Excel fixture builder**

The PowerShell builder must:

- create `Excel.Application` with `Visible = $false`, alerts off, and macro security forced off;
- create a new workbook, automatic calculation, and `PrecisionAsDisplayed = $false`;
- write named precision/formula cells, call `CalculateFullRebuild()`, and save format 51 (`.xlsx`)
  to a temporary path before replacement;
- record Excel version/build, file version, locale/decimal separator, calculation settings, UTC
  creation time, authoritative cell map/comparison mode, and SHA-256 in JSON;
- close every workbook and quit/release every COM object in `finally`; and
- never run in default CI.

The README must state exact regeneration and validation commands and that OpenPyXL must never resave
the reference.

- [ ] **Step 3: Generate the fixture through desktop Excel**

Run the builder on this Windows machine, then:

```powershell
uv run --all-extras pytest tests/xlsx/test_excel_numeric_reference.py -q
```

Expected: all reference, cache, XML, checksum, renderer, and exact-operation comparisons pass.

- [ ] **Step 4: Commit**

```powershell
git add tests/fixtures/numeric_fidelity tests/xlsx/test_excel_numeric_reference.py tests/adapters/test_polars.py
git commit -m "test: add frozen excel numeric reference"
```

---

### Task 6: Update maintained samples, user docs, and run the complete gate

**Files:**
- Modify: `samples/scalar_values.py`
- Modify: `samples/scalar_values_template.xlsx`
- Modify: `samples/scalar_values_output.xlsx`
- Modify: `samples/polars_dataframe.py`
- Modify: `samples/polars_dataframe_template.xlsx`
- Modify: `samples/polars_dataframe_output.xlsx`
- Modify: `samples/README.md`
- Modify: `tests/test_samples.py`
- Modify: `README.md`
- Modify: `docs/directives.md`
- Modify: `docs/explained.md`
- Modify: `docs/DEVELOPMENT.md`

**Interfaces:**
- Consumes: completed public numeric behavior and diagnostics from Tasks 2-5.
- Produces: executable examples, generated matching workbooks, user-facing contract documentation,
  fixture maintenance guidance, and complete verification evidence.

- [ ] **Step 1: Write failing maintained-sample assertions**

Add `test_scalar_values_sample_demonstrates_numeric_fidelity` and
`test_polars_sample_demonstrates_computed_numeric_writeback`. Locate cases by their visible labels,
then assert numeric data type, exact float bits or Decimal token equality, and authored format.

Run:

```powershell
uv run --all-extras pytest tests/test_samples.py -q
```

Expected: FAIL because the committed samples do not contain the new labeled cases.

- [ ] **Step 2: Extend executable samples**

In `scalar_values.py`:

- make the existing invoice amount an actual Decimal;
- append a formerly lossy `100000.00000000001` float;
- append a supported 15-digit financial Decimal such as `1234567890123.45`;
- append `Decimal("12.50")` with an authored two-decimal format; and
- assert numeric type, float bits/Decimal amount, and formats after reopen.

In `polars_dataframe.py`, compute an explicit Polars Decimal result column before rendering, assert
its dtype and scalar, render that column, and verify the reopened numeric amount and format. Keep
input preprocessing in Polars, not in the template language.

- [ ] **Step 3: Regenerate only the two affected sample pairs**

Run:

```powershell
uv run --all-extras python -m samples.scalar_values
uv run --all-extras python -m samples.polars_dataframe
uv run --all-extras pytest tests/test_samples.py -q
```

Expected: both builders and all sample catalog tests pass; unrelated binary samples are untouched.

- [ ] **Step 4: Update user and contributor documentation**

Add concise numeric-output guidance to:

- `README.md`: float bit fidelity, Decimal 15-digit eligibility, numeric-not-text output, and the
  library-versus-desktop-Excel distinction;
- `docs/directives.md`: stored value versus number format, caller-owned rounding, and E3203-E3207;
- `docs/explained.md`: pure render plan -> numeric preflight -> activation -> writer flow; and
- `docs/DEVELOPMENT.md`: frozen Excel fixture provenance, checksum, cached-formula inspection, and
  no-OpenPyXL-resave rule.

- [ ] **Step 5: Run focused checks**

Run:

```powershell
uv run --all-extras pytest tests/test_samples.py tests/xlsx/test_numeric_preflight.py tests/xlsx/test_openpyxl_numeric_compat.py tests/xlsx/test_numeric_fidelity.py tests/xlsx/test_excel_numeric_reference.py tests/adapters/test_polars.py -q
```

Expected: all focused contract, integration, reference, adapter, and sample tests pass.

- [ ] **Step 6: Run the repository's complete gate**

Run:

```powershell
uv run --all-extras pytest
uv run --all-extras ruff check src tests samples scratch
uv run --all-extras ruff format --check src tests samples scratch
uv run --all-extras ty check
```

Expected: every command exits zero.

- [ ] **Step 7: Commit**

```powershell
git add README.md docs/DEVELOPMENT.md docs/directives.md docs/explained.md samples/scalar_values.py samples/scalar_values_template.xlsx samples/scalar_values_output.xlsx samples/polars_dataframe.py samples/polars_dataframe_template.xlsx samples/polars_dataframe_output.xlsx samples/README.md tests/test_samples.py
git commit -m "docs: demonstrate deterministic numeric writing"
```
