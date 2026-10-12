# Missing Collection Recovery Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Extend render-scoped `MissingOutputPolicy.BLANK` and `.PRESERVE` to unavailable repeat collections and null collection-filter receivers, while keeping strict rendering as the default.

**Architecture:** The expression evaluator distinguishes a null collection receiver from other filter type errors. The pure renderer recovers a missing or null `ForNode` collection into one source-sized block, with raw authored values or blank values and identity presentation provenance. The XLSX writer consumes the ordinary completed plan without new layout logic.

**Tech Stack:** CPython 3.12, `uv`, pytest, Ruff, ty, OpenPyXL 3.1.5.

**Spec:** `docs/superpowers/specs/2026-09-29-missing-collection-recovery-design.md`

## Global Constraints

- Update normative `SPEC.md` before changing rendering behavior; the current text explicitly rejects missing repeat collections under tolerant policies.
- Keep the interpreter and layout planner independent of OpenPyXL; parse, validate, and plan before workbook mutation.
- Preserve exact authored rectangles, source cell presentation, row dimensions, and contained merges in recovered repeats; never infer geometry from blank cells or styles.
- `error` remains the default; missing `if` conditions and independent template, type, layout, and workbook errors remain fatal.
- Every user-visible change needs executable Python and matching template/output XLSX workbooks under `samples/`.
- Run focused checks and the full configured suite through `uv run --all-extras`.

## Review Focus

- Missing root `tables` and missing indexed key `tables["some_table"]` should each recover once at the loop opening cell; Task 3 tests both.
- An unavailable child repeat inside a valid parent repeat should recover and warn once per parent instance; Task 3 tests this.
- Recovery under `rows`, `cells`, and `none` should keep source-sized geometry and leave adjacent content at its original destination; Task 3 tests each shift mode.
- A wrong non-null collection value or an independent error should still prevent a plan or workbook; Tasks 2–4 test these boundaries.
- Missing `if` conditions and ordinary empty collections should retain their existing behavior; Task 3 tests both.

---

### Task 1: Make the normative contract agree with the reviewed design

**Files:**
- Modify: `SPEC.md`

**Interfaces:**
- Consumes: the approved design linked above and the current `MissingOutputPolicy` vocabulary.
- Produces: normative rules for unavailable repeat and null collection-filter recovery used by all later tasks.

- [ ] **Step 1: Update `SPEC.md`** in sections 2, 5.5, 8.6, 9.7, 15.7, and 16. Define missing versus present null, the `error`/`blank`/`preserve` matrix, one-block identity geometry, descendant suppression, warning locations and cardinality, strict `if` behavior, and the `W1302 NULL_COLLECTION_RENDERED` code. Keep existing scalar output-tag and empty-collection rules explicit.
- [ ] **Step 2: Check the contract for contradictions.** Run `rg -n 'missing.*for|MissingOutputPolicy|W1302|NULL_COLLECTION_RENDERED|collection filter' SPEC.md` and review every match against the approved design. Expected: no remaining statement says a missing `for` collection is fatal under all policies.
- [ ] **Step 3: Commit** `SPEC.md` with `docs: define missing collection recovery semantics`.

### Task 2: Distinguish null collection-filter receivers from wrong types

**Files:**
- Modify: `src/excel_template_writer/expressions.py`
- Modify: `src/excel_template_writer/render.py`
- Modify: `src/excel_template_writer/diagnostics.py`
- Test: `tests/test_expressions.py`
- Test: `tests/test_missing_output.py`

**Interfaces:**
- Consumes: existing `evaluate_expression(expression: Expression, scope: Mapping[str, Any]) -> Any` and `MissingOutputPolicy`.
- Produces: `NullCollectionError(FilterTypeError)` for a canonical-null receiver of `join`, `sum`, `min`, `max`, or `count`; `DiagnosticCode.NULL_COLLECTION_RENDERED = "W1302"`; output-tag recovery for this exception only under `blank`/`preserve`.

- [ ] **Step 1: Write failing evaluator tests** parameterized over `join`, `sum`, `min`, `max`, and `count`: `{{ items | FILTER }}` with `items=None` raises `NullCollectionError`; `items=7` raises `FilterTypeError`, never a raw `TypeError`. Assert existing empty collection results stay unchanged.
- [ ] **Step 2: Run** `uv run --all-extras pytest tests/test_expressions.py -q`. Expected: the new null-receiver tests fail for the missing dedicated exception, and the existing tests pass.
- [ ] **Step 3: Implement the dedicated exception and receiver checks** in `expressions.py`. Check the collection receiver before `join` iteration or aggregate dispatch; do not make non-null invalid inputs recoverable.
- [ ] **Step 4: Write failing render tests** in `tests/test_missing_output.py`: for each filter, `error` reports `E1304` with no plan; `blank` yields `None` and `W1302`; `preserve` yields the exact output tag and `W1302`. Assert warning source location and severity. Add one wrong-type output test that remains `E1304` under both tolerant modes.
- [ ] **Step 5: Run** `uv run --all-extras pytest tests/test_missing_output.py -q`. Expected: new tolerant null-filter cases fail; existing missing-output tests pass.
- [ ] **Step 6: Add `W1302` and catch `NullCollectionError` before `FilterTypeError`** at the output-tag boundary in `render.py`. Preserve the existing whole-tag and mixed-cell assembly rules; strict mode must still emit `E1304`.
- [ ] **Step 7: Run** `uv run --all-extras pytest tests/test_expressions.py tests/test_missing_output.py -q`. Expected: all pass.
- [ ] **Step 8: Commit** the expression, renderer, diagnostic, and test changes with `feat: recover null collection filters at output tags`.

### Task 3: Recover an unavailable repeat as one identity-sized block

**Files:**
- Modify: `src/excel_template_writer/render.py`
- Test: `tests/test_missing_output.py`
- Test: `tests/test_render_plan.py`

**Interfaces:**
- Consumes: `ForNode.rectangle`, immutable `compiled.template.cells`, existing `_Block`, `PlannedCell`, `PlannedRow`, `PlannedMerge`, and `W1302` from Task 2.
- Produces: `_Renderer.render_unavailable_repeat(self, node: ForNode, path: tuple[int, ...], *, preserve: bool) -> _Block`, a pure source-sized block with exact raw values or `None`, source presentation coordinates, source rows, contained merges, and the normal resource-limit and collision checks.

- [ ] **Step 1: Write failing policy tests** for missing `tables` and missing `tables["some_table"]` in a multi-cell, multi-row loop. Assert default `E1301` and no plan; `blank` clears every authored value, including static text and directives, and emits one `W1301`; `preserve` retains exact raw cell strings and emits one `W1301`. Assert warning at the opening directive.
- [ ] **Step 2: Run** `uv run --all-extras pytest tests/test_missing_output.py -q`. Expected: the new tolerant repeat tests fail.
- [ ] **Step 3: Implement missing-repeat recovery** in the `ForNode` evaluation path before empty-repeat planning. Build one local identity block from the immutable source rectangle, skip all descendant evaluation, and keep the existing error path for `error`. Do not route a failed collection through the empty-list placeholder.
- [ ] **Step 4: Run** `uv run --all-extras pytest tests/test_missing_output.py -q`. Expected: missing-repeat tests pass.
- [ ] **Step 5: Write failing null-repeat and layout tests**: `None` uses `W1302` only under tolerant modes; `[]` retains its warning-free placeholder; a non-null wrong type remains `E1303`; missing `if` remains `E1301`; recovery preserves rows, contained merges, presentation provenance, and zero displacement under `rows`, `cells`, and `none`. Test nested child recovery once per parent instance, suppression of descendant errors, independent fatal errors, and resource/collision limits.
- [ ] **Step 6: Run** `uv run --all-extras pytest tests/test_missing_output.py tests/test_render_plan.py -q`. Expected: new null-repeat and layout tests fail.
- [ ] **Step 7: Complete null-repeat warning and identity-block planning** using `render_unavailable_repeat(...)`. Count each recovered occurrence under repeat limits, reserve fixed footprints for `shift="none"`, and preserve existing parent allocation checks.
- [ ] **Step 8: Run** `uv run --all-extras pytest tests/test_missing_output.py tests/test_render_plan.py -q`. Expected: all pass.
- [ ] **Step 9: Commit** the renderer and tests with `feat: recover missing and null repeat collections`.

### Task 4: Prove workbook publication and presentation preservation

**Files:**
- Test: `tests/xlsx/test_workbook_renderer.py`
- Modify only if a failing integration test proves it necessary: `src/excel_template_writer/xlsx/writer.py` or `src/excel_template_writer/xlsx/validation.py`

**Interfaces:**
- Consumes: completed adapter-neutral `RenderPlan` from Task 3 and existing `render_workbook(..., missing_output=...)`.
- Produces: reopened XLSX evidence for strict non-publication and tolerant identity geometry; no new writer layout policy.

- [ ] **Step 1: Write a minimal workbook integration test** with a two-row unavailable repeat, static text, output/directive tags, a styled blank, a contained merge, custom row height and column width, and an unaffected neighbor. Under `blank` and `preserve`, save and reopen; assert values, types, direct fills/borders, number format, row height, column width, merge, unchanged neighbor, and warning code/location. Inspect OOXML only if effective blank/default styling cannot be proven through OpenPyXL.
- [ ] **Step 2: Add strict and independent-error cases** proving no output file is published or existing destination changed when `error` is active or a separate fatal error remains.
- [ ] **Step 3: Run** `uv run --all-extras pytest tests/xlsx/test_workbook_renderer.py -q`. Expected: all pass if the writer already consumes the new plans correctly; if not, fix only the proven adapter issue and rerun.
- [ ] **Step 4: Commit** integration tests and any necessary adapter fix with `test: verify recovered repeats in xlsx output`.

### Task 5: Maintain author docs and executable workbook samples

**Files:**
- Modify: `README.md`
- Modify: `docs/directives.md`
- Modify: `docs/explained.md`
- Modify: `samples/missing_output_policies.py`
- Modify: `samples/missing_output_policies_template.xlsx` and `samples/missing_output_policies_output.xlsx` via the Python builder only
- Modify: `samples/README.md`
- Modify: `tests/test_samples.py`

**Interfaces:**
- Consumes: final language contract and production workbook API.
- Produces: documented examples and a maintained visible-tag template/output pair demonstrating missing and null repeats, null collection filters, and existing scalar output recovery.

- [ ] **Step 1: Extend the sample builder** with a styled multi-cell missing repeat, a null repeat, and a null collection-filter output. Verify temporary strict and blank renders plus the committed preserve render after save/reload. Assert warning codes, source locations, exact directive/output text, styles, dimensions, merges, and neighboring cells.
- [ ] **Step 2: Update `tests/test_samples.py`** so the preserve-mode sample may contain expected `{% ... %}` text, while all other sample outputs still reject unresolved directives. Keep assertions for exact expected tags.
- [ ] **Step 3: Update README and both author documents** with the policy matrix and examples. Explain that output recovery is per tag, repeat recovery is per rectangle, and `if`/wrong types remain strict.
- [ ] **Step 4: Regenerate only the affected pair** with `uv run --all-extras python -m samples.missing_output_policies`, then run `uv run --all-extras pytest tests/test_samples.py tests/test_documentation.py -q`. Expected: all pass, and the generated template/output workbooks reopen with the asserted properties.
- [ ] **Step 5: Commit** docs, Python sample, tests, and generated workbook pair with `docs: demonstrate missing collection recovery`.

### Task 6: Run the full gate and review the branch

**Files:**
- No new files; fix only failures caused by this change.

**Interfaces:**
- Consumes: Tasks 1–5.
- Produces: a verified branch ready for review.

- [ ] **Step 1: Run** `uv run --all-extras pytest`. Expected: all tests pass.
- [ ] **Step 2: Run** `uv run --all-extras ruff check src tests samples scratch`. Expected: all checks pass.
- [ ] **Step 3: Run** `uv run --all-extras ruff format --check src tests samples scratch`. Expected: all files formatted.
- [ ] **Step 4: Run** `uv run --all-extras ty check`. Expected: all checks pass.
- [ ] **Step 5: Review** `git diff main...HEAD --check` and `git status --short`. Expected: no whitespace errors or unrelated changes; user-authored untracked files remain untouched.
