# Development

## Status

The repository has an executable language model and a production-oriented `.xlsx` adapter. The
adapter snapshots supported workbook presentation, invokes the pure compiler and layout planner,
validates coordinate-dependent workbook features, writes a separate workbook atomically, and
reopens the result for package-integrity verification.

## Runtime and package management

- Target CPython 3.12.
- Use `uv` for the environment, dependency management, lockfile, and project commands.
- `.python-version` pins the requested interpreter line for local tooling.

Manage packages with:

```powershell
uv python install 3.12
uv sync
uv sync --all-groups --all-extras
uv add <package>
uv add --dev <development-package>
uv add --optional <extra> <package>
uv remove <package>
uv run <command>
```

Commit both `pyproject.toml` and `uv.lock`. Do not use direct `pip install` commands to change the project environment or maintain a parallel dependency list.

The selected quality tools are:

- `pytest` for tests;
- Ruff for linting and formatting;
- `ty` for static type checking.

Run the complete local gate with:

```powershell
uv run --all-extras pytest
uv run --all-extras ruff check src tests samples scratch
uv run --all-extras ruff format --check src tests samples scratch
uv run --all-extras ty check
```

The base environment may omit optional integrations; their test modules skip when the corresponding
extra is unavailable. The complete gate above installs and tests every supported integration.

## Before changing the project

1. Read `SPEC.md` for the relevant contract.
2. Read `AGENTS.md` for repository-wide constraints.
3. Use the language-semantics skill for changes that affect template meaning or layout.
4. Use the XLSX-integration skill for changes that affect actual workbook files.
5. Use both when a feature spans the interpreter and workbook adapter.

## Package boundaries

Preserve these responsibilities:

- immutable canonical render-context normalization, resource limits, and caller-supplied platform
  adapters;
- workbook reader and immutable workbook model;
- cell lexer and expression/directive parser;
- spatial marker linker and semantic validator;
- AST and safe evaluator;
- pure layout planner and render-plan IR;
- `openpyxl` workbook writer;
- structured diagnostics.

The interpreter and layout layers must be testable without opening or saving an `.xlsx` file.

## Testing layers

- Language unit tests: normalization/adapters, canonical values, tokens, grammar, expressions,
  scopes, AST, and diagnostics.
- Spatial tests: rectangle pairing, containment, ambiguity, nesting, measurement, shifting, and collisions.
- Workbook integration tests: typed cells, styles, dimensions, merged ranges, supported ordered
  chart/image/text-shape drawings, embedded-media identity, and save/reload integrity.
- Resource-limit tests: fail-fast context paths, pure-plan boundaries, package preflight, and
  unpublished oversized output.
- End-to-end fixtures: only for representative user-visible behavior spanning all layers.

Prefer small semantic assertions over whole-workbook binary comparisons. Each invalid case should
assert its stable diagnostic code and either its source location or canonical context path.

## XLSX fixtures

- Give each fixture one narrow purpose.
- Keep source templates separate from generated outputs.
- Generate programmatic fixtures through a documented helper once a test package exists.
- Do not manually patch binary workbook contents.
- Reopen rendered workbooks and verify values, types, styles, dimensions, merges, supported chart
  properties, image and text-shape anchors, static shape content, drawing order, and embedded-media
  bytes.
- For any change that removes directive text, repeats cells, or shifts layout, assert direct fill
  and border properties after save/reload at representative rendered, directive-only, formatted
  blank, shifted, and unaffected cells. Also verify the effective workbook default style for
  absent blank cells, including row/column overrides, and retain explicit no-fill overrides
  alongside directly filled controls. When Normal is customized, inspect both `cellXfs[0]` and
  the built-in Normal entry in `cellStyleXfs`; checking only one missed a visible gridline
  regression. Value-only assertions do not prove presentation preservation and can miss cells
  that silently revert to Excel's default appearance. Inspect saved style definitions and cell
  records in OOXML when synthesized blank cells in the public workbook model cannot prove the
  effective appearance.
- Inspect OOXML parts only when the public workbook model cannot prove the behavior.

Numeric fidelity tests distinguish stored tokens from reader interpretation. Every finite built-in
float token preserves binary64 bits, and every accepted Decimal token preserves its decimal amount.
Reopen-as-float type/bit assertions require an effective authored format that pinned OpenPyXL does
not interpret as a date, time, or duration. For recognized temporal formats, assert numeric XML
and preserved formats independently from the returned temporal Python value or `#VALUE!` for an
out-of-range serial; characterize existing behavior without adding a temporal fidelity guarantee.
The pinned-library contract does not guarantee desktop Excel resave fidelity.

### Frozen desktop-Excel numeric reference

`tests/fixtures/numeric_fidelity/excel_numeric_reference.xlsx` is an immutable reference created,
fully recalculated, and saved by desktop Excel. Its adjacent JSON manifest records the Excel
version/build, locale, UTC creation time, calculation settings (including disabled Precision as
displayed), authoritative named cells, comparison modes, and SHA-256 checksum. See the
[fixture maintenance instructions](../tests/fixtures/numeric_fidelity/README.md) and documented
PowerShell COM builder; review the regenerated workbook and manifest together.

Portable tests verify provenance/checksum, formula text with `data_only=False`, cached numeric
results with `data_only=True`, and selected OOXML formula/cache nodes. Never resave the reference
through OpenPyXL: it does not calculate formulas and saving can discard Excel-authored caches.
Do not hand-edit its ZIP/XML. Regenerate it only with the documented desktop-Excel builder, outside
default CI. When updating maintained samples, regenerate only affected sample modules; leave the
frozen reference unchanged unless its maintenance is explicitly in scope.

The reference covers 1–15-digit Decimal amounts and four single operations on identical exact
binary-fraction inputs. It is evidence from the recorded Excel build, not a universal formula
oracle or a promise of arbitrary Polars/Excel arithmetic parity or Excel resave fidelity. Default
CI requires no live Excel or `pywin32`; optional live checks use a temporary copy, run serially,
and record the Excel build.

### DuckDB group-by comparison protocol

DuckDB is not a renderer dependency and the project does not currently bundle a DuckDB adapter.
When testing a caller pipeline of Polars ingestion → DuckDB transformation → XLSX output, keep the
grouping and business rounding outside the renderer and materialize the final ordered records before
calling the production workbook API. A DuckDB result returned as an eager Polars frame can then use
the existing Polars adapter.

A focused comparison test should prove these boundaries separately:

1. the source Polars schema and scalar values;
2. the explicit DuckDB input and output types, deterministic `ORDER BY`, and exact transformed
   results;
3. the eager Polars schema, row order, and scalar values returned by DuckDB `.pl()`;
4. the canonical values produced by `polars_adapters()`;
5. selected numeric worksheet XML tokens; and
6. values, types, and number formats after reopening with pinned OpenPyXL.

For fixed-scale financial group-bys, prefer explicit `DECIMAL(p, s)` casts and a documented rounding
scale. Include keys with duplicates, positive and negative rows, cancellation, zero, null-only
groups, rounding boundaries, every final precision from 1 through 15, and an unsupported final value
that proves the writer fails at the destination cell. Define null behavior explicitly: SQL `SUM` of
only nulls is null, whereas an Excel `SUMIFS` comparison commonly yields zero for blank inputs. A
final `ORDER BY` only stabilizes result-row order; an exact `DOUBLE` reduction probe also needs a
stable source ordinal and aggregate input ordering because DuckDB floating-point `SUM` is
order-sensitive.

Compare a DuckDB aggregate with desktop Excel only through a frozen, fully recalculated workbook
whose source rows, group keys, formula, rounding rule, build, and calculation settings are recorded.
For example, pair an explicitly cast and rounded DuckDB `SUM` with an Excel `ROUND(SUMIFS(...), s)`
result at the same reporting scale. DuckDB `DECIMAL` and Excel binary64 formulas have different
arithmetic semantics, so this is a bounded financial-scenario oracle, not a raw-bit or generic
DuckDB/Polars/Excel equivalence claim. See
[`NUMERIC_FIDELITY.md`](handoffs/numeric-fidelity/NUMERIC_FIDELITY.md#polars-ingestion-through-duckdb-to-excel)
for the complete boundary matrix.

DuckDB is not currently a runtime, optional-extra, or test dependency of this repository. These are
the required assertions for an application integration or a future pinned test; do not describe a
DuckDB path as covered by the default suite until an executable test and its DuckDB version are added
to the project lock.

If pytest's default temporary directory is inaccessible, pass a fresh workspace path such as
`--basetemp .venv/pytest-numeric-20261002` to each pytest invocation.

## Maintained user samples

`samples/` is the maintained, executable catalog of supported user-visible behavior. `scratch/`
remains useful for exploratory or disposable demonstrations, but it does not satisfy sample
coverage for a completed feature.

Every user-visible language, layout, value-adapter, formatting, or XLSX feature change must add or
update:

- executable Python under `samples/` that builds and renders through the production API;
- a visible-tag `*_template.xlsx` workbook;
- the corresponding `*_output.xlsx` workbook;
- save/reload assertions for the behavior being demonstrated; and
- the catalog in `samples/README.md`.

Regenerate the complete catalog with:

```powershell
uv run --all-extras python -m samples.generate_all
```

Samples depend on the current specification, implementation, and tests. They are explanatory
artifacts, not a competing source of language semantics.

## Definition of done

A change is complete when:

- behavior agrees with `SPEC.md`;
- architecture boundaries remain intact;
- valid, invalid, empty, boundary, and nesting cases are covered where relevant;
- workbook changes pass save/reload checks;
- user-visible features have matching maintained samples;
- supported checks run through `uv run`;
- limitations and unsupported behavior fail explicitly or are documented.
