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
- Inspect OOXML parts only when the public workbook model cannot prove the behavior.

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
