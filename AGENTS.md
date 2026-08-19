# Repository Instructions

## Sources of truth

- `SPEC.md` defines the template language and rendering semantics.
- `docs/DEVELOPMENT.md` defines the contributor workflow.
- `pyproject.toml` will define package metadata, dependencies, and tool configuration once application scaffolding begins.
- Do not introduce behavior that contradicts `SPEC.md`. Update and review the specification before implementing a semantic change.

## Environment

- Target CPython 3.12 only.
- Use `uv` for Python installation, virtual environments, dependency management, locking, and command execution.
- Use `uv add` and `uv remove` for dependency changes once `pyproject.toml` exists.
- Run project tools with `uv run`.
- Commit `uv.lock` once dependencies are introduced.
- Do not add Poetry, Pipenv, Conda, or a hand-maintained `requirements.txt` as a competing source of dependency truth.

## Architecture boundaries

- Preserve the pipeline: workbook reader/model → lexer/parser → spatial linker → semantic validator → AST/evaluator → layout IR/planner → workbook writer.
- Keep the interpreter, AST, validator, and layout planner independent of `openpyxl`.
- Parse, validate, and plan the complete render before mutating a workbook.
- Never infer rectangular regions from blank cells, styles, or worksheet used ranges.
- Never implement template behavior with ad hoc string replacement.
- Fail explicitly for ambiguity, collisions, and unsupported workbook behavior.

## Repository skills

- For syntax, directives, AST, expression, scope, region, condition, repeat, shift, layout, or diagnostic changes, follow `.agents/skills/language-semantics-change/SKILL.md`.
- For `openpyxl`, workbook loading/writing, cell types, styles, dimensions, merged cells, formula boundaries, or `.xlsx` fixture changes, follow `.agents/skills/xlsx-integration-change/SKILL.md`.
- Follow both skills when a change crosses the language and workbook boundary.

## Change discipline

- Keep input preprocessing separate from rendering. Sorting, grouping, joining, keyed merging, and reconciliation are not renderer responsibilities.
- Preserve user-authored changes and keep unrelated edits out of a change.
- Add the smallest tests that prove the contract at the lowest applicable layer.
- Every user-visible language, layout, adapter, formatting, or XLSX feature change must add or
  update executable Python and matching template/output workbooks under `samples/`. Samples follow
  the current implementation; they do not override `SPEC.md`.
- Use stable diagnostic codes and include worksheet/cell source locations.
- Do not edit binary `.xlsx` fixtures as ordinary source files.
- Before handing off, run available focused checks and then the full configured suite. State clearly when tooling or tests have not been configured yet.
