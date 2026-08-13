---
name: language-semantics-change
description: Design, implement, or review changes to the Excel template language, including delimiters, expressions, directives, scopes, AST nodes, rectangular block geometry, conditions, repeats, shifting, layout behavior, validation, and diagnostics. Use whenever a change can alter what a template means or how its rendered cells are positioned.
---

# Language Semantics Change

Evolve the language as a spatial interpreter. Keep syntax, semantic validation, layout, and diagnostics aligned with the published contract.

## Workflow

1. Read the relevant sections of `SPEC.md` before proposing code.
2. Describe the change across all affected surfaces:
   - author-facing syntax and examples;
   - lexer/parser representation;
   - AST node fields and lexical scope;
   - rectangular geometry and nesting constraints;
   - evaluation and measured output size;
   - shift or isolation behavior;
   - validation rules and stable diagnostics.
3. Resolve unclear semantics in `SPEC.md` before implementation. Do not let tests or current code become an accidental language specification.
4. Implement through the proper pipeline: tokens, parsed syntax, spatial linking, semantic validation, AST/evaluation, layout IR, then adapter consumption.
5. Keep the interpreter and layout core independent of `openpyxl`.
6. Add focused tests before broad integration fixtures:
   - ordinary valid use;
   - boundary and empty cases;
   - invalid syntax and invalid geometry;
   - ambiguous pairing or partial overlap;
   - nesting and sibling interaction;
   - stable diagnostic code and source location.
7. Run the applicable checks through `uv run` and report any checks that do not yet exist.

## Compatibility rules

- Treat changes to existing syntax, output types, empty behavior, movement, or errors as compatibility changes.
- Prefer an explicit validation error to inferred or silently changed behavior.
- Preserve input collection order; sorting and data reconciliation belong to preprocessing.
- Never infer block boundaries from blank cells, styles, or worksheet used ranges.
- Never implement directives as ad hoc cell-string replacement.
- Keep author syntax declarative and usable by technically capable Excel users.

## Definition of done

- `SPEC.md` and implementation agree.
- The AST and layout effects are explicit.
- Invalid and ambiguous cases have stable diagnostics.
- Unit tests cover semantics without requiring workbook serialization where possible.
- Relevant integration tests demonstrate the final cell geometry.
