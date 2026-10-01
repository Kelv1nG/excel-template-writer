# Numeric Fidelity Implementation Package

This directory is a self-contained handoff for the Excel Template Writer.

## Start here

- [Implementation handoff](handoff/excel-template-writer/README.md)
- [Numeric-fidelity background](NUMERIC_FIDELITY.md)

The implementation handoff contains the selected OpenPyXL compatibility-shim design, known risks,
contract decisions that must be resolved in `SPEC.md`, test coverage, acceptance criteria, and
primary references. The background note explains the Excel, XLSX XML, and binary-floating-point
mental model in more detail.

## Ready-to-paste Codex request

```text
Read AGENTS.md and docs/handoffs/numeric-fidelity/README.md, then execute the numeric-fidelity
implementation handoff. Follow the repository's XLSX integration skill and preserve unrelated
worktree changes. Update SPEC.md before changing public semantics. Do not guess the explicitly
listed float-range, integer-range, or Decimal policy decisions; surface those decisions to me
first. After they are resolved, implement the guarded OpenPyXL compatibility shim and its complete
regression coverage, then run the full quality gate from docs/DEVELOPMENT.md.
```
