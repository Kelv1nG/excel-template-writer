from __future__ import annotations

import ast
import re
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
DOCUMENTED_ROOTS = (ROOT / "src", ROOT / "samples")


def _python_files() -> list[Path]:
    return sorted(path for root in DOCUMENTED_ROOTS for path in root.rglob("*.py"))


def _functions(path: Path) -> list[ast.FunctionDef | ast.AsyncFunctionDef]:
    return [
        node
        for node in ast.walk(ast.parse(path.read_text(encoding="utf-8")))
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef))
    ]


def _documented_parameters(node: ast.FunctionDef | ast.AsyncFunctionDef) -> set[str]:
    docstring = ast.get_docstring(node) or ""
    return {
        match.group("name")
        for match in re.finditer(r"^\s*(?P<name>[A-Za-z_][A-Za-z0-9_]*):", docstring, re.MULTILINE)
    }


@pytest.mark.parametrize("path", _python_files(), ids=lambda path: str(path.relative_to(ROOT)))
def test_maintained_python_functions_have_parameter_documentation(path: Path) -> None:
    failures: list[str] = []
    for node in _functions(path):
        docstring = ast.get_docstring(node)
        if docstring is None:
            failures.append(f"{node.name} at line {node.lineno} has no docstring")
            continue
        parameters = [
            argument.arg
            for argument in (*node.args.posonlyargs, *node.args.args, *node.args.kwonlyargs)
            if argument.arg not in {"self", "cls"}
        ]
        if node.args.vararg is not None:
            parameters.append(node.args.vararg.arg)
        if node.args.kwarg is not None:
            parameters.append(node.args.kwarg.arg)
        missing = sorted(set(parameters) - _documented_parameters(node))
        if missing:
            failures.append(
                f"{node.name} at line {node.lineno} does not document: {', '.join(missing)}"
            )
    assert not failures, "\n".join(failures)
