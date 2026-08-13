import ast
from pathlib import Path


def test_phase_zero_core_does_not_import_openpyxl() -> None:
    package = Path(__file__).parents[1] / "src" / "excel_template_writer"
    forbidden = []

    for path in package.glob("*.py"):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                forbidden.extend(
                    (path.name, alias.name) for alias in node.names if alias.name == "openpyxl"
                )
            elif isinstance(node, ast.ImportFrom) and node.module == "openpyxl":
                forbidden.append((path.name, node.module))

    assert forbidden == []
