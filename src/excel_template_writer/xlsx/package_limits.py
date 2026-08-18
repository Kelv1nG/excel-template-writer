"""Read-only XLSX ZIP preflight for deterministic package resource limits."""

from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree
from zipfile import ZipFile

from excel_template_writer.diagnostics import Diagnostic, DiagnosticCode, SourceLocation
from excel_template_writer.limits import ResourceLimits


def _diagnostic(code: DiagnosticCode, message: str) -> Diagnostic:
    return Diagnostic(code, message, SourceLocation("<workbook>", "A1"))


def _worksheet_count(archive: ZipFile) -> int:
    try:
        workbook_xml = archive.open("xl/workbook.xml")
    except KeyError:
        return 0
    count = 0
    with workbook_xml:
        for _, element in ElementTree.iterparse(workbook_xml, events=("end",)):
            if element.tag.rpartition("}")[2] == "sheet":
                count += 1
            element.clear()
    return count


def inspect_xlsx_package(
    path: str | Path,
    limits: ResourceLimits,
    *,
    description: str,
) -> Diagnostic | None:
    """Return the first package-limit failure without extracting the archive."""

    package_path = Path(path)
    compressed_bytes = package_path.stat().st_size
    if compressed_bytes > limits.max_xlsx_file_bytes:
        return _diagnostic(
            DiagnosticCode.XLSX_PACKAGE_LIMIT_EXCEEDED,
            f"{description} exceeds max_xlsx_file_bytes={limits.max_xlsx_file_bytes:,}",
        )

    with ZipFile(package_path) as archive:
        members = archive.infolist()
        if len(members) > limits.max_xlsx_archive_members:
            return _diagnostic(
                DiagnosticCode.XLSX_PACKAGE_LIMIT_EXCEEDED,
                f"{description} exceeds "
                f"max_xlsx_archive_members={limits.max_xlsx_archive_members:,}",
            )
        uncompressed_bytes = sum(member.file_size for member in members)
        if uncompressed_bytes > limits.max_xlsx_uncompressed_bytes:
            return _diagnostic(
                DiagnosticCode.XLSX_PACKAGE_LIMIT_EXCEEDED,
                f"{description} exceeds "
                f"max_xlsx_uncompressed_bytes={limits.max_xlsx_uncompressed_bytes:,}",
            )
        worksheets = _worksheet_count(archive)
        if worksheets > limits.max_worksheets:
            return _diagnostic(
                DiagnosticCode.WORKSHEET_COUNT_LIMIT_EXCEEDED,
                f"{description} exceeds max_worksheets={limits.max_worksheets:,}",
            )
    return None


__all__ = ["inspect_xlsx_package"]
