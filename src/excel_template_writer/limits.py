"""Shared deterministic resource limits for normalization, planning, and XLSX I/O."""

from __future__ import annotations

from dataclasses import dataclass, fields

XLSX_MAX_ROWS = 1_048_576
XLSX_MAX_COLUMNS = 16_384
XLSX_MAX_CELL_TEXT_LENGTH = 32_767


@dataclass(frozen=True, slots=True)
class ResourceLimits:
    """Permissive safety ceilings for one complete workbook render operation."""

    max_context_depth: int = 64
    max_context_nodes: int = 1_000_000
    max_container_items: int = 100_000
    max_input_string_length: int = 1_000_000
    max_repeat_iterations_per_sheet: int = 100_000
    max_planned_cells_per_sheet: int = 500_000
    max_planned_cells_per_workbook: int = 1_000_000
    max_output_rows_per_sheet: int = 250_000
    max_output_columns_per_sheet: int = 4_096
    max_worksheets: int = 100
    max_xlsx_file_bytes: int = 50 * 1024 * 1024
    max_xlsx_uncompressed_bytes: int = 250 * 1024 * 1024
    max_xlsx_archive_members: int = 10_000

    def __post_init__(self) -> None:
        """Reject non-integer or non-positive resource ceilings.

        Raises:
            ValueError: If any configured ceiling is not a positive integer.
        """

        for field in fields(self):
            value = getattr(self, field.name)
            if type(value) is not int or value <= 0:
                raise ValueError(f"{field.name} must be a positive integer")


DEFAULT_RESOURCE_LIMITS = ResourceLimits()


__all__ = [
    "DEFAULT_RESOURCE_LIMITS",
    "XLSX_MAX_CELL_TEXT_LENGTH",
    "XLSX_MAX_COLUMNS",
    "XLSX_MAX_ROWS",
    "ResourceLimits",
]
