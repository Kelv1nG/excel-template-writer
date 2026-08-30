"""Regenerate and verify every maintained sample workbook."""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from samples import (
    cell_shift_lanes,
    conditions_and_nesting,
    fixed_range_charts,
    regions,
    repeated_blocks,
    scalar_values,
    template_images,
    template_text_shapes,
)


def _run_sample(
    build: Callable[[], Path],
    render: Callable[[Path], Path],
) -> tuple[Path, Path]:
    """Build and render one sample pair.

    Args:
        build: Zero-argument template builder.
        render: Renderer accepting the generated template path.

    Returns:
        The generated ``(template_path, output_path)`` pair.
    """

    template = build()
    return template, render(template)


def main() -> None:
    """Regenerate all core samples and the optional Polars sample when installed."""

    pairs = [
        _run_sample(scalar_values.build_template, scalar_values.render_sample),
        _run_sample(repeated_blocks.build_template, repeated_blocks.render_sample),
        _run_sample(
            conditions_and_nesting.build_template,
            conditions_and_nesting.render_sample,
        ),
        _run_sample(cell_shift_lanes.build_template, cell_shift_lanes.render_sample),
        _run_sample(fixed_range_charts.build_template, fixed_range_charts.render_sample),
        _run_sample(template_images.build_template, template_images.render_sample),
        _run_sample(
            template_text_shapes.build_template,
            template_text_shapes.render_sample,
        ),
        _run_sample(regions.build_template, regions.render_sample),
    ]
    try:
        from samples import polars_dataframe
    except ModuleNotFoundError as error:
        if error.name != "polars":
            raise
        print("Polars is not installed; skipped the optional polars_dataframe sample.")
    else:
        pairs.append(_run_sample(polars_dataframe.build_template, polars_dataframe.render_sample))

    for template, output in pairs:
        print(f"Generated {template.name} and {output.name}")


if __name__ == "__main__":
    main()
