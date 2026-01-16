#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Mapping loader for PPT Automation Demo.

This module parses the markdown mapping specification in
`docs/data_mapping.md` and exposes it as Python-friendly structures.

Design goal: the markdown table serves as both human documentation and
machine-readable configuration (documentation-as-code pattern).
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional


@dataclass
class ElementMapping:
    slide_index: int
    element_type: str  # text | chart | table
    logical_name: str
    shape_name: str
    data_key: str
    number_format: Optional[str]
    title_keyword: Optional[str]
    description: str


def _parse_mapping_table(lines: List[str]) -> List[ElementMapping]:
    """Parse the first markdown table in the mapping document."""
    header_idx: Optional[int] = None

    for idx, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("|") and "slide_index" in stripped and "element_type" in stripped:
            header_idx = idx
            break

    if header_idx is None:
        raise ValueError("Cannot find mapping table header in data_mapping.md")

    # Skip header line and separator line
    data_start = header_idx + 2

    mappings: List[ElementMapping] = []

    for line in lines[data_start:]:
        stripped = line.strip()
        if not stripped:
            # Stop at first empty line after table
            break
        if not stripped.startswith("|"):
            # End of table section
            break

        # Split markdown row: | a | b | c |
        parts = [cell.strip() for cell in stripped.split("|")[1:-1]]
        if len(parts) < 8:
            # Skip malformed rows
            continue

        slide_index_str, element_type, logical_name, shape_name, data_key, number_format, title_keyword, description = parts[:8]

        # Basic normalization
        try:
            slide_index = int(slide_index_str)
        except ValueError:
            # Skip header-like or invalid rows
            continue

        number_format = number_format or None
        title_keyword = title_keyword or None

        mappings.append(
            ElementMapping(
                slide_index=slide_index,
                element_type=element_type,
                logical_name=logical_name,
                shape_name=shape_name,
                data_key=data_key,
                number_format=number_format,
                title_keyword=title_keyword,
                description=description,
            )
        )

    if not mappings:
        raise ValueError("No valid mapping rows found in data_mapping.md")

    return mappings


def load_mapping(path: str = "docs/data_mapping.md") -> Dict[str, object]:
    """
    Load mapping configuration from markdown file.

    Returns a dictionary with:
      - 'elements': List[ElementMapping]
      - 'shape_names': Dict[logical_name, shape_name]
      - 'charts': List[ElementMapping]   (element_type == 'chart')
      - 'texts': List[ElementMapping]    (element_type == 'text')
      - 'tables': List[ElementMapping]   (element_type == 'table')
    """
    with open(path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    elements = _parse_mapping_table(lines)

    shape_names: Dict[str, str] = {
        e.logical_name: e.shape_name for e in elements if e.shape_name
    }

    charts = [e for e in elements if e.element_type == "chart"]
    texts = [e for e in elements if e.element_type == "text"]
    tables = [e for e in elements if e.element_type == "table"]

    return {
        "elements": elements,
        "shape_names": shape_names,
        "charts": charts,
        "texts": texts,
        "tables": tables,
    }


if __name__ == "__main__":
    # Simple manual test / debug helper
    cfg = load_mapping()
    print(f"Loaded {len(cfg['elements'])} mapping rows")
    print("Shape names:")
    for k, v in cfg["shape_names"].items():
        print(f"  {k}: {v}")

