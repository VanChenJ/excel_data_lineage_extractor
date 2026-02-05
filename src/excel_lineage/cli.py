from __future__ import annotations

import argparse
import json
from dataclasses import asdict

from .lineage import extract_lineage


def _format_markdown(items) -> str:
    lines = ["# Excel Lineage", ""]
    for item in items:
        lines.append(f"## {item.name}")
        lines.append(f"- Sheet: `{item.sheet}`")
        lines.append(f"- Target: `{item.target}`")
        if item.formula:
            lines.append(f"- Formula: `{item.formula}`")
        if item.references:
            lines.append("- References:")
            for ref in item.references:
                lines.append(f"  - `{ref.sheet}!{ref.ref}`")
        if item.metadata.top_headers or item.metadata.left_headers:
            lines.append("- Metadata:")
            if item.metadata.top_headers:
                lines.append(f"  - Top headers: {', '.join(item.metadata.top_headers)}")
            if item.metadata.left_headers:
                lines.append(f"  - Left headers: {', '.join(item.metadata.left_headers)}")
        lines.append("")
    return "\n".join(lines).strip() + "\n"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Extract Excel lineage to stdout.")
    parser.add_argument("path", help="Path to the Excel file (.xlsx)")
    parser.add_argument(
        "--format",
        choices=("json", "markdown"),
        default="json",
        help="Output format for stdout.",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    lineages = extract_lineage(args.path)
    if args.format == "markdown":
        print(_format_markdown(lineages))
        return
    payload = [asdict(item) for item in lineages]
    print(json.dumps(payload, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
