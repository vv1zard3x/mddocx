"""docxmd — convert a Markdown folder into a GOST-styled .docx document.

Usage::

    docxmd path/to/source_dir [-o out.docx]

Where ``source_dir`` contains exactly one ``*.md`` file and optionally
a ``media/`` directory referenced by image links. See
``MARKDOWN_RULES.md`` for the markdown contract the converter follows.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from ._convert import ConversionError, convert

__all__ = ["convert", "ConversionError", "main"]
__version__ = "0.1.0"


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="docxmd",
        description=(
            "Convert a folder with one Markdown file (and optional media/) "
            "into a GOST-formatted .docx document."
        ),
    )
    parser.add_argument(
        "source",
        type=Path,
        help="Directory containing exactly one .md file and optional media/.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="Output .docx path. Defaults to <md filename>.docx next to source.",
    )
    args = parser.parse_args(argv)

    try:
        out_path = convert(args.source, args.output)
    except ConversionError as exc:
        print(f"docxmd: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # pragma: no cover - last-resort safety net
        print(f"docxmd: unexpected error: {exc}", file=sys.stderr)
        return 1

    print(out_path)
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
