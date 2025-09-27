# md_to_docx/__main__.py
from __future__ import annotations
import sys
import argparse
from pathlib import Path

from . import MarkdownToDocx
from .errors import (
    MdToDocxError,
    InputNotFoundError,
    TemplateError,
    ImageDownloadError,
    NumberingError,
    MathConversionError,
)


def _build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="python -m md_to_docx",
        description="Convert Markdown (.md) files to DOCX using an optional Word template."
    )
    p.add_argument("input", help="Path to the input Markdown file (.md)")
    p.add_argument("-o", "--output",
                   help="Output file (.docx). Defaults to input filename with .docx extension")
    p.add_argument("-t", "--template",
                   help="Optional Word template (.docx) to reuse styles and numbering")
    p.add_argument("--apply-base-font", action="store_true",
                   help="Force base font name/size on all styles (use with care on corporate templates)")
    p.add_argument("--font-name", default="Calibri",
                   help="Base font name (only relevant with --apply-base-font)")
    p.add_argument("--font-size", type=int, default=11,
                   help="Base font size in points (only relevant with --apply-base-font)")
    return p


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    try:
        in_path = Path(args.input)
        if not in_path.exists():
            raise InputNotFoundError(f"Input file not found: {in_path}")
        if in_path.suffix.lower() != ".md":
            print("WARNING: Input file does not have a .md extension. Attempting anyway.", file=sys.stderr)

        out_path = Path(args.output) if args.output else in_path.with_suffix(".docx")
        tpl_path = args.template

        md_text = in_path.read_text(encoding="utf-8")

        conv = MarkdownToDocx(
            template_path=tpl_path,
            apply_base_font=args.apply_base_font,
            base_font=args.font_name,
            base_size_pt=args.font_size,
        )

        conv.convert_text(md_text, str(out_path))
        print(f"Done: {out_path}")
        return 0

    except InputNotFoundError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 2
    except TemplateError as e:
        print(f"TEMPLATE ERROR: {e}", file=sys.stderr)
        return 3
    except ImageDownloadError as e:
        print(f"IMAGE ERROR: {e}", file=sys.stderr)
        return 4
    except NumberingError as e:
        print(f"LIST ERROR: {e}", file=sys.stderr)
        return 5
    except MathConversionError as e:
        print(f"MATH ERROR: {e}", file=sys.stderr)
        return 6
    except MdToDocxError as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 7
    except Exception as e:
        print(f"UNEXPECTED: {e}", file=sys.stderr)
        return 10


if __name__ == "__main__":
    raise SystemExit(main())
