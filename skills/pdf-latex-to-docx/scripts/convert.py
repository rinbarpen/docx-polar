#!/usr/bin/env python3
"""Convert PDF or LaTeX files to DOCX preserving content and formatting.

Routes:
  .pdf  → pdf2docx (primary) or PyMuPDF text extraction (fallback)
  .tex  → pandoc (primary) or compile→PDF→pdf2docx (fallback)
"""

import argparse
import logging
import os
import shutil
import subprocess
import sys
from pathlib import Path

logging.basicConfig(format="%(levelname)s: %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)


def _run(cmd: list[str], verbose: bool = False, **kwargs):
    if verbose:
        cmd_str = " ".join(cmd)
        dim = "\033[2m" if sys.stderr.isatty() else ""
        reset = "\033[0m" if sys.stderr.isatty() else ""
        log.info("  %s%s%s", dim, cmd_str, reset)
    return subprocess.run(cmd, **kwargs)


def _find_latex_compiler(compiler: str) -> str:
    order = [compiler, "pdflatex", "xelatex", "lualatex"]
    for name in order:
        path = shutil.which(name)
        if path:
            return path
    raise SystemExit("No LaTeX compiler found (tried: pdflatex, xelatex, lualatex)")


# ---------------------------------------------------------------------------
# PDF → DOCX
# ---------------------------------------------------------------------------


def convert_pdf_pdf2docx(input_path: Path, output_path: Path, verbose: bool = False) -> dict:
    """Convert PDF to DOCX using pdf2docx (page-level layout preservation)."""
    from pdf2docx import Converter

    if verbose:
        log.info("Converting %s → %s via pdf2docx", input_path, output_path)

    cv = Converter(str(input_path))
    try:
        cv.convert(str(output_path), start=0, end=None)
    finally:
        cv.close()

    return _stats_pdf(input_path)


def convert_pdf_fitz_fallback(input_path: Path, output_path: Path, verbose: bool = False) -> dict:
    """Extract text from PDF via PyMuPDF and write to a basic DOCX.

    Used as a fallback when pdf2docx fails. Layout and images are not preserved.
    """
    import fitz
    from docx import Document

    if verbose:
        log.info("Extracting text from %s via PyMuPDF", input_path)

    out = Document()
    with fitz.open(str(input_path)) as doc:
        num_pages = len(doc)
        for page_num, page in enumerate(doc):
            text = page.get_text("text")
            if not text.strip():
                continue
            if page_num > 0:
                out.add_page_break()
            for block in text.splitlines():
                block = block.strip()
                if block:
                    out.add_paragraph(block)

    out.save(str(output_path))
    return {"pages": num_pages, "method": "fitz text extraction"}


# ---------------------------------------------------------------------------
# LaTeX → DOCX
# ---------------------------------------------------------------------------


def convert_tex_pandoc(input_path: Path, output_path: Path, verbose: bool = False) -> dict:
    """Convert LaTeX to DOCX using pandoc."""
    if verbose:
        log.info("Converting %s → %s via pandoc", input_path, output_path)

    result = _run(
        ["pandoc", str(input_path), "-o", str(output_path), "--from=latex"],
        capture_output=not verbose,
        verbose=verbose,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"pandoc conversion failed:\n{result.stderr.decode(errors='replace')}"
        )
    return {"converter": "pandoc"}


def _compile_latex_to_pdf(tex_path: Path, compiler: str, verbose: bool = False) -> Path:
    """Run LaTeX compiler twice for cross-references; return path to the produced PDF."""
    tex_dir = str(tex_path.parent.resolve()) or "."
    tex_name = tex_path.name

    compiler_path = _find_latex_compiler(compiler)
    comp_name = os.path.basename(compiler_path)

    for _ in range(2):
        result = _run(
            [
                compiler_path,
                "-interaction=nonstopmode",
                "-output-directory",
                tex_dir,
                tex_name,
            ],
            cwd=tex_dir,
            capture_output=not verbose,
            verbose=verbose,
        )
        if result.returncode != 0:
            stderr_text = result.stderr.decode(errors="replace") if result.stderr else ""
            log.warning(
                "LaTeX compiler %s returned non-zero (attempt %d): %s",
                comp_name,
                _ + 1,
                stderr_text[:500],
            )

    pdf_path = Path(tex_dir) / tex_path.with_suffix(".pdf").name
    if not pdf_path.exists():
        raise FileNotFoundError(
            f"LaTeX compilation did not produce {pdf_path}. Check the .log file for errors."
        )
    return pdf_path


def convert_tex_compile_pdf(
    input_path: Path, output_path: Path, latex_compiler: str, verbose: bool = False
) -> dict:
    """Compile LaTeX to PDF, then convert PDF to DOCX via pdf2docx."""
    if verbose:
        log.info("Compiling %s → PDF via %s", input_path, latex_compiler)

    pdf_path = _compile_latex_to_pdf(input_path, latex_compiler, verbose=verbose)
    return convert_pdf_pdf2docx(pdf_path, output_path, verbose=verbose)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _stats_pdf(pdf_path: Path) -> dict:
    import fitz

    with fitz.open(str(pdf_path)) as doc:
        return {"pages": len(doc)}


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


class ConversionError(Exception):
    """Expected conversion failure (triggers fallback)."""


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert PDF or LaTeX files to DOCX preserving content and formatting."
    )
    parser.add_argument("input", help="Input file path (.pdf or .tex)")
    parser.add_argument(
        "--output", "-o", default=None, help="Output DOCX path (default: <input_stem>.docx)"
    )
    parser.add_argument(
        "--method",
        choices=["auto", "pdf2docx", "pandoc", "compile-pdf"],
        default="auto",
        help="Conversion method (default: auto-detect)",
    )
    parser.add_argument(
        "--latex-compiler",
        choices=["pdflatex", "xelatex", "lualatex"],
        default="pdflatex",
        help="LaTeX compiler for compile-pdf method",
    )
    parser.add_argument("--verbose", "-v", action="store_true", help="Print progress output")
    args = parser.parse_args()

    if args.verbose:
        log.setLevel(logging.DEBUG)

    input_path = Path(args.input).resolve()
    if not input_path.exists():
        raise SystemExit(f"Input file not found: {args.input}")

    suffix = input_path.suffix.lower()
    if suffix not in (".pdf", ".tex"):
        raise SystemExit(f"Unsupported input type: {suffix}. Expected .pdf or .tex")

    output_path = Path(args.output).resolve() if args.output else input_path.with_suffix(".docx")

    method = args.method
    if method == "auto":
        method = "pdf2docx" if suffix == ".pdf" else "pandoc"

    if suffix == ".pdf" and method not in ("pdf2docx", "auto"):
        log.warning("PDF input always uses pdf2docx; ignoring --method=%s", method)
        method = "pdf2docx"

    if suffix == ".tex" and method not in ("pandoc", "compile-pdf", "auto"):
        log.warning("LaTeX input with unsupported method %s; falling back to pandoc", method)
        method = "pandoc"

    print(f"  Input:  {input_path}")
    print(f"  Output: {output_path}")
    print(f"  Method: {method}")

    stats: dict = {}
    try:
        if suffix == ".pdf":
            stats = convert_pdf_pdf2docx(input_path, output_path, verbose=args.verbose)
        elif method == "pandoc":
            stats = convert_tex_pandoc(input_path, output_path, verbose=args.verbose)
        else:
            stats = convert_tex_compile_pdf(
                input_path, output_path, args.latex_compiler, verbose=args.verbose
            )

        print(f"  Done:   {output_path}")
        if "pages" in stats:
            print(f"  Pages:  {stats['pages']}")

    except ConversionError as exc:
        # Already a known failure; no fallback
        raise SystemExit(str(exc))

    except (ImportError, ModuleNotFoundError) as exc:
        raise SystemExit(
            f"Missing dependency: {exc}\n"
            "Install with: pip install pdf2docx python-docx PyMuPDF"
        )

    except (RuntimeError, FileNotFoundError, OSError, SystemExit) as exc:
        print(f"  Error: {exc}", file=sys.stderr)

        if method == "pdf2docx":
            print("  Falling back to PyMuPDF text extraction...", file=sys.stderr)
            try:
                stats = convert_pdf_fitz_fallback(input_path, output_path, verbose=args.verbose)
                print(f"  Done:   {output_path} (text-only)", file=sys.stderr)
                if "pages" in stats:
                    print(f"  Pages:  {stats['pages']}", file=sys.stderr)
            except Exception as exc2:
                raise SystemExit(f"Both methods failed.\n  pdf2docx: {exc}\n  fitz: {exc2}")
        elif method == "pandoc":
            print("  Falling back to compile→PDF→DOCX method...", file=sys.stderr)
            try:
                stats = convert_tex_compile_pdf(
                    input_path, output_path, args.latex_compiler, verbose=args.verbose
                )
                print(f"  Done:   {output_path}", file=sys.stderr)
            except Exception as exc2:
                raise SystemExit(f"Both methods failed.\n  pandoc: {exc}\n  compile-pdf: {exc2}")
        elif method == "compile-pdf":
            print("  Falling back to pandoc method...", file=sys.stderr)
            try:
                stats = convert_tex_pandoc(input_path, output_path, verbose=args.verbose)
                print(f"  Done:   {output_path}", file=sys.stderr)
            except Exception as exc2:
                raise SystemExit(f"Both methods failed.\n  compile-pdf: {exc}\n  pandoc: {exc2}")

    except Exception:
        log.exception("Unexpected error during conversion")
        raise SystemExit(
            "An unexpected error occurred during conversion. "
            "Re-run with --verbose for details."
        )


if __name__ == "__main__":
    main()
