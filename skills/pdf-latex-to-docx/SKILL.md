---
name: pdf-latex-to-docx
description: "Convert PDF or LaTeX (.tex) files to DOCX while preserving content, structure, and formatting. Supports figures, tables, equations, citations, and page layout. Trigger: 'convert pdf to docx', 'latex to docx', 'tex to word', 'pdf to word', 'convert this pdf', 'turn this latex into docx'."
---

# PDF/LaTeX to DOCX

Convert PDF or LaTeX documents to editable DOCX format with content and formatting preserved.

## Conversion Paths

| Input | Primary Method | Fallback | Best For |
|-------|---------------|----------|----------|
| `.pdf` | `pdf2docx` (page-level layout extraction) | PyMuPDF text extraction | Text-heavy PDFs, article reprints |
| `.tex` | `pandoc` (math → OMML, citations, crossrefs) | Compile → PDF → pdf2docx | Standard LaTeX articles, reports |
| `.tex` (complex) | Compile → PDF → pdf2docx | pandoc | Custom classes, complex layouts, non-pandoc packages |

## Quick Start

```bash
# PDF to DOCX
python skills/pdf-latex-to-docx/scripts/convert.py document.pdf

# LaTeX to DOCX
python skills/pdf-latex-to-docx/scripts/convert.py paper.tex

# Specify output path
python skills/pdf-latex-to-docx/scripts/convert.py document.pdf --output result.docx
```

## Usage

### PDF Conversion

```bash
# Default: pdf2docx with layout preservation
python skills/pdf-latex-to-docx/scripts/convert.py article.pdf

# Multi-page PDF with progress
python skills/pdf-latex-to-docx/scripts/convert.py thesis.pdf -o thesis.docx --verbose
```

### LaTeX Conversion

```bash
# Standard LaTeX (uses pandoc)
python skills/pdf-latex-to-docx/scripts/convert.py paper.tex

# Complex LaTeX with custom class (compile → PDF → convert)
python skills/pdf-latex-to-docx/scripts/convert.py nsfc_proposal.tex --method compile-pdf

# Specify LaTeX compiler
python skills/pdf-latex-to-docx/scripts/convert.py document.tex --latex-compiler xelatex
```

### Batch Conversion

```bash
# Convert all PDFs in a directory
for f in inputs/*.pdf; do
  python skills/pdf-latex-to-docx/scripts/convert.py "$f" -o "outputs/$(basename "${f%.pdf}.docx")"
done
```

## CLI Reference

| Flag | Default | Description |
|------|---------|-------------|
| `input` | _(required)_ | Input file path (.pdf or .tex) |
| `--output`, `-o` | `input_stem.docx` | Output DOCX path |
| `--method` | `auto` | Conversion method: `auto`, `pdf2docx`, `pandoc`, `compile-pdf` |
| `--latex-compiler` | `pdflatex` | LaTeX compiler for compile-pdf fallback: `pdflatex`, `xelatex`, `lualatex` |
| `--verbose`, `-v` | off | Print detailed progress output |
| `--keep-temp` | off | Keep intermediate files (compiled PDF, etc.) |

## Limitations

### PDF → DOCX

- Scanned PDFs (image-only) fall back to PyMuPDF text extraction (text-only, no layout/images)
- Complex multi-column layouts may reorder text blocks
- Embedded fonts are not preserved; DOCX uses default fonts
- Vector graphics may rasterize or lose fidelity

### LaTeX → DOCX

- Custom macros and `\def` commands are not expanded by pandoc
- TikZ/PGF figures become images or placeholders
- BibLaTeX citations are converted but formatting may differ
- Custom class files (`.cls`) may cause compile failures
- Pseudo-code environments (algorithm2e, algorithmicx) may lose formatting

## Dependencies

```bash
pip install pdf2docx python-docx PyMuPDF
```

System requirements:
- `pandoc` (for LaTeX → DOCX)
- `pdflatex` / `xelatex` (for compile-pdf fallback)

## See Also

- [docx-pipeline](../docx-pipeline/SKILL.md) — Full DOCX formatting workflow
- [docx-write](../docx-write/SKILL.md) — Generate DOCX from YAML/JSON specs
- [docx-review](../docx-review/SKILL.md) — Validate DOCX content and formatting
