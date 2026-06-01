---
name: docx-template-apply
description: "Apply a formatting template (from docx-template-extract) to modify and style a DOCX document. Updates styles, fonts, colors, page layout, margins, and numbering to match a reference template. Trigger when the user says 'apply template to docx', 'format this document like', 'make this document match the style of', 'apply formatting template', or wants to restyle a DOCX to match another document's look and feel."
---

# DOCX Template Application

Apply a formatting template (produced by **docx-template-extract**) to a target DOCX document, giving it the same visual style as the original template source.

Templates may include `body_paragraph_formats`, which are detailed paragraph-index profiles extracted from `document.xml`. By default these profiles enrich reusable styles and prevent destructive cleanup of body-level direct fonts. Apply them by exact paragraph index only when the target has the same paragraph structure as the reference.

Templates may also include `semantic_format_rules`. These are strict role-based rules inferred from instruction-style journal templates, such as Chinese title, author line, affiliation, abstract, keywords, section headings, captions, references, table cells, and body text. The apply script rewrites matching paragraphs to the exact rule fonts and sizes instead of preserving accidental source formatting.

## Prerequisites

- **docx skill** (required): Provides `unpack.py` and `pack.py` for DOCX manipulation
- **Python packages**: `defusedxml`, `pyyaml`

```bash
pip install defusedxml pyyaml
```

## Workflow

### Step 1: Ensure you have a template

You need a YAML template file. Either:

- Extract one from a reference document:
  ```bash
  cd /home/rczx/workspace/rinbarpen/work/docx-polar/.claude/skills/docx-template-extract
  python scripts/extract_template.py reference.docx --output my-template.yaml
  ```
- Hand-craft one following the format in the [Template Reference](#template-format-reference) below

### Step 2: Apply the template

**Single document (backward-compatible):**
```bash
cd /home/rczx/workspace/rinbarpen/work/docx-polar/.claude/skills/docx-template-apply

python scripts/apply_template.py target.docx my-template.yaml
```

This creates `{target_stem}-formatted.docx` beside the source. Specify a custom output path:

```bash
python scripts/apply_template.py target.docx my-template.yaml --output result.docx
```

**Applying with category (template/article name):**

Organize output by template name or article name with automatic version management:

```bash
# Auto-detect next version, auto-locate template from styles/北交模板/
python scripts/apply_template.py target.docx --category 北交模板

# Specify version explicitly
python scripts/apply_template.py target.docx --category 道路交通场景双级样条网络 --version 2

# Specify template explicitly with category output
python scripts/apply_template.py target.docx my-template.yaml --category 北交模板 --version 1

# Preserve existing direct formatting (skip content cleanup)
python scripts/apply_template.py target.docx --category 北交模板 --preserve-formatting

# Exact same-structure documents only: apply extracted body profiles by paragraph index
python scripts/apply_template.py target.docx --category 北交模板 --apply-body-formats
```

This creates `outputs/<category>/v<version>/<stem>-formatted.docx`.

**Batch mode (parallel per-document):**
```bash
# Glob pattern
python scripts/apply_template.py --batch "chapter*.docx" my-template.yaml --output-dir ./out

# Batch with category output
python scripts/apply_template.py --batch "chapters/*.docx" --category journal --jobs 4

# Multiple files
python scripts/apply_template.py --batch doc1.docx doc2.docx my-template.yaml --jobs 4
```

The `--batch` flag enables **per-document parallelism**: each file runs in its own thread. Control concurrency with `--jobs N` (defaults to min(files, CPU count)).

### Directory Structure Convention

```
styles/
  北交模板/
    <stem>-template.yaml
  道路交通场景双级样条网络/
    <stem>-template.yaml

outputs/
  北交模板/
    v1/
      <stem>-formatted.docx
    v2/
      <stem>-formatted.docx
  道路交通场景双级样条网络/
    v1/
      <stem>-formatted.docx
```

Templates are stored in `styles/<name>/` and generated outputs go to `outputs/<name>/v<version>/`. When `--category` is provided and no template file is specified, the script automatically searches for the single YAML template in `styles/<name>/`.

### Body Content Formatting Cleanup

By default, the script also strips **direct formatting** from the document body to ensure the template styles take full effect:

| Level | Stripped Properties |
|-------|-------------------|
| **Run** (`w:rPr`) | Font (`w:rFonts`), size (`w:sz`/`w:szCs`), color (`w:color`), bold/italic/underline/strikethrough/caps |
| **Paragraph** (`w:pPr`) | Spacing (`w:spacing`), indentation (`w:ind`), alignment (`w:jc`) |
| **Paragraph run** (`w:pPr/w:rPr`) | Paragraph-level font, size, color, bold/italic overrides |

Use `--preserve-formatting` to skip this cleanup and keep existing direct formatting.

When the template contains `body_paragraph_formats`, default cleanup is conservative: body-level direct font formatting is preserved because it may be the only reliable source of visible Word formatting. Use `--apply-body-formats` only for same-structure documents; it removes stale direct formatting first, then reapplies the extracted paragraph-index profiles from the reference document.

### Step 3: Verify the result

Check the output document:

```bash
# Text extraction
pandoc result.docx -o result.md

# Visual inspection — unpack and review XML
~/.claude/skills/docx/scripts/office/unpack.py result.docx /tmp/verify/
```

## What Gets Applied

| Document Aspect | What Changes |
|----------------|-------------|
| **Document defaults** | Default font, size, color, paragraph spacing |
| **Paragraph styles** | Heading 1-9, Normal, List Paragraph, etc. |
| **Character styles** | Hyperlink, Emphasis, Strong, etc. |
| **Page layout** | Paper size, margins, orientation |
| **Numbering** | Bullet and numbered list formatting |
| **Theme colors** | Color scheme (accent colors, hyperlink colors) |
| **Theme fonts** | Major/minor font scheme |
| **Font table** | Declared fonts for compatibility |
| **Body paragraph formats** | Paragraph-indexed direct font/size/alignment profiles for title, abstract, headings, captions, and other unstyled content |
| **Semantic format rules** | Role-based exact fonts/sizes for journal structures; used for strict template application |

## Template Format Reference

A minimal template looks like:

```yaml
document_defaults:
  run:
    font: "Calibri"
    fontSize: 20
    fontColor: "000000"
  paragraph:
    spacing_after: 200
    alignment: "both"

paragraph_styles:
  - id: "Normal"
    name: "Normal"
    run:
      font: "Calibri"
      fontSize: 20
    paragraph:
      spacing_after: 200

page:
  width: 12240
  height: 15840
  margins:
    top: 1440
    bottom: 1440
    left: 1440
    right: 1440

numbering:
  - id: 0
    abstractNumId: 0
    levels:
      - level: 0
        format: "bullet"
        text: "•"

theme:
  schemeName: "Office"
  colors:
    a:dk1: "000000"
    a:lt1: "FFFFFF"
    a:accent1: "2E75B6"
  fonts:
    major: "Calibri Light"
    minor: "Calibri"
```

See the **[docx-template-extract SKILL.md](../docx-template-extract/SKILL.md)** for a complete property reference.

## Editable Template Values

You can hand-edit the YAML template to fine-tune styling. Common customizations:

- **Font**: Change `font` and `major`/`minor` values
- **Colors**: Update hex values in `colors` and `fontColor`
- **Margins**: Adjust `page.margins` values (in DXA, 1440 = 1 inch)
- **Spacing**: Modify `spacing_before`/`spacing_after` (in DXA)
- **Heading styles**: Edit `paragraph_styles[].run` or `.paragraph` values for headings

### DXA Units Reference

| DXA | Inches | Purpose |
|-----|--------|---------|
| 1440 | 1" | Standard margin |
| 720 | 0.5" | Half-inch indent |
| 360 | 0.25" | Hanging indent |
| 240 | — | Single line spacing |
| 480 | — | 1.5 line spacing |
| 20 | — | 1pt font size |
| 12240 | 8.5" | US Letter width |
| 15840 | 11" | US Letter height |
| 11906 | — | A4 width |
| 16838 | — | A4 height |

## Hand-Editing a Template

You can create a template from scratch without extracting from a reference document. Start with this minimal template:

```yaml
document_defaults:
  run:
    font: "Arial"
    fontSize: 22
    fontColor: "333333"
  paragraph:
    spacing_after: 120

paragraph_styles:
  - id: "Normal"
    name: "Normal"
    run:
      font: "Arial"
      fontSize: 22

page:
  width: 12240
  height: 15840
  margins:
    top: 1440
    bottom: 1440
    left: 1440
    right: 1440
```

Save it as a `.yaml` file and pass it to `apply_template.py`.

## Parallel Processing

The script has two levels of parallelism:

### Level 1: Per-File XML Parallelism (always on)

Within each document, 5 independent XML files are processed in parallel threads:

| Thread | XML File | Operation |
|--------|----------|-----------|
| 1 | `styles.xml` | Document defaults + paragraph/character styles |
| 2 | `document.xml` | Page layout (size, margins, orientation) |
| 3 | `numbering.xml` | Bullet and numbered list definitions |
| 4 | `theme/theme1.xml` | Color scheme and font scheme |
| 5 | `fontTable.xml` | Declared fonts |

### Level 2: Batch Mode (per-document parallelism)

Use `--batch` to process multiple DOCX files concurrently:

```bash
# All files in parallel, 4 workers
python scripts/apply_template.py --batch "reports/*.docx" template.yaml --jobs 4
```

Each document runs the full pipeline (unpack → parallel XML modify → repack) in its own thread. Output goes to `--output-dir` or alongside each source file.

### CLI Reference

| Flag | Default | Description |
|------|---------|-------------|
| `--category`, `-c` | (none) | Output category; saves to `outputs/<category>/v<version>/`; enables template auto-lookup in `styles/<category>/` |
| `--version`, `-v` | auto-detect | Version number (creates `v<N>` subdirectory) |
| `--preserve-formatting` | off | Skip body content formatting cleanup |
| `--batch`, `-b` | off | Enable batch mode (target may be a glob pattern) |
| `--jobs`, `-j` | min(files, CPUs) | Number of parallel workers (batch) |
| `--output`, `-o` | `{stem}-formatted.docx` | Output path (single-doc mode; overrides category/version) |
| `--output-dir` | source folder | Output directory (batch mode; overrides category/version) |
| `--chunks` | 1 | Split body into N parallel chunks (extension point)

## Dependencies

- `~/.claude/skills/docx/scripts/office/unpack.py` — DOCX unpacking
- `~/.claude/skills/docx/scripts/office/pack.py` — DOCX repacking with validation
- Python: `defusedxml`, `pyyaml`

## See Also

- [docx-template-extract](../docx-template-extract/SKILL.md) — Extract templates from reference documents
- [docx](../../../.claude/skills/docx/SKILL.md) — General DOCX creation and editing
