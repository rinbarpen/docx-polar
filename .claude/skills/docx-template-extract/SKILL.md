---
name: docx-template-extract
description: "Extract formatting templates from reference DOCX documents. Captures fonts, paragraph/character styles, page layout, margins, numbering (bullets/lists), theme colors, and font declarations into a YAML template file. Trigger when the user says 'extract template from docx', 'capture formatting from docx', 'create a template from this word document', or wants to reuse the visual style of an existing DOCX file in other documents."
---

# DOCX Template Extraction

Extract a reusable formatting template from a reference Word document. The template captures styles, page layout, numbering, theme information, and the visible body-level paragraph/run formatting that often carries real fonts in Word files. This is important for Chinese journal templates and legacy `.doc` files where title, abstract, headings, captions, and body text may be formatted directly instead of through named styles.

## Prerequisites

- **docx skill** (required): Provides `unpack.py` and `pack.py` for DOCX manipulation
- **Python packages**: `defusedxml`, `pyyaml`

```bash
pip install defusedxml pyyaml
```

## Workflow

### Step 1: Run the extraction script

```bash
cd /home/rczx/workspace/rinbarpen/work/docx-polar/.claude/skills/docx-template-extract

python scripts/extract_template.py path/to/reference.docx
```

This accepts `.docx` directly. Legacy `.doc`, `.odt`, and `.rtf` references are converted through LibreOffice/soffice before extraction. It creates `{filename}-template.yaml` in the current directory, or inside `styles/<category>/` when `--category` is used. Specify a custom output path:

```bash
python scripts/extract_template.py path/to/reference.docx --output my-template.yaml
```

### Extracting with Category

Save the template under a specific template name or article name:

```bash
cd /home/rczx/workspace/rinbarpen/work/docx-polar/.claude/skills/docx-template-extract

# Extract template into styles/北交模板/
python scripts/extract_template.py path/to/reference.docx --category 北交模板

# Extract template into styles/道路交通场景/
python scripts/extract_template.py path/to/reference.docx --category 道路交通场景
```

This creates `styles/<category>/<stem>-template.yaml`, auto-creating directories as needed.

This creates `styles/<category>/<stem>-template.yaml`, auto-creating directories as needed.

### Step 2: Review the template (optional)

The generated YAML is human-readable and can be edited. Key sections:

```yaml
document_defaults:
  run:
    font: "Calibri"
    fontSize: 20          # 10pt in half-points
    fontColor: "000000"
  paragraph:
    spacing_after: 200

paragraph_styles:
  - id: "Heading1"
    name: "heading 1"
    run:
      font: "Calibri Light"
      fontSize: 32        # 16pt
      bold: true
      fontColor: "2E75B6"
    paragraph:
      spacing_before: 360
      outlineLevel: 0

character_styles:
  - id: "Hyperlink"
    name: "Hyperlink"
    run:
      fontColor: "0563C1"
      underline: "single"

page:
  width: 12240            # DXA (8.5 inches)
  height: 15840           # DXA (11 inches)
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
        indent_left: 720
        indent_hanging: 360

theme:
  schemeName: "Office"
  colors:
    a:dk1: "000000"
    a:lt1: "FFFFFF"
    a:dk2: "44546A"
    a:accent1: "2E75B6"
    ...
fonts:
    major: "Calibri Light"
    minor: "Calibri"

body_paragraph_formats:
  - index: 0
    style: "Normal"
    textSample: "论文标题"
    paragraph:
      alignment: "center"
    paragraph_run:
      font: "黑体"
      fontEastAsia: "黑体"
      fontSize: 28
    run:
      font: "Times New Roman"
      fontEastAsia: "黑体"
      fontSize: 28
body_paragraph_formats_mode: "style_enrichment"

semantic_format_rules:
  - role: "cn_title"
    match:
      type: "front_index"
      index: 0
    run:
      fontEastAsia: "黑体;SimHei"
      fontSize: 40
      bold: true
```

You can manually tweak any values before applying. For example, change fonts, colors, margins, or spacing.

## Template Format Reference

| Section | XML Source | Description |
|---------|-----------|-------------|
| `document_defaults.run` | `styles.xml` → `w:docDefaults/w:rPrDefault` | Default font, size, color for runs |
| `document_defaults.paragraph` | `styles.xml` → `w:docDefaults/w:pPrDefault` | Default spacing, alignment |
| `paragraph_styles[]` | `styles.xml` → `w:style[@type="paragraph"]` | Named paragraph styles (Heading 1, Normal, etc.) |
| `character_styles[]` | `styles.xml` → `w:style[@type="character"]` | Named character styles (Hyperlink, etc.) |
| `page` | `document.xml` → `w:sectPr` | Page size, margins, orientation |
| `numbering[]` | `numbering.xml` → `w:num` | Bullet and numbered list definitions |
| `theme` | `theme/theme1.xml` → `a:themeElements` | Color scheme and font scheme |
| `fonts[]` | `fontTable.xml` → `w:font` | Declared font names |
| `body_paragraph_formats[]` | `document.xml` → body `w:p`, `w:pPr/w:rPr`, `w:r/w:rPr` | Visible direct body formatting by paragraph index, including fonts/sizes for unstyled templates |
| `semantic_format_rules[]` | Template body instruction text + extracted formats | Strict role-based rules for title, authors, abstract, headings, captions, references, table cells, and body text |

## Run Properties (w:rPr)

| Template Key | XML Element | Description |
|-------------|-------------|-------------|
| `font` | `w:rFonts[w:ascii]` | Font name |
| `fontSize` | `w:sz[w:val]` | Size in half-points (e.g., 20 = 10pt) |
| `fontSizeCs` | `w:szCs[w:val]` | Size for complex scripts |
| `fontColor` | `w:color[w:val]` | Hex color (e.g., "2E75B6") |
| `bold` | `w:b` | Bold (presence = true) |
| `italic` | `w:i` | Italic (presence = true) |
| `underline` | `w:u[w:val]` | Underline style ("single", "double", etc.) |
| `strike` | `w:strike` | Strikethrough |
| `smallCaps` | `w:smallCaps` | Small capitals |
| `caps` | `w:caps` | All capitals |
| `fontSpacing` | `w:spacing[w:val]` | Character spacing in DXA |

## Paragraph Properties (w:pPr)

| Template Key | XML Element | Description |
|-------------|-------------|-------------|
| `spacing_before` | `w:spacing[w:before]` | Space before paragraph (DXA) |
| `spacing_after` | `w:spacing[w:after]` | Space after paragraph (DXA) |
| `spacing_line` | `w:spacing[w:line]` | Line spacing (240 = single) |
| `spacing_lineRule` | `w:spacing[w:lineRule]` | Line rule ("auto", "exact", "atLeast") |
| `alignment` | `w:jc[w:val]` | "left", "center", "right", "both" |
| `indent_left` | `w:ind[w:left]` | Left indent (DXA) |
| `indent_right` | `w:ind[w:right]` | Right indent (DXA) |
| `indent_firstLine` | `w:ind[w:firstLine]` | First line indent (DXA) |
| `indent_hanging` | `w:ind[w:hanging]` | Hanging indent (DXA) |
| `outlineLevel` | `w:outlineLvl[w:val]` | Heading level (0-8) |
| `keepNext` | `w:keepNext` | Keep with next paragraph |
| `pageBreakBefore` | `w:pageBreakBefore` | Page break before paragraph |

## Dependencies

- `~/.claude/skills/docx/scripts/office/unpack.py` — DOCX unpacking
- Python: `defusedxml`, `pyyaml`

## See Also

- [docx-template-apply](../docx-template-apply/SKILL.md) — Apply this template to another DOCX
- [docx](../../../.claude/skills/docx/SKILL.md) — General DOCX creation and editing
