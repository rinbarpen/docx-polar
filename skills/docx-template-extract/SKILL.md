---
name: docx-template-extract
description: "Extract formatting templates from reference DOCX documents. Captures fonts, paragraph/character styles, page layout, margins, numbering (bullets/lists), theme colors, font declarations, and math formula presence into a YAML template file. Trigger when the user says 'extract template from docx', 'capture formatting from docx', 'create a template from this word document', or wants to reuse the visual style of an existing DOCX file in other documents."
---

# DOCX Template Extraction

Extract a reusable formatting template from a reference Word document. The template captures styles, page layout, numbering, theme information, body-level paragraph/run formatting, and math formula detection. This is important for Chinese journal templates and legacy `.doc` files where title, abstract, headings, captions, and body text may be formatted directly instead of through named styles.

## Prerequisites

- **docx skill** (required): Provides `unpack.py` for DOCX unpacking
- **Python packages**: `defusedxml`, `pyyaml`

```bash
pip install defusedxml pyyaml
```

## Workflow

### Step 1: Run the extraction script

```bash
cd /home/rczx/workspace/rinbarpen/work/docx-polar/skills/docx-template-extract

python scripts/extract_template.py path/to/reference.docx
```

This accepts `.docx` directly. Legacy `.doc`, `.odt`, and `.rtf` references are converted through LibreOffice/soffice before extraction. It creates `{filename}-template.yaml` in the current directory, or inside `styles/<category>/` when `--category` is used.

```bash
# Custom output path
python scripts/extract_template.py path/to/reference.docx --output my-template.yaml

# With category (saves to styles/<category>/)
python scripts/extract_template.py path/to/reference.docx --category 北交模板
```

### Step 2: Review the template (optional)

The generated YAML is human-readable and can be edited. Key sections:

```yaml
document_defaults:
  run:
    font: "Calibri"
    fontSize: 20
    fontColor: "000000"

paragraph_styles:
  - id: "Heading1"
    name: "heading 1"
    run:
      font: "Calibri Light"
      fontSize: 32
      bold: true
      fontColor: "2E75B6"

page:
  width: 12240
  height: 15840
  margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 }

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
    dk1: "000000"
    lt1: "FFFFFF"

body_paragraph_formats:
  - index: 0
    style: "Normal"
    textSample: "论文标题"
    has_math: false
    paragraph: { alignment: "center" }
    run: { font: "Times New Roman", fontSize: 28 }

body_paragraph_formats_mode: "style_enrichment"

semantic_format_rules:
  - role: "cn_title"
    match: { type: "front_index", index: 0 }
    run: { fontEastAsia: "黑体;SimHei", fontSize: 40, bold: true }
```

### Style Config (caption_templates, toc, features)

Since the extraction auto-detects caption numbering patterns, the output also includes reusable config for `docx-write`:

```yaml
caption_templates:
  figure_cn:
    prefix: "图 "
    suffix: ""
    seq_name: "Figure"
  figure_en:
    prefix: "Fig. "
    suffix: ". "
    seq_name: "Figure"
  table_cn:
    prefix: "表 "
    suffix: ""
    seq_name: "Table"

toc:
  label: "目录"
  headingStyleRange: "1-3"

features:
  keep_together: true
  auto_numbering: true
  clickable_toc: true
  clickable_cross_refs: true
```

These feed directly into `docx-write --template` for style-consistent generation and `--preview` for format verification.

## Math Formula Detection

The extraction detects native OMML (Office Math Markup Language) formulas in the template DOCX. Paragraphs containing math are tagged in `body_paragraph_formats`:

```yaml
body_paragraph_formats:
  - index: 15
    style: "Normal"
    textSample: "E = mc²"
    has_math: true
    math_type: "display"      # "display" for <m:oMathPara>, "inline" for <m:oMath>
    paragraph: { alignment: "center" }
```

| Field | Values | Description |
|-------|--------|-------------|
| `has_math` | `true`/absent | Whether the paragraph contains OMML math |
| `math_type` | `"display"` / `"inline"` | Display math (`<m:oMathPara>`) or inline (`<m:oMath>`) |

A summary is printed during extraction when formulas are detected:

```
  Detected 3 paragraph(s) containing math formulas
```

## Template Format Reference

| Section | XML Source | Description |
|---------|-----------|-------------|
| `document_defaults.run` | `styles.xml` → `w:docDefaults/w:rPrDefault` | Default font, size, color |
| `paragraph_styles[]` | `styles.xml` → `w:style[@type="paragraph"]` | Named paragraph styles |
| `character_styles[]` | `styles.xml` → `w:style[@type="character"]` | Named character styles |
| `page` | `document.xml` → `w:sectPr` | Page size, margins |
| `numbering[]` | `numbering.xml` → `w:num` | Bullet/numbered list defs |
| `theme` | `theme/theme1.xml` → `a:themeElements` | Color and font scheme |
| `fonts[]` | `fontTable.xml` → `w:font` | Declared font names |
| `body_paragraph_formats[]` | `document.xml` → body `w:p` | Direct body formatting per paragraph index, including math detection |
| `semantic_format_rules[]` | Template body text + extracted formats | Role-based rules for title, abstract, headings, etc. |
| `caption_templates` | Auto-detected from rules/text | Figure/table caption patterns (`prefix`, `suffix`, `seq_name`) |
| `toc` | Language-detected | TOC label and heading range (e.g. `label: "目录"`) |
| `features` | Always included | Behavioral flags for docx-write generation |

## Dependencies

- `~/.claude/skills/docx/scripts/office/unpack.py` — DOCX unpacking
- Python: `defusedxml`, `pyyaml`

## See Also

- [docx-write](../docx-write/SKILL.md) — Generate DOCX from YAML spec (supports math)
- [docx-pipeline](../docx-pipeline/SKILL.md) — Full formatting pipeline
