---
name: docx-write
description: "Generate a submission-ready DOCX from a structured YAML/JSON specification. Supports full paper structure (bilingual titles, authors, abstracts, keywords, sections, references) AND patent documents (CNIPA Chinese + USPTO/EPO/PCT English, all patent types: invention, utility model, design). Uses docx-js (Node.js) for document creation. Trigger when the user says 'generate docx from spec', 'write a paper docx', 'create document from yaml', 'write a patent docx', '生成专利申请文件', or wants to create a NEW DOCX from scratch with formatted content."
---

# DOCX Write

Generate a new DOCX document from a structured YAML or JSON specification using `docx-js` (Node.js). This is a standalone skill for creating documents from scratch — it does not depend on any other skill.

## Prerequisites

- **Node.js**: v18+ with `docx` npm package installed globally
- **Python packages**: `pyyaml`

```bash
npm install -g docx
pip install pyyaml
```

## Input Specification

The skill accepts a YAML (or JSON) file describing the complete paper structure:

```yaml
# Optional metadata for output path resolution
meta:
  paper_name: "my-paper"
  category: "北交模板"
  version: 1

# Optional TOC configuration (also can come from --template)
toc:
  label: "目录"
  headingStyleRange: "1-3"

# Page setup (defaults to A4: 11906 x 16838 DXA, 1-inch margins)
page:
  width: 11906
  height: 16838
  margins: { top: 1440, bottom: 1440, left: 1440, right: 1440 }

# Document style defaults
style:
  default_font: "Times New Roman"
  default_font_size: 24          # Half-points (24 = 12pt)
  default_font_color: "000000"

# Title (bilingual)
title:
  cn: "中文标题"
  en: "English Title"

# Authors
authors:
  - name: "作者一"
    affiliation: "北京交通大学"
    email: "author@bjtu.edu.cn"
    corresponding: true
  - name: "作者二"
    affiliation: "北京交通大学"

# Abstract (bilingual)
abstract:
  cn: "中文摘要内容..."
  en: "English abstract content..."

# Keywords (bilingual)
keywords:
  cn: ["关键词1", "关键词2"]
  en: ["keyword1", "keyword2"]

# Body sections
sections:
  - id: "intro"                     # NEW: optional bookmark anchor
    heading: "Introduction"
    level: 1
    content:
      - type: paragraph
        text: "This is the first paragraph of the introduction."

      - type: paragraph
        text: "As shown in Fig. {{ref:fig-example}}, the architecture..."

      - type: image
        src: "figures/figure1.png"
        bookmark_id: "fig-example"  # optional cross-reference target
        caption: "Fig. {{seq:Figure}}. Example figure."
        width: 500

      - type: table
        bookmark_id: "tab-results"  # optional cross-reference target
        caption: "Table {{seq:Table}}. Experimental results."
        columns: ["Metric", "Value", "Notes"]
        rows:
          - ["Accuracy", "95.2%", "On test set"]
          - ["Precision", "93.1%", "Macro average"]

      - type: list
        ordered: false
        items:
          - "First bullet point"
          - "Second bullet point"

      # Block/display math equation
      - type: math
        latex: "\\mathbf{f}_i = \\sum_{j=1}^{d_{in}} \\Phi_{i,j}(x_j)"
        display: block
        label: "(1)"           # optional equation number

      # Inline math within paragraph text (uses $...$ delimiters)
      - type: paragraph
        text: "We compute $\\sum_{i=1}^n x_i$ as the total loss."

      # Standalone inline math
      - type: math
        latex: "E = mc^2"
        display: inline

  - heading: "Related Work"
    level: 1
    content:
      - type: paragraph
        text: "Prior work in this area..."

  - heading: "Detailed Analysis"
    level: 2
    content:
      - type: paragraph
        text: "A sub-section with detailed analysis."

# References
references:
  - text: "[1] Smith, J. et al. 'Deep Learning for 3D Detection.' CVPR 2024."
    doi: "10.xxxx/xxxxx"
  - text: "[2] 张伟. '点云目标检测方法综述.' 自动化学报, 2023."
```

## Workflow

### Basic generation

```bash
python scripts/write_docx.py paper.yaml --output paper.docx
```

### Generate with category/paper/version path

```bash
python scripts/write_docx.py paper.yaml \
  --category 北交模板 --paper 论文名 --version 1
```

This produces `outputs/论文名/北交模板/v1/论文名-generated.docx`.

### Skip validation

```bash
python scripts/write_docx.py paper.yaml --skip-validate
```

### See generated JS code

```bash
python scripts/write_docx.py paper.yaml --verbose
```

## Patent Documents

When `type: patent` is set in the spec, `docx-write` dispatches to the `patent-writer` subskill (see `patent-writer/SKILL.md`). This supports:

| Jurisdiction | Patent Types | Language |
|-------------|-------------|----------|
| CN (CNIPA) | invention, utility_model, design | Chinese (zh) |
| US (USPTO) | invention, design | English (en) |
| EP (EPO) | invention | English (en) |
| PCT | invention | English (en) |

### Patent spec example (CN invention)

```yaml
type: patent
patent:
  jurisdiction: cn
  patent_type: invention
  language: zh
  title: "一种基于深度学习的焊缝缺陷检测方法及系统"
  applicant: "{{申请人}}"
  inventors:
    - name: "{{发明人}}"
  abstract: "本发明公开了一种..."
  claims:
    - "1. 一种...方法，其特征在于，包括以下步骤：..."
  drawings:
    - ref: "图1"
      src: "figs/figure_1.png"
      caption: "本发明实施例中系统整体结构示意图"

sections:
  - heading: "技术领域"
    level: 1
    content:
      - type: paragraph
        text: "本发明涉及..."
```

### Template-based workflow

```bash
# Copy a template
cp patent-writer/templates/cn/invention.yaml my_patent.yaml

# Fill in {{PLACEHOLDER}} values, then generate
python scripts/write_docx.py my_patent.yaml --output patent.docx
```

Templates are available in `patent-writer/templates/cn/` and `patent-writer/templates/en/`. See `patent-writer/references/` for writing guides.

## Math Formulas

Math formulas are written in LaTeX and converted to native OMML (Office Math Markup Language) via pandoc during post-processing. The `docx` npm package generates placeholder markers, and a Python step replaces them with live equations.

### Content types

| Field | Type | Description |
|-------|------|-------------|
| `type` | `"math"` | Math content item |
| `latex` | string | LaTeX equation string |
| `display` | `"block"` or `"inline"` | Display mode (default: `"block"`) |
| `label` | string (optional) | Equation number, appended as `\qquad (N)` |
| `alignment` | string (optional) | Paragraph alignment for block math |
| `spacing_before` | int (optional) | Spacing before in DXA |
| `spacing_after` | int (optional) | Spacing after in DXA |

### Inline math in paragraphs

Use `$...$` delimiters within paragraph text for inline math — no separate content item needed.

### Requirements

- **pandoc** must be on `PATH` (used for LaTeX → OMML conversion)
- Formulas without pandoc will produce a warning and the doc will contain placeholder text

## Table of Contents

Set a top-level `toc` key to generate a clickable Table of Contents:

```yaml
toc:
  label: "目录"                  # TOC heading text (default: "Table of Contents")
  headingStyleRange: "1-3"       # which heading outline levels to include
```

The TOC appears after the Keywords section and before body sections. In Word, Ctrl+click any TOC entry to jump to that heading. Press F9 to update page numbers.

## Auto-Numbering & Cross-References

The generator supports Word field-based auto-numbering and clickable cross-references via inline markers in caption and paragraph text. When you open the document in Word and press **Ctrl+A then F9**, all numbers and references recalculate automatically.

### Marker Syntax

| Marker | Purpose | Example Usage |
|--------|---------|---------------|
| `{{seq:NAME}}` | Auto-incrementing sequence number (SEQ field) | `"图 {{seq:Figure}}"` → "图 1", "图 2"... |
| `{{ref:ID}}` | Cross-reference to a bookmark (clickable) | `"如{{ref:fig-arch}}所示"` → "如图 1所示" |
| `{{pageref:ID}}` | Page number of a bookmark | `"(第{{pageref:fig-arch}}页)"` → "(第3页)" |

### Bookmark IDs

Sections, images, and tables can carry optional identifiers for cross-referencing:

```yaml
sections:
  - id: "intro"              # bookmark target: _bm_intro
    heading: "引言"
    level: 1
    content:
      - type: image
        bookmark_id: "fig-arch"   # bookmark target: _bm_fig-arch
        caption: "图 {{seq:Figure}}. 系统架构图"
```

Bookmark IDs are automatically prefixed with `_bm_` to avoid collisions. Use the same ID in `{{ref:ID}}` and `{{pageref:ID}}` markers to create clickable cross-references.

### How It Works

1. **SEQ fields**: `{{seq:Figure}}` generates a Word `{ SEQ Figure \* ARABIC }` field. Each SEQ with the same name auto-increments.
2. **Bookmarks**: The first `{{seq:NAME}}` in a caption is wrapped in a `Bookmark` if `bookmark_id` is set. Section headings are bookmarked when `id` is set.
3. **Cross-references**: `{{ref:ID}}` generates a clickable `{ REF _bm_ID \h }` field that shows the bookmarked number and hyperlinks to it.
4. **Page references**: `{{pageref:ID}}` generates a `PAGEREF` field showing the page where the bookmark is located.

### Backward Compatibility

All new fields and markers are optional. Existing specs without `bookmark_id`, `id`, `toc`, or `{{...}}` markers continue to work unchanged.

## Image & Table Keep-Together

Images and tables automatically stay on the same page as their captions:

- **Image paragraphs** use `keepNext: true` to prevent page breaks between the image and its caption
- **Caption paragraphs** use `keepLines: true` to prevent captions from splitting across pages
- **Table captions** use `keepNext: true` to prevent page breaks between caption and table

These properties are always applied and require no special configuration.

## Style Templates

Use `--template` to load a style template YAML (extracted via `docx-template-extract`). The template provides defaults for page layout, fonts, caption patterns, TOC, and feature flags:

```bash
# Generate from spec with template defaults
python scripts/write_docx.py paper.yaml --template styles/北交模板/北交模板.yaml

# Generate a preview DOCX to verify style understanding
python scripts/write_docx.py --template styles/北交模板/北交模板.yaml --preview
```

**Template fields used:**

| Template field | Effect |
|---------------|--------|
| `document_defaults.run` | Default font and font size |
| `page` | Page dimensions and margins |
| `caption_templates` | Figure/table caption numbering patterns (`prefix`, `suffix`, `seq_name`) |
| `toc` | TOC label and heading range |
| `features` | Behavioral flags (keep_together, auto_numbering, clickable_toc, clickable_cross_refs) |

Spec fields always take priority over template defaults — the template acts as a base layer.

### Preview Mode

The `--preview` flag generates a sample document demonstrating all style features from the template. No spec file is needed:

```bash
python scripts/write_docx.py --template styles/北交模板/北交模板.yaml --preview -o preview.docx
```

Preview content includes: bilingual title, authors, abstracts, keywords, TOC, sample image with caption, sample table with caption, cross-references, page references, multi-level headings, lists, references, and inline math.

---

## How it works

1. **Load spec** — Python reads the YAML/JSON specification
2. **Dispatch** — If `type: patent`, delegates to `gen_patent.py`; otherwise proceeds with paper generation
3. **Generate JS** — Produces a Node.js script using the `docx-js` API, with support for:
   - SEQ fields (`{{seq:NAME}}`) for auto-numbering figures, tables, equations
   - Bookmarks and clickable cross-references (`{{ref:ID}}`, `{{pageref:ID}}`)
   - Table of Contents with heading hyperlinks
   - Keep-together paragraph properties for image/table caption binding
4. **Execute** — Runs the JS script via `node` to create the DOCX
5. **Post-process math** — Replaces formula placeholder markers with native OMML via pandoc
6. **Validate** — Verifies the output is a well-formed DOCX (ZIP structure check)

The generated JS script follows critical `docx-js` rules: explicit page sizes, proper numbering config for lists (never unicode bullets), dual table widths with DXA units, `ShadingType.CLEAR` for table cells, and `PageBreak` inside `Paragraph` elements.

## CLI Reference

| Flag | Default | Description |
|------|---------|-------------|
| `--output`, `-o` | `{stem}.docx` | Output DOCX path |
| `--template`, `-t` | none | Style template YAML for defaults (page, font, captions, TOC) |
| `--preview` | off | Generate preview DOCX from template (requires `--template`) |
| `--category`, `-c` | none | Template category for output path (e.g. `北交模板`) |
| `--paper`, `-p` | spec stem | Paper name for output path |
| `--version`, `-v` | auto-detect | Version number |
| `--skip-validate` | off | Skip DOCX validation |
| `--verbose`, `-V` | off | Show generated JavaScript code |

## Dependencies

- Node.js: `npm install -g docx`
- Python: `pyyaml`
