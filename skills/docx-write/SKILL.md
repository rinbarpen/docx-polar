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
  - heading: "Introduction"
    level: 1
    content:
      - type: paragraph
        text: "This is the first paragraph of the introduction."

      - type: paragraph
        text: "Another paragraph."

      - type: image
        src: "figures/figure1.png"
        caption: "Fig. 1. Example figure."
        width: 500

      - type: table
        caption: "Table 1. Experimental results."
        columns: ["Metric", "Value", "Notes"]
        rows:
          - ["Accuracy", "95.2%", "On test set"]
          - ["Precision", "93.1%", "Macro average"]

      - type: list
        ordered: false
        items:
          - "First bullet point"
          - "Second bullet point"

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

---

## How it works

1. **Load spec** — Python reads the YAML/JSON specification
2. **Dispatch** — If `type: patent`, delegates to `gen_patent.py`; otherwise proceeds with paper generation
3. **Generate JS** — Produces a Node.js script using the `docx-js` API
4. **Execute** — Runs the JS script via `node` to create the DOCX
5. **Validate** — Verifies the output is a well-formed DOCX (ZIP structure check)

The generated JS script follows critical `docx-js` rules: explicit page sizes, proper numbering config for lists (never unicode bullets), dual table widths with DXA units, `ShadingType.CLEAR` for table cells, and `PageBreak` inside `Paragraph` elements.

## CLI Reference

| Flag | Default | Description |
|------|---------|-------------|
| `--output`, `-o` | `{stem}.docx` | Output DOCX path |
| `--category`, `-c` | none | Template category for output path (e.g. `北交模板`) |
| `--paper`, `-p` | spec stem | Paper name for output path |
| `--version`, `-v` | auto-detect | Version number |
| `--skip-validate` | off | Skip DOCX validation |
| `--verbose`, `-V` | off | Show generated JavaScript code |

## Dependencies

- Node.js: `npm install -g docx`
- Python: `pyyaml`
