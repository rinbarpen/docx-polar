# Patent Writer — YAML Spec Templates (Subskill of docx-write)

Generate submission-ready patent DOCX files from YAML specifications. Lives as a subdirectory within `docx-write` — dispatch is automatic when the spec contains `type: patent`.

Supports both **Chinese (CNIPA)** and **English (USPTO/EPO/PCT)** patent formats, covering all three patent types.

## Prerequisites

- Node.js v18+ with `docx` npm package installed globally
- Python `pyyaml`
- Run via the parent skill:
  ```bash
  python scripts/write_docx.py your_patent_spec.yaml --output patent.docx
  ```

## Quick Start

1. Choose a template from `templates/cn/` or `templates/en/`
2. Copy it and fill in the `{{PLACEHOLDER}}` values
3. Run `write_docx.py` on your spec

## Patent Types Supported

| Type | Chinese (CNIPA) | English (USPTO/EPO/PCT) |
|------|----------------|------------------------|
| Invention (发明专利) | `templates/cn/invention.yaml` | `templates/en/invention.yaml` |
| Utility Model (实用新型) | `templates/cn/utility_model.yaml` | — (no direct equivalent) |
| Design (外观设计) | `templates/cn/design.yaml` | `templates/en/design.yaml` |

## Specification Schema

Top-level keys:
- `type: patent` (required — triggers patent generation)
- `meta` — metadata for output path resolution (same as paper)
- `page` — page dimensions and margins (same as paper)
- `patent` — all patent-specific content (see below)

### `patent` fields (common)

| Field | Required | Description |
|-------|----------|-------------|
| `jurisdiction` | Yes | `cn` (CNIPA), `us` (USPTO), `ep` (EPO), or `pct` (PCT) |
| `patent_type` | Yes | `invention`, `utility_model`, or `design` |
| `language` | Yes | `zh` or `en` |
| `title` | Yes | Patent title |
| `applicant` | No | Applicant/assignee name |
| `inventors` | No | List of `{name: "..."}` objects |
| `abstract` | No | Patent abstract |
| `claims` | No | List of claim strings (numbered automatically) |
| `drawings` | No | List of `{ref, src, caption}` objects |

### `patent` fields (CN design only)

| Field | Required | Description |
|-------|----------|-------------|
| `product_name` | Yes | Product name |
| `product_purpose` | Yes | Brief description of product purpose |
| `design_features` | Yes | Design points / key features |
| `primary_view` | Yes | Name of the primary reference view |
| `views` | No | List of `{name, src}` for each view |

### `patent` fields (EN design only)

| Field | Required | Description |
|-------|----------|-------------|
| `figure_description` | No | String describing all figures |
| `claim` | No | Single design claim |
| `related_applications` | No | Cross-reference info (USPTO) |

### `sections` (invention/utility model)

Standard `docx-write` sections array. Each section has `heading`, `level`, and `content` (paragraphs, images, tables, lists). The body sections form the patent specification.

## Formatting Notes

### Chinese (CNIPA)
- Body font: 宋体 (SimSun), 12pt
- Heading font: 黑体 (SimHei)
- Section order: 权利要求书 → 技术领域 → 背景技术 → 发明内容 → 附图说明 → 具体实施方式 → 说明书附图 → 说明书摘要
- Claims with hanging indent and "根据权利要求X所述的..." format

### English (USPTO/EPO/PCT)
- Body font: Times New Roman, 12pt
- Heading font: Times New Roman, bold, all caps
- Section order: CROSS-REFERENCE → ABSTRACT → body sections → CLAIMS → DRAWINGS
- Claims with hanging indent and "The method of claim X, wherein..." format

## CLI Reference

Same as docx-write — all flags work identically:

| Flag | Description |
|------|-------------|
| `--output`, `-o` | Output DOCX path |
| `--category`, `-c` | Template category for output path |
| `--paper`, `-p` | Paper/patent name for output path |
| `--version`, `-v` | Version number |
| `--skip-validate` | Skip DOCX validation |
| `--verbose`, `-V` | Show generated JavaScript code |

## References

For writing guidance, see the `references/` directory:
- `claims_guide.md` — How to write effective claims (jurisdiction-neutral)
- `cn_invention_guide.md` — Chinese invention patent writing guide
- `cn_utility_model_guide.md` — Chinese utility model writing guide
- `cn_design_guide.md` — Chinese design patent writing guide
- `en_patent_guide.md` — English patent writing guide (USPTO/EPO/PCT)
