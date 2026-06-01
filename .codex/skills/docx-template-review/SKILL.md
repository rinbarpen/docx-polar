---
name: docx-template-review
description: "Validate a formatted DOCX against a template YAML. Checks page layout, document defaults, paragraph/character styles, theme colors/fonts, font table, numbering, and body direct formatting. Trigger when the user says 'review docx', 'validate formatting', 'check template application', 'verify DOCX style', or wants to confirm formatted output matches the reference template."
---

# DOCX Template Review

Validate a formatted DOCX document against its template YAML. This skill checks that all template styles, layout, and formatting have been properly applied, producing a structured pass/warn/fail report.

If the template includes `body_paragraph_formats` with `body_paragraph_formats_mode: index`, review compares the paragraph-indexed direct formatting in `document.xml` as well. Templates extracted in the default `style_enrichment` mode use those profiles to improve styles and avoid destructive cleanup, but do not compare sample-template paragraphs against unrelated manuscript paragraphs.

If the template includes `semantic_format_rules`, review is strict: every matched manuscript paragraph must use the exact role font family, East Asian font, font size, bold/italic state, and selected paragraph properties. Any mismatch is `FAIL`, not `WARN`.

## Prerequisites

- **docx skill** (required): Provides `unpack.py` for DOCX manipulation
- **Python packages**: `defusedxml`, `pyyaml`

```bash
pip install defusedxml pyyaml
```

## Workflow

### Basic review

```bash
cd /home/rczx/workspace/rinbarpen/work/docx-polar/.claude/skills/docx-template-review

python scripts/review_template.py path/to/formatted.docx path/to/template.yaml
```

### Review with category auto-lookup

```bash
# Auto-locate template from styles/北交模板/
python scripts/review_template.py path/to/formatted.docx --category 北交模板
```

### Save report to file

```bash
python scripts/review_template.py formatted.docx --category 北交模板 \
  --output outputs/论文名/北交模板/v1/review-report.txt
```

### Verbose mode (show full details)

```bash
python scripts/review_template.py formatted.docx --category 北交模板 --verbose
```

### JSON output (machine-readable)

```bash
python scripts/review_template.py formatted.docx --category 北交模板 --json
```

## Report Format

```
=== DOCX Template Review Report ===
File: outputs/my-paper/北交模板/v1/target-formatted.docx
Template: styles/北交模板/template.yaml

PAGE LAYOUT ............... PASS  width=11906 ✓, height=16838 ✓, margins ✓
DOCUMENT DEFAULTS ......... PASS  font=Liberation Serif ✓
PARAGRAPH STYLES .......... WARN  40/40 exist, 2 mismatches
  - 'Heading1'.run.font: expected 'Calibri', got 'Times New Roman'
CHARACTER STYLES .......... PASS  11/11
THEME COLORS .............. PASS  12/12
THEME FONTS ............... PASS  major ✓, minor ✓
FONT TABLE ................ PASS  9/9
NUMBERING ................. WARN  2/3 found (missing: id=3)
DIRECT FORMATTING ......... WARN  47 instances (42 run-level, 5 para-level)
BODY PARAGRAPH FORMATS .... PASS  86/86 paragraph profile(s) match

Overall: PASS
```

## Checks

| Check | What It Validates | Severity |
|-------|------------------|----------|
| **Page Layout** | Page dimensions (width/height) and margins match template | PASS/FAIL |
| **Document Defaults** | Default font, size, color, paragraph spacing from `w:docDefaults` | PASS/FAIL |
| **Paragraph Styles** | All template paragraph styles exist in output + property match | PASS/WARN/FAIL |
| **Character Styles** | All template character styles exist in output + property match | PASS/WARN/FAIL |
| **Theme Colors** | Color scheme hex values (accent colors, hyperlink, etc.) | PASS/FAIL |
| **Theme Fonts** | Major/minor font scheme typefaces | PASS/FAIL |
| **Font Table** | All required fonts declared in `fontTable.xml` | PASS/WARN |
| **Numbering** | All numbered/bullet list definitions present | PASS/WARN |
| **Body Paragraph Formats** | Direct paragraph/run formatting captured from reference body matches by paragraph index | PASS/FAIL/SKIP |
| **Semantic Format Rules** | Role-based actual manuscript fonts/sizes match exactly, including title/author/abstract/headings/captions/references/body/table cells | PASS/FAIL |
| **Direct Formatting** | Count of hardcoded font/size/color/bold overrides in body | WARN only |

### Edge Cases

| Scenario | Handling |
|----------|----------|
| Template has no `page` key | Skip page layout check |
| Template style has `run: {}` (empty) | Skip property comparison — "any value acceptable" |
| Style exists in template but missing in output | FAIL with missing style ID |
| `--preserve-formatting` was used | Direct formatting WARN is expected |
| Output has no `numbering.xml` | All numbering entries reported as FAIL |

## CLI Reference

| Flag | Default | Description |
|------|---------|-------------|
| `--category`, `-c` | (none) | Template name; auto-lookup in `styles/<name>/` |
| `--output`, `-o` | (none) | Write report to file (also prints to terminal) |
| `--json` | off | Output report as JSON array |
| `--verbose`, `-v` | off | Show full diff details even on PASS |

## Dependencies

- `~/.claude/skills/docx/scripts/office/unpack.py` — DOCX unpacking
- Python: `defusedxml`, `pyyaml`

## See Also

- [docx-template-extract](../docx-template-extract/SKILL.md) — Extract templates from reference documents
- [docx-template-apply](../docx-template-apply/SKILL.md) — Apply template to target documents
- [docx-pipeline](../docx-pipeline/SKILL.md) — Run extract → apply → review in one command
- [docx](../../../.claude/skills/docx/SKILL.md) — General DOCX creation and editing
