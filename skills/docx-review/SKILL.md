---
name: docx-review
description: "Unified DOCX review dispatcher. Supports content (images+tables), style (formatting validation), and citation (CrossRef verification). Trigger: 'review docx', 'docx-review:content', 'docx-review:style', 'docx-review:citation'."
---

# DOCX Review

Unified dispatcher for three review types on formatted DOCX documents. Call as `docx-review:<type>` or invoke directly with the sub-type as an argument.

## Sub-types

| Invocation | Script | Purpose |
|---|---|---|
| `docx-review:content` | `scripts/review_content.py` | Image/table content review — DPI, alt-text, numbering, cross-references |
| `docx-review:style` | `scripts/review_style.py` | Style/formatting validation against template YAML |
| `docx-review:citation` | `scripts/review_citation.py` | CrossRef citation verification |

## Unified Interface

```bash
# Run all review types
python scripts/review_dispatcher.py formatted.docx --template-file template.yaml

# Run specific types
python scripts/review_dispatcher.py formatted.docx --type content,style --category 北交模板

# JSON output
python scripts/review_dispatcher.py formatted.docx --type all --json

# With thresholds
python scripts/review_dispatcher.py formatted.docx \
  --type citation --citation-threshold 0.7 \
  --type content --dpi-threshold 300
```

## Individual Script Usage

### Content Review
```bash
python scripts/review_content.py formatted.docx
python scripts/review_content.py formatted.docx --json --dpi-threshold 150
```

### Style Review
```bash
python scripts/review_style.py formatted.docx --template-file template.yaml
python scripts/review_style.py formatted.docx --category 北交模板 --json
```

### Citation Review
```bash
python scripts/review_citation.py formatted.docx --threshold 0.7
python scripts/review_citation.py formatted.docx --output report.json
```

## CLI Reference (dispatcher)

| Flag | Default | Description |
|------|---------|-------------|
| `--type`, `-t` | all | Review types: content,style,citation (comma-separated) |
| `--template-file` | none | Template YAML for style review |
| `--category`, `-c` | none | Template name for auto-lookup in styles/ |
| `--dpi-threshold` | 300 | Minimum DPI for print quality |
| `--citation-threshold` | 0.6 | Minimum CrossRef confidence score |
| `--output`, `-o` | none | Write report to file |
| `--json` | off | Output as JSON |
| `--verbose`, `-v` | off | Show full details |

## Programmatic API

```python
from review_dispatcher import review_docx
from review_content import review_content
from review_style import review_style
from review_citation import review_citations

# Unified
report = review_docx("doc.docx", template_file="tmpl.yaml")
print(report.overall_pass)

# Individual
results = review_content("doc.docx", dpi_threshold=300)
results = review_style("doc.docx", template_file="tmpl.yaml")
citations = review_citations("doc.docx", threshold=0.6)
```

## Dependencies

- Python: `defusedxml`, `pyyaml`, `Pillow` (optional)
- `docx skill`: `~/.claude/skills/docx/scripts/office/unpack.py`

## See Also

- [docx-pipeline](../docx-pipeline/SKILL.md) — end-to-end formatting pipeline
- [docx-write](../docx-write/SKILL.md) — content generation and template application
