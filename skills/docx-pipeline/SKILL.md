---
name: docx-pipeline
description: "Orchestrate the complete DOCX formatting workflow: create template rules → write content (LLM-driven per chapter) → apply style → review (content+style+citation) → iterate until PASS. Trigger: 'pipeline this docx', 'format and validate', 'generate and review docx', 'run the full workflow'."
---

# DOCX Pipeline

Orchestrate content generation, style application, and review in a single command.

## Workflow

```
pipeline --generate spec.yaml --template 北交模板 --iterate
  │
  ├─ Phase 0: Create template rules (extract from reference or load existing)
  ├─ Phase 1: Write content per chapter (LLM-driven, optional)
  ├─ Phase 2: Apply style template
  ├─ Phase 3: Review (content + style + citation)
  └─ Phase 4: Iterate until PASS or max_iterations
```

## Usage

### Full LLM-driven workflow
```bash
cd /home/rczx/workspace/rinbarpen/work/docx-polar

python skills/docx-pipeline/scripts/pipeline.py \
  --generate paper.yaml --template 北交模板 --iterate
```

### Template-only workflow (preserved)
```bash
python skills/docx-pipeline/scripts/pipeline.py \
  reference.docx target.docx --template 北交模板 --iterate
```

### With chapter template
```bash
python skills/docx-pipeline/scripts/pipeline.py \
  --generate paper.yaml --template 北交模板 \
  --chapter-spec templates/chapters/bjtu_journal.yaml \
  --iterate
```

### Skip phases
```bash
python skills/docx-pipeline/scripts/pipeline.py \
  --generate paper.yaml --template 北交模板 \
  --skip-content --skip-review-citation
```

## CLI Reference

| Flag | Default | Description |
|------|---------|-------------|
| `--generate`, `-g` | none | YAML/JSON spec to generate DOCX from |
| `--template`, `-t` | auto | Template name (subdirectory under `styles/`) |
| `--paper`, `-p` | auto | Paper/article name (subdirectory under `outputs/`) |
| `--version`, `-v` | auto | Version string |
| `--chapter-spec` | none | Chapter template YAML (e.g. `templates/chapters/bjtu_journal.yaml`) |
| `--skip-extract` | off | Skip template extraction |
| `--skip-content` | off | Skip LLM content generation |
| `--skip-style` | off | Skip template application |
| `--skip-review-content` | off | Skip content review (images/tables) |
| `--skip-review-style` | off | Skip style review (formatting) |
| `--skip-review-citation` | off | Skip citation review (CrossRef) |
| `--preserve-formatting` | off | Keep existing direct formatting |
| `--apply-body-formats` | off | Apply body_paragraph_formats by index |
| `--iterate` | off | Iterative mode: loop write→review until PASS |
| `--max-iterations` | 10 | Max iterations in iterate mode |
| `--citation-threshold` | 0.6 | CrossRef confidence threshold |
| `--content-dpi-threshold` | 300 | Minimum DPI for print quality |
| `--llm-model` | claude-sonnet-4-6 | LLM model for content generation |

## See Also

- [docx-review](../docx-review/SKILL.md) — Unified review (content, style, citation)
- [docx-write](../docx-write/SKILL.md) — Content generation and style application
- [omni-draw](../omni-draw/SKILL.md) — AI figure generation
- [omni-sheet](../omni-sheet/SKILL.md) — AI table generation
