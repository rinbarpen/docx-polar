---
name: docx-pipeline
description: "Orchestrate the complete DOCX formatting workflow: extract template from reference DOCX → apply to target DOCX → review the result — all in one command. Trigger when the user says 'pipeline this docx', 'format and validate', 'extract apply review', 'run the full workflow', 'process this paper', or wants end-to-end DOCX formatting."
---

# DOCX Pipeline

Run the complete formatting pipeline — extract, apply, and review — in a single command. The pipeline delegates to the individual skills:

1. **Extract** — calls `docx-template-extract` to create a YAML template from the reference Word document
2. **Apply** — calls `docx-template-apply` to format the target DOCX using the template
3. **Review** — calls `docx-template-review` to validate the output and save a report

## Prerequisites

- **docx skill** (required): Provides `unpack.py` and `pack.py`
- **docx-template-extract** skill
- **docx-template-apply** skill
- **docx-template-review** skill (for review phase)
- **Python packages**: `defusedxml`, `pyyaml`

```bash
pip install defusedxml pyyaml
```

## Workflow

### Basic pipeline

Run the full workflow with default naming (template name = reference filename, paper name = target filename):

```bash
cd /home/rczx/workspace/rinbarpen/work/docx-polar/.claude/skills/docx-pipeline

python scripts/pipeline.py reference.docx target.docx
```

### With explicit template and paper names

```bash
python scripts/pipeline.py 北交模板.docx target.docx \
  --template 北交模板 \
  --paper 道路交通场景双级样条网络三维目标检测
```

### Specify version

```bash
python scripts/pipeline.py ref.docx target.docx \
  --template 北交模板 --paper 论文名 --version 2
```

### Skip extract (use existing template)

```bash
python scripts/pipeline.py ref.docx target.docx \
  --template 北交模板 --skip-extract
```

### Skip review (apply only)

```bash
python scripts/pipeline.py ref.docx target.docx \
  --template 北交模板 --skip-review
```

### Preserve existing direct formatting

```bash
python scripts/pipeline.py ref.docx target.docx \
  --template 北交模板 --preserve-formatting
```

### Apply detailed body paragraph profiles by index

Only use this when the reference and target have the same paragraph structure:

```bash
python scripts/pipeline.py ref.docx target.docx \
  --template 北交模板 --apply-body-formats
```

### Iterative mode: apply → review → fix → re-apply until PASS

Use `--iterate` to loop through multiple versions, pausing after each iteration so you can edit the template YAML and re-apply:

```bash
python scripts/pipeline.py ref.docx target.docx \
  --template 北交模板 --paper 论文名 --iterate
```

Each iteration:
1. **Apply** template → `outputs/<paper>/<template>/v<N>/`
2. **Review** the output and parse PASS/FAIL programmatically
3. If **FAIL**: print the issues, show the template path, prompt user to edit → press Enter → next version (vN+1)
4. If **PASS**: print success and exit

Control max iterations:

```bash
python scripts/pipeline.py ref.docx target.docx \
  --template 北交模板 --iterate --max-iterations 20
```

This workflow allows you to:
1. Run pipeline once → get v1 + review report
2. Read the review report to see what's wrong (e.g., "Heading1 font mismatch")
3. Edit the template YAML in `styles/<template>/` to fix
4. Press Enter → pipeline auto-creates v2 with updated template
5. Repeat until review passes

## Pipeline Output

```
=== docx-pipeline ===
Template: 北交模板
Paper:    道路交通场景双级样条网络三维目标检测
Reference: /home/.../ref.docx
Target:    /home/.../target.docx

── Phase 1/3: Extracting template ──
Unpacking ref.docx...
  Extracted 37 paragraph styles
  Extracted 11 character styles
  Extracted page layout: 11906x16838
  Extracted 3 numbering definitions
  Extracted theme: Custom
  Extracted 9 declared fonts
Template written to styles/北交模板/ref-template.yaml

── Phase 2/3: Applying template ──
...

── Phase 3/3: Reviewing output ──
=== DOCX Template Review Report ===
...
Overall: PASS
Review report: outputs/论文名/北交模板/v1/review-report.txt

Pipeline complete.
Output: outputs/论文名/北交模板/v1/target-formatted.docx
```

## CLI Reference

| Flag | Default | Description |
|------|---------|-------------|
| `--template`, `-t` | reference docx stem | Template name (subdirectory under `styles/`) |
| `--paper`, `-p` | target docx stem | Paper/article name (subdirectory under `outputs/`) |
| `--version`, `-v` | auto-detect | Version number for output |
| `--skip-extract` | off | Skip extraction; use existing template in `styles/` |
| `--skip-review` | off | Skip review phase |
| `--preserve-formatting` | off | Pass through to apply step |
| `--apply-body-formats` | off | Pass through to apply step; applies `body_paragraph_formats` by paragraph index for same-structure documents only |
| `--iterate` | off | Iterative mode: loop apply→review until PASS |
| `--max-iterations` | 10 | Max iterations in iterate mode |

## Dependencies

- `~/.claude/skills/docx/scripts/office/unpack.py` — DOCX unpacking
- `~/.claude/skills/docx/scripts/office/pack.py` — DOCX repacking
- Python: `defusedxml`, `pyyaml`

## See Also

- [docx-template-extract](../docx-template-extract/SKILL.md) — Extract templates from reference documents
- [docx-template-apply](../docx-template-apply/SKILL.md) — Apply template to target documents
- [docx-template-review](../docx-template-review/SKILL.md) — Review formatted output against template
- [docx](../../../.claude/skills/docx/SKILL.md) — General DOCX creation and editing
