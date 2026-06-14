---
name: omni-sheet
description: "AI-driven table generation for academic papers. Generates experiment result tables, comparison tables, and structured data tables with pre-configured style schemes for different journals and conferences. Trigger: 'generate table', 'create comparison table', 'omni-sheet'."
---

# Omni-Sheet

AI-driven table generation for academic papers. Generates structured table data (columns + rows) that can be embedded into DOCX documents.

## Table Types

| Type | Description |
|------|-------------|
| `results` | Experiment results table with metrics |
| `comparison` | Method comparison table |
| `parameters` | Parameter/configuration table |
| `ablation` | Ablation study results |
| `dataset` | Dataset statistics table |

## Schemes

Pre-configured style presets in `schemes/`:

- `default.yaml` — Clean modern tabular style
- `bjtu_journal.yaml` — BJTU学报 three-line table, SimSun 9pt

## Usage

```bash
python scripts/generate_table.py \
  --description "Ablation study results for modality fusion" \
  --columns "Model,P,R,mAP" --rows 4 --scheme bjtu_journal
```

## See Also

- [omni-draw](../omni-draw/SKILL.md) — AI figure generation
- [docx-write](../docx-write/SKILL.md) — Uses omni-sheet for table insertion
