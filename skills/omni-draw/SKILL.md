---
name: omni-draw
description: "AI-driven figure and chart generation for academic papers. Generates data visualizations (bar, line, scatter), flowcharts, architecture diagrams, and comparison charts using matplotlib. Supports pre-configured style schemes for different journals and conferences. Trigger: 'generate figure', 'draw chart', 'create diagram', 'omni-draw'."
---

# Omni-Draw

AI-driven figure generation for academic papers. Uses matplotlib for precise, publication-quality charts.

## Supported Chart Types

| Type | Description |
|------|-------------|
| `bar_chart` | Vertical/horizontal bar chart |
| `line_chart` | Line chart with markers |
| `scatter` | Scatter plot |
| `flowchart` | Process flowchart |
| `architecture_diagram` | System/network architecture diagram |
| `comparison_chart` | Side-by-side comparison bar chart |
| `boxplot` | Box and whisker plot |

## Schemes

Pre-configured style presets in `schemes/`:

- `default.yaml` — Clean, modern academic style
- `bjtu_journal.yaml` — BJTU学报 grayscale, SimSun 9pt, 300 DPI

## Usage

```bash
python scripts/generate_chart.py \
  --type bar --title "实验结果对比" \
  --data '{"labels":["A","B","C"],"values":[95.2,87.1,78.5]}' \
  --scheme bjtu_journal --output figures/result.png
```

## See Also

- [omni-sheet](../omni-sheet/SKILL.md) — AI table generation
- [docx-write](../docx-write/SKILL.md) — Uses omni-draw for figure insertion
