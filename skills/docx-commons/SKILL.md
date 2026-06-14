---
name: docx-commons
description: "Internal shared library for DOCX skills. Provides XML helpers, unpack/pack wrappers, CheckResult dataclass, template utilities, and semantic format rule matching. Not user-invokable — imported by docx-pipeline, docx-review, docx-write, omni-draw, and omni-sheet."
internal: true
---

# DOCX Commons

Internal shared library. All other docx skills import from `docx_helpers` and `template_utils` here.

## Modules

- `docx_helpers.py` — unpack/pack wrappers, XML parsing helpers, CheckResult/ReviewReport/FixInstruction dataclasses
- `template_utils.py` — semantic format rules, font matching, property comparison, run/paragraph property extraction

## Usage

```python
from docx_commons.scripts.docx_helpers import unpack_docx, parse_xml, CheckResult
from docx_commons.scripts.template_utils import (
    extract_run_props, extract_para_props, compare_props,
    default_semantic_format_rules, select_semantic_rule,
    font_matches, text_has_cjk, text_has_latin,
)
```
