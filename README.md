# docx-polar

DOCX template formatting pipeline — extract, apply, review, and citation-check formatting for Chinese academic journal submissions.

Extracts formatting templates from reference `.doc`/`.docx` files, applies semantic formatting rules to target documents, validates the results, and verifies that all citations are real via CrossRef.

## Structure

```
├── styles/               # Extracted formatting templates (YAML)
├── templates/            # Reference template documents
├── outputs/              # Processed output documents
├── tests/                # Unit tests
├── skills/               # Agent skill definitions (mirror of .claude/skills/)
└── .claude/skills/       # Agent skill definitions
```

## Workflow

1. **Extract** — capture formatting from a reference template into a YAML style file
2. **Apply** — restyle a target document using the extracted YAML template
3. **Review** — validate that the output matches expected formatting
4. **Citation Check** — verify all references are real publications via CrossRef

## Usage with AI Code CLI

This project is designed to work with AI-powered code assistants like **Claude Code** and **Codex CLI**. The skills in `.claude/skills/` define the operations these assistants can perform.

### Claude Code

```bash
# Install Claude Code (if not already installed)
npm install -g @anthropic-ai/claude-code

# Run the full formatting pipeline
claude -p "run the docx pipeline on reference.docx and target.docx"

# Format with iterative refinement
claude -p "pipeline this docx: extract template from ref.docx, apply to paper.docx, iterate until formatting passes"

# Check citations only
claude -p "check if all citations in paper.docx are real"

# Full workflow with citation verification
claude -p "process this paper: format paper.docx like template.docx, then verify all citations are real"
```

### Codex CLI

```bash
# Install Codex CLI (if not already installed)
npm install -g @openai/codex

# Run pipeline
codex "run the docx formatting pipeline on ref.docx and paper.docx"

# Check citations
codex "verify citations in paper.docx against CrossRef"
```

### Available Agent Skills

| Skill | Description |
|-------|-------------|
| `docx-write` | Generate a new DOCX from a structured YAML/JSON specification |
| `docx-pipeline` | Orchestrate the full workflow: generate → extract → apply → review → citation-check |
| `docx-template-extract` | Extract formatting template from a reference DOCX |
| `docx-template-apply` | Apply formatting template to a target DOCX |
| `docx-template-review` | Review formatted output against the template |
| `docx-citation-claim` | Verify citations against CrossRef API |

Each skill has its own `SKILL.md` with detailed usage, CLI reference, and examples.

## Python Usage

```bash
# Install dependencies
pip install defusedxml pyyaml requests

# Run the pipeline directly (without AI CLI)
cd .claude/skills/docx-pipeline
python scripts/pipeline.py reference.docx target.docx

# Run tests
python -m pytest tests/
```
