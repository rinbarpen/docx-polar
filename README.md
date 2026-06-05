# docx-polar

DOCX template formatting pipeline — extract, apply, and review formatting for Chinese academic journal submissions.

Extracts formatting templates from reference `.doc`/`.docx` files, applies semantic formatting rules to target documents, and validates the results.

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

## Usage

```python
# Run tests
python -m pytest tests/
```
