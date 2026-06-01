#!/usr/bin/env python3
"""Orchestrate the complete DOCX formatting workflow.

Pipeline: extract template → apply template → review output.
Supports iterative mode: loop apply → review → fix → re-apply until PASS.

Usage:
    python scripts/pipeline.py reference.docx target.docx
    python scripts/pipeline.py ref.docx target.docx --template 北交模板 --paper 论文名
    python scripts/pipeline.py ref.docx target.docx --iterate

Dependencies:
    - docx skill (~/.claude/skills/docx/)
    - docx-template-extract skill
    - docx-template-apply skill
    - docx-template-review skill
"""

import argparse
import json
import subprocess
import sys
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent
_SKILLS_DIR = _PROJECT_ROOT / ".claude" / "skills"

EXTRACT_SCRIPT = _SKILLS_DIR / "docx-template-extract" / "scripts" / "extract_template.py"
APPLY_SCRIPT = _SKILLS_DIR / "docx-template-apply" / "scripts" / "apply_template.py"
REVIEW_SCRIPT = _SKILLS_DIR / "docx-template-review" / "scripts" / "review_template.py"


def _check_skills():
    missing = []
    for name, path in [("docx-template-extract", EXTRACT_SCRIPT),
                       ("docx-template-apply", APPLY_SCRIPT),
                       ("docx-template-review", REVIEW_SCRIPT)]:
        if not path.exists():
            missing.append(name)
    if missing:
        print(f"Error: required skills not found: {', '.join(missing)}", file=sys.stderr)
        sys.exit(1)


def _detect_next_major(paper: str | None, template: str) -> int:
    """Detect the highest major version number in outputs dir."""
    if paper:
        outputs_dir = _PROJECT_ROOT / "outputs" / paper / template
    else:
        outputs_dir = _PROJECT_ROOT / "outputs" / template
    if not outputs_dir.exists():
        return 1
    max_major = 0
    for d in outputs_dir.iterdir():
        if d.is_dir() and d.name.startswith("v"):
            parts = d.name[1:].split(".", 1)
            try:
                major = int(parts[0])
                max_major = max(max_major, major)
            except ValueError:
                continue
    return max_major + 1


def _run_phase(label: str, cmd: list[str]) -> str:
    """Run a subprocess phase, print output, return stdout."""
    print(f"\n── {label} ──")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr, end="")
    if result.returncode != 0:
        print(f"FAILED: {label} (exit code {result.returncode})", file=sys.stderr)
        sys.exit(result.returncode)
    return result.stdout


def _apply_once(tgt_path: Path, template_name: str, paper_name: str,
                version: str | None, preserve_formatting: bool,
                apply_body_formats: bool) -> tuple[str, Path]:
    """Run apply phase, return (actual_version_str, output_path)."""
    apply_cmd = [
        sys.executable, str(APPLY_SCRIPT), str(tgt_path),
        "--category", template_name, "--paper", paper_name,
    ]
    if version:
        apply_cmd.extend(["--version", str(version)])
    if preserve_formatting:
        apply_cmd.append("--preserve-formatting")
    if apply_body_formats:
        apply_cmd.append("--apply-body-formats")
    _run_phase("Applying template", apply_cmd)

    # Read back the actual version from output dirs (apply_template creates it)
    output_base = _PROJECT_ROOT / "outputs" / paper_name / template_name
    if version:
        actual_version = version
    else:
        # Detect from apply_template output: find highest dir
        actual_version = _detect_actual_version_str(output_base)
    output_dir = output_base / f"v{actual_version}"
    output_path = output_dir / f"{tgt_path.stem}-formatted.docx"
    print(f"Output: {output_path}")
    return actual_version, output_path


def _detect_actual_version_str(output_base: Path) -> str:
    """Detect the highest version string (e.g. '1.2') in output dir."""
    if not output_base.exists():
        return "1.1"
    max_major = 0
    max_minor = 0
    for d in output_base.iterdir():
        if d.is_dir() and d.name.startswith("v"):
            parts = d.name[1:].split(".", 1)
            try:
                major = int(parts[0])
                minor = int(parts[1]) if len(parts) > 1 else 1
                if major > max_major or (major == max_major and minor > max_minor):
                    max_major = major
                    max_minor = minor
            except ValueError:
                continue
    if max_major == 0:
        return "1.1"
    return f"{max_major}.{max_minor}"


def _review_once(output_path: Path, template_name: str, paper_name: str,
                 output_dir: Path) -> bool:
    """Run review phase, return True if PASS (no FAIL)."""
    review_report = output_dir / "review-report.txt"
    # Use --json to parse status programmatically
    review_cmd = [
        sys.executable, str(REVIEW_SCRIPT), str(output_path),
        "--category", template_name,
        "--output", str(review_report),
        "--json",
    ]
    stdout = _run_phase("Reviewing output", review_cmd)

    # Parse JSON result to check PASS/FAIL
    has_fail = True  # default to fail-safe
    try:
        data = json.loads(stdout)
        for entry in data:
            if entry.get("category") == "overall":
                has_fail = entry.get("status") == "FAIL"
                break
    except (json.JSONDecodeError, IndexError):
        pass  # fallback: assume FAIL

    # Also print human-readable report
    if review_report.exists():
        print(review_report.read_text(encoding="utf-8"))

    return not has_fail


def run_pipeline(
    reference_file: str,
    target_file: str,
    *,
    template_name: str | None = None,
    paper_name: str | None = None,
    version: str | None = None,
    skip_extract: bool = False,
    skip_review: bool = False,
    preserve_formatting: bool = False,
    apply_body_formats: bool = False,
    iterate: bool = False,
    max_iterations: int = 10,
):
    """Run the extract → apply → review pipeline, optionally iterating."""
    _check_skills()

    ref_path = Path(reference_file).resolve()
    tgt_path = Path(target_file).resolve()

    if not ref_path.exists():
        print(f"Error: reference not found: {ref_path}", file=sys.stderr)
        sys.exit(1)
    if not tgt_path.exists():
        print(f"Error: target not found: {tgt_path}", file=sys.stderr)
        sys.exit(1)

    # Default names from file stems
    if template_name is None:
        template_name = ref_path.stem
    if paper_name is None:
        paper_name = tgt_path.stem

    print(f"=== docx-pipeline ===")
    print(f"Template: {template_name}")
    print(f"Paper:    {paper_name}")
    if version:
        print(f"Version:  {version}")
    print(f"Reference: {ref_path}")
    print(f"Target:   {tgt_path}")

    # Phase 1: Extract
    if not skip_extract:
        styles_dir = _PROJECT_ROOT / "styles" / template_name
        styles_dir.mkdir(parents=True, exist_ok=True)
        template_output = styles_dir / f"{template_name}.yaml"
        extract_cmd = [
            sys.executable, str(EXTRACT_SCRIPT), str(ref_path),
            "--category", template_name,
            "--output", str(template_output),
        ]
        _run_phase("Phase 1: Extracting template", extract_cmd)
    else:
        print("\n── Phase 1: Skipped (--skip-extract) ──")

    # Phase 2+3: Apply → Review (possibly iterative)
    if iterate:
        # Find the styles template for user reference
        styles_dir = _PROJECT_ROOT / "styles" / template_name
        yaml_files = sorted(styles_dir.glob("*.yaml")) + sorted(styles_dir.glob("*.yml"))
        tpl_path = yaml_files[0] if yaml_files else styles_dir / "*.yaml"

        current_major = _detect_next_major(paper_name, template_name)
        # version string: v{major}.1 for the first iteration of this pipeline run
        current_version = version if version else f"{current_major}.1"

        for iteration in range(1, max_iterations + 1):
            print(f"\n{'='*50}")
            print(f"Iteration {iteration}/{max_iterations} (version v{current_version})")
            print(f"{'='*50}")

            # Apply
            actual_ver, output_path = _apply_once(
                tgt_path, template_name, paper_name,
                current_version, preserve_formatting, apply_body_formats,
            )

            # Review
            output_base = _PROJECT_ROOT / "outputs" / paper_name / template_name
            output_dir = output_base / f"v{actual_ver}"
            passed = _review_once(output_path, template_name, paper_name, output_dir)

            if passed:
                print(f"\n✅ Review PASSED — formatting is correct!")
                print(f"Final output: {output_path}")
                return

            # Review failed — prepare for next iteration
            # Increment minor version: 1.1 → 1.2 → 1.3
            ver_parts = current_version.split(".", 1)
            current_major = int(ver_parts[0])
            current_minor = int(ver_parts[1]) if len(ver_parts) > 1 else 1
            next_minor = current_minor + 1
            print(f"\n❌ Review found issues (iteration {iteration}).")
            print(f"\nEdit the template and press Enter to re-apply as v{current_major}.{next_minor}.")
            print(f"Template file: {tpl_path}")
            print(f"Current output: {output_path}")

            try:
                input("Press Enter when ready (or Ctrl+C to stop)... ")
            except (EOFError, KeyboardInterrupt):
                print("\nStopped.")
                sys.exit(1)

            current_version = f"{current_major}.{next_minor}"

        print(f"\nReached max iterations ({max_iterations}). Stopping.")
        print(f"Last output: {output_path}")
    else:
        # Single pass mode
        actual_ver, output_path = _apply_once(
            tgt_path, template_name, paper_name, version, preserve_formatting,
            apply_body_formats,
        )

        if not skip_review:
            output_dir = _PROJECT_ROOT / "outputs" / paper_name / template_name / f"v{actual_ver}"
            _review_once(output_path, template_name, paper_name, output_dir)

        print(f"\nPipeline complete.")
        print(f"Output: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="DOCX formatting pipeline: extract → apply → review"
    )
    parser.add_argument("reference_docx", help="Reference Word document to extract template from")
    parser.add_argument("target_docx", help="Target DOCX to format")
    parser.add_argument("--template", "-t",
                        help="Template name (default: stem of reference docx)")
    parser.add_argument("--paper", "-p",
                        help="Paper/result name (default: stem of target docx)")
    parser.add_argument("--version", "-v", type=str, default=None,
                        help="Version string in major.minor format (default: auto-detect next)")
    parser.add_argument("--skip-extract", action="store_true",
                        help="Skip extraction; use existing template")
    parser.add_argument("--skip-review", action="store_true",
                        help="Skip review phase")
    parser.add_argument("--preserve-formatting", action="store_true",
                        help="Pass through to apply step")
    parser.add_argument("--apply-body-formats", action="store_true",
                        help="Apply body_paragraph_formats by paragraph index. "
                             "Use only for same-structure reference/target documents.")
    parser.add_argument("--iterate", action="store_true",
                        help="Iterative mode: loop apply→review until PASS. "
                             "Pauses after each iteration for template editing.")
    parser.add_argument("--max-iterations", type=int, default=10,
                        help="Max iterations in iterate mode (default: 10)")
    args = parser.parse_args()

    run_pipeline(
        args.reference_docx,
        args.target_docx,
        template_name=args.template,
        paper_name=args.paper,
        version=args.version,
        skip_extract=args.skip_extract,
        skip_review=args.skip_review,
        preserve_formatting=args.preserve_formatting,
        apply_body_formats=args.apply_body_formats,
        iterate=args.iterate,
        max_iterations=args.max_iterations,
    )


if __name__ == "__main__":
    main()
