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

import yaml

_SELF_DIR = Path(__file__).resolve().parent  # .../skills/docx-pipeline/scripts/
_SKILLS_DIR = _SELF_DIR.parent.parent  # .../skills/
_PROJECT_ROOT = _SKILLS_DIR.parent  # docx-polar root

EXTRACT_SCRIPT = _SKILLS_DIR / "docx-template-extract" / "scripts" / "extract_template.py"
APPLY_SCRIPT = _SKILLS_DIR / "docx-template-apply" / "scripts" / "apply_template.py"
REVIEW_SCRIPT = _SKILLS_DIR / "docx-template-review" / "scripts" / "review_template.py"
CITATION_SCRIPT = _SKILLS_DIR / "docx-citation-claim" / "scripts" / "check_citations.py"
CONTENT_REVIEW_SCRIPT = _SKILLS_DIR / "docx-review" / "scripts" / "review_content.py"
GENERATE_SCRIPT = _SKILLS_DIR / "docx-write" / "scripts" / "write_docx.py"


def _check_skills(*, check_generate: bool = False):
    missing = []
    for name, path in [("docx-template-extract", EXTRACT_SCRIPT),
                       ("docx-template-apply", APPLY_SCRIPT),
                       ("docx-template-review", REVIEW_SCRIPT),
                       ("docx-citation-claim", CITATION_SCRIPT)]:
        if not path.exists():
            missing.append(name)
    if check_generate and not GENERATE_SCRIPT.exists():
        missing.append("docx-write")
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
        return "1.0"
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
        return "1.0"
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


def _check_citations_once(docx_path: Path, threshold: float = 0.6,
                          output_dir: Path | None = None) -> bool:
    """Run citation check phase, return True if all citations are REAL."""
    citation_output = (output_dir / "citation-report.json") if output_dir else None
    citation_cmd = [
        sys.executable, str(CITATION_SCRIPT), str(docx_path),
        "--threshold", str(threshold),
        "--source", "auto",
    ]
    if citation_output:
        citation_cmd.extend(["--output", str(citation_output)])

    # Run without _run_phase to allow non-zero exit (fake citations)
    print(f"\n── Checking citations ──")
    result = subprocess.run(citation_cmd, capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(result.stderr, file=sys.stderr, end="")

    passed = True
    try:
        report = json.loads(result.stdout.strip())
        passed = report.get("passed", False)
        total = report.get("total", 0)
        real = report.get("real", 0)
        fake = report.get("fake", 0)
        unverified = report.get("unverified", 0)
        errors = report.get("error", 0)
        print(f"\n  Citation Summary: {total} total, {real} REAL, {unverified} UNVERIFIED, {fake} FAKE, {errors} ERROR")
        if fake > 0 or errors > 0:
            print(f"\n  ❌ {fake} fake / {errors} error citations found — needs fixing.")
            for c in report.get("citations", []):
                if c.get("status") in ("FAKE", "ERROR"):
                    print(f"     [{c['index']}] {c['status']}: {c['text'][:80]}...")
        else:
            print(f"\n  ✅ All citations verified as REAL.")
    except (json.JSONDecodeError, KeyError) as e:
        print(f"  Warning: could not parse citation report: {e}", file=sys.stderr)
        passed = False

    return passed


def _review_content_once(output_path: Path, output_dir: Path,
                        dpi_threshold: int = 300) -> bool:
    """Run content review phase, return True if PASS (no FAIL)."""
    if not CONTENT_REVIEW_SCRIPT.exists():
        print("  Warning: docx-review skill not found, skipping content review.",
              file=sys.stderr)
        return True

    content_report = output_dir / "content-review-report.txt"
    cmd = [
        sys.executable, str(CONTENT_REVIEW_SCRIPT), str(output_path),
        "--output", str(content_report),
        "--json",
        "--dpi-threshold", str(dpi_threshold),
    ]
    stdout = _run_phase("Phase 5: Reviewing content (images & tables)", cmd)

    has_fail = True
    try:
        data = json.loads(stdout)
        for entry in data:
            if entry.get("category") == "overall":
                has_fail = entry.get("status") == "FAIL"
                break
    except (json.JSONDecodeError, IndexError):
        pass

    if content_report.exists():
        print(content_report.read_text(encoding="utf-8"))

    return not has_fail


def _fmt(val):
    """Format a value for table display."""
    if val is None:
        return "-"
    if isinstance(val, bool):
        return "Yes" if val else "No"
    return str(val)


def _dxa_to_mm(dxa):
    """Convert DXA (twentieth of a point) to approximate mm."""
    if dxa is None:
        return "-"
    try:
        return f"{float(str(dxa).replace('pt','')) * 25.4 / 72:.1f}mm"
    except (ValueError, TypeError):
        return str(dxa)


def _show_template_review(template_path: Path):
    """Display extracted template in table format for human review."""
    if not template_path.exists():
        print(f"\n  Template file not found: {template_path}")
        return

    with open(template_path, "r", encoding="utf-8") as f:
        tpl = yaml.safe_load(f)

    if not tpl:
        print("\n  Empty template.")
        return

    print(f"\n{'=' * 70}")
    print("  Template Review")
    print(f"{'=' * 70}")

    # Page Layout
    page = tpl.get("page", {})
    if page:
        m = page.get("margins", {})
        print('\n  [Page Layout]')
        sep = "  " + "─" * 60
        print(sep)
        print(f"  {'Size':<20} {page.get('width', '-')} x {page.get('height', '-')} DXA")
        if m:
            print(f"  {'Margins':<20} T:{_dxa_to_mm(m.get('top'))}  B:{_dxa_to_mm(m.get('bottom'))}  "
                  f"L:{_dxa_to_mm(m.get('left'))}  R:{_dxa_to_mm(m.get('right'))}")

    # Document Defaults
    dd = tpl.get("document_defaults", {})
    if dd:
        dr = dd.get("run", {})
        dp = dd.get("paragraph", {})
        print('\n  [Document Defaults]')
        print(sep)
        if dr:
            print(f"  {'Font':<20} {_fmt(dr.get('font'))}")
            print(f"  {'Font (EA)':<20} {_fmt(dr.get('fontEastAsia'))}")
            sz = dr.get('fontSize')
            try:
                pt_display = f"({int(int(sz) / 2)}pt)" if sz else ""
            except (ValueError, TypeError):
                pt_display = ""
            print(f"  {'Font Size':<20} {_fmt(sz)} {pt_display}")
            print(f"  {'Font Color':<20} {_fmt(dr.get('fontColor'))}")
        if dp:
            print(f"  {'Spacing After':<20} {_fmt(dp.get('spacing_after'))} DXA")
            print(f"  {'Alignment':<20} {_fmt(dp.get('alignment'))}")

    # Paragraph Styles
    pstyles = tpl.get("paragraph_styles", [])
    if pstyles:
        print(f'\n  [Paragraph Styles] ({len(pstyles)} total)')
        sep = "  " + "─" * 60
        print(sep)
        print(f"  {'ID':<20} {'Name':<28} {'Font':<18} {'Size':<8} {'EA Font':<18}")
        print(sep)
        for ps in pstyles:
            pid = ps.get("id", "-")
            pname = ps.get("name") or "-"
            r = ps.get("run", {})
            font = r.get("font", "-")
            ea = r.get("fontEastAsia", "-")
            sz = str(r.get("fontSize", "-"))
            print(f"  {pid:<20} {_fmt(pname):<28} {font:<18} {sz:<8} {ea:<18}")

    # Character Styles
    cstyles = tpl.get("character_styles", [])
    if cstyles:
        print(f'\n  [Character Styles] ({len(cstyles)} total)')
        for cs in cstyles:
            cid = cs.get("id", "-")
            cname = cs.get("name")
            display = f"  {cid}" + (f" ({cname})" if cname else "")
            print(display)

    # Fonts
    fonts = tpl.get("fonts", [])
    if fonts:
        print(f'\n  [Font Table] ({len(fonts)} fonts)')
        for f in fonts:
            print(f"  {f}")

    # Numbering
    numbering = tpl.get("numbering", [])
    if numbering:
        print(f'\n  [Numbering Definitions] ({len(numbering)} total)')
        for nd in numbering:
            nid = nd.get("id", "-")
            levels = nd.get("levels", [])
            for lv in levels[:3]:
                lvl = lv.get("level", "-")
                fmt_lv = lv.get("format", "-")
                txt = lv.get("text", "-")
                print(f"  id={nid}  level={lvl}  format={fmt_lv}  text={txt}")

    # Theme
    theme = tpl.get("theme", {})
    if theme:
        colors = theme.get("colors", {})
        fonts_theme = theme.get("fonts", {})
        print('\n  [Theme]')
        if fonts_theme:
            print(f"  {'Major':<20} {_fmt(fonts_theme.get('major'))}")
            print(f"  {'Minor':<20} {_fmt(fonts_theme.get('minor'))}")
        if colors:
            for ck, cv in list(colors.items())[:8]:
                print(f"  {ck:<20} #{cv}")


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
    citation_check: bool = False,
    citation_iterate: bool = False,
    citation_threshold: float = 0.6,
    citation_max_iterations: int = 10,
    content_review: bool = False,
    content_dpi_threshold: int = 300,
    generate_spec: str | None = None,
):
    """Run the generate → extract → apply → review → citation-check → content-review pipeline."""
    _check_skills(check_generate=bool(generate_spec))

    # Phase 0: Generate DOCX from spec
    if generate_spec:
        spec_path = Path(generate_spec).resolve()
        if not spec_path.exists():
            print(f"Error: spec file not found: {generate_spec}", file=sys.stderr)
            sys.exit(1)

        if paper_name is None:
            paper_name = spec_path.stem

        gen_cmd = [
            sys.executable, str(GENERATE_SCRIPT), str(spec_path),
            "--skip-validate",
        ]
        if template_name:
            gen_cmd.extend(["--category", template_name])
        if paper_name:
            gen_cmd.extend(["--paper", paper_name])
        if version:
            gen_cmd.extend(["--version", str(version)])

        stdout = _run_phase("Phase 0: Generating DOCX from spec", gen_cmd)
        # Extract output path from stdout
        generated_path = None
        for line in stdout.splitlines():
            if line.startswith("Generated:"):
                generated_path = line.split("Generated:", 1)[-1].strip()
                break

        if not generated_path:
            # Fallback: construct from spec stem
            generated_path = str(spec_path.with_suffix(".docx"))

        target_file = generated_path
        print(f"Generated: {target_file}")

    # Now target_file is resolved
    if not target_file:
        print("Error: no target document specified (use --generate or provide target_docx)", file=sys.stderr)
        sys.exit(1)

    ref_path = Path(reference_file).resolve() if reference_file else None
    tgt_path = Path(target_file).resolve()

    if reference_file and not ref_path.exists():
        print(f"Error: reference not found: {ref_path}", file=sys.stderr)
        sys.exit(1)
    if not tgt_path.exists():
        print(f"Error: target not found: {tgt_path}", file=sys.stderr)
        sys.exit(1)

    # Default names from file stems
    if template_name is None:
        template_name = ref_path.stem if ref_path else paper_name or "default"
    if paper_name is None:
        paper_name = tgt_path.stem

    flags = []
    if generate_spec:
        flags.append("generate")
    if citation_check:
        flags.append("citation-check")
    if citation_iterate:
        flags.append("citation-iterate")
    flags_str = f" [{', '.join(flags)}]" if flags else ""

    print(f"=== docx-pipeline ==={flags_str}")
    print(f"Template: {template_name}")
    print(f"Paper:    {paper_name}")
    if version:
        print(f"Version:  {version}")
    if ref_path:
        print(f"Reference: {ref_path}")
    if generate_spec:
        print(f"Spec:      {generate_spec}")
    print(f"Target:   {tgt_path}")

    # Phase 1: Extract
    if not skip_extract and ref_path:
        styles_dir = _PROJECT_ROOT / "styles" / template_name
        styles_dir.mkdir(parents=True, exist_ok=True)
        template_output = styles_dir / f"{template_name}.yaml"
        extract_cmd = [
            sys.executable, str(EXTRACT_SCRIPT), str(ref_path),
            "--category", template_name,
            "--output", str(template_output),
        ]
        _run_phase("Phase 1: Extracting template", extract_cmd)
    elif not ref_path:
        print("\n── Phase 1: Skipped (no reference, using template if available) ──")
        skip_extract = True
    else:
        print("\n── Phase 1: Skipped (--skip-extract) ──")

    # Human review of template (always, regardless of skip-extract)
    tpl_path_for_review = _PROJECT_ROOT / "styles" / template_name / f"{template_name}.yaml"
    _show_template_review(tpl_path_for_review)
    print(f"\n{'─' * 70}")
    print(f"  Template: {tpl_path_for_review}")
    print(f"{'─' * 70}")
    # Skip interactive prompt in non-interactive mode
    import os
    if sys.stdin.isatty():
        print("\n  Edit the template YAML above if needed, then press Enter to apply.")
        print("  (Ctrl+C to stop)")
        try:
            input()
        except (EOFError, KeyboardInterrupt):
            print("\nStopped.")
            sys.exit(1)
    else:
        print("\n  Non-interactive mode: proceeding automatically.")

    # Phase 2+3: Apply → Review
    if iterate:
        # Find the styles template for user reference
        styles_dir = _PROJECT_ROOT / "styles" / template_name
        yaml_files = sorted(styles_dir.glob("*.yaml")) + sorted(styles_dir.glob("*.yml"))
        tpl_path = yaml_files[0] if yaml_files else styles_dir / "*.yaml"

        current_major = _detect_next_major(paper_name, template_name)
        # version string: v{major}.0 for the first iteration of this pipeline run
        current_version = version if version else f"{current_major}.0"
        formatting_passed = False

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
            formatting_passed = _review_once(output_path, template_name, paper_name, output_dir)

            if formatting_passed:
                print(f"\n✅ Review PASSED — formatting is correct!")
                break

            # Review failed — prepare for next iteration
            ver_parts = current_version.split(".", 1)
            current_major = int(ver_parts[0])
            current_minor = int(ver_parts[1]) if len(ver_parts) > 1 else 1
            next_minor = current_minor + 1
            print(f"\n❌ Review found issues (iteration {iteration}).")
            print(f"\nEdit the template and press Enter to re-apply as v{current_major}.{next_minor}.")
            print(f"Template file: {tpl_path}")
            print(f"Current output: {output_path}")

            if sys.stdin.isatty():
                try:
                    input("Press Enter when ready (or Ctrl+C to stop)... ")
                except (EOFError, KeyboardInterrupt):
                    print("\nStopped.")
                    sys.exit(1)
            else:
                print("Non-interactive: stopping after this iteration.")

            current_version = f"{current_major}.{next_minor}"

        if not formatting_passed:
            print(f"\nReached max iterations ({max_iterations}). Stopping.")
            print(f"Last output: {output_path}")
            return
    else:
        # Single pass mode
        actual_ver, output_path = _apply_once(
            tgt_path, template_name, paper_name, version, preserve_formatting,
            apply_body_formats,
        )

        if not skip_review:
            output_dir = _PROJECT_ROOT / "outputs" / paper_name / template_name / f"v{actual_ver}"
            _review_once(output_path, template_name, paper_name, output_dir)

    # Phase 4: Citation check
    if citation_check:
        output_base = _PROJECT_ROOT / "outputs" / paper_name / template_name
        output_dir = output_base / f"v{actual_ver}"

        print(f"\n{'─' * 70}")
        print("  Phase 4: Checking citations")
        print(f"{'─' * 70}")

        if citation_iterate:
            for cit_iter in range(1, citation_max_iterations + 1):
                print(f"\n{'='*50}")
                print(f"Citation iteration {cit_iter}/{citation_max_iterations}")
                print(f"{'='*50}")

                citation_passed = _check_citations_once(
                    output_path, citation_threshold, output_dir,
                )

                if citation_passed:
                    print(f"\n✅ Citation check PASSED — all references are real!")
                    break

                if cit_iter < citation_max_iterations:
                    print(f"\n❌ Citation check found issues (iteration {cit_iter}).")
                    print(f"\nFix the references in the target DOCX, then save and press Enter to re-check.")
                    print(f"Target file: {tgt_path}")
                    print(f"Formatted output: {output_path}")

                    if sys.stdin.isatty():
                        try:
                            input("Press Enter when ready (or Ctrl+C to stop)... ")
                        except (EOFError, KeyboardInterrupt):
                            print("\nStopped.")
                            sys.exit(1)
                    else:
                        print("Non-interactive: stopping after this iteration.")
                        break
                else:
                    print(f"\nReached max citation iterations ({citation_max_iterations}). Stopping.")

            if not citation_passed:
                print(f"\n⚠️  Citation check did NOT pass all checks.")
        else:
            citation_passed = _check_citations_once(
                output_path, citation_threshold, output_dir,
            )

    # Phase 5: Content review (images & tables)
    if content_review:
        output_base = _PROJECT_ROOT / "outputs" / paper_name / template_name
        output_dir = output_base / f"v{actual_ver}"
        _review_content_once(output_path, output_dir, content_dpi_threshold)

    print(f"\nPipeline complete.")
    print(f"Output: {output_path}")


def main():
    parser = argparse.ArgumentParser(
        description="DOCX formatting pipeline: generate → extract → apply → review"
    )
    parser.add_argument(
        "reference_docx", nargs="?",
        help="Reference Word document to extract template from (optional with --generate)",
    )
    parser.add_argument(
        "target_docx", nargs="?",
        help="Target DOCX to format (optional with --generate)",
    )
    parser.add_argument(
        "--generate", "-g",
        help="YAML/JSON spec to generate DOCX from (Phase 0)",
    )
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
    parser.add_argument("--citation-check", action="store_true",
                        help="Enable citation checking phase (Phase 4)")
    parser.add_argument("--skip-citation-check", action="store_true",
                        help="Skip citation checking phase")
    parser.add_argument("--citation-iterate", action="store_true",
                        help="Iterative mode for citation checking: loop check→fix→re-check until PASS")
    parser.add_argument("--citation-threshold", type=float, default=0.6,
                        help="Minimum confidence score (0.0–1.0) for REAL citation (default: 0.6)")
    parser.add_argument("--citation-max-iterations", type=int, default=10,
                        help="Max iterations in citation iterate mode (default: 10)")
    parser.add_argument("--content-review", action="store_true",
                        help="Enable content review phase (Phase 5) for images and tables")
    parser.add_argument("--content-dpi-threshold", type=int, default=300,
                        help="Minimum DPI for print-quality images (default: 300)")
    args = parser.parse_args()

    if not args.generate and not args.reference_docx:
        parser.error("either --generate or reference_docx is required")
    if not args.generate and not args.target_docx:
        parser.error("either --generate or target_docx is required")

    citation_enabled = args.citation_check and not args.skip_citation_check

    run_pipeline(
        args.reference_docx or "",
        args.target_docx or "",
        template_name=args.template,
        paper_name=args.paper,
        version=args.version,
        skip_extract=args.skip_extract,
        skip_review=args.skip_review,
        preserve_formatting=args.preserve_formatting,
        apply_body_formats=args.apply_body_formats,
        iterate=args.iterate,
        max_iterations=args.max_iterations,
        citation_check=citation_enabled,
        citation_iterate=args.citation_iterate,
        citation_threshold=args.citation_threshold,
        citation_max_iterations=args.citation_max_iterations,
        content_review=args.content_review,
        content_dpi_threshold=args.content_dpi_threshold,
        generate_spec=args.generate,
    )


if __name__ == "__main__":
    main()
