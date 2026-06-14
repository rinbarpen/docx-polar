#!/usr/bin/env python3
"""Review images and tables in a formatted DOCX document.

Checks: image count/metadata, DPI, alt-text, figure numbering,
figure cross-references, table structure, table numbering,
table cross-references. Produces a structured pass/warn/fail report.

Usage:
    python review_content.py formatted.docx
    python review_content.py formatted.docx --json
    python review_content.py formatted.docx --output report.txt --verbose

Dependencies:
    - pip install defusedxml Pillow (Pillow optional)
    - ~/.claude/skills/docx/scripts/office/unpack.py
"""

import argparse
import json
import os
import re
import struct
import subprocess
import sys
import tempfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

import defusedxml.minidom

UNPACK_PY = os.path.expanduser("~/.claude/skills/docx/scripts/office/unpack.py")


@dataclass
class CheckResult:
    category: str
    status: str  # PASS | WARN | FAIL | SKIP
    summary: str
    details: list[str] = field(default_factory=list)


def _ensure_docx_skill():
    if not Path(UNPACK_PY).exists():
        print("Error: docx skill not found.", file=sys.stderr)
        sys.exit(1)


def _unpack(input_file: str, output_dir: str) -> str:
    result = subprocess.run(
        [sys.executable, UNPACK_PY, input_file, output_dir,
         "--merge-runs", "false", "--simplify-redlines", "false"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(f"Error unpacking: {result.stderr or result.stdout}", file=sys.stderr)
        sys.exit(1)
    return result.stdout.strip()


def _parse_xml(path: Path):
    return defusedxml.minidom.parseString(path.read_text(encoding="utf-8"))


def _get_attr(elem, attr: str) -> str | None:
    val = elem.getAttribute(attr)
    return val if val else None


def _direct_children(elem, tag: str | None = None) -> list:
    children = [n for n in elem.childNodes if n.nodeType == n.ELEMENT_NODE]
    if tag is None:
        return children
    return [n for n in children if n.tagName == tag]


def _paragraph_text(p) -> str:
    parts = []
    for t in p.getElementsByTagName("w:t"):
        if t.firstChild:
            parts.append(t.firstChild.nodeValue)
    return "".join(parts)


def _emu_to_inches(emu: int) -> float:
    return emu / 914400.0


# --- Image metadata extraction ---

def _get_image_dimensions_pillow(media_path: Path) -> tuple[int, int] | None:
    try:
        from PIL import Image
        with Image.open(media_path) as img:
            return img.size
    except Exception:
        return None


def _get_png_dimensions(media_path: Path) -> tuple[int, int] | None:
    try:
        with open(media_path, "rb") as f:
            header = f.read(24)
            if header[:8] != b"\x89PNG\r\n\x1a\n":
                return None
            return struct.unpack(">II", header[16:24])
    except Exception:
        return None


def _get_image_dimensions(media_path: Path) -> tuple[int, int] | None:
    dims = _get_image_dimensions_pillow(media_path)
    if dims is not None:
        return dims
    if media_path.suffix.lower() == ".png":
        return _get_png_dimensions(media_path)
    return None


def _get_file_size_kb(media_path: Path) -> float:
    try:
        return media_path.stat().st_size / 1024.0
    except OSError:
        return 0.0


def _format_bytes(size_kb: float) -> str:
    if size_kb >= 1024:
        return f"{size_kb / 1024:.1f}MB"
    return f"{size_kb:.0f}KB"


# --- Image checks ---

def check_images(word_dir: Path) -> CheckResult:
    media_dir = word_dir / "media"
    doc_xml = word_dir / "document.xml"
    rels_xml = word_dir / "_rels" / "document.xml.rels"

    if not doc_xml.exists():
        return CheckResult("images", "FAIL", "document.xml not found")

    # Parse relationships to map rId -> media filename
    rid_to_file: dict[str, str] = {}
    if rels_xml.exists():
        rels_dom = _parse_xml(rels_xml)
        for rel in rels_dom.getElementsByTagName("Relationship"):
            rel_type = _get_attr(rel, "Type") or ""
            if "image" in rel_type.lower():
                rid = _get_attr(rel, "Id")
                target = _get_attr(rel, "Target")
                if rid and target:
                    rid_to_file[rid] = target.split("/")[-1]

    dom = _parse_xml(doc_xml)
    drawings = dom.getElementsByTagName("w:drawing")

    if not drawings:
        return CheckResult("images", "PASS", "0 images found", [])

    details = []
    for i, drawing in enumerate(drawings):
        blips = drawing.getElementsByTagName("a:blip")
        rid = None
        for blip in blips:
            rid = _get_attr(blip, "r:embed")
            if rid:
                break

        extents = drawing.getElementsByTagName("wp:extent")
        cx_emu = cy_emu = None
        if extents:
            cx_emu = int(_get_attr(extents[0], "cx") or 0)
            cy_emu = int(_get_attr(extents[0], "cy") or 0)

        doc_prs = drawing.getElementsByTagName("wp:docPr")
        img_name = _get_attr(doc_prs[0], "name") if doc_prs else "unknown"
        img_descr = _get_attr(doc_prs[0], "descr") if doc_prs else None

        media_file = rid_to_file.get(rid, "") if rid else ""
        media_path = media_dir / media_file if media_file else None

        fmt = ""
        size_str = "unknown"
        px_str = "unknown"
        rendered_str = "unknown"
        dpi_str = ""

        if media_path and media_path.exists():
            fmt = media_path.suffix.lower().lstrip(".")
            size_str = _format_bytes(_get_file_size_kb(media_path))
            dims = _get_image_dimensions(media_path)
            if dims:
                pw, ph = dims
                px_str = f"{pw}x{ph}px"
                if cx_emu and cy_emu and cx_emu > 0 and cy_emu > 0:
                    rw_in = _emu_to_inches(cx_emu)
                    rh_in = _emu_to_inches(cy_emu)
                    dpi_w = pw / rw_in if rw_in > 0 else 0
                    dpi_h = ph / rh_in if rh_in > 0 else 0
                    rendered_str = f"{rw_in * 2.54:.1f}x{rh_in * 2.54:.1f}cm"
                    dpi_str = f", {dpi_w:.0f}/{dpi_h:.0f} DPI"

        detail = f"  image[{i}]: {img_name}"
        if fmt:
            detail += f", {fmt}"
        if size_str != "unknown":
            detail += f", {size_str}"
        if px_str != "unknown":
            detail += f", {px_str}"
        if rendered_str != "unknown":
            detail += f", rendered {rendered_str}{dpi_str}"
        if img_descr:
            detail += f", alt-text OK"
        details.append(detail)

    return CheckResult("images", "PASS", f"{len(drawings)} total images", details)


def check_image_dpi(word_dir: Path, dpi_threshold: int = 300) -> CheckResult:
    media_dir = word_dir / "media"
    doc_xml = word_dir / "document.xml"
    rels_xml = word_dir / "_rels" / "document.xml.rels"

    if not doc_xml.exists():
        return CheckResult("image_dpi", "SKIP", "no document.xml")

    rid_to_file: dict[str, str] = {}
    if rels_xml.exists():
        rels_dom = _parse_xml(rels_xml)
        for rel in rels_dom.getElementsByTagName("Relationship"):
            rel_type = _get_attr(rel, "Type") or ""
            if "image" in rel_type.lower():
                rid = _get_attr(rel, "Id")
                target = _get_attr(rel, "Target")
                if rid and target:
                    rid_to_file[rid] = target.split("/")[-1]

    dom = _parse_xml(doc_xml)
    drawings = dom.getElementsByTagName("w:drawing")

    if not drawings:
        return CheckResult("image_dpi", "SKIP", "no images to check", [])

    details = []
    below_threshold = 0
    unknown_count = 0
    ok_count = 0

    for drawing in drawings:
        blips = drawing.getElementsByTagName("a:blip")
        rid = None
        for blip in blips:
            rid = _get_attr(blip, "r:embed")
            if rid:
                break

        extents = drawing.getElementsByTagName("wp:extent")
        if not extents:
            unknown_count += 1
            continue
        cx_emu = int(_get_attr(extents[0], "cx") or 0)
        if cx_emu <= 0:
            unknown_count += 1
            continue

        media_file = rid_to_file.get(rid, "") if rid else ""
        media_path = media_dir / media_file if media_file else None

        doc_prs = drawing.getElementsByTagName("wp:docPr")
        img_name = _get_attr(doc_prs[0], "name") if doc_prs else "unknown"

        if not media_path or not media_path.exists():
            unknown_count += 1
            details.append(f"  {img_name}: media file not found")
            continue

        dims = _get_image_dimensions(media_path)
        if not dims:
            unknown_count += 1
            details.append(f"  {img_name}: could not read pixel dimensions")
            continue

        pw = dims[0]
        rw_in = _emu_to_inches(cx_emu)
        if rw_in <= 0:
            unknown_count += 1
            continue

        dpi = pw / rw_in
        if dpi < dpi_threshold:
            below_threshold += 1
            details.append(f"  {img_name}: {dpi:.0f} DPI ({pw}px / {rw_in:.1f}in) -- below {dpi_threshold} DPI threshold")
        else:
            ok_count += 1

    if unknown_count > 0 and not below_threshold:
        return CheckResult("image_dpi", "PASS",
                          f"{ok_count} OK ({unknown_count} unknown)", details)
    if below_threshold > 0:
        return CheckResult("image_dpi", "WARN",
                          f"{below_threshold} below {dpi_threshold} DPI ({ok_count} OK, {unknown_count} unknown)",
                          details)
    return CheckResult("image_dpi", "PASS", f"all {ok_count} images meet {dpi_threshold} DPI", details)


def check_image_alt_text(word_dir: Path) -> CheckResult:
    doc_xml = word_dir / "document.xml"
    if not doc_xml.exists():
        return CheckResult("image_alt_text", "SKIP", "no document.xml")

    dom = _parse_xml(doc_xml)
    drawings = dom.getElementsByTagName("w:drawing")
    if not drawings:
        return CheckResult("image_alt_text", "SKIP", "no images to check")

    details = []
    missing = 0
    for drawing in drawings:
        doc_prs = drawing.getElementsByTagName("wp:docPr")
        img_name = _get_attr(doc_prs[0], "name") if doc_prs else "unknown"
        descr = _get_attr(doc_prs[0], "descr") if doc_prs else None
        if not descr:
            missing += 1
            details.append(f"  {img_name}: missing alt-text (wp:docPr descr)")

    if missing > 0:
        return CheckResult("image_alt_text", "WARN",
                          f"{missing} image(s) without alt-text", details)
    return CheckResult("image_alt_text", "PASS", "all images have alt-text")


def _extract_figure_numbers(word_dir: Path) -> dict[int, list[str]]:
    """Extract figure numbers from captions. Returns {number: [caption_text]}."""
    doc_xml = word_dir / "document.xml"
    if not doc_xml.exists():
        return {}
    dom = _parse_xml(doc_xml)
    figures: dict[int, list[str]] = {}
    # Chinese: 图1, 图 1, 图1, 图 1
    cn_pattern = re.compile(r"^图\s*(\d+)")
    # English: Fig.1, Fig. 1, Figure 1, Figure.1
    en_pattern = re.compile(r"^(?:Fig\.?|Figure)\s*(\d+)", re.IGNORECASE)
    for p in dom.getElementsByTagName("w:p"):
        text = _paragraph_text(p).strip()
        for pat in (cn_pattern, en_pattern):
            m = pat.match(text)
            if m:
                num = int(m.group(1))
                figures.setdefault(num, []).append(text)
                break
    return figures


def check_figure_numbering(word_dir: Path) -> CheckResult:
    figures = _extract_figure_numbers(word_dir)
    if not figures:
        return CheckResult("figure_numbering", "SKIP", "no figure captions found")

    nums = sorted(figures.keys())
    expected = list(range(1, max(nums) + 1))
    gaps = sorted(set(expected) - set(nums))
    duplicates = {n: texts for n, texts in figures.items() if len(texts) > 1}

    details = []
    if duplicates:
        for n, texts in duplicates.items():
            details.append(f"  figure {n}: duplicated caption ({len(texts)} occurrences)")

    if gaps:
        details.append(f"  missing figures: {gaps}")
        return CheckResult("figure_numbering", "FAIL",
                          f"figures {nums[0]}-{nums[-1]}, {len(gaps)} gap(s)",
                          details)

    if duplicates:
        return CheckResult("figure_numbering", "WARN",
                          f"figures {nums[0]}-{nums[-1]} sequential, {len(duplicates)} duplicate(s)",
                          details)

    return CheckResult("figure_numbering", "PASS",
                      f"figures {nums[0]}-{nums[-1]} sequential and unique")


def _extract_figure_references(word_dir: Path) -> set[int]:
    """Extract figure numbers referenced in body text (excludes captions)."""
    doc_xml = word_dir / "document.xml"
    if not doc_xml.exists():
        return set()
    dom = _parse_xml(doc_xml)
    refs: set[int] = set()
    cn_ref = re.compile(r"图\s*(\d+)")
    en_ref = re.compile(r"(?:Fig\.?|Figure)\s*(\d+)", re.IGNORECASE)
    caption_start = re.compile(r"^(?:图\s*\d+|Fig\.?\s*\d+|Figure\s*\d+)", re.IGNORECASE)
    for p in dom.getElementsByTagName("w:p"):
        text = _paragraph_text(p).strip()
        if caption_start.match(text):
            continue
        for pat in (cn_ref, en_ref):
            for m in pat.finditer(text):
                refs.add(int(m.group(1)))
    return refs


def check_figure_cross_references(word_dir: Path) -> CheckResult:
    figures = _extract_figure_numbers(word_dir)
    refs = _extract_figure_references(word_dir)

    if not figures:
        return CheckResult("figure_cross_references", "SKIP", "no figures to cross-reference")
    if not refs:
        return CheckResult("figure_cross_references", "PASS", "no figure references in body text")

    fig_nums = set(figures.keys())
    missing_refs = refs - fig_nums  # referenced but doesn't exist
    unreferenced = fig_nums - refs  # exists but never referenced

    details = []
    if missing_refs:
        details.append(f"  referenced but missing: figures {sorted(missing_refs)}")
    if unreferenced:
        details.append(f"  exist but never referenced: figures {sorted(unreferenced)}")

    if missing_refs:
        return CheckResult("figure_cross_references", "FAIL",
                          f"{len(missing_refs)} missing reference(s)",
                          details)
    if unreferenced:
        return CheckResult("figure_cross_references", "WARN",
                          f"{len(unreferenced)} unreferenced figure(s)",
                          details)
    return CheckResult("figure_cross_references", "PASS",
                      f"all {len(refs)} figure references match existing figures")


# --- Table checks ---

def check_tables(word_dir: Path) -> CheckResult:
    doc_xml = word_dir / "document.xml"
    if not doc_xml.exists():
        return CheckResult("tables", "FAIL", "document.xml not found")

    dom = _parse_xml(doc_xml)
    tables = dom.getElementsByTagName("w:tbl")

    if not tables:
        return CheckResult("tables", "PASS", "0 tables found")

    details = []
    for i, tbl in enumerate(tables):
        rows = tbl.getElementsByTagName("w:tr")
        row_count = len(rows)
        col_count = 0
        grid = tbl.getElementsByTagName("w:tblGrid")
        if grid:
            col_count = len(grid[0].getElementsByTagName("w:gridCol"))

        merged_cells = 0
        for tc in tbl.getElementsByTagName("w:tc"):
            tc_pr = tc.getElementsByTagName("w:tcPr")
            if tc_pr:
                grid_span = tc_pr[0].getElementsByTagName("w:gridSpan")
                v_merge = tc_pr[0].getElementsByTagName("w:vMerge")
                if grid_span:
                    span = int(_get_attr(grid_span[0], "w:val") or 1)
                    if span > 1:
                        merged_cells += 1
                if v_merge:
                    merged_cells += 1

        detail = f"  table[{i}]: {row_count}x{col_count}"
        if merged_cells > 0:
            detail += f", {merged_cells} merged cell(s)"
        details.append(detail)

    return CheckResult("tables", "PASS",
                      f"{len(tables)} tables: {min(len(t.getElementsByTagName('w:tr')) for t in tables)}-{max(len(t.getElementsByTagName('w:tr')) for t in tables)} rows",
                      details)


def _extract_table_numbers(word_dir: Path) -> dict[int, list[str]]:
    doc_xml = word_dir / "document.xml"
    if not doc_xml.exists():
        return {}
    dom = _parse_xml(doc_xml)
    tables: dict[int, list[str]] = {}
    cn_pattern = re.compile(r"^表\s*(\d+)")
    en_pattern = re.compile(r"^(?:Table|Tab\.)\s*(\d+)", re.IGNORECASE)
    for p in dom.getElementsByTagName("w:p"):
        text = _paragraph_text(p).strip()
        for pat in (cn_pattern, en_pattern):
            m = pat.match(text)
            if m:
                num = int(m.group(1))
                tables.setdefault(num, []).append(text)
                break
    return tables


def check_table_numbering(word_dir: Path) -> CheckResult:
    tables = _extract_table_numbers(word_dir)
    if not tables:
        return CheckResult("table_numbering", "SKIP", "no table captions found")

    nums = sorted(tables.keys())
    expected = list(range(1, max(nums) + 1))
    gaps = sorted(set(expected) - set(nums))
    duplicates = {n: texts for n, texts in tables.items() if len(texts) > 1}

    details = []
    if duplicates:
        for n, texts in duplicates.items():
            details.append(f"  table {n}: duplicated caption ({len(texts)} occurrences)")

    if gaps:
        details.append(f"  missing tables: {gaps}")
        return CheckResult("table_numbering", "FAIL",
                          f"tables {nums[0]}-{nums[-1]}, {len(gaps)} gap(s)",
                          details)

    if duplicates:
        return CheckResult("table_numbering", "WARN",
                          f"tables {nums[0]}-{nums[-1]} sequential, {len(duplicates)} duplicate(s)",
                          details)

    return CheckResult("table_numbering", "PASS",
                      f"tables {nums[0]}-{nums[-1]} sequential and unique")


def _extract_table_references(word_dir: Path) -> set[int]:
    doc_xml = word_dir / "document.xml"
    if not doc_xml.exists():
        return set()
    dom = _parse_xml(doc_xml)
    refs: set[int] = set()
    cn_ref = re.compile(r"表\s*(\d+)")
    en_ref = re.compile(r"(?:Table|Tab\.)\s*(\d+)", re.IGNORECASE)
    caption_start = re.compile(r"^(?:表\s*\d+|Table\s*\d+|Tab\.\s*\d+)", re.IGNORECASE)
    for p in dom.getElementsByTagName("w:p"):
        text = _paragraph_text(p).strip()
        if caption_start.match(text):
            continue
        for pat in (cn_ref, en_ref):
            for m in pat.finditer(text):
                refs.add(int(m.group(1)))
    return refs


def check_table_cross_references(word_dir: Path) -> CheckResult:
    tables = _extract_table_numbers(word_dir)
    refs = _extract_table_references(word_dir)

    if not tables:
        return CheckResult("table_cross_references", "SKIP", "no tables to cross-reference")
    if not refs:
        return CheckResult("table_cross_references", "PASS", "no table references in body text")

    tbl_nums = set(tables.keys())
    missing_refs = refs - tbl_nums
    unreferenced = tbl_nums - refs

    details = []
    if missing_refs:
        details.append(f"  referenced but missing: tables {sorted(missing_refs)}")
    if unreferenced:
        details.append(f"  exist but never referenced: tables {sorted(unreferenced)}")

    if missing_refs:
        return CheckResult("table_cross_references", "FAIL",
                          f"{len(missing_refs)} missing reference(s)",
                          details)
    if unreferenced:
        return CheckResult("table_cross_references", "WARN",
                          f"{len(unreferenced)} unreferenced table(s)",
                          details)
    return CheckResult("table_cross_references", "PASS",
                      f"all {len(refs)} table references match existing tables")


# --- Orchestration ---

def review_content(
    formatted_file: str,
    *,
    verbose: bool = False,
    json_output: bool = False,
    output_file: str | None = None,
    dpi_threshold: int = 300,
) -> list[CheckResult]:
    _ensure_docx_skill()

    input_path = Path(formatted_file).resolve()
    if not input_path.exists():
        print(f"Error: {formatted_file} not found", file=sys.stderr)
        sys.exit(1)

    with tempfile.TemporaryDirectory(prefix="docx-content-review-") as tmp_dir:
        tmp_path = Path(tmp_dir)
        unpack_dir = tmp_path / "unpacked"
        _unpack(str(input_path), str(unpack_dir))
        word_dir = unpack_dir / "word"

        results = [
            check_images(word_dir),
            check_image_dpi(word_dir, dpi_threshold),
            check_image_alt_text(word_dir),
            check_figure_numbering(word_dir),
            check_figure_cross_references(word_dir),
            check_tables(word_dir),
            check_table_numbering(word_dir),
            check_table_cross_references(word_dir),
        ]

    return results


def _format_report(results: list[CheckResult], docx_path: str, verbose: bool) -> str:
    lines = [
        "=== DOCX Content Review Report ===",
        f"File: {docx_path}",
        "",
    ]
    for r in results:
        if r.status == "SKIP":
            continue
        dots = "." * max(1, 30 - len(r.category))
        line = f"{r.category.upper()} {dots} {r.status}  {r.summary}"
        lines.append(line)
        if verbose or r.status in ("WARN", "FAIL"):
            for d in r.details:
                lines.append(d)
    lines.append("")
    has_fail = any(r.status == "FAIL" for r in results)
    lines.append(f"Overall: {'FAIL' if has_fail else 'PASS'}")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(
        description="Review images and tables in a formatted DOCX document"
    )
    parser.add_argument("formatted_docx", help="Formatted DOCX file to review")
    parser.add_argument("--output", "-o", help="Write report to file")
    parser.add_argument("--json", action="store_true", help="Output as JSON")
    parser.add_argument("--verbose", "-v", action="store_true",
                       help="Show full details even on PASS")
    parser.add_argument("--dpi-threshold", type=int, default=300,
                       help="Minimum DPI for print quality (default: 300)")
    args = parser.parse_args()

    results = review_content(
        args.formatted_docx,
        verbose=args.verbose,
        json_output=args.json,
        output_file=args.output,
        dpi_threshold=args.dpi_threshold,
    )

    docx_path = str(Path(args.formatted_docx).resolve())

    if args.json:
        report_data = [
            {"category": r.category, "status": r.status,
             "summary": r.summary, "details": r.details}
            for r in results
        ]
        has_fail = any(r.status == "FAIL" for r in results)
        report_data.append({"category": "overall",
                           "status": "FAIL" if has_fail else "PASS"})
        output = json.dumps(report_data, ensure_ascii=False, indent=2)
    else:
        output = _format_report(results, docx_path, args.verbose)

    print(output)

    if args.output:
        out_path = Path(args.output)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(output, encoding="utf-8")
        print(f"\nReport saved to {out_path}")


if __name__ == "__main__":
    main()
