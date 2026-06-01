#!/usr/bin/env python3
"""Extract a formatting template from a reference DOCX document.

Produces a YAML template capturing styles, page layout, numbering,
and theme — which can be applied to other DOCX files via apply_template.py.

Usage:
    python extract_template.py reference.docx [--output template.yaml]

Dependencies:
    - pip install defusedxml pyyaml
    - ~/.claude/skills/docx/scripts/office/unpack.py (for unpacking)
"""

import argparse
from collections import Counter
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import defusedxml.minidom
import yaml


# --- Helpers ---

UNPACK_PY = os.path.expanduser(
    "~/.claude/skills/docx/scripts/office/unpack.py"
)

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent


def _ensure_docx_skill():
    if not Path(UNPACK_PY).exists():
        print(
            "Error: docx skill not found. Install it first:\n"
            "  Run a docx-related task in Claude Code to auto-install, or\n"
            "  clone from github.com/anthropics/skills",
            file=sys.stderr,
        )
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


def _get_text(elem, tag: str) -> str | None:
    """Get text content of first child element with given tag."""
    children = elem.getElementsByTagName(tag)
    if children and children[0].firstChild:
        return children[0].firstChild.nodeValue.strip()
    return None


def _get_attr(elem, attr: str) -> str | None:
    val = elem.getAttribute(attr)
    return val if val else None


def _direct_children(elem, tag: str | None = None) -> list:
    children = [
        n for n in elem.childNodes
        if n.nodeType == n.ELEMENT_NODE
    ]
    if tag is None:
        return children
    return [n for n in children if n.tagName == tag]


def _direct_child(elem, tag: str):
    children = _direct_children(elem, tag)
    return children[0] if children else None


def _paragraph_text(p) -> str:
    parts = []
    for t in p.getElementsByTagName("w:t"):
        if t.firstChild:
            parts.append(t.firstChild.nodeValue)
    return "".join(parts)


def _run_text(r) -> str:
    parts = []
    for t in r.getElementsByTagName("w:t"):
        if t.firstChild:
            parts.append(t.firstChild.nodeValue)
    return "".join(parts)


def _props_key(props: dict) -> tuple:
    return tuple(sorted(props.items()))


def _props_from_key(key: tuple) -> dict:
    return dict(key)


def _dominant_props(weighted_props: list[tuple[dict, int]]) -> dict:
    counter = Counter()
    for props, weight in weighted_props:
        if props:
            counter[_props_key(props)] += max(1, weight)
    if not counter:
        return {}
    return _props_from_key(counter.most_common(1)[0][0])


def _ensure_docx_input(input_path: Path, tmp_path: Path) -> Path:
    """Return a DOCX path, converting legacy Office files when needed."""
    if input_path.suffix.lower() == ".docx":
        return input_path

    if input_path.suffix.lower() not in {".doc", ".odt", ".rtf"}:
        print(
            f"Error: unsupported reference format: {input_path.suffix}. "
            "Use .docx or a LibreOffice-convertible document.",
            file=sys.stderr,
        )
        sys.exit(1)

    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        print(
            "Error: LibreOffice/soffice is required to convert non-DOCX templates.",
            file=sys.stderr,
        )
        sys.exit(1)

    converted_dir = tmp_path / "converted"
    converted_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [
            soffice, "--headless", "--convert-to", "docx",
            "--outdir", str(converted_dir), str(input_path),
        ],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        print(
            f"Error converting {input_path} to DOCX: {result.stderr or result.stdout}",
            file=sys.stderr,
        )
        sys.exit(1)

    expected = converted_dir / f"{input_path.stem}.docx"
    if expected.exists():
        return expected
    matches = sorted(converted_dir.glob("*.docx"))
    if matches:
        return matches[0]
    print(f"Error: conversion produced no DOCX for {input_path}", file=sys.stderr)
    sys.exit(1)


# --- Extraction Functions ---


def extract_doc_defaults(styles_dom) -> dict:
    """Extract document defaults (default run and paragraph properties)."""
    result = {}
    doc_defaults = styles_dom.getElementsByTagName("w:docDefaults")
    if not doc_defaults:
        return result

    dd = doc_defaults[0]

    # Default run properties
    r_pr_default = dd.getElementsByTagName("w:rPrDefault")
    if r_pr_default:
        rpr = r_pr_default[0].getElementsByTagName("w:rPr")
        if rpr:
            result["run"] = _extract_run_props(rpr[0])

    # Default paragraph properties
    p_pr_default = dd.getElementsByTagName("w:pPrDefault")
    if p_pr_default:
        ppr = p_pr_default[0].getElementsByTagName("w:pPr")
        if ppr:
            result["paragraph"] = _extract_para_props(ppr[0])

    return result


def _extract_run_props(rpr) -> dict:
    """Extract run properties from a w:rPr element."""
    props = {}

    # Font
    r_fonts = rpr.getElementsByTagName("w:rFonts")
    if r_fonts:
        fonts = {}
        for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
            val = _get_attr(r_fonts[0], f"w:{attr}")
            if val:
                fonts[attr] = val
        if fonts:
            latin_font = fonts.get("ascii") or fonts.get("hAnsi")
            if latin_font:
                props["font"] = latin_font
            if "eastAsia" in fonts:
                props["fontEastAsia"] = fonts["eastAsia"]

    # Size (in half-points)
    sz = rpr.getElementsByTagName("w:sz")
    if sz:
        props["fontSize"] = int(_get_attr(sz[0], "w:val"))

    sz_cs = rpr.getElementsByTagName("w:szCs")
    if sz_cs:
        props["fontSizeCs"] = int(_get_attr(sz_cs[0], "w:val"))

    # Color
    color = rpr.getElementsByTagName("w:color")
    if color:
        props["fontColor"] = _get_attr(color[0], "w:val")

    # Bold / Italic
    if rpr.getElementsByTagName("w:b"):
        props["bold"] = True
    if rpr.getElementsByTagName("w:i"):
        props["italic"] = True

    # Underline
    u = rpr.getElementsByTagName("w:u")
    if u:
        props["underline"] = _get_attr(u[0], "w:val")

    # Font spacing (in DXA)
    spacing = rpr.getElementsByTagName("w:spacing")
    if spacing:
        val = _get_attr(spacing[0], "w:val")
        if val:
            props["fontSpacing"] = int(val)

    # Small caps / caps
    if rpr.getElementsByTagName("w:smallCaps"):
        props["smallCaps"] = True
    if rpr.getElementsByTagName("w:caps"):
        props["caps"] = True

    # Strikethrough
    if rpr.getElementsByTagName("w:strike"):
        props["strike"] = True

    return props


def _extract_para_props(ppr) -> dict:
    """Extract paragraph properties from a w:pPr element."""
    props = {}

    spacing = ppr.getElementsByTagName("w:spacing")
    if spacing:
        for attr in ("before", "after", "line", "lineRule"):
            val = _get_attr(spacing[0], f"w:{attr}")
            if val is not None:
                try:
                    props[f"spacing_{attr}"] = int(val)
                except ValueError:
                    props[f"spacing_{attr}"] = val

    indent = ppr.getElementsByTagName("w:ind")
    if indent:
        for attr in ("left", "right", "firstLine", "hanging"):
            val = _get_attr(indent[0], f"w:{attr}")
            if val is not None:
                props[f"indent_{attr}"] = int(val)

    jc = ppr.getElementsByTagName("w:jc")
    if jc:
        props["alignment"] = _get_attr(jc[0], "w:val")

    outline_lvl = ppr.getElementsByTagName("w:outlineLvl")
    if outline_lvl:
        val = _get_attr(outline_lvl[0], "w:val")
        if val is not None:
            props["outlineLevel"] = int(val)

    # Keep with next / page break before / widow control
    if ppr.getElementsByTagName("w:keepNext"):
        props["keepNext"] = True
    if ppr.getElementsByTagName("w:pageBreakBefore"):
        props["pageBreakBefore"] = True
    if ppr.getElementsByTagName("w:widowControl"):
        props["widowControl"] = True

    # Shading
    shading = ppr.getElementsByTagName("w:shd")
    if shading:
        shd = {}
        for attr in ("val", "color", "fill"):
            val = _get_attr(shading[0], f"w:{attr}")
            if val:
                shd[attr] = val
        if shd:
            props["shading"] = shd

    return props


def extract_styles(styles_dom) -> dict:
    """Extract paragraph and character style definitions."""
    para_styles = []
    char_styles = []

    for style in styles_dom.getElementsByTagName("w:style"):
        style_type = _get_attr(style, "w:type")
        style_id = _get_attr(style, "w:styleId")
        if not style_id:
            continue

        entry = {
            "id": style_id,
            "name": _get_text(style, "w:name"),
        }

        based_on = style.getElementsByTagName("w:basedOn")
        if based_on:
            entry["basedOn"] = _get_attr(based_on[0], "w:val")

        next_style = style.getElementsByTagName("w:next")
        if next_style:
            entry["next"] = _get_attr(next_style[0], "w:val")

        # Run properties
        rpr = style.getElementsByTagName("w:rPr")
        if rpr:
            entry["run"] = _extract_run_props(rpr[0])

        # Paragraph properties
        ppr = style.getElementsByTagName("w:pPr")
        if ppr:
            entry["paragraph"] = _extract_para_props(ppr[0])

        if style_type == "paragraph":
            para_styles.append(entry)
        elif style_type == "character":
            char_styles.append(entry)

    return {
        "paragraph_styles": para_styles,
        "character_styles": char_styles,
    }


def extract_body_paragraph_formats(doc_dom) -> list:
    """Extract document-body direct paragraph and run formatting by index.

    Many Word templates, especially legacy .doc files converted to DOCX,
    store the visible font/size on paragraph or run nodes instead of reusable
    styles. Capturing these body-level formats makes the template reflect what
    the reference document actually displays.
    """
    formats = []
    body_list = doc_dom.getElementsByTagName("w:body")
    if not body_list:
        return formats

    paragraphs = [
        n for n in _direct_children(body_list[0], "w:p")
        if _paragraph_text(n).strip() or _direct_child(n, "w:pPr")
    ]

    for index, p in enumerate(paragraphs):
        ppr = _direct_child(p, "w:pPr")
        style = "Normal"
        paragraph_props = {}
        paragraph_run_props = {}

        if ppr:
            pstyle = _direct_child(ppr, "w:pStyle")
            if pstyle:
                style = _get_attr(pstyle, "w:val") or style
            paragraph_props = _extract_para_props(ppr)
            paragraph_run = _direct_child(ppr, "w:rPr")
            if paragraph_run:
                paragraph_run_props = _extract_run_props(paragraph_run)

        run_entries = []
        weighted = []
        for run_index, r in enumerate(_direct_children(p, "w:r")):
            rpr = _direct_child(r, "w:rPr")
            if not rpr:
                continue
            run_props = _extract_run_props(rpr)
            if not run_props:
                continue
            text = _run_text(r)
            weighted.append((run_props, len(text)))
            entry = {
                "index": run_index,
                "run": run_props,
            }
            if text.strip():
                entry["textSample"] = text.strip()[:80]
            run_entries.append(entry)

        dominant_run = _dominant_props(weighted)
        text = _paragraph_text(p).strip()
        entry = {
            "index": index,
            "style": style,
        }
        if text:
            entry["textSample"] = text[:120]
        if paragraph_props:
            entry["paragraph"] = paragraph_props
        if paragraph_run_props:
            entry["paragraph_run"] = paragraph_run_props
        if dominant_run:
            entry["run"] = dominant_run
        if run_entries:
            entry["runs"] = run_entries

        if any(k in entry for k in ("paragraph", "paragraph_run", "run", "runs")):
            formats.append(entry)

    return formats


def enrich_styles_from_body_formats(template: dict) -> int:
    """Fold dominant body-level formatting back into paragraph styles.

    This keeps older consumers useful even if they only understand
    paragraph_styles, while preserving detailed per-paragraph data under
    body_paragraph_formats for exact review/application.
    """
    body_formats = template.get("body_paragraph_formats", [])
    para_styles = template.get("paragraph_styles", [])
    if not body_formats or not para_styles:
        return 0

    by_style = {s.get("id"): s for s in para_styles if s.get("id")}
    buckets: dict[str, dict[str, list[tuple[dict, int]]]] = {}

    for fmt in body_formats:
        style_id = fmt.get("style") or "Normal"
        if style_id not in by_style:
            continue
        weight = max(1, len(fmt.get("textSample", "")))
        bucket = buckets.setdefault(style_id, {"run": [], "paragraph": []})
        run_props = fmt.get("run") or fmt.get("paragraph_run") or {}
        para_props = fmt.get("paragraph") or {}
        if run_props:
            bucket["run"].append((run_props, weight))
        if para_props:
            bucket["paragraph"].append((para_props, weight))

    changed = 0
    for style_id, bucket in buckets.items():
        style = by_style[style_id]
        dominant_run = _dominant_props(bucket["run"])
        dominant_para = _dominant_props(bucket["paragraph"])
        if dominant_run:
            style["run"] = {**style.get("run", {}), **dominant_run}
            changed += 1
        if dominant_para:
            style["paragraph"] = {**style.get("paragraph", {}), **dominant_para}
            changed += 1
    return changed


def build_semantic_format_rules(template: dict) -> list[dict]:
    """Create strict semantic rules when the reference is an instruction template."""
    samples = "\n".join(
        fmt.get("textSample", "")
        for fmt in template.get("body_paragraph_formats", [])
    )
    if "一级标题二号黑体" not in samples and "北京交通大学学报" not in samples:
        return []
    return [
        {"role": "cn_title", "match": {"type": "front_index", "index": 0}, "paragraph": {"alignment": "center", "spacing_line": 360, "spacing_lineRule": "auto"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 40, "fontSizeCs": 40, "bold": True}},
        {"role": "cn_authors", "match": {"type": "front_index", "index": 1}, "paragraph": {"alignment": "center"}, "run": {"font": "楷体_GB2312;楷体", "fontEastAsia": "楷体_GB2312;楷体", "fontSize": 28, "fontSizeCs": 28}},
        {"role": "cn_affiliation", "match": {"type": "front_index", "index": 2}, "paragraph": {"alignment": "center"}, "run": {"font": "Times New Roman", "fontEastAsia": "宋体;SimSun", "fontSize": 18, "fontSizeCs": 18}},
        {"role": "cn_abstract", "match": {"type": "regex", "pattern": r"^摘\s*要[:：]"}, "paragraph": {"alignment": "both", "indent_left": 425, "indent_right": 425}, "run": {"font": "Times New Roman", "fontEastAsia": "楷体_GB2312;楷体", "fontSize": 21, "fontSizeCs": 21}, "prefixes": [{"text": "摘 要：", "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}}, {"text": "摘  要：", "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}}, {"text": "摘要：", "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}}]},
        {"role": "cn_keywords", "match": {"type": "regex", "pattern": r"^关键词[:：]"}, "paragraph": {"alignment": "both", "indent_left": 425, "indent_right": 425}, "run": {"font": "Times New Roman", "fontEastAsia": "楷体_GB2312;楷体", "fontSize": 21, "fontSizeCs": 21}, "prefixes": [{"text": "关键词：", "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}}, {"text": "关键词:", "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}}]},
        {"role": "classification", "match": {"type": "regex", "pattern": r"^中图分类号"}, "paragraph": {"alignment": "both"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}},
        {"role": "en_title", "match": {"type": "front_index", "index": 6}, "paragraph": {"alignment": "center"}, "run": {"font": "Times New Roman", "fontSize": 32, "fontSizeCs": 32, "bold": True}},
        {"role": "en_authors", "match": {"type": "front_index", "index": 7}, "paragraph": {"alignment": "center"}, "run": {"font": "Times New Roman", "fontSize": 28, "fontSizeCs": 28, "italic": True}},
        {"role": "en_affiliation", "match": {"type": "front_index", "index": 8}, "paragraph": {"alignment": "center"}, "run": {"font": "Times New Roman", "fontSize": 18, "fontSizeCs": 18}},
        {"role": "en_abstract", "match": {"type": "regex", "pattern": r"^Abstract[:：]"}, "paragraph": {"alignment": "both"}, "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21}},
        {"role": "en_keywords", "match": {"type": "regex", "pattern": r"^(Key words|Keywords)[:：]"}, "paragraph": {"alignment": "both"}, "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21}},
        {"role": "references_heading", "match": {"type": "regex", "pattern": r"^参考文献"}, "paragraph": {"alignment": "left"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21, "bold": True}, "sets_context": {"after_references": True}},
        {"role": "reference_entry", "match": {"type": "after_references_regex", "pattern": r"^\[?\d+\]"}, "paragraph": {"alignment": "both"}, "run": {"font": "Times New Roman", "fontEastAsia": "宋体;SimSun", "fontSize": 18, "fontSizeCs": 18}},
        {"role": "section_heading", "match": {"type": "regex", "pattern": r"^\d+\s+\S"}, "paragraph": {"alignment": "left", "spacing_line": 220, "spacing_lineRule": "exact"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 28, "fontSizeCs": 28}},
        {"role": "subsection_heading", "match": {"type": "regex", "pattern": r"^\d+\.\d+(\.\d+)?\s+\S"}, "paragraph": {"alignment": "left"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21, "bold": True}},
        {"role": "figure_caption", "match": {"type": "regex", "pattern": r"^(图\s*\d+|Fig\.\s*\d+)"}, "paragraph": {"alignment": "center"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 18, "fontSizeCs": 18}},
        {"role": "table_caption", "match": {"type": "regex", "pattern": r"^(表\s*\d+|Tab\.\s*\d+)"}, "paragraph": {"alignment": "center"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 18, "fontSizeCs": 18, "bold": True}},
        {"role": "table_cell", "match": {"type": "in_table"}, "paragraph": {"alignment": "center"}, "run": {"font": "Times New Roman", "fontEastAsia": "宋体;SimSun", "fontSize": 18, "fontSizeCs": 18, "bold": True}},
        {"role": "body", "match": {"type": "default"}, "paragraph": {"alignment": "both", "indent_firstLine": 420, "spacing_line": 314, "spacing_lineRule": "exact"}, "run": {"font": "Times New Roman", "fontEastAsia": "宋体;SimSun", "fontSize": 21, "fontSizeCs": 21}},
    ]


def extract_page_layout(doc_dom) -> dict:
    """Extract page size, margins, orientation from document.xml."""
    # Look in the last sectPr (most sections override earlier ones)
    sect_prs = doc_dom.getElementsByTagName("w:sectPr")
    if not sect_prs:
        return {}

    sect_pr = sect_prs[-1]

    result = {}

    pg_sz = sect_pr.getElementsByTagName("w:pgSz")
    if pg_sz:
        sz = pg_sz[0]
        result["width"] = int(_get_attr(sz, "w:w"))
        result["height"] = int(_get_attr(sz, "w:h"))
        orient = _get_attr(sz, "w:orient")
        if orient and orient != "portrait":
            result["orientation"] = orient

    pg_mar = sect_pr.getElementsByTagName("w:pgMar")
    if pg_mar:
        mar = pg_mar[0]
        margins = {}
        for attr in ("top", "bottom", "left", "right", "header", "footer", "gutter"):
            val = _get_attr(mar, f"w:{attr}")
            if val is not None:
                margins[attr] = int(val)
        if margins:
            result["margins"] = margins

    # Columns
    cols = sect_pr.getElementsByTagName("w:cols")
    if cols:
        col_count = _get_attr(cols[0], "w:num")
        if col_count and int(col_count) > 1:
            result["columns"] = {"count": int(col_count)}
            space = _get_attr(cols[0], "w:space")
            if space:
                result["columns"]["space"] = int(space)

    return result


def extract_numbering(num_dom) -> list:
    """Extract numbering definitions (bullet and numbered lists)."""
    if not num_dom:
        return []

    definitions = []

    # Get abstract numbering definitions
    abstract_nums = {}
    for anum in num_dom.getElementsByTagName("w:abstractNum"):
        anum_id = _get_attr(anum, "w:abstractNumId")
        if anum_id is None:
            continue

        entry = {"abstractNumId": int(anum_id), "levels": []}

        for lvl in anum.getElementsByTagName("w:lvl"):
            level_data = {}
            level_data["level"] = int(_get_attr(lvl, "w:ilvl") or 0)

            start = lvl.getElementsByTagName("w:start")
            if start:
                level_data["start"] = int(_get_attr(start[0], "w:val") or 1)

            num_fmt = lvl.getElementsByTagName("w:numFmt")
            if num_fmt:
                level_data["format"] = _get_attr(num_fmt[0], "w:val")

            lvl_text = lvl.getElementsByTagName("w:lvlText")
            if lvl_text:
                level_data["text"] = _get_attr(lvl_text[0], "w:val")

            lvl_jc = lvl.getElementsByTagName("w:lvlJc")
            if lvl_jc:
                level_data["alignment"] = _get_attr(lvl_jc[0], "w:val")

            # Paragraph properties for the level
            ppr = lvl.getElementsByTagName("w:pPr")
            if ppr:
                indent = ppr[0].getElementsByTagName("w:ind")
                if indent:
                    for attr in ("left", "right", "firstLine", "hanging"):
                        val = _get_attr(indent[0], f"w:{attr}")
                        if val is not None:
                            level_data[f"indent_{attr}"] = int(val)

            # Run properties for the level
            rpr = lvl.getElementsByTagName("w:rPr")
            if rpr:
                rp = rpr[0]
                r_fonts = rp.getElementsByTagName("w:rFonts")
                if r_fonts:
                    for attr in ("ascii", "hAnsi", "cs"):
                        val = _get_attr(r_fonts[0], f"w:{attr}")
                        if val:
                            level_data[f"font_{attr}"] = val
                sz = rp.getElementsByTagName("w:sz")
                if sz:
                    level_data["fontSize"] = int(_get_attr(sz[0], "w:val"))

            entry["levels"].append(level_data)

        abstract_nums[entry["abstractNumId"]] = entry

    # Get concrete numbering instances and link to abstract defs
    for num in num_dom.getElementsByTagName("w:num"):
        num_id = _get_attr(num, "w:numId")
        if num_id is None:
            continue

        entry = {"id": int(num_id)}

        anum_ref = num.getElementsByTagName("w:abstractNumId")
        if anum_ref:
            ref_id = int(_get_attr(anum_ref[0], "w:val") or 0)
            entry["abstractNumId"] = ref_id
            if ref_id in abstract_nums:
                entry["levels"] = abstract_nums[ref_id]["levels"]

        # Override levels
        for lvl_override in num.getElementsByTagName("w:lvlOverride"):
            lvl_idx = int(_get_attr(lvl_override, "w:ilvl") or 0)
            start_override = lvl_override.getElementsByTagName("w:startOverride")
            if start_override:
                if "levelOverrides" not in entry:
                    entry["levelOverrides"] = {}
                entry["levelOverrides"][lvl_idx] = {
                    "start": int(_get_attr(start_override[0], "w:val") or 1)
                }

        definitions.append(entry)

    return definitions


def extract_theme(theme_dom) -> dict:
    """Extract color scheme and font scheme from theme XML."""
    if not theme_dom:
        return {}

    result = {}
    theme_el = theme_dom.getElementsByTagName("a:theme")
    if not theme_el:
        return result

    theme_elements = theme_el[0].getElementsByTagName("a:themeElements")
    if not theme_elements:
        return result

    te = theme_elements[0]

    # Color scheme
    clr_scheme = te.getElementsByTagName("a:clrScheme")
    if clr_scheme:
        colors = {}
        cs = clr_scheme[0]
        name = _get_attr(cs, "name")
        if name:
            result["schemeName"] = name
        for child in cs.childNodes:
            if child.nodeType == child.ELEMENT_NODE:
                tag = child.tagName
                # The color value could be in a:srgbClr, a:sysClr, etc.
                srgb = child.getElementsByTagName("a:srgbClr")
                if srgb:
                    colors[tag] = _get_attr(srgb[0], "val")
                else:
                    sys_clr = child.getElementsByTagName("a:sysClr")
                    if sys_clr:
                        colors[tag] = _get_attr(sys_clr[0], "lastClr") or _get_attr(
                            sys_clr[0], "val"
                        )
        if colors:
            result["colors"] = colors

    # Font scheme
    font_scheme = te.getElementsByTagName("a:fontScheme")
    if font_scheme:
        fonts = {}
        fs = font_scheme[0]
        name = _get_attr(fs, "name")
        if name:
            result["fontSchemeName"] = name

        major = fs.getElementsByTagName("a:majorFont")
        if major:
            latin = major[0].getElementsByTagName("a:latin")
            if latin:
                fonts["major"] = _get_attr(latin[0], "typeface")

        minor = fs.getElementsByTagName("a:minorFont")
        if minor:
            latin = minor[0].getElementsByTagName("a:latin")
            if latin:
                fonts["minor"] = _get_attr(latin[0], "typeface")

        if fonts:
            result["fonts"] = fonts

    return result


def extract_template(input_file: str, output_file: str | None = None,
                     category: str | None = None) -> str:
    """Extract a formatting template from a DOCX file and write to YAML."""
    input_path = Path(input_file).resolve()
    if not input_path.exists():
        print(f"Error: {input_file} does not exist", file=sys.stderr)
        sys.exit(1)

    _ensure_docx_skill()

    with tempfile.TemporaryDirectory(prefix="docx-extract-") as tmp_dir:
        tmp_path = Path(tmp_dir)
        unpack_dir = tmp_path / "unpacked"
        docx_path = _ensure_docx_input(input_path, tmp_path)

        if docx_path != input_path:
            print(f"Converted {input_path.name} → {docx_path.name}")

        print(f"Unpacking {docx_path}...")
        _unpack(str(docx_path), str(unpack_dir))

        word_dir = unpack_dir / "word"
        template = {}

        # 1. Styles
        styles_xml = word_dir / "styles.xml"
        if styles_xml.exists():
            dom = _parse_xml(styles_xml)
            template["document_defaults"] = extract_doc_defaults(dom)
            template.update(extract_styles(dom))
            print(f"  Extracted {len(template.get('paragraph_styles', []))} paragraph styles")
            print(f"  Extracted {len(template.get('character_styles', []))} character styles")

        # 2. Page layout from document.xml
        doc_xml = word_dir / "document.xml"
        if doc_xml.exists():
            dom = _parse_xml(doc_xml)
            page = extract_page_layout(dom)
            if page:
                template["page"] = page
                print(f"  Extracted page layout: {page.get('width')}x{page.get('height')}")
            body_formats = extract_body_paragraph_formats(dom)
            if body_formats:
                template["body_paragraph_formats"] = body_formats
                template["body_paragraph_formats_mode"] = "style_enrichment"
                print(f"  Extracted {len(body_formats)} body paragraph format profiles")
                enriched = enrich_styles_from_body_formats(template)
                if enriched:
                    print(f"  Enriched {enriched} style format section(s) from body content")
                semantic_rules = build_semantic_format_rules(template)
                if semantic_rules:
                    template["semantic_format_rules"] = semantic_rules
                    print(f"  Built {len(semantic_rules)} strict semantic format rule(s)")

        # 3. Numbering
        num_xml = word_dir / "numbering.xml"
        if num_xml.exists():
            dom = _parse_xml(num_xml)
            numbering = extract_numbering(dom)
            if numbering:
                template["numbering"] = numbering
                print(f"  Extracted {len(numbering)} numbering definitions")

        # 4. Theme
        theme_xml = word_dir / "theme" / "theme1.xml"
        if theme_xml.exists():
            dom = _parse_xml(theme_xml)
            theme = extract_theme(dom)
            if theme:
                template["theme"] = theme
                print(f"  Extracted theme: {theme.get('schemeName', 'unnamed')}")

        # 5. Font table
        font_table_xml = word_dir / "fontTable.xml"
        if font_table_xml.exists():
            dom = _parse_xml(font_table_xml)
            fonts = []
            for font_el in dom.getElementsByTagName("w:font"):
                name = _get_attr(font_el, "w:name")
                if name:
                    fonts.append(name)
            if fonts:
                template["fonts"] = fonts
                print(f"  Extracted {len(fonts)} declared fonts")

    # Write output
    if output_file is None:
        stem = input_path.stem
        if category:
            out_dir = _PROJECT_ROOT / "styles" / category
            out_dir.mkdir(parents=True, exist_ok=True)
            output_file = str(out_dir / f"{stem}-template.yaml")
        else:
            output_file = f"{stem}-template.yaml"

    output_path = Path(output_file)
    with open(output_path, "w", encoding="utf-8") as f:
        yaml.dump(template, f, default_flow_style=False, allow_unicode=True, sort_keys=False)

    print(f"\nTemplate written to {output_path}")
    return str(output_path)


def main():
    parser = argparse.ArgumentParser(
        description="Extract formatting template from a reference DOCX document"
    )
    parser.add_argument("input_file", help="Reference DOCX file")
    parser.add_argument(
        "--output", "-o",
        help="Output YAML file (default: {input_stem}-template.yaml)",
    )
    parser.add_argument(
        "--category", "-c",
        help="Template name (e.g., 北交模板). "
             "Saves template to styles/<name>/<stem>-template.yaml",
    )
    args = parser.parse_args()

    extract_template(args.input_file, args.output, category=args.category)


if __name__ == "__main__":
    main()
