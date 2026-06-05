#!/usr/bin/env python3
"""Validate a formatted DOCX against a template YAML.

Checks: page layout, document defaults, paragraph/character styles,
theme colors/fonts, font table, numbering definitions, and body
direct formatting. Produces a structured pass/warn/fail report.

Usage:
    python review_template.py formatted.docx template.yaml
    python review_template.py formatted.docx --category 北交模板
    python review_template.py formatted.docx --category 北交模板 --json

Dependencies:
    - pip install defusedxml pyyaml
    - ~/.claude/skills/docx/scripts/office/unpack.py
"""

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

import defusedxml.minidom
import yaml

UNPACK_PY = os.path.expanduser(
    "~/.claude/skills/docx/scripts/office/unpack.py"
)
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent


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


def _get_text(elem, tag: str) -> str | None:
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


def _extract_run_props(rpr) -> dict:
    props = {}
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
    sz = rpr.getElementsByTagName("w:sz")
    if sz:
        props["fontSize"] = int(_get_attr(sz[0], "w:val"))
    sz_cs = rpr.getElementsByTagName("w:szCs")
    if sz_cs:
        props["fontSizeCs"] = int(_get_attr(sz_cs[0], "w:val"))
    color = rpr.getElementsByTagName("w:color")
    if color:
        props["fontColor"] = _get_attr(color[0], "w:val")
    if rpr.getElementsByTagName("w:b"):
        props["bold"] = True
    if rpr.getElementsByTagName("w:i"):
        props["italic"] = True
    u = rpr.getElementsByTagName("w:u")
    if u:
        props["underline"] = _get_attr(u[0], "w:val")
    spacing = rpr.getElementsByTagName("w:spacing")
    if spacing:
        val = _get_attr(spacing[0], "w:val")
        if val:
            props["fontSpacing"] = int(val)
    if rpr.getElementsByTagName("w:smallCaps"):
        props["smallCaps"] = True
    if rpr.getElementsByTagName("w:caps"):
        props["caps"] = True
    if rpr.getElementsByTagName("w:strike"):
        props["strike"] = True
    return props


def _extract_para_props(ppr) -> dict:
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
    if ppr.getElementsByTagName("w:keepNext"):
        props["keepNext"] = True
    if ppr.getElementsByTagName("w:pageBreakBefore"):
        props["pageBreakBefore"] = True
    if ppr.getElementsByTagName("w:widowControl"):
        props["widowControl"] = True
    if ppr.getElementsByTagName("w:keepLines"):
        props["keepLines"] = True
    outline_lvl = ppr.getElementsByTagName("w:outlineLvl")
    if outline_lvl:
        val = _get_attr(outline_lvl[0], "w:val")
        if val is not None:
            props["outlineLevel"] = int(val)
    return props


def _compare_props(expected: dict, actual: dict, path: str) -> list[str]:
    """Compare expected vs actual property dicts, return diff lines."""
    diffs = []
    all_keys = set(expected.keys()) | set(actual.keys())
    for key in sorted(all_keys):
        ev = expected.get(key)
        av = actual.get(key)
        if ev is None and av is None:
            continue
        if ev is None and av is not None:
            diffs.append(f"  {path}.{key}: unexpected value '{av}'")
        elif ev is not None and av is None:
            diffs.append(f"  {path}.{key}: missing (expected '{ev}')")
        elif str(ev) != str(av):
            diffs.append(f"  {path}.{key}: expected '{ev}', got '{av}'")
    return diffs


# --- Checks ---


def check_page_layout(word_dir: Path, template: dict) -> CheckResult:
    if "page" not in template:
        return CheckResult("page_layout", "SKIP", "no page section in template")
    doc_xml = word_dir / "document.xml"
    if not doc_xml.exists():
        return CheckResult("page_layout", "FAIL", "document.xml not found")
    dom = _parse_xml(doc_xml)
    sect_prs = dom.getElementsByTagName("w:sectPr")
    if not sect_prs:
        return CheckResult("page_layout", "FAIL", "no sectPr found")
    sp = sect_prs[-1]
    details = []
    tpl = template["page"]

    pg_sz = sp.getElementsByTagName("w:pgSz")
    if pg_sz and "width" in tpl and "height" in tpl:
        tw = int(_get_attr(pg_sz[0], "w:w") or 0)
        th = int(_get_attr(pg_sz[0], "w:h") or 0)
        if tw != tpl.get("width"):
            details.append(f"  width: expected {tpl['width']}, got {tw}")
        if th != tpl.get("height"):
            details.append(f"  height: expected {tpl['height']}, got {th}")

    margins = tpl.get("margins", {})
    if margins:
        pg_mar = sp.getElementsByTagName("w:pgMar")
        if pg_mar:
            for attr in ("top", "bottom", "left", "right"):
                if attr in margins:
                    av = int(_get_attr(pg_mar[0], f"w:{attr}") or 0)
                    if av != margins[attr]:
                        details.append(f"  margin.{attr}: expected {margins[attr]}, got {av}")

    if details:
        return CheckResult("page_layout", "FAIL", f"{len(details)} mismatch(es)", details)
    return CheckResult("page_layout", "PASS", "all dimensions match")


def check_doc_defaults(word_dir: Path, template: dict) -> CheckResult:
    dd = template.get("document_defaults", {})
    if not dd:
        return CheckResult("doc_defaults", "SKIP", "no document_defaults in template")
    styles_xml = word_dir / "styles.xml"
    if not styles_xml.exists():
        return CheckResult("doc_defaults", "FAIL", "styles.xml not found")
    dom = _parse_xml(styles_xml)
    dd_list = dom.getElementsByTagName("w:docDefaults")
    if not dd_list:
        return CheckResult("doc_defaults", "FAIL", "no docDefaults in output")
    details = []

    run_tpl = dd.get("run", {})
    if run_tpl:
        rpr_def = dd_list[0].getElementsByTagName("w:rPrDefault")
        if rpr_def:
            rpr = rpr_def[0].getElementsByTagName("w:rPr")
            if rpr:
                actual = _extract_run_props(rpr[0])
                details.extend(_compare_props(run_tpl, actual, "run"))

    para_tpl = dd.get("paragraph", {})
    if para_tpl:
        ppr_def = dd_list[0].getElementsByTagName("w:pPrDefault")
        if ppr_def:
            ppr = ppr_def[0].getElementsByTagName("w:pPr")
            if ppr:
                actual = _extract_para_props(ppr[0])
                details.extend(_compare_props(para_tpl, actual, "paragraph"))

    if details:
        return CheckResult("doc_defaults", "FAIL", f"{len(details)} mismatch(es)", details)
    return CheckResult("doc_defaults", "PASS", "all defaults match")


def _get_styles_by_id(styles_dom) -> dict:
    result = {}
    for style in styles_dom.getElementsByTagName("w:style"):
        sid = _get_attr(style, "w:styleId")
        if sid:
            result[sid] = style
    return result


def check_styles(word_dir: Path, template: dict, style_type: str) -> CheckResult:
    key = "paragraph_styles" if style_type == "paragraph" else "character_styles"
    style_list = template.get(key, [])
    if not style_list:
        return CheckResult(f"{style_type}_styles", "SKIP", f"no {key} in template")

    styles_xml = word_dir / "styles.xml"
    if not styles_xml.exists():
        return CheckResult(f"{style_type}_styles", "FAIL", "styles.xml not found")
    dom = _parse_xml(styles_xml)
    existing = _get_styles_by_id(dom)

    exist_total = 0
    details = []
    for s in style_list:
        sid = s.get("id")
        if not sid:
            continue
        if sid in existing:
            exist_total += 1
            style_el = existing[sid]
            # Check run properties
            run_tpl = s.get("run", {})
            if run_tpl:
                rpr_list = style_el.getElementsByTagName("w:rPr")
                actual = _extract_run_props(rpr_list[0]) if rpr_list else {}
                details.extend(_compare_props(run_tpl, actual, f"'{sid}'.run"))
            # Check paragraph properties
            para_tpl = s.get("paragraph", {})
            if para_tpl:
                ppr_list = style_el.getElementsByTagName("w:pPr")
                actual = _extract_para_props(ppr_list[0]) if ppr_list else {}
                details.extend(_compare_props(para_tpl, actual, f"'{sid}'.paragraph"))
        else:
            details.append(f"  style '{sid}': missing in output")

    total = len(style_list)
    missing_count = total - exist_total
    if missing_count > 0:
        return CheckResult(
            f"{style_type}_styles", "FAIL",
            f"{exist_total}/{total} exist, {missing_count} missing",
            details,
        )
    if details:
        return CheckResult(
            f"{style_type}_styles", "WARN",
            f"{exist_total}/{total} exist, {len(details)} property mismatch(es)",
            details,
        )
    return CheckResult(f"{style_type}_styles", "PASS", f"{exist_total}/{total} exist")


def check_theme(word_dir: Path, template: dict) -> CheckResult:
    if "theme" not in template:
        return CheckResult("theme", "SKIP", "no theme in template")
    theme_xml = word_dir / "theme" / "theme1.xml"
    if not theme_xml.exists():
        return CheckResult("theme", "FAIL", "theme/theme1.xml not found")
    dom = _parse_xml(theme_xml)
    tpl = template["theme"]
    details = []
    te_list = dom.getElementsByTagName("a:themeElements")
    if not te_list:
        return CheckResult("theme", "FAIL", "no themeElements")
    te = te_list[0]

    # Colors
    tpl_colors = tpl.get("colors", {})
    if tpl_colors:
        cs_list = te.getElementsByTagName("a:clrScheme")
        if cs_list:
            for child in cs_list[0].childNodes:
                if child.nodeType != child.ELEMENT_NODE:
                    continue
                tag = child.tagName
                if tag not in tpl_colors:
                    continue
                srgb = child.getElementsByTagName("a:srgbClr")
                if srgb:
                    val = _get_attr(srgb[0], "val")
                    if val and val.upper() != tpl_colors[tag].upper():
                        details.append(f"  color {tag}: expected {tpl_colors[tag]}, got {val}")
                else:
                    sys_clr = child.getElementsByTagName("a:sysClr")
                    if sys_clr:
                        val = _get_attr(sys_clr[0], "lastClr") or _get_attr(sys_clr[0], "val")
                        if val and val.upper() != tpl_colors[tag].upper():
                            details.append(f"  color {tag}: expected {tpl_colors[tag]}, got {val}")

    # Fonts
    tpl_fonts = tpl.get("fonts", {})
    if tpl_fonts:
        fs_list = te.getElementsByTagName("a:fontScheme")
        if fs_list:
            for ft in ("major", "minor"):
                expected = tpl_fonts.get(ft)
                if not expected:
                    continue
                container = fs_list[0].getElementsByTagName(f"a:{ft}Font")
                if container:
                    latin = container[0].getElementsByTagName("a:latin")
                    if latin:
                        actual = _get_attr(latin[0], "typeface")
                        if actual and actual != expected:
                            details.append(f"  font.{ft}: expected '{expected}', got '{actual}'")

    if details:
        return CheckResult("theme", "FAIL", f"{len(details)} mismatch(es)", details)
    return CheckResult("theme", "PASS", "colors and fonts match")


def check_font_table(word_dir: Path, template: dict) -> CheckResult:
    if "fonts" not in template:
        return CheckResult("font_table", "SKIP", "no fonts in template")
    ft_xml = word_dir / "fontTable.xml"
    if not ft_xml.exists():
        return CheckResult("font_table", "FAIL", "fontTable.xml not found")
    dom = _parse_xml(ft_xml)
    declared = set()
    for f in dom.getElementsByTagName("w:font"):
        name = _get_attr(f, "w:name")
        if name:
            declared.add(name)
    details = []
    for fn in template["fonts"]:
        if fn not in declared:
            details.append(f"  font '{fn}': not declared in output")
    if details:
        return CheckResult("font_table", "WARN", f"{len(details)} missing", details)
    return CheckResult("font_table", "PASS", f"{len(template['fonts'])} fonts declared")


def check_numbering(word_dir: Path, template: dict) -> CheckResult:
    if "numbering" not in template:
        return CheckResult("numbering", "SKIP", "no numbering in template")
    num_xml = word_dir / "numbering.xml"
    tpl_nums = template["numbering"]
    if not num_xml.exists():
        ids = [str(n.get("id")) for n in tpl_nums]
        return CheckResult("numbering", "FAIL",
                           f"numbering.xml not found, {len(tpl_nums)} missing: {', '.join(ids)}")
    dom = _parse_xml(num_xml)
    output_ids = set()
    for num in dom.getElementsByTagName("w:num"):
        nid = _get_attr(num, "w:numId")
        if nid is not None:
            output_ids.add(int(nid))
    details = []
    found = 0
    for n in tpl_nums:
        nid = n.get("id")
        if nid in output_ids:
            found += 1
        else:
            details.append(f"  numbering id={nid}: missing")
    if details:
        return CheckResult("numbering", "WARN",
                           f"{found}/{len(tpl_nums)} found", details)
    return CheckResult("numbering", "PASS", f"{found}/{len(tpl_nums)} found")


def _props_key(props: dict) -> tuple:
    return tuple(sorted(props.items()))


def _props_from_key(key: tuple) -> dict:
    return dict(key)


def _dominant_run_props(p) -> dict:
    counter = Counter()
    for r in _direct_children(p, "w:r"):
        rpr = _direct_child(r, "w:rPr")
        if not rpr:
            continue
        props = _extract_run_props(rpr)
        if not props:
            continue
        text_len = 0
        for t in r.getElementsByTagName("w:t"):
            if t.firstChild:
                text_len += len(t.firstChild.nodeValue)
        counter[_props_key(props)] += max(1, text_len)
    if not counter:
        return {}
    return _props_from_key(counter.most_common(1)[0][0])


def check_body_paragraph_formats(word_dir: Path, template: dict) -> CheckResult:
    body_formats = template.get("body_paragraph_formats", [])
    if not body_formats:
        return CheckResult("body_paragraph_formats", "SKIP", "no body_paragraph_formats in template")
    if template.get("body_paragraph_formats_mode") != "index":
        return CheckResult(
            "body_paragraph_formats",
            "SKIP",
            "body_paragraph_formats used for style enrichment, not index review",
        )
    doc_xml = word_dir / "document.xml"
    if not doc_xml.exists():
        return CheckResult("body_paragraph_formats", "FAIL", "document.xml not found")
    dom = _parse_xml(doc_xml)
    body = dom.getElementsByTagName("w:body")
    if not body:
        return CheckResult("body_paragraph_formats", "FAIL", "no body found")
    paragraphs = _direct_children(body[0], "w:p")

    details = []
    checked = 0
    for fmt in body_formats:
        idx = fmt.get("index")
        if idx is None:
            continue
        if idx < 0 or idx >= len(paragraphs):
            details.append(f"  paragraph[{idx}]: missing in output")
            continue
        checked += 1
        p = paragraphs[idx]
        ppr = _direct_child(p, "w:pPr")
        actual_para = _extract_para_props(ppr) if ppr else {}
        expected_para = fmt.get("paragraph") or {}
        if expected_para:
            details.extend(_compare_props(expected_para, actual_para, f"paragraph[{idx}].paragraph"))

        expected_para_run = fmt.get("paragraph_run") or {}
        if expected_para_run:
            actual_para_run = {}
            if ppr:
                paragraph_rpr = _direct_child(ppr, "w:rPr")
                if paragraph_rpr:
                    actual_para_run = _extract_run_props(paragraph_rpr)
            details.extend(_compare_props(expected_para_run, actual_para_run, f"paragraph[{idx}].paragraph_run"))

        expected_run = fmt.get("run") or {}
        if expected_run:
            actual_run = _dominant_run_props(p)
            details.extend(_compare_props(expected_run, actual_run, f"paragraph[{idx}].run"))

        runs = _direct_children(p, "w:r")
        for run_fmt in fmt.get("runs", []):
            run_idx = run_fmt.get("index")
            expected = run_fmt.get("run") or {}
            if not expected or run_idx is None:
                continue
            if run_idx < 0 or run_idx >= len(runs):
                details.append(f"  paragraph[{idx}].runs[{run_idx}]: missing in output")
                continue
            rpr = _direct_child(runs[run_idx], "w:rPr")
            actual = _extract_run_props(rpr) if rpr else {}
            details.extend(_compare_props(expected, actual, f"paragraph[{idx}].runs[{run_idx}]"))

    if details:
        return CheckResult(
            "body_paragraph_formats",
            "FAIL",
            f"{len(details)} mismatch(es) across {checked}/{len(body_formats)} paragraph profile(s)",
            details,
        )
    return CheckResult(
        "body_paragraph_formats",
        "PASS",
        f"{checked}/{len(body_formats)} paragraph profile(s) match",
    )


def check_body_direct_formatting(word_dir: Path, template: dict | None = None) -> CheckResult:
    doc_xml = word_dir / "document.xml"
    if not doc_xml.exists():
        return CheckResult("direct_formatting", "FAIL", "document.xml not found")
    dom = _parse_xml(doc_xml)
    run_format_tags = {
        "w:rFonts", "w:sz", "w:szCs", "w:color",
        "w:b", "w:i", "w:u", "w:strike",
        "w:smallCaps", "w:caps", "w:spacing",
    }
    para_format_tags = {"w:spacing", "w:ind", "w:jc"}
    para_run_format_tags = run_format_tags
    total = 0
    run_count = 0
    para_count = 0
    para_run_count = 0
    for p in dom.getElementsByTagName("w:p"):
        for r in p.getElementsByTagName("w:r"):
            rpr_list = r.getElementsByTagName("w:rPr")
            if not rpr_list:
                continue
            for child in rpr_list[0].childNodes:
                if child.nodeType == child.ELEMENT_NODE and child.tagName in run_format_tags:
                    run_count += 1
                    total += 1
        ppr_list = p.getElementsByTagName("w:pPr")
        if not ppr_list:
            continue
        for child in ppr_list[0].childNodes:
            if child.nodeType != child.ELEMENT_NODE:
                continue
            if child.tagName in para_format_tags:
                para_count += 1
                total += 1
            elif child.tagName == "w:rPr":
                for run_child in child.childNodes:
                    if (
                        run_child.nodeType == run_child.ELEMENT_NODE
                        and run_child.tagName in para_run_format_tags
                    ):
                        para_run_count += 1
                        total += 1
    if total == 0:
        return CheckResult("direct_formatting", "PASS", "no direct formatting found")
    if template and template.get("body_paragraph_formats"):
        return CheckResult(
            "direct_formatting",
            "PASS",
            "direct formatting governed by body_paragraph_formats",
            [f"  {run_count} run-level, {para_count} paragraph-level, {para_run_count} paragraph-run-level"],
        )
    details = [f"  {run_count} run-level, {para_count} paragraph-level, {para_run_count} paragraph-run-level"]
    return CheckResult(
        "direct_formatting",
        "WARN",
        f"{total} direct formatting instance(s) found "
        f"({run_count} run-level, {para_count} paragraph-level, {para_run_count} paragraph-run-level)",
        details,
    )


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


def _has_ancestor(elem, tag: str) -> bool:
    node = elem.parentNode
    while node is not None:
        if getattr(node, "tagName", None) == tag:
            return True
        node = node.parentNode
    return False


def _default_semantic_format_rules() -> list[dict]:
    base_body = {
        "font": "Times New Roman",
        "fontEastAsia": "宋体;SimSun",
        "fontSize": 21,
        "fontSizeCs": 21,
    }
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
        {"role": "en_abstract", "match": {"type": "regex", "pattern": r"^Abstract[:：]"}, "paragraph": {"alignment": "both"}, "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21}, "prefixes": [{"text": "Abstract:", "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21, "bold": True}}, {"text": "Abstract：", "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21, "bold": True}}]},
        {"role": "en_keywords", "match": {"type": "regex", "pattern": r"^(Key words|Keywords)[:：]"}, "paragraph": {"alignment": "both"}, "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21}, "prefixes": [{"text": "Key words:", "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21, "bold": True}}, {"text": "Keywords:", "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21, "bold": True}}, {"text": "Key words：", "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21, "bold": True}}, {"text": "Keywords：", "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21, "bold": True}}]},
        {"role": "references_heading", "match": {"type": "regex", "pattern": r"^参考文献"}, "paragraph": {"alignment": "left"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21, "bold": True}, "sets_context": {"after_references": True}},
        {"role": "reference_entry", "match": {"type": "after_references"}, "paragraph": {"alignment": "both", "indent_left": 425, "indent_hanging": 425}, "run": {"font": "Times New Roman", "fontEastAsia": "宋体;SimSun", "fontSize": 18, "fontSizeCs": 18}},
        {"role": "section_heading", "match": {"type": "regex", "pattern": r"^\d+\s+\S"}, "paragraph": {"alignment": "left", "spacing_line": 220, "spacing_lineRule": "exact"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 28, "fontSizeCs": 28}},
        {"role": "subsection_heading", "match": {"type": "regex", "pattern": r"^\d+\.\d+(\.\d+)?\s+\S"}, "paragraph": {"alignment": "left"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}},
        {"role": "figure_caption_cn", "match": {"type": "regex", "pattern": r"^图\s*\d+"}, "paragraph": {"alignment": "center"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 18, "fontSizeCs": 18}},
        {"role": "figure_caption_en", "match": {"type": "regex", "pattern": r"^(Fig\.|Figure)\s*\d+"}, "paragraph": {"alignment": "center"}, "run": {"font": "Times New Roman", "fontSize": 18, "fontSizeCs": 18}},
        {"role": "table_caption_cn", "match": {"type": "regex", "pattern": r"^表\s*\d+"}, "paragraph": {"alignment": "center"}, "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 18, "fontSizeCs": 18}},
        {"role": "table_caption_en", "match": {"type": "regex", "pattern": r"^(Table|Tab\.)\s*\d+"}, "paragraph": {"alignment": "center"}, "run": {"font": "Times New Roman", "fontSize": 18, "fontSizeCs": 18}},
        {"role": "table_cell", "match": {"type": "in_table"}, "paragraph": {"alignment": "center"}, "run": {"font": "Times New Roman", "fontEastAsia": "宋体;SimSun", "fontSize": 18, "fontSizeCs": 18}},
        {"role": "body", "match": {"type": "default"}, "paragraph": {"alignment": "both", "indent_firstLine": 420, "spacing_line": 314, "spacing_lineRule": "exact"}, "run": base_body},
    ]


def _semantic_rules(template: dict) -> list[dict]:
    return template.get("semantic_format_rules") or _default_semantic_format_rules()


def _rule_matches(rule: dict, text: str, front_index: int | None,
                  in_table: bool, context: dict) -> bool:
    match = rule.get("match", {})
    match_type = match.get("type")
    if match_type == "front_index":
        return front_index == match.get("index")
    if match_type == "regex":
        return bool(re.match(match.get("pattern", ""), text))
    if match_type == "after_references_regex":
        return context.get("after_references") and bool(re.match(match.get("pattern", ""), text))
    if match_type == "after_references":
        return bool(context.get("after_references"))
    if match_type == "in_table":
        return in_table
    if match_type == "default":
        return not in_table
    return False


def _select_semantic_rule(rules: list[dict], text: str, front_index: int | None,
                          in_table: bool, context: dict) -> dict | None:
    for rule in rules:
        if rule.get("match", {}).get("type") == "default":
            continue
        if _rule_matches(rule, text, front_index, in_table, context):
            return rule
    for rule in rules:
        if rule.get("match", {}).get("type") == "default" and _rule_matches(rule, text, front_index, in_table, context):
            return rule
    return None


def _font_matches(expected: str, actual: str | None) -> bool:
    if not expected:
        return True
    if not actual:
        return False
    expected_parts = {p.strip() for p in expected.split(";") if p.strip()}
    actual_parts = {p.strip() for p in actual.split(";") if p.strip()}
    return bool(expected_parts & actual_parts) or expected == actual


def _text_has_cjk(text: str) -> bool:
    return any("\u4e00" <= ch <= "\u9fff" for ch in text)


def _text_has_latin(text: str) -> bool:
    return any(("A" <= ch <= "Z") or ("a" <= ch <= "z") or ch.isdigit() for ch in text)


def _effective_run_props(p, r) -> dict:
    props = {}
    ppr = _direct_child(p, "w:pPr")
    if ppr:
        pr = _direct_child(ppr, "w:rPr")
        if pr:
            props.update(_extract_run_props(pr))
    rpr = _direct_child(r, "w:rPr")
    if rpr:
        props.update(_extract_run_props(rpr))
    return props


def _expected_run_for_offset(rule: dict, text: str, offset: int, run_text: str) -> tuple[dict, str]:
    for prefix in rule.get("prefixes", []):
        prefix_text = prefix.get("text", "")
        if text.startswith(prefix_text) and offset < len(prefix_text):
            if offset + len(run_text) > len(prefix_text):
                return prefix.get("run", {}), "split-required"
            return prefix.get("run", {}), "prefix"
    return rule.get("run", {}), "default"


def _compare_run_strict(expected: dict, actual: dict, text: str, path: str) -> list[str]:
    details = []
    if "fontEastAsia" in expected and _text_has_cjk(text):
        if not _font_matches(expected["fontEastAsia"], actual.get("fontEastAsia")):
            details.append(f"  {path}.fontEastAsia: expected '{expected['fontEastAsia']}', got '{actual.get('fontEastAsia')}'")
    if "font" in expected and _text_has_latin(text):
        if not _font_matches(expected["font"], actual.get("font")):
            details.append(f"  {path}.font: expected '{expected['font']}', got '{actual.get('font')}'")
    for key in ("fontSize", "fontSizeCs"):
        if key in expected and str(actual.get(key)) != str(expected[key]):
            details.append(f"  {path}.{key}: expected '{expected[key]}', got '{actual.get(key)}'")
    for key in ("bold", "italic"):
        expected_bool = bool(expected.get(key))
        actual_bool = bool(actual.get(key))
        if expected_bool != actual_bool:
            details.append(f"  {path}.{key}: expected '{expected_bool}', got '{actual_bool}'")
    return details


def check_semantic_format_rules(word_dir: Path, template: dict) -> CheckResult:
    rules = _semantic_rules(template)
    if not rules:
        return CheckResult("semantic_format_rules", "SKIP", "no semantic_format_rules in template")
    doc_xml = word_dir / "document.xml"
    if not doc_xml.exists():
        return CheckResult("semantic_format_rules", "FAIL", "document.xml not found")
    dom = _parse_xml(doc_xml)
    details = []
    checked = 0
    front_index = -1
    context = {"after_references": False}

    for p_idx, p in enumerate(dom.getElementsByTagName("w:p")):
        text = _paragraph_text(p).strip()
        if not text:
            continue
        in_table = _has_ancestor(p, "w:tc")
        if not in_table:
            front_index += 1
        rule = _select_semantic_rule(rules, text, front_index if not in_table else None, in_table, context)
        if not rule:
            continue
        checked += 1

        expected_para = rule.get("paragraph") or {}
        if expected_para:
            ppr = _direct_child(p, "w:pPr")
            actual_para = _extract_para_props(ppr) if ppr else {}
            for key, expected_value in expected_para.items():
                if str(actual_para.get(key)) != str(expected_value):
                    details.append(f"  paragraph[{p_idx}]({rule['role']}).paragraph.{key}: expected '{expected_value}', got '{actual_para.get(key)}'")

        offset = 0
        for r_idx, r in enumerate(_direct_children(p, "w:r")):
            run_text = _run_text(r)
            if not run_text:
                continue
            expected_run, segment = _expected_run_for_offset(rule, text, offset, run_text)
            if segment == "split-required":
                details.append(f"  paragraph[{p_idx}]({rule['role']}).runs[{r_idx}]: prefix formatting must be split into a separate run")
            actual_run = _effective_run_props(p, r)
            details.extend(_compare_run_strict(
                expected_run, actual_run, run_text,
                f"paragraph[{p_idx}]({rule['role']}).runs[{r_idx}]",
            ))
            offset += len(run_text)

        for key, value in (rule.get("sets_context") or {}).items():
            context[key] = value

    if details:
        return CheckResult(
            "semantic_format_rules",
            "FAIL",
            f"{len(details)} strict formatting mismatch(es) across {checked} paragraph(s)",
            details[:500],
        )
    return CheckResult("semantic_format_rules", "PASS", f"{checked} paragraph(s) match strict semantic rules")


def review_docx(
    formatted_file: str,
    template_file: str | None = None,
    *,
    category: str | None = None,
    verbose: bool = False,
    json_output: bool = False,
    output_file: str | None = None,
) -> list[CheckResult]:
    """Review a formatted DOCX against a template YAML."""
    _ensure_docx_skill()

    # Resolve template
    if template_file is None and category:
        styles_dir = _PROJECT_ROOT / "styles" / category
        if not styles_dir.exists():
            print(f"Error: styles/{category}/ not found", file=sys.stderr)
            sys.exit(1)
        yaml_files = sorted(styles_dir.glob("*.yaml")) + sorted(styles_dir.glob("*.yml"))
        if not yaml_files:
            print(f"Error: no .yaml files in styles/{category}/", file=sys.stderr)
            sys.exit(1)
        if len(yaml_files) > 1:
            print(f"Error: multiple templates in styles/{category}/, specify one",
                  file=sys.stderr)
            sys.exit(1)
        template_file = str(yaml_files[0])
    elif template_file is None:
        print("Error: template_file or --category required", file=sys.stderr)
        sys.exit(1)

    with open(template_file, "r", encoding="utf-8") as f:
        template = yaml.safe_load(f)
    if not template:
        print("Error: Empty template", file=sys.stderr)
        sys.exit(1)

    input_path = Path(formatted_file).resolve()
    if not input_path.exists():
        print(f"Error: {formatted_file} not found", file=sys.stderr)
        sys.exit(1)

    with tempfile.TemporaryDirectory(prefix="docx-review-") as tmp_dir:
        tmp_path = Path(tmp_dir)
        unpack_dir = tmp_path / "unpacked"
        _unpack(str(input_path), str(unpack_dir))
        word_dir = unpack_dir / "word"

        results = [
            check_page_layout(word_dir, template),
            check_doc_defaults(word_dir, template),
            check_styles(word_dir, template, "paragraph"),
            check_styles(word_dir, template, "character"),
            check_theme(word_dir, template),
            check_font_table(word_dir, template),
            check_numbering(word_dir, template),
            check_body_paragraph_formats(word_dir, template),
            check_semantic_format_rules(word_dir, template),
            check_body_direct_formatting(word_dir, template),
        ]

    return results


def _format_report(results: list[CheckResult], docx_path: str, tpl_path: str,
                   verbose: bool) -> str:
    lines = ["=== DOCX Template Review Report ===",
             f"File: {docx_path}",
             f"Template: {tpl_path}",
             ""]
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
    overall = "PASS" if all(r.status in ("PASS", "SKIP", "WARN") for r in results) else "FAIL"
    # WARN-only is still overall PASS
    has_fail = any(r.status == "FAIL" for r in results)
    lines.append(f"Overall: {'FAIL' if has_fail else 'PASS'}")
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(
        description="Review a formatted DOCX against a template YAML"
    )
    parser.add_argument("formatted_docx", help="Formatted DOCX file to review")
    parser.add_argument("template_file", nargs="?",
                        help="Template YAML file (optional if --category is given)")
    parser.add_argument("--category", "-c",
                        help="Template name; auto-lookup in styles/<name>/")
    parser.add_argument("--output", "-o",
                        help="Write report to file (also printed to terminal)")
    parser.add_argument("--json", action="store_true",
                        help="Output report as JSON (machine-readable)")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="Show full diff details even on PASS")
    args = parser.parse_args()

    results = review_docx(
        args.formatted_docx,
        args.template_file,
        category=args.category,
        verbose=args.verbose,
        json_output=args.json,
        output_file=args.output,
    )

    # Resolve paths for display
    docx_path = Path(args.formatted_docx).resolve()
    tpl_path = "unknown"
    if args.template_file:
        tpl_path = str(Path(args.template_file).resolve())
    elif args.category:
        styles_dir = _PROJECT_ROOT / "styles" / args.category
        yaml_files = sorted(styles_dir.glob("*.yaml")) + sorted(styles_dir.glob("*.yml"))
        if yaml_files:
            tpl_path = str(yaml_files[0])

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
        output = _format_report(results, str(docx_path), tpl_path, args.verbose)

    # Print to terminal
    print(output)

    # Write to file if requested
    if args.output:
        out_path = Path(args.output)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(output, encoding="utf-8")
        print(f"\nReport saved to {out_path}")


if __name__ == "__main__":
    main()
