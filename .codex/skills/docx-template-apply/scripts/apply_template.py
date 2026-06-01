#!/usr/bin/env python3
"""Apply a formatting template to DOCX documents.

Two levels of parallelism:

  1. Per-document parallel XML processing (default):
     styles, page layout, numbering, theme, fontTable are
     modified in parallel threads for each document.

  2. Multi-document batch mode (--batch):
     Multiple DOCX files are processed in parallel. Accepts
     a glob pattern (e.g. "chapter*.docx") or multiple paths.

Usage (single document):
    python apply_template.py target.docx template.yaml [--output result.docx]

Usage (batch mode — parallel per-document):
    python apply_template.py --batch "chapter*.docx" template.yaml --output-dir ./out
    python apply_template.py --batch doc1.docx doc2.docx template.yaml --jobs 4

Dependencies:
    - pip install defusedxml pyyaml
    - ~/.claude/skills/docx/scripts/office/unpack.py
    - ~/.claude/skills/docx/scripts/office/pack.py
"""

import argparse
import concurrent.futures
import glob as glob_module
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import defusedxml.minidom
import yaml

# --- Paths ---

UNPACK_PY = os.path.expanduser(
    "~/.claude/skills/docx/scripts/office/unpack.py"
)
PACK_PY = os.path.expanduser(
    "~/.claude/skills/docx/scripts/office/pack.py"
)

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent.parent.parent

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _create_w(doc, local_tag: str):
    return doc.createElementNS(NS["w"], f"w:{local_tag}")


def _create_a(doc, local_tag: str):
    return doc.createElementNS(A_NS, f"a:{local_tag}")


# --- Helpers ---


def _ensure_docx_skill():
    for script in (UNPACK_PY, PACK_PY):
        if not Path(script).exists():
            print(
                f"Error: Required script not found: {script}\n"
                "Install the docx skill first.",
                file=sys.stderr,
            )
            sys.exit(1)


def _unpack(input_file: str, output_dir: str):
    subprocess.run(
        [sys.executable, UNPACK_PY, input_file, output_dir,
         "--merge-runs", "false", "--simplify-redlines", "false"],
        check=True, capture_output=True, text=True,
    )


def _pack(input_dir: str, output_file: str, original: str | None = None):
    cmd = [sys.executable, PACK_PY, input_dir, output_file]
    if original:
        cmd.extend(["--original", original])
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        if result.stdout:
            print(result.stdout, file=sys.stderr)
        if result.stderr:
            print(result.stderr, file=sys.stderr)
        raise subprocess.CalledProcessError(
            result.returncode, cmd, output=result.stdout, stderr=result.stderr
        )


def _parse_xml(path: Path):
    return defusedxml.minidom.parseString(path.read_text(encoding="utf-8"))


def _write_bytes(dom, path: Path):
    path.write_bytes(dom.toxml(encoding="UTF-8"))


def _get_attr(elem, attr: str) -> str | None:
    val = elem.getAttribute(attr)
    return val if val else None


def _set_attr(elem, attr: str, val):
    elem.setAttribute(attr, str(val))


def _ensure_element(parent, tag: str) -> object:
    children = parent.getElementsByTagName(tag)
    if children:
        return children[0]
    doc = parent.ownerDocument
    if tag.startswith("w:"):
        new_el = _create_w(doc, tag[2:])
    elif tag.startswith("a:"):
        new_el = _create_a(doc, tag[2:])
    else:
        new_el = doc.createElementNS(None, tag)
    parent.appendChild(new_el)
    return new_el


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


def _ensure_direct_element(parent, tag: str):
    existing = _direct_child(parent, tag)
    if existing:
        return existing
    doc = parent.ownerDocument
    if tag.startswith("w:"):
        new_el = _create_w(doc, tag[2:])
    elif tag.startswith("a:"):
        new_el = _create_a(doc, tag[2:])
    else:
        new_el = doc.createElementNS(None, tag)
    parent.appendChild(new_el)
    return new_el


def _clear_element(el):
    while el.firstChild:
        el.removeChild(el.firstChild)


# --- Template Application Functions ---


def apply_doc_defaults(styles_dom, defaults: dict):
    if not defaults:
        return False
    doc = styles_dom
    dd_list = styles_dom.getElementsByTagName("w:docDefaults")
    if dd_list:
        dd = dd_list[0]
    else:
        dd = _create_w(doc, "docDefaults")
        styles_root = styles_dom.getElementsByTagName("w:styles")[0]
        styles_root.insertBefore(dd, styles_root.firstChild)

    modified = False
    run_def = defaults.get("run", {})
    if run_def:
        rpr_default = _ensure_element(dd, "w:rPrDefault")
        rpr = _ensure_element(rpr_default, "w:rPr")
        _apply_run_props(rpr, run_def)
        modified = True

    para_def = defaults.get("paragraph", {})
    if para_def:
        ppr_default = _ensure_element(dd, "w:pPrDefault")
        ppr = _ensure_element(ppr_default, "w:pPr")
        _apply_para_props(ppr, para_def)
        modified = True
    return modified


def _apply_run_props(rpr, props: dict):
    if "font" in props or "fontEastAsia" in props:
        rfonts = _ensure_element(rpr, "w:rFonts")
        if "font" in props:
            _set_attr(rfonts, "w:ascii", props["font"])
            _set_attr(rfonts, "w:hAnsi", props["font"])
            _set_attr(rfonts, "w:cs", props["font"])
        if "fontEastAsia" in props:
            _set_attr(rfonts, "w:eastAsia", props["fontEastAsia"])
    if "fontSize" in props:
        sz = _ensure_element(rpr, "w:sz")
        _set_attr(sz, "w:val", str(props["fontSize"]))
    if "fontSizeCs" in props:
        sz_cs = _ensure_element(rpr, "w:szCs")
        _set_attr(sz_cs, "w:val", str(props["fontSizeCs"]))
    if "fontColor" in props:
        color = _ensure_element(rpr, "w:color")
        _set_attr(color, "w:val", props["fontColor"])
    if props.get("bold"):
        _ensure_element(rpr, "w:b")
    else:
        for el in rpr.getElementsByTagName("w:b"):
            rpr.removeChild(el)
    if props.get("italic"):
        _ensure_element(rpr, "w:i")
    else:
        for el in rpr.getElementsByTagName("w:i"):
            rpr.removeChild(el)
    if "underline" in props:
        u = _ensure_element(rpr, "w:u")
        _set_attr(u, "w:val", props["underline"])
    if props.get("smallCaps"):
        _ensure_element(rpr, "w:smallCaps")
    if props.get("caps"):
        _ensure_element(rpr, "w:caps")
    if props.get("strike"):
        _ensure_element(rpr, "w:strike")
    if "fontSpacing" in props:
        sp = _ensure_element(rpr, "w:spacing")
        _set_attr(sp, "w:val", str(props["fontSpacing"]))


def _apply_para_props(ppr, props: dict):
    """Set paragraph properties on a w:pPr element in correct OOXML order.

    CT_PPr sequence:
      pStyle, keepNext, keepLines, pageBreakBefore, widowControl,
      numPr, spacing, ind, jc, outlineLvl, ...
    """
    # Clear and rebuild in schema order
    _clear_element(ppr)
    doc = ppr.ownerDocument

    if props.get("keepNext"):
        _ensure_element(ppr, "w:keepNext")
    if props.get("pageBreakBefore"):
        _ensure_element(ppr, "w:pageBreakBefore")
    if props.get("widowControl"):
        _ensure_element(ppr, "w:widowControl")

    spacing_attrs = {k[8:]: v for k, v in props.items() if k.startswith("spacing_")}
    if spacing_attrs:
        sp = _ensure_element(ppr, "w:spacing")
        for attr, val in spacing_attrs.items():
            _set_attr(sp, f"w:{attr}", str(val))

    indent_attrs = {k[7:]: v for k, v in props.items() if k.startswith("indent_")}
    if indent_attrs:
        ind = _ensure_element(ppr, "w:ind")
        for attr, val in indent_attrs.items():
            _set_attr(ind, f"w:{attr}", str(val))

    if "alignment" in props:
        jc = _ensure_element(ppr, "w:jc")
        _set_attr(jc, "w:val", props["alignment"])

    if "outlineLevel" in props and props["outlineLevel"] is not None:
        ol = _ensure_element(ppr, "w:outlineLvl")
        _set_attr(ol, "w:val", str(props["outlineLevel"]))


_BODY_PARA_FORMAT_TAGS = {
    "w:spacing", "w:ind", "w:jc", "w:outlineLvl",
    "w:keepNext", "w:pageBreakBefore", "w:widowControl", "w:keepLines",
    "w:shd", "w:rPr",
}


def _remove_direct_children(parent, tags: set[str]):
    for child in list(parent.childNodes):
        if child.nodeType == child.ELEMENT_NODE and child.tagName in tags:
            parent.removeChild(child)


def _ensure_paragraph_ppr(p):
    ppr = _direct_child(p, "w:pPr")
    if ppr:
        return ppr
    ppr = _create_w(p.ownerDocument, "pPr")
    if p.firstChild:
        p.insertBefore(ppr, p.firstChild)
    else:
        p.appendChild(ppr)
    return ppr


def _ensure_run_rpr(r):
    rpr = _direct_child(r, "w:rPr")
    if rpr:
        return rpr
    rpr = _create_w(r.ownerDocument, "rPr")
    if r.firstChild:
        r.insertBefore(rpr, r.firstChild)
    else:
        r.appendChild(rpr)
    return rpr


def _apply_para_props_overlay(ppr, props: dict):
    """Apply body-level paragraph formatting without removing pStyle/numPr."""
    if not props:
        return
    _remove_direct_children(ppr, _BODY_PARA_FORMAT_TAGS)
    doc = ppr.ownerDocument

    if props.get("keepNext"):
        ppr.appendChild(_create_w(doc, "keepNext"))
    if props.get("pageBreakBefore"):
        ppr.appendChild(_create_w(doc, "pageBreakBefore"))
    if props.get("widowControl"):
        ppr.appendChild(_create_w(doc, "widowControl"))
    if props.get("keepLines"):
        ppr.appendChild(_create_w(doc, "keepLines"))

    spacing_attrs = {
        "before": props.get("spacing_before"),
        "after": props.get("spacing_after"),
        "line": props.get("spacing_line"),
        "lineRule": props.get("spacing_lineRule"),
    }
    if any(v is not None for v in spacing_attrs.values()):
        spacing = _create_w(doc, "spacing")
        for attr, val in spacing_attrs.items():
            if val is not None:
                _set_attr(spacing, f"w:{attr}", str(val))
        ppr.appendChild(spacing)

    indent_attrs = {
        "left": props.get("indent_left"),
        "right": props.get("indent_right"),
        "firstLine": props.get("indent_firstLine"),
        "hanging": props.get("indent_hanging"),
    }
    if any(v is not None for v in indent_attrs.values()):
        ind = _create_w(doc, "ind")
        for attr, val in indent_attrs.items():
            if val is not None:
                _set_attr(ind, f"w:{attr}", str(val))
        ppr.appendChild(ind)

    if "alignment" in props:
        jc = _create_w(doc, "jc")
        _set_attr(jc, "w:val", props["alignment"])
        ppr.appendChild(jc)

    if "outlineLevel" in props:
        ol = _create_w(doc, "outlineLvl")
        _set_attr(ol, "w:val", str(props["outlineLevel"]))
        ppr.appendChild(ol)

    if "shading" in props:
        shd = _create_w(doc, "shd")
        for attr, val in (props.get("shading") or {}).items():
            _set_attr(shd, f"w:{attr}", val)
        ppr.appendChild(shd)


def apply_body_paragraph_formats(doc_dom, body_formats: list) -> bool:
    """Apply per-paragraph formatting captured from the reference body.

    Matching is index-based so legacy templates that encode title, abstract,
    heading, caption, and body fonts as direct formatting can be reproduced.
    """
    if not body_formats:
        return False
    body_list = doc_dom.getElementsByTagName("w:body")
    if not body_list:
        return False
    paragraphs = _direct_children(body_list[0], "w:p")
    modified = False

    for fmt in body_formats:
        idx = fmt.get("index")
        if idx is None or idx < 0 or idx >= len(paragraphs):
            continue
        p = paragraphs[idx]
        ppr = _ensure_paragraph_ppr(p)

        para_props = fmt.get("paragraph") or {}
        if para_props:
            _apply_para_props_overlay(ppr, para_props)
            modified = True

        para_run_props = fmt.get("paragraph_run") or {}
        if para_run_props:
            paragraph_rpr = _ensure_direct_element(ppr, "w:rPr")
            _clear_element(paragraph_rpr)
            _apply_run_props(paragraph_rpr, para_run_props)
            modified = True

        default_run_props = fmt.get("run") or {}
        runs_by_index = {
            item.get("index"): item.get("run") or {}
            for item in fmt.get("runs", [])
            if item.get("index") is not None
        }
        if default_run_props or runs_by_index:
            for run_index, r in enumerate(_direct_children(p, "w:r")):
                run_props = runs_by_index.get(run_index) or default_run_props
                if not run_props:
                    continue
                rpr = _ensure_run_rpr(r)
                _clear_element(rpr)
                _apply_run_props(rpr, run_props)
                modified = True

    return modified


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
        {
            "role": "cn_title",
            "match": {"type": "front_index", "index": 0},
            "paragraph": {"alignment": "center", "spacing_line": 360, "spacing_lineRule": "auto"},
            "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 40, "fontSizeCs": 40, "bold": True},
        },
        {
            "role": "cn_authors",
            "match": {"type": "front_index", "index": 1},
            "paragraph": {"alignment": "center"},
            "run": {"font": "楷体_GB2312;楷体", "fontEastAsia": "楷体_GB2312;楷体", "fontSize": 28, "fontSizeCs": 28},
        },
        {
            "role": "cn_affiliation",
            "match": {"type": "front_index", "index": 2},
            "paragraph": {"alignment": "center"},
            "run": {"font": "Times New Roman", "fontEastAsia": "宋体;SimSun", "fontSize": 18, "fontSizeCs": 18},
        },
        {
            "role": "cn_abstract",
            "match": {"type": "regex", "pattern": r"^摘\s*要[:：]"},
            "paragraph": {"alignment": "both", "indent_left": 425, "indent_right": 425},
            "run": {"font": "Times New Roman", "fontEastAsia": "楷体_GB2312;楷体", "fontSize": 21, "fontSizeCs": 21},
            "prefixes": [
                {"text": "摘 要：", "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}},
                {"text": "摘  要：", "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}},
                {"text": "摘要：", "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}},
            ],
        },
        {
            "role": "cn_keywords",
            "match": {"type": "regex", "pattern": r"^关键词[:：]"},
            "paragraph": {"alignment": "both", "indent_left": 425, "indent_right": 425},
            "run": {"font": "Times New Roman", "fontEastAsia": "楷体_GB2312;楷体", "fontSize": 21, "fontSizeCs": 21},
            "prefixes": [
                {"text": "关键词：", "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}},
                {"text": "关键词:", "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21}},
            ],
        },
        {
            "role": "classification",
            "match": {"type": "regex", "pattern": r"^中图分类号"},
            "paragraph": {"alignment": "both"},
            "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21},
        },
        {
            "role": "en_title",
            "match": {"type": "front_index", "index": 6},
            "paragraph": {"alignment": "center"},
            "run": {"font": "Times New Roman", "fontSize": 32, "fontSizeCs": 32, "bold": True},
        },
        {
            "role": "en_authors",
            "match": {"type": "front_index", "index": 7},
            "paragraph": {"alignment": "center"},
            "run": {"font": "Times New Roman", "fontSize": 28, "fontSizeCs": 28, "italic": True},
        },
        {
            "role": "en_affiliation",
            "match": {"type": "front_index", "index": 8},
            "paragraph": {"alignment": "center"},
            "run": {"font": "Times New Roman", "fontSize": 18, "fontSizeCs": 18},
        },
        {
            "role": "en_abstract",
            "match": {"type": "regex", "pattern": r"^Abstract[:：]"},
            "paragraph": {"alignment": "both"},
            "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21},
        },
        {
            "role": "en_keywords",
            "match": {"type": "regex", "pattern": r"^(Key words|Keywords)[:：]"},
            "paragraph": {"alignment": "both"},
            "run": {"font": "Times New Roman", "fontSize": 21, "fontSizeCs": 21},
        },
        {
            "role": "references_heading",
            "match": {"type": "regex", "pattern": r"^参考文献"},
            "paragraph": {"alignment": "left"},
            "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21, "bold": True},
            "sets_context": {"after_references": True},
        },
        {
            "role": "reference_entry",
            "match": {"type": "after_references_regex", "pattern": r"^\[?\d+\]"},
            "paragraph": {"alignment": "both"},
            "run": {"font": "Times New Roman", "fontEastAsia": "宋体;SimSun", "fontSize": 18, "fontSizeCs": 18},
        },
        {
            "role": "section_heading",
            "match": {"type": "regex", "pattern": r"^\d+\s+\S"},
            "paragraph": {"alignment": "left", "spacing_line": 220, "spacing_lineRule": "exact"},
            "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 28, "fontSizeCs": 28},
        },
        {
            "role": "subsection_heading",
            "match": {"type": "regex", "pattern": r"^\d+\.\d+(\.\d+)?\s+\S"},
            "paragraph": {"alignment": "left"},
            "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 21, "fontSizeCs": 21, "bold": True},
        },
        {
            "role": "figure_caption",
            "match": {"type": "regex", "pattern": r"^(图\s*\d+|Fig\.\s*\d+)"},
            "paragraph": {"alignment": "center"},
            "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 18, "fontSizeCs": 18},
        },
        {
            "role": "table_caption",
            "match": {"type": "regex", "pattern": r"^(表\s*\d+|Tab\.\s*\d+)"},
            "paragraph": {"alignment": "center"},
            "run": {"font": "黑体;SimHei", "fontEastAsia": "黑体;SimHei", "fontSize": 18, "fontSizeCs": 18, "bold": True},
        },
        {
            "role": "table_cell",
            "match": {"type": "in_table"},
            "paragraph": {"alignment": "center"},
            "run": {"font": "Times New Roman", "fontEastAsia": "宋体;SimSun", "fontSize": 18, "fontSizeCs": 18, "bold": True},
        },
        {
            "role": "body",
            "match": {"type": "default"},
            "paragraph": {"alignment": "both", "indent_firstLine": 420, "spacing_line": 314, "spacing_lineRule": "exact"},
            "run": base_body,
        },
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


_EXACT_RUN_TAGS = {
    "w:rFonts", "w:sz", "w:szCs", "w:color", "w:b", "w:i", "w:u",
    "w:strike", "w:smallCaps", "w:caps", "w:spacing",
}


def _apply_run_props_exact(rpr, props: dict):
    _remove_direct_children(rpr, _EXACT_RUN_TAGS)
    _apply_run_props(rpr, props)


def _split_run_at_text_offset(run, offset: int):
    texts = run.getElementsByTagName("w:t")
    if not texts or not texts[0].firstChild:
        return None
    value = texts[0].firstChild.nodeValue
    if offset <= 0 or offset >= len(value):
        return None
    texts[0].firstChild.nodeValue = value[:offset]
    clone = run.cloneNode(deep=True)
    clone_texts = clone.getElementsByTagName("w:t")
    if clone_texts and clone_texts[0].firstChild:
        clone_texts[0].firstChild.nodeValue = value[offset:]
    parent = run.parentNode
    if run.nextSibling:
        parent.insertBefore(clone, run.nextSibling)
    else:
        parent.appendChild(clone)
    return clone


def _apply_prefix_formats(p, rule: dict):
    text = _paragraph_text(p)
    prefix_rule = None
    for candidate in rule.get("prefixes", []):
        if text.startswith(candidate.get("text", "")):
            prefix_rule = candidate
            break
    if not prefix_rule:
        return False

    remaining = len(prefix_rule["text"])
    modified = False
    for r in list(_direct_children(p, "w:r")):
        if remaining <= 0:
            break
        run_text = _run_text(r)
        if not run_text:
            continue
        if len(run_text) > remaining:
            _split_run_at_text_offset(r, remaining)
            run_text = _run_text(r)
        rpr = _ensure_run_rpr(r)
        _apply_run_props_exact(rpr, prefix_rule["run"])
        remaining -= len(run_text)
        modified = True
    return modified


def apply_semantic_format_rules(doc_dom, template: dict) -> bool:
    rules = _semantic_rules(template)
    if not rules:
        return False
    body = doc_dom.getElementsByTagName("w:body")
    if not body:
        return False

    modified = False
    front_index = -1
    context = {"after_references": False}
    paragraphs = doc_dom.getElementsByTagName("w:p")
    for p in paragraphs:
        text = _paragraph_text(p).strip()
        if not text:
            continue
        in_table = _has_ancestor(p, "w:tc")
        if not in_table:
            front_index += 1
        rule = _select_semantic_rule(rules, text, front_index if not in_table else None, in_table, context)
        if not rule:
            continue

        ppr = _ensure_paragraph_ppr(p)
        if rule.get("paragraph"):
            _apply_para_props_overlay(ppr, rule["paragraph"])
            modified = True
        if rule.get("run"):
            paragraph_rpr = _ensure_direct_element(ppr, "w:rPr")
            _apply_run_props_exact(paragraph_rpr, rule["run"])
            for r in _direct_children(p, "w:r"):
                if not _run_text(r):
                    continue
                rpr = _ensure_run_rpr(r)
                _apply_run_props_exact(rpr, rule["run"])
            modified = True
        if _apply_prefix_formats(p, rule):
            modified = True

        for key, value in (rule.get("sets_context") or {}).items():
            context[key] = value

    return modified


def apply_styles(styles_dom, template: dict):
    modified = False
    doc = styles_dom
    styles_root = styles_dom.getElementsByTagName("w:styles")[0]

    existing_by_id = {}
    for style in styles_root.getElementsByTagName("w:style"):
        sid = _get_attr(style, "w:styleId")
        if sid:
            existing_by_id[sid] = style

    def _get_name(style_data: dict) -> str | None:
        return style_data.get("name")

    for ps in template.get("paragraph_styles", []):
        style_id = ps.get("id")
        style_name = _get_name(ps)
        if not style_id:
            continue
        existing = existing_by_id.get(style_id)
        if existing and style_name:
            # Overwrite only if same-named (same purpose)
            existing_name = ""
            n = existing.getElementsByTagName("w:name")
            if n:
                existing_name = _get_attr(n[0], "w:val") or ""
            if existing_name == style_name:
                _clear_element(existing)
                _set_attr(existing, "w:type", "paragraph")
                _set_attr(existing, "w:styleId", style_id)
                _append_style_children(existing, ps)
                modified = True
        elif existing and not style_name:
            _clear_element(existing)
            _set_attr(existing, "w:type", "paragraph")
            _set_attr(existing, "w:styleId", style_id)
            _append_style_children(existing, ps)
            modified = True
        elif not existing:
            style_el = _create_w(doc, "style")
            _set_attr(style_el, "w:type", "paragraph")
            _set_attr(style_el, "w:styleId", style_id)
            styles_root.appendChild(style_el)
            _append_style_children(style_el, ps)
            modified = True

    for cs in template.get("character_styles", []):
        style_id = cs.get("id")
        style_name = _get_name(cs)
        if not style_id:
            continue
        existing = existing_by_id.get(style_id)
        if existing and style_name:
            existing_name = ""
            n = existing.getElementsByTagName("w:name")
            if n:
                existing_name = _get_attr(n[0], "w:val") or ""
            if existing_name == style_name:
                _clear_element(existing)
                _set_attr(existing, "w:type", "character")
                _set_attr(existing, "w:styleId", style_id)
                _append_style_children(existing, cs)
                modified = True
        elif existing and not style_name:
            _clear_element(existing)
            _set_attr(existing, "w:type", "character")
            _set_attr(existing, "w:styleId", style_id)
            _append_style_children(existing, cs)
            modified = True
        elif not existing and style_name:
            style_el = _create_w(doc, "style")
            _set_attr(style_el, "w:type", "character")
            _set_attr(style_el, "w:styleId", style_id)
            styles_root.appendChild(style_el)
            _append_style_children(style_el, cs)
            modified = True
        elif not existing:
            style_el = _create_w(doc, "style")
            _set_attr(style_el, "w:type", "character")
            _set_attr(style_el, "w:styleId", style_id)
            styles_root.appendChild(style_el)
            _append_style_children(style_el, cs)
            modified = True

    return modified


def _append_style_children(style_el, style_data: dict):
    doc = style_el.ownerDocument
    if style_data.get("name"):
        name_el = _create_w(doc, "name")
        _set_attr(name_el, "w:val", style_data["name"])
        style_el.appendChild(name_el)
    if style_data.get("basedOn"):
        bo = _create_w(doc, "basedOn")
        _set_attr(bo, "w:val", style_data["basedOn"])
        style_el.appendChild(bo)
    if style_data.get("next"):
        nx = _create_w(doc, "next")
        _set_attr(nx, "w:val", style_data["next"])
        style_el.appendChild(nx)
    para = style_data.get("paragraph", {})
    if para:
        ppr = _create_w(doc, "pPr")
        style_el.appendChild(ppr)
        _apply_para_props(ppr, para)
    run = style_data.get("run", {})
    if run:
        rpr = _create_w(doc, "rPr")
        style_el.appendChild(rpr)
        _apply_run_props(rpr, run)


def apply_page_layout(doc_dom, page: dict):
    if not page:
        return False
    modified = False
    for sect_pr in doc_dom.getElementsByTagName("w:sectPr"):
        if "width" in page or "height" in page or "orientation" in page:
            pg_sz = _ensure_element(sect_pr, "w:pgSz")
            if "width" in page:
                _set_attr(pg_sz, "w:w", str(page["width"]))
            if "height" in page:
                _set_attr(pg_sz, "w:h", str(page["height"]))
            if "orientation" in page:
                _set_attr(pg_sz, "w:orient", page["orientation"])
            modified = True
        margins = page.get("margins", {})
        if margins:
            pg_mar_list = sect_pr.getElementsByTagName("w:pgMar")
            if pg_mar_list:
                pg_mar = pg_mar_list[0]
            else:
                pg_mar = _create_w(sect_pr.ownerDocument, "pgMar")
                sect_pr.insertBefore(pg_mar, sect_pr.firstChild)
            for attr in ("top", "bottom", "left", "right", "header", "footer", "gutter"):
                if attr in margins:
                    _set_attr(pg_mar, f"w:{attr}", str(margins[attr]))
            modified = True
    return modified


def apply_numbering(num_dom, numbering: list):
    if not numbering:
        return False
    doc = num_dom
    num_root = num_dom.getElementsByTagName("w:numbering")[0]
    _clear_element(num_root)

    num_parts = {}
    for num_def in numbering:
        anum_id = num_def.get("abstractNumId")
        if anum_id is None:
            continue
        if anum_id not in num_parts:
            num_parts[anum_id] = {"abstract": None, "concrete": []}
        num_parts[anum_id]["concrete"].append(num_def)
        if num_parts[anum_id]["abstract"] is not None:
            continue

        anum_el = _create_w(doc, "abstractNum")
        _set_attr(anum_el, "w:abstractNumId", str(anum_id))
        for lvl in num_def.get("levels", []):
            lvl_el = _create_w(doc, "lvl")
            _set_attr(lvl_el, "w:ilvl", str(lvl.get("level", 0)))
            if "start" in lvl:
                el = _create_w(doc, "start")
                _set_attr(el, "w:val", str(lvl["start"]))
                lvl_el.appendChild(el)
            if "format" in lvl:
                el = _create_w(doc, "numFmt")
                _set_attr(el, "w:val", lvl["format"])
                lvl_el.appendChild(el)
            if "text" in lvl:
                el = _create_w(doc, "lvlText")
                _set_attr(el, "w:val", lvl["text"])
                lvl_el.appendChild(el)
            if "alignment" in lvl:
                el = _create_w(doc, "lvlJc")
                _set_attr(el, "w:val", lvl["alignment"])
                lvl_el.appendChild(el)
            indent_attrs = {k[7:]: v for k, v in lvl.items() if k.startswith("indent_")}
            if indent_attrs:
                ppr_el = _create_w(doc, "pPr")
                ind_el = _create_w(doc, "ind")
                for attr, val in indent_attrs.items():
                    _set_attr(ind_el, f"w:{attr}", str(val))
                ppr_el.appendChild(ind_el)
                lvl_el.appendChild(ppr_el)
            font_attrs = {k[5:]: v for k, v in lvl.items() if k.startswith("font_")}
            has_font = font_attrs or "fontSize" in lvl
            if has_font:
                rpr_el = _create_w(doc, "rPr")
                if font_attrs:
                    rfonts = _create_w(doc, "rFonts")
                    for attr, val in font_attrs.items():
                        _set_attr(rfonts, f"w:{attr}", val)
                    rpr_el.appendChild(rfonts)
                if "fontSize" in lvl:
                    sz = _create_w(doc, "sz")
                    _set_attr(sz, "w:val", str(lvl["fontSize"]))
                    rpr_el.appendChild(sz)
                lvl_el.appendChild(rpr_el)
            anum_el.appendChild(lvl_el)
        num_parts[anum_id]["abstract"] = anum_el

    for anum_id in sorted(num_parts.keys()):
        if num_parts[anum_id]["abstract"] is not None:
            num_root.appendChild(num_parts[anum_id]["abstract"])
    for anum_id in sorted(num_parts.keys()):
        for num_def in num_parts[anum_id]["concrete"]:
            num_el = _create_w(doc, "num")
            _set_attr(num_el, "w:numId", str(num_def.get("id", 0)))
            anum_ref = _create_w(doc, "abstractNumId")
            _set_attr(anum_ref, "w:val", str(anum_id))
            num_el.appendChild(anum_ref)
            for lvl_idx, override in num_def.get("levelOverrides", {}).items():
                lo = _create_w(doc, "lvlOverride")
                _set_attr(lo, "w:ilvl", str(lvl_idx))
                so = _create_w(doc, "startOverride")
                _set_attr(so, "w:val", str(override.get("start", 1)))
                lo.appendChild(so)
                num_el.appendChild(lo)
            num_root.appendChild(num_el)
    return True


def apply_theme(theme_dom, theme: dict):
    if not theme:
        return False
    modified = False
    te_list = theme_dom.getElementsByTagName("a:themeElements")
    if not te_list:
        return False
    te = te_list[0]
    doc = te.ownerDocument

    colors = theme.get("colors", {})
    if colors:
        for ex in te.getElementsByTagName("a:clrScheme"):
            te.removeChild(ex)
        cs = _create_a(doc, "clrScheme")
        _set_attr(cs, "name", theme.get("schemeName", "Custom"))
        for tag_name, color_val in colors.items():
            child = doc.createElementNS(A_NS, tag_name)
            srgb = _create_a(doc, "srgbClr")
            _set_attr(srgb, "val", color_val)
            child.appendChild(srgb)
            cs.appendChild(child)
        te.insertBefore(cs, te.firstChild)
        modified = True

    fonts = theme.get("fonts", {})
    if fonts:
        for ex in te.getElementsByTagName("a:fontScheme"):
            te.removeChild(ex)
        fs = _create_a(doc, "fontScheme")
        _set_attr(fs, "name", theme.get("fontSchemeName", "Custom"))
        major = _create_a(doc, "majorFont")
        major_latin = _create_a(doc, "latin")
        _set_attr(major_latin, "typeface", fonts.get("major", "Calibri Light"))
        major.appendChild(major_latin)
        major.appendChild(_create_a(doc, "ea"))
        fs.appendChild(major)
        minor = _create_a(doc, "minorFont")
        minor_latin = _create_a(doc, "latin")
        _set_attr(minor_latin, "typeface", fonts.get("minor", "Calibri"))
        minor.appendChild(minor_latin)
        minor.appendChild(_create_a(doc, "ea"))
        fs.appendChild(minor)
        te.appendChild(fs)
        modified = True
    return modified


def apply_font_table(font_table_dom, fonts: list):
    if not fonts:
        return False
    modified = False
    doc = font_table_dom
    root = font_table_dom.getElementsByTagName("w:fonts")[0]
    existing = {_get_attr(f, "w:name") for f in root.getElementsByTagName("w:font")}
    for font_name in fonts:
        if font_name not in existing:
            font_el = _create_w(doc, "font")
            _set_attr(font_el, "w:name", font_name)
            root.appendChild(font_el)
            modified = True
    return modified


# --- Version and Template Resolution ---


def _parse_version_dir(name: str) -> tuple[int, int] | None:
    """Parse a version directory name like v1.2 into (major, minor)."""
    if not name.startswith("v"):
        return None
    parts = name[1:].split(".", 1)
    if len(parts) == 1:
        try:
            return (int(parts[0]), 1)
        except ValueError:
            return None
    try:
        return (int(parts[0]), int(parts[1]))
    except ValueError:
        return None


def _detect_next_version(category: str, paper: str | None = None) -> str:
    """Find the next version string for a given category and optional paper.

    Looks for existing v{major}.{minor} directories and returns
    f\"{max_major + 1}.1\" (new major version, minor starts at 1).
    """
    if paper:
        outputs_dir = _PROJECT_ROOT / "outputs" / paper / category
    else:
        outputs_dir = _PROJECT_ROOT / "outputs" / category
    if not outputs_dir.exists():
        return "1.1"
    max_major = 0
    for d in outputs_dir.iterdir():
        if d.is_dir():
            parsed = _parse_version_dir(d.name)
            if parsed is not None:
                max_major = max(max_major, parsed[0])
    return f"{max_major + 1}.1"


def _resolve_template(template_file: str | None, category: str | None) -> str:
    """Resolve template YAML path from explicit arg or category lookup."""
    if template_file:
        return template_file

    if not category:
        print(
            "Error: template_file is required when --category is not specified",
            file=sys.stderr,
        )
        sys.exit(1)

    styles_dir = _PROJECT_ROOT / "styles" / category
    if not styles_dir.exists():
        print(
            f"Error: styles/{category}/ directory not found: {styles_dir}",
            file=sys.stderr,
        )
        sys.exit(1)

    yaml_files = sorted(styles_dir.glob("*.yaml")) + sorted(styles_dir.glob("*.yml"))
    if not yaml_files:
        print(
            f"Error: no .yaml files found in styles/{category}/",
            file=sys.stderr,
        )
        sys.exit(1)
    if len(yaml_files) > 1:
        print(
            f"Error: multiple templates found in styles/{category}/. "
            f"Please specify one explicitly:\n  "
            + "\n  ".join(str(f) for f in yaml_files),
            file=sys.stderr,
        )
        sys.exit(1)

    return str(yaml_files[0])


# --- Body Formatting Cleanup ---

# Run-level direct formatting tags to strip (these override style definitions)
_RUN_FORMAT_TAGS = {
    "w:rFonts", "w:sz", "w:szCs", "w:color",
    "w:b", "w:i", "w:u", "w:strike",
    "w:smallCaps", "w:caps", "w:spacing",
}

# Paragraph-level direct formatting tags to strip
_PARA_FORMAT_TAGS = {
    "w:spacing", "w:ind", "w:jc", "w:rPr",
}


def _strip_run_formatting(dom):
    """Remove direct run formatting from all paragraphs in document body.

    Strips w:rPr child elements that override style-defined font, size,
    color, bold, italic, underline, strikethrough, caps, and spacing.
    """
    modified_any = False
    for p in dom.getElementsByTagName("w:p"):
        for r in p.getElementsByTagName("w:r"):
            rpr_list = r.getElementsByTagName("w:rPr")
            if not rpr_list:
                continue
            rpr = rpr_list[0]
            children = list(rpr.childNodes)
            for child in children:
                if child.nodeType == child.ELEMENT_NODE and child.tagName in _RUN_FORMAT_TAGS:
                    rpr.removeChild(child)
                    modified_any = True
            # Remove empty w:rPr
            if not rpr.childNodes or all(
                n.nodeType == n.TEXT_NODE and not n.nodeValue.strip()
                for n in rpr.childNodes
            ):
                rpr.parentNode.removeChild(rpr)
    return modified_any


def _strip_para_formatting(dom):
    """Remove direct paragraph formatting from all paragraphs in document body.

    Strips w:pPr child elements that override style-defined spacing,
    indentation, and alignment. Preserves pStyle, numPr, outlineLvl,
    keepNext, pageBreakBefore, widowControl, and other structural props.
    """
    preserved_tags = {
        "w:pStyle", "w:numPr", "w:outlineLvl",
        "w:keepNext", "w:pageBreakBefore", "w:widowControl",
        "w:keepLines", "w:suppressAutoHyphens",
    }

    modified_any = False
    for p in dom.getElementsByTagName("w:p"):
        ppr_list = p.getElementsByTagName("w:pPr")
        if not ppr_list:
            continue
        ppr = ppr_list[0]
        children = list(ppr.childNodes)
        for child in children:
            if child.nodeType == child.ELEMENT_NODE:
                if child.tagName in _PARA_FORMAT_TAGS:
                    ppr.removeChild(child)
                    modified_any = True
        # Remove empty w:pPr (only preserved tags remain = keep it)
    return modified_any


# --- Per-document pipeline ---


def _make_numbering_dom(numbering: list):
    """Create a minimal w:numbering DOM from template data."""
    ns_w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    parts = []
    seen_abstract = set()
    for num_def in numbering:
        anum_id = num_def.get("abstractNumId")
        if anum_id is not None and anum_id not in seen_abstract:
            parts.append(f'<w:abstractNum w:abstractNumId="{anum_id}">')
            for lvl in num_def.get("levels", []):
                parts.append(f'<w:lvl w:ilvl="{lvl.get("level", 0)}">')
                if "start" in lvl:
                    parts.append(f'<w:start w:val="{lvl["start"]}"/>')
                if "format" in lvl:
                    parts.append(f'<w:numFmt w:val="{lvl["format"]}"/>')
                if "text" in lvl and lvl["text"]:
                    parts.append(f'<w:lvlText w:val="{lvl["text"]}"/>')
                parts.append("</w:lvl>")
            parts.append("</w:abstractNum>")
            seen_abstract.add(anum_id)

    for num_def in numbering:
        parts.append(
            f'<w:num w:numId="{num_def.get("id", 0)}">'
            f'<w:abstractNumId w:val="{num_def.get("abstractNumId", 0)}"/>'
            f"</w:num>"
        )

    xml_str = f'<w:numbering {ns_w}>{"".join(parts)}</w:numbering>'
    return defusedxml.minidom.parseString(xml_str)


def _process_one_document(
    target_path: Path,
    template: dict,
    output_path: Path | None,
    *,
    chunks: int = 1,
    preserve_formatting: bool = False,
    apply_body_formats: bool = False,
):
    """Run the full apply-template pipeline for a single document.

    Within-document parallelism:
      - level 1: XML files (styles, page, numbering, theme, fonts) in parallel
      - level 2: document body paragraph chunks when chunks > 1
    """
    if output_path is None:
        stem = target_path.stem
        output_path = target_path.with_stem(f"{stem}-formatted")

    with tempfile.TemporaryDirectory(prefix="docx-apply-") as tmp_dir:
        tmp_path = Path(tmp_dir)
        unpack_dir = tmp_path / "unpacked"

        _unpack(str(target_path), str(unpack_dir))
        word_dir = unpack_dir / "word"
        word_dir_str = str(word_dir)
        unpack_dir_str = str(unpack_dir)

        # --- Level 1: parallel XML file modifications ---
        def _task_styles():
            p = Path(word_dir_str) / "styles.xml"
            if not p.exists():
                return None
            dom = _parse_xml(p)
            msgs = []
            if apply_doc_defaults(dom, template.get("document_defaults", {})):
                msgs.append("document defaults")
            if apply_styles(dom, template):
                n_ps = len(template.get("paragraph_styles", []))
                n_cs = len(template.get("character_styles", []))
                msgs.append(f"{n_ps} paragraph styles, {n_cs} character styles")
            if msgs:
                _write_bytes(dom, p)
            return "; ".join(msgs) if msgs else None

        def _task_page():
            p = Path(word_dir_str) / "document.xml"
            if not p.exists():
                return None
            dom = _parse_xml(p)
            msgs = []

            # Apply page layout
            if "page" in template and apply_page_layout(dom, template["page"]):
                msgs.append("page layout")

            # Strip direct formatting (rPr/pPr) — must run in same task
            # to avoid race with parallel writers of document.xml
            has_body_formats = bool(template.get("body_paragraph_formats"))
            should_strip = (
                not preserve_formatting
                and (not has_body_formats or apply_body_formats)
            )
            if should_strip:
                run_modified = _strip_run_formatting(dom)
                para_modified = _strip_para_formatting(dom)
                if run_modified or para_modified:
                    parts = []
                    if run_modified:
                        parts.append("run formatting")
                    if para_modified:
                        parts.append("para formatting")
                    msgs.append("cleaned " + ", ".join(parts))

            if apply_body_formats and apply_body_paragraph_formats(dom, template.get("body_paragraph_formats", [])):
                msgs.append("body paragraph formats")

            if apply_semantic_format_rules(dom, template):
                msgs.append("semantic format rules")

            if msgs:
                _write_bytes(dom, p)
                return "; ".join(msgs)
            return None

        def _task_numbering():
            p = Path(word_dir_str) / "numbering.xml"
            if "numbering" not in template:
                return None
            if not p.exists():
                # Create minimal numbering.xml from scratch
                _write_bytes(
                    _make_numbering_dom(template["numbering"]), p
                )
                # Register in relationships
                rels_p = Path(word_dir_str) / "_rels" / "document.xml.rels"
                if rels_p.exists():
                    dom = _parse_xml(rels_p)
                    root = dom.documentElement
                    max_rid = 0
                    for rel in dom.getElementsByTagName("Relationship"):
                        rid = rel.getAttribute("Id")
                        if rid.startswith("rId"):
                            try:
                                max_rid = max(max_rid, int(rid[3:]))
                            except ValueError:
                                pass
                    rel_el = dom.createElement("Relationship")
                    rel_el.setAttribute("Id", f"rId{max_rid + 1}")
                    rel_el.setAttribute("Type",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering")
                    rel_el.setAttribute("Target", "numbering.xml")
                    root.appendChild(rel_el)
                    _write_bytes(dom, rels_p)
                # Register in content types
                ct_p = Path(unpack_dir_str) / "[Content_Types].xml"
                if ct_p.exists():
                    dom = _parse_xml(ct_p)
                    root = dom.documentElement
                    ov = dom.createElement("Override")
                    ov.setAttribute("PartName", "/word/numbering.xml")
                    ov.setAttribute("ContentType",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml")
                    root.appendChild(ov)
                    _write_bytes(dom, ct_p)
                return f"{len(template['numbering'])} numbering definitions (created)"
            dom = _parse_xml(p)
            if apply_numbering(dom, template["numbering"]):
                _write_bytes(dom, p)
                return f"{len(template['numbering'])} numbering definitions"
            return None

        def _task_theme():
            p = Path(word_dir_str) / "theme" / "theme1.xml"
            if not p.exists() or "theme" not in template:
                return None
            dom = _parse_xml(p)
            if apply_theme(dom, template["theme"]):
                _write_bytes(dom, p)
                return "theme colors/fonts"
            return None

        def _task_fonts():
            p = Path(word_dir_str) / "fontTable.xml"
            if not p.exists() or "fonts" not in template:
                return None
            dom = _parse_xml(p)
            if apply_font_table(dom, template["fonts"]):
                _write_bytes(dom, p)
                return f"{len(template['fonts'])} fonts added"
            return None

        tasks = {
            "styles": _task_styles,
            "page": _task_page,
            "numbering": _task_numbering,
            "theme": _task_theme,
            "fonts": _task_fonts,
        }

        with concurrent.futures.ThreadPoolExecutor(max_workers=len(tasks)) as xml_ex:
            fmap = {xml_ex.submit(fn): nm for nm, fn in tasks.items()}
            xml_results = []
            for future in concurrent.futures.as_completed(fmap):
                nm = fmap[future]
                try:
                    r = future.result()
                    xml_results.append(f"    {nm}: {r}" if r else f"    {nm}: -")
                except Exception as e:
                    xml_results.append(f"    {nm}: ERROR - {e}")

        for msg in sorted(xml_results):
            print(msg)

        # --- Level 2: document body chunking (per-section parallelism) ---
        if chunks > 1:
            n_edited = _apply_body_chunks_parallel(
                word_dir / "document.xml", chunks
            )
            if n_edited:
                print(f"    body: edited {n_edited} paragraphs across {chunks} chunks")

        # Repack
        _pack(str(unpack_dir), str(output_path), original=str(target_path))

    return str(output_path)


def _apply_body_chunks_parallel(doc_xml: Path, chunks: int) -> int:
    """Split document body paragraphs into *chunks* and process in parallel.

    Each chunk re-serialises the paragraph XML, modifies paragraph-level
    properties in a thread, and the results are merged back.

    Returns the number of modified paragraph elements.
    """
    dom = _parse_xml(doc_xml)
    body = dom.getElementsByTagName("w:body")
    if not body:
        return 0
    body = body[0]

    paragraphs = [n for n in body.childNodes if n.nodeType == n.ELEMENT_NODE]
    if not paragraphs:
        return 0

    # Split into chunks
    chunk_size = max(1, len(paragraphs) // chunks)
    groups = [
        paragraphs[i: i + chunk_size]
        for i in range(0, len(paragraphs), chunk_size)
    ]

    # Serialize each chunk to XML string for thread-safe processing
    chunk_data = []
    for group in groups:
        fragment = "".join(
            n.toxml() for n in group
        )
        chunk_data.append(fragment)

    def _process_chunk(xml_fragment: str) -> str | None:
        """Parse a paragraph fragment and return modified XML or None."""
        # Wrap in a minimal body so minidom can parse
        wrapped = f'<w:p xmlns:w="{NS["w"]}">{xml_fragment.replace("<w:p ", "<w:p ").replace("</w:p>", "</w:p>")}</w:p>'
        # Actually, the fragment is already a set of <w:p>... elements.
        # We need a proper wrapper
        inner = f'<w:body xmlns:w="{NS["w"]}">{xml_fragment}</w:body>'
        try:
            frag_dom = defusedxml.minidom.parseString(inner)
            # Nothing substantive to modify at paragraph level for now —
            # styles are applied globally via styles.xml.
            # This is an extension point for future per-paragraph operations.
            return None
        except Exception:
            return None

    with concurrent.futures.ThreadPoolExecutor(max_workers=chunks) as chunk_ex:
        futures = [chunk_ex.submit(_process_chunk, cd) for cd in chunk_data]
        for f in concurrent.futures.as_completed(futures):
            f.result()  # collect results / exceptions

    return 0  # no paragraph-level modifications yet (extension point)


# --- Batch entry point ---


def _apply_one(args: tuple) -> str:
    """Wrapper for *one* document in a batch run."""
    target_path, template, output_dir, chunks, category, version, paper, preserve_formatting, apply_body_formats = args
    if output_dir:
        out = Path(output_dir) / f"{target_path.stem}-formatted.docx"
    elif category:
        ver = version if version is not None else _detect_next_version(category, paper)
        if paper:
            out_dir = _PROJECT_ROOT / "outputs" / paper / category / f"v{ver}"
        else:
            out_dir = _PROJECT_ROOT / "outputs" / category / f"v{ver}"
        out_dir.mkdir(parents=True, exist_ok=True)
        out = out_dir / f"{target_path.stem}-formatted.docx"
    else:
        out = None  # auto-name beside source
    print(f"\n── {target_path.name} ──")
    try:
        result = _process_one_document(
            target_path, template, out, chunks=chunks,
            preserve_formatting=preserve_formatting,
            apply_body_formats=apply_body_formats,
        )
        return f"OK  {target_path.name} → {result}"
    except Exception as e:
        return f"FAIL {target_path.name}: {e}"


# --- Public entry point ---


def apply_template(
    target_file: str,
    template_file: str | None = None,
    output_file: str | None = None,
    *,
    batch: bool = False,
    chunks: int = 1,
    jobs: int | None = None,
    output_dir: str | None = None,
    category: str | None = None,
    version: str | None = None,
    paper: str | None = None,
    preserve_formatting: bool = False,
    apply_body_formats: bool = False,
):
    """Apply a formatting template to one or more DOCX files."""
    _ensure_docx_skill()

    # Resolve template from explicit path or category lookup
    template_file = _resolve_template(template_file, category)
    template_path = Path(template_file)
    if not template_path.exists():
        print(f"Error: template not found: {template_file}", file=sys.stderr)
        sys.exit(1)

    with open(template_path, "r", encoding="utf-8") as f:
        template = yaml.safe_load(f)
    if not template:
        print("Error: Empty template", file=sys.stderr)
        sys.exit(1)

    # Resolve target files
    target_paths = _resolve_targets(target_file, batch)

    if len(target_paths) == 1 and not batch:
        # Single document — simple call
        target_path = target_paths[0]
        out = None
        if output_file:
            out = Path(output_file)
        elif category:
            ver = version if version is not None else _detect_next_version(category, paper)
            if paper:
                out_dir = _PROJECT_ROOT / "outputs" / paper / category / f"v{ver}"
            else:
                out_dir = _PROJECT_ROOT / "outputs" / category / f"v{ver}"
            out_dir.mkdir(parents=True, exist_ok=True)
            out = out_dir / f"{target_path.stem}-formatted.docx"
        _process_one_document(
            target_path, template, out, chunks=chunks,
            preserve_formatting=preserve_formatting,
            apply_body_formats=apply_body_formats,
        )
        return

    # Batch mode: multiple documents in parallel
    if output_dir:
        Path(output_dir).mkdir(parents=True, exist_ok=True)

    n_workers = jobs or min(len(target_paths), os.cpu_count() or 4)
    work_items = [
        (tp, template, output_dir, chunks, category, version, paper, preserve_formatting, apply_body_formats)
        for tp in target_paths
    ]

    print(f"\nBatch processing {len(work_items)} document(s) ({n_workers} workers)...")
    with concurrent.futures.ThreadPoolExecutor(max_workers=n_workers) as batch_ex:
        futures = [batch_ex.submit(_apply_one, wi) for wi in work_items]
        results = []
        for future in concurrent.futures.as_completed(futures):
            results.append(future.result())

    print("\n" + "=" * 50)
    print("Batch results:")
    for r in sorted(results):
        print(f"  {r}")


def _resolve_targets(target_file: str, batch: bool) -> list[Path]:
    """Resolve target_file into a list of existing file paths.

    In batch mode, *target_file* may be a glob pattern like ``chapter*.docx``.
    """
    candidate = Path(target_file)
    if candidate.exists():
        return [candidate.resolve()]

    if not batch:
        print(f"Error: target not found: {target_file}", file=sys.stderr)
        sys.exit(1)

    # Try as glob
    matches = sorted(Path(p).resolve() for p in glob_module.glob(target_file))
    if not matches:
        print(
            f"Error: no files match pattern: {target_file}",
            file=sys.stderr,
        )
        sys.exit(1)
    return matches


def main():
    parser = argparse.ArgumentParser(
        description="Apply a formatting template to DOCX document(s)",
    )
    parser.add_argument(
        "target_file",
        help=(
            "Target DOCX file. In --batch mode this can be a glob pattern "
            "(e.g. 'chapter*.docx' or '/path/to/*.docx')"
        ),
    )
    parser.add_argument(
        "template_file",
        nargs="?",
        help="Template YAML file (optional if --category is given; "
             "auto-looked up in styles/<category>/)",
    )
    parser.add_argument(
        "--output", "-o",
        help="Output DOCX file (single-document mode only)",
    )
    parser.add_argument(
        "--batch", "-b",
        action="store_true",
        help="Batch mode: process multiple files (target_file may be a glob)",
    )
    parser.add_argument(
        "--chunks",
        type=int,
        default=1,
        help="Split document body into N chunks for per-section parallelism (default: 1 = off)",
    )
    parser.add_argument(
        "--jobs", "-j",
        type=int,
        default=None,
        help="Number of parallel workers for batch mode (default: min(files, CPU count))",
    )
    parser.add_argument(
        "--output-dir",
        help="Output directory for batch mode (default: same folder as each source)",
    )
    parser.add_argument(
        "--category", "-c",
        help="Template name (e.g., 北交模板). "
             "Output goes to outputs/<paper>/<name>/v<version>/; "
             "template auto-located in styles/<name>/ if not specified",
    )
    parser.add_argument(
        "--version", "-v",
        type=str,
        default=None,
        help="Version string in the format major.minor (e.g. 1.1). "
             "Default: auto-detect next version in outputs/<paper>/<name>/)",
    )
    parser.add_argument(
        "--paper", "-p",
        help="Paper or article name (e.g., 道路交通场景双级样条网络). "
             "Output goes to outputs/<paper>/<category>/v<version>/; "
             "when omitted uses <category> as root: outputs/<category>/v<version>/",
    )
    parser.add_argument(
        "--preserve-formatting",
        action="store_true",
        help="Skip body content formatting cleanup (default: strip direct formatting unless template has body_paragraph_formats)",
    )
    parser.add_argument(
        "--apply-body-formats",
        action="store_true",
        help=(
            "Apply body_paragraph_formats by paragraph index. Use only when "
            "the target has the same paragraph structure as the reference."
        ),
    )

    args = parser.parse_args()

    apply_template(
        args.target_file,
        args.template_file,
        args.output,
        batch=args.batch,
        chunks=args.chunks,
        jobs=args.jobs,
        output_dir=args.output_dir,
        category=args.category,
        version=args.version,
        paper=args.paper,
        preserve_formatting=args.preserve_formatting,
        apply_body_formats=args.apply_body_formats,
    )


if __name__ == "__main__":
    main()
