#!/usr/bin/env python3
"""Generate a DOCX from a structured YAML/JSON specification.

Uses docx-js (Node.js npm package) to create the document.
Generates a temporary JS script and executes it via subprocess.

Usage:
    python scripts/write_docx.py paper.yaml --output paper.docx
    python scripts/write_docx.py paper.yaml -c 北交模板 -p 论文名 -v 1
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import defusedxml.minidom
import yaml

# A4 page size in DXA
A4_WIDTH = 11906
A4_HEIGHT = 16838
DEFAULT_MARGIN = 1440  # 1 inch in DXA
DEFAULT_FONT = "Times New Roman"
DEFAULT_FONT_SIZE = 24  # half-points, 24 = 12pt

MATH_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"

_math_placeholders: dict[int, dict] = {}
_INLINE_MATH_RE = re.compile(r"\$(.+?)\$")
_math_next_idx: int = 0

_SEQ_MARKER_RE = re.compile(r"\{\{seq:(\w+)\}\}")
_REF_MARKER_RE = re.compile(r"\{\{ref:([\w-]+)\}\}")
_PAGEREF_MARKER_RE = re.compile(r"\{\{pageref:([\w-]+)\}\}")

_seq_counters: dict[str, int] = {}
_auto_bookmark_ids: dict[str, str] = {}


def _reset_counters():
    """Reset SEQ counters and auto-bookmark tracking for a new generation run."""
    global _seq_counters, _auto_bookmark_ids
    _seq_counters = {}
    _auto_bookmark_ids = {}


def _bmid(user_id: str) -> str:
    """Convert user-facing bookmark ID to internal safe ID with prefix."""
    return f"_bm_{user_id}"


def load_spec(path: str) -> dict:
    """Load a YAML or JSON specification file."""
    with open(path, "r", encoding="utf-8") as f:
        if path.endswith((".yaml", ".yml")):
            return yaml.safe_load(f)
        return json.load(f)


def load_template(path: str) -> dict:
    """Load a style template YAML file."""
    return load_spec(path)


def _apply_template_defaults(spec: dict, template: dict) -> dict:
    """Merge template defaults into spec. Spec values take priority."""
    result = dict(spec)

    # Page layout: template defaults, spec overrides
    if "page" not in result:
        result["page"] = template.get("page", {})

    # Style: template document_defaults, spec overrides
    if "style" not in result:
        defaults = template.get("document_defaults", {})
        run = defaults.get("run", {})
        style = {}
        if "font" in run:
            style["default_font"] = run["font"]
        if "fontSize" in run:
            style["default_font_size"] = run["fontSize"]
        result["style"] = style

    # TOC: template defaults, spec overrides
    if "toc" not in result:
        tmpl_toc = template.get("toc", {})
        if tmpl_toc:
            result["toc"] = tmpl_toc

    # Features: template defaults, spec overrides
    if "features" not in result:
        result["features"] = template.get("features", {})

    # Caption templates stored for use during generation
    result["_caption_templates"] = template.get("caption_templates", {})

    return result


def _caption_prefix_suffix(template: dict, lang: str, item_type: str) -> tuple[str, str, str]:
    """Get (prefix, suffix, seq_name) for a caption from template config.

    lang: "cn" or "en"
    item_type: "figure" or "table"
    """
    ct = template.get("_caption_templates", {})
    key = f"{item_type}_{lang}"
    cfg = ct.get(key, {})
    prefix = cfg.get("prefix", "")
    suffix = cfg.get("suffix", "")
    seq_name = cfg.get("seq_name", item_type.capitalize())
    return prefix, suffix, seq_name


def _build_preview_spec(template: dict, output_dir: str) -> dict:
    """Build an in-memory spec that demonstrates all style features."""
    ct = template.get("caption_templates", template.get("_caption_templates", {}))
    features = template.get("features", {})

    fig_cn = ct.get("figure_cn", {})
    tab_cn = ct.get("table_cn", {})
    fig_prefix = fig_cn.get("prefix", "图 ")
    fig_suffix = fig_cn.get("suffix", "")
    tab_prefix = tab_cn.get("prefix", "表 ")
    tab_suffix = tab_cn.get("suffix", "")

    # Create a small placeholder image for preview
    import io
    preview_png = str(Path(output_dir) / "_preview_placeholder.png")
    try:
        from PIL import Image
    except ImportError:
        Image = None
    if Image:
        img = Image.new("RGB", (400, 300), color="#4472C4")
        img.save(preview_png)

    has_toc = bool(template.get("toc")) and features.get("clickable_toc", True)
    has_cross_refs = features.get("clickable_cross_refs", True)
    has_auto_num = features.get("auto_numbering", True)

    fig_caption = f"{fig_prefix}{{{{seq:Figure}}}}{fig_suffix} 示例图片 — 格式预览"
    tab_caption = f"{tab_prefix}{{{{seq:Table}}}}{tab_suffix} 示例表格 — 格式预览"

    sections = [
        {
            "id": "sec-sample",
            "heading": "第一节  示例章节",
            "level": 1,
            "content": [
                {
                    "type": "paragraph",
                    "text": (
                        "这是一个示例正文段落，用于验证正文字体、字号、行距是否正确。"
                        "本段包含行内数学公式 $E=mc^2$ 以验证公式渲染。"
                    ),
                },
                {
                    "type": "paragraph",
                    "text": (
                        "下面是一张示例图片，其题注使用了模板定义的格式："
                    ) if has_cross_refs else "下面是一张示例图片。",
                },
                {
                    "type": "image",
                    "src": preview_png if Image else "",
                    "width": 350,
                    "bookmark_id": "fig-sample" if has_cross_refs else None,
                    "caption": fig_caption if has_auto_num else "图 1  示例图片 — 格式预览",
                },
                {
                    "type": "paragraph",
                    "text": (
                        f"如{'{{{{ref:fig-sample}}}}' if has_cross_refs else '图1'}所示，"
                        "示例图片展示了模板的图表题注格式。"
                        f"{'（见第{{{{pageref:fig-sample}}}}}页）' if has_cross_refs else ''}"
                    ),
                },
                {
                    "type": "paragraph",
                    "text": "下面是一张示例表格：",
                },
                {
                    "type": "table",
                    "bookmark_id": "tab-sample" if has_cross_refs else None,
                    "caption": tab_caption if has_auto_num else "表 1  示例表格 — 格式预览",
                    "columns": ["指标", "数值", "说明"],
                    "rows": [
                        ["准确率", "95.2%", "测试集"],
                        ["召回率", "93.1%", "测试集"],
                        ["F1 分数", "94.1%", "宏平均"],
                    ],
                },
                {
                    "type": "paragraph",
                    "text": (
                        f"如{'{{{{ref:tab-sample}}}}' if has_cross_refs else '表1'}所示，"
                        "表格格式符合模板要求。"
                    ),
                },
            ],
        },
        {
            "id": "sec-lists",
            "heading": "第二节  列表与子标题",
            "level": 1,
            "content": [
                {
                    "type": "paragraph",
                    "text": "以下为二级标题和列表的格式预览。",
                },
            ],
        },
        {
            "heading": "二级标题示例",
            "level": 2,
            "content": [
                {
                    "type": "paragraph",
                    "text": "二级标题下的正文内容。",
                },
                {
                    "type": "list",
                    "ordered": True,
                    "items": [
                        "有序列表第一项，验证列表编号格式",
                        "有序列表第二项，验证多行文本的缩进",
                        "有序列表第三项",
                    ],
                },
                {
                    "type": "list",
                    "ordered": False,
                    "items": [
                        "无序列表第一项",
                        "无序列表第二项",
                        "无序列表第三项",
                    ],
                },
            ],
        },
    ]

    return {
        "meta": {"paper_name": "style-preview"},
        **({"toc": template["toc"]} if has_toc else {}),
        "title": {
            "cn": "格式预览文档 — Style Preview",
            "en": "Style Preview Document — 格式预览",
        },
        "authors": [
            {"name": "预览作者", "affiliation": "北京交通大学 计算机与信息技术学院"},
        ],
        "abstract": {
            "cn": "这是一个格式预览文档，用于验证从参考模板中提取的样式是否正确。包括标题、摘要、关键词、目录、正文、图表题注、列表、参考文献等所有元素的格式。",
            "en": "This is a style preview document to verify that extracted template formatting is correct. It includes all document elements: titles, abstracts, keywords, TOC, body text, figure/table captions, lists, and references.",
        },
        "keywords": {
            "cn": ["格式预览", "模板验证", "样式检查"],
            "en": ["style preview", "template verification", "format check"],
        },
        "sections": sections,
        "references": [
            {"text": "[1] 作者甲, 作者乙. 示例论文标题[J]. 示例期刊, 2025, 40(1): 1-10."},
            {"text": "[2] Smith J, Doe J. Sample Paper Title[C]. Sample Conference, 2025."},
            {"text": "[3] 作者丙. 深度学习在自然语言处理中的应用[M]. 北京: 示例出版社, 2024."},
        ],
    }

    """Escape a string for safe embedding in JavaScript template literal."""
    return json.dumps(text, ensure_ascii=False)


def _escape_js_string(text: str) -> str:
    """Escape a string for safe embedding in JavaScript template literal."""
    return json.dumps(text, ensure_ascii=False)


def _validate_docx(path: str) -> bool:
    """Basic validation: check the file is a well-formed ZIP (DOCX is a ZIP)."""
    try:
        with zipfile.ZipFile(path, "r") as zf:
            bad = zf.testzip()
            if bad:
                print(f"  Warning: corrupt entry in DOCX: {bad}", file=sys.stderr)
                return False
            names = zf.namelist()
            if "[Content_Types].xml" not in names:
                print("  Warning: missing [Content_Types].xml", file=sys.stderr)
                return False
        return True
    except (zipfile.BadZipFile, OSError) as e:
        print(f"  Validation failed: {e}", file=sys.stderr)
        return False


def generate_js(spec: dict, output_path: str, spec_dir: str) -> str:
    """Generate a complete Node.js script using docx-js to create the DOCX.

    Dispatches to gen_patent.generate_patent_js() when the spec type is 'patent'.
    """
    if spec.get("type") == "patent":
        from gen_patent import generate_patent_js
        return generate_patent_js(spec, output_path, spec_dir)

    global _math_placeholders, _math_next_idx
    _math_placeholders = {}
    _math_next_idx = 0
    _reset_counters()

    page = spec.get("page", {})
    page_width = page.get("width", A4_WIDTH)
    page_height = page.get("height", A4_HEIGHT)
    margins = page.get("margins", {
        "top": DEFAULT_MARGIN,
        "bottom": DEFAULT_MARGIN,
        "left": DEFAULT_MARGIN,
        "right": DEFAULT_MARGIN,
    })
    content_width = page_width - margins.get("left", 0) - margins.get("right", 0)

    style = spec.get("style", {})
    font = style.get("default_font", DEFAULT_FONT)
    font_size = style.get("default_font_size", DEFAULT_FONT_SIZE)
    font_color = style.get("default_font_color", "000000")

    lines = []
    lines.append("// Auto-generated by docx-write/write_docx.py")
    lines.append("const fs = require('fs');")
    lines.append("const path = require('path');")
    lines.append("const {")
    lines.append("  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,")
    lines.append("  Header, Footer, AlignmentType, HeadingLevel, LevelFormat,")
    lines.append(
        "  BorderStyle, WidthType, ShadingType, PageBreak, PageNumber,"
    )
    lines.append(
        "  ImageRun, TabStopType, TabStopPosition,"
    )
    lines.append(
        "  Bookmark, InternalHyperlink, TableOfContents,"
    )
    lines.append(
        "  SequentialIdentifier, SimpleField, PageReference"
    )
    lines.append("} = require('docx');")
    lines.append("")
    lines.append("const doc = new Document({")

    # Styles
    lines.append("  styles: {")
    lines.append("    default: {")
    lines.append("      document: {")
    lines.append(f'        run: {{ font: {_escape_js_string(font)}, size: {font_size}, color: "{font_color}" }},')
    lines.append("      },")
    lines.append("    },")
    lines.append("    paragraphStyles: [")
    heading_styles = [
        ("Heading1", "Heading 1", 32, 240),
        ("Heading2", "Heading 2", 28, 180),
        ("Heading3", "Heading 3", 26, 120),
        ("Heading4", "Heading 4", 24, 120),
    ]
    for idx, (h_id, h_name, h_size, h_space) in enumerate(heading_styles):
        lines.append("      {")
        lines.append(f'        id: "{h_id}",')
        lines.append(f'        name: "{h_name}",')
        lines.append('        basedOn: "Normal",')
        lines.append('        next: "Normal",')
        lines.append("        quickFormat: true,")
        lines.append(f"        run: {{ size: {h_size}, bold: true, font: {_escape_js_string(font)} }},")
        lines.append(
            f"        paragraph: {{ spacing: {{ before: {h_space}, after: {h_space // 2} }}, outlineLevel: {idx} }},"
        )
        lines.append("      },")
    lines.append("    ],")
    lines.append("  },")

    # Numbering config (for lists)
    lines.append("  numbering: {")
    lines.append("    config: [")
    lines.append("      {")
    lines.append('        reference: "bullets",')
    lines.append("        levels: [{")
    lines.append("          level: 0,")
    lines.append('          format: LevelFormat.BULLET,')
    lines.append('          text: "•",')
    lines.append("          alignment: AlignmentType.LEFT,")
    lines.append(
        "          style: { paragraph: { indent: { left: 720, hanging: 360 } } },"
    )
    lines.append("        }],")
    lines.append("      },")
    lines.append("      {")
    lines.append('        reference: "numbers",')
    lines.append("        levels: [{")
    lines.append("          level: 0,")
    lines.append('          format: LevelFormat.DECIMAL,')
    lines.append('          text: "%1.",')
    lines.append("          alignment: AlignmentType.LEFT,")
    lines.append(
        "          style: { paragraph: { indent: { left: 720, hanging: 360 } } },"
    )
    lines.append("        }],")
    lines.append("      },")
    lines.append("    ],")
    lines.append("  },")

    # Section content
    lines.append("  sections: [{")
    lines.append("    properties: {")
    lines.append("      page: {")
    lines.append(f"        size: {{ width: {page_width}, height: {page_height} }},")
    lines.append(
        f"        margin: {{ top: {margins['top']}, bottom: {margins['bottom']}, "
        f"left: {margins['left']}, right: {margins['right']} }},"
    )
    lines.append("      },")
    lines.append("    },")

    # Header
    lines.append("    headers: {")
    lines.append("      default: new Header({")
    lines.append("        children: [new Paragraph({")
    title_en = spec.get("title", {})
    title_text = title_en.get("en", "") or title_en.get("cn", "") or ""
    lines.append(
        f"          alignment: AlignmentType.CENTER,"
    )
    lines.append(
        f"          children: [new TextRun({{ text: {_escape_js_string(title_text)}, "
        f"font: {_escape_js_string(font)}, size: 18, color: \"888888\" }})],"
    )
    lines.append("        })],")
    lines.append("      }),")
    lines.append("    },")

    # Footer with page numbers
    lines.append("    footers: {")
    lines.append("      default: new Footer({")
    lines.append("        children: [new Paragraph({")
    lines.append("          alignment: AlignmentType.CENTER,")
    lines.append("          children: [")
    lines.append(
        "            new TextRun({ text: \"— \", font: "
        + _escape_js_string(font)
        + ", size: 18, color: \"888888\" }),"
    )
    lines.append("            new TextRun({ children: [PageNumber.CURRENT],")
    lines.append(
        f"              font: {_escape_js_string(font)}, size: 18, color: \"888888\" }}),"
    )
    lines.append(
        "            new TextRun({ text: \" —\", font: "
        + _escape_js_string(font)
        + ", size: 18, color: \"888888\" }),"
    )
    lines.append("          ],")
    lines.append("        })],")
    lines.append("      }),")
    lines.append("    },")

    # Body children
    lines.append("    children: [")
    lines.append("      // ===== TITLE =====")
    lines += _gen_title(spec, font)
    lines.append("      // ===== AUTHORS =====")
    lines += _gen_authors(spec, font)
    lines.append("      // ===== ABSTRACT =====")
    lines += _gen_abstract(spec, font)
    lines.append("      // ===== KEYWORDS =====")
    lines += _gen_keywords(spec, font)
    lines.append("      // ===== TOC =====")
    lines += _gen_toc(spec, font)
    lines.append("      // ===== BODY SECTIONS =====")
    lines += _gen_sections(spec, font, content_width, spec_dir)
    lines.append("      // ===== REFERENCES =====")
    lines += _gen_references(spec, font)

    lines.append("    ],")
    lines.append("  }],")
    lines.append("});")
    lines.append("")
    lines.append(
        f"Packer.toBuffer(doc).then(buf => fs.writeFileSync({_escape_js_string(output_path)}, buf));"
    )

    return "\n".join(lines)


def _gen_caption_children(
    caption: str,
    font: str,
    font_size: int = 20,
    bookmark_id: str | None = None,
    color: str | None = None,
    bold: bool = False,
) -> list[str]:
    """Generate paragraph children JS for a caption with optional {{seq:NAME}} markers.

    If a bookmark_id is provided, the first {{seq:NAME}} marker's SequentialIdentifier
    is wrapped in a Bookmark for cross-reference targeting.
    """
    parts = _SEQ_MARKER_RE.split(caption)
    has_seq = len(parts) > 1

    if not has_seq:
        extra = ""
        if color:
            extra += f', color: "{color}"'
        if bold:
            extra += ", bold: true"
        return [
            f"          new TextRun({{ text: {_escape_js_string(caption)}, "
            f"font: {_escape_js_string(font)}, size: {font_size}{extra} }}),"
        ]

    children = []
    bookmark_used = False
    for i, part in enumerate(parts):
        if not part:
            continue
        if i % 2 == 0:
            extra = ""
            if color:
                extra += f', color: "{color}"'
            if bold:
                extra += ", bold: true"
            children.append(
                f"          new TextRun({{ text: {_escape_js_string(part)}, "
                f"font: {_escape_js_string(font)}, size: {font_size}{extra} }}),"
            )
        else:
            seq = f"new SequentialIdentifier({_escape_js_string(part)})"
            if bookmark_id and not bookmark_used:
                safe_id = _bmid(bookmark_id)
                children.append(
                    f"          new Bookmark({{ id: {_escape_js_string(safe_id)}, "
                    f"children: [{seq}] }}),"
                )
                bookmark_used = True
            else:
                children.append(f"          {seq},")
    return children


def _gen_paragraph_with_fields(
    text: str,
    font: str,
    font_size: int,
    spacing_after: int,
    next_math_idx: int,
) -> tuple[list[str], int]:
    """Generate JS for a paragraph with {{ref:ID}}, {{pageref:ID}}, and {{seq:NAME}} markers.

    Also handles $...$ inline math via delegation to _gen_paragraph_with_math().
    Returns (lines, next_idx).
    """
    has_ref = _REF_MARKER_RE.search(text)
    has_pageref = _PAGEREF_MARKER_RE.search(text)
    has_seq = _SEQ_MARKER_RE.search(text)
    has_math = _INLINE_MATH_RE.search(text)

    if not has_ref and not has_pageref and not has_seq:
        return _gen_paragraph_with_math(text, font, font_size, spacing_after, next_math_idx)

    idx = next_math_idx
    # Build a single marker pattern that splits on all three
    combined_re = re.compile(r"(\{\{ref:[\w-]+\}\}|\{\{pageref:[\w-]+\}\}|\{\{seq:\w+\}\})")
    segments = combined_re.split(text)

    children: list[str] = []
    for seg in segments:
        if not seg:
            continue
        ref_m = _REF_MARKER_RE.fullmatch(seg)
        pageref_m = _PAGEREF_MARKER_RE.fullmatch(seg)
        seq_m = _SEQ_MARKER_RE.fullmatch(seg)

        if ref_m:
            safe_id = _bmid(ref_m.group(1))
            children.append(
                f"          new InternalHyperlink({{ anchor: {_escape_js_string(safe_id)}, "
                f"children: [new SimpleField({_escape_js_string(f' REF {safe_id} \\\\h ')})] }}),"
            )
        elif pageref_m:
            safe_id = _bmid(pageref_m.group(1))
            children.append(
                f"          new PageReference({_escape_js_string(safe_id)}),"
            )
        elif seq_m:
            children.append(
                f"          new SequentialIdentifier({_escape_js_string(seq_m.group(1))}),"
            )
        else:
            math_child_lines, idx = _inline_math_children(seg, font, font_size, idx)
            children.extend(math_child_lines)

    lines = []
    lines.append("      new Paragraph({")
    lines.append(f"        spacing: {{ after: {spacing_after} }},")
    lines.append("        children: [")
    lines.extend(children)
    lines.append("        ],")
    lines.append("      }),")
    return lines, idx


def _inline_math_children(
    text: str, font: str, font_size: int, start_idx: int
) -> tuple[list[str], int]:
    """Generate JS children for literal text that may contain $...$ inline math.

    Caller should use this for the non-marker segments within a paragraph.
    Returns (child_lines, next_idx).
    """
    parts = _INLINE_MATH_RE.split(text)
    has_math = len(parts) > 1
    idx = start_idx
    children = []

    for i, part in enumerate(parts):
        if not part:
            continue
        if i % 2 == 0:
            children.append(
                f"          new TextRun({{ text: {_escape_js_string(part)}, "
                f"font: {_escape_js_string(font)}, size: {font_size} }}),"
            )
        else:
            _math_placeholders[idx] = {"latex": part, "display": "inline"}
            marker = f"##MATH_PLACEHOLDER_{idx}##"
            children.append(
                f"          new TextRun({{ text: {_escape_js_string(marker)}, "
                f"font: {_escape_js_string(font)}, size: {font_size} }}),"
            )
            idx += 1
    return children, idx


def _gen_toc(spec: dict, font: str) -> list[str]:
    """Generate JS for a Table of Contents with heading paragraph."""
    toc_config = spec.get("toc", {})
    if not toc_config:
        return []

    label = toc_config.get("label", "Table of Contents")
    heading_range = toc_config.get("headingStyleRange", "1-3")

    lines = [
        "      new Paragraph({",
        f"        heading: HeadingLevel.HEADING_1,",
        f"        children: [new TextRun({{ text: {_escape_js_string(label)}, "
        f"font: {_escape_js_string(font)}, size: 32, bold: true }})],",
        "      }),",
        f"      new TableOfContents({_escape_js_string(label)}, {{",
        f"        headingStyleRange: {_escape_js_string(heading_range)},",
        "      }),",
        "      new Paragraph({ children: [new PageBreak()] }),",
    ]
    return lines


def _gen_title(spec: dict, font: str) -> list[str]:
    """Generate the title paragraph(s) in JS."""
    lines = []
    title = spec.get("title", {})
    cn_title = title.get("cn", "")
    en_title = title.get("en", "")

    lines.append("      // Title paragraph(s)")
    if cn_title:
        lines.append("      new Paragraph({")
        lines.append("        alignment: AlignmentType.CENTER,")
        lines.append("        spacing: { after: 120 },")
        lines.append(
            f"        children: [new TextRun({{ text: {_escape_js_string(cn_title)}, "
            f"font: {_escape_js_string(font)}, size: 36, bold: true }})],"
        )
        lines.append("      }),")
    if en_title:
        lines.append("      new Paragraph({")
        lines.append("        alignment: AlignmentType.CENTER,")
        en_size = 36 if not cn_title else 28
        lines.append("        spacing: { after: 200 },")
        lines.append(
            f"        children: [new TextRun({{ text: {_escape_js_string(en_title)}, "
            f"font: {_escape_js_string(font)}, size: {en_size}, bold: true }})],"
        )
        lines.append("      }),")

    return lines


def _gen_authors(spec: dict, font: str) -> list[str]:
    """Generate author paragraph(s)."""
    lines = []
    authors = spec.get("authors", [])
    if not authors:
        return lines

    name_parts = [a.get("name", "") for a in authors]
    name_line = ", ".join(name_parts)
    lines.append("      new Paragraph({")
    lines.append("        alignment: AlignmentType.CENTER,")
    lines.append("        spacing: { after: 120 },")
    lines.append(
        f"        children: [new TextRun({{ text: {_escape_js_string(name_line)}, "
        f"font: {_escape_js_string(font)}, size: 24 }})],"
    )
    lines.append("      }),")

    affiliations = {}
    for author in authors:
        aff = author.get("affiliation", "")
        if aff and aff not in affiliations:
            affiliations[aff] = len(affiliations) + 1

    for aff, num in affiliations.items():
        lines.append("      new Paragraph({")
        lines.append("        alignment: AlignmentType.CENTER,")
        lines.append("        spacing: { after: 80 },")
        label = f"{num}. {aff}"
        lines.append(
            f"        children: [new TextRun({{ text: {_escape_js_string(label)}, "
            f"font: {_escape_js_string(font)}, size: 20, color: \"555555\" }})],"
        )
        lines.append("      }),")

    return lines


def _gen_abstract(spec: dict, font: str) -> list[str]:
    """Generate abstract section."""
    lines = []
    abstract = spec.get("abstract", {})
    if not abstract:
        return lines

    cn_abstract = abstract.get("cn", "")
    en_abstract = abstract.get("en", "")

    if cn_abstract:
        lines.append("      new Paragraph({")
        lines.append("        heading: HeadingLevel.HEADING_1,")
        lines.append("        children: [new TextRun({ text: \"摘要\",")
        lines.append(
            f"          font: {_escape_js_string(font)}, size: 32, bold: true }})],"
        )
        lines.append("      }),")
        lines.append("      new Paragraph({")
        lines.append("        spacing: { after: 200 },")
        lines.append(
            f"        children: [new TextRun({{ text: {_escape_js_string(cn_abstract)}, "
            f"font: {_escape_js_string(font)}, size: 24 }})],"
        )
        lines.append("      }),")

    if en_abstract:
        lines.append("      new Paragraph({")
        lines.append("        heading: HeadingLevel.HEADING_1,")
        lines.append("        children: [new TextRun({ text: \"Abstract\",")
        lines.append(
            f"          font: {_escape_js_string(font)}, size: 32, bold: true }})],"
        )
        lines.append("      }),")
        lines.append("      new Paragraph({")
        lines.append("        spacing: { after: 200 },")
        lines.append(
            f"        children: [new TextRun({{ text: {_escape_js_string(en_abstract)}, "
            f"font: {_escape_js_string(font)}, size: 24 }})],"
        )
        lines.append("      }),")

    return lines


def _gen_keywords(spec: dict, font: str) -> list[str]:
    """Generate keywords section."""
    lines = []
    keywords = spec.get("keywords", {})
    if not keywords:
        return lines

    cn_kw = keywords.get("cn", [])
    en_kw = keywords.get("en", [])

    if cn_kw:
        kw_text = "关键词：" + "；".join(cn_kw)
        lines.append("      new Paragraph({")
        lines.append("        spacing: { after: 200 },")
        lines.append("        children: [")
        lines.append(
            f"          new TextRun({{ text: {_escape_js_string(kw_text)}, "
            f"font: {_escape_js_string(font)}, size: 24, bold: false }}),"
        )
        lines.append("        ],")
        lines.append("      }),")

    if en_kw:
        kw_text = "Keywords: " + "; ".join(en_kw)
        lines.append("      new Paragraph({")
        lines.append("        spacing: { after: 200 },")
        lines.append("        children: [")
        lines.append(
            f"          new TextRun({{ text: {_escape_js_string(kw_text)}, "
            f"font: {_escape_js_string(font)}, size: 24, bold: false }}),"
        )
        lines.append("        ],")
        lines.append("      }),")

    return lines


def _gen_math_item(
    idx: int,
    latex: str,
    display: str,
    label: str | None = None,
    alignment: str = "CENTER",
    spacing_before: int = 120,
    spacing_after: int = 120,
    **_kwargs,
) -> list[str]:
    """Generate JS code for a type: math content item.

    Emits a placeholder marker text run (for post-processing to replace
    with real OMML). Block math gets its own paragraph; inline math
    emits just the text run for embedding inside another paragraph.
    """
    resolved_latex = latex
    if label:
        resolved_latex = f"{latex} \\qquad {label}"
    _math_placeholders[idx] = {"latex": resolved_latex, "display": display}

    marker = f"##MATH_PLACEHOLDER_{idx}##"

    if display == "inline":
        return [
            f"          new TextRun({{ text: {_escape_js_string(marker)} }}),"
        ]

    lines = []
    lines.append("      new Paragraph({")
    lines.append(f"        spacing: {{ before: {spacing_before}, after: {spacing_after} }},")
    if alignment:
        lines.append(f"        alignment: AlignmentType.{alignment.upper()},")
    lines.append(
        f"        children: [new TextRun({{ text: {_escape_js_string(marker)} }})],"
    )
    lines.append("      }),")
    return lines


def _gen_paragraph_with_math(
    text: str,
    font: str,
    font_size: int,
    spacing_after: int,
    next_math_idx: int,
) -> tuple[list[str], int]:
    """Generate JS for a paragraph that may contain $...$ inline math.

    Returns (lines, next_idx) with next_idx advanced past any math segments.
    """
    parts = _INLINE_MATH_RE.split(text)
    has_math = len(parts) > 1

    if not has_math:
        lines = []
        lines.append("      new Paragraph({")
        lines.append(f"        spacing: {{ after: {spacing_after} }},")
        lines.append(
            f"        children: [new TextRun({{ text: {_escape_js_string(text)}, "
            f"font: {_escape_js_string(font)}, size: {font_size} }})],"
        )
        lines.append("      }),")
        return lines, next_math_idx

    idx = next_math_idx
    children = []
    for i, part in enumerate(parts):
        if i % 2 == 0:
            if part:
                children.append(
                    f"          new TextRun({{ text: {_escape_js_string(part)}, "
                    f"font: {_escape_js_string(font)}, size: {font_size} }}),"
                )
        else:
            _math_placeholders[idx] = {"latex": part, "display": "inline"}
            marker = f"##MATH_PLACEHOLDER_{idx}##"
            children.append(
                f"          new TextRun({{ text: {_escape_js_string(marker)}, "
                f"font: {_escape_js_string(font)}, size: {font_size} }}),"
            )
            idx += 1

    lines = []
    lines.append("      new Paragraph({")
    lines.append(f"        spacing: {{ after: {spacing_after} }},")
    lines.append("        children: [")
    lines.extend(children)
    lines.append("        ],")
    lines.append("      }),")
    return lines, idx


def _gen_sections(spec: dict, font: str, content_width: int, spec_dir: str) -> list[str]:
    """Generate body sections from spec."""
    global _math_next_idx
    lines = []
    sections = spec.get("sections", [])

    heading_map = {
        1: "HeadingLevel.HEADING_1",
        2: "HeadingLevel.HEADING_2",
        3: "HeadingLevel.HEADING_3",
        4: "HeadingLevel.HEADING_4",
    }

    for sec_idx, section in enumerate(sections):
        heading = section.get("heading", "")
        level = section.get("level", 1)
        heading_level = heading_map.get(level, "HeadingLevel.HEADING_1")
        content_items = section.get("content", [])

        if heading:
            if level == 1 and sec_idx > 0:
                lines.append(
                    "      new Paragraph({ children: [new PageBreak()] }),"
                )

            sec_id = section.get("id", "")
            lines.append("      new Paragraph({")
            lines.append(f"        heading: {heading_level},")
            if sec_id:
                safe_id = _bmid(sec_id)
                lines.append(
                    f"        children: [new Bookmark({{ id: {_escape_js_string(safe_id)}, "
                    f"children: [new TextRun({{ text: {_escape_js_string(heading)}, "
                    f"font: {_escape_js_string(font)}, bold: true }})] }})],"
                )
            else:
                lines.append(
                    f"        children: [new TextRun({{ text: {_escape_js_string(heading)}, "
                    f"font: {_escape_js_string(font)}, bold: true }})],"
                )
            lines.append("      }),")

        for item in content_items:
            item_type = item.get("type", "paragraph")

            if item_type == "paragraph":
                text = item.get("text", "")
                para_lines, _math_next_idx = _gen_paragraph_with_fields(
                    text, font, 24, 120, _math_next_idx
                )
                lines.extend(para_lines)

            elif item_type == "math":
                latex = item.get("latex", "")
                display = item.get("display", "block")
                label = item.get("label")
                alignment = item.get("alignment", "CENTER" if display == "block" else "")
                spacing_before = item.get("spacing_before", 120)
                spacing_after = item.get("spacing_after", 120)
                math_lines = _gen_math_item(
                    _math_next_idx, latex, display, label,
                    alignment=alignment,
                    spacing_before=spacing_before,
                    spacing_after=spacing_after,
                )
                lines.extend(math_lines)
                _math_next_idx += 1

            elif item_type == "image":
                src = item.get("src", "")
                caption = item.get("caption", "")
                width_px = item.get("width", 400)
                bookmark_id = item.get("bookmark_id")

                img_path = os.path.join(spec_dir, src) if not os.path.isabs(src) else src

                lines.append("      new Paragraph({")
                lines.append("        alignment: AlignmentType.CENTER,")
                lines.append("        spacing: { before: 200, after: 100 },")
                lines.append("        keepNext: true,")
                ext = os.path.splitext(src)[1].lstrip(".").lower()
                img_type = {"jpg": "jpg", "jpeg": "jpeg", "png": "png", "gif": "gif", "bmp": "bmp", "svg": "svg"}.get(ext, "png")
                lines.append("        children: [new ImageRun({")
                lines.append(f'          type: "{img_type}",')
                lines.append(
                    f"          data: fs.readFileSync({_escape_js_string(img_path)}),"
                )
                lines.append(
                    f"          transformation: {{ width: {width_px}, height: {int(width_px * 0.75)} }},"
                )
                lines.append(f"          altText: {{ title: {_escape_js_string(caption)}, description: {_escape_js_string(caption)}, name: {_escape_js_string(caption)} }},")
                lines.append("        })],")
                lines.append("      }),")

                if caption:
                    lines.append("      new Paragraph({")
                    lines.append("        alignment: AlignmentType.CENTER,")
                    lines.append("        spacing: { after: 200 },")
                    lines.append("        keepLines: true,")
                    cap_children = _gen_caption_children(
                        caption, font, 20, bookmark_id=bookmark_id,
                        color="555555"
                    )
                    lines.append("        children: [")
                    lines.extend(cap_children)
                    lines.append("        ],")
                    lines.append("      }),")

            elif item_type == "table":
                caption = item.get("caption", "")
                bookmark_id = item.get("bookmark_id")
                columns = item.get("columns", [])
                rows = item.get("rows", [])
                num_cols = len(columns)
                if num_cols == 0:
                    continue

                col_width_base = content_width // num_cols
                col_widths = [col_width_base] * num_cols
                remainder = content_width - col_width_base * num_cols
                col_widths[-1] += remainder
                col_widths_str = ", ".join(str(w) for w in col_widths)

                if caption:
                    lines.append("      new Paragraph({")
                    lines.append("        alignment: AlignmentType.CENTER,")
                    lines.append("        spacing: { before: 200, after: 100 },")
                    lines.append("        keepNext: true,")
                    cap_children = _gen_caption_children(
                        caption, font, 20, bookmark_id=bookmark_id,
                        color="555555", bold=True
                    )
                    lines.append("        children: [")
                    lines.extend(cap_children)
                    lines.append("        ],")
                    lines.append("      }),")

                # Table
                lines.append("      new Table({")
                lines.append(
                    f"        width: {{ size: {content_width}, type: WidthType.DXA }},"
                )
                lines.append(
                    f"        columnWidths: [{col_widths_str}],"
                )
                lines.append("        rows: [")

                # Header row
                lines.append("          new TableRow({")
                lines.append("            tableHeader: true,")
                lines.append("            children: [")
                border = "BorderStyle.SINGLE"
                for idx, col_name in enumerate(columns):
                    cw = col_widths[idx]
                    lines.append("              new TableCell({")
                    lines.append(
                        f"                borders: {{ top: {{ style: {border}, size: 1, color: \"CCCCCC\" }}, bottom: {{ style: {border}, size: 1, color: \"CCCCCC\" }}, left: {{ style: {border}, size: 1, color: \"CCCCCC\" }}, right: {{ style: {border}, size: 1, color: \"CCCCCC\" }} }},"
                    )
                    lines.append(
                        f"                width: {{ size: {cw}, type: WidthType.DXA }},"
                    )
                    lines.append(
                        "                shading: { fill: \"EEEEEE\", type: ShadingType.CLEAR },"
                    )
                    lines.append(
                        "                margins: { top: 80, bottom: 80, left: 120, right: 120 },"
                    )
                    lines.append("                children: [new Paragraph({")
                    lines.append(
                        f"                  children: [new TextRun({{ text: {_escape_js_string(col_name)}, "
                        f"font: {_escape_js_string(font)}, size: 22, bold: true }})],"
                    )
                    lines.append("                })],")
                    lines.append("              }),")
                lines.append("            ],")
                lines.append("          }),")

                # Data rows
                for row in rows:
                    lines.append("          new TableRow({")
                    lines.append("            children: [")
                    for idx, cell_text in enumerate(row[:num_cols]):
                        cw = col_widths[idx]
                        lines.append("              new TableCell({")
                        lines.append(
                            f"                borders: {{ top: {{ style: {border}, size: 1, color: \"CCCCCC\" }}, bottom: {{ style: {border}, size: 1, color: \"CCCCCC\" }}, left: {{ style: {border}, size: 1, color: \"CCCCCC\" }}, right: {{ style: {border}, size: 1, color: \"CCCCCC\" }} }},"
                        )
                        lines.append(
                            f"                width: {{ size: {cw}, type: WidthType.DXA }},"
                        )
                        lines.append(
                            "                margins: { top: 80, bottom: 80, left: 120, right: 120 },"
                        )
                        lines.append("                children: [new Paragraph({")
                        lines.append(
                            f"                  children: [new TextRun({{ text: {_escape_js_string(str(cell_text))}, "
                            f"font: {_escape_js_string(font)}, size: 22 }})],"
                        )
                        lines.append("                })],")
                        lines.append("              }),")
                    lines.append("            ],")
                    lines.append("          }),")

                lines.append("        ],")
                lines.append("      }),")

            elif item_type == "list":
                ordered = item.get("ordered", False)
                ref = "numbers" if ordered else "bullets"
                list_items = item.get("items", [])

                for li in list_items:
                    lines.append("      new Paragraph({")
                    lines.append(
                        f'        numbering: {{ reference: "{ref}", level: 0 }},'
                    )
                    lines.append("        spacing: { after: 60 },")
                    lines.append(
                        f"        children: [new TextRun({{ text: {_escape_js_string(li)}, "
                        f"font: {_escape_js_string(font)}, size: 24 }})],"
                    )
                    lines.append("      }),")

    return lines


def _gen_references(spec: dict, font: str) -> list[str]:
    """Generate references section."""
    lines = []
    references = spec.get("references", [])
    if not references:
        return lines

    lines.append(
        "      new Paragraph({ children: [new PageBreak()] }),"
    )
    lines.append("      new Paragraph({")
    lines.append("        heading: HeadingLevel.HEADING_1,")
    lines.append("        children: [new TextRun({ text: \"参考文献\",")
    lines.append(
        f"          font: {_escape_js_string(font)}, size: 32, bold: true }})],"
    )
    lines.append("      }),")

    for ref in references:
        text = ref.get("text", "")
        lines.append("      new Paragraph({")
        lines.append("        spacing: { after: 60 },")
        lines.append(
            f"        children: [new TextRun({{ text: {_escape_js_string(text)}, "
            f"font: {_escape_js_string(font)}, size: 20 }})],"
        )
        lines.append("      }),")

    return lines


def _find_placeholder_run(p):
    """Find a <w:r> element in paragraph containing a math placeholder."""
    for r in p.getElementsByTagName("w:r"):
        text = _dom_get_run_text(r)
        if "##MATH_PLACEHOLDER_" in text:
            return r
    return None


def _dom_get_run_text(r) -> str:
    """Get the concatenated text content of a <w:r> element."""
    parts = []
    for t in r.getElementsByTagName("w:t"):
        if t.firstChild and t.firstChild.nodeType == t.TEXT_NODE:
            parts.append(t.firstChild.nodeValue)
    return "".join(parts)


def _replace_paragraph_content_with_omml(p, omml_node):
    """Replace paragraph content children with an OMML node, preserving w:pPr."""
    ppr = None
    for child in list(p.childNodes):
        if child.nodeType == child.ELEMENT_NODE and child.tagName == "w:pPr":
            ppr = child
        elif child.nodeType == child.ELEMENT_NODE:
            p.removeChild(child)
        elif child.nodeType == child.TEXT_NODE:
            p.removeChild(child)

    if ppr:
        p.appendChild(ppr)
    p.appendChild(omml_node)


def _latex_to_omml_node(latex: str, display: str, owner_dom):
    """Convert LaTeX to an OMML DOM node via pandoc.

    Returns <m:oMathPara> for block math, <m:oMath> for inline math.
    Returns None on failure.
    """
    if not shutil.which("pandoc"):
        print("  Warning: pandoc not found, skipping math conversion",
              file=sys.stderr)
        return None

    with tempfile.TemporaryDirectory(prefix="docx-omml-") as tmp_dir:
        tmp = Path(tmp_dir)
        md_path = tmp / "equation.md"
        docx_path = tmp / "equation.docx"

        delimiter = f"$${latex}$$" if display == "block" else f"${latex}$"
        md_path.write_text(f"{delimiter}\n", encoding="utf-8")

        result = subprocess.run(
            ["pandoc", str(md_path), "-o", str(docx_path)],
            capture_output=True, text=True,
        )
        if result.returncode != 0 or not docx_path.exists():
            return None

        with zipfile.ZipFile(docx_path) as zf:
            xml = zf.read("word/document.xml")

    equation_dom = defusedxml.minidom.parseString(xml)

    if display == "block":
        nodes = equation_dom.getElementsByTagName("m:oMathPara")
        if not nodes:
            nodes = equation_dom.getElementsByTagName("m:oMath")
        if not nodes:
            return None
    else:
        nodes = equation_dom.getElementsByTagName("m:oMath")
        if not nodes:
            return None

    node = nodes[0]

    # Clean up empty <m:sty> elements
    for sty in list(node.getElementsByTagName("m:sty")):
        parent = sty.parentNode
        parent.removeChild(sty)
        if parent.tagName == "m:rPr" and not any(
            n.nodeType == n.ELEMENT_NODE for n in parent.childNodes
        ):
            parent.parentNode.removeChild(parent)

    return owner_dom.importNode(node, deep=True)


def _post_process_math(docx_path: str) -> bool:
    """Replace math placeholders in a generated DOCX with live OMML via pandoc.

    Returns True if any formula was processed.
    """
    if not _math_placeholders:
        return False

    print("  Post-processing math formulas...", end="", flush=True)

    with zipfile.ZipFile(docx_path, "r") as zf:
        doc_xml = zf.read("word/document.xml")

    dom = defusedxml.minidom.parseString(doc_xml)

    root = dom.documentElement
    if not root.hasAttribute("xmlns:m"):
        root.setAttribute("xmlns:m", MATH_NS)

    modified = False
    paragraphs = dom.getElementsByTagName("w:p")

    for p in paragraphs:
        while True:
            placeholder_run = _find_placeholder_run(p)
            if placeholder_run is None:
                break

            marker_text = _dom_get_run_text(placeholder_run)
            match = re.search(r"##MATH_PLACEHOLDER_(\d+)##", marker_text)
            if not match:
                break

            idx = int(match.group(1))
            info = _math_placeholders.get(idx)
            if not info:
                break

            omml_node = _latex_to_omml_node(info["latex"], info["display"], dom)
            if omml_node is None:
                print(
                    f"\n  Warning: could not convert formula #{idx}: "
                    f"{info['latex'][:60]}", file=sys.stderr
                )
                break

            if info["display"] == "block":
                _replace_paragraph_content_with_omml(p, omml_node)
            else:
                p.replaceChild(omml_node, placeholder_run)

            modified = True

    if not modified:
        print(" (none found)")
        return False

    xml_bytes = dom.toxml(encoding="UTF-8")
    tmp_path = docx_path + ".math_tmp.docx"
    with zipfile.ZipFile(docx_path, "r") as zf_in:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                if item.filename == "word/document.xml":
                    zf_out.writestr(item, xml_bytes)
                else:
                    zf_out.writestr(item, zf_in.read(item.filename))

    shutil.move(tmp_path, docx_path)
    formula_count = sum(
        1 for k, v in _math_placeholders.items() if v is not None
    )
    print(f" {formula_count} formula(s) processed")
    return True


def resolve_output_path(spec: dict, input_path: str | None, args: argparse.Namespace) -> str:
    """Determine output DOCX path from CLI args or spec metadata."""
    if args.output:
        return args.output

    meta = spec.get("meta", {})
    paper_name = args.paper or meta.get("paper_name", "")
    category = args.category or meta.get("category", "")
    version = args.version or meta.get("version", 1)

    if category and paper_name:
        out_dir = Path.cwd() / "outputs" / paper_name / category / f"v{version}"
        out_dir.mkdir(parents=True, exist_ok=True)
        return str(out_dir / f"{paper_name}-generated.docx")

    if input_path:
        input_path_obj = Path(input_path)
        return str(input_path_obj.with_suffix(".docx"))

    return "output.docx"


def main():
    parser = argparse.ArgumentParser(
        description="Generate a DOCX from a structured YAML/JSON specification."
    )
    parser.add_argument(
        "input_file", nargs="?", help="YAML or JSON specification file (optional with --preview)"
    )
    parser.add_argument(
        "--output", "-o", help="Output DOCX path (default: derived from input name)"
    )
    parser.add_argument(
        "--template", "-t",
        help="Style template YAML path for defaults (page, font, caption patterns, TOC, features)",
    )
    parser.add_argument(
        "--preview", action="store_true",
        help="Generate a preview DOCX demonstrating all style features (requires --template)",
    )
    parser.add_argument(
        "--category", "-c",
        help="Template category for output path (e.g. 北交模板)",
    )
    parser.add_argument(
        "--paper", "-p",
        help="Paper name for output path",
    )
    parser.add_argument(
        "--version", "-v", type=int,
        help="Version number for output path",
    )
    parser.add_argument(
        "--humanize", action="store_true", default=None,
        help="Enable text humanization (default: on, use --no-humanize to disable)",
    )
    parser.add_argument(
        "--no-humanize", action="store_false", dest="humanize", default=None,
        help="Disable text humanization",
    )
    parser.add_argument(
        "--skip-validate", action="store_true",
        help="Skip DOCX validation",
    )
    parser.add_argument(
        "--verbose", "-V", action="store_true",
        help="Show generated JS script content",
    )
    args = parser.parse_args()

    # Load template if specified
    template = {}
    if args.template:
        tmpl_path = Path(args.template).resolve()
        if not tmpl_path.exists():
            print(f"Template not found: {args.template}", file=sys.stderr)
            sys.exit(1)
        template = load_template(str(tmpl_path))
        print(f"Loaded template: {args.template}")

    # Preview mode: generate a sample spec internally
    _preview_tmp_dir = None
    if args.preview:
        if not args.template:
            print("Error: --preview requires --template", file=sys.stderr)
            sys.exit(1)
        _preview_tmp_dir = tempfile.mkdtemp(prefix="docx-preview-")
        spec = _build_preview_spec(template, _preview_tmp_dir)
        spec_dir = _preview_tmp_dir
        output_path = args.output or str(
            Path(tmpl_path).parent / f"{tmpl_path.stem}-preview.docx"
        )
        if "page" not in spec:
            spec["page"] = template.get("page", {})
        if "style" not in spec:
            spec["style"] = {}
        defaults = template.get("document_defaults", {}).get("run", {})
        if "font" in defaults and "default_font" not in spec["style"]:
            spec["style"]["default_font"] = defaults["font"]
        print(f"Generating preview from template: {args.template}")
    else:
        # Load spec from file
        if not args.input_file:
            print("Error: input_file required (or use --preview with --template)", file=sys.stderr)
            sys.exit(1)
        spec_path = Path(args.input_file).resolve()
        if not spec_path.exists():
            print(f"Spec file not found: {args.input_file}", file=sys.stderr)
            sys.exit(1)
        spec = load_spec(str(spec_path))
        spec_dir = str(spec_path.parent)

    # Apply template defaults to spec (spec takes priority)
    if template:
        spec = _apply_template_defaults(spec, template)

    # Humanization status: spec field takes priority, falls back to CLI flag, default ON
    humanize = spec.get("humanize", True)
    if args.humanize is not None:
        humanize = args.humanize
    spec["humanize"] = humanize
    humanize_status = "ON" if humanize else "OFF"
    print(f"Humanization: {humanize_status}")

    output_path = args.output or resolve_output_path(spec, args.input_file or "preview", args)

    print(f"Output: {output_path}")

    # Generate JS script
    js_code = generate_js(spec, str(Path(output_path).resolve()), spec_dir)

    if args.verbose:
        print("\n--- Generated JS ---")
        print(js_code)
        print("--- End JS ---\n")

    # Write and execute JS
    tmp = tempfile.NamedTemporaryFile(
        mode="w", suffix=".js", delete=False, encoding="utf-8"
    )
    try:
        tmp.write(js_code)
        tmp.close()

        env = os.environ.copy()
        npm_root = subprocess.run(
            ["npm", "root", "-g"], capture_output=True, text=True
        ).stdout.strip()
        if "NODE_PATH" in env:
            env["NODE_PATH"] = npm_root + os.pathsep + env["NODE_PATH"]
        else:
            env["NODE_PATH"] = npm_root
        result = subprocess.run(
            ["node", tmp.name],
            capture_output=True,
            text=True,
            env=env,
        )
        if result.returncode != 0:
            print(f"Node.js generation failed:\n{result.stderr}", file=sys.stderr)
            sys.exit(1)
        if result.stdout:
            print(result.stdout)
    finally:
        os.unlink(tmp.name)

    # Post-process math formulas (LaTeX → OMML via pandoc)
    _post_process_math(output_path)

    # Validate
    if not args.skip_validate:
        print("Validating output...")
        if _validate_docx(output_path):
            print("Validation passed.")
        else:
            print("(Use --skip-validate to bypass)", file=sys.stderr)

    print(f"\nGenerated: {output_path}")

    # Cleanup preview temp directory
    if _preview_tmp_dir and os.path.isdir(_preview_tmp_dir):
        shutil.rmtree(_preview_tmp_dir, ignore_errors=True)


if __name__ == "__main__":
    main()
