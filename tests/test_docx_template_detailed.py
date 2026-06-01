import importlib.util
import struct
import tempfile
import unittest
from pathlib import Path

import defusedxml.minidom


ROOT = Path(__file__).resolve().parents[1]
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


extract = load_module(
    "extract_template",
    ROOT / ".claude/skills/docx-template-extract/scripts/extract_template.py",
)
apply = load_module(
    "apply_template",
    ROOT / ".claude/skills/docx-template-apply/scripts/apply_template.py",
)
review = load_module(
    "review_template",
    ROOT / ".claude/skills/docx-template-review/scripts/review_template.py",
)


def parse_body(inner: str):
    return defusedxml.minidom.parseString(
        f'<w:document xmlns:w="{W_NS}"><w:body>{inner}</w:body></w:document>'
    )


def first_child(elem, tag: str):
    nodes = [
        n for n in elem.childNodes
        if n.nodeType == n.ELEMENT_NODE and n.tagName == tag
    ]
    return nodes[0] if nodes else None


def paragraph_text(p):
    return "".join(
        n.firstChild.nodeValue
        for n in p.getElementsByTagName("w:t")
        if n.firstChild
    )


class DetailedDocxTemplateTests(unittest.TestCase):
    def test_extracts_body_paragraph_direct_font_formatting(self):
        dom = parse_body(
            """
            <w:p>
              <w:pPr>
                <w:pStyle w:val="Body"/>
                <w:jc w:val="both"/>
                <w:rPr>
                  <w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体" w:cs="宋体"/>
                  <w:sz w:val="21"/>
                </w:rPr>
              </w:pPr>
              <w:r>
                <w:rPr>
                  <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>
                  <w:sz w:val="21"/>
                </w:rPr>
                <w:t>正文 text</w:t>
              </w:r>
            </w:p>
            """
        )

        formats = extract.extract_body_paragraph_formats(dom)

        self.assertEqual(formats[0]["style"], "Body")
        self.assertEqual(formats[0]["paragraph"]["alignment"], "both")
        self.assertEqual(formats[0]["paragraph_run"]["font"], "宋体")
        self.assertEqual(formats[0]["run"]["font"], "Times New Roman")
        self.assertEqual(formats[0]["run"]["fontEastAsia"], "宋体")
        self.assertEqual(formats[0]["runs"][0]["run"]["fontSize"], 21)

    def test_apply_body_paragraph_formats_sets_direct_fonts_by_index(self):
        dom = parse_body(
            """
            <w:p>
              <w:r><w:t>标题</w:t></w:r>
            </w:p>
            """
        )
        formats = [
            {
                "index": 0,
                "style": "Normal",
                "paragraph": {"alignment": "center"},
                "paragraph_run": {
                    "font": "黑体",
                    "fontEastAsia": "黑体",
                    "fontSize": 28,
                },
                "run": {
                    "font": "Times New Roman",
                    "fontEastAsia": "黑体",
                    "fontSize": 28,
                },
                "runs": [
                    {
                        "index": 0,
                        "run": {
                            "font": "Times New Roman",
                            "fontEastAsia": "黑体",
                            "fontSize": 28,
                        },
                    }
                ],
            }
        ]

        changed = apply.apply_body_paragraph_formats(dom, formats)

        self.assertTrue(changed)
        paragraph = dom.getElementsByTagName("w:p")[0]
        ppr = first_child(paragraph, "w:pPr")
        self.assertIsNotNone(ppr)
        self.assertEqual(ppr.getElementsByTagName("w:jc")[0].getAttribute("w:val"), "center")
        ppr_fonts = ppr.getElementsByTagName("w:rPr")[0].getElementsByTagName("w:rFonts")[0]
        self.assertEqual(ppr_fonts.getAttribute("w:eastAsia"), "黑体")
        run = paragraph.getElementsByTagName("w:r")[0]
        run_fonts = run.getElementsByTagName("w:rPr")[0].getElementsByTagName("w:rFonts")[0]
        self.assertEqual(run_fonts.getAttribute("w:ascii"), "Times New Roman")
        self.assertEqual(run.getElementsByTagName("w:sz")[0].getAttribute("w:val"), "28")

    def test_review_reports_paragraph_run_direct_formatting(self):
        dom = parse_body(
            """
            <w:p>
              <w:pPr>
                <w:rPr>
                  <w:rFonts w:ascii="宋体" w:eastAsia="宋体"/>
                  <w:sz w:val="21"/>
                </w:rPr>
              </w:pPr>
              <w:r><w:t>正文</w:t></w:r>
            </w:p>
            """
        )

        with tempfile.TemporaryDirectory() as tmp:
            word_dir = Path(tmp)
            (word_dir / "document.xml").write_bytes(dom.toxml(encoding="UTF-8"))

            result = review.check_body_direct_formatting(word_dir)

        self.assertEqual(result.status, "WARN")
        self.assertIn("paragraph-run-level", result.summary)

    def test_apply_semantic_rules_replaces_wrong_title_font(self):
        dom = parse_body(
            """
            <w:p>
              <w:pPr><w:jc w:val="center"/></w:pPr>
              <w:r>
                <w:rPr>
                  <w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/>
                  <w:sz w:val="28"/>
                  <w:b/>
                </w:rPr>
                <w:t>道路交通场景双级样条网络三维目标检测</w:t>
              </w:r>
            </w:p>
            """
        )
        template = {
            "semantic_format_rules": [
                {
                    "role": "cn_title",
                    "match": {"type": "front_index", "index": 0},
                    "run": {
                        "font": "黑体;SimHei",
                        "fontEastAsia": "黑体;SimHei",
                        "fontSize": 40,
                        "fontSizeCs": 40,
                        "bold": True,
                    },
                    "paragraph": {"alignment": "center"},
                }
            ]
        }

        changed = apply.apply_semantic_format_rules(dom, template)

        self.assertTrue(changed)
        run = dom.getElementsByTagName("w:r")[0]
        rpr = first_child(run, "w:rPr")
        fonts = rpr.getElementsByTagName("w:rFonts")[0]
        self.assertEqual(fonts.getAttribute("w:eastAsia"), "黑体;SimHei")
        self.assertEqual(rpr.getElementsByTagName("w:sz")[0].getAttribute("w:val"), "40")

    def test_strict_review_fails_wrong_semantic_font(self):
        dom = parse_body(
            """
            <w:p>
              <w:pPr><w:jc w:val="center"/></w:pPr>
              <w:r>
                <w:rPr>
                  <w:rFonts w:ascii="Times New Roman" w:eastAsia="宋体"/>
                  <w:sz w:val="28"/>
                </w:rPr>
                <w:t>道路交通场景双级样条网络三维目标检测</w:t>
              </w:r>
            </w:p>
            """
        )
        template = {
            "semantic_format_rules": [
                {
                    "role": "cn_title",
                    "match": {"type": "front_index", "index": 0},
                    "run": {
                        "font": "黑体;SimHei",
                        "fontEastAsia": "黑体;SimHei",
                        "fontSize": 40,
                    },
                }
            ]
        }

        with tempfile.TemporaryDirectory() as tmp:
            word_dir = Path(tmp)
            (word_dir / "document.xml").write_bytes(dom.toxml(encoding="UTF-8"))

            result = review.check_semantic_format_rules(word_dir, template)

        self.assertEqual(result.status, "FAIL")
        self.assertIn("fontEastAsia", "\n".join(result.details))

    def test_bjtu_semantic_rules_match_requested_journal_formats(self):
        template = {
            "body_paragraph_formats": [
                {"textSample": "北京交通大学学报"},
                {"textSample": "一级标题二号黑体"},
            ]
        }

        rules = extract.build_semantic_format_rules(template)
        by_role = {rule["role"]: rule for rule in rules}

        self.assertIn("en_abstract", by_role)
        self.assertEqual(
            by_role["en_abstract"]["prefixes"][0]["run"]["bold"],
            True,
        )
        self.assertNotIn("bold", by_role["subsection_heading"]["run"])
        self.assertEqual(by_role["table_caption_en"]["paragraph"]["alignment"], "center")
        self.assertEqual(by_role["table_caption_en"]["run"]["font"], "Times New Roman")
        self.assertEqual(by_role["table_caption_en"]["run"]["fontSize"], 18)
        self.assertNotIn("bold", by_role["table_caption_en"]["run"])
        self.assertNotIn("bold", by_role["table_cell"]["run"])
        self.assertEqual(by_role["reference_entry"]["paragraph"]["indent_left"], 425)
        self.assertEqual(by_role["reference_entry"]["paragraph"]["indent_hanging"], 425)

    def test_bjtu_normalization_removes_intro_renumbers_headings_and_conclusion_items(self):
        dom = parse_body(
            """
            <w:p><w:r><w:t>题名</w:t></w:r></w:p>
            <w:p><w:r><w:t>作者</w:t></w:r></w:p>
            <w:p><w:r><w:t>单位</w:t></w:r></w:p>
            <w:p><w:r><w:t>摘要：文本</w:t></w:r></w:p>
            <w:p><w:r><w:t>关键词：词</w:t></w:r></w:p>
            <w:p><w:r><w:t>中图分类号：U491</w:t></w:r></w:p>
            <w:p><w:r><w:t>English Title</w:t></w:r></w:p>
            <w:p><w:r><w:t>AUTHOR</w:t></w:r></w:p>
            <w:p><w:r><w:t>Affiliation</w:t></w:r></w:p>
            <w:p><w:r><w:t>Abstract: Text</w:t></w:r></w:p>
            <w:p><w:r><w:t>Key words: word</w:t></w:r></w:p>
            <w:p><w:r><w:t>1 引言</w:t></w:r></w:p>
            <w:p><w:r><w:t>正文引用([1-2])和普通引用[14]。</w:t></w:r></w:p>
            <w:p><w:r><w:t>2 本文算法</w:t></w:r></w:p>
            <w:p><w:r><w:t>2.1 算法总体架构</w:t></w:r></w:p>
            <w:p><w:r><w:t>4 结论</w:t></w:r></w:p>
            <w:p><w:r><w:t>第一条结论</w:t></w:r></w:p>
            <w:p><w:r><w:t>第二条结论</w:t></w:r></w:p>
            <w:p><w:r><w:t>参考文献</w:t></w:r></w:p>
            <w:p><w:r><w:t>[1] 作者. 文题[J].</w:t></w:r></w:p>
            """
        )

        changed = apply.apply_bjtu_journal_normalizations(dom, {}, None)

        self.assertTrue(changed)
        texts = [paragraph_text(p) for p in dom.getElementsByTagName("w:p") if paragraph_text(p)]
        self.assertNotIn("1 引言", texts)
        self.assertIn("1 本文算法", texts)
        self.assertIn("1.1 算法总体架构", texts)
        self.assertIn("3 结论", texts)
        self.assertIn("1）第一条结论", texts)
        self.assertIn("2）第二条结论", texts)
        self.assertTrue(any("正文引用[1-2]和普通引用[14]。" == t for t in texts))

        citation_runs = [
            r for r in dom.getElementsByTagName("w:r")
            if paragraph_text(r) in {"[1-2]", "[14]"}
        ]
        self.assertEqual(len(citation_runs), 2)
        for run in citation_runs:
            vert = run.getElementsByTagName("w:vertAlign")[0]
            self.assertEqual(vert.getAttribute("w:val"), "superscript")

    def test_bjtu_formula_image_is_replaced_with_omml_and_alt_text_is_cleared(self):
        dom = parse_body(
            """
            <w:p>
              <w:r>
                <w:drawing>
                  <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
                    <wp:docPr id="1" name="图片 2" descr="2026-05-09"/>
                    <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                      <a:graphicData>
                        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                          <pic:nvPicPr>
                            <pic:cNvPr id="1" name="图片 2" descr="2026-05-09"/>
                          </pic:nvPicPr>
                          <pic:blipFill>
                            <a:blip r:embed="rIdFormula" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
                          </pic:blipFill>
                        </pic:pic>
                      </a:graphicData>
                    </a:graphic>
                  </wp:inline>
                </w:drawing>
              </w:r>
            </w:p>
            """
        )

        with tempfile.TemporaryDirectory() as tmp:
            word_dir = Path(tmp)
            (word_dir / "_rels").mkdir()
            (word_dir / "media").mkdir()
            (word_dir / "_rels" / "document.xml.rels").write_text(
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rIdFormula" Type="x" Target="media/image2.png"/>'
                "</Relationships>",
                encoding="utf-8",
            )
            png_header = b"\x89PNG\r\n\x1a\n" + b"\x00\x00\x00\rIHDR" + struct.pack(">II", 189, 83)
            (word_dir / "media" / "image2.png").write_bytes(png_header)

            original = apply._latex_to_omml_para
            apply._latex_to_omml_para = lambda latex, owner: defusedxml.minidom.parseString(
                f'<w:document xmlns:w="{W_NS}" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">'
                "<w:body><m:oMathPara><m:oMath><m:r><m:t>f</m:t></m:r></m:oMath></m:oMathPara></w:body></w:document>"
            ).getElementsByTagName("m:oMathPara")[0]
            try:
                changed = apply.apply_bjtu_journal_normalizations(dom, {}, word_dir)
            finally:
                apply._latex_to_omml_para = original

        self.assertTrue(changed)
        self.assertEqual(len(dom.getElementsByTagName("w:drawing")), 0)
        self.assertEqual(len(dom.getElementsByTagName("m:oMathPara")), 1)


if __name__ == "__main__":
    unittest.main()
