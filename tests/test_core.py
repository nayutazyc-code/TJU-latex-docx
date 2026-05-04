from pathlib import Path
import tempfile
import unittest
from zipfile import ZIP_DEFLATED, ZipFile

from latex_docx_converter.citation import audit_citations
from latex_docx_converter.converter import (
    ConversionConfig,
    build_pandoc_command,
    make_log_path,
    normalize_config,
    resolve_export_docx,
)
from latex_docx_converter.defaults import find_default_bibliography, find_default_csl, find_default_reference_docx
from latex_docx_converter.scanner import find_main_tex_candidates
from latex_docx_converter.tjuthesis import build_expanded_tex, extract_tjuthesis_fields, prepare_tjuthesis_input
from latex_docx_converter.word_postprocess import WordPostprocessProfile, postprocess_docx


class ScannerTests(unittest.TestCase):
    def test_finds_main_tex_before_chapter_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "chapter1.tex").write_text("\\section{Intro}", encoding="utf-8")
            (root / "main.tex").write_text(
                "\\documentclass{article}\n\\begin{document}\nHello\n\\end{document}",
                encoding="utf-8",
            )

            candidates = find_main_tex_candidates(root)

            self.assertEqual(candidates[0].name, "main.tex")


class ConverterCommandTests(unittest.TestCase):
    def test_builds_pandoc_command_with_optional_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            main = root / "main.tex"
            reference = root / "reference.docx"
            bibliography = root / "refs.bib"
            csl = root / "style.csl"
            for path in (main, reference, bibliography, csl):
                path.write_text("x", encoding="utf-8")

            config = normalize_config(
                ConversionConfig(
                    project_dir=root,
                    main_tex=Path("main.tex"),
                    output_docx=root / "out.docx",
                    reference_docx=reference,
                    bibliography=bibliography,
                    csl=csl,
                )
            )

            command = build_pandoc_command(config, "/usr/bin/pandoc")

            self.assertEqual(command[0], "/usr/bin/pandoc")
            self.assertIn("main.tex", command)
            self.assertIn("-f", command)
            self.assertIn("latex", command)
            self.assertIn("-t", command)
            self.assertIn("docx", command)
            self.assertIn("--citeproc", command)
            self.assertIn(f"--reference-doc={reference.resolve()}", command)
            self.assertIn(f"--bibliography={bibliography.resolve()}", command)
            self.assertIn(f"--csl={csl.resolve()}", command)

    def test_normalizes_output_and_log_under_export_directories(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            main = root / "main.tex"
            main.write_text("\\documentclass{article}", encoding="utf-8")

            config = normalize_config(
                ConversionConfig(project_dir=root, main_tex=main, output_docx=root / "main.docx")
            )
            log_path = make_log_path(config.output_docx)

            self.assertEqual(config.output_docx, (root / "docx导出" / "main.docx").resolve())
            self.assertEqual(log_path.parent, (root / "docx导出" / "logs").resolve())

    def test_resolve_export_docx_uses_selected_filename_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            chosen = Path("/elsewhere/custom-name.docx")

            output = resolve_export_docx(chosen, root)

            self.assertEqual(output, (root / "docx导出" / "custom-name.docx").resolve())

    def test_normalize_config_uses_default_optional_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            main = root / "main.tex"
            main.write_text("\\documentclass{article}", encoding="utf-8")
            reference = root / "【2025修订版】1-1天津大学本科生毕业设计模板.docx"
            reference.write_text("x", encoding="utf-8")
            bibliography = root / "reference.bib"
            bibliography.write_text("@article{a, title={T}}", encoding="utf-8")
            csl = root / "china-national-standard-gb-t-7714-2015-numeric.csl"
            csl.write_text("x", encoding="utf-8")

            config = normalize_config(
                ConversionConfig(project_dir=root, main_tex=main, output_docx=root / "main.docx")
            )

            self.assertEqual(config.reference_docx, reference.resolve())
            self.assertEqual(config.bibliography, bibliography.resolve())
            self.assertEqual(config.csl, csl.resolve())

    def test_builds_command_with_toc_and_figures_resource_path(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "figures").mkdir()
            main = root / "main.tex"
            main.write_text("\\documentclass{article}", encoding="utf-8")
            config = normalize_config(
                ConversionConfig(project_dir=root, main_tex=main, output_docx=root / "out.docx")
            )

            command = build_pandoc_command(config, "/usr/bin/pandoc", add_toc=True)

            self.assertIn("--toc", command)
            resource_index = command.index("--resource-path") + 1
            self.assertIn(str((root / "figures").resolve()), command[resource_index])


class TjuThesisTests(unittest.TestCase):
    def test_extracts_frontmatter_fields(self):
        text = "\\ctitle{测试题目}\n\\cauthor{张三}\n\\cabstractcn{中文摘要\\\\下一段}"

        fields = extract_tjuthesis_fields(text)

        self.assertEqual(fields["title"], "测试题目")
        self.assertEqual(fields["author"], "张三")
        self.assertIn("下一段", fields["abstract_cn"])

    def test_expands_cover_abstract_and_body_includes(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            contents = root / "contents"
            contents.mkdir()
            (contents / "chapter1.tex").write_text("\\chapter{第一章}\n正文", encoding="utf-8")
            main_text = (
                "\\documentclass{tjuthesis-Bachelor}\n"
                "\\begin{document}\n"
                "\\include{contents/introduction}\n"
                "\\makecover\n"
                "\\makeabstract\n"
                "\\mainmatter\n"
                "\\include{contents/chapter1}\n"
                "\\printbibliography[heading=bibintoc]\n"
                "\\end{document}\n"
            )
            intro_text = "\\ctitle{题目}\n\\cauthor{作者}\n\\cabstractcn{摘要内容}\n\\ckeywordcn{关键词}"

            expanded = build_expanded_tex(root, main_text, intro_text)

            self.assertIn("本科生毕业论文", expanded)
            self.assertIn("题目：题目", expanded)
            self.assertIn("\\section*{摘 要}", expanded)
            self.assertIn("TJU_DOCX_TOC_PLACEHOLDER", expanded)
            self.assertIn("TJU_DOCX_BIB_PLACEHOLDER", expanded)
            self.assertIn("\\chapter{第一章}", expanded)

    def test_prepare_tjuthesis_marks_postprocessing_required(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            main = root / "main.tex"
            main.write_text(
                "\\documentclass{tjuthesis-Bachelor}\n\\begin{document}\\end{document}",
                encoding="utf-8",
            )

            with prepare_tjuthesis_input(root, main) as prepared:
                self.assertTrue(prepared.postprocess_docx)

    def test_prepare_plain_tex_does_not_mark_postprocessing_required(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            main = root / "main.tex"
            main.write_text(
                "\\documentclass{article}\n\\begin{document}\\end{document}",
                encoding="utf-8",
            )

            with prepare_tjuthesis_input(root, main) as prepared:
                self.assertFalse(prepared.postprocess_docx)


class DefaultDiscoveryTests(unittest.TestCase):
    def test_finds_default_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            (root / "reference.bib").write_text("@article{a, title={T}}", encoding="utf-8")
            ref = root / "【2025修订版】1-1天津大学本科生毕业设计模板.docx"
            ref.write_text("x", encoding="utf-8")
            csl = root / "china-national-standard-gb-t-7714-2015-numeric.csl"
            csl.write_text("x", encoding="utf-8")

            self.assertEqual(find_default_bibliography(root), (root / "reference.bib").resolve())
            self.assertEqual(find_default_reference_docx(root), ref.resolve())
            self.assertEqual(find_default_csl(root), csl.resolve())


class CitationAuditTests(unittest.TestCase):
    def test_audit_citations_reports_missing_keys(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            bib = root / "refs.bib"
            bib.write_text("@article{known, title={T}}", encoding="utf-8")

            audit = audit_citations("\\cite{known,missing}", bib)

            self.assertEqual(audit.cited_keys, ("known", "missing"))
            self.assertEqual(audit.missing_keys, ("missing",))
            self.assertTrue(audit.warnings)


class WordPostprocessTests(unittest.TestCase):
    def test_postprocess_inserts_toc_and_moves_bibliography(self):
        with tempfile.TemporaryDirectory() as tmp:
            docx = Path(tmp) / "sample.docx"
            create_minimal_docx(
                docx,
                [
                    ("目 录", None),
                    ("第一章 绪论概述", None),
                    ("第一章 绪论", "2"),
                    ("正文内容", None),
                    ("参考文献", "3"),
                    ("附 录", "2"),
                    ("附录正文", None),
                    ("[1] ZHANG San. Title[J]. Journal, 2024.", None),
                ],
            )

            result = postprocess_docx(docx, WordPostprocessProfile())
            document_xml = read_docx_xml(docx, "word/document.xml")
            settings_xml = read_docx_xml(docx, "word/settings.xml")

            self.assertTrue(result.toc_inserted)
            self.assertTrue(result.bibliography_moved)
            self.assertIn('TOC \\o "1-3" \\h \\u', document_xml)
            self.assertIn('w:val="true"', settings_xml)
            self.assertLess(document_xml.find("参考文献"), document_xml.find("[1] ZHANG"))
            self.assertLess(document_xml.find("[1] ZHANG"), document_xml.find("附  录"))


def create_minimal_docx(path: Path, paragraphs: list[tuple[str, str | None]]) -> None:
    body = "".join(make_paragraph_xml(text, style) for text, style in paragraphs)
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{body}<w:sectPr/></w:body></w:document>"
    )
    settings = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    )
    with ZipFile(path, "w", ZIP_DEFLATED) as docx:
        docx.writestr("word/document.xml", document)
        docx.writestr("word/settings.xml", settings)


def make_paragraph_xml(text: str, style: str | None = None) -> str:
    style_xml = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
    return f"<w:p>{style_xml}<w:r><w:t>{text}</w:t></w:r></w:p>"


def read_docx_xml(path: Path, name: str) -> str:
    with ZipFile(path) as docx:
        return docx.read(name).decode("utf-8", errors="ignore")


if __name__ == "__main__":
    unittest.main()
