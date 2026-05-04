from pathlib import Path
import tempfile
import unittest

from latex_docx_converter.converter import (
    ConversionConfig,
    build_pandoc_command,
    make_log_path,
    normalize_config,
    resolve_export_docx,
)
from latex_docx_converter.scanner import find_main_tex_candidates
from latex_docx_converter.tjuthesis import build_expanded_tex, extract_tjuthesis_fields


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
                "\\end{document}\n"
            )
            intro_text = "\\ctitle{题目}\n\\cauthor{作者}\n\\cabstractcn{摘要内容}\n\\ckeywordcn{关键词}"

            expanded = build_expanded_tex(root, main_text, intro_text)

            self.assertIn("本科生毕业论文", expanded)
            self.assertIn("题目：题目", expanded)
            self.assertIn("\\section*{摘 要}", expanded)
            self.assertIn("\\chapter{第一章}", expanded)


if __name__ == "__main__":
    unittest.main()
