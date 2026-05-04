from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
import re
import shutil
import subprocess
import tempfile


@dataclass(frozen=True)
class PreparedInput:
    main_tex: Path
    notes: tuple[str, ...] = ()
    add_toc: bool = False


@contextmanager
def prepare_tjuthesis_input(project_dir: Path, main_tex: Path):
    main_text = read_text(main_tex)
    if not is_tjuthesis_project(main_text):
        yield PreparedInput(main_tex=main_tex)
        return

    introduction_path = find_introduction_file(project_dir, main_text)
    introduction_text = read_text(introduction_path) if introduction_path else ""

    with tempfile.TemporaryDirectory(prefix="latex-docx-tjuthesis-") as tmp:
        temp_dir = Path(tmp)
        cover_image = render_pdf_cover(project_dir, main_tex, temp_dir)
        expanded = build_expanded_tex(project_dir, main_text, introduction_text, cover_image)
        expanded_path = Path(tmp) / "tjuthesis-pandoc-expanded.tex"
        expanded_path.write_text(expanded, encoding="utf-8")
        yield PreparedInput(
            main_tex=expanded_path,
            notes=(
                "TJUThesis compatibility preprocessing enabled.",
                f"Expanded temporary input: {expanded_path}",
                f"PDF cover image: {cover_image}" if cover_image else "PDF cover image: unavailable; using text fallback.",
            ),
            add_toc=False,
        )


def is_tjuthesis_project(main_text: str) -> bool:
    return "tjuthesis-Bachelor" in main_text or "\\makecover" in main_text or "\\makeabstract" in main_text


def find_introduction_file(project_dir: Path, main_text: str) -> Path | None:
    for include_name in find_command_group_values(main_text, "include") + find_command_group_values(main_text, "input"):
        if "introduction" not in include_name.lower():
            continue
        candidate = resolve_tex_path(project_dir, include_name)
        if candidate.is_file():
            return candidate
    fallback = project_dir / "contents" / "introduction.tex"
    return fallback if fallback.is_file() else None


def build_expanded_tex(
    project_dir: Path,
    main_text: str,
    introduction_text: str,
    cover_image: Path | None = None,
) -> str:
    fields = extract_tjuthesis_fields(introduction_text)
    body = extract_body_after_mainmatter(main_text)
    body = expand_includes(body, project_dir)
    body = cleanup_template_latex(body)

    parts = [
        "\\documentclass{book}",
        "\\usepackage{graphicx}",
        "\\usepackage{booktabs}",
        "\\usepackage{amsmath}",
        "\\usepackage{hyperref}",
        "\\begin{document}",
        build_cover_tex(fields, cover_image),
        "\\newpage",
        build_declaration_note(),
        "\\newpage",
        build_abstract_tex(fields),
        "\\newpage",
        build_manual_toc(body),
        "\\newpage",
        body,
        "\\end{document}",
    ]
    return "\n\n".join(part for part in parts if part.strip())


def extract_tjuthesis_fields(text: str) -> dict[str, str]:
    commands = {
        "title": "ctitle",
        "affiliation": "caffil",
        "subject": "csubject",
        "grade": "cgrade",
        "author": "cauthor",
        "student_number": "cnumber",
        "supervisor": "csupervisor",
        "abstract_cn": "cabstractcn",
        "keyword_cn": "ckeywordcn",
        "abstract_en": "cabstracten",
        "keyword_en": "ckeyworden",
    }
    return {
        field: normalize_frontmatter_text(extract_command_argument(text, command) or "")
        for field, command in commands.items()
    }


def build_cover_tex(fields: dict[str, str], cover_image: Path | None = None) -> str:
    if cover_image is not None:
        return "\n".join(
            [
                "\\begin{titlepage}",
                "\\begin{center}",
                f"\\includegraphics[width=0.95\\textwidth]{{{cover_image.as_posix()}}}",
                "\\end{center}",
                "\\end{titlepage}",
            ]
        )

    rows = [
        ("学院", fields.get("affiliation", "")),
        ("专业", fields.get("subject", "")),
        ("年级", fields.get("grade", "")),
        ("姓名", fields.get("author", "")),
        ("学号", fields.get("student_number", "")),
        ("指导教师", fields.get("supervisor", "")),
    ]
    table_rows = "\n".join(f"{label} & {value} \\\\" for label, value in rows if value)
    return "\n".join(
        [
            "\\begin{titlepage}",
            "\\begin{center}",
            "{\\LARGE 天津大学}\\\\[1em]",
            "{\\Large 本科生毕业论文}\\\\[3em]",
            f"{{\\large 题目：{fields.get('title', '')}}}\\\\[3em]",
            "\\begin{tabular}{ll}",
            table_rows,
            "\\end{tabular}",
            "\\end{center}",
            "\\end{titlepage}",
        ]
    )


def build_manual_toc(body: str) -> str:
    entries = extract_heading_entries(body)
    if not entries:
        return "\\section*{目 录}"

    lines = ["\\section*{目 录}"]
    for level, title in entries:
        indent = "\\quad " * max(level - 1, 0)
        lines.append(f"{indent}{title}\\\\")
    return "\n".join(lines)


def build_declaration_note() -> str:
    return "\n".join(
        [
            "\\section*{独创性声明}",
            "原 LaTeX 项目通过 \\texttt{\\textbackslash includepdf\\{独创性声明.pdf\\}} 插入声明页。",
            "Pandoc 不能直接把该 PDF 页面嵌入 DOCX，请在 Word 中按学校模板核对或插入声明页。",
        ]
    )


def build_abstract_tex(fields: dict[str, str]) -> str:
    return "\n".join(
        [
            "\\section*{摘 要}",
            fields.get("abstract_cn", ""),
            "",
            f"\\noindent \\textbf{{关键词：}} {fields.get('keyword_cn', '')}",
            "\\newpage",
            "\\section*{ABSTRACT}",
            fields.get("abstract_en", ""),
            "",
            f"\\noindent \\textbf{{KEY WORDS:}} {fields.get('keyword_en', '')}",
        ]
    )


def extract_body_after_mainmatter(main_text: str) -> str:
    document = text_between(main_text, "\\begin{document}", "\\end{document}") or main_text
    if "\\mainmatter" in document:
        return document.split("\\mainmatter", 1)[1]
    return document


def expand_includes(text: str, project_dir: Path) -> str:
    lines: list[str] = []
    for line in text.splitlines():
        stripped = line.strip()
        command = "include" if stripped.startswith("\\include") else "input" if stripped.startswith("\\input") else None
        if not command:
            lines.append(line)
            continue

        include_name = extract_command_argument(stripped, command)
        if not include_name:
            lines.append(line)
            continue

        include_path = resolve_tex_path(project_dir, include_name)
        if include_path.is_file():
            lines.append(f"% Expanded from {include_path.relative_to(project_dir)}")
            lines.append(expand_includes(read_text(include_path), project_dir))
        else:
            lines.append(line)
    return "\n".join(lines)


def cleanup_template_latex(text: str) -> str:
    text = strip_comments(text)
    text = replace_multi_argument_macro(text, "figuremacro", replace_figuremacro)
    text = replace_multi_argument_macro(text, "tablemacro", replace_tablemacro)
    text = replace_multi_argument_macro(text, "texorpdfstring", lambda groups: groups[0], group_count=2)
    text = number_mainmatter_chapters(text)
    text = re.sub(r"\\(?:begin|end)\{appendixenv\}", "", text)
    text = re.sub(r"\\(?:frontmatter|mainmatter|backmatter|clearpage)\b", "", text)
    text = re.sub(r"\\vspace\*?\{[^{}]*\}", "", text)
    text = re.sub(r"\\(?:songti|heiti|zihao)\*?(?:\{[^{}]*\})?", "", text)
    text = re.sub(r"\\(?:songtibf|heitibf)\{([^{}]*)\}", r"\\textbf{\1}", text)
    text = re.sub(r"\\printbibliography(?:\[[^\]]*\])?", "\\\\section*{参考文献}", text)
    return text


def number_mainmatter_chapters(text: str) -> str:
    result: list[str] = []
    index = 0
    chapter_number = 1
    numbered = True
    token = "\\chapter"

    while True:
        backmatter_index = text.find("\\backmatter", index)
        chapter_index = text.find(token, index)
        if chapter_index == -1:
            result.append(text[index:])
            return "".join(result)

        if backmatter_index != -1 and backmatter_index < chapter_index:
            result.append(text[index:backmatter_index])
            result.append("\\backmatter")
            index = backmatter_index + len("\\backmatter")
            numbered = False
            continue

        parsed = parse_command_groups(text, chapter_index, "chapter", 1)
        if parsed is None:
            result.append(text[index : chapter_index + len(token)])
            index = chapter_index + len(token)
            continue

        groups, end = parsed
        title = clean_heading_title(groups[0])
        result.append(text[index:chapter_index])
        if numbered and title:
            if title.startswith("第") and "章" in title[:4]:
                result.append(f"\\chapter{{{title}}}")
            else:
                prefix = f"第{to_chinese_number(chapter_number)}章"
                result.append(f"\\chapter{{{prefix} {title}}}")
            chapter_number += 1
        else:
            result.append(f"\\chapter{{{title}}}")
        index = end


def extract_heading_entries(text: str) -> list[tuple[int, str]]:
    entries: list[tuple[int, str]] = []
    index = 0
    commands = (("chapter", 1), ("section", 2), ("subsection", 3))

    while index < len(text):
        matches = [
            (text.find(f"\\{command}", index), command, level)
            for command, level in commands
            if text.find(f"\\{command}", index) != -1
        ]
        if not matches:
            break

        start, command, level = min(matches, key=lambda item: item[0])
        if text[start + len(command) + 1 : start + len(command) + 2] == "*":
            index = start + len(command) + 2
            continue
        parsed = parse_command_groups(text, start, command, 1)
        if parsed is None:
            index = start + len(command) + 1
            continue
        groups, end = parsed
        title = clean_heading_title(groups[0])
        if title:
            entries.append((level, title))
        index = end
    return entries


def clean_heading_title(title: str) -> str:
    title = title.replace("\\quad", " ")
    title = title.replace("\\qquad", " ")
    title = re.sub(r"\\[a-zA-Z]+\*?(?:\[[^\]]*\])?", "", title)
    title = re.sub(r"\s+", " ", title)
    return title.strip()


def to_chinese_number(number: int) -> str:
    values = {
        1: "一",
        2: "二",
        3: "三",
        4: "四",
        5: "五",
        6: "六",
        7: "七",
        8: "八",
        9: "九",
        10: "十",
    }
    return values.get(number, str(number))


def render_pdf_cover(project_dir: Path, main_tex: Path, temp_dir: Path) -> Path | None:
    pdf = main_tex.with_suffix(".pdf")
    if not pdf.is_file():
        candidates = sorted(project_dir.glob("*.pdf"))
        pdf = candidates[0] if candidates else pdf
    if not pdf.is_file() or shutil.which("sips") is None:
        return None

    cover_png = temp_dir / "tju-cover.png"
    process = subprocess.run(
        ["sips", "-s", "format", "png", str(pdf), "--out", str(cover_png)],
        text=True,
        capture_output=True,
        check=False,
    )
    if process.returncode != 0 or not cover_png.is_file():
        return None
    return cover_png


def replace_figuremacro(groups: list[str]) -> str:
    _placement, image, caption, width, label = groups
    return "\n".join(
        [
            "\\begin{figure}",
            "\\centering",
            f"\\includegraphics[width={width}\\textwidth]{{{image}}}",
            f"\\caption{{{caption}}}",
            f"\\label{{{label}}}",
            "\\end{figure}",
        ]
    )


def replace_tablemacro(groups: list[str]) -> str:
    _placement, caption, header, rows, label = groups
    return "\n".join(
        [
            "\\begin{table}",
            "\\centering",
            f"\\caption{{{caption}}}",
            "\\begin{tabular}{ccc}",
            "\\toprule",
            header,
            "\\midrule",
            rows,
            "\\bottomrule",
            "\\end{tabular}",
            f"\\label{{{label}}}",
            "\\end{table}",
        ]
    )


def replace_multi_argument_macro(text: str, command: str, replacer, group_count: int = 5) -> str:
    result: list[str] = []
    index = 0
    token = f"\\{command}"
    while True:
        start = text.find(token, index)
        if start == -1:
            result.append(text[index:])
            return "".join(result)

        parsed = parse_command_groups(text, start, command, group_count)
        if parsed is None:
            result.append(text[index : start + len(token)])
            index = start + len(token)
            continue

        groups, end = parsed
        result.append(text[index:start])
        result.append(replacer(groups))
        index = end


def parse_command_groups(text: str, start: int, command: str, group_count: int) -> tuple[list[str], int] | None:
    index = start + len(command) + 1
    groups: list[str] = []
    for _ in range(group_count):
        while index < len(text) and text[index].isspace():
            index += 1
        if index >= len(text) or text[index] != "{":
            return None
        parsed = parse_group_at(text, index)
        if parsed is None:
            return None
        group, index = parsed
        groups.append(group)
    return groups, index


def find_command_group_values(text: str, command: str) -> list[str]:
    values: list[str] = []
    index = 0
    token = f"\\{command}"
    while True:
        start = text.find(token, index)
        if start == -1:
            return values
        value = extract_command_argument(text[start:], command)
        if value is not None:
            values.append(value.strip())
        index = start + len(token)


def extract_command_argument(text: str, command: str) -> str | None:
    token = f"\\{command}"
    start = text.find(token)
    if start == -1:
        return None
    index = start + len(token)
    while index < len(text) and text[index].isspace():
        index += 1
    if index >= len(text) or text[index] != "{":
        return None
    parsed = parse_group_at(text, index)
    return parsed[0] if parsed else None


def parse_group_at(text: str, start: int) -> tuple[str, int] | None:
    if text[start] != "{":
        return None

    depth = 0
    chars: list[str] = []
    index = start
    while index < len(text):
        char = text[index]
        if char == "\\" and index + 1 < len(text):
            if depth > 0:
                chars.append(char)
                chars.append(text[index + 1])
            index += 2
            continue
        if char == "{":
            depth += 1
            if depth > 1:
                chars.append(char)
        elif char == "}":
            depth -= 1
            if depth == 0:
                return "".join(chars), index + 1
            chars.append(char)
        else:
            chars.append(char)
        index += 1
    return None


def normalize_frontmatter_text(text: str) -> str:
    text = strip_comments(text).strip()
    text = text.replace("\\hspace*{\\parindent}", "")
    text = text.replace("\\hspace{\\parindent}", "")
    text = text.replace("\\\\", "\n\n")
    text = text.replace("\\qquad", " ")
    text = text.replace("\\quad", " ")
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def strip_comments(text: str) -> str:
    return "\n".join(strip_comment_line(line) for line in text.splitlines())


def strip_comment_line(line: str) -> str:
    escaped = False
    for index, char in enumerate(line):
        if char == "\\" and not escaped:
            escaped = True
            continue
        if char == "%" and not escaped:
            return line[:index].rstrip()
        escaped = False
    return line


def text_between(text: str, start_token: str, end_token: str) -> str | None:
    start = text.find(start_token)
    if start == -1:
        return None
    start += len(start_token)
    end = text.find(end_token, start)
    if end == -1:
        return text[start:]
    return text[start:end]


def resolve_tex_path(project_dir: Path, include_name: str) -> Path:
    clean_name = include_name.strip()
    path = project_dir / clean_name
    if path.suffix:
        return path
    return path.with_suffix(".tex")


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore")
