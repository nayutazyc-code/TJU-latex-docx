from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
import re
import tempfile

from .tikz_renderer import render_tikz_figures


@dataclass(frozen=True)
class PreparedInput:
    main_tex: Path
    notes: tuple[str, ...] = ()
    warnings: tuple[str, ...] = ()
    add_toc: bool = False
    postprocess_docx: bool = False


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
        tikz_output_dir = temp_dir / "rendered-tikz"
        expanded, tikz_report = build_expanded_tex(
            project_dir,
            main_text,
            introduction_text,
            tikz_output_dir=tikz_output_dir,
            return_tikz_report=True,
        )
        expanded_path = temp_dir / "tjuthesis-pandoc-expanded.tex"
        expanded_path.write_text(expanded, encoding="utf-8")
        yield PreparedInput(
            main_tex=expanded_path,
            notes=(
                "TJUThesis compatibility preprocessing enabled.",
                f"Expanded temporary input: {expanded_path}",
                "Cover uses editable text fallback; please check final cover against the school template.",
                *tikz_report.notes,
            ),
            warnings=tikz_report.warnings,
            add_toc=False,
            postprocess_docx=True,
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
    tikz_output_dir: Path | None = None,
    return_tikz_report: bool = False,
):
    fields = extract_tjuthesis_fields(introduction_text)
    body = extract_body_after_mainmatter(main_text)
    body = expand_includes(body, project_dir)
    body = cleanup_template_latex(body)
    if tikz_output_dir is not None:
        body, tikz_report = render_tikz_figures(
            body,
            project_dir=project_dir,
            output_dir=tikz_output_dir,
            preamble_text=main_text,
        )
    else:
        from .tikz_renderer import TikzRenderReport

        tikz_report = TikzRenderReport()

    parts = [
        "\\documentclass{book}",
        "\\usepackage{graphicx}",
        "\\usepackage{booktabs}",
        "\\usepackage{amsmath}",
        "\\usepackage{hyperref}",
        "\\begin{document}",
        build_frontmatter_placeholder(),
        "\\newpage",
        build_abstract_tex(fields),
        "\\newpage",
        build_toc_placeholder(),
        "\\newpage",
        body,
        "\\end{document}",
    ]
    expanded = "\n\n".join(part for part in parts if part.strip())
    if return_tikz_report:
        return expanded, tikz_report
    return expanded


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


def build_toc_placeholder() -> str:
    return "\\section*{目  录}\n\nTJU_DOCX_TOC_PLACEHOLDER"


def build_frontmatter_placeholder() -> str:
    return "\n".join(
        [
            "\\section*{封面与独创性声明}",
            "请复制粘贴学校 Word 模板中对应的封面、独创性声明部分，并在最终提交前按学院要求核对。",
        ]
    )


def build_declaration_note() -> str:
    return "\n".join(
        [
            "\\section*{独创性声明}",
            "原 LaTeX 项目通过 \\texttt{\\textbackslash includepdf\\{独创性声明.pdf\\}} 插入声明页。",
            "Pandoc 不能直接把该 PDF 页面嵌入 DOCX，请在 Word 中按学校模板核对或插入声明页。",
        ]
    )


def build_abstract_tex(fields: dict[str, str]) -> str:
    keyword_cn = normalize_chinese_keywords(fields.get("keyword_cn", ""))
    keyword_en = normalize_english_keywords(fields.get("keyword_en", ""))
    return "\n".join(
        [
            "\\section*{摘 要}",
            fields.get("abstract_cn", ""),
            "",
            f"\\noindent \\textbf{{关键词：}} {keyword_cn}",
            "\\newpage",
            "\\section*{ABSTRACT}",
            fields.get("abstract_en", ""),
            "",
            f"\\noindent \\textbf{{KEY WORDS:}} {keyword_en}",
        ]
    )


def normalize_chinese_keywords(value: str) -> str:
    keywords = split_keyword_text(value)
    return "，".join(keywords) if keywords else value.strip()


def normalize_english_keywords(value: str) -> str:
    keywords = split_keyword_text(value)
    return "; ".join(capitalize_keyword(keyword) for keyword in keywords) if keywords else value.strip()


def split_keyword_text(value: str) -> list[str]:
    cleaned = re.sub(r"^(关键词|关键字|KEY\s*WORDS?)\s*[:：]\s*", "", value.strip(), flags=re.IGNORECASE)
    return [part.strip() for part in re.split(r"[，,；;、]+", cleaned) if part.strip()]


def capitalize_keyword(value: str) -> str:
    for index, char in enumerate(value):
        if char.isalpha():
            return value[:index] + char.upper() + value[index + 1 :]
    return value


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
    text = number_mainmatter_sections(text)
    text = number_float_captions(text)
    text = re.sub(r"\\(?:begin|end)\{appendixenv\}", "", text)
    text = re.sub(r"\\(?:frontmatter|mainmatter|backmatter|clearpage)\b", "", text)
    text = re.sub(r"\\vspace\*?\{[^{}]*\}", "", text)
    text = re.sub(r"\\(?:songti|heiti|zihao)\*?(?:\{[^{}]*\})?", "", text)
    text = re.sub(r"\\(?:songtibf|heitibf)\{([^{}]*)\}", r"\\textbf{\1}", text)
    text = re.sub(
        r"\\printbibliography(?:\[[^\]]*\])?",
        "\\\\section*{参考文献}\n\nTJU_DOCX_BIB_PLACEHOLDER",
        text,
    )
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


def number_mainmatter_sections(text: str) -> str:
    result: list[str] = []
    index = 0
    chapter_number = 0
    section_number = 0
    subsection_number = 0
    numbered = True
    commands = ("chapter", "section", "subsection")

    while index < len(text):
        matches = [
            (text.find(f"\\{command}", index), command)
            for command in commands
            if text.find(f"\\{command}", index) != -1
        ]
        backmatter_index = text.find("\\backmatter", index)
        if backmatter_index != -1:
            matches.append((backmatter_index, "backmatter"))
        if not matches:
            result.append(text[index:])
            return "".join(result)

        start, command = min(matches, key=lambda item: item[0])
        result.append(text[index:start])
        if command == "backmatter":
            result.append("\\backmatter")
            index = start + len("\\backmatter")
            numbered = False
            continue

        if text[start + len(command) + 1 : start + len(command) + 2] == "*":
            result.append(text[start : start + len(command) + 2])
            index = start + len(command) + 2
            continue

        parsed = parse_command_groups(text, start, command, 1)
        if parsed is None:
            result.append(text[start : start + len(command) + 1])
            index = start + len(command) + 1
            continue

        groups, end = parsed
        title = clean_heading_title(groups[0])
        if command == "chapter":
            if numbered:
                chapter_number = chapter_number_from_title(title) or chapter_number + 1
                section_number = 0
                subsection_number = 0
            result.append(f"\\chapter{{{title}}}")
        elif command == "section":
            if numbered and chapter_number:
                section_number += 1
                subsection_number = 0
                title = strip_heading_number(title)
                result.append(f"\\section{{{chapter_number}.{section_number}  {title}}}")
            else:
                result.append(f"\\section{{{title}}}")
        elif command == "subsection":
            if numbered and chapter_number and section_number:
                subsection_number += 1
                title = strip_heading_number(title)
                result.append(f"\\subsection{{{chapter_number}.{section_number}.{subsection_number}  {title}}}")
            else:
                result.append(f"\\subsection{{{title}}}")
        index = end

    return "".join(result)


def chapter_number_from_title(title: str) -> int | None:
    match = re.match(r"^第([一二三四五六七八九十百\d]+)章", title)
    if not match:
        return None
    value = match.group(1)
    if value.isdigit():
        return int(value)
    return chinese_number_to_int(value)


def chinese_number_to_int(value: str) -> int | None:
    mapping = {
        "一": 1,
        "二": 2,
        "三": 3,
        "四": 4,
        "五": 5,
        "六": 6,
        "七": 7,
        "八": 8,
        "九": 9,
        "十": 10,
    }
    return mapping.get(value)


def strip_heading_number(title: str) -> str:
    return re.sub(r"^\d+(?:\.\d+)*\s+", "", title).strip()


def number_float_captions(text: str) -> str:
    result: list[str] = []
    index = 0
    chapter_number = 0
    figure_number = 0
    table_number = 0
    numbered = True

    while index < len(text):
        matches = find_next_float_caption_targets(text, index)
        if not matches:
            result.append(text[index:])
            return "".join(result)

        start, kind = min(matches, key=lambda item: item[0])
        result.append(text[index:start])
        if kind == "backmatter":
            result.append("\\backmatter")
            index = start + len("\\backmatter")
            numbered = False
            continue
        if kind == "chapter":
            parsed = parse_command_groups(text, start, "chapter", 1)
            if parsed is None:
                result.append(text[start : start + len("\\chapter")])
                index = start + len("\\chapter")
                continue
            groups, end = parsed
            chapter_number = chapter_number_from_title(clean_heading_title(groups[0])) or chapter_number + 1
            figure_number = 0
            table_number = 0
            result.append(text[start:end])
            index = end
            continue

        environment_end = find_environment_end(text, start, kind)
        if environment_end is None:
            result.append(text[start:])
            return "".join(result)
        environment_text = text[start:environment_end]
        if numbered and chapter_number:
            if kind == "figure":
                figure_number += 1
                environment_text = number_caption_in_environment(
                    environment_text,
                    f"图{chapter_number}-{figure_number}",
                )
            elif kind == "table":
                table_number += 1
                environment_text = number_caption_in_environment(
                    environment_text,
                    f"表{chapter_number}-{table_number}",
                )
        result.append(environment_text)
        index = environment_end

    return "".join(result)


def find_next_float_caption_targets(text: str, index: int) -> list[tuple[int, str]]:
    targets: list[tuple[int, str]] = []
    for token, kind in (
        ("\\chapter", "chapter"),
        ("\\backmatter", "backmatter"),
        ("\\begin{figure", "figure"),
        ("\\begin{table", "table"),
    ):
        found = text.find(token, index)
        if found != -1:
            targets.append((found, kind))
    return targets


def find_environment_end(text: str, start: int, environment: str) -> int | None:
    token = f"\\end{{{environment}}}"
    end = text.find(token, start)
    if end == -1:
        return None
    return end + len(token)


def number_caption_in_environment(text: str, prefix: str) -> str:
    parsed = parse_caption_command(text)
    if parsed is None:
        return text
    caption, group_start, group_end = parsed
    clean_caption = clean_heading_title(caption)
    if re.match(r"^(图|表)\s*\d+[-.]\d+", clean_caption):
        return text
    return text[:group_start] + "{" + f"{prefix}  {clean_caption}" + "}" + text[group_end:]


def parse_caption_command(text: str) -> tuple[str, int, int] | None:
    start = text.find("\\caption")
    if start == -1:
        return None
    index = start + len("\\caption")
    while index < len(text) and text[index].isspace():
        index += 1
    if index < len(text) and text[index] == "[":
        optional_end = text.find("]", index)
        if optional_end == -1:
            return None
        index = optional_end + 1
    while index < len(text) and text[index].isspace():
        index += 1
    if index >= len(text) or text[index] != "{":
        return None
    parsed = parse_group_at(text, index)
    if parsed is None:
        return None
    caption, end = parsed
    return caption, index, end


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
