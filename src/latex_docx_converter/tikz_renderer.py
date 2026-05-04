from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import shutil
import subprocess


@dataclass(frozen=True)
class TikzRenderReport:
    notes: tuple[str, ...] = ()
    warnings: tuple[str, ...] = ()
    rendered_count: int = 0
    failed_count: int = 0


@dataclass(frozen=True)
class TikzFigure:
    start: int
    end: int
    content: str
    caption: str
    label: str
    index: int


def render_tikz_figures(
    text: str,
    project_dir: Path,
    output_dir: Path,
    preamble_text: str = "",
) -> tuple[str, TikzRenderReport]:
    figures = find_tikz_figures(text)
    if not figures:
        return text, TikzRenderReport(notes=("No TikZ figures detected.",))

    xelatex = shutil.which("xelatex")
    gs = shutil.which("gs")
    sips = shutil.which("sips")
    if not xelatex:
        return text, TikzRenderReport(
            warnings=("TikZ figures detected, but xelatex was not found; figures were left as LaTeX source."),
            failed_count=len(figures),
        )
    if not gs and not sips:
        return text, TikzRenderReport(
            warnings=("TikZ figures detected, but neither Ghostscript nor sips was found; figures were left as LaTeX source."),
            failed_count=len(figures),
        )

    output_dir.mkdir(parents=True, exist_ok=True)
    notes: list[str] = []
    warnings: list[str] = []
    replacements: list[tuple[int, int, str]] = []
    rendered_count = 0
    failed_count = 0

    tikz_libraries = extract_tikz_libraries(preamble_text)
    for figure in figures:
        image_name = f"tikz-{safe_name(figure.label) or figure.index}.png"
        image_path = output_dir / image_name
        error = render_one_figure(
            figure,
            project_dir=project_dir,
            work_dir=output_dir / f"_build-{figure.index}",
            image_path=image_path,
            xelatex=xelatex,
            gs=gs,
            sips=sips,
            tikz_libraries=tikz_libraries,
        )
        if error is None:
            replacements.append((figure.start, figure.end, make_includegraphics_figure(figure, image_path)))
            rendered_count += 1
            notes.append(
                f"TikZ figure rendered: label={figure.label or '<none>'}, "
                f"caption={figure.caption or '<none>'}, image={image_path}, dpi=300"
            )
        else:
            failed_count += 1
            warning = (
                f"TikZ figure render failed: label={figure.label or '<none>'}, "
                f"caption={figure.caption or '<none>'}. {error} A warning placeholder was inserted."
            )
            warnings.append(warning)
            replacements.append((figure.start, figure.end, make_failed_figure_placeholder(figure, warning)))

    return apply_replacements(text, replacements), TikzRenderReport(
        notes=tuple(notes),
        warnings=tuple(warnings),
        rendered_count=rendered_count,
        failed_count=failed_count,
    )


def find_tikz_figures(text: str) -> list[TikzFigure]:
    figures: list[TikzFigure] = []
    index = 0
    figure_index = 1
    while True:
        start = text.find("\\begin{figure", index)
        if start == -1:
            return figures
        begin_end = text.find("}", start)
        if begin_end == -1:
            index = start + 1
            continue
        end_token = "\\end{figure}"
        end = text.find(end_token, begin_end)
        if end == -1:
            index = start + 1
            continue
        end += len(end_token)
        content = text[start:end]
        if "\\begin{tikzpicture}" in content or "\\begin{tikzpicture}[" in content:
            figures.append(
                TikzFigure(
                    start=start,
                    end=end,
                    content=content,
                    caption=extract_command_argument(content, "caption") or "",
                    label=extract_command_argument(content, "label") or "",
                    index=figure_index,
                )
            )
            figure_index += 1
        index = end


def render_one_figure(
    figure: TikzFigure,
    project_dir: Path,
    work_dir: Path,
    image_path: Path,
    xelatex: str,
    gs: str | None,
    sips: str | None,
    tikz_libraries: str,
) -> str | None:
    work_dir.mkdir(parents=True, exist_ok=True)
    tex_path = work_dir / "figure.tex"
    pdf_path = work_dir / "figure.pdf"
    tex_path.write_text(build_standalone_tex(figure.content, tikz_libraries), encoding="utf-8")

    compile_process = subprocess.run(
        [xelatex, "-interaction=nonstopmode", "-halt-on-error", tex_path.name],
        cwd=work_dir,
        text=True,
        capture_output=True,
        check=False,
    )
    if compile_process.returncode != 0 or not pdf_path.exists():
        output = (compile_process.stdout or "") + "\n" + (compile_process.stderr or "")
        (work_dir / "xelatex-output.log").write_text(output, encoding="utf-8")
        return f"xelatex failed: {first_error_line(output)}"

    return convert_pdf_to_png(pdf_path, image_path, project_dir, work_dir, gs, sips)


def convert_pdf_to_png(
    pdf_path: Path,
    image_path: Path,
    project_dir: Path,
    work_dir: Path,
    gs: str | None,
    sips: str | None,
) -> str | None:
    if gs:
        process = subprocess.run(
            [
                gs,
                "-dSAFER",
                "-dBATCH",
                "-dNOPAUSE",
                "-sDEVICE=pngalpha",
                "-r300",
                "-dTextAlphaBits=4",
                "-dGraphicsAlphaBits=4",
                f"-sOutputFile={image_path}",
                str(pdf_path),
            ],
            cwd=project_dir,
            text=True,
            capture_output=True,
            check=False,
        )
        if process.returncode == 0 and image_path.exists():
            return None
        output = (process.stdout or "") + "\n" + (process.stderr or "")
        (work_dir / "ghostscript-output.log").write_text(output, encoding="utf-8")
        if not sips:
            return f"ghostscript failed: {first_error_line(output)}"

    if sips:
        process = subprocess.run(
            [sips, "-s", "format", "png", str(pdf_path), "--out", str(image_path)],
            cwd=project_dir,
            text=True,
            capture_output=True,
            check=False,
        )
        if process.returncode == 0 and image_path.exists():
            return None
        output = (process.stdout or "") + "\n" + (process.stderr or "")
        (work_dir / "sips-output.log").write_text(output, encoding="utf-8")
        return f"sips failed: {first_error_line(output)}"
    return "no PDF-to-PNG converter was available"


def build_standalone_tex(figure_content: str, tikz_libraries: str) -> str:
    visual = figure_visual_content(figure_content)
    return "\n".join(
        [
            "\\documentclass[varwidth=true,border=6pt]{standalone}",
            "\\usepackage[fontset=fandol]{ctex}",
            "\\usepackage{graphicx}",
            "\\usepackage{xcolor}",
            "\\usepackage{tikz}",
            tikz_libraries,
            "\\begin{document}",
            visual,
            "\\end{document}",
        ]
    )


def figure_visual_content(figure_content: str) -> str:
    content = re.sub(r"\\begin\{figure\}(?:\[[^\]]*\])?", "", figure_content)
    content = content.replace("\\end{figure}", "")
    content = content.replace("\\centering", "")
    content = remove_command_with_group(content, "caption")
    content = remove_command_with_group(content, "label")
    return content.replace("\\textwidth", "\\linewidth").strip()


def make_includegraphics_figure(figure: TikzFigure, image_path: Path) -> str:
    lines = [
        "\\begin{figure}",
        "\\centering",
        f"\\includegraphics[width=0.95\\textwidth]{{{image_path.as_posix()}}}",
    ]
    if figure.caption:
        lines.append(f"\\caption{{{figure.caption}}}")
    if figure.label:
        lines.append(f"\\label{{{figure.label}}}")
    lines.append("\\end{figure}")
    return "\n".join(lines)


def make_failed_figure_placeholder(figure: TikzFigure, warning: str) -> str:
    lines = [
        "\\begin{figure}",
        "\\centering",
        f"\\textbf{{TikZ 图渲染失败：}} {latex_escape(warning)}",
    ]
    if figure.caption:
        lines.append(f"\\caption{{{figure.caption}}}")
    if figure.label:
        lines.append(f"\\label{{{figure.label}}}")
    lines.append("\\end{figure}")
    return "\n".join(lines)


def extract_tikz_libraries(text: str) -> str:
    libraries: list[str] = []
    for match in re.finditer(r"\\usetikzlibrary\{([^{}]+)\}", text):
        libraries.extend(part.strip() for part in match.group(1).split(",") if part.strip())
    if not libraries:
        return ""
    unique = sorted(set(libraries))
    return "\\usetikzlibrary{" + ", ".join(unique) + "}"


def apply_replacements(text: str, replacements: list[tuple[int, int, str]]) -> str:
    result: list[str] = []
    index = 0
    for start, end, replacement in sorted(replacements, key=lambda item: item[0]):
        result.append(text[index:start])
        result.append(replacement)
        index = end
    result.append(text[index:])
    return "".join(result)


def remove_command_with_group(text: str, command: str) -> str:
    result: list[str] = []
    index = 0
    token = f"\\{command}"
    while True:
        start = text.find(token, index)
        if start == -1:
            result.append(text[index:])
            return "".join(result)
        result.append(text[index:start])
        parsed = parse_command_group(text, start, command)
        if parsed is None:
            index = start + len(token)
        else:
            index = parsed[1]


def extract_command_argument(text: str, command: str) -> str | None:
    start = text.find(f"\\{command}")
    if start == -1:
        return None
    parsed = parse_command_group(text, start, command)
    return parsed[0] if parsed else None


def parse_command_group(text: str, start: int, command: str) -> tuple[str, int] | None:
    index = start + len(command) + 1
    while index < len(text) and text[index].isspace():
        index += 1
    if command == "caption" and index < len(text) and text[index] == "[":
        optional_end = text.find("]", index)
        if optional_end != -1:
            index = optional_end + 1
            while index < len(text) and text[index].isspace():
                index += 1
    if index >= len(text) or text[index] != "{":
        return None
    return parse_group_at(text, index)


def parse_group_at(text: str, start: int) -> tuple[str, int] | None:
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


def safe_name(value: str) -> str:
    clean = re.sub(r"[^A-Za-z0-9_.-]+", "-", value).strip("-")
    return clean[:80]


def latex_escape(value: str) -> str:
    return (
        value.replace("\\", "\\textbackslash{}")
        .replace("_", "\\_")
        .replace("%", "\\%")
        .replace("&", "\\&")
        .replace("#", "\\#")
    )


def first_error_line(output: str) -> str:
    for line in output.splitlines():
        stripped = line.strip()
        if stripped.startswith("!") or "Error" in stripped or "error" in stripped:
            return stripped
    for line in output.splitlines():
        stripped = line.strip()
        if stripped:
            return stripped[:160]
    return "no diagnostic output"
