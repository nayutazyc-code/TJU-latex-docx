from __future__ import annotations

from dataclasses import asdict, dataclass
import json
from pathlib import Path
import re

from .citation import CitationAudit, audit_citations, collect_project_tex_text, extract_bib_keys
from .docx_review import ReviewReport, parse_docx, review_docx


MAX_SOURCE_FILES = 40
MAX_SOURCE_CHARS_PER_FILE = 30000


@dataclass(frozen=True)
class AiReviewBundle:
    bundle_dir: Path
    context_path: Path
    docx_structure_path: Path
    latex_sources_path: Path
    bibliography_summary_path: Path
    prompt_path: Path


def build_ai_review_bundle(
    output_docx: Path,
    project_dir: Path | None = None,
    main_tex: Path | None = None,
    log_path: Path | None = None,
    review_report: ReviewReport | None = None,
    bibliography: Path | None = None,
    csl: Path | None = None,
    reference_docx: Path | None = None,
    citation_audit: CitationAudit | None = None,
    output_dir: Path | None = None,
) -> AiReviewBundle:
    output_docx = output_docx.expanduser().resolve()
    project_dir = project_dir.expanduser().resolve() if project_dir is not None else None
    main_tex = main_tex.expanduser().resolve() if main_tex is not None else None
    log_path = log_path.expanduser().resolve() if log_path is not None else None
    bibliography = bibliography.expanduser().resolve() if bibliography is not None else None
    csl = csl.expanduser().resolve() if csl is not None else None
    reference_docx = reference_docx.expanduser().resolve() if reference_docx is not None else None
    bundle_dir = (output_dir or output_docx.parent / "ai-review-bundle").expanduser().resolve()
    bundle_dir.mkdir(parents=True, exist_ok=True)

    if review_report is None:
        review_report = review_docx(output_docx, output_docx.parent / "review")

    if citation_audit is None and project_dir is not None:
        citation_audit = audit_citations(collect_project_tex_text(project_dir), bibliography)

    bundle = AiReviewBundle(
        bundle_dir=bundle_dir,
        context_path=bundle_dir / "review-context.md",
        docx_structure_path=bundle_dir / "docx-structure.json",
        latex_sources_path=bundle_dir / "latex-sources.md",
        bibliography_summary_path=bundle_dir / "bibliography-summary.md",
        prompt_path=bundle_dir / "prompt.md",
    )

    write_context(bundle, output_docx, project_dir, main_tex, log_path, review_report, bibliography, csl, reference_docx)
    write_docx_structure(bundle, output_docx)
    write_latex_sources(bundle, project_dir, main_tex)
    write_bibliography_summary(bundle, bibliography, citation_audit)
    write_prompt(bundle, output_docx, project_dir, main_tex, review_report)
    return bundle


def write_context(
    bundle: AiReviewBundle,
    output_docx: Path,
    project_dir: Path | None,
    main_tex: Path | None,
    log_path: Path | None,
    review_report: ReviewReport,
    bibliography: Path | None,
    csl: Path | None,
    reference_docx: Path | None,
) -> None:
    lines = [
        "# TJU 论文 AI 审稿上下文",
        "",
        "## 文件",
        f"- DOCX：`{output_docx}`",
        f"- LaTeX 项目：`{project_dir}`" if project_dir else "- LaTeX 项目：未提供",
        f"- 主 LaTeX：`{main_tex}`" if main_tex else "- 主 LaTeX：未提供",
        f"- 转换日志：`{log_path}`" if log_path else "- 转换日志：未提供",
        f"- 参考文献：`{bibliography}`" if bibliography else "- 参考文献：未提供",
        f"- CSL：`{csl}`" if csl else "- CSL：未提供",
        f"- Word 模板：`{reference_docx}`" if reference_docx else "- Word 模板：未提供",
        "",
        "## 自动格式检查摘要",
        f"- 严重问题：{review_report.error_count}",
        f"- 警告：{review_report.warning_count}",
        f"- Markdown 报告：`{review_report.markdown_path}`",
        f"- JSON 报告：`{review_report.json_path}`",
        "",
        "## 优先关注",
        "- 先看 `review/report.md` 中已有的结构化格式问题。",
        "- 再结合 `latex-sources.md` 检查章节逻辑、摘要、图表说明、参考文献一致性和送审风险。",
        "- DOCX 是 Pandoc 转换结果，封面、声明页、目录页码和复杂版式仍需要最终人工核对。",
        "",
    ]
    bundle.context_path.write_text("\n".join(lines), encoding="utf-8")


def write_docx_structure(bundle: AiReviewBundle, output_docx: Path) -> None:
    paragraphs, styles = parse_docx(output_docx)
    data = {
        "docx_path": str(output_docx),
        "paragraph_count": len(paragraphs),
        "paragraphs": [asdict(paragraph) for paragraph in paragraphs],
        "styles": styles,
    }
    bundle.docx_structure_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def write_latex_sources(bundle: AiReviewBundle, project_dir: Path | None, main_tex: Path | None) -> None:
    if project_dir is None or not project_dir.is_dir():
        bundle.latex_sources_path.write_text("# LaTeX 源码摘录\n\n未提供可读取的 LaTeX 项目目录。\n", encoding="utf-8")
        return

    files = collect_tex_files(project_dir, main_tex)
    lines = [
        "# LaTeX 源码摘录",
        "",
        f"- 项目目录：`{project_dir}`",
        f"- 收录文件数：{len(files)}",
        "",
    ]
    for tex_file in files:
        relative = tex_file.relative_to(project_dir)
        text = tex_file.read_text(encoding="utf-8", errors="ignore")
        truncated = len(text) > MAX_SOURCE_CHARS_PER_FILE
        if truncated:
            text = text[:MAX_SOURCE_CHARS_PER_FILE]
        lines.extend(
            [
                f"## `{relative}`",
                "",
                "```tex",
                text,
                "```",
                "",
            ]
        )
        if truncated:
            lines.extend(["该文件内容过长，审稿包中已截断。", ""])
    bundle.latex_sources_path.write_text("\n".join(lines), encoding="utf-8")


def collect_tex_files(project_dir: Path, main_tex: Path | None) -> list[Path]:
    files = []
    if main_tex is not None:
        resolved_main = main_tex.expanduser().resolve()
        if resolved_main.is_file():
            files.append(resolved_main)

    for tex_file in sorted(project_dir.rglob("*.tex")):
        if should_skip_tex_file(project_dir, tex_file):
            continue
        resolved = tex_file.resolve()
        if resolved not in files:
            files.append(resolved)
        if len(files) >= MAX_SOURCE_FILES:
            break
    return files


def should_skip_tex_file(project_dir: Path, tex_file: Path) -> bool:
    try:
        parts = tex_file.relative_to(project_dir).parts
    except ValueError:
        return True
    return any(part.startswith(".") or part == "docx导出" for part in parts)


def write_bibliography_summary(
    bundle: AiReviewBundle,
    bibliography: Path | None,
    citation_audit: CitationAudit | None,
) -> None:
    lines = ["# 参考文献与引用摘要", ""]
    if bibliography is None or not bibliography.is_file():
        lines.append("未提供可读取的 `.bib` 文件。")
        lines.append("")
        if citation_audit is not None and citation_audit.cited_keys:
            lines.append("## 正文引用 key")
            lines.extend(f"- `{key}`" for key in citation_audit.cited_keys)
            lines.append("")
        bundle.bibliography_summary_path.write_text("\n".join(lines), encoding="utf-8")
        return

    bib_keys = sorted(extract_bib_keys(bibliography))
    lines.extend(
        [
            f"- Bib 文件：`{bibliography}`",
            f"- Bib 条目数：{len(bib_keys)}",
            "",
        ]
    )
    if citation_audit is not None:
        lines.extend(
            [
                "## 引用统计",
                f"- 正文引用 key 数：{len(citation_audit.cited_keys)}",
                f"- 缺失 key 数：{len(citation_audit.missing_keys)}",
                "",
            ]
        )
        if citation_audit.missing_keys:
            lines.append("## 缺失引用 key")
            lines.extend(f"- `{key}`" for key in citation_audit.missing_keys)
            lines.append("")

    lines.append("## Bib 条目 key")
    lines.extend(f"- `{key}`" for key in bib_keys)
    lines.append("")
    bundle.bibliography_summary_path.write_text("\n".join(lines), encoding="utf-8")


def write_prompt(
    bundle: AiReviewBundle,
    output_docx: Path,
    project_dir: Path | None,
    main_tex: Path | None,
    review_report: ReviewReport,
) -> None:
    prompt = f"""# 天津大学本科毕业论文 AI 综合 Review Prompt

请按天津大学本科毕业设计（论文）常见规范，对这次导出的论文做一轮送审前综合 review。

## 输入材料
- DOCX 文件：`{output_docx}`
- LaTeX 项目：`{project_dir or "未提供"}`
- 主 LaTeX：`{main_tex or "未提供"}`
- 自动格式检查：`{review_report.markdown_path}`
- DOCX 结构 JSON：`{bundle.docx_structure_path}`
- LaTeX 源码摘录：`{bundle.latex_sources_path}`
- 参考文献摘要：`{bundle.bibliography_summary_path}`

## Review 要求
1. 先阅读 `review-context.md` 和 `review/report.md`，不要重复罗列已经明显正确的项目。
2. 重点检查标题层级、摘要和关键词、目录、图表题注、公式编号、参考文献、附录、致谢、页码/页眉页脚风险。
3. 检查 LaTeX 正文是否存在口语化、备忘录式表达、本地路径、内部文件名、未解释缩写、图表未在正文引用等送审风险。
4. 检查文内引用和参考文献是否存在不一致、缺失、顺序异常、英文作者格式异常。
5. 输出时按严重程度分组：必须修改、建议修改、提交前人工核对。
6. 每条问题尽量说明位置、原因和具体修改建议。

## 输出格式
请输出 Markdown：

```markdown
# 论文送审前 Review 报告

## 必须修改
- ...

## 建议修改
- ...

## 提交前人工核对
- ...

## 已检查但未发现明显问题
- ...
```
"""
    bundle.prompt_path.write_text(prompt, encoding="utf-8")


def bundle_summary(bundle: AiReviewBundle) -> str:
    return re.sub(r"\s+", " ", f"AI review bundle generated at {bundle.bundle_dir}").strip()
