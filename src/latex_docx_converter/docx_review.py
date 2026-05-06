from __future__ import annotations

from dataclasses import asdict, dataclass
import json
from pathlib import Path
import re
from zipfile import ZipFile
import xml.etree.ElementTree as ET

from .word_postprocess import NS, W_NS, current_style, element_text, q


@dataclass(frozen=True)
class ParagraphInfo:
    index: int
    text: str
    style: str | None
    alignment: str | None
    spacing: dict[str, str]
    indentation: dict[str, str]
    run_properties: dict[str, str]
    page_break_before: bool
    has_math: bool
    math_text: str
    has_toc_field: bool
    tab_stops: tuple[dict[str, str], ...]


@dataclass(frozen=True)
class ReviewIssue:
    severity: str
    rule_id: str
    message: str
    paragraph_index: int | None = None
    text: str | None = None
    expected: str | None = None
    actual: str | None = None


@dataclass(frozen=True)
class ReviewReport:
    docx_path: Path
    markdown_path: Path
    json_path: Path
    issues: tuple[ReviewIssue, ...]

    @property
    def error_count(self) -> int:
        return sum(1 for issue in self.issues if issue.severity == "error")

    @property
    def warning_count(self) -> int:
        return sum(1 for issue in self.issues if issue.severity == "warning")


def review_docx(docx_path: Path, output_dir: Path | None = None) -> ReviewReport:
    docx_path = docx_path.expanduser().resolve()
    report_dir = (output_dir or docx_path.parent / "review").expanduser().resolve()
    report_dir.mkdir(parents=True, exist_ok=True)

    paragraphs, styles = parse_docx(docx_path)
    issues = collect_review_issues(paragraphs, styles)

    report = ReviewReport(
        docx_path=docx_path,
        markdown_path=report_dir / "report.md",
        json_path=report_dir / "report.json",
        issues=tuple(issues),
    )
    write_review_report(report)
    return report


def parse_docx(docx_path: Path) -> tuple[tuple[ParagraphInfo, ...], dict[str, dict[str, dict[str, str] | str | None]]]:
    with ZipFile(docx_path, "r") as docx:
        document_root = ET.fromstring(docx.read("word/document.xml"))
        styles_root = (
            ET.fromstring(docx.read("word/styles.xml"))
            if "word/styles.xml" in docx.namelist()
            else ET.Element(q("styles"))
        )

    paragraphs = tuple(
        paragraph_info(index, paragraph)
        for index, paragraph in enumerate(document_root.findall(".//w:p", NS), start=1)
    )
    styles = parse_styles(styles_root)
    return paragraphs, styles


def paragraph_info(index: int, paragraph: ET.Element) -> ParagraphInfo:
    ppr = paragraph.find("w:pPr", NS)
    alignment = child_attr(ppr, "jc", "val")
    spacing = child_attrs(ppr, "spacing")
    indentation = child_attrs(ppr, "ind")
    return ParagraphInfo(
        index=index,
        text=element_text(paragraph),
        style=current_style(paragraph),
        alignment=alignment,
        spacing=spacing,
        indentation=indentation,
        run_properties=run_properties(paragraph.find("w:r/w:rPr", NS)),
        page_break_before=ppr is not None and ppr.find("w:pageBreakBefore", NS) is not None,
        has_math=paragraph.find(".//m:oMath", NS) is not None,
        math_text="".join(node.text or "" for node in paragraph.findall(".//m:t", NS)),
        has_toc_field=bool("TOC " in "".join(node.text or "" for node in paragraph.findall(".//w:instrText", NS))),
        tab_stops=tuple(tab_attrs(tab) for tab in paragraph.findall("w:pPr/w:tabs/w:tab", NS)),
    )


def parse_styles(root: ET.Element) -> dict[str, dict[str, dict[str, str] | str | None]]:
    styles: dict[str, dict[str, dict[str, str] | str | None]] = {}
    for style in root.findall("w:style", NS):
        style_id = style.get(q("styleId"))
        if not style_id:
            continue
        styles[style_id] = {
            "spacing": child_attrs(style.find("w:pPr", NS), "spacing"),
            "indentation": child_attrs(style.find("w:pPr", NS), "ind"),
            "run": run_properties(style.find("w:rPr", NS)),
            "based_on": child_attr(style, "basedOn", "val"),
        }
    return styles


def collect_review_issues(
    paragraphs: tuple[ParagraphInfo, ...],
    styles: dict[str, dict[str, dict[str, str] | str | None]],
) -> list[ReviewIssue]:
    issues: list[ReviewIssue] = []
    check_toc(paragraphs, issues)
    check_headings(paragraphs, issues)
    check_abstracts(paragraphs, issues)
    check_captions(paragraphs, issues)
    check_equations(paragraphs, issues)
    check_backmatter(paragraphs, styles, issues)
    check_bibliography(paragraphs, issues)
    return issues


def check_toc(paragraphs: tuple[ParagraphInfo, ...], issues: list[ReviewIssue]) -> None:
    if any(paragraph.has_toc_field for paragraph in paragraphs):
        return
    issues.append(
        ReviewIssue(
            severity="warning",
            rule_id="toc.field.missing",
            message="没有检测到 Word 自动目录字段，请确认目录不是静态文本。",
            expected='TOC \\o "1-3" \\h \\u',
            actual="未检测到 TOC 字段",
        )
    )


def check_headings(paragraphs: tuple[ParagraphInfo, ...], issues: list[ReviewIssue]) -> None:
    for paragraph in paragraphs:
        text = normalized_visible_text(paragraph.text)
        if is_chapter_heading(text):
            expect(paragraph, issues, "heading.chapter.style", paragraph.style == "37", "章标题应使用模板大标题样式。", "37", paragraph.style)
            expect(paragraph, issues, "heading.chapter.align", paragraph.alignment == "center", "章标题应居中。", "center", paragraph.alignment)
            expect_spacing(paragraph, issues, "heading.chapter.before", "before", "600", "章标题段前应为 30 磅。")
            expect_spacing(paragraph, issues, "heading.chapter.after", "after", "600", "章标题段后应为 30 磅。")
            expect_spacing(paragraph, issues, "heading.chapter.line", "line", "240", "章标题应为单倍行距。")
            expect(paragraph, issues, "heading.chapter.page_break", paragraph.page_break_before, "章标题前应分页。", "pageBreakBefore", "missing")
            expect(paragraph, issues, "heading.chapter.number_space", has_heading_number_gap(paragraph.text, 1), "章标题编号和标题文字之间应空两格。", "第一章  标题", paragraph.text)
        elif is_section_heading(text):
            expect(paragraph, issues, "heading.section.style", paragraph.style == "38", "二级标题应使用模板二级标题样式。", "38", paragraph.style)
            expect(paragraph, issues, "heading.section.align", paragraph.alignment in {"left", None}, "二级标题应顶格左对齐。", "left", paragraph.alignment)
            expect_spacing(paragraph, issues, "heading.section.line", "line", "240", "二级标题应为单倍行距。")
            expect_indent_zero(paragraph, issues, "heading.section.indent", "二级标题不应首行缩进。")
            expect(paragraph, issues, "heading.section.number_space", has_heading_number_gap(paragraph.text, 2), "二级标题编号和标题文字之间应空两格。", "1.1  标题", paragraph.text)
        elif is_subsection_heading(text):
            expect(paragraph, issues, "heading.subsection.style", paragraph.style == "39", "三级标题应使用模板三级标题样式。", "39", paragraph.style)
            expect(paragraph, issues, "heading.subsection.align", paragraph.alignment in {"left", None}, "三级标题应顶格左对齐。", "left", paragraph.alignment)
            expect_indent_zero(paragraph, issues, "heading.subsection.indent", "三级标题不应首行缩进。")
            expect(paragraph, issues, "heading.subsection.number_space", has_heading_number_gap(paragraph.text, 3), "三级标题编号和标题文字之间应空两格。", "1.1.1  标题", paragraph.text)


def check_abstracts(paragraphs: tuple[ParagraphInfo, ...], issues: list[ReviewIssue]) -> None:
    cn_title = next((paragraph for paragraph in paragraphs if normalized_visible_text(paragraph.text) in {"摘 要", "摘  要", "摘要"}), None)
    en_title = next((paragraph for paragraph in paragraphs if normalized_visible_text(paragraph.text) == "ABSTRACT"), None)
    if cn_title is not None:
        expect(cn_title, issues, "abstract.cn.title.text", cn_title.text in {"摘  要", "摘　要"}, "中文摘要标题两字之间应空一格。", "摘  要", cn_title.text)
        expect(cn_title, issues, "abstract.cn.title.style", cn_title.style == "36", "中文摘要标题应使用不编号章标题样式。", "36", cn_title.style)
        expect(cn_title, issues, "abstract.cn.title.align", cn_title.alignment == "center", "中文摘要标题应居中。", "center", cn_title.alignment)
        expect(cn_title, issues, "abstract.cn.title.font", cn_title.run_properties.get("font_eastAsia") == "宋体", "中文摘要标题应使用宋体。", "宋体", cn_title.run_properties.get("font_eastAsia"), severity="warning")
        expect(cn_title, issues, "abstract.cn.title.size", cn_title.run_properties.get("size") == "44", "中文摘要标题应为二号字。", "44", cn_title.run_properties.get("size"), severity="warning")
        expect(cn_title, issues, "abstract.cn.title.bold", cn_title.run_properties.get("bold") == "true", "中文摘要标题应加粗。", "bold=true", cn_title.run_properties.get("bold"), severity="warning")
        expect(cn_title, issues, "abstract.cn.page_break", cn_title.page_break_before, "中文摘要标题前应分页。", "pageBreakBefore", "missing")
    else:
        issues.append(ReviewIssue("warning", "abstract.cn.missing", "没有检测到中文摘要标题。"))

    if en_title is not None:
        expect(en_title, issues, "abstract.en.title.style", en_title.style == "36", "英文摘要标题应使用不编号章标题样式。", "36", en_title.style)
        expect(en_title, issues, "abstract.en.title.align", en_title.alignment == "center", "英文摘要标题应居中。", "center", en_title.alignment)
        expect(en_title, issues, "abstract.en.title.font", en_title.run_properties.get("font_ascii") == "Times New Roman", "英文摘要标题应使用 Times New Roman。", "Times New Roman", en_title.run_properties.get("font_ascii"), severity="warning")
        expect(en_title, issues, "abstract.en.title.size", en_title.run_properties.get("size") == "44", "英文摘要标题应为二号字。", "44", en_title.run_properties.get("size"), severity="warning")
        expect(en_title, issues, "abstract.en.title.bold", en_title.run_properties.get("bold") == "true", "英文摘要标题应加粗。", "bold=true", en_title.run_properties.get("bold"), severity="warning")
        expect(en_title, issues, "abstract.en.page_break", en_title.page_break_before, "英文摘要标题前应分页。", "pageBreakBefore", "missing")
    else:
        issues.append(ReviewIssue("warning", "abstract.en.missing", "没有检测到 ABSTRACT 标题。"))

    for paragraph in paragraphs:
        text = normalized_visible_text(paragraph.text)
        if is_chinese_keyword(text):
            expect(paragraph, issues, "abstract.cn.keywords.indent", is_zero_indentation(paragraph), "中文关键词段不应首行缩进。", "firstLine=0", paragraph.indentation.get("firstLine"))
            expect(paragraph, issues, "abstract.cn.keywords.separator", "，" in text or "," in text, "中文关键词之间应使用逗号分隔。", "关键词：A，B", text)
            expect_spacing(paragraph, issues, "abstract.cn.keywords.before", "beforeLines", "20", "中文关键词段前应为 0.2 行。", severity="warning")
        elif is_english_keyword(text):
            expect(paragraph, issues, "abstract.en.keywords.indent", is_zero_indentation(paragraph), "英文关键词段不应首行缩进。", "firstLine=0", paragraph.indentation.get("firstLine"))
            expect(paragraph, issues, "abstract.en.keywords.separator", ";" in text, "英文关键词之间应使用分号分隔。", "KEY WORDS: A; B", text)
            expect_spacing(paragraph, issues, "abstract.en.keywords.before", "beforeLines", "20", "英文关键词段前应为 0.2 行。", severity="warning")


def check_captions(paragraphs: tuple[ParagraphInfo, ...], issues: list[ReviewIssue]) -> None:
    for paragraph in paragraphs:
        text = normalized_visible_text(paragraph.text)
        if not re.match(r"^[图表]\d+-\d+", text):
            continue
        expect(paragraph, issues, "caption.style", paragraph.style == "8", "图表题注应使用题注样式。", "8", paragraph.style)
        expect(paragraph, issues, "caption.align", paragraph.alignment == "center", "图表题注应居中。", "center", paragraph.alignment)
        expect_spacing(paragraph, issues, "caption.line", "line", "400", "图表题注行距应为 20 磅。")
        expect_spacing(paragraph, issues, "caption.line_rule", "lineRule", "exact", "图表题注应使用固定行距。")
        expect(paragraph, issues, "caption.font.cn", paragraph.run_properties.get("font_eastAsia") == "宋体", "图表题注中文字体应为宋体。", "宋体", paragraph.run_properties.get("font_eastAsia"), severity="warning")
        expect(paragraph, issues, "caption.font.en", paragraph.run_properties.get("font_ascii") == "Times New Roman", "图表题注英文字体应为 Times New Roman。", "Times New Roman", paragraph.run_properties.get("font_ascii"), severity="warning")
        expect(paragraph, issues, "caption.font.size", paragraph.run_properties.get("size") == "21", "图表题注应为五号字。", "21", paragraph.run_properties.get("size"), severity="warning")
        expect(paragraph, issues, "caption.number_space", bool(re.match(r"^[图表]\d+-\d+\s{2,}\S", paragraph.text)), "图表编号和题名之间应空两格。", "图1-1  标题", paragraph.text)


def check_equations(paragraphs: tuple[ParagraphInfo, ...], issues: list[ReviewIssue]) -> None:
    for paragraph in paragraphs:
        text = normalized_visible_text(paragraph.text)
        math_text = text or paragraph.math_text
        if not re.search(r"\(\d+-\d+\)", math_text):
            continue
        expect(paragraph, issues, "equation.number.format", bool(re.search(r"\(\d+-\d+\)", math_text)), "公式编号应使用英文小括号。", "(3-1)", math_text)
        right_tabs = [tab for tab in paragraph.tab_stops if tab.get("val") == "right"]
        expect(paragraph, issues, "equation.number.right_tab", bool(right_tabs), "公式编号应通过右对齐制表位对齐到正文版心右侧。", "right tab stop", "missing")
        expect(paragraph, issues, "equation.number.no_leader", all(tab.get("leader") in {None, "none"} for tab in right_tabs), "公式编号右对齐不应使用引导符。", "leader=none", str(right_tabs))


def check_backmatter(
    paragraphs: tuple[ParagraphInfo, ...],
    styles: dict[str, dict[str, dict[str, str] | str | None]],
    issues: list[ReviewIssue],
) -> None:
    style_36 = styles.get("36", {})
    if style_36.get("based_on") is not None:
        issues.append(
            ReviewIssue(
                severity="warning",
                rule_id="backmatter.style36.based_on",
                message="不编号章标题样式仍继承其他样式，可能导致段前段后显示异常。",
                expected="basedOn 为空",
                actual=str(style_36.get("based_on")),
            )
        )

    if not any(normalized_visible_text(paragraph.text) == "参考文献" for paragraph in paragraphs):
        issues.append(ReviewIssue("warning", "backmatter.references.missing", "没有检测到参考文献标题。"))

    for paragraph in paragraphs:
        text = normalized_visible_text(paragraph.text)
        if text not in {"参考文献", "附 录", "附  录", "致 谢", "致  谢"}:
            continue
        expect(paragraph, issues, "backmatter.style", paragraph.style == "36", f"{text} 应使用不编号章标题样式。", "36", paragraph.style)
        expect(paragraph, issues, "backmatter.align", paragraph.alignment == "center", f"{text} 应居中。", "center", paragraph.alignment)
        expect(paragraph, issues, "backmatter.page_break", paragraph.page_break_before, f"{text} 前应分页。", "pageBreakBefore", "missing")
        expect_spacing(paragraph, issues, "backmatter.before", "before", "600", f"{text} 段前应为 30 磅。")
        expect_spacing(paragraph, issues, "backmatter.after", "after", "600", f"{text} 段后应为 30 磅。")
        expect_spacing(paragraph, issues, "backmatter.line", "line", "240", f"{text} 应为单倍行距。")
        expect_indent_zero(paragraph, issues, "backmatter.indent", f"{text} 不应首行缩进。")


def check_bibliography(paragraphs: tuple[ParagraphInfo, ...], issues: list[ReviewIssue]) -> None:
    entries = [paragraph for paragraph in paragraphs if re.match(r"^\[\d+\]\s+", normalized_visible_text(paragraph.text))]
    expected_number = 1
    for paragraph in entries:
        text = normalized_visible_text(paragraph.text)
        match = re.match(r"^\[(\d+)\]\s+", text)
        if match is None:
            continue
        number = int(match.group(1))
        expect(paragraph, issues, "bibliography.number.sequence", number == expected_number, "参考文献编号应连续。", f"[{expected_number}]", f"[{number}]")
        expected_number += 1
        expect(paragraph, issues, "bibliography.style", paragraph.style == "44", "参考文献条目应使用参考文献段落样式。", "44", paragraph.style)
        expect(paragraph, issues, "bibliography.indent", paragraph.indentation.get("left") == "420" and paragraph.indentation.get("hanging") == "420", "参考文献条目应使用悬挂缩进，避免编号后大空白。", "left=420 hanging=420", str(paragraph.indentation))
        if is_english_reference(text):
            expect(paragraph, issues, "bibliography.english.et_al", "等." not in text and "等。" not in text, "英文参考文献多人作者应使用 et al.。", "et al.", text, severity="warning")
            expect(paragraph, issues, "bibliography.english.case", not has_all_caps_author_block(text), "英文作者名不应整体强制大写。", "Yan H, Ding G", text, severity="warning")


def write_review_report(report: ReviewReport) -> None:
    report.markdown_path.write_text(render_markdown_report(report), encoding="utf-8")
    report.json_path.write_text(
        json.dumps(
            {
                "docx_path": str(report.docx_path),
                "error_count": report.error_count,
                "warning_count": report.warning_count,
                "issues": [asdict(issue) for issue in report.issues],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )


def render_markdown_report(report: ReviewReport) -> str:
    lines = [
        "# DOCX 格式检查报告",
        "",
        f"- 文件：`{report.docx_path}`",
        f"- 严重问题：{report.error_count}",
        f"- 警告：{report.warning_count}",
        "",
    ]
    if not report.issues:
        lines.append("未发现第一版规则覆盖范围内的格式问题。")
        lines.append("")
        return "\n".join(lines)

    lines.append("## 问题列表")
    lines.append("")
    for issue in report.issues:
        location = f"第 {issue.paragraph_index} 段" if issue.paragraph_index is not None else "文档级"
        lines.append(f"- **{issue.severity.upper()}** `{issue.rule_id}`：{issue.message}")
        lines.append(f"  - 位置：{location}")
        if issue.text:
            lines.append(f"  - 文本：{issue.text}")
        if issue.expected is not None or issue.actual is not None:
            lines.append(f"  - 期望：{issue.expected or ''}")
            lines.append(f"  - 实际：{issue.actual or ''}")
    lines.append("")
    return "\n".join(lines)


def expect(
    paragraph: ParagraphInfo,
    issues: list[ReviewIssue],
    rule_id: str,
    condition: bool,
    message: str,
    expected: str | None,
    actual: str | None,
    severity: str = "error",
) -> None:
    if condition:
        return
    issues.append(
        ReviewIssue(
            severity=severity,
            rule_id=rule_id,
            message=message,
            paragraph_index=paragraph.index,
            text=paragraph.text,
            expected=expected,
            actual=actual,
        )
    )


def expect_spacing(
    paragraph: ParagraphInfo,
    issues: list[ReviewIssue],
    rule_id: str,
    attr: str,
    value: str,
    message: str,
    severity: str = "error",
) -> None:
    expect(paragraph, issues, rule_id, paragraph.spacing.get(attr) == value, message, f"{attr}={value}", f"{attr}={paragraph.spacing.get(attr)}", severity=severity)


def expect_indent_zero(paragraph: ParagraphInfo, issues: list[ReviewIssue], rule_id: str, message: str) -> None:
    expect(paragraph, issues, rule_id, is_zero_indentation(paragraph), message, "firstLine=0 或无首行缩进", str(paragraph.indentation))


def child_attr(parent: ET.Element | None, child_tag: str, attr: str) -> str | None:
    if parent is None:
        return None
    child = parent.find(f"w:{child_tag}", NS)
    if child is None:
        return None
    return child.get(q(attr))


def child_attrs(parent: ET.Element | None, child_tag: str) -> dict[str, str]:
    if parent is None:
        return {}
    child = parent.find(f"w:{child_tag}", NS)
    if child is None:
        return {}
    return local_attrs(child)


def run_properties(rpr: ET.Element | None) -> dict[str, str]:
    if rpr is None:
        return {}
    props: dict[str, str] = {}
    fonts = rpr.find("w:rFonts", NS)
    if fonts is not None:
        props.update({f"font_{key}": value for key, value in local_attrs(fonts).items()})
    size = rpr.find("w:sz", NS)
    if size is not None and size.get(q("val")) is not None:
        props["size"] = size.get(q("val"), "")
    if rpr.find("w:b", NS) is not None:
        props["bold"] = "true"
    return props


def tab_attrs(tab: ET.Element) -> dict[str, str]:
    return local_attrs(tab)


def local_attrs(element: ET.Element) -> dict[str, str]:
    return {
        key.rsplit("}", 1)[-1]: value
        for key, value in element.attrib.items()
        if key.startswith(f"{{{W_NS}}}")
    }


def normalized_visible_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\u00a0", " ")).strip()


def is_chapter_heading(text: str) -> bool:
    return bool(re.match(r"^第[一二三四五六七八九十百\d]+章\b", text))


def is_section_heading(text: str) -> bool:
    return bool(re.match(r"^\d+\.\d+\b", text)) and not is_subsection_heading(text)


def is_subsection_heading(text: str) -> bool:
    return bool(re.match(r"^\d+\.\d+\.\d+\b", text))


def has_heading_number_gap(text: str, level: int) -> bool:
    text = text.replace("\u00a0", " ")
    if level == 1:
        return bool(re.match(r"^第[一二三四五六七八九十百\d]+章(?: {2}|\u3000)\S", text))
    return bool(re.match(r"^\d+(?:\.\d+){1,2}(?: {2}|\u3000)\S", text))


def is_chinese_keyword(text: str) -> bool:
    return bool(re.match(r"^(关键词|关键字)\s*[:：]", text))


def is_english_keyword(text: str) -> bool:
    return bool(re.match(r"^KEY\s+WORDS?\s*:", text, flags=re.IGNORECASE))


def is_zero_indentation(paragraph: ParagraphInfo) -> bool:
    return paragraph.indentation.get("firstLine") in {None, "0"} and paragraph.indentation.get("firstLineChars") in {None, "0"}


def is_english_reference(text: str) -> bool:
    match = re.match(r"^\[\d+\]\s*(.+?)\.\s+", text)
    if match is None:
        return False
    authors = match.group(1).replace("et al", "").replace("等", "")
    return bool(re.search(r"[A-Za-z]", authors)) and not re.search(r"[\u4e00-\u9fff]", authors)


def has_all_caps_author_block(text: str) -> bool:
    match = re.match(r"^\[\d+\]\s*(.+?)\.\s+", text)
    if match is None:
        return False
    authors = re.sub(r"\bet\s+al\b", "", match.group(1), flags=re.IGNORECASE)
    letters = re.findall(r"[A-Za-z]", authors)
    if len(letters) < 6:
        return False
    uppercase = sum(1 for letter in letters if letter.isupper())
    return uppercase / len(letters) > 0.85
