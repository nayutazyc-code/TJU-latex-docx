from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import shutil
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

from .citation import CitationAudit


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
NS = {"w": W_NS}

ET.register_namespace("w", W_NS)


@dataclass(frozen=True)
class WordPostprocessProfile:
    reference_docx: Path | None = None


@dataclass(frozen=True)
class PostprocessResult:
    warnings: tuple[str, ...] = ()
    notes: tuple[str, ...] = ()
    toc_inserted: bool = False
    bibliography_moved: bool = False


def postprocess_docx(
    output_docx: Path,
    profile: WordPostprocessProfile,
    audit_result: CitationAudit | None = None,
) -> PostprocessResult:
    del profile
    del audit_result
    warnings: list[str] = []
    notes: list[str] = []

    with tempfile.TemporaryDirectory(prefix="latex-docx-postprocess-") as tmp:
        temp_docx = Path(tmp) / output_docx.name
        with ZipFile(output_docx, "r") as source:
            document_xml = source.read("word/document.xml")
            settings_xml = source.read("word/settings.xml") if "word/settings.xml" in source.namelist() else None
            styles_xml = source.read("word/styles.xml") if "word/styles.xml" in source.namelist() else None

            document_root = ET.fromstring(document_xml)
            toc_inserted, bibliography_moved = process_document_xml(document_root)
            new_document_xml = xml_bytes(document_root)

            new_settings_xml = ensure_update_fields(settings_xml)
            new_styles_xml = process_styles_xml(styles_xml) if styles_xml is not None else None

            with ZipFile(temp_docx, "w", ZIP_DEFLATED) as target:
                for item in source.infolist():
                    if item.filename == "word/document.xml":
                        target.writestr(item, new_document_xml)
                    elif item.filename == "word/settings.xml":
                        target.writestr(item, new_settings_xml)
                    elif item.filename == "word/styles.xml" and new_styles_xml is not None:
                        target.writestr(item, new_styles_xml)
                    else:
                        target.writestr(item, source.read(item.filename))
                if "word/settings.xml" not in source.namelist():
                    target.writestr("word/settings.xml", new_settings_xml)

        shutil.move(str(temp_docx), output_docx)

    if toc_inserted:
        notes.append("Inserted Word TOC field at 目录.")
    else:
        warnings.append("TOC placeholder/title was not found; Word TOC field was not inserted.")

    if bibliography_moved:
        notes.append("Moved generated bibliography entries under 参考文献.")
    else:
        warnings.append("Generated bibliography entries were not detected for repositioning.")

    return PostprocessResult(
        warnings=tuple(warnings),
        notes=tuple(notes),
        toc_inserted=toc_inserted,
        bibliography_moved=bibliography_moved,
    )


def process_document_xml(root: ET.Element) -> tuple[bool, bool]:
    body = root.find("w:body", NS)
    if body is None:
        return False, False

    remove_table_of_contents_heading(body)
    bibliography_moved = move_bibliography_entries(body)
    apply_tju_styles(body)
    toc_inserted = insert_toc_field(body)
    remove_marker_paragraphs(body)
    return toc_inserted, bibliography_moved


def process_styles_xml(styles_xml: bytes) -> bytes:
    root = ET.fromstring(styles_xml)
    for style_id in ("2", "3", "4", "37", "38", "39"):
        style = root.find(f"w:style[@w:styleId='{style_id}']", NS)
        if style is None:
            continue
        ppr = style.find("w:pPr", NS)
        if ppr is not None:
            remove_child(ppr, "numPr")
    reference_style = root.find("w:style[@w:styleId='44']", NS)
    if reference_style is not None:
        apply_reference_paragraph_format(reference_style)
    return xml_bytes(root)


def remove_table_of_contents_heading(body: ET.Element) -> None:
    for child in list(body):
        if strip_text(element_text(child)).lower() == "table of contents":
            body.remove(child)


def move_bibliography_entries(body: ET.Element) -> bool:
    children = list(body)
    bib_entries = [child for child in children if is_paragraph(child) and is_bibliography_entry(element_text(child))]
    if not bib_entries:
        return False

    for entry in bib_entries:
        body.remove(entry)
        set_paragraph_style(entry, "44")
        apply_reference_paragraph_format(entry)

    children = list(body)
    insert_at = None
    for index, child in enumerate(children):
        if normalized_text(element_text(child)) == "参考文献":
            insert_at = index + 1
            break
    if insert_at is None:
        sect_pr = find_section_properties(body)
        insert_at = len(children) - (1 if sect_pr is not None else 0)
        body.insert(insert_at, make_text_paragraph("参考文献", "36"))
        insert_at += 1

    for offset, entry in enumerate(bib_entries):
        body.insert(insert_at + offset, entry)
    return True


def apply_tju_styles(body: ET.Element) -> None:
    children = list(body)
    for index, child in enumerate(children):
        if not is_paragraph(child):
            continue
        text = normalized_text(element_text(child))
        if not text:
            continue
        style = current_style(child)
        previous = children[index - 1] if index > 0 else None
        next_element = children[index + 1] if index + 1 < len(children) else None
        if text in {
            "封面与独创性声明",
            "独创性声明",
            "摘 要",
            "摘要",
            "ABSTRACT",
            "目 录",
            "目录",
            "参考文献",
            "附 录",
            "附录",
            "致 谢",
            "致谢",
        }:
            if text in {"目 录", "目录"}:
                replace_paragraph_text(child, "目  录")
            if text in {"摘 要", "摘要"}:
                replace_paragraph_text(child, "摘  要")
            if text in {"附 录", "附录"}:
                replace_paragraph_text(child, "附  录")
            if text in {"致 谢", "致谢"}:
                replace_paragraph_text(child, "致  谢")
            set_paragraph_style(child, "36", outline_level=0, clear_numbering=True)
        elif is_caption_like_paragraph(text, child, previous, next_element):
            set_paragraph_style(child, "8")
            set_paragraph_alignment(child, "center")
        elif style == "2" or re.match(r"^第[一二三四五六七八九十百\d]+章\b", text):
            set_paragraph_style(child, "2", outline_level=0, clear_numbering=True)
            apply_heading_paragraph_format(child, 1)
        elif style == "3":
            set_paragraph_style(child, "3", outline_level=1, clear_numbering=True)
            apply_heading_paragraph_format(child, 2)
        elif style == "4":
            set_paragraph_style(child, "4", outline_level=2, clear_numbering=True)
            apply_heading_paragraph_format(child, 3)
        elif style == "38":
            set_paragraph_style(child, "3", outline_level=1, clear_numbering=True)
            apply_heading_paragraph_format(child, 2)
        elif style == "39":
            set_paragraph_style(child, "4", outline_level=2, clear_numbering=True)
            apply_heading_paragraph_format(child, 3)
        elif re.match(r"^第[一二三四五六七八九十百\d]+章\b", text):
            set_paragraph_style(child, "2", outline_level=0, clear_numbering=True)
            apply_heading_paragraph_format(child, 1)
        elif is_bibliography_entry(text):
            set_paragraph_style(child, "44")
            apply_reference_paragraph_format(child)
        elif is_body_like_paragraph(text, child):
            set_paragraph_style(child, "40")


def insert_toc_field(body: ET.Element) -> bool:
    children = list(body)
    for index, child in enumerate(children):
        text = normalized_text(element_text(child))
        if text not in {"目 录", "目录"}:
            continue

        remove_static_toc_following(body, index + 1)
        body.insert(index + 1, make_toc_field_paragraph())
        return True

    for index, child in enumerate(children):
        if normalized_text(element_text(child)) == "TJU_DOCX_TOC_PLACEHOLDER":
            body.remove(child)
            body.insert(index, make_text_paragraph("目  录", "36"))
            body.insert(index + 1, make_toc_field_paragraph())
            return True
    return False


def remove_static_toc_following(body: ET.Element, start: int) -> None:
    while True:
        children = list(body)
        if start >= len(children):
            return
        child = children[start]
        text = normalized_text(element_text(child))
        if not text or text == "TJU_DOCX_TOC_PLACEHOLDER":
            body.remove(child)
            continue
        if re.match(r"^第[一二三四五六七八九十百\d]+章\b", text):
            return
        if len(text) > 30 and ("第一章" in text or "第二章" in text):
            body.remove(child)
            continue
        return


def remove_marker_paragraphs(body: ET.Element) -> None:
    for child in list(body):
        text = normalized_text(element_text(child))
        if text in {"TJU_DOCX_TOC_PLACEHOLDER", "TJU_DOCX_BIB_PLACEHOLDER"}:
            body.remove(child)


def ensure_update_fields(settings_xml: bytes | None) -> bytes:
    if settings_xml:
        root = ET.fromstring(settings_xml)
    else:
        root = ET.Element(q("settings"))
    existing = root.find("w:updateFields", NS)
    if existing is None:
        existing = ET.SubElement(root, q("updateFields"))
    existing.set(q("val"), "true")
    return xml_bytes(root)


def make_toc_field_paragraph() -> ET.Element:
    p = ET.Element(q("p"))
    add_run_with_field_char(p, "begin")
    run = ET.SubElement(p, q("r"))
    instr = ET.SubElement(run, q("instrText"))
    instr.set(XML_SPACE, "preserve")
    instr.text = ' TOC \\o "1-3" \\h \\u '
    add_run_with_field_char(p, "separate")
    run = ET.SubElement(p, q("r"))
    text = ET.SubElement(run, q("t"))
    text.text = "请在 Word 中右键更新目录。"
    add_run_with_field_char(p, "end")
    return p


def add_run_with_field_char(paragraph: ET.Element, field_type: str) -> None:
    run = ET.SubElement(paragraph, q("r"))
    fld = ET.SubElement(run, q("fldChar"))
    fld.set(q("fldCharType"), field_type)


def make_text_paragraph(text: str, style_id: str | None = None) -> ET.Element:
    p = ET.Element(q("p"))
    if style_id:
        set_paragraph_style(p, style_id)
    run = ET.SubElement(p, q("r"))
    text_node = ET.SubElement(run, q("t"))
    text_node.text = text
    return p


def replace_paragraph_text(paragraph: ET.Element, text: str) -> None:
    for run in paragraph.findall("w:r", NS):
        paragraph.remove(run)
    run = ET.SubElement(paragraph, q("r"))
    text_node = ET.SubElement(run, q("t"))
    text_node.text = text


def set_paragraph_style(
    paragraph: ET.Element,
    style_id: str,
    outline_level: int | None = None,
    clear_numbering: bool = False,
) -> None:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(q("pPr"))
        paragraph.insert(0, ppr)
    if clear_numbering:
        remove_child(ppr, "numPr")
    pstyle = ppr.find("w:pStyle", NS)
    if pstyle is None:
        pstyle = ET.Element(q("pStyle"))
        ppr.insert(0, pstyle)
    pstyle.set(q("val"), style_id)
    if outline_level is not None:
        outline = ppr.find("w:outlineLvl", NS)
        if outline is None:
            outline = ET.Element(q("outlineLvl"))
            ppr.append(outline)
        outline.set(q("val"), str(outline_level))


def apply_heading_paragraph_format(paragraph: ET.Element, level: int) -> None:
    ppr = ensure_ppr(paragraph)
    remove_child(ppr, "numPr")
    if level == 1:
        set_paragraph_alignment(paragraph, "center")
        set_paragraph_spacing(paragraph, before="600", after="600")
        set_paragraph_indentation(paragraph, left="0", first_line="0")
        set_page_break_before(paragraph)
    elif level == 2:
        set_paragraph_alignment(paragraph, "left")
        set_paragraph_spacing(paragraph, before="360", after="360")
        set_paragraph_indentation(paragraph, left="0", first_line="0", first_line_chars="0")
    elif level == 3:
        set_paragraph_alignment(paragraph, "left")
        set_paragraph_spacing(paragraph, before="240", after="240")
        set_paragraph_indentation(paragraph, left="0", first_line="0", first_line_chars="0")


def apply_reference_paragraph_format(element: ET.Element) -> None:
    normalize_bibliography_text(element)
    set_paragraph_indentation(element, left="420", hanging="420", first_line_chars="0")
    set_paragraph_spacing(element, before_lines="0", line="400", line_rule="exact")
    ind = ensure_ppr(element).find("w:ind", NS)
    if ind is not None and q("firstLine") in ind.attrib:
        del ind.attrib[q("firstLine")]


def normalize_bibliography_text(element: ET.Element) -> None:
    if not is_paragraph(element):
        return
    text = element_text(element)
    if not text:
        return
    text = re.sub(r"\s*\t\s*", " ", text)
    text = re.sub(r"^(\[\d+\])\s+", r"\1 ", text)
    ppr = element.find("w:pPr", NS)
    for child in list(element):
        if child is not ppr:
            element.remove(child)
    run = ET.SubElement(element, q("r"))
    text_node = ET.SubElement(run, q("t"))
    text_node.text = text


def set_paragraph_alignment(paragraph: ET.Element, value: str) -> None:
    ppr = ensure_ppr(paragraph)
    alignment = ppr.find("w:jc", NS)
    if alignment is None:
        alignment = ET.Element(q("jc"))
        ppr.append(alignment)
    alignment.set(q("val"), value)


def set_paragraph_spacing(
    paragraph: ET.Element,
    before: str | None = None,
    after: str | None = None,
    before_lines: str | None = None,
    line: str | None = None,
    line_rule: str | None = None,
) -> None:
    ppr = ensure_ppr(paragraph)
    spacing = ppr.find("w:spacing", NS)
    if spacing is None:
        spacing = ET.Element(q("spacing"))
        ppr.append(spacing)
    for attr, value in (
        ("before", before),
        ("after", after),
        ("beforeLines", before_lines),
        ("line", line),
        ("lineRule", line_rule),
    ):
        if value is not None:
            spacing.set(q(attr), value)


def set_paragraph_indentation(
    paragraph: ET.Element,
    left: str | None = None,
    first_line: str | None = None,
    first_line_chars: str | None = None,
    hanging: str | None = None,
) -> None:
    ppr = ensure_ppr(paragraph)
    ind = ppr.find("w:ind", NS)
    if ind is None:
        ind = ET.Element(q("ind"))
        ppr.append(ind)
    for attr, value in (
        ("left", left),
        ("firstLine", first_line),
        ("firstLineChars", first_line_chars),
        ("hanging", hanging),
    ):
        if value is not None:
            ind.set(q(attr), value)


def set_page_break_before(paragraph: ET.Element) -> None:
    ppr = ensure_ppr(paragraph)
    if ppr.find("w:pageBreakBefore", NS) is None:
        ppr.append(ET.Element(q("pageBreakBefore")))


def ensure_ppr(element: ET.Element) -> ET.Element:
    ppr = element.find("w:pPr", NS)
    if ppr is None:
        ppr = ET.Element(q("pPr"))
        element.insert(0, ppr)
    return ppr


def remove_child(parent: ET.Element, tag: str) -> None:
    child = parent.find(f"w:{tag}", NS)
    if child is not None:
        parent.remove(child)


def current_style(paragraph: ET.Element) -> str | None:
    pstyle = paragraph.find("w:pPr/w:pStyle", NS)
    return pstyle.get(q("val")) if pstyle is not None else None


def is_body_like_paragraph(text: str, paragraph: ET.Element) -> bool:
    if text.startswith("请在 Word 中"):
        return False
    if current_style(paragraph) in {"8", "CaptionedFigure", "ImageCaption", "TableCaption", "Compact"}:
        return False
    return bool(re.search(r"[\u4e00-\u9fffA-Za-z]", text))


def is_caption_like_paragraph(
    text: str,
    paragraph: ET.Element,
    previous: ET.Element | None,
    next_element: ET.Element | None,
) -> bool:
    if current_style(paragraph) in {"8", "CaptionedFigure", "ImageCaption", "TableCaption"}:
        return True
    if is_bibliography_entry(text) or len(text) > 90:
        return False
    if re.match(r"^(图|表)\s*[\d一二三四五六七八九十]+(?:[.-]\d+)?", text):
        return True
    if previous is not None and contains_visual(previous):
        return bool(re.search(r"[\u4e00-\u9fffA-Za-z]", text))
    if next_element is not None and next_element.tag == q("tbl"):
        return bool(re.search(r"[\u4e00-\u9fffA-Za-z]", text))
    return False


def contains_visual(element: ET.Element) -> bool:
    return element.find(".//w:drawing", NS) is not None or element.find(".//w:pict", NS) is not None


def is_bibliography_entry(text: str) -> bool:
    return bool(re.match(r"^\[\d+\]\s+", strip_text(text)))


def is_paragraph(element: ET.Element) -> bool:
    return element.tag == q("p")


def element_text(element: ET.Element) -> str:
    return "".join(text.text or "" for text in element.findall(".//w:t", NS))


def normalized_text(text: str) -> str:
    return re.sub(r"\s+", " ", strip_text(text)).strip()


def strip_text(text: str) -> str:
    return text.replace("\u00a0", " ").strip()


def find_section_properties(body: ET.Element) -> ET.Element | None:
    for child in reversed(list(body)):
        if child.tag == q("sectPr"):
            return child
    return None


def q(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
