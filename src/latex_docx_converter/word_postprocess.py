from __future__ import annotations

from dataclasses import dataclass
import copy
from pathlib import Path
import re
import shutil
import tempfile
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET

from .citation import CitationAudit


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
NS = {"w": W_NS, "r": R_NS, "rel": PKG_REL_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)


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
    del audit_result
    warnings: list[str] = []
    notes: list[str] = []

    with tempfile.TemporaryDirectory(prefix="latex-docx-postprocess-") as tmp:
        temp_docx = Path(tmp) / output_docx.name
        with ZipFile(output_docx, "r") as source:
            document_xml = source.read("word/document.xml")
            settings_xml = source.read("word/settings.xml") if "word/settings.xml" in source.namelist() else None
            styles_xml = source.read("word/styles.xml") if "word/styles.xml" in source.namelist() else None
            rels_xml = (
                source.read("word/_rels/document.xml.rels")
                if "word/_rels/document.xml.rels" in source.namelist()
                else None
            )

            document_root = ET.fromstring(document_xml)
            package_additions: dict[str, bytes] = {}
            if profile.reference_docx is not None and profile.reference_docx.is_file():
                frontmatter = load_reference_frontmatter(profile.reference_docx, rels_xml)
                if frontmatter is not None:
                    insert_reference_frontmatter(document_root, frontmatter.elements)
                    rels_xml = frontmatter.rels_xml
                    package_additions.update(frontmatter.package_additions)
                    notes.append("Copied first two pages from reference DOCX.")
                else:
                    warnings.append("Could not detect the first two template pages in reference DOCX.")
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
                    elif item.filename == "word/_rels/document.xml.rels" and rels_xml is not None:
                        target.writestr(item, rels_xml)
                    else:
                        target.writestr(item, source.read(item.filename))
                if "word/settings.xml" not in source.namelist():
                    target.writestr("word/settings.xml", new_settings_xml)
                if "word/_rels/document.xml.rels" not in source.namelist() and rels_xml is not None:
                    target.writestr("word/_rels/document.xml.rels", rels_xml)
                for name, data in package_additions.items():
                    target.writestr(name, data)

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


@dataclass(frozen=True)
class FrontmatterCopy:
    elements: tuple[ET.Element, ...]
    rels_xml: bytes
    package_additions: dict[str, bytes]


def load_reference_frontmatter(reference_docx: Path, target_rels_xml: bytes | None) -> FrontmatterCopy | None:
    with ZipFile(reference_docx, "r") as reference:
        document_root = ET.fromstring(reference.read("word/document.xml"))
        body = document_root.find("w:body", NS)
        if body is None:
            return None

        elements = first_two_section_elements(body)
        if not elements:
            return None

        target_rels_root = parse_relationships(target_rels_xml)
        reference_rels_root = parse_relationships(
            reference.read("word/_rels/document.xml.rels")
            if "word/_rels/document.xml.rels" in reference.namelist()
            else None
        )
        package_additions: dict[str, bytes] = {}
        mapped_elements = [copy.deepcopy(element) for element in elements]
        remove_header_footer_references(mapped_elements)
        used_relationship_ids = collect_relationship_ids(mapped_elements)
        remap_relationships(
            mapped_elements,
            used_relationship_ids,
            reference_rels_root,
            target_rels_root,
            reference,
            package_additions,
        )
        return FrontmatterCopy(
            elements=tuple(mapped_elements),
            rels_xml=xml_bytes(target_rels_root),
            package_additions=package_additions,
        )


def first_two_section_elements(body: ET.Element) -> list[ET.Element]:
    elements: list[ET.Element] = []
    section_count = 0
    for child in list(body):
        elements.append(child)
        if contains_section_properties(child):
            section_count += 1
            if section_count >= 2:
                return elements
    return []


def contains_section_properties(element: ET.Element) -> bool:
    return element.tag == q("sectPr") or element.find(".//w:sectPr", NS) is not None


def insert_reference_frontmatter(root: ET.Element, elements: tuple[ET.Element, ...]) -> None:
    body = root.find("w:body", NS)
    if body is None:
        return
    remove_generated_frontmatter_placeholder(body)
    insert_at = 0
    for offset, element in enumerate(elements):
        body.insert(insert_at + offset, element)


def remove_generated_frontmatter_placeholder(body: ET.Element) -> None:
    remove_named_bookmarks(body, "封面与独创性声明")
    children = list(body)
    start = None
    for index, child in enumerate(children):
        if normalized_text(element_text(child)) == "封面与独创性声明":
            start = index
            break
    if start is None:
        return
    end = start + 1
    while end < len(children):
        text = normalized_text(element_text(children[end]))
        if text in {"摘 要", "摘要", "ABSTRACT", "目 录", "目录"}:
            break
        end += 1
    for child in children[start:end]:
        body.remove(child)


def remove_named_bookmarks(body: ET.Element, name: str) -> None:
    bookmark_ids: set[str] = set()
    for child in list(body):
        if child.tag == q("bookmarkStart") and child.get(q("name")) == name and child.get(q("id")):
            bookmark_ids.add(child.get(q("id")))
            body.remove(child)
            continue
        for bookmark in child.findall(".//w:bookmarkStart", NS):
            if bookmark.get(q("name")) == name and bookmark.get(q("id")):
                bookmark_ids.add(bookmark.get(q("id")))
                remove_descendant(child, bookmark)
    for child in list(body):
        if child.tag == q("bookmarkEnd") and child.get(q("id")) in bookmark_ids:
            body.remove(child)
            continue
        for bookmark in child.findall(".//w:bookmarkEnd", NS):
            if bookmark.get(q("id")) in bookmark_ids:
                remove_descendant(child, bookmark)


def remove_descendant(root: ET.Element, target: ET.Element) -> bool:
    for parent in root.iter():
        for child in list(parent):
            if child is target:
                parent.remove(child)
                return True
    return False


def parse_relationships(rels_xml: bytes | None) -> ET.Element:
    if rels_xml:
        return ET.fromstring(rels_xml)
    return ET.Element(f"{{{PKG_REL_NS}}}Relationships")


def collect_relationship_ids(elements: list[ET.Element]) -> set[str]:
    ids: set[str] = set()
    for element in elements:
        for node in element.iter():
            for key, value in node.attrib.items():
                if key in {rq("id"), rq("embed"), rq("link")}:
                    ids.add(value)
    return ids


def remove_header_footer_references(elements: list[ET.Element]) -> None:
    for element in elements:
        for sect_pr in element.findall(".//w:sectPr", NS):
            for child in list(sect_pr):
                if child.tag in {q("headerReference"), q("footerReference")}:
                    sect_pr.remove(child)


def remap_relationships(
    elements: list[ET.Element],
    used_relationship_ids: set[str],
    reference_rels_root: ET.Element,
    target_rels_root: ET.Element,
    reference_zip: ZipFile,
    package_additions: dict[str, bytes],
) -> None:
    reference_relationships = {
        relationship.get("Id"): relationship
        for relationship in reference_rels_root.findall("rel:Relationship", NS)
        if relationship.get("Id")
    }
    for old_id in sorted(used_relationship_ids):
        relationship = reference_relationships.get(old_id)
        if relationship is None:
            continue
        new_id = next_relationship_id(target_rels_root)
        new_target = copy_relationship_target(relationship, new_id, reference_zip, package_additions)
        ET.SubElement(
            target_rels_root,
            f"{{{PKG_REL_NS}}}Relationship",
            {
                "Id": new_id,
                "Type": relationship.get("Type", ""),
                "Target": new_target,
                **({"TargetMode": relationship.get("TargetMode")} if relationship.get("TargetMode") else {}),
            },
        )
        replace_relationship_id(elements, old_id, new_id)


def copy_relationship_target(
    relationship: ET.Element,
    new_id: str,
    reference_zip: ZipFile,
    package_additions: dict[str, bytes],
) -> str:
    target = relationship.get("Target", "")
    if relationship.get("TargetMode") == "External" or not target:
        return target
    source_name = f"word/{target}" if not target.startswith("/") else target.lstrip("/")
    suffix = Path(target).suffix
    if target.startswith("media/"):
        new_target = f"media/{new_id}{suffix}"
    else:
        target_path = Path(target)
        new_target = target_path.with_name(f"{target_path.stem}-{new_id}{suffix}").as_posix()
    new_name = f"word/{new_target}"
    if source_name in reference_zip.namelist():
        package_additions[new_name] = reference_zip.read(source_name)
        return new_target
    return target


def replace_relationship_id(elements: list[ET.Element], old_id: str, new_id: str) -> None:
    for element in elements:
        for node in element.iter():
            for key, value in list(node.attrib.items()):
                if key in {rq("id"), rq("embed"), rq("link")} and value == old_id:
                    node.set(key, new_id)


def next_relationship_id(root: ET.Element) -> str:
    used = {
        relationship.get("Id")
        for relationship in root.findall("rel:Relationship", NS)
        if relationship.get("Id")
    }
    index = 1
    while f"rIdFront{index}" in used:
        index += 1
    return f"rIdFront{index}"


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
    caption_style = root.find("w:style[@w:styleId='8']", NS)
    if caption_style is not None:
        apply_caption_paragraph_format(caption_style)
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
            apply_caption_paragraph_format(child)
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
    if "  " in text:
        text_node.set(XML_SPACE, "preserve")
    text_node.text = text
    return p


def replace_paragraph_text(paragraph: ET.Element, text: str) -> None:
    for run in paragraph.findall("w:r", NS):
        paragraph.remove(run)
    run = ET.SubElement(paragraph, q("r"))
    text_node = ET.SubElement(run, q("t"))
    if "  " in text:
        text_node.set(XML_SPACE, "preserve")
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


def apply_caption_paragraph_format(paragraph: ET.Element) -> None:
    normalize_caption_text(paragraph)
    set_paragraph_alignment(paragraph, "center")
    set_paragraph_indentation(paragraph, left="0", first_line="0", first_line_chars="0")
    set_paragraph_spacing(paragraph, before="120", after="120", line="400", line_rule="exact")
    set_run_format(paragraph, east_asia_font="宋体", ascii_font="Times New Roman", size="21")


def normalize_caption_text(paragraph: ET.Element) -> None:
    text = element_text(paragraph)
    normalized = re.sub(r"^([图表]\d+-\d+)\s+", r"\1  ", text)
    if re.match(r"^[图表]\d+-\d+", normalized):
        replace_paragraph_text(paragraph, normalized)


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


def set_run_format(element: ET.Element, east_asia_font: str, ascii_font: str, size: str) -> None:
    if element.tag == q("style"):
        set_rpr_format(ensure_rpr(element), east_asia_font, ascii_font, size)
        return

    runs = element.findall("w:r", NS)
    if not runs:
        runs = [ET.SubElement(element, q("r"))]
    for run in runs:
        rpr = run.find("w:rPr", NS)
        if rpr is None:
            rpr = ET.Element(q("rPr"))
            run.insert(0, rpr)
        set_rpr_format(rpr, east_asia_font, ascii_font, size)


def ensure_rpr(element: ET.Element) -> ET.Element:
    rpr = element.find("w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(q("rPr"))
        element.append(rpr)
    return rpr


def set_rpr_format(rpr: ET.Element, east_asia_font: str, ascii_font: str, size: str) -> None:
    fonts = rpr.find("w:rFonts", NS)
    if fonts is None:
        fonts = ET.Element(q("rFonts"))
        rpr.insert(0, fonts)
    fonts.set(q("eastAsia"), east_asia_font)
    fonts.set(q("ascii"), ascii_font)
    fonts.set(q("hAnsi"), ascii_font)
    fonts.set(q("cs"), ascii_font)
    sz = rpr.find("w:sz", NS)
    if sz is None:
        sz = ET.Element(q("sz"))
        rpr.append(sz)
    sz.set(q("val"), size)
    sz_cs = rpr.find("w:szCs", NS)
    if sz_cs is None:
        sz_cs = ET.Element(q("szCs"))
        rpr.append(sz_cs)
    sz_cs.set(q("val"), size)


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


def rq(tag: str) -> str:
    return f"{{{R_NS}}}{tag}"


def xml_bytes(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
