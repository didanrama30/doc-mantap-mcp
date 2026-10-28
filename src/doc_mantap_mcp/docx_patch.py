"""Conservative OOXML patching that leaves unrelated DOCX package parts untouched."""

from __future__ import annotations

import os
import shutil
import tempfile
import uuid
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree
from docx.oxml.ns import qn

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "http://www.w3.org/XML/1998/namespace"


def _text(paragraph) -> str:
    return "".join(node.text or "" for node in paragraph.iter(qn("w:t")))


def _all_paragraphs(body) -> list:
    return list(body.iter(qn("w:p")))


def _tables(body) -> list:
    return [e for e in body if e.tag == qn("w:tbl")]


def _resolve_paragraph(body, anchor: str):
    paragraphs = _all_paragraphs(body)
    if anchor.startswith("p:") and anchor[2:].isdigit():
        index = int(anchor[2:])
        if index >= len(paragraphs):
            raise ValueError(f"Paragraph anchor out of range: {anchor}")
        return paragraphs[index]
    if anchor.startswith("bookmark:"):
        name = anchor.split(":", 1)[1]
        for p in paragraphs:
            if any(b.get(qn("w:name")) == name for b in p.iter(qn("w:bookmarkStart"))):
                return p
        raise ValueError(f"Bookmark not found: {name}")
    matches = [p for p in paragraphs if anchor in _text(p)]
    if not matches:
        raise ValueError(f"Anchor not found: {anchor!r}")
    if len(matches) > 1:
        raise ValueError(f"Anchor is ambiguous ({len(matches)} paragraphs): {anchor!r}; use p:index")
    return matches[0]


def _style_id(styles_root, name: str) -> str:
    for style in styles_root.iter(qn("w:style")):
        style_name = style.find(qn("w:name"))
        if (style.get(qn("w:type")) == "paragraph" and style_name is not None
                and style_name.get(qn("w:val")) == name):
            return style.get(qn("w:styleId"))
    raise ValueError(f"Paragraph style not found: {name!r}")


def _code_style(styles_root) -> tuple[str | None, str]:
    for style in styles_root.iter(qn("w:style")):
        name_el = style.find(qn("w:name"))
        name = name_el.get(qn("w:val"), "") if name_el is not None else ""
        rpr = style.find(qn("w:rPr"))
        fonts = rpr.find(qn("w:rFonts")) if rpr is not None else None
        font = fonts.get(qn("w:ascii")) if fonts is not None else None
        if (style.get(qn("w:type")) == "paragraph" and font
                and ("code" in name.lower() or "mono" in name.lower()
                     or "preformatted" in name.lower())):
            return style.get(qn("w:styleId")), font
    for style in styles_root.iter(qn("w:style")):
        rpr = style.find(qn("w:rPr"))
        fonts = rpr.find(qn("w:rFonts")) if rpr is not None else None
        font = fonts.get(qn("w:ascii")) if fonts is not None else None
        if font and any(token in font.lower() for token in ("courier", "consolas", "mono")):
            style_id = (style.get(qn("w:styleId"))
                        if style.get(qn("w:type")) == "paragraph" else None)
            return style_id, font
    return None, "Consolas"


def _set_text(paragraph, text: str, style_id: str | None = None,
              inherit: bool = True, font: str | None = None, run_rpr=None) -> None:
    p = etree.Element(qn("w:p"))
    if inherit and paragraph.find(qn("w:pPr")) is not None:
        p.append(deepcopy(paragraph.find(qn("w:pPr"))))
    if style_id:
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            ppr = etree.Element(qn("w:pPr"))
            p.insert(0, ppr)
        pstyle = ppr.find(qn("w:pStyle"))
        if pstyle is None:
            pstyle = etree.Element(qn("w:pStyle"))
            ppr.insert(0, pstyle)
        pstyle.set(qn("w:val"), style_id)
    run = etree.SubElement(p, qn("w:r"))
    if run_rpr is not None:
        run.append(deepcopy(run_rpr))
    if font:
        rpr = run.find(qn("w:rPr"))
        if rpr is None:
            rpr = etree.Element(qn("w:rPr"))
            run.insert(0, rpr)
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = etree.SubElement(rpr, qn("w:rFonts"))
        for attr in ("ascii", "hAnsi", "cs", "eastAsia"):
            fonts.set(qn(f"w:{attr}"), font)
    for n, line in enumerate(str(text).split("\n")):
        if n:
            etree.SubElement(run, qn("w:br"))
        t = etree.SubElement(run, qn("w:t"))
        if line[:1].isspace() or line[-1:].isspace():
            t.set(f"{{{XML_SPACE}}}space", "preserve")
        t.text = line
    return p


def _replace_in_paragraph(paragraph, needle: str, replacement: str) -> bool:
    nodes = list(paragraph.iter(qn("w:t")))
    texts = [node.text or "" for node in nodes]
    joined = "".join(texts)
    start = joined.find(needle)
    if start < 0:
        return False
    end = start + len(needle)
    starts, cursor = [], 0
    for text in texts:
        starts.append(cursor)
        cursor += len(text)
    first = next(i for i, text in enumerate(texts) if starts[i] <= start < starts[i] + len(text))
    last = next(i for i, text in enumerate(texts) if starts[i] < end <= starts[i] + len(text))
    prefix = texts[first][:start - starts[first]]
    suffix = texts[last][end - starts[last]:]
    if first == last:
        nodes[first].text = prefix + replacement + suffix
    else:
        nodes[first].text = prefix + replacement
        for i in range(first + 1, last):
            nodes[i].text = ""
        nodes[last].text = suffix
    for node in (nodes[first], nodes[last]):
        if (node.text or "")[:1].isspace() or (node.text or "")[-1:].isspace():
            node.set(f"{{{XML_SPACE}}}space", "preserve")
    return True


def _target_cell(body, operation: dict):
    tables = _tables(body)
    index = int(operation.get("table", 0))
    row, col = int(operation["row"]), int(operation["col"])
    if index < 0 or index >= len(tables):
        raise ValueError(f"Table index out of range: {index}")
    rows = tables[index].findall(qn("w:tr"))
    if row < 0 or row >= len(rows):
        raise ValueError(f"Table {index} row out of range: {row}")
    cells = rows[row].findall(qn("w:tc"))
    logical_col = 0
    for cell in cells:
        tcpr = cell.find(qn("w:tcPr"))
        span = tcpr.find(qn("w:gridSpan")) if tcpr is not None else None
        width = int(span.get(qn("w:val"))) if span is not None else 1
        if logical_col <= col < logical_col + width:
            return cell
        logical_col += width
    raise ValueError(f"Table {index} column out of range: {col}")


def _insert_after(anchor, element) -> None:
    anchor.addnext(element)


def _paragraph_specs(value) -> list[dict]:
    if isinstance(value, str):
        return [{"text": value}]
    if isinstance(value, list):
        return [({"text": item} if isinstance(item, str) else item) for item in value]
    raise ValueError("paragraphs/content must be a string or list of strings/objects")


def _apply(body, styles_root, operation: dict) -> int:
    action = operation.get("op") or operation.get("action")
    if action == "replace_text":
        anchor = str(operation["anchor"])
        replacement = str(operation["new_text"])
        target = _resolve_paragraph(body, anchor)
        if not _replace_in_paragraph(target, anchor, replacement):
            raise ValueError(f"Text not found in resolved anchor: {anchor!r}")
        return 0
    if action in ("insert_after", "insert_code_block"):
        target = _resolve_paragraph(body, str(operation["anchor"]))
        if action == "insert_code_block":
            code_style, code_font = _code_style(styles_root)
            specs = [{"text": str(operation["text"]), "style_id": code_style,
                      "font": code_font}]
        else:
            specs = _paragraph_specs(operation.get("paragraphs", []))
        previous = target
        for spec in specs:
            style = spec.get("style")
            style_id = spec.get("style_id") or (_style_id(styles_root, style) if style else None)
            paragraph = _set_text(target, spec.get("text", ""), style_id=style_id,
                                  inherit=not bool(style_id),
                                  font=spec.get("font"))
            _insert_after(previous, paragraph)
            previous = paragraph
        return len(specs)
    if action in ("fill_table_cell", "append_to_cell"):
        cell = _target_cell(body, operation)
        content = str(operation.get("content", ""))
        lines = content.split("\n")
        paragraphs = cell.findall(qn("w:p"))
        if action == "fill_table_cell":
            first = paragraphs[0] if paragraphs else etree.Element(qn("w:p"))
            ppr = first.find(qn("w:pPr"))
            old_run = first.find(qn("w:r"))
            run_rpr = old_run.find(qn("w:rPr")) if old_run is not None else None
            for child in list(first):
                if child is not ppr:
                    first.remove(child)
            if ppr is None:
                first.append(etree.Element(qn("w:r")))
                run = first[-1]
            else:
                run = etree.SubElement(first, qn("w:r"))
            if run_rpr is not None:
                run.insert(0, deepcopy(run_rpr))
            for n, line in enumerate(lines):
                if n:
                    etree.SubElement(run, qn("w:br"))
                t = etree.SubElement(run, qn("w:t"))
                t.text = line
            if not paragraphs:
                cell.append(first)
            for p in paragraphs[1:]:
                cell.remove(p)
            return 0
        for line in lines:
            last_run = (paragraphs[-1].find(qn("w:r")) if paragraphs else None)
            run_rpr = last_run.find(qn("w:rPr")) if last_run is not None else None
            p = _set_text(paragraphs[-1] if paragraphs else etree.Element(qn("w:p")),
                          line, inherit=bool(paragraphs), run_rpr=run_rpr)
            cell.append(p)
        return len(lines)
    raise ValueError(f"Unknown patch operation: {action!r}")


def _counts(root) -> tuple[int, int]:
    return len(list(root.iter(qn("w:p")))), len(list(root.iter(qn("w:tbl"))))


def _heading_texts(root) -> list[str]:
    headings = []
    for p in root.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        style = ppr.find(qn("w:pStyle")) if ppr is not None else None
        if style is not None and style.get(qn("w:val"), "").lower().startswith("heading"):
            text = _text(p)
            if text:
                headings.append(text)
    return headings


def patch_docx(path: str, operations: list[dict], output_path: str | None = None,
               backup: bool = True, overwrite: bool = False) -> dict:
    """Apply anchored changes to DOCX OOXML while preserving untouched ZIP parts."""
    src = Path(path).expanduser().resolve()
    if not src.is_file():
        raise ValueError(f"DOCX file not found: {src}")
    if src.suffix.lower() not in (".docx", ".docm"):
        raise ValueError("docx_patch supports .docx and .docm files")
    dest = Path(output_path).expanduser().resolve() if output_path else src.with_name(
        f"{src.stem}.patched{src.suffix}")
    dest.parent.mkdir(parents=True, exist_ok=True)
    same = dest == src
    if dest.exists() and not same and not overwrite:
        raise ValueError(f"Output already exists: {dest}. Set overwrite=true to replace it.")
    backup_path = None
    if same:
        if not backup:
            raise ValueError("In-place edits require backup=true; refusing silent overwrite.")
        backup_path = src.with_suffix(src.suffix + ".bak")
        if backup_path.exists():
            raise ValueError(f"Backup already exists: {backup_path}; refusing to overwrite it.")
    if not operations:
        raise ValueError("operations must contain at least one patch operation")

    with zipfile.ZipFile(src) as original:
        try:
            document_xml = original.read("word/document.xml")
            styles_xml = original.read("word/styles.xml")
        except KeyError as exc:
            raise ValueError(f"Invalid Word package: missing {exc}") from None
        parser = etree.XMLParser(resolve_entities=False, no_network=True, recover=False)
        root = etree.fromstring(document_xml, parser)
        styles_root = etree.fromstring(styles_xml, parser)
        body = root.find(f"{{{W}}}body")
        before_p, before_t = _counts(root)
        before_headings = _heading_texts(root)
        added = 0
        for operation in operations:
            if not isinstance(operation, dict):
                raise ValueError("Each operation must be an object")
            added += _apply(body, styles_root, operation)
        after_p, after_t = _counts(root)
        after_headings = _heading_texts(root)
        missing = [heading for heading in before_headings if heading not in after_headings]
        # A requested replace_text may intentionally edit its own heading anchor.
        for operation in operations:
            if (operation.get("op") or operation.get("action")) == "replace_text":
                anchor = str(operation.get("anchor", ""))
                replacement = str(operation.get("new_text", ""))
                if anchor in missing and replacement in after_headings:
                    missing.remove(anchor)
        if missing:
            raise ValueError("Patch would remove or alter unrequested heading(s): "
                             + ", ".join(repr(x) for x in missing))
        updated_xml = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                     standalone=True)
        # Validate all XML package members before writing and preserve every ZIP member.
        for name in (n for n in original.namelist()
                     if n.endswith((".xml", ".rels")) and n != "word/document.xml"):
            etree.fromstring(original.read(name), parser)

        fd, temporary = tempfile.mkstemp(prefix=f".{dest.name}.", suffix=".tmp",
                                         dir=str(dest.parent))
        os.close(fd)
        try:
            with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as result:
                for info in original.infolist():
                    content = updated_xml if info.filename == "word/document.xml" else original.read(
                        info.filename)
                    result.writestr(info, content)
            with zipfile.ZipFile(temporary) as check:
                bad = check.testzip()
                if bad:
                    raise ValueError(f"Output DOCX ZIP validation failed at {bad}")
                etree.fromstring(check.read("word/document.xml"), parser)
            if backup_path:
                shutil.copy2(src, backup_path)
            os.replace(temporary, dest)
        except Exception:
            if os.path.exists(temporary):
                os.unlink(temporary)
            if backup_path and backup_path.exists() and dest == src:
                # Keep a completed backup for recovery; do not remove user data.
                pass
            raise
    return {"path": str(dest), "backup": str(backup_path) if backup_path else None,
            "paragraphs_before": before_p, "paragraphs_after": after_p,
            "paragraphs_added": after_p - before_p, "paragraphs_removed": before_p - after_p,
            "tables_before": before_t, "tables_after": after_t,
            "headings_preserved": True}


def insert_image(path: str, image_path: str, anchor: str | None = None,
                 position: str = "after", width_cm: float | None = None,
                 caption: str | None = None, alignment: str = "center",
                 table: int = 0, row: int | None = None, col: int | None = None,
                 output_path: str | None = None, backup: bool = True,
                 overwrite: bool = False) -> dict:
    """Insert an image by editing document XML and relationships, preserving all ZIP parts."""
    from docx import Document
    from docx.shared import Cm, Emu
    from docx.image.image import Image as DocxImage

    src = Path(path).expanduser().resolve()
    image = Path(image_path).expanduser().resolve()
    if not src.is_file():
        raise ValueError(f"DOCX file not found: {src}")
    if not image.is_file():
        raise ValueError(f"Image file not found: {image}")
    if src.suffix.lower() not in (".docx", ".docm"):
        raise ValueError("docx_insert_image supports .docx and .docm files")
    if position not in ("after", "before", "replace", "append_cell"):
        raise ValueError("position must be after, before, replace, or append_cell")
    if alignment not in ("left", "center", "right"):
        raise ValueError("alignment must be left, center, or right")
    if position == "append_cell" and (row is None or col is None):
        raise ValueError("append_cell requires row and col")
    if position != "append_cell" and not anchor:
        raise ValueError("anchor is required unless position='append_cell'")

    dest = Path(output_path).expanduser().resolve() if output_path else src.with_name(
        f"{src.stem}.patched{src.suffix}")
    dest.parent.mkdir(parents=True, exist_ok=True)
    same = dest == src
    if dest.exists() and not same and not overwrite:
        raise ValueError(f"Output already exists: {dest}. Set overwrite=true to replace it.")
    backup_path = None
    if same:
        if not backup:
            raise ValueError("In-place image insertion requires backup=true")
        backup_path = src.with_suffix(src.suffix + ".bak")
        if backup_path.exists():
            raise ValueError(f"Backup already exists: {backup_path}; refusing to overwrite it")

    doc = Document(str(src))
    doc_root = doc.element.body
    target_cell = None
    target_paragraph = None
    if position == "append_cell":
        tables = _tables(doc_root)
        if table < 0 or table >= len(tables):
            raise ValueError(f"Table index out of range: {table}")
        target_cell = _target_cell(doc_root, {"table": table, "row": row, "col": col})
        container_width = None
        tcpr = target_cell.find(qn("w:tcPr"))
        tcw = tcpr.find(qn("w:tcW")) if tcpr is not None else None
        if tcw is not None and tcw.get(qn("w:w")):
            container_width = int(tcw.get(qn("w:w"))) * 635
    else:
        target_paragraph = _resolve_paragraph(doc_root, str(anchor))
        sect = doc.sections[-1]
        container_width = sect.page_width - sect.left_margin - sect.right_margin

    img_desc = DocxImage.from_file(str(image))
    max_width = int(container_width) if container_width else int(Cm(16))
    requested_width = int(Cm(width_cm)) if width_cm else img_desc.width
    picture_width = min(requested_width, max_width)
    if picture_width <= 0:
        raise ValueError("Image width must be greater than zero")

    # Ask python-docx only to produce standards-compliant DrawingML. Package is not saved;
    # its image bytes and drawing are transplanted into the original ZIP below.
    temp_doc = Document()
    drawing = temp_doc.part.new_pic_inline(str(image), width=Emu(picture_width))
    temp_rel_id = next(blip.get(qn("r:embed")) for blip in drawing.iter(qn("a:blip")))
    temp_rel = temp_doc.part.rels[temp_rel_id]
    image_blob = temp_rel.target_part.blob
    ext = Path(str(temp_rel.target_part.partname)).suffix.lstrip(".").lower() or "png"
    media_name = f"image-doc-mantap-{uuid.uuid4().hex}.{ext}"
    new_rel_id = None

    parser = etree.XMLParser(resolve_entities=False, no_network=True, recover=False)
    with zipfile.ZipFile(src) as original:
        document_xml = original.read("word/document.xml")
        rels_xml = original.read("word/_rels/document.xml.rels")
        types_xml = original.read("[Content_Types].xml")
        root = etree.fromstring(document_xml, parser)
        rels_root = etree.fromstring(rels_xml, parser)
        types_root = etree.fromstring(types_xml, parser)
        body = root.find(f"{{{W}}}body")
        before_p, before_t = _counts(root)
        headings_before = _heading_texts(root)

        if position == "append_cell":
            target = _target_cell(body, {"table": table, "row": row, "col": col})
            parent_for_insert = target
            insertion_anchor = target.findall(qn("w:p"))[-1] if target.findall(qn("w:p")) else None
        else:
            target_paragraph = _resolve_paragraph(body, str(anchor))
            parent_for_insert = target_paragraph.getparent()
            insertion_anchor = target_paragraph

        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        image_rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        used = [int(rel.get("Id", "rId0")[3:]) for rel in rels_root
                if rel.get("Id", "").startswith("rId") and rel.get("Id", "")[3:].isdigit()]
        new_rel_id = f"rId{max(used, default=0) + 1}"
        relation = etree.SubElement(rels_root, f"{{{rel_ns}}}Relationship")
        relation.set("Id", new_rel_id)
        relation.set("Type", image_rel_type)
        relation.set("Target", f"media/{media_name}")
        for blip in drawing.iter(qn("a:blip")):
            blip.set(qn("r:embed"), new_rel_id)

        content_type = temp_rel.target_part.content_type
        if not any(el.tag.endswith("Default") and el.get("Extension", "").lower() == ext
                   for el in types_root):
            default = etree.SubElement(types_root,
                "{http://schemas.openxmlformats.org/package/2006/content-types}Default")
            default.set("Extension", ext)
            default.set("ContentType", content_type)

        paragraph = etree.Element(qn("w:p"))
        ppr = etree.SubElement(paragraph, qn("w:pPr"))
        jc = etree.SubElement(ppr, qn("w:jc"))
        jc.set(qn("w:val"), alignment)
        run = etree.SubElement(paragraph, qn("w:r"))
        drawing_container = etree.SubElement(run, qn("w:drawing"))
        drawing_container.append(deepcopy(drawing))
        docpr = drawing_container.find(".//" + qn("wp:docPr"))
        if docpr is not None:
            ids = [int(el.get("id", "0")) for el in root.iter(qn("wp:docPr"))
                   if el.get("id", "").isdigit()]
            docpr.set("id", str(max(ids, default=0) + 1))
        c_nv_pr = drawing_container.find(".//" + qn("pic:cNvPr"))
        if c_nv_pr is not None:
            ids = [int(el.get("id", "0")) for el in root.iter(qn("pic:cNvPr"))
                   if el.get("id", "").isdigit()]
            c_nv_pr.set("id", str(max(ids, default=0) + 1))
        if position == "append_cell":
            if insertion_anchor is not None:
                insertion_anchor.addnext(paragraph)
            else:
                parent_for_insert.append(paragraph)
        elif position == "before":
            insertion_anchor.addprevious(paragraph)
        elif position == "after":
            insertion_anchor.addnext(paragraph)
        else:
            insertion_anchor.addprevious(paragraph)
            parent_for_insert.remove(insertion_anchor)

        added = 1
        if caption:
            cap = etree.Element(qn("w:p"))
            cppr = etree.SubElement(cap, qn("w:pPr"))
            cstyle = etree.SubElement(cppr, qn("w:pStyle"))
            cstyle.set(qn("w:val"), "Caption")
            cjc = etree.SubElement(cppr, qn("w:jc"))
            cjc.set(qn("w:val"), alignment)
            crun = etree.SubElement(cap, qn("w:r"))
            ctext = etree.SubElement(crun, qn("w:t"))
            ctext.text = caption
            paragraph.addnext(cap)
            added += 1

        after_p, after_t = _counts(root)
        headings_after = _heading_texts(root)
        if any(h not in headings_after for h in headings_before):
            raise ValueError("Image insertion would remove or alter an existing heading")
        updated_document = etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                                         standalone=True)
        updated_rels = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8",
                                      standalone=True)
        updated_types = etree.tostring(types_root, xml_declaration=True, encoding="UTF-8",
                                       standalone=True)
        fd, temporary = tempfile.mkstemp(prefix=f".{dest.name}.", suffix=".tmp",
                                         dir=str(dest.parent))
        os.close(fd)
        try:
            with zipfile.ZipFile(temporary, "w", zipfile.ZIP_DEFLATED) as result:
                for info in original.infolist():
                    if info.filename == "word/document.xml":
                        content = updated_document
                    elif info.filename == "word/_rels/document.xml.rels":
                        content = updated_rels
                    elif info.filename == "[Content_Types].xml":
                        content = updated_types
                    else:
                        content = original.read(info.filename)
                    result.writestr(info, content)
                result.writestr(f"word/media/{media_name}", image_blob)
            with zipfile.ZipFile(temporary) as check:
                if check.testzip():
                    raise ValueError("Output DOCX ZIP validation failed")
                etree.fromstring(check.read("word/document.xml"), parser)
                etree.fromstring(check.read("word/_rels/document.xml.rels"), parser)
                etree.fromstring(check.read("[Content_Types].xml"), parser)
            if backup_path:
                shutil.copy2(src, backup_path)
            os.replace(temporary, dest)
        except Exception:
            if os.path.exists(temporary):
                os.unlink(temporary)
            raise

    return {"path": str(dest), "backup": str(backup_path) if backup_path else None,
            "paragraphs_before": before_p, "paragraphs_after": after_p,
            "paragraphs_added": after_p - before_p, "paragraphs_removed": before_p - after_p,
            "tables_before": before_t, "tables_after": after_t,
            "images_added": 1, "caption_added": bool(caption), "headings_preserved": True}
