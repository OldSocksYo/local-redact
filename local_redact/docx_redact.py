"""读写 docx 压缩包：正文、页眉页脚、图片、水印、背景。"""

from __future__ import annotations

import posixpath
import re
import zipfile
from pathlib import Path

from lxml import etree

from local_redact.fields import COLUMN_RULES, HEADER_WORDS, LABELS, Rule, normalize_label
from local_redact.textutil import Redactor

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
V = "{urn:schemas-microsoft-com:vml}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_SKIP_PARTS = {
    "word/styles.xml",
    "word/settings.xml",
    "word/webSettings.xml",
    "word/fontTable.xml",
    "word/numbering.xml",
}
_DROP_REL_SUFFIXES = ("/image", "/oleObject", "/package")
_IDENTITY_ATTRS = {"author", "initials", "userid"}
_SENSITIVE_PROP_NAMES = {
    "company",
    "manager",
    "creator",
    "lastauthor",
    "owner",
    "operator",
    "lastmodifiedby",
}
_SKIP_VALUE_ATTRS = {
    "fmtid",
    "pid",
    "id",
    "idref",
    "contenttype",
    "uri",
    "type",
    "target",
    "targetmode",
}


class DocxError(Exception):
    pass


def redact_docx(src: Path, dest: Path, redactor: Redactor) -> dict[str, int]:
    if src.suffix.lower() == ".doc":
        raise DocxError("这里只处理已经转好的 docx。")
    if src.resolve() == dest.resolve():
        raise DocxError("不能覆盖原稿")
    try:
        package = _DocxPackage.read(src)
    except zipfile.BadZipFile as exc:
        raise DocxError("不是有效的 docx") from exc

    redactor.begin_file()
    parts = [name for name in package.names() if _is_content_xml(name)]
    roots: dict[str, etree._Element] = {}
    for name in parts:
        roots[name] = _parse_xml(package.files[name])

    for root in roots.values():
        _redact_tables(root, redactor)
    # 两遍：第一遍从栏目里收集名称，第二遍才能改到先出现的页眉、批注。
    for _ in range(2):
        for root in roots.values():
            _redact_paragraphs(root, redactor)
            _redact_side_text(root, redactor)

    shapes = 0
    for name, root in roots.items():
        shapes += _remove_visuals(root)
        _redact_identity_attrs(root, redactor)
        package.files[name] = _serialize_xml(root)

    _redact_properties(package, redactor)
    _redact_settings(package, redactor)
    _redact_custom_parts(package, redactor)
    _scrub_package_leaks(package, redactor)
    media = _cleanup_package(package, redactor)
    dest.parent.mkdir(parents=True, exist_ok=True)
    package.save(dest)
    return {"shapes": shapes, "media": media, "fields": redactor.field_hits}


def _is_content_xml(name: str) -> bool:
    if not name.startswith("word/") or not name.endswith(".xml"):
        return False
    if name in _SKIP_PARTS or "/theme/" in name:
        return False
    return True


def _parse_xml(data: bytes) -> etree._Element:
    if data.startswith(b"\xef\xbb\xbf"):
        data = data[3:]
    return etree.fromstring(data)


def _serialize_xml(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _nearest(element: etree._Element, tag: str) -> etree._Element | None:
    parent = element.getparent()
    while parent is not None:
        if parent.tag == tag:
            return parent
        parent = parent.getparent()
    return None


def _direct_text_nodes(paragraph: etree._Element) -> list[etree._Element]:
    nodes = []
    for node in paragraph.iter(f"{W}t"):
        if _nearest(node, f"{W}p") is not paragraph:
            continue
        if _nearest(node, f"{W}del") is not None:
            continue
        nodes.append(node)
    return nodes


def _paragraph_text(paragraph: etree._Element) -> str:
    return "".join(node.text or "" for node in _direct_text_nodes(paragraph))


def _set_paragraph_text(paragraph: etree._Element, text: str) -> None:
    nodes = _direct_text_nodes(paragraph)
    if not nodes:
        if not text:
            return
        run = etree.SubElement(paragraph, f"{W}r")
        node = etree.SubElement(run, f"{W}t")
        node.text = text
        if text[:1].isspace() or text[-1:].isspace():
            node.set(XML_SPACE, "preserve")
        return
    nodes[0].text = text
    if text[:1].isspace() or text[-1:].isspace():
        nodes[0].set(XML_SPACE, "preserve")
    for node in nodes[1:]:
        node.text = ""


def _cell_paragraphs(cell: etree._Element) -> list[etree._Element]:
    return [child for child in cell if child.tag == f"{W}p"]


def _cell_text(cell: etree._Element) -> str:
    return "".join(_paragraph_text(paragraph) for paragraph in _cell_paragraphs(cell))


def _set_cell_text(cell: etree._Element, text: str) -> None:
    paragraphs = _cell_paragraphs(cell)
    if not paragraphs:
        paragraphs = [etree.SubElement(cell, f"{W}p")]
    _set_paragraph_text(paragraphs[0], text)
    for paragraph in paragraphs[1:]:
        _set_paragraph_text(paragraph, "")


def _table_rows(table: etree._Element) -> list[list[tuple[etree._Element, int, int]]]:
    rows = []
    for tr in table.findall(f"{W}tr"):
        cells = []
        col = 0
        for tc in tr.findall(f"{W}tc"):
            span = 1
            props = tc.find(f"{W}tcPr")
            if props is not None:
                grid = props.find(f"{W}gridSpan")
                if grid is not None and grid.get(f"{W}val", "").isdigit():
                    span = int(grid.get(f"{W}val"))
            cells.append((tc, col, span))
            col += span
        rows.append(cells)
    return rows


def _cell_at(row: list[tuple[etree._Element, int, int]], col: int) -> etree._Element | None:
    for cell, start, span in row:
        if start <= col < start + span:
            return cell
    return None


def _is_header_row(row: list[tuple[etree._Element, int, int]]) -> bool:
    labels = []
    for cell, _, _ in row:
        text = _cell_text(cell).strip()
        if text:
            labels.append(normalize_label(text))
    return len(labels) >= 2 and all(label in HEADER_WORDS for label in labels)


def _is_note_row(row: list[tuple[etree._Element, int, int]]) -> bool:
    """表下「备注：…」说明行，不按人员/坐标等列规则打码。"""
    for cell, _, _ in row:
        text = _cell_text(cell).strip()
        if text.startswith(("备注", "注：", "注:")):
            return True
    return False


def _replace_cell(cell: etree._Element, redactor: Redactor, rule: Rule) -> None:
    raw = _cell_text(cell).strip()
    if not raw:
        return
    key = normalize_label(raw)
    if key in LABELS or key in HEADER_WORDS:
        return
    _set_cell_text(cell, redactor.apply_rule(raw, rule))


def _redact_tables(root: etree._Element, redactor: Redactor) -> None:
    for table in root.iter(f"{W}tbl"):
        rows = _table_rows(table)
        header_at = [index for index, row in enumerate(rows) if _is_header_row(row)]
        for position, start_row in enumerate(header_at):
            end_row = header_at[position + 1] if position + 1 < len(header_at) else len(rows)
            columns = {}
            for cell, col, _span in rows[start_row]:
                key = normalize_label(_cell_text(cell))
                if key in COLUMN_RULES:
                    columns[col] = COLUMN_RULES[key]
            for row_index in range(start_row + 1, end_row):
                if _is_note_row(rows[row_index]):
                    continue
                for col, rule in columns.items():
                    cell = _cell_at(rows[row_index], col)
                    if cell is not None:
                        _replace_cell(cell, redactor, rule)

        header_set = set(header_at)
        for row_index, row in enumerate(rows):
            if row_index in header_set:
                continue
            for cell, col, span in row:
                _redact_label_cell(rows, row_index, cell, col, span, redactor)


def _redact_label_cell(
    rows: list[list[tuple[etree._Element, int, int]]],
    row_index: int,
    cell: etree._Element,
    col: int,
    span: int,
    redactor: Redactor,
) -> None:
    paragraphs = _cell_paragraphs(cell)
    texts = [_paragraph_text(paragraph).strip() for paragraph in paragraphs]
    key = normalize_label(texts[0]) if texts else ""
    if key in LABELS and any(texts[1:]):
        value = "".join(texts[1:]).strip()
        if value:
            _set_paragraph_text(paragraphs[1], redactor.apply_rule(value, LABELS[key]))
            for paragraph in paragraphs[2:]:
                _set_paragraph_text(paragraph, "")
        return
    if normalize_label(_cell_text(cell)) not in LABELS:
        return
    rule = LABELS[normalize_label(_cell_text(cell))]
    target = _cell_at(rows[row_index], col + span) if row_index < len(rows) else None
    target_text = _cell_text(target).strip() if target is not None else ""
    target_key = normalize_label(target_text) if target_text else ""
    if target is None or not target_text or target_key in LABELS or target_key in HEADER_WORDS:
        if row_index + 1 < len(rows):
            target = _cell_at(rows[row_index + 1], col)
            target_text = _cell_text(target).strip() if target is not None else ""
            target_key = normalize_label(target_text) if target_text else ""
            if target is None or not target_text or target_key in LABELS or target_key in HEADER_WORDS:
                return
        else:
            return
    _set_cell_text(target, redactor.apply_rule(target_text, rule))


def _redact_paragraphs(root: etree._Element, redactor: Redactor) -> None:
    for paragraph in root.iter(f"{W}p"):
        old = _paragraph_text(paragraph)
        new = redactor.redact_string(old)
        if new != old:
            _set_paragraph_text(paragraph, new)


def _redact_side_text(root: etree._Element, redactor: Redactor) -> None:
    for node in root.iter(f"{W}delText", f"{W}instrText"):
        old = node.text or ""
        new = redactor.redact_string(old)
        if new != old:
            node.text = new
    for node in root.iter(f"{W}t"):
        if _nearest(node, f"{W}del") is None:
            continue
        old = node.text or ""
        new = redactor.redact_string(old)
        if new != old:
            node.text = new
    for node in root.iter(f"{V}textpath"):
        old = node.get("string") or ""
        new = redactor.redact_string(old)
        if new != old:
            node.set("string", new)
    for node in root.iter(f"{W}hyperlink"):
        tip = node.get(f"{W}tooltip")
        if not tip:
            continue
        new = redactor.redact_string(tip)
        if new != tip:
            node.set(f"{W}tooltip", new)


def _has_image(element: etree._Element) -> bool:
    for node in element.iter():
        if node.tag in {f"{A}blip", f"{V}imagedata"}:
            return True
        if node.tag == f"{V}fill" and node.get(f"{R}id"):
            return True
    return False


def _is_watermark(element: etree._Element) -> bool:
    for node in element.iter():
        if node.tag == f"{V}textpath":
            return True
        name = f"{node.get('name') or ''}{node.get('id') or ''}"
        if "watermark" in name.lower():
            return True
    return False


def _should_remove(element: etree._Element) -> bool:
    if element.tag == f"{W}object":
        return True
    if element.tag == f"{W}background":
        return _has_image(element)
    if element.tag in {f"{W}drawing", f"{W}pict"}:
        return _has_image(element) or _is_watermark(element)
    return False


def _remove_visuals(root: etree._Element) -> int:
    targets = [node for node in root.iter() if _should_remove(node)]
    removed = 0
    for node in targets:
        parent = node.getparent()
        if parent is not None and node.getparent() is parent:
            parent.remove(node)
            removed += 1
    for alt in list(root.iter(f"{MC}AlternateContent")):
        if any(node.tag == f"{W}t" and (node.text or "").strip() for node in alt.iter(f"{W}t")):
            continue
        if any(node.tag in {f"{W}drawing", f"{W}pict"} for node in alt.iter()):
            continue
        parent = alt.getparent()
        if parent is not None:
            parent.remove(alt)
    return removed


def _redact_properties(package: _DocxPackage, redactor: Redactor) -> None:
    core_name = "docProps/core.xml"
    if core_name in package.files:
        root = _parse_xml(package.files[core_name])
        for element in root.iter():
            local = element.tag.rsplit("}", 1)[-1]
            if local in {"creator", "lastModifiedBy"} and (element.text or "").strip():
                redactor.mapping.drop("文档属性", element.text.strip())
                element.text = ""
                continue
            if element.text and element.text.strip():
                updated = redactor.redact_string(element.text)
                if updated != element.text:
                    element.text = updated
        _redact_identity_attrs(root, redactor)
        package.files[core_name] = _serialize_xml(root)
    app_name = "docProps/app.xml"
    if app_name in package.files:
        root = _parse_xml(package.files[app_name])
        for element in root.iter():
            local = element.tag.rsplit("}", 1)[-1]
            if local in {"Company", "Manager"} and (element.text or "").strip():
                redactor.mapping.drop("文档属性", element.text.strip())
                element.text = ""
        _redact_identity_attrs(root, redactor)
        package.files[app_name] = _serialize_xml(root)


def _attr_local(element: etree._Element, local_name: str) -> str:
    for attr, value in element.attrib.items():
        if attr.rsplit("}", 1)[-1].lower() == local_name.lower():
            return value
    return ""


def _named_attr(element: etree._Element, local_name: str) -> str | None:
    for attr in element.attrib:
        if attr.rsplit("}", 1)[-1] == local_name:
            return attr
    return None


def _redact_identity_attrs(root: etree._Element, redactor: Redactor) -> None:
    """批注、修订、人员部件上的作者名。正文替换不会碰到这些属性。"""
    for element in root.iter():
        for attr, value in list(element.attrib.items()):
            local = attr.rsplit("}", 1)[-1].lower()
            if local not in _IDENTITY_ATTRS:
                continue
            text = value.strip()
            if not text or text == "已脱敏":
                continue
            redactor.mapping.drop("文档属性", text)
            element.set(attr, "已脱敏")


def _redact_settings(package: _DocxPackage, redactor: Redactor) -> None:
    name = "word/settings.xml"
    if name not in package.files:
        return
    root = _parse_xml(package.files[name])
    for element in root.iter():
        if element.tag.rsplit("}", 1)[-1] != "docVar":
            continue
        val_attr = _named_attr(element, "val")
        if val_attr is None:
            continue
        old = element.get(val_attr) or ""
        if not old.strip():
            continue
        if _is_path_text(old):
            redactor.mapping.drop("文档属性", old.strip())
            element.set(val_attr, "")
            continue
        updated = redactor.redact_string(old)
        if updated != old:
            element.set(val_attr, updated)
    _redact_identity_attrs(root, redactor)
    package.files[name] = _serialize_xml(root)


def _redact_custom_parts(package: _DocxPackage, redactor: Redactor) -> None:
    names = [
        name
        for name in package.files
        if name == "docProps/custom.xml" or (name.startswith("customXml/") and name.endswith(".xml"))
    ]
    for name in names:
        root = _parse_xml(package.files[name])
        _redact_loose_xml(root, redactor)
        _redact_identity_attrs(root, redactor)
        package.files[name] = _serialize_xml(root)


def _redact_loose_xml(root: etree._Element, redactor: Redactor) -> None:
    cleared: set[int] = set()
    for element in root.iter():
        if id(element) in cleared:
            continue
        local = element.tag.rsplit("}", 1)[-1].lower()
        prop_name = _attr_local(element, "name").strip().lower()
        if local == "property" and prop_name in _SENSITIVE_PROP_NAMES:
            for node in element.iter():
                cleared.add(id(node))
                if node.text and node.text.strip():
                    redactor.mapping.drop("文档属性", node.text.strip())
                    node.text = ""
            continue
        if element.text and element.text.strip():
            updated = redactor.redact_string(element.text)
            if updated != element.text:
                element.text = updated
        for attr, value in list(element.attrib.items()):
            attr_name = attr.rsplit("}", 1)[-1].lower()
            if attr_name in _SKIP_VALUE_ATTRS or attr_name in _IDENTITY_ATTRS or not value.strip():
                continue
            updated = redactor.redact_string(value)
            if updated != value:
                element.set(attr, updated)


def _is_local_target(target: str) -> bool:
    text = target.strip()
    if not text:
        return False
    lowered = text.lower()
    if lowered.startswith("file:"):
        return True
    if text.startswith("\\\\") or text.startswith("//"):
        return True
    return re.match(r"^[A-Za-z]:[\\/]", text) is not None


def _is_path_text(value: str) -> bool:
    return _is_local_target(value) or ":\\" in value or "/Users/" in value or "\\Users\\" in value


def _target_is_thumbnail(target: str) -> bool:
    leaf = target.replace("\\", "/").split("?", 1)[0].rstrip("/").rsplit("/", 1)[-1].lower()
    return leaf.startswith("thumbnail.")


def _unlink_rel_ids(root: etree._Element, removed_ids: set[str]) -> bool:
    if not removed_ids:
        return False
    changed = False
    for element in list(root.iter()):
        for attr, value in list(element.attrib.items()):
            if value not in removed_ids or not attr.startswith(R):
                continue
            parent = element.getparent()
            if parent is not None:
                parent.remove(element)
            else:
                del element.attrib[attr]
            changed = True
            break
    return changed


def _scrub_package_leaks(package: _DocxPackage, redactor: Redactor) -> None:
    """首页缩略图，以及指向本机路径的外部关系（模板、邮件合并数据源）。"""
    for name in list(package.files):
        leaf = name.rsplit("/", 1)[-1].lower()
        if name.startswith("docProps/") and leaf.startswith("thumbnail"):
            del package.files[name]
    for name in list(package.files):
        if not name.endswith(".rels"):
            continue
        root = _parse_xml(package.files[name])
        removed_ids: set[str] = set()
        changed = False
        for rel in list(root):
            if rel.tag != f"{PKG_REL}Relationship":
                continue
            target = rel.get("Target") or ""
            rel_type = rel.get("Type") or ""
            mode = rel.get("TargetMode") or ""
            drop = False
            if "thumbnail" in rel_type.lower() or _target_is_thumbnail(target):
                drop = True
            elif mode == "External" and _is_local_target(target):
                redactor.mapping.drop("链接", target)
                drop = True
            if not drop:
                continue
            rel_id = rel.get("Id")
            if rel_id:
                removed_ids.add(rel_id)
            root.remove(rel)
            changed = True
        if not changed:
            continue
        package.files[name] = _serialize_xml(root)
        part = _part_of_rels(name)
        if part and part in package.files and part.endswith(".xml"):
            part_root = _parse_xml(package.files[part])
            if _unlink_rel_ids(part_root, removed_ids):
                package.files[part] = _serialize_xml(part_root)


def _referenced_ids(root: etree._Element) -> set[str]:
    found = set()
    for element in root.iter():
        for attr, value in element.attrib.items():
            if attr.startswith(R):
                found.add(value)
    return found


def _part_of_rels(rels_name: str) -> str | None:
    if not rels_name.endswith(".rels") or "/_rels/" not in rels_name:
        return None
    parent, base = rels_name.rsplit("/_rels/", 1)
    part_name = base[: -len(".rels")]
    return f"{parent}/{part_name}" if parent else part_name


def _resolve_target(rels_name: str, target: str) -> str | None:
    if not target or "://" in target:
        return None
    target = target.split("?", 1)[0].replace("\\", "/")
    if target.startswith("/"):
        return target.lstrip("/")
    parent = rels_name.rsplit("/_rels/", 1)[0] if "/_rels/" in rels_name else ""
    joined = posixpath.normpath(posixpath.join(parent, target))
    if joined.startswith("../") or joined == "..":
        return None
    return joined


def _cleanup_package(package: _DocxPackage, redactor: Redactor) -> int:
    remaining: set[str] = set()
    for name in list(package.files):
        if not name.endswith(".rels"):
            continue
        root = _parse_xml(package.files[name])
        part = _part_of_rels(name)
        referenced: set[str] | None = None
        if part and part in package.files and part.endswith(".xml"):
            referenced = _referenced_ids(_parse_xml(package.files[part]))
        changed = False
        if referenced is not None:
            for rel in list(root):
                if rel.tag != f"{PKG_REL}Relationship":
                    continue
                rel_type = rel.get("Type") or ""
                mode = rel.get("TargetMode") or ""
                target = rel.get("Target") or ""
                if mode == "External" and rel_type.endswith("/hyperlink"):
                    if target and "redacted.invalid" not in target:
                        redactor.mapping.drop("链接", target)
                        rel.set("Target", "http://redacted.invalid")
                        changed = True
                    continue
                if rel.get("Id") in referenced:
                    continue
                if rel_type.endswith(_DROP_REL_SUFFIXES):
                    root.remove(rel)
                    changed = True
        if changed:
            package.files[name] = _serialize_xml(root)
        for rel in root:
            if rel.tag != f"{PKG_REL}Relationship":
                continue
            if (rel.get("TargetMode") or "") == "External":
                continue
            resolved = _resolve_target(name, rel.get("Target") or "")
            if resolved:
                remaining.add(resolved)

    removed_media = 0
    for name in list(package.files):
        if name.startswith("word/media/") or name.startswith("word/embeddings/"):
            if name not in remaining:
                del package.files[name]
                removed_media += 1

    content_types = "[Content_Types].xml"
    if content_types in package.files:
        root = _parse_xml(package.files[content_types])
        changed = False
        for element in list(root):
            if element.tag != f"{CT}Override":
                continue
            part_name = (element.get("PartName") or "").lstrip("/")
            if part_name and part_name not in package.files:
                root.remove(element)
                changed = True
        if changed:
            package.files[content_types] = _serialize_xml(root)
    return removed_media


class _DocxPackage:
    def __init__(self, files: dict[str, bytes]):
        self.files = files

    @classmethod
    def read(cls, path: Path) -> _DocxPackage:
        with zipfile.ZipFile(path) as archive:
            files = {info.filename: archive.read(info.filename) for info in archive.infolist() if not info.is_dir()}
        if "[Content_Types].xml" not in files or "word/document.xml" not in files:
            raise DocxError("不是有效的 docx")
        return cls(files)

    def names(self) -> list[str]:
        return list(self.files)

    def save(self, path: Path) -> None:
        def sort_key(name: str) -> tuple[int, str]:
            if name == "[Content_Types].xml":
                return (0, name)
            if name == "_rels/.rels":
                return (1, name)
            return (2, name)

        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name in sorted(self.files, key=sort_key):
                archive.writestr(name, self.files[name])
