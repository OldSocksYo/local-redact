"""读写 xlsx 压缩包：单元格、工作表名、批注、图片和缓存。

格子上的字仍交给 Redactor。共享字符串按「还被哪些格子用着」重建，
避免改一条字符串时把表头一起改掉，也避免没人引用的原文留在包里。
"""

from __future__ import annotations

import posixpath
import re
import zipfile
from copy import deepcopy
from pathlib import Path

from lxml import etree

from local_redact.fields import COLUMN_RULES, HEADER_WORDS, LABELS, Rule, normalize_label
from local_redact.patterns import mask_sensitive
from local_redact.textutil import Redactor

S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

_DROP_PREFIXES = (
    "xl/media/",
    "xl/embeddings/",
    "xl/printersettings/",
    "xl/activex/",
    "xl/pivotcache/",
    "xl/pivottables/",
    "xl/externallinks/",
)
_DROP_REL_SUFFIXES = (
    "/image",
    "/oleObject",
    "/package",
    "/printerSettings",
    "/vbaProject",
    "/pivotCacheDefinition",
    "/pivotCacheRecords",
    "/pivotTable",
    "/externalLink",
    "/externalLinkPath",
    "/control",
)
_MACRO_CONTENT_TYPES = {
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
    "application/vnd.ms-excel.template.macroEnabled.main+xml":
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
}
_SHARED_STRINGS_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
_SHARED_STRINGS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
_HF_TAGS = {"oddHeader", "evenHeader", "firstHeader", "oddFooter", "evenFooter", "firstFooter"}
_FORMULA_TAGS = {"f", "formula1", "formula2", "definedName"}
_CHART_NUMBER_CACHE = {"numCache", "numLit"}
_CHART_TEXT_CACHE = {"strCache", "strLit"}
_IDENTITY_ATTRS = {"author", "initials", "userid"}
_SENSITIVE_PROP_NAMES = {"company", "manager", "creator", "lastauthor", "owner", "operator", "lastmodifiedby"}
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
_REF = re.compile(r"\$?([A-Za-z]{1,3})\$?(\d+)")
_FORMULA_STRING = re.compile(r'"(?:[^"]|"")*"')


class XlsxError(Exception):
    pass


class _Cell:
    def __init__(
        self,
        element: etree._Element,
        row: int,
        col: int,
        text: str,
        kind: str,
        shared_index: int | None,
    ):
        self.element = element
        self.row = row
        self.col = col
        self.text = text
        self.kind = kind
        self.shared_index = shared_index
        self.span_cols = 1
        self.span_rows = 1
        self.edit_text: str | None = None
        self.edit_mode: str | None = None


class _Sheet:
    def __init__(self, name: str, element: etree._Element, path: str, root: etree._Element):
        self.name = name
        self.element = element
        self.path = path
        self.root = root
        self.cells: list[_Cell] = []


class _StringTable:
    """只保留仍被格子引用的原字符串，新文本另建条目。"""

    def __init__(self, originals: list[etree._Element]):
        self.originals = originals
        self.output: list[etree._Element] = []
        self.references = 0
        self._orig_map: dict[int, int] = {}
        self._text_map: dict[str, int] = {}

    def use_original(self, index: int) -> int:
        if index not in self._orig_map:
            self._orig_map[index] = len(self.output)
            self.output.append(deepcopy(self.originals[index]))
        self.references += 1
        return self._orig_map[index]

    def use_text(self, text: str) -> int:
        if text not in self._text_map:
            si = etree.Element(f"{{{S_NS}}}si")
            node = etree.SubElement(si, f"{{{S_NS}}}t")
            _write_text(node, text)
            self._text_map[text] = len(self.output)
            self.output.append(si)
        self.references += 1
        return self._text_map[text]


def redact_xlsx(src: Path, dest: Path, redactor: Redactor) -> dict[str, int]:
    if src.suffix.lower() == ".xls":
        raise XlsxError("这里只处理已经转好的 xlsx。")
    if src.resolve() == dest.resolve():
        raise XlsxError("不能覆盖原稿")
    try:
        package = _XlsxPackage.read(src)
    except zipfile.BadZipFile as exc:
        raise XlsxError("不是有效的 xlsx，文件可能已加密或已损坏。") from exc

    redactor.begin_file()
    workbook = _parse_xml(package.files["xl/workbook.xml"])
    strings, string_nodes = _load_shared_strings(package)
    sheets = _load_sheets(package, workbook, strings)
    for sheet in sheets:
        _redact_sheet_structure(sheet, redactor)
    # 先收齐栏目里的名称，再改公式字面量、其余格子和工作表名。
    _redact_formula_tree(workbook, redactor)
    for sheet in sheets:
        _redact_sheet_text(sheet, redactor)

    renames = _rename_sheets(sheets, redactor)
    _rename_formulas(workbook, renames)
    for sheet in sheets:
        _rename_formulas(sheet.root, renames)
        _redact_header_footer(sheet.root, redactor)
        _redact_hyperlink_labels(sheet.root, redactor)

    _redact_charts(package, redactor, renames)
    _redact_drawing_text(package, redactor)
    _redact_comments(package, redactor)
    _redact_persons(package, redactor)
    _redact_properties(package, redactor)
    _redact_loose_parts(package, redactor)

    table = _StringTable(string_nodes)
    for sheet in sheets:
        _commit_cells(sheet, table)
        package.files[sheet.path] = _serialize_xml(sheet.root)
    _write_shared_strings(package, table)
    package.files["xl/workbook.xml"] = _serialize_xml(workbook)

    shapes = _remove_pictures(package)
    media = _cleanup_package(package, redactor)
    dest.parent.mkdir(parents=True, exist_ok=True)
    package.save(dest)
    return {"shapes": shapes, "media": media, "fields": redactor.field_hits}


def _load_shared_strings(package: _XlsxPackage) -> tuple[list[str], list[etree._Element]]:
    path = _shared_strings_path(package)
    if path is None or path not in package.files:
        return [], []
    root = _parse_xml(package.files[path])
    nodes = [node for node in root if _local(node) == "si"]
    return [_si_text(node) for node in nodes], nodes


def _shared_strings_path(package: _XlsxPackage) -> str | None:
    rels = "xl/_rels/workbook.xml.rels"
    if rels in package.files:
        root = _parse_xml(package.files[rels])
        for rel in root:
            if _local(rel) != "Relationship":
                continue
            if (rel.get("Type") or "").endswith("/sharedStrings"):
                resolved = _resolve_target(rels, rel.get("Target") or "")
                if resolved:
                    return resolved
    if "xl/sharedStrings.xml" in package.files:
        return "xl/sharedStrings.xml"
    return None


def _load_sheets(
    package: _XlsxPackage,
    workbook: etree._Element,
    strings: list[str],
) -> list[_Sheet]:
    rels_name = "xl/_rels/workbook.xml.rels"
    rels = _parse_xml(package.files[rels_name]) if rels_name in package.files else etree.Element(f"{{{PKG_NS}}}Relationships")
    targets: dict[str, str] = {}
    for rel in rels:
        if _local(rel) != "Relationship":
            continue
        rel_id = rel.get("Id")
        if not rel_id:
            continue
        resolved = _resolve_target(rels_name, rel.get("Target") or "")
        if resolved:
            targets[rel_id] = resolved

    sheets: list[_Sheet] = []
    for element in workbook.iter():
        if _local(element) != "sheet":
            continue
        name = element.get("name") or ""
        rel_id = _rel_attr(element)
        path = targets.get(rel_id or "")
        if not path or path not in package.files or not path.endswith(".xml"):
            continue
        root = _parse_xml(package.files[path])
        sheet = _Sheet(name, element, path, root)
        sheet.cells = _read_cells(root, strings)
        sheets.append(sheet)
    return sheets


def _read_cells(root: etree._Element, strings: list[str]) -> list[_Cell]:
    merges = _merge_spans(root)
    covered: set[tuple[int, int]] = set()
    for (row, col), (span_cols, span_rows) in merges.items():
        for r in range(row, row + span_rows):
            for c in range(col, col + span_cols):
                if (r, c) != (row, col):
                    covered.add((r, c))

    cells: list[_Cell] = []
    for row_el in root.iter():
        if _local(row_el) != "row":
            continue
        next_col = 0
        for cell_el in row_el:
            if _local(cell_el) != "c":
                continue
            ref = _split_ref(cell_el.get("r") or "")
            if ref is None:
                raw_row = row_el.get("r") or ""
                if not raw_row.isdigit():
                    continue
                row, col = int(raw_row), next_col
            else:
                row, col = ref
            next_col = col + 1
            if (row, col) in covered:
                continue
            text, kind, index = _cell_text(cell_el, strings)
            cell = _Cell(cell_el, row, col, text, kind, index)
            span = merges.get((row, col))
            if span is not None:
                cell.span_cols, cell.span_rows = span
            cells.append(cell)
    return cells


def _merge_spans(root: etree._Element) -> dict[tuple[int, int], tuple[int, int]]:
    spans: dict[tuple[int, int], tuple[int, int]] = {}
    for element in root.iter():
        if _local(element) != "mergeCell":
            continue
        ref = element.get("ref") or ""
        start_text, _, end_text = ref.partition(":")
        start = _split_ref(start_text)
        end = _split_ref(end_text) if end_text else start
        if start is None or end is None:
            continue
        span_cols = end[1] - start[1] + 1
        span_rows = end[0] - start[0] + 1
        if span_cols >= 1 and span_rows >= 1:
            spans[start] = (span_cols, span_rows)
    return spans


def _cell_text(element: etree._Element, strings: list[str]) -> tuple[str, str, int | None]:
    kind = element.get("t") or "n"
    if kind == "s":
        raw = _direct_text(element, "v").strip()
        if raw.isdigit():
            index = int(raw)
            if 0 <= index < len(strings):
                return strings[index], "s", index
        return "", "s", None
    if kind == "inlineStr":
        return _inline_text(element), "inline", None
    if kind in {"str", "b", "e", "d"}:
        return _direct_text(element, "v"), kind, None
    return _direct_text(element, "v"), "n", None


def _redact_sheet_structure(sheet: _Sheet, redactor: Redactor) -> None:
    grouped: dict[int, list[_Cell]] = {}
    for cell in sheet.cells:
        grouped.setdefault(cell.row, []).append(cell)
    rows = [(row, grouped[row]) for row in sorted(grouped)]
    for cells in grouped.values():
        cells.sort(key=lambda item: item.col)

    header_at = [index for index, (_, cells) in enumerate(rows) if _is_header_row(cells)]
    for position, start in enumerate(header_at):
        end = header_at[position + 1] if position + 1 < len(header_at) else len(rows)
        columns: dict[int, Rule] = {}
        for cell in rows[start][1]:
            key = normalize_label(cell.text)
            if key not in COLUMN_RULES:
                continue
            for col in range(cell.col, cell.col + cell.span_cols):
                columns[col] = COLUMN_RULES[key]
        for index in range(start + 1, end):
            cells = rows[index][1]
            if _is_note_row(cells):
                continue
            for col, rule in columns.items():
                target = _covering(cells, col)
                if target is not None:
                    _replace_cell(target, redactor, rule)

    header_set = set(header_at)
    for index, (_, cells) in enumerate(rows):
        if index in header_set:
            continue
        for cell in cells:
            _redact_label(rows, index, cell, redactor)


def _is_header_row(cells: list[_Cell]) -> bool:
    labels = [normalize_label(cell.text.strip()) for cell in cells if cell.text.strip()]
    return len(labels) >= 2 and all(label in HEADER_WORDS for label in labels)


def _is_note_row(cells: list[_Cell]) -> bool:
    for cell in cells:
        text = cell.text.strip()
        if text.startswith(("备注", "注：", "注:")):
            return True
    return False


def _covering(cells: list[_Cell], col: int) -> _Cell | None:
    for cell in cells:
        if cell.col <= col < cell.col + cell.span_cols:
            return cell
    return None


def _replace_cell(cell: _Cell, redactor: Redactor, rule: Rule) -> None:
    if cell.edit_text is not None:
        return
    raw = cell.text.strip()
    if not raw:
        return
    key = normalize_label(raw)
    if key in LABELS or key in HEADER_WORDS:
        return
    updated = redactor.apply_rule(raw, rule)
    if updated != raw:
        cell.edit_text = updated
        cell.edit_mode = "value"


def _usable_target(target: _Cell | None, source: _Cell) -> bool:
    if target is None or target is source or target.edit_text is not None:
        return False
    text = target.text.strip()
    if not text:
        return False
    key = normalize_label(text)
    return key not in LABELS and key not in HEADER_WORDS


def _redact_label(
    rows: list[tuple[int, list[_Cell]]],
    index: int,
    cell: _Cell,
    redactor: Redactor,
) -> None:
    key = normalize_label(cell.text.strip())
    if key not in LABELS:
        return
    rule = LABELS[key]
    target = _covering(rows[index][1], cell.col + cell.span_cols)
    if not _usable_target(target, cell):
        if index + 1 >= len(rows):
            return
        target = _covering(rows[index + 1][1], cell.col)
        if not _usable_target(target, cell):
            return
    _replace_cell(target, redactor, rule)


def _redact_sheet_text(sheet: _Sheet, redactor: Redactor) -> None:
    _redact_formula_tree(sheet.root, redactor)
    for cell in sheet.cells:
        if cell.edit_mode == "value":
            continue
        if cell.edit_text is not None or cell.kind in {"b", "e"}:
            continue
        original = cell.text
        if not original:
            continue
        updated = redactor.redact_string(original)
        if updated == original:
            continue
        cell.edit_text = updated
        cell.edit_mode = "cache" if cell.kind == "str" else "value"


def _rename_sheets(sheets: list[_Sheet], redactor: Redactor) -> list[tuple[str, str]]:
    used: set[str] = set()
    pairs: list[tuple[str, str]] = []
    for sheet in sheets:
        renamed = redactor.redact_string(sheet.name)
        cleaned = _unique_sheet_name(renamed, used)
        if cleaned != sheet.name:
            pairs.append((sheet.name, cleaned))
        sheet.element.set("name", cleaned)
        sheet.name = cleaned
    return pairs


def _unique_sheet_name(name: str, used: set[str]) -> str:
    cleaned = re.sub(r"[\\/?*\[\]:]", "", name).strip().strip("'")
    cleaned = cleaned[:31] or "工作表"
    candidate = cleaned
    number = 2
    while candidate.casefold() in used:
        suffix = f"_{number}"
        candidate = f"{cleaned[: 31 - len(suffix)]}{suffix}"
        number += 1
    used.add(candidate.casefold())
    return candidate


def _rename_formulas(root: etree._Element, pairs: list[tuple[str, str]]) -> None:
    if not pairs:
        return
    for element in root.iter():
        local = _local(element)
        if local in _FORMULA_TAGS and element.text and "!" in element.text:
            updated = _rename_sheets_in_formula(element.text, pairs)
            if updated != element.text:
                element.text = updated
        if local != "hyperlink":
            continue
        attr = _attr_name(element, "location")
        if attr is None:
            continue
        old = element.get(attr) or ""
        if "!" not in old:
            continue
        updated = _rename_sheets_in_formula(old, pairs)
        if updated != old:
            element.set(attr, updated)


def _rename_sheets_in_formula(formula: str, pairs: list[tuple[str, str]]) -> str:
    mapping = {old: new for old, new in pairs if old != new}
    if not mapping or "!" not in formula:
        return formula

    def rename_code(code: str) -> str:
        def repl_quoted(match: re.Match[str]) -> str:
            name = match.group(1).replace("''", "'")
            new = mapping.get(name)
            if new is None:
                return match.group(0)
            return "'" + new.replace("'", "''") + "'"

        updated = re.sub(r"'((?:[^']|'')*)'(?=!)", repl_quoted, code)
        for old, new in sorted(mapping.items(), key=lambda item: len(item[0]), reverse=True):
            if re.fullmatch(r"[A-Za-z0-9_\u4e00-\u9fff]+", old) is None:
                continue
            pattern = rf"(?<![A-Za-z0-9_\u4e00-\u9fff']){re.escape(old)}(?=!)"
            updated = re.sub(pattern, lambda _match, replacement=new: replacement, updated)
        return updated

    parts: list[str] = []
    last = 0
    for match in _FORMULA_STRING.finditer(formula):
        parts.append(rename_code(formula[last : match.start()]))
        parts.append(match.group(0))
        last = match.end()
    parts.append(rename_code(formula[last:]))
    return "".join(parts)


def _redact_header_footer(root: etree._Element, redactor: Redactor) -> None:
    for element in root.iter():
        if _local(element) not in _HF_TAGS or not (element.text or "").strip():
            continue
        updated = redactor.redact_string(element.text or "")
        if updated != element.text:
            element.text = updated


def _redact_hyperlink_labels(root: etree._Element, redactor: Redactor) -> None:
    for element in root.iter():
        if _local(element) != "hyperlink":
            continue
        for local_name in ("display", "tooltip"):
            attr = _attr_name(element, local_name)
            if attr is None:
                continue
            old = element.get(attr) or ""
            if not old.strip():
                continue
            updated = redactor.redact_string(old)
            if updated != old:
                element.set(attr, updated)


def _redact_charts(package: _XlsxPackage, redactor: Redactor, pairs: list[tuple[str, str]]) -> None:
    for name in list(package.files):
        if not name.lower().startswith("xl/charts/") or not name.endswith(".xml"):
            continue
        root = _parse_xml(package.files[name])
        for element in root.iter():
            local = _local(element)
            if local == "t" and element.text:
                updated = redactor.redact_string(element.text)
                if updated != element.text:
                    _write_text(element, updated)
                continue
            if local != "v" or not (element.text or ""):
                continue
            cache = _nearest_local(element, _CHART_TEXT_CACHE | _CHART_NUMBER_CACHE)
            if cache in _CHART_TEXT_CACHE:
                updated = redactor.redact_string(element.text or "")
                if updated != element.text:
                    element.text = updated
            elif cache in _CHART_NUMBER_CACHE:
                updated = mask_sensitive(element.text or "", redactor.mapping.drop)
                if updated != element.text:
                    element.text = updated
        _redact_formula_tree(root, redactor)
        _rename_formulas(root, pairs)
        package.files[name] = _serialize_xml(root)


def _redact_drawing_text(package: _XlsxPackage, redactor: Redactor) -> None:
    for name in list(package.files):
        lower = name.lower()
        if not lower.endswith((".xml", ".vml")):
            continue
        if "/drawings/" not in lower and not lower.endswith(".vml"):
            continue
        root = _parse_xml(package.files[name])
        changed = False
        for element in root.iter():
            local = _local(element)
            if local == "t" and element.text:
                updated = redactor.redact_string(element.text)
                if updated != element.text:
                    _write_text(element, updated)
                    changed = True
            if local == "textpath":
                attr = _attr_name(element, "string")
                if attr is None:
                    continue
                old = element.get(attr) or ""
                updated = redactor.redact_string(old)
                if updated != old:
                    element.set(attr, updated)
                    changed = True
        if changed:
            package.files[name] = _serialize_xml(root)


def _redact_comments(package: _XlsxPackage, redactor: Redactor) -> None:
    for name in list(package.files):
        lower = name.lower()
        if not lower.endswith(".xml"):
            continue
        if "/comments" not in lower and "/threadedcomments/" not in lower:
            continue
        root = _parse_xml(package.files[name])
        for element in root.iter():
            local = _local(element)
            if local == "author" and (element.text or "").strip():
                redactor.mapping.drop("文档属性", element.text.strip())
                element.text = "已脱敏"
            elif local in {"comment", "threadedComment"}:
                _redact_comment_body(element, redactor)
        _redact_identity_attrs(root, redactor)
        package.files[name] = _serialize_xml(root)


def _redact_comment_body(element: etree._Element, redactor: Redactor) -> None:
    texts = [
        node
        for node in element.iter()
        if _local(node) == "t" and _nearest_local(node, {"rPh"}) is None
    ]
    if texts:
        combined = "".join(node.text or "" for node in texts)
        updated = redactor.redact_string(combined)
        if updated != combined:
            _write_text(texts[0], updated)
            for extra in texts[1:]:
                extra.text = ""
        return
    for node in element.iter():
        if _local(node) != "text" or not (node.text or "").strip():
            continue
        updated = redactor.redact_string(node.text or "")
        if updated != node.text:
            _write_text(node, updated)


def _redact_persons(package: _XlsxPackage, redactor: Redactor) -> None:
    for name in list(package.files):
        if "/persons/" not in name.lower() or not name.endswith(".xml"):
            continue
        root = _parse_xml(package.files[name])
        for element in root.iter():
            for attr in list(element.attrib):
                local = attr.rsplit("}", 1)[-1]
                value = element.get(attr) or ""
                if not value.strip():
                    continue
                if local == "displayName" and value != "已脱敏":
                    redactor.mapping.drop("文档属性", value.strip())
                    element.set(attr, "已脱敏")
                elif local in {"userId", "providerId"}:
                    updated = redactor.redact_string(value)
                    if updated != value:
                        element.set(attr, updated)
        package.files[name] = _serialize_xml(root)


def _redact_properties(package: _XlsxPackage, redactor: Redactor) -> None:
    core_name = "docProps/core.xml"
    if core_name in package.files:
        root = _parse_xml(package.files[core_name])
        for element in root.iter():
            local = _local(element).lower()
            if local in {"creator", "lastmodifiedby"} and (element.text or "").strip():
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
            if _local(element) in {"Company", "Manager"} and (element.text or "").strip():
                redactor.mapping.drop("文档属性", element.text.strip())
                element.text = ""
        _redact_identity_attrs(root, redactor)
        package.files[app_name] = _serialize_xml(root)


def _redact_loose_parts(package: _XlsxPackage, redactor: Redactor) -> None:
    for name in list(package.files):
        lower = name.lower()
        if not lower.endswith(".xml"):
            continue
        if name == "docProps/custom.xml" or lower.startswith("customxml/") or lower == "xl/connections.xml":
            root = _parse_xml(package.files[name])
            _redact_loose_xml(root, redactor)
            _redact_identity_attrs(root, redactor)
            package.files[name] = _serialize_xml(root)
        elif lower.startswith("xl/tables/") or lower.startswith("xl/querytables/"):
            root = _parse_xml(package.files[name])
            _redact_table_text(root, redactor)
            package.files[name] = _serialize_xml(root)


def _redact_table_text(root: etree._Element, redactor: Redactor) -> None:
    for element in root.iter():
        if element.text and element.text.strip():
            updated = redactor.redact_string(element.text)
            if updated != element.text:
                element.text = updated
        for local_name in ("name", "displayName", "uniqueName"):
            attr = _attr_name(element, local_name)
            if attr is None:
                continue
            old = element.get(attr) or ""
            if not old.strip():
                continue
            updated = redactor.redact_string(old)
            if updated != old:
                element.set(attr, updated)


def _redact_loose_xml(root: etree._Element, redactor: Redactor) -> None:
    cleared: set[int] = set()
    for element in root.iter():
        if id(element) in cleared:
            continue
        local = _local(element).lower()
        prop_name = _attr_value(element, "name").strip().lower()
        if local == "property" and prop_name in _SENSITIVE_PROP_NAMES:
            for node in element.iter():
                cleared.add(id(node))
                if node.text and node.text.strip():
                    redactor.mapping.drop("文档属性", node.text.strip())
                    node.text = ""
            continue
        if element.text and element.text.strip():
            updated = _scrub_value(element.text, redactor)
            if updated != element.text:
                element.text = updated
        for attr, value in list(element.attrib.items()):
            attr_name = attr.rsplit("}", 1)[-1].lower()
            if attr_name in _SKIP_VALUE_ATTRS or attr_name in _IDENTITY_ATTRS or not value.strip():
                continue
            updated = _scrub_value(value, redactor)
            if updated != value:
                element.set(attr, updated)


def _scrub_value(value: str, redactor: Redactor) -> str:
    if _is_path_text(value):
        redactor.mapping.drop("链接", value.strip())
        return ""
    return redactor.redact_string(value)


def _redact_identity_attrs(root: etree._Element, redactor: Redactor) -> None:
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


def _commit_cells(sheet: _Sheet, table: _StringTable) -> None:
    for cell in sheet.cells:
        element = cell.element
        if cell.edit_mode == "cache" and cell.edit_text is not None:
            _set_cached_text(element, cell.edit_text)
            continue
        if cell.edit_mode == "value" and cell.edit_text is not None:
            _set_shared_value(element, table.use_text(cell.edit_text))
            continue
        if cell.kind == "s" and cell.shared_index is not None:
            _set_shared_index(element, table.use_original(cell.shared_index))


def _set_cached_text(element: etree._Element, text: str) -> None:
    _drop_inline(element)
    _set_type(element, "str")
    node = _ensure_child(element, "v")
    _write_text(node, text)


def _set_shared_value(element: etree._Element, index: int) -> None:
    for child in list(element):
        if _local(child) in {"f", "is"}:
            element.remove(child)
    _set_shared_index(element, index)


def _set_shared_index(element: etree._Element, index: int) -> None:
    _set_type(element, "s")
    node = _ensure_child(element, "v")
    node.text = str(index)
    if XML_SPACE in node.attrib:
        del node.attrib[XML_SPACE]


def _set_type(element: etree._Element, value: str) -> None:
    for attr in list(element.attrib):
        if attr.rsplit("}", 1)[-1] == "t":
            del element.attrib[attr]
    element.set("t", value)


def _drop_inline(element: etree._Element) -> None:
    for child in list(element):
        if _local(child) == "is":
            element.remove(child)


def _ensure_child(element: etree._Element, local: str) -> etree._Element:
    found = _direct_child(element, local)
    if found is not None:
        return found
    return etree.SubElement(element, _qualified(element, local))


def _write_shared_strings(package: _XlsxPackage, table: _StringTable) -> None:
    path = _shared_strings_path(package) or "xl/sharedStrings.xml"
    if not table.output:
        if path in package.files:
            del package.files[path]
        _drop_rel_type(package, "xl/_rels/workbook.xml.rels", "/sharedStrings")
        return
    root = etree.Element(f"{{{S_NS}}}sst", nsmap={None: S_NS})
    root.set("count", str(max(table.references, len(table.output))))
    root.set("uniqueCount", str(len(table.output)))
    for node in table.output:
        root.append(node)
    package.files[path] = _serialize_xml(root)
    _ensure_shared_strings_rel(package, path)


def _ensure_shared_strings_rel(package: _XlsxPackage, path: str) -> None:
    rels_name = "xl/_rels/workbook.xml.rels"
    if rels_name not in package.files:
        root = etree.Element(f"{{{PKG_NS}}}Relationships", nsmap={None: PKG_NS})
    else:
        root = _parse_xml(package.files[rels_name])
    for rel in root:
        if _local(rel) == "Relationship" and (rel.get("Type") or "").endswith("/sharedStrings"):
            rel.set("Target", posixpath.relpath(path, "xl"))
            package.files[rels_name] = _serialize_xml(root)
            _ensure_content_override(package, path, _SHARED_STRINGS_TYPE)
            return
    ids = {rel.get("Id") for rel in root if _local(rel) == "Relationship"}
    number = 1
    while f"rId{number}" in ids:
        number += 1
    rel = etree.SubElement(root, _qualified(root, "Relationship"))
    rel.set("Id", f"rId{number}")
    rel.set("Type", _SHARED_STRINGS_REL)
    rel.set("Target", posixpath.relpath(path, "xl"))
    package.files[rels_name] = _serialize_xml(root)
    _ensure_content_override(package, path, _SHARED_STRINGS_TYPE)


def _drop_rel_type(package: _XlsxPackage, rels_name: str, type_suffix: str) -> None:
    if rels_name not in package.files:
        return
    root = _parse_xml(package.files[rels_name])
    changed = False
    for rel in list(root):
        if _local(rel) == "Relationship" and (rel.get("Type") or "").endswith(type_suffix):
            root.remove(rel)
            changed = True
    if changed:
        package.files[rels_name] = _serialize_xml(root)


def _remove_pictures(package: _XlsxPackage) -> int:
    removed = 0
    for name in list(package.files):
        if not _may_contain_pictures(name):
            continue
        try:
            root = _parse_xml(package.files[name])
        except etree.XMLSyntaxError:
            continue
        taken = 0
        for element in list(root.iter()):
            if _local(element) not in {"blip", "imagedata"}:
                continue
            target = _picture_ancestor(element)
            parent = target.getparent() if target is not None else None
            if target is None or parent is None:
                continue
            parent.remove(target)
            taken += 1
        if taken:
            package.files[name] = _serialize_xml(root)
            removed += taken
    return removed


def _may_contain_pictures(name: str) -> bool:
    lower = name.lower()
    if not lower.endswith((".xml", ".vml")):
        return False
    return "/drawings/" in lower or lower.endswith(".vml") or "/worksheets/" in lower or "/charts/" in lower


def _picture_ancestor(element: etree._Element) -> etree._Element | None:
    node: etree._Element | None = element
    pic: etree._Element | None = None
    while node is not None:
        local = _local(node)
        if local == "pic":
            pic = node
        if local in {"twoCellAnchor", "oneCellAnchor", "absoluteAnchor"}:
            if any(_local(child) == "chart" for child in node.iter()):
                return pic
            return node
        if local in {"shape", "pict"}:
            pic = node
        node = node.getparent()
    return pic


def _cleanup_package(package: _XlsxPackage, redactor: Redactor) -> int:
    drop_names = {name for name in package.files if _is_dropped_part(name)}
    for name in list(package.files):
        if not name.endswith(".rels"):
            continue
        root = _parse_xml(package.files[name])
        changed = _scrub_rels(root, name, drop_names, redactor)
        if changed:
            package.files[name] = _serialize_xml(root)

    if "xl/workbook.xml" in package.files:
        root = _parse_xml(package.files["xl/workbook.xml"])
        changed = False
        for element in list(root):
            if _local(element) in {"pivotCaches", "externalReferences"}:
                root.remove(element)
                changed = True
        if changed:
            package.files["xl/workbook.xml"] = _serialize_xml(root)

    removed = 0
    for name in list(package.files):
        if name not in drop_names and not _is_dropped_part(name):
            continue
        lower = name.lower()
        if lower.startswith(("xl/media/", "xl/embeddings/")) or (
            lower.startswith("docprops/") and "thumbnail" in lower
        ):
            removed += 1
        del package.files[name]
    _scrub_content_types(package)
    return removed


def _is_dropped_part(name: str) -> bool:
    lower = name.lower()
    if lower.startswith(_DROP_PREFIXES):
        return True
    if lower == "xl/vbaproject.bin" or lower.endswith("/vbaproject.bin"):
        return True
    leaf = lower.rsplit("/", 1)[-1]
    return lower.startswith("docprops/") and leaf.startswith("thumbnail")


def _scrub_rels(
    root: etree._Element,
    rels_name: str,
    drop_names: set[str],
    redactor: Redactor,
) -> bool:
    changed = False
    for rel in list(root):
        if _local(rel) != "Relationship":
            continue
        rel_type = rel.get("Type") or ""
        target = rel.get("Target") or ""
        mode = rel.get("TargetMode") or ""
        resolved = _resolve_target(rels_name, target)
        if rel_type.endswith("/hyperlink") and mode == "External":
            if target and "redacted.invalid" not in target:
                redactor.mapping.drop("链接", target)
                rel.set("Target", "http://redacted.invalid")
                changed = True
            continue
        drop = rel_type.endswith(_DROP_REL_SUFFIXES)
        if resolved is not None and (resolved in drop_names or _is_dropped_part(resolved)):
            drop = True
        if mode == "External" and _is_local_target(target):
            redactor.mapping.drop("链接", target)
            drop = True
        if not drop:
            continue
        root.remove(rel)
        changed = True
    return changed


def _scrub_content_types(package: _XlsxPackage) -> None:
    name = "[Content_Types].xml"
    if name not in package.files:
        return
    root = _parse_xml(package.files[name])
    changed = False
    has_bin = any(part.lower().endswith(".bin") for part in package.files)
    for element in list(root):
        local = _local(element)
        if local == "Override":
            part_name = (element.get("PartName") or "").lstrip("/")
            content_type = element.get("ContentType") or ""
            if part_name and part_name not in package.files:
                root.remove(element)
                changed = True
                continue
            replacement = _MACRO_CONTENT_TYPES.get(content_type)
            if replacement:
                element.set("ContentType", replacement)
                changed = True
        elif local == "Default":
            content_type = (element.get("ContentType") or "").lower()
            extension = (element.get("Extension") or "").lower()
            if "vbaproject" in content_type or (extension == "bin" and not has_bin):
                root.remove(element)
                changed = True
    if changed:
        package.files[name] = _serialize_xml(root)


def _ensure_content_override(package: _XlsxPackage, part: str, content_type: str) -> None:
    name = "[Content_Types].xml"
    if name not in package.files:
        return
    root = _parse_xml(package.files[name])
    part_name = "/" + part.lstrip("/")
    for element in root:
        if _local(element) == "Override" and (element.get("PartName") or "") == part_name:
            element.set("ContentType", content_type)
            package.files[name] = _serialize_xml(root)
            return
    element = etree.SubElement(root, _qualified(root, "Override"))
    element.set("PartName", part_name)
    element.set("ContentType", content_type)
    package.files[name] = _serialize_xml(root)


def _redact_formula_tree(root: etree._Element, redactor: Redactor) -> None:
    for element in root.iter():
        if _local(element) not in _FORMULA_TAGS or not element.text:
            continue
        updated = _redact_formula_literals(element.text, redactor)
        # 数据有效性里的名单常常不带引号，也不是单元格公式。
        if (
            _local(element) in {"formula1", "formula2"}
            and updated == element.text
            and "!" not in element.text
            and not element.text.startswith("=")
        ):
            updated = redactor.redact_string(element.text)
        if updated != element.text:
            element.text = updated


def _redact_formula_literals(formula: str, redactor: Redactor) -> str:
    def repl(match: re.Match[str]) -> str:
        raw = match.group(0)
        inner = raw[1:-1].replace('""', '"')
        updated = redactor.redact_string(inner)
        if updated == inner:
            return raw
        return '"' + updated.replace('"', '""') + '"'

    return _FORMULA_STRING.sub(repl, formula)


def _si_text(node: etree._Element) -> str:
    parts = []
    for element in node.iter():
        if _local(element) != "t" or _nearest_local(element, {"rPh"}) is not None:
            continue
        parts.append(element.text or "")
    return "".join(parts)


def _inline_text(element: etree._Element) -> str:
    inline = _direct_child(element, "is")
    if inline is None:
        return ""
    return _si_text(inline)


def _direct_text(element: etree._Element, local: str) -> str:
    child = _direct_child(element, local)
    return "" if child is None or child.text is None else child.text


def _direct_child(element: etree._Element, local: str) -> etree._Element | None:
    for child in element:
        if _local(child) == local:
            return child
    return None


def _split_ref(ref: str) -> tuple[int, int] | None:
    match = _REF.fullmatch(ref.strip())
    if match is None:
        return None
    col = 0
    for char in match.group(1).upper():
        col = col * 26 + (ord(char) - 64)
    return int(match.group(2)), col - 1


def _local(element: etree._Element) -> str:
    tag = element.tag
    if not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


def _qualified(element: etree._Element, local: str) -> str:
    tag = element.tag
    if isinstance(tag, str) and tag.startswith("{"):
        namespace = tag[1:].split("}", 1)[0]
        return f"{{{namespace}}}{local}"
    return local


def _nearest_local(element: etree._Element, names: set[str]) -> str | None:
    parent = element.getparent()
    while parent is not None:
        local = _local(parent)
        if local in names:
            return local
        parent = parent.getparent()
    return None


def _attr_name(element: etree._Element, local_name: str) -> str | None:
    for attr in element.attrib:
        if attr.rsplit("}", 1)[-1] == local_name:
            return attr
    return None


def _attr_value(element: etree._Element, local_name: str) -> str:
    attr = _attr_name(element, local_name)
    if attr is None:
        return ""
    return element.get(attr) or ""


def _rel_attr(element: etree._Element) -> str:
    for attr, value in element.attrib.items():
        local = attr.rsplit("}", 1)[-1]
        if local == "id" and (attr == "id" or "relationships" in attr):
            return value
    return ""


def _write_text(element: etree._Element, text: str) -> None:
    element.text = text
    if text[:1].isspace() or text[-1:].isspace():
        element.set(XML_SPACE, "preserve")
    elif XML_SPACE in element.attrib:
        del element.attrib[XML_SPACE]


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


def _parse_xml(data: bytes) -> etree._Element:
    if data.startswith(b"\xef\xbb\xbf"):
        data = data[3:]
    return etree.fromstring(data)


def _serialize_xml(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


class _XlsxPackage:
    def __init__(self, files: dict[str, bytes]):
        self.files = files

    @classmethod
    def read(cls, path: Path) -> _XlsxPackage:
        with zipfile.ZipFile(path) as archive:
            files = {info.filename: archive.read(info.filename) for info in archive.infolist() if not info.is_dir()}
        if "[Content_Types].xml" not in files or "xl/workbook.xml" not in files:
            raise XlsxError("不是有效的 xlsx")
        return cls(files)

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
