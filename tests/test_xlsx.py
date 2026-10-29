import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from local_redact.cli import main
from local_redact.mapping import MappingBook
from local_redact.names import redact_relative
from local_redact.textutil import Redactor
from local_redact.xlsx_redact import XlsxError, redact_xlsx

S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"

COMPANY = "杭州示例公司"
SHEET = f"{COMPANY}原始记录"
SECRET = b"0123456789abcdef"

STRINGS = [
    COMPANY,
    "委托单位",
    "电话",
    "地址",
    "文三路1号",
    "序号",
    "样品编号",
    "检测项目",
    "检测结果",
    "单位",
    "采样人员",
    "东坐标",
    "北坐标",
    "方法依据",
    "S2024-001",
    "化学需氧量",
    "mg/L",
    "李伟",
    "HJ 828-2017",
    "S2024-002",
    "pH",
    "无量纲",
    "张超",
    "GB 3838-2002",
    "备注：晴",
    "本项目由示例环保承担",
    "138 0013 8000",
]


def q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def _xml(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _shared_cell(ref: str, index: int) -> etree._Element:
    cell = etree.Element(q(S, "c"), r=ref, t="s")
    etree.SubElement(cell, q(S, "v")).text = str(index)
    return cell


def _number_cell(ref: str, value: str, formula: str | None = None) -> etree._Element:
    cell = etree.Element(q(S, "c"), r=ref)
    if formula:
        etree.SubElement(cell, q(S, "f")).text = formula
    etree.SubElement(cell, q(S, "v")).text = value
    return cell


def _text_formula(ref: str, formula: str, cached: str) -> etree._Element:
    cell = etree.Element(q(S, "c"), r=ref, t="str")
    etree.SubElement(cell, q(S, "f")).text = formula
    etree.SubElement(cell, q(S, "v")).text = cached
    return cell


def _inline_cell(ref: str, text: str) -> etree._Element:
    cell = etree.Element(q(S, "c"), r=ref, t="inlineStr")
    inline = etree.SubElement(cell, q(S, "is"))
    etree.SubElement(inline, q(S, "t")).text = text
    return cell


def _row(number: int, *cells: etree._Element) -> etree._Element:
    row = etree.Element(q(S, "row"), r=str(number))
    for cell in cells:
        row.append(cell)
    return row


def _relationship(rel_id: str, rel_type: str, target: str, external: bool = False) -> etree._Element:
    rel = etree.Element(q(PKG, "Relationship"), Id=rel_id, Type=rel_type, Target=target)
    if external:
        rel.set("TargetMode", "External")
    return rel


def _rels(*items: etree._Element) -> bytes:
    root = etree.Element(q(PKG, "Relationships"), nsmap={None: PKG})
    for item in items:
        root.append(item)
    return _xml(root)


def build_workbook(path: Path) -> None:
    sst = etree.Element(q(S, "sst"), nsmap={None: S}, count=str(len(STRINGS)), uniqueCount=str(len(STRINGS)))
    for text in STRINGS:
        si = etree.SubElement(sst, q(S, "si"))
        etree.SubElement(si, q(S, "t")).text = text

    sheet = etree.Element(q(S, "worksheet"), nsmap={None: S, "r": R})
    etree.SubElement(sheet, q(S, "headerFooter")).append(
        etree.Element(q(S, "oddHeader"))
    )
    sheet.find(q(S, "headerFooter")).find(q(S, "oddHeader")).text = f"&C{COMPANY}"
    data = etree.SubElement(sheet, q(S, "sheetData"))
    data.append(_row(1, _shared_cell("A1", 0)))
    data.append(_row(2, _shared_cell("A2", 1), _shared_cell("B2", 0)))
    data.append(_row(3, _shared_cell("A3", 2), _shared_cell("B3", 26)))
    data.append(_row(4, _shared_cell("A4", 3), _shared_cell("B4", 4)))
    data.append(_row(5, _shared_cell("A5", 25)))
    data.append(_row(6, _inline_cell("A6", "1#排气筒")))
    data.append(
        _row(
            7,
            _shared_cell("A7", 5),
            _shared_cell("B7", 6),
            _shared_cell("C7", 7),
            _shared_cell("D7", 8),
            _shared_cell("E7", 9),
            _shared_cell("F7", 10),
            _shared_cell("G7", 11),
            _shared_cell("H7", 12),
            _shared_cell("I7", 13),
        )
    )
    data.append(
        _row(
            8,
            _number_cell("A8", "1"),
            _shared_cell("B8", 14),
            _shared_cell("C8", 15),
            _number_cell("D8", "35.2", "10+25.2"),
            _shared_cell("E8", 16),
            _shared_cell("F8", 17),
            _number_cell("G8", "405678.12"),
            _number_cell("H8", "345678.91"),
            _shared_cell("I8", 18),
        )
    )
    data.append(
        _row(
            9,
            _number_cell("A9", "2"),
            _shared_cell("B9", 19),
            _shared_cell("C9", 20),
            _number_cell("D9", "7.26"),
            _shared_cell("E9", 21),
            _shared_cell("F9", 22),
            _number_cell("G9", "405679.12"),
            _number_cell("H9", "345679.91"),
            _shared_cell("I9", 23),
        )
    )
    data.append(_row(10, _shared_cell("A10", 24)))
    data.append(
        _row(
            11,
            _text_formula("AA11", f'="联系{COMPANY}"', f"联系{COMPANY}"),
            _number_cell("AB11", "35.2", f"'{SHEET}'!D8"),
        )
    )
    merges = etree.SubElement(sheet, q(S, "mergeCells"), count="1")
    etree.SubElement(merges, q(S, "mergeCell"), ref="A1:C1")
    validations = etree.SubElement(sheet, q(S, "dataValidations"), count="1")
    validation = etree.SubElement(validations, q(S, "dataValidation"), type="list", sqref="A20")
    etree.SubElement(validation, q(S, "formula1")).text = f'"{COMPANY}"'
    links = etree.SubElement(sheet, q(S, "hyperlinks"))
    link = etree.SubElement(links, q(S, "hyperlink"), ref="A1", display=COMPANY)
    link.set(q(R, "id"), "rId5")

    hidden = etree.Element(q(S, "worksheet"), nsmap={None: S})
    hidden_data = etree.SubElement(hidden, q(S, "sheetData"))
    hidden_data.append(_row(1, _number_cell("A1", "13800138000")))

    workbook = etree.Element(q(S, "workbook"), nsmap={None: S, "r": R})
    sheets = etree.SubElement(workbook, q(S, "sheets"))
    first = etree.SubElement(sheets, q(S, "sheet"), name=SHEET, sheetId="1")
    first.set(q(R, "id"), "rId1")
    second = etree.SubElement(sheets, q(S, "sheet"), name="内部", sheetId="2", state="hidden")
    second.set(q(R, "id"), "rId2")
    names = etree.SubElement(workbook, q(S, "definedNames"))
    defined = etree.SubElement(names, q(S, "definedName"), name="客户区")
    defined.text = f"{SHEET}!$A$1"
    caches = etree.SubElement(workbook, q(S, "pivotCaches"))
    cache = etree.SubElement(caches, q(S, "pivotCache"), cacheId="1")
    cache.set(q(R, "id"), "rId9")
    external_refs = etree.SubElement(workbook, q(S, "externalReferences"))
    external_ref = etree.SubElement(external_refs, q(S, "externalReference"))
    external_ref.set(q(R, "id"), "rId8")

    comments = etree.Element(q(S, "comments"), nsmap={None: S})
    authors = etree.SubElement(comments, q(S, "authors"))
    etree.SubElement(authors, q(S, "author")).text = "ONLYAUTHOR"
    comment = etree.SubElement(etree.SubElement(comments, q(S, "commentList")), q(S, "comment"), ref="A2", authorId="0")
    text = etree.SubElement(comment, q(S, "text"))
    run = etree.SubElement(text, q(S, "r"))
    etree.SubElement(run, q(S, "t")).text = "请联系"
    run = etree.SubElement(text, q(S, "r"))
    etree.SubElement(run, q(S, "t")).text = COMPANY

    persons = etree.Element(q(S, "personList"), nsmap={None: S})
    etree.SubElement(persons, q(S, "person"), displayName="ONLYPERSON", id="{1}", userId="only@example.com")

    chart = etree.Element(q(CHART, "chartSpace"), nsmap={None: CHART, "a": A, "c": CHART})
    title = etree.SubElement(etree.SubElement(chart, q(CHART, "title")), q(CHART, "tx"))
    rich = etree.SubElement(title, q(CHART, "rich"))
    paragraph = etree.SubElement(rich, q(A, "p"))
    etree.SubElement(etree.SubElement(paragraph, q(A, "r")), q(A, "t")).text = f"{COMPANY}排放图"
    series = etree.SubElement(chart, q(CHART, "ser"))
    num_cache = etree.SubElement(etree.SubElement(etree.SubElement(series, q(CHART, "val")), q(CHART, "numRef")), q(CHART, "numCache"))
    etree.SubElement(etree.SubElement(num_cache, q(CHART, "pt"), idx="0"), q(CHART, "v")).text = "35.2"
    str_cache = etree.SubElement(etree.SubElement(etree.SubElement(series, q(CHART, "cat")), q(CHART, "strRef")), q(CHART, "strCache"))
    etree.SubElement(etree.SubElement(str_cache, q(CHART, "pt"), idx="0"), q(CHART, "v")).text = COMPANY
    etree.SubElement(series, q(CHART, "f")).text = f"'{SHEET}'!$D$8"

    drawing = etree.Element(q(XDR, "wsDr"), nsmap={None: XDR, "a": A, "r": R})
    picture_anchor = etree.SubElement(drawing, q(XDR, "twoCellAnchor"))
    etree.SubElement(etree.SubElement(etree.SubElement(picture_anchor, q(XDR, "pic")), q(XDR, "blipFill")), q(A, "blip")).set(
        q(R, "embed"), "rId1"
    )
    text_anchor = etree.SubElement(drawing, q(XDR, "twoCellAnchor"))
    text_body = etree.SubElement(etree.SubElement(text_anchor, q(XDR, "sp")), q(XDR, "txBody"))
    etree.SubElement(etree.SubElement(etree.SubElement(text_body, q(A, "p")), q(A, "r")), q(A, "t")).text = COMPANY

    core = etree.Element(q(CP, "coreProperties"), nsmap={"cp": CP, "dc": DC})
    etree.SubElement(core, q(DC, "creator")).text = "ONLYCREATOR"
    etree.SubElement(core, q(CP, "lastModifiedBy")).text = "ONLYMODIFIED"
    app = etree.Element(q(EP, "Properties"), nsmap={None: EP})
    etree.SubElement(app, q(EP, "Company")).text = "ONLYPROPCOMPANY"
    etree.SubElement(app, q(EP, "Manager")).text = "ONLYMANAGER"

    content_types = etree.Element(q(CT, "Types"), nsmap={None: CT})
    etree.SubElement(content_types, q(CT, "Default"), Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
    etree.SubElement(content_types, q(CT, "Default"), Extension="xml", ContentType="application/xml")
    etree.SubElement(content_types, q(CT, "Default"), Extension="bin", ContentType="application/vnd.ms-office.vbaProject")
    overrides = {
        "/xl/workbook.xml": "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        "/xl/worksheets/sheet1.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        "/xl/worksheets/sheet2.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        "/xl/sharedStrings.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
        "/xl/comments1.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
        "/xl/charts/chart1.xml": "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
        "/xl/drawings/drawing1.xml": "application/vnd.openxmlformats-officedocument.drawing+xml",
        "/xl/vbaProject.bin": "application/vnd.ms-office.vbaProject",
        "/docProps/core.xml": "application/vnd.openxmlformats-package.core-properties+xml",
        "/docProps/app.xml": "application/vnd.openxmlformats-officedocument.extended-properties+xml",
        "/xl/pivotCache/pivotCacheRecords1.xml": "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml",
    }
    for part, content_type in overrides.items():
        etree.SubElement(content_types, q(CT, "Override"), PartName=part, ContentType=content_type)

    office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    package = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/"
    files = {
        "[Content_Types].xml": _xml(content_types),
        "_rels/.rels": _rels(
            _relationship("rId1", f"{office}officeDocument", "xl/workbook.xml"),
            _relationship("rId2", f"{package}core-properties", "docProps/core.xml"),
            _relationship("rId3", f"{office}extended-properties", "docProps/app.xml"),
            _relationship("rId4", f"{office}thumbnail", "docProps/thumbnail.jpeg"),
        ),
        "xl/workbook.xml": _xml(workbook),
        "xl/_rels/workbook.xml.rels": _rels(
            _relationship("rId1", f"{office}worksheet", "worksheets/sheet1.xml"),
            _relationship("rId2", f"{office}worksheet", "worksheets/sheet2.xml"),
            _relationship("rId3", f"{office}sharedStrings", "sharedStrings.xml"),
            _relationship("rId8", f"{office}externalLink", "externalLinks/externalLink1.xml"),
            _relationship("rId9", f"{office}pivotCacheDefinition", "pivotCache/pivotCacheDefinition1.xml"),
        ),
        "xl/sharedStrings.xml": _xml(sst),
        "xl/worksheets/sheet1.xml": _xml(sheet),
        "xl/worksheets/sheet2.xml": _xml(hidden),
        "xl/worksheets/_rels/sheet1.xml.rels": _rels(
            _relationship("rId1", f"{office}drawing", "../drawings/drawing1.xml"),
            _relationship("rId2", f"{office}comments", "../comments1.xml"),
            _relationship("rId5", f"{office}hyperlink", "file:///D:/客户/site", external=True),
        ),
        "xl/comments1.xml": _xml(comments),
        "xl/persons/person.xml": _xml(persons),
        "xl/charts/chart1.xml": _xml(chart),
        "xl/drawings/drawing1.xml": _xml(drawing),
        "xl/drawings/_rels/drawing1.xml.rels": _rels(
            _relationship("rId1", f"{office}image", "../media/photo.png"),
        ),
        "xl/media/photo.png": b"ONLYMEDIA",
        "xl/vbaProject.bin": b"ONLYMACRO",
        "xl/pivotCache/pivotCacheRecords1.xml": f'<s v="ONLYPIVOT{COMPANY}"/>'.encode("utf-8"),
        "xl/externalLinks/externalLink1.xml": (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            "<file>ONLYEXTERNALPATH D:\\秘密\\客户.xlsx</file></externalLink>"
        ).encode("utf-8"),
        "docProps/core.xml": _xml(core),
        "docProps/app.xml": _xml(app),
        "docProps/thumbnail.jpeg": b"ONLYTHUMB",
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as archive:
        for name, payload in files.items():
            archive.writestr(name, payload)


def package_blob(path: Path) -> bytes:
    with zipfile.ZipFile(path) as archive:
        return b"\n".join(archive.read(name) for name in archive.namelist())


class XlsxRedactTests(unittest.TestCase):
    def test_redacts_sheet_keeps_results_and_strips_leaks(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            src = root / "记录.xlsx"
            dest = root / "out.xlsx"
            build_workbook(src)
            book = MappingBook(root / "map.json")
            redactor = Redactor([("示例环保", None)], SECRET, book)
            stats = redact_xlsx(src, dest, redactor)

            blob = package_blob(dest)
            self.assertGreater(stats["fields"], 0)
            self.assertGreater(stats["shapes"], 0)
            self.assertGreater(stats["media"], 0)
            self.assertNotIn(COMPANY.encode(), blob)
            self.assertNotIn("示例环保".encode(), blob)
            self.assertNotIn("文三路1号".encode(), blob)
            self.assertNotIn("李伟".encode(), blob)
            self.assertNotIn("张超".encode(), blob)
            self.assertNotIn("405678.12".encode(), blob)
            self.assertNotIn("345678.91".encode(), blob)
            self.assertNotIn("13800138000".encode(), blob)
            self.assertNotIn("138 0013 8000".encode(), blob)
            self.assertNotIn("ONLYAUTHOR".encode(), blob)
            self.assertNotIn("ONLYPERSON".encode(), blob)
            self.assertNotIn("only@example.com".encode(), blob)
            self.assertNotIn("ONLYCREATOR".encode(), blob)
            self.assertNotIn("ONLYMODIFIED".encode(), blob)
            self.assertNotIn("ONLYPROPCOMPANY".encode(), blob)
            self.assertNotIn("ONLYMANAGER".encode(), blob)
            self.assertNotIn(b"ONLYMEDIA", blob)
            self.assertNotIn(b"ONLYMACRO", blob)
            self.assertNotIn(b"ONLYTHUMB", blob)
            self.assertNotIn("ONLYPIVOT".encode(), blob)
            self.assertNotIn("ONLYEXTERNALPATH".encode(), blob)
            self.assertNotIn("file:///D:/客户/site".encode(), blob)
            self.assertIn("35.2".encode(), blob)
            self.assertIn("7.26".encode(), blob)
            self.assertIn("10+25.2".encode(), blob)
            self.assertIn("S2024-001".encode(), blob)
            self.assertIn("化学需氧量".encode(), blob)
            self.assertIn("mg/L".encode(), blob)
            self.assertIn("HJ 828-2017".encode(), blob)
            self.assertIn("GB 3838-2002".encode(), blob)
            self.assertIn("采样人员".encode(), blob)
            self.assertIn("1#排气筒".encode(), blob)
            self.assertIn("备注：晴".encode(), blob)
            self.assertIn("[坐标已移除]".encode(), blob)
            self.assertIn("[电话已移除]".encode(), blob)
            self.assertIn("[地址已移除]".encode(), blob)
            self.assertIn("已脱敏".encode(), blob)
            token = next(item for item, original in book.tokens.items() if original == COMPANY)
            self.assertIn(token.encode(), blob)

            with zipfile.ZipFile(dest) as archive:
                names = set(archive.namelist())
                self.assertFalse(any(name.startswith("xl/media/") for name in names))
                self.assertFalse(any(name.startswith("xl/pivotCache/") for name in names))
                self.assertFalse(any(name.startswith("xl/externalLinks/") for name in names))
                self.assertNotIn("xl/vbaProject.bin", names)
                self.assertNotIn("docProps/thumbnail.jpeg", names)
                workbook = etree.fromstring(archive.read("xl/workbook.xml"))
                self.assertIsNone(workbook.find(q(S, "pivotCaches")))
                self.assertIsNone(workbook.find(q(S, "externalReferences")))
                sheet_names = [node.get("name") for node in workbook.iter(q(S, "sheet"))]
                self.assertIn("内部", sheet_names)
                self.assertTrue(all(COMPANY not in (name or "") for name in sheet_names))
                types = etree.fromstring(archive.read("[Content_Types].xml"))
                content_types = " ".join(node.get("ContentType") or "" for node in types)
                self.assertNotIn("macroEnabled", content_types)
                self.assertNotIn("vbaProject", content_types)
                self.assertIn("spreadsheetml.sheet.main+xml", content_types)

    def test_output_suffix(self) -> None:
        book = MappingBook(Path("unused-map.json"))
        plain = Redactor([], SECRET, book)
        self.assertEqual(redact_relative(Path("检测报告.doc"), plain), Path("检测报告.docx"))
        self.assertEqual(redact_relative(Path("记录.xlsx"), plain), Path("记录.xlsx"))
        self.assertEqual(redact_relative(Path("记录.xlsm"), plain), Path("记录.xlsx"))
        self.assertEqual(redact_relative(Path("记录.xls"), plain), Path("记录.xlsx"))

    def test_rejects_legacy_xls_bytes(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            src = root / "记录.xls"
            src.write_bytes(b"not-a-zip")
            book = MappingBook(root / "map.json")
            with self.assertRaises(XlsxError):
                redact_xlsx(src, root / "out.xlsx", Redactor([], SECRET, book))

    def test_cli_redacts_xlsx_name_and_cells(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            src = root / "in"
            build_workbook(src / f"{COMPANY}原始记录.xlsx")
            lexicon = root / "lexicon.txt"
            lexicon.write_text("示例环保\n", encoding="utf-8")
            out = root / "out"
            code = main(
                [
                    str(src),
                    "-o",
                    str(out),
                    "--lexicon",
                    str(lexicon),
                    "--map",
                    str(root / "mapping.json"),
                    "--secret",
                    str(root / "secret.key"),
                ]
            )
            self.assertEqual(code, 0)
            outputs = list(out.rglob("*.xlsx"))
            self.assertEqual(len(outputs), 1)
            self.assertNotIn(COMPANY, outputs[0].name)
            self.assertNotIn("示例环保", outputs[0].name)
            blob = package_blob(outputs[0])
            self.assertNotIn(COMPANY.encode(), blob)
            self.assertIn("35.2".encode(), blob)
            mapping = __import__("json").loads((root / "mapping.json").read_text(encoding="utf-8"))
            stored = mapping["files"][f"{COMPANY}原始记录.xlsx"]
            self.assertTrue(stored.endswith(".xlsx"))
            self.assertNotIn(COMPANY, stored)

    def test_cli_converts_xls_then_redacts(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            template = root / "template.xlsx"
            build_workbook(template)
            src = root / "in"
            original = src / f"{COMPANY}.xls"
            original.parent.mkdir()
            original.write_bytes(b"dummy-xls")

            class FakeSheet:
                def start(self) -> None:
                    return None

                def convert(self, source: Path, dest: Path) -> None:
                    dest.write_bytes(template.read_bytes())

                def close(self) -> None:
                    return None

            lexicon = root / "lexicon.txt"
            lexicon.write_text("# 空词表\n", encoding="utf-8")
            out = root / "out"
            with patch("local_redact.cli.SheetSession", FakeSheet):
                code = main(
                    [
                        str(src),
                        "-o",
                        str(out),
                        "--lexicon",
                        str(lexicon),
                        "--map",
                        str(root / "mapping.json"),
                        "--secret",
                        str(root / "secret.key"),
                    ]
                )
            self.assertEqual(code, 0)
            self.assertEqual(original.read_bytes(), b"dummy-xls")
            outputs = list(out.rglob("*.xlsx"))
            self.assertEqual(len(outputs), 1)
            self.assertNotIn(COMPANY, outputs[0].name)
            self.assertNotIn(COMPANY.encode(), package_blob(outputs[0]))


if __name__ == "__main__":
    unittest.main()
