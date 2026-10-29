import json
import shutil
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from lxml import etree

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from local_redact.cli import main
from local_redact.docx_redact import DocxError, redact_docx
from local_redact.mapping import MappingBook
from local_redact.names import OutputNames, redact_relative
from local_redact.patterns import is_valid_id_card, is_valid_uscc, mask_sensitive, uscc_check_char
from local_redact.textutil import Redactor, load_lexicon

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
V = "urn:schemas-microsoft-com:vml"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
O = "urn:schemas-microsoft-com:office:office"
PKG = "http://schemas.openxmlformats.org/package/2006/relationships"

TABLE_CLIENT = "表格里的委托单位有限公司"
INLINE_CLIENT = "句子里的委托单位有限公司"
STANDARDS = "依据 GB 3838-2002、HJ 828-2017，化学需氧量 35.2 mg/L，pH 7.26，采样日期 2024-03-07，排放浓度 120.5 mg/m3。"


def q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def paragraph(*runs: str) -> etree._Element:
    node = etree.Element(q(W, "p"))
    for text in runs:
        run = etree.SubElement(node, q(W, "r"))
        text_node = etree.SubElement(run, q(W, "t"))
        text_node.text = text
    return node


def cell(text: str) -> etree._Element:
    node = etree.Element(q(W, "tc"))
    node.append(paragraph(text))
    return node


def row(*texts: str) -> etree._Element:
    node = etree.Element(q(W, "tr"))
    for text in texts:
        node.append(cell(text))
    return node


def table(*rows: etree._Element) -> etree._Element:
    node = etree.Element(q(W, "tbl"))
    for item in rows:
        node.append(item)
    return node


def image_paragraph(rel_id: str) -> etree._Element:
    node = paragraph()
    run = node.find(q(W, "r"))
    if run is None:
        run = etree.SubElement(node, q(W, "r"))
    drawing = etree.SubElement(run, q(W, "drawing"))
    inline = etree.SubElement(drawing, q(WP, "inline"))
    graphic = etree.SubElement(inline, q(A, "graphic"))
    data = etree.SubElement(graphic, q(A, "graphicData"))
    blip = etree.SubElement(data, q(A, "blip"))
    blip.set(q(R, "embed"), rel_id)
    return node


def text_watermark(text: str) -> etree._Element:
    node = paragraph()
    run = etree.SubElement(node, q(W, "r"))
    pict = etree.SubElement(run, q(W, "pict"))
    shape = etree.SubElement(pict, q(V, "shape"))
    shape.set("id", "PowerPlusWaterMarkObject1")
    path = etree.SubElement(shape, q(V, "textpath"))
    path.set("string", text)
    return node


def image_watermark(rel_id: str) -> etree._Element:
    node = paragraph()
    run = etree.SubElement(node, q(W, "r"))
    pict = etree.SubElement(run, q(W, "pict"))
    shape = etree.SubElement(pict, q(V, "shape"))
    shape.set("id", "WordPictureWatermark1")
    image = etree.SubElement(shape, q(V, "imagedata"))
    image.set(q(R, "id"), rel_id)
    return node


def drawing_watermark(text: str) -> etree._Element:
    node = paragraph()
    run = etree.SubElement(node, q(W, "r"))
    drawing = etree.SubElement(run, q(W, "drawing"))
    anchor = etree.SubElement(drawing, q(WP, "anchor"))
    props = etree.SubElement(anchor, q(WP, "docPr"))
    props.set("name", "PowerPlusWaterMarkObject2")
    content = etree.SubElement(anchor, q(W, "txbxContent"))
    content.append(paragraph(text))
    return node


def textbox(text: str) -> etree._Element:
    node = paragraph()
    run = etree.SubElement(node, q(W, "r"))
    pict = etree.SubElement(run, q(W, "pict"))
    shape = etree.SubElement(pict, q(V, "shape"))
    shape.set("id", "TextBox1")
    box = etree.SubElement(shape, q(V, "textbox"))
    content = etree.SubElement(box, q(W, "txbxContent"))
    content.append(paragraph(text))
    return node


def ole_paragraph(rel_id: str) -> etree._Element:
    node = paragraph()
    run = etree.SubElement(node, q(W, "r"))
    obj = etree.SubElement(run, q(W, "object"))
    ole = etree.SubElement(obj, q(O, "OLEObject"))
    ole.set(q(R, "id"), rel_id)
    return node


def background(rel_id: str) -> etree._Element:
    node = etree.Element(q(W, "background"))
    node.set(q(W, "color"), "FFFFFF")
    holder = etree.SubElement(node, q(V, "background"))
    fill = etree.SubElement(holder, q(V, "fill"))
    fill.set(q(R, "id"), rel_id)
    return node


def xml_bytes(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def build_report(path: Path, uscc: str, id_card: str) -> None:
    document = etree.Element(q(W, "document"))
    document.append(background("rIdBg"))
    body = etree.SubElement(document, q(W, "body"))
    body.append(
        table(
            row("委托单位", TABLE_CLIENT, "联系电话", "13800138000"),
            row("采样人", "张三", "编制", "李四"),
            row("项目地址", "杭州市西湖区文三路1号"),
        )
    )
    body.append(
        table(
            row("样品编号", "检测项目", "结果", "单位"),
            row("S2024-001", "化学需氧量", "35.2", "mg/L"),
            row("S2024-002", "pH", "7.26", "无量纲"),
        )
    )
    body.append(table(row("点位", "1#排气筒", "经纬度", "东经120.15° 北纬30.25°")))
    body.append(paragraph(f"委托单位：{INLINE_CLIENT}。结论见 GB 3838-2002。"))
    body.append(paragraph("本次受测", TABLE_CLIENT, "废水。"))
    body.append(paragraph(STANDARDS))
    body.append(paragraph("审核：符合要求。样品 S2024-001 的化学需氧量超标。现场联系人张三到场。"))
    body.append(paragraph("点位描述120°15′32″，东经120.1532°，北纬30.2511°，N30.25 E120.15，排放浓度 120.5 mg/m3。"))
    body.append(paragraph("联系号码１３８１２３４５６７８完"))
    body.append(
        paragraph(
            "敏感行开始 report@example.com https://verify.example.com/r/1 "
            f"192.168.1.20 8.8.8.8 {id_card} {uscc} api_key=abcdef123456 "
            "sk-abcdefghijklmnopqrstuvwxyz 敏感行结束"
        )
    )
    body.append(textbox(f"备注保留 {TABLE_CLIENT}"))
    body.append(image_paragraph("rIdBody"))
    body.append(ole_paragraph("rIdOle"))
    sect = etree.SubElement(body, q(W, "sectPr"))
    header_ref = etree.SubElement(sect, q(W, "headerReference"))
    header_ref.set(q(W, "type"), "default")
    header_ref.set(q(R, "id"), "rIdHeader")
    footer_ref = etree.SubElement(sect, q(W, "footerReference"))
    footer_ref.set(q(W, "type"), "default")
    footer_ref.set(q(R, "id"), "rIdFooter")

    header = etree.Element(q(W, "hdr"))
    header.append(text_watermark("机密水印示例公司"))
    header.append(image_watermark("rIdWm"))
    header.append(drawing_watermark("矢量水印不应留下"))
    header.append(paragraph("保留检测机构"))
    header.append(paragraph(TABLE_CLIENT))
    header.append(paragraph(INLINE_CLIENT))
    header.append(paragraph("138", "00138000"))

    even_header = etree.Element(q(W, "hdr"))
    even_header.append(paragraph("偶数页13800138001"))

    footer = etree.Element(q(W, "ftr"))
    footer.append(paragraph("保留检测机构  地址保留路1号  电话：0571-88886666"))

    comments = etree.Element(q(W, "comments"))
    comment = etree.SubElement(comments, q(W, "comment"))
    comment.set(q(W, "id"), "1")
    comment.set(q(W, "author"), "报告编制人甲")
    comment.set(q(W, "initials"), "甲")
    comment.append(paragraph(f"请复核{INLINE_CLIENT}的地址"))

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/embeddings/oleObject1.bin" ContentType="application/vnd.openxmlformats-officedocument.oleObject"/>
</Types>""",
        )
        archive.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rIdThumb" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.jpeg"/>
 </Relationships>""",
        )
        archive.writestr("word/comments.xml", xml_bytes(comments))
        archive.writestr("word/header1.xml", xml_bytes(header))
        archive.writestr("word/header2.xml", xml_bytes(even_header))
        archive.writestr("word/footer1.xml", xml_bytes(footer))
        archive.writestr("word/document.xml", xml_bytes(document))
        archive.writestr(
            "word/_rels/document.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
  <Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
  <Relationship Id="rIdBg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/bg.png"/>
  <Relationship Id="rIdBody" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/body.png"/>
  <Relationship Id="rIdOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="embeddings/oleObject1.bin"/>
</Relationships>""",
        )
        archive.writestr(
            "word/_rels/header1.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/wm.png"/>
</Relationships>""",
        )
        archive.writestr("word/media/bg.png", b"BGIMG-MARKER")
        archive.writestr("word/media/body.png", b"BODYIMG-MARKER")
        archive.writestr("word/media/wm.png", b"WMIMG-MARKER")
        archive.writestr("word/embeddings/oleObject1.bin", b"OLE-MARKER")
        archive.writestr(
            "docProps/core.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/">
  <dc:creator>报告编制人甲</dc:creator>
  <cp:lastModifiedBy>报告编制人甲</cp:lastModifiedBy>
  <dc:title>{INLINE_CLIENT}检测报告</dc:title>
</cp:coreProperties>""".encode("utf-8"),
        )
        archive.writestr(
            "docProps/app.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Company>{TABLE_CLIENT}</Company>
  <Manager>张三</Manager>
</Properties>""".encode("utf-8"),
        )
        archive.writestr("docProps/thumbnail.jpeg", b"THUMB-MARKER")
        archive.writestr(
            "docProps/custom.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <property fmtid="{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}" pid="2" name="Company">
    <vt:lpwstr>{TABLE_CLIENT}</vt:lpwstr>
  </property>
</Properties>""".encode("utf-8"),
        )
        archive.writestr(
            "customXml/item1.xml",
            f"<client>{TABLE_CLIENT}</client>".encode("utf-8"),
        )
        archive.writestr(
            "word/settings.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:attachedTemplate r:id="rIdTemplate"/>
  <w:docVars>
    <w:docVar w:name="客户" w:val="{TABLE_CLIENT}"/>
  </w:docVars>
</w:settings>""".encode("utf-8"),
        )
        archive.writestr(
            "word/_rels/settings.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTemplate" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="file:///C:/Users/报告编制人甲/template.dotx" TargetMode="External"/>
</Relationships>""",
        )


def package_bytes(path: Path) -> bytes:
    with zipfile.ZipFile(path) as archive:
        return b"".join(archive.read(name) for name in archive.namelist())


def drop_marker(kind: str, _original: str) -> str:
    return f"[{kind}已移除]"


class PatternTests(unittest.TestCase):
    def test_keeps_results_and_standards(self) -> None:
        self.assertEqual(mask_sensitive(STANDARDS, drop_marker), STANDARDS)
        self.assertIn("120.5 mg/m3", mask_sensitive("排放浓度 120.5 mg/m3", drop_marker))

    def test_masks_coordinates_and_contacts(self) -> None:
        masked = mask_sensitive("点位描述120°15′32″，东经120.1532°，N30.25 E120.15", drop_marker)
        self.assertNotIn("120°15", masked)
        self.assertNotIn("东经120.1532", masked)
        self.assertNotIn("N30.25", masked)

    def test_phone_masks_standalone_keeps_instrument_serial(self) -> None:
        masked = mask_sensitive("电话13800138000，固话0571-88886666", drop_marker)
        self.assertNotIn("13800138000", masked)
        self.assertNotIn("0571-88886666", masked)
        self.assertIn("[电话已移除]", masked)
        serial = "671021S13800138000"
        self.assertEqual(mask_sensitive(serial, drop_marker), serial)
        self.assertEqual(mask_sensitive("742800NB023060061", drop_marker), "742800NB023060061")

    def test_grouped_phones_keep_instrument_serial(self) -> None:
        samples = [
            "138 0013 8000",
            "138-0013-8000",
            "+86 138 0013 8000",
            "(0571)88886666",
            "（0571）88886666",
            "0571-8888 6666",
        ]
        for sample in samples:
            masked = mask_sensitive(f"前{sample}后", drop_marker)
            self.assertNotIn(sample, masked, sample)
            self.assertIn("[电话已移除]", masked)
            self.assertIn("前", masked)
            self.assertIn("后", masked)
        mixed = mask_sensitive("编号671021S13800138000，电话138 0013 8000", drop_marker)
        self.assertIn("671021S13800138000", mixed)
        self.assertNotIn("138 0013 8000", mixed)
        self.assertEqual(mask_sensitive(STANDARDS, drop_marker), STANDARDS)

    def test_id_and_uscc_checksum(self) -> None:
        self.assertTrue(is_valid_id_card("110101199003078515"))
        self.assertFalse(is_valid_id_card("110101199003078516"))
        body = "91330000MA27U1XX0"
        code = body + uscc_check_char(body)
        self.assertTrue(is_valid_uscc(code))
        self.assertFalse(is_valid_uscc(body + ("0" if code[-1] != "0" else "1")))


class FieldRedactTests(unittest.TestCase):
    def _redactor(self) -> Redactor:
        return Redactor([], b"unit-test-secret-key!!", MappingBook(Path("mapping-unused.json")))

    def test_project_number_and_spaced_address(self) -> None:
        redactor = self._redactor()
        text = (
            "项目编号：CDHZHJJCFWYXGS 7601-0001。"
            "地 址：成都市青白江区华金大道二段562号。"
            "邮政编码：610300。"
            "电 话：028-12345678。"
        )
        out = redactor.redact_string(text)
        self.assertNotIn("CDHZHJJCFWYXGS", out)
        self.assertIn("项目号_", out)
        self.assertNotIn("华金大道", out)
        self.assertIn("[地址已移除]", out)
        self.assertIn("610300", out)
        self.assertNotIn("[邮编已移除]", out)
        self.assertIn("[电话已移除]", out)

    def test_header_report_serial_kept(self) -> None:
        redactor = self._redactor()
        out = redactor.redact_string("华展环检字（2026）第0002-5号  第 2 页 共 9 页")
        self.assertIn("华展环检字（2026）第0002-5号", out)
        self.assertNotIn("报告_", out)
        self.assertIn("第 2 页 共 9 页", out)

    def test_staff_column_and_person_list(self) -> None:
        from local_redact.fields import COLUMN_RULES
        from local_redact.docx_redact import _redact_tables

        redactor = self._redactor()
        names = "李伟\n田超\n吴广"
        note = (
            "备注：检测类别为气体样品时，注明样品介质（如：吸附管、气袋、真空瓶、"
            "全玻注射器、滤膜、滤筒、吸收液等），其他检测类别（噪声、振动除外）注明样品状态。"
        )
        root = etree.Element(q(W, "document"))
        body = etree.SubElement(root, q(W, "body"))
        body.append(
            table(
                row(
                    "检测类别",
                    "采样/检测位置",
                    "点位编号",
                    "样品介质/状态",
                    "采样/检测时间",
                    "采样/检测人员",
                    "检测项目",
                    "检测频次",
                ),
                row("有组织废气", "厂界", "YG260017-01", "气袋", "2026.6.22", names, "非甲烷总烃", "1天3次"),
                row(note, "", "", "", "", "", "", ""),
            )
        )
        _redact_tables(root, redactor)
        texts = ["".join(node.text or "" for node in root.iter(q(W, "t")))]
        joined = texts[0]
        self.assertNotIn("李伟", joined)
        self.assertNotIn("田超", joined)
        self.assertNotIn("吴广", joined)
        self.assertIn("人员_", joined)
        self.assertIn("气袋", joined)
        self.assertIn("滤膜", joined)
        self.assertIn("真空瓶", joined)
        self.assertIn("YG260017-01", joined)
        self.assertIn("非甲烷总烃", joined)
        self.assertIn("采样/检测人员", COLUMN_RULES)
        self.assertNotIn("点位编号", COLUMN_RULES)
        self.assertNotIn("样品编号", COLUMN_RULES)
        self.assertNotIn("气袋", set(redactor.mapping.tokens.values()))
        self.assertNotIn("滤膜", set(redactor.mapping.tokens.values()))

    def test_person_rule_keeps_prose_and_non_names(self) -> None:
        from local_redact.fields import COLUMN_RULES

        redactor = self._redactor()
        rule = COLUMN_RULES["采样/检测人员"]
        note = "注明样品介质（如：吸附管、气袋、滤膜、滤筒、吸收液等）"
        self.assertEqual(redactor.apply_rule(note, rule), note)
        self.assertEqual(redactor.apply_rule("气袋", rule), "气袋")
        self.assertEqual(redactor.apply_rule("气袋、滤膜、滤筒", rule), "气袋、滤膜、滤筒")
        self.assertIn("人员_", redactor.apply_rule("李伟田超吴广", rule))
        self.assertIn("人员_", redactor.apply_rule("李伟、田超、吴广", rule))
        # 二字名+三字名无分隔（曾跃杰、许红强、宁建荣）
        mixed = "李伟田超吴广曾跃杰许红强宁建荣"
        self.assertIn("人员_", redactor.apply_rule(mixed, rule))
        self.assertNotIn("李伟", redactor.apply_rule(mixed, rule))

    def test_serial_column_still_redacts_staff_and_modern_surnames(self) -> None:
        from local_redact.docx_redact import _redact_tables
        from local_redact.fields import HEADER_WORDS

        redactor = self._redactor()
        root = etree.Element(q(W, "document"))
        body = etree.SubElement(root, q(W, "body"))
        body.append(
            table(
                row("序号", "样品编号", "检测项目", "结果", "采样人员"),
                row("1", "S2024-001", "化学需氧量", "35.2", "李伟"),
                row("2", "S2024-002", "pH", "7.26", "肖强"),
                row("3", "S2024-003", "氨氮", "0.12", "闫丽"),
                row("4", "S2024-004", "总磷", "0.08", "付伟"),
                row("5", "S2024-005", "总氮", "1.10", "王芳"),
            )
        )
        _redact_tables(root, redactor)
        joined = "".join(node.text or "" for node in root.iter(q(W, "t")))
        for secret in ("李伟", "肖强", "闫丽", "付伟", "王芳"):
            self.assertNotIn(secret, joined, secret)
        self.assertIn("人员_", joined)
        self.assertIn("S2024-001", joined)
        self.assertIn("35.2", joined)
        self.assertIn("序号", HEADER_WORDS)
        self.assertIn("样品编号", HEADER_WORDS)

    def test_projected_coordinate_column_keeps_results(self) -> None:
        from local_redact.docx_redact import _redact_tables
        from local_redact.fields import COLUMN_RULES

        redactor = self._redactor()
        root = etree.Element(q(W, "document"))
        body = etree.SubElement(root, q(W, "body"))
        body.append(
            table(
                row("点位", "东坐标", "北坐标", "结果"),
                row("厂界", "3456789.12", "567890.34", "35.2"),
            )
        )
        _redact_tables(root, redactor)
        joined = "".join(node.text or "" for node in root.iter(q(W, "t")))
        self.assertNotIn("3456789.12", joined)
        self.assertNotIn("567890.34", joined)
        self.assertIn("[坐标已移除]", joined)
        self.assertIn("35.2", joined)
        self.assertIn("高斯坐标", COLUMN_RULES)

    def test_cover_labels(self) -> None:
        redactor = self._redactor()
        text = (
            "法定代表人：李伟。项目负责人：肖强。"
            "送样单位：杭州示例环保科技有限公司。企业名称：示例企业有限公司。客户名称：示例客户。"
        )
        out = redactor.redact_string(text)
        for secret in ("李伟", "肖强", "杭州示例环保科技有限公司", "示例企业有限公司", "示例客户"):
            self.assertNotIn(secret, out, secret)
        self.assertIn("人员_", out)
        self.assertIn("机构_", out)

    def test_lexicon_gb18030_and_bom(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "lexicon.txt"
            path.write_bytes("张三 => 人员甲\n".encode("gb18030"))
            self.assertEqual(load_lexicon(path), [("张三", "人员甲")])
            path.write_bytes(b"\xef\xbb\xbf" + "示例企业\n".encode("utf-8"))
            self.assertEqual(load_lexicon(path), [("示例企业", None)])


class ReportRedactTests(unittest.TestCase):
    def test_redacts_environmental_report(self) -> None:
        body = "91330000MA27U1XX0"
        uscc = body + uscc_check_char(body)
        id_card = "110101199003078515"
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            src_dir = root / "in"
            src = src_dir / "检测报告.docx"
            build_report(src, uscc, id_card)
            out = root / "out"
            code = main(
                [
                    str(src_dir),
                    "-o",
                    str(out),
                    "--lexicon",
                    str(self._lexicon(root)),
                    "--map",
                    str(root / "mapping.json"),
                    "--secret",
                    str(root / "secret.key"),
                ]
            )
            self.assertEqual(code, 0)
            dest = out / "检测报告.docx"
            self.assertTrue(dest.exists())
            blob = package_bytes(dest)
            text = blob.decode("utf-8", errors="ignore")

            for secret in (
                TABLE_CLIENT,
                INLINE_CLIENT,
                "张三",
                "李四",
                "杭州市西湖区文三路1号",
                "13800138000",
                "13800138001",
                "13812345678",
                "１３８１２３４５６７８",
                "0571-88886666",
                "东经",
                "120°15",
                "N30.25",
                "机密水印示例公司",
                "矢量水印不应留下",
                "报告编制人甲",
                id_card,
                uscc,
                "report@example.com",
                "https://verify.example.com/r/1",
                "192.168.1.20",
                "api_key=abcdef123456",
                "sk-abcdefghijklmnopqrstuvwxyz",
            ):
                self.assertNotIn(secret, text, secret)

            for marker in (b"BGIMG-MARKER", b"BODYIMG-MARKER", b"WMIMG-MARKER", b"OLE-MARKER", b"THUMB-MARKER"):
                self.assertNotIn(marker, blob)
            self.assertNotIn("file:///C:/Users/", text)
            self.assertNotIn("thumbnail.jpeg", text)

            for kept in (
                "保留检测机构",
                "地址保留路1号",
                "1#排气筒",
                "备注保留",
                "偶数页",
                "符合要求",
                "GB 3838-2002",
                "HJ 828-2017",
                "35.2",
                "7.26",
                "无量纲",
                "2024-03-07",
                "120.5 mg/m3",
                "8.8.8.8",
                "化学需氧量",
                "S2024-001",
                "S2024-002",
                "本次受测",
                "废水",
                "联系号码",
                "完",
            ):
                self.assertIn(kept, text, kept)

            mapping = json.loads((root / "mapping.json").read_text(encoding="utf-8"))
            originals = set(mapping["tokens"].values())
            self.assertIn(TABLE_CLIENT, originals)
            self.assertNotIn("S2024-001", originals)
            self.assertNotIn("mapping.json", [path.name for path in out.rglob("*")])
            self.assertFalse((out / ".redact-tmp").exists())

            with zipfile.ZipFile(dest) as archive:
                names = set(archive.namelist())
                self.assertNotIn("word/media/bg.png", names)
                self.assertNotIn("word/embeddings/oleObject1.bin", names)
                self.assertNotIn("docProps/thumbnail.jpeg", names)
                rels = etree.fromstring(archive.read("word/_rels/document.xml.rels"))
                targets = [rel.get("Target") for rel in rels]
                self.assertIn("header1.xml", targets)
                self.assertIn("footer1.xml", targets)
                self.assertFalse(any((target or "").startswith("media/") for target in targets))
                document = etree.fromstring(archive.read("word/document.xml"))
                self.assertIsNotNone(document.find(f".//{q(W, 'body')}"))

    def test_redacts_filenames_and_folders(self) -> None:
        uscc = "91330000MA27U1XX0"
        uscc = uscc + uscc_check_char(uscc)
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            src = root / "in"
            build_report(src / TABLE_CLIENT / f"{TABLE_CLIENT}检测报告.docx", uscc, "110101199003078515")
            build_report(src / "批次" / "报告13800138000.docx", uscc, "110101199003078515")
            lexicon = root / "lexicon.txt"
            lexicon.write_text("# 空词表\n", encoding="utf-8")
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
            self.assertFalse((out / ".redact-tmp").exists())
            for path in out.rglob("*"):
                self.assertNotIn(TABLE_CLIENT, path.name, path)
                self.assertNotIn("13800138000", path.name, path)

            mapping = json.loads((root / "mapping.json").read_text(encoding="utf-8"))
            client_key = f"{TABLE_CLIENT}/{TABLE_CLIENT}检测报告.docx"
            self.assertIn(client_key, mapping["files"])
            self.assertNotIn(TABLE_CLIENT, mapping["files"][client_key])
            phone_key = "批次/报告13800138000.docx"
            self.assertIn(phone_key, mapping["files"])
            self.assertNotIn("13800138000", mapping["files"][phone_key])
            self.assertIn("[电话已移除]", mapping["files"][phone_key])
            token = next(item for item, original in mapping["tokens"].items() if original == TABLE_CLIENT)
            self.assertIn(token, mapping["files"][client_key])
            self.assertTrue((out / "批次").is_dir())

    def test_filename_rules(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            book = MappingBook(Path(folder) / "map.json")
            redactor = Redactor([("某某检测有限公司", "机构甲")], b"0123456789abcdef", book)
            renamed = redact_relative(
                Path("某某检测有限公司") / "某某检测有限公司报告13800138000.docx",
                redactor,
            )
            text = renamed.as_posix()
            self.assertNotIn("某某检测有限公司", text)
            self.assertNotIn("13800138000", text)
            self.assertIn("机构甲", text)
            self.assertTrue(text.endswith(".docx"))

            plain = Redactor([], b"0123456789abcdef", book)
            self.assertEqual(redact_relative(Path("检测报告.docx"), plain), Path("检测报告.docx"))
            self.assertEqual(redact_relative(Path("检测报告.doc"), plain), Path("检测报告.docx"))
            self.assertEqual(redact_relative(Path("GB 3838-2002.docx"), plain), Path("GB 3838-2002.docx"))

        names = OutputNames(b"0123456789abcdef")
        first = names.place(Path("a/1.docx"), Path("报告.docx"))
        second = names.place(Path("b/2.docx"), Path("报告.docx"))
        self.assertEqual(first, Path("报告.docx"))
        self.assertNotEqual(first, second)
        again = OutputNames(b"0123456789abcdef")
        self.assertEqual(again.place(Path("a/1.docx"), Path("报告.docx")), first)
        self.assertEqual(again.place(Path("b/2.docx"), Path("报告.docx")), second)

    def test_converts_doc_then_redacts(self) -> None:
        uscc = "91330000MA27U1XX0" + uscc_check_char("91330000MA27U1XX0")
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            template = root / "template.docx"
            build_report(template, uscc, "110101199003078515")
            src = root / "in"
            original = src / f"{TABLE_CLIENT}.doc"
            original.parent.mkdir()
            original.write_bytes(b"dummy-doc")

            class FakeWord:
                def start(self) -> None:
                    return None

                def convert(self, source: Path, dest: Path) -> None:
                    shutil.copyfile(template, dest)

                def close(self) -> None:
                    return None

            lexicon = root / "lexicon.txt"
            lexicon.write_text("# 空词表\n", encoding="utf-8")
            out = root / "out"
            with patch("local_redact.cli.WordSession", FakeWord):
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
            self.assertEqual(original.read_bytes(), b"dummy-doc")
            outputs = list(out.rglob("*.docx"))
            self.assertEqual(len(outputs), 1)
            self.assertNotIn(TABLE_CLIENT, outputs[0].name)
            self.assertNotIn(TABLE_CLIENT, package_bytes(outputs[0]).decode("utf-8", errors="ignore"))
            mapping = json.loads((root / "mapping.json").read_text(encoding="utf-8"))
            self.assertTrue(mapping["files"][f"{TABLE_CLIENT}.doc"].endswith(".docx"))

    def test_same_secret_same_token(self) -> None:
        secret = b"0123456789abcdef"
        with tempfile.TemporaryDirectory() as folder:
            first = Redactor([], secret, MappingBook(Path(folder) / "a.json")).token_for("机构", "同一家公司")
            second = Redactor([], secret, MappingBook(Path(folder) / "b.json")).token_for("机构", "同一家公司")
        self.assertEqual(first, second)

    def test_load_empty_mapping_file(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "mapping.json"
            path.write_text("", encoding="utf-8")
            book = MappingBook(path)
            book.load()
            self.assertEqual(book.tokens, {})
            self.assertEqual(book.files, {})
            self.assertEqual(book.removed, {})

    def test_rejects_legacy_doc(self) -> None:
        redactor = Redactor([], b"0123456789abcdef", MappingBook(Path("mapping.json")))
        with self.assertRaises(DocxError):
            redact_docx(Path("原稿.doc"), Path("副本.docx"), redactor)

    def _lexicon(self, root: Path) -> Path:
        path = root / "lexicon.txt"
        path.write_text("张三 => 人员甲\n# 注释不要当成词\n", encoding="utf-8")
        return path


if __name__ == "__main__":
    unittest.main()
