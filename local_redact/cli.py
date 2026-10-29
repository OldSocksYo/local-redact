"""命令行入口。"""

from __future__ import annotations

import argparse
import os
import secrets
import shutil
import sys
import tempfile
import traceback
from pathlib import Path

from local_redact.convert import ConversionError, SheetSession, WordSession
from local_redact.docx_redact import DocxError, redact_docx
from local_redact.mapping import MappingBook
from local_redact.names import EXCEL_SUFFIXES, WORD_SUFFIXES, OutputNames, output_suffix, redact_relative
from local_redact.textutil import Redactor, load_lexicon
from local_redact.xlsx_redact import XlsxError, redact_xlsx

ROOT = Path(__file__).resolve().parents[1]


def _load_secret(path: Path) -> bytes:
    if path.exists():
        data = path.read_bytes()
        if len(data) >= 16:
            return data
    path.parent.mkdir(parents=True, exist_ok=True)
    data = secrets.token_bytes(32)
    path.write_bytes(data)
    return data


def _is_office_document(path: Path) -> bool:
    suffix = path.suffix.lower()
    return path.is_file() and suffix in WORD_SUFFIXES | EXCEL_SUFFIXES and not path.name.startswith("~$")


def _collect_documents(src: Path) -> list[Path]:
    if src.is_file():
        return [src] if _is_office_document(src) else []
    return sorted(path for path in src.rglob("*") if _is_office_document(path))


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="local_redact",
        description="本地脱敏 Word 报告和 Excel 表格。保留检测数值和标准号，去掉身份、地址、坐标、图片、水印，并脱敏输出文件名。",
    )
    parser.add_argument("src", type=Path, help="一份 Word / Excel，或一个文件夹")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=ROOT / "output",
        help="脱敏副本的输出目录，默认是项目下的 output",
    )
    parser.add_argument("--lexicon", type=Path, default=None, help="词表，默认使用项目里的 lexicon.txt（如果有）")
    parser.add_argument("--map", type=Path, default=ROOT / "data" / "mapping.json", help="对照表路径，不要放进输出目录")
    parser.add_argument("--secret", type=Path, default=ROOT / "data" / "secret.key", help="稳定代号用的本机密钥")
    return parser


def _path_inside(child: Path, parent: Path) -> bool:
    try:
        child_resolved = child.resolve()
        parent_resolved = parent.resolve()
    except OSError:
        return False
    return child_resolved == parent_resolved or child_resolved.is_relative_to(parent_resolved)


def _output_conflict(
    src: Path,
    output: Path,
    mapping_path: Path,
    secret_path: Path,
    lexicon_path: Path | None,
) -> str | None:
    try:
        resolved = output.resolve()
    except OSError as exc:
        return f"无法使用输出目录：{exc}"
    if resolved == Path(resolved.anchor):
        return "输出目录不能是磁盘根目录。"
    try:
        if resolved == ROOT.resolve():
            return "输出目录不能是项目目录。"
    except OSError:
        pass
    if _path_inside(mapping_path, output):
        return "对照表不能放在输出目录里。"
    if _path_inside(secret_path, output):
        return "密钥不能放在输出目录里。"
    if lexicon_path is not None and _path_inside(lexicon_path, output):
        return "词表不能放在输出目录里。"
    if src.is_dir() and _path_inside(output, src):
        return "输出目录不要放在输入目录里面。"
    if _path_inside(src, output):
        return "输出目录里还有输入文件，已停止，避免误删原稿。"
    return None


def _clear_output(output: Path) -> bool:
    removed = False
    for child in list(output.iterdir()):
        removed = True
        if child.is_dir():
            shutil.rmtree(child)
        else:
            child.unlink()
    return removed


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError):
            pass


def _redact_one(
    path: Path,
    dest: Path,
    tmp_dir: Path,
    index: int,
    redactor: Redactor,
    word: WordSession | None,
    word_startup_error: str,
    excel: SheetSession | None,
    excel_startup_error: str,
) -> dict[str, int]:
    source = path
    converted: Path | None = None
    kind = path.suffix.lower()
    try:
        if kind == ".doc":
            if word is None:
                raise ConversionError(word_startup_error or "没有检测到可用的 Word 或 WPS，无法转换 .doc。")
            converted = tmp_dir / f"conv-{index}.docx"
            word.convert(path, converted)
            source = converted
            return redact_docx(source, dest, redactor)
        if kind == ".xls":
            if excel is None:
                raise ConversionError(excel_startup_error or "没有检测到可用的 Excel 或 WPS 表格，无法转换 .xls。")
            converted = tmp_dir / f"conv-{index}.xlsx"
            excel.convert(path, converted)
            source = converted
            return redact_xlsx(source, dest, redactor)
        if kind in EXCEL_SUFFIXES:
            return redact_xlsx(source, dest, redactor)
        return redact_docx(source, dest, redactor)
    finally:
        if converted is not None and converted.exists():
            converted.unlink()


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = _build_parser().parse_args(argv)
    src: Path = args.src
    output: Path = args.output
    if not src.exists():
        print(f"找不到输入：{src}", file=sys.stderr)
        return 1

    lexicon_path = args.lexicon
    if lexicon_path is None:
        default_lexicon = ROOT / "lexicon.txt"
        lexicon_path = default_lexicon if default_lexicon.exists() else None
    if lexicon_path is not None and not lexicon_path.exists():
        print(f"找不到词表：{lexicon_path}", file=sys.stderr)
        return 1
    conflict = _output_conflict(src, output, args.map, args.secret, lexicon_path)
    if conflict is not None:
        print(conflict, file=sys.stderr)
        return 1

    lexicon = []
    if lexicon_path is not None:
        try:
            lexicon = load_lexicon(lexicon_path)
        except UnicodeDecodeError:
            print(f"词表编码无法识别，请另存为 UTF-8：{lexicon_path}", file=sys.stderr)
            return 1
        except OSError as exc:
            print(f"无法读取词表：{exc}", file=sys.stderr)
            return 1
        print(f"词表：{lexicon_path}（{len(lexicon)} 条）")
    else:
        print("未使用词表，只按栏目和固定格式脱敏。两三字的人名需要写入词表。")

    files = _collect_documents(src)
    if not files:
        print("没有找到 Word 或 Excel 文件。", file=sys.stderr)
        return 1

    try:
        cleared = _clear_output(output) if output.exists() else False
    except OSError as exc:
        print(f"无法清空输出目录：{exc}", file=sys.stderr)
        return 1
    if cleared:
        print(f"已清空输出目录：{output}")
    output.mkdir(parents=True, exist_ok=True)
    mapping = MappingBook(args.map)
    mapping.load()
    redactor = Redactor(lexicon, _load_secret(args.secret), mapping)
    names = OutputNames(redactor.secret)
    tmp_dir = output / ".redact-tmp"
    ok = 0
    failed = 0
    word: WordSession | None = None
    excel: SheetSession | None = None
    word_startup_error = ""
    excel_startup_error = ""
    if any(path.suffix.lower() == ".doc" for path in files):
        word = WordSession()
        try:
            word.start()
        except ConversionError as exc:
            word_startup_error = str(exc)
            word = None
    if any(path.suffix.lower() == ".xls" for path in files):
        excel = SheetSession()
        try:
            excel.start()
        except ConversionError as exc:
            excel_startup_error = str(exc)
            excel = None
    try:
        tmp_dir.mkdir(exist_ok=True)
        for index, path in enumerate(files, start=1):
            relative = path.relative_to(src) if src.is_dir() else Path(path.name)
            fd, tmp_name = tempfile.mkstemp(prefix="part-", suffix=output_suffix(path.name), dir=tmp_dir)
            os.close(fd)
            tmp_path = Path(tmp_name)
            try:
                stats = _redact_one(
                    path,
                    tmp_path,
                    tmp_dir,
                    index,
                    redactor,
                    word,
                    word_startup_error,
                    excel,
                    excel_startup_error,
                )
                redacted = names.place(relative, redact_relative(relative, redactor))
                dest = output / redacted
                dest.parent.mkdir(parents=True, exist_ok=True)
                os.replace(tmp_path, dest)
            except (DocxError, XlsxError, ConversionError) as exc:
                failed += 1
                print(f"[{index}/{len(files)}] 失败 {relative}：{exc}")
                continue
            except Exception:  # noqa: BLE001
                failed += 1
                print(f"[{index}/{len(files)}] 失败 {relative}")
                traceback.print_exc()
                continue
            finally:
                if tmp_path.exists():
                    tmp_path.unlink()
            mapping.remember_file(relative.as_posix(), redacted.as_posix())
            ok += 1
            print(
                f"[{index}/{len(files)}] {relative} -> {redacted}  "
                f"图形 {stats['shapes']}，图片文件 {stats['media']}，栏目 {stats['fields']}"
            )
    finally:
        if word is not None:
            word.close()
        if excel is not None:
            excel.close()
        shutil.rmtree(tmp_dir, ignore_errors=True)

    mapping.save()
    print(f"完成 {ok} 个，失败 {failed} 个。")
    print(f"输出目录：{output}")
    print(f"对照表：{args.map}")
    print("对照表里是原文，不要和脱敏文件一起上传。")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
