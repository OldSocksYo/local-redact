"""命令行入口。"""

from __future__ import annotations

import argparse
import os
import secrets
import shutil
import sys
import tempfile
from pathlib import Path

from local_redact.convert import DocConversionError, WordSession
from local_redact.docx_redact import DocxError, redact_docx
from local_redact.mapping import MappingBook
from local_redact.names import OutputNames, redact_relative
from local_redact.textutil import Redactor, load_lexicon

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


def _is_word_document(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in {".doc", ".docx"} and not path.name.startswith("~$")


def _collect_documents(src: Path) -> list[Path]:
    if src.is_file():
        return [src] if _is_word_document(src) else []
    return sorted(path for path in src.rglob("*") if _is_word_document(path))


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="local_redact",
        description="本地脱敏 Word 报告。保留检测数值和标准号，去掉身份、地址、坐标、图片、水印，并脱敏输出文件名。",
    )
    parser.add_argument("src", type=Path, help="一份 .doc / .docx，或一个文件夹")
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


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError):
            pass


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = _build_parser().parse_args(argv)
    src: Path = args.src
    output: Path = args.output
    if not src.exists():
        print(f"找不到输入：{src}", file=sys.stderr)
        return 1
    if src.is_dir():
        try:
            if output.resolve().is_relative_to(src.resolve()):
                print("输出目录不要放在输入目录里面。", file=sys.stderr)
                return 1
        except OSError:
            pass
    try:
        if args.map.resolve().is_relative_to(output.resolve()):
            print("对照表不能放在输出目录里。", file=sys.stderr)
            return 1
    except OSError:
        pass

    lexicon_path = args.lexicon
    if lexicon_path is None:
        default_lexicon = ROOT / "lexicon.txt"
        lexicon_path = default_lexicon if default_lexicon.exists() else None
    lexicon = []
    if lexicon_path is not None:
        if not lexicon_path.exists():
            print(f"找不到词表：{lexicon_path}", file=sys.stderr)
            return 1
        lexicon = load_lexicon(lexicon_path)
        print(f"词表：{lexicon_path}（{len(lexicon)} 条）")
    else:
        print("未使用词表，只按栏目和固定格式脱敏。两三字的人名需要写入词表。")

    files = _collect_documents(src)
    if not files:
        print("没有找到 .doc 或 .docx 文件。", file=sys.stderr)
        return 1

    output.mkdir(parents=True, exist_ok=True)
    mapping = MappingBook(args.map)
    mapping.load()
    redactor = Redactor(lexicon, _load_secret(args.secret), mapping)
    names = OutputNames(redactor.secret)
    tmp_dir = output / ".redact-tmp"
    ok = 0
    failed = 0
    word: WordSession | None = None
    word_startup_error = ""
    if any(path.suffix.lower() == ".doc" for path in files):
        word = WordSession()
        try:
            word.start()
        except DocConversionError as exc:
            word_startup_error = str(exc)
            word = None
    try:
        tmp_dir.mkdir(exist_ok=True)
        for index, path in enumerate(files, start=1):
            relative = path.relative_to(src) if src.is_dir() else Path(path.name)
            fd, tmp_name = tempfile.mkstemp(prefix="part-", suffix=".docx", dir=tmp_dir)
            os.close(fd)
            tmp_path = Path(tmp_name)
            converted: Path | None = None
            try:
                source = path
                if path.suffix.lower() == ".doc":
                    if word is None:
                        raise DocConversionError(word_startup_error)
                    converted = tmp_dir / f"conv-{index}.docx"
                    word.convert(path, converted)
                    source = converted
                stats = redact_docx(source, tmp_path, redactor)
                redacted = names.place(relative, redact_relative(relative, redactor))
                dest = output / redacted
                dest.parent.mkdir(parents=True, exist_ok=True)
                os.replace(tmp_path, dest)
            except (DocxError, DocConversionError) as exc:
                failed += 1
                print(f"[{index}/{len(files)}] 失败 {relative}：{exc}")
                continue
            except Exception as exc:  # noqa: BLE001
                failed += 1
                print(f"[{index}/{len(files)}] 失败 {relative}：{exc}")
                continue
            finally:
                if tmp_path.exists():
                    tmp_path.unlink()
                if converted is not None and converted.exists():
                    converted.unlink()
            mapping.remember_file(relative.as_posix(), redacted.as_posix())
            ok += 1
            print(
                f"[{index}/{len(files)}] {relative} -> {redacted}  "
                f"图形 {stats['shapes']}，图片文件 {stats['media']}，栏目 {stats['fields']}"
            )
    finally:
        if word is not None:
            word.close()
        shutil.rmtree(tmp_dir, ignore_errors=True)

    mapping.save()
    print(f"完成 {ok} 个，失败 {failed} 个。")
    print(f"输出目录：{output}")
    print(f"对照表：{args.map}")
    print("对照表里是原文，不要和脱敏文件一起上传。")
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
