"""输出路径里的文件名和目录名。扩展名保持 .docx。"""

from __future__ import annotations

import hashlib
import hmac
import re
from pathlib import Path

from local_redact.textutil import Redactor

_INVALID = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
_MAX_PART = 80


def sanitize_part(text: str) -> str:
    cleaned = _INVALID.sub("", text)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    if cleaned.upper() in _RESERVED:
        cleaned = f"_{cleaned}"
    if len(cleaned) > _MAX_PART:
        cleaned = cleaned[:_MAX_PART].rstrip(" .")
    return cleaned


def redact_relative(relative: Path, redactor: Redactor) -> Path:
    """脱敏相对路径的每一段。最后一段按文件名处理，只改主文件名。"""
    raw_parts = list(relative.parts)
    redacted: list[str] = []
    for index, part in enumerate(raw_parts):
        is_file = index == len(raw_parts) - 1
        if is_file:
            stem = sanitize_part(redactor.redact_string(Path(part).stem)) or "报告"
            redacted.append(f"{stem}.docx")
        else:
            redacted.append(sanitize_part(redactor.redact_string(part)) or "目录")
    return Path(*redacted)


class OutputNames:
    """同一批里脱敏后重名时，用原路径的稳定后缀区分。"""

    def __init__(self, secret: bytes):
        self.secret = secret
        self._taken: dict[str, str] = {}

    def place(self, original: Path, redacted: Path) -> Path:
        original_key = original.as_posix()
        candidate = redacted
        owner = self._taken.get(candidate.as_posix())
        if owner not in (None, original_key):
            digest = hmac.new(self.secret, original_key.encode("utf-8"), hashlib.sha256).hexdigest()[:6]
            candidate = redacted.with_name(f"{redacted.stem}_{digest}{redacted.suffix}")
            number = 2
            while self._taken.get(candidate.as_posix()) not in (None, original_key):
                candidate = redacted.with_name(f"{redacted.stem}_{digest}_{number}{redacted.suffix}")
                number += 1
        self._taken[candidate.as_posix()] = original_key
        return candidate
