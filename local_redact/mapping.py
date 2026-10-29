"""本地对照表。和脱敏后的 Word 分开存放，里面是原文，不能上传。"""

from __future__ import annotations

import hashlib
import hmac
import json
from pathlib import Path


class MappingBook:
    def __init__(self, path: Path):
        self.path = path
        self.tokens: dict[str, str] = {}
        self.removed: dict[str, list[str]] = {}
        self.files: dict[str, str] = {}

    def load(self) -> None:
        if not self.path.exists():
            return
        raw = self.path.read_text(encoding="utf-8").strip()
        # 允许用户直接清空文件；空内容按新对照表处理
        if not raw:
            return
        try:
            data = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ValueError(
                f"对照表不是合法 JSON：{self.path}。可删掉该文件或清空后重试。"
            ) from exc
        if not isinstance(data, dict):
            raise ValueError(f"对照表格式不对（应为对象）：{self.path}")
        self.tokens.update(data.get("tokens") or {})
        self.files.update(data.get("files") or {})
        for kind, values in (data.get("removed") or {}).items():
            bucket = self.removed.setdefault(kind, [])
            for value in values:
                if value not in bucket:
                    bucket.append(value)

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "version": 1,
            "warning": "含原文，不要和脱敏文件一起上传或提交",
            "tokens": self.tokens,
            "removed": self.removed,
            "files": self.files,
        }
        self.path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def clear(self) -> None:
        """丢掉已记录的原文、代号和文件名对应，并写回文件。"""
        self.tokens.clear()
        self.removed.clear()
        self.files.clear()
        self.save()

    def token_for(self, prefix: str, original: str, secret: bytes) -> str:
        digest = hmac.new(secret, f"{prefix}\n{original}".encode("utf-8"), hashlib.sha256).hexdigest()[:8]
        token = f"{prefix}_{digest}"
        self.tokens[token] = original
        return token

    def remember(self, token: str, original: str) -> None:
        self.tokens[token] = original

    def remember_file(self, original: str, renamed: str) -> None:
        self.files[original] = renamed

    def drop(self, kind: str, original: str) -> str:
        bucket = self.removed.setdefault(kind, [])
        if original not in bucket:
            bucket.append(original)
        return f"[{kind}已移除]"
