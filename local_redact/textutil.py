"""把栏目、词表和正则套到一段文字上。不读文件。"""

from __future__ import annotations

import re
from pathlib import Path

from local_redact.fields import LABELS, TOKEN_RE, Rule, normalize_label
from local_redact.mapping import MappingBook
from local_redact.patterns import mask_sensitive

_PHONE_VALUE = re.compile(r"(?:\+?86[-\s]?)?1[3-9]\d{9}|0\d{2,3}[-\s]?\d{7,8}")
_PERSON_VALUE = re.compile(r"[\u4e00-\u9fffA-Za-z·•.]{2,20}")
_COORD_VALUE = re.compile(r"[\d°度′'’＇″\"”＂.\s,，;；、东西南北经纬EeWwNnSs+\-]{4,80}")
# 页眉常见报告文号：华展环检字（2026）第0002-5号
_REPORT_SERIAL = re.compile(
    r"(?<![\u4e00-\u9fffA-Za-z0-9_])"
    r"[\u4e00-\u9fffA-Za-z0-9]{2,40}字\s*[（(]\d{4}[）)]\s*第\s*[\d\-—～~]+\s*号"
)


def load_lexicon(path: Path) -> list[tuple[str, str | None]]:
    items: list[tuple[str, str | None]] = []
    for line in path.read_text(encoding="utf-8").splitlines():
        raw = line.strip()
        if not raw or raw.startswith("#"):
            continue
        if "=>" in raw:
            left, right = raw.split("=>", 1)
            word, token = left.strip(), right.strip()
            if word and token:
                items.append((word, token))
        else:
            items.append((raw, None))
    return items


def fold_digits(text: str) -> str:
    chars = []
    for char in text:
        code = ord(char)
        if 0xFF10 <= code <= 0xFF19:
            chars.append(chr(code - 0xFEE0))
        else:
            chars.append(char)
    return "".join(chars)


def _already_masked(value: str) -> bool:
    return TOKEN_RE.fullmatch(value.strip()) is not None


def _should_propagate(text: str) -> bool:
    if len(text) < 4 or _already_masked(text):
        return False
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        return False
    return True


class Redactor:
    def __init__(self, lexicon: list[tuple[str, str | None]], secret: bytes, mapping: MappingBook):
        self.lexicon = [(word, token) for word, token in lexicon if len(word) >= 2]
        self.secret = secret
        self.mapping = mapping
        self.aliases: list[tuple[str, str]] = []
        self.field_hits = 0
        self._inline_re = _compile_inline()

    def begin_file(self) -> None:
        self.aliases = []
        self.field_hits = 0

    def token_for(self, prefix: str, original: str) -> str:
        return self.mapping.token_for(prefix, original, self.secret)

    def lexicon_lookup(self, word: str) -> str | None:
        for src, fixed in self.lexicon:
            if src != word:
                continue
            if fixed:
                self.mapping.remember(fixed, word)
                return fixed
            return self.token_for("名称", word)
        return None

    def propagate(self, original: str, replacement: str) -> None:
        text = original.strip()
        if _should_propagate(text):
            self.aliases.append((text, replacement))

    def apply_rule(self, raw: str, rule: Rule) -> str:
        value = raw.strip()
        if not value or _already_masked(value):
            return value
        # 人员列常见多人：换行或顿号分隔，逐个打码并传播到全文
        if rule.value_mode == "person" and re.search(r"[\n\r、，,;；/／]", value):
            return self._apply_person_list(value, rule)
        return self._apply_one(value, rule)

    def _apply_one(self, value: str, rule: Rule) -> str:
        if not value or _already_masked(value):
            return value
        fixed = self.lexicon_lookup(value)
        if fixed:
            replacement = fixed
        elif rule.drop:
            replacement = self.mapping.drop(rule.drop, value)
        else:
            replacement = self.token_for(rule.prefix or "名称", value)
        self.propagate(value, replacement)
        self.field_hits += 1
        return replacement

    def _apply_person_list(self, value: str, rule: Rule) -> str:
        chunks = re.split(r"([\n\r]+|[、，,;；/／]+)", value)
        out: list[str] = []
        for chunk in chunks:
            if not chunk:
                continue
            if re.fullmatch(r"[\n\r]+|[、，,;；/／]+", chunk):
                out.append(chunk)
                continue
            person = chunk.strip()
            if len(person) < 2:
                out.append(chunk)
                continue
            out.append(self._apply_one(person, rule))
        return "".join(out)

    def redact_string(self, text: str) -> str:
        if not text:
            return text
        updated = fold_digits(text)
        updated = self._inline(updated)
        updated = self._report_serials(updated)
        updated = self._terms(updated)
        updated = mask_sensitive(updated, self.mapping.drop)
        return updated

    def _inline(self, text: str) -> str:
        if self._inline_re is None or ("：" not in text and ":" not in text):
            return text

        def repl(match: re.Match[str]) -> str:
            rule = LABELS[normalize_label(match.group("label"))]
            extracted = _extract(match.group("value"), rule.value_mode)
            if not extracted or _already_masked(extracted):
                return match.group(0)
            replacement = self.apply_rule(extracted, rule)
            return match.group(0).replace(extracted, replacement, 1)

        return self._inline_re.sub(repl, text)

    def _report_serials(self, text: str) -> str:
        def repl(match: re.Match[str]) -> str:
            found = match.group(0)
            if _already_masked(found):
                return found
            replacement = self.token_for("报告", found)
            self.propagate(found, replacement)
            return replacement

        return _REPORT_SERIAL.sub(repl, text)

    def _terms(self, text: str) -> str:
        pairs: list[tuple[str, str]] = []
        seen: set[str] = set()
        for word, token in self.aliases:
            if word and word not in seen:
                seen.add(word)
                pairs.append((word, token))
        for word, fixed in self.lexicon:
            if word in seen or word not in text:
                continue
            seen.add(word)
            if fixed:
                self.mapping.remember(fixed, word)
                pairs.append((word, fixed))
            else:
                pairs.append((word, self.token_for("名称", word)))
        if not pairs:
            return text
        return _replace_outside_tokens(text, pairs)


def _flex_label(label: str) -> str:
    """允许「地 址」「电 话」这类字间空格。"""
    return r"\s*".join(re.escape(ch) for ch in label)


def _compile_inline() -> re.Pattern[str] | None:
    labels = [label for label, rule in LABELS.items() if rule.inline]
    if not labels:
        return None
    labels.sort(key=len, reverse=True)
    alt = "|".join(_flex_label(label) for label in labels)
    return re.compile(
        rf"(?P<label>{alt})\s*[:：]\s*(?P<value>.*?)(?=(?:{alt})\s*[:：]|[。；;\n]|$)"
    )


def _extract(value: str, mode: str) -> str:
    value = value.strip().strip("，, ")
    if not value:
        return ""
    if mode == "phone":
        match = _PHONE_VALUE.match(value)
        return match.group(0) if match else ""
    if mode == "person":
        match = _PERSON_VALUE.match(value)
        return match.group(0) if match else ""
    if mode == "coord":
        match = _COORD_VALUE.match(value)
        return match.group(0).strip(" ，,;；") if match else ""
    return value[:80].strip()


def _replace_outside_tokens(text: str, pairs: list[tuple[str, str]]) -> str:
    ordered = sorted(pairs, key=lambda item: len(item[0]), reverse=True)
    mapping = {word: token for word, token in ordered}
    pattern = re.compile("|".join(re.escape(word) for word, _ in ordered))
    parts = TOKEN_RE.split(text)
    kept = TOKEN_RE.findall(text)

    def repl_gap(gap: str) -> str:
        return pattern.sub(lambda match: mapping[match.group(0)], gap)

    chunks = []
    for index, gap in enumerate(parts):
        chunks.append(repl_gap(gap))
        if index < len(kept):
            chunks.append(kept[index])
    return "".join(chunks)
