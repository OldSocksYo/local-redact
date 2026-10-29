"""格式固定的敏感信息。只打码能校验得上的号码，避免误伤检测数值和标准号。"""

from __future__ import annotations

import re
from collections.abc import Callable

USCC_ALPHABET = "0123456789ABCDEFGHJKLMNPQRTUWXY"
USCC_WEIGHTS = [1, 3, 9, 27, 19, 26, 16, 17, 20, 29, 25, 13, 8, 24, 10, 30, 28]

ID_WEIGHTS = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
ID_CHECKS = "10X98765432"

DropFn = Callable[[str, str], str]

_DMS = re.compile(
    r"\d{1,3}\s*[°度]\s*\d{1,2}\s*[′'’＇]\s*(?:\d{1,2}(?:\.\d+)?\s*[″\"”＂]?)?"
)
_LABELED_COORD = re.compile(
    r"(?:东经|西经|南纬|北纬|经度|纬度)\s*[:：]?\s*-?\d{1,3}(?:\.\d+)?\s*[°度]?"
)
_NE_COORD = re.compile(
    r"(?<![A-Za-z0-9])[NnSs]\s*-?\d{1,2}(?:\.\d+)?\s*[°度]?\s*[,，/\s]\s*[EeWw]\s*-?\d{1,3}(?:\.\d+)?\s*[°度]?"
)
_EW_COORD = re.compile(
    r"(?<![A-Za-z0-9])[EeWw]\s*-?\d{1,3}(?:\.\d+)?\s*[°度]?\s*[,，/\s]\s*[NnSs]\s*-?\d{1,2}(?:\.\d+)?\s*[°度]?"
)
_EMAIL = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
_URL = re.compile(r"https?://[^\s<>\"']+")
_ID_CARD = re.compile(r"(?<!\d)\d{17}[\dXx](?!\d)")
_USCC = re.compile(r"(?<![0-9A-Z])[0-9A-Z]{18}(?![0-9A-Z])")
# 前后都不能是字母/数字，避免仪器编号里「S138…」「NB023…」这类片段被当成电话
_MOBILE = re.compile(r"(?<![A-Za-z0-9])(?:\+?86[-\s]?)?1[3-9]\d{9}(?![A-Za-z0-9])")
_LANDLINE = re.compile(r"(?<![A-Za-z0-9])0\d{2,3}[-\s]?\d{7,8}(?![A-Za-z0-9])")
_LONG_DIGITS = re.compile(r"(?<!\d)\d{16,19}(?!\d)")
_SECRET_ASSIGN = re.compile(
    r"(?i)(?:api[_-]?key|secret|password|passwd|pwd|access[_-]?key)\s*[:=]\s*[^\s,;，；]{6,}"
)
_SK_KEY = re.compile(r"(?<![A-Za-z0-9])sk-[A-Za-z0-9]{16,}")
_PRIVATE_IP = re.compile(
    r"(?<!\d)(?:10\.\d{1,3}\.\d{1,3}\.\d{1,3}|192\.168\.\d{1,3}\.\d{1,3}|172\.(?:1[6-9]|2\d|3[0-1])\.\d{1,3}\.\d{1,3}|127\.\d{1,3}\.\d{1,3}\.\d{1,3})(?!\d)"
)


def uscc_check_char(body17: str) -> str:
    total = sum(USCC_ALPHABET.index(ch) * weight for ch, weight in zip(body17, USCC_WEIGHTS))
    return USCC_ALPHABET[(31 - total % 31) % 31]


def is_valid_uscc(value: str) -> bool:
    if len(value) != 18 or any(ch not in USCC_ALPHABET for ch in value):
        return False
    return uscc_check_char(value[:17]) == value[17]


def is_valid_id_card(value: str) -> bool:
    upper = value.upper()
    if re.fullmatch(r"\d{17}[\dX]", upper) is None:
        return False
    total = sum(int(num) * weight for num, weight in zip(upper[:17], ID_WEIGHTS))
    return ID_CHECKS[total % 11] == upper[-1]


def luhn_ok(number: str) -> bool:
    if not number.isdigit():
        return False
    checksum = 0
    double = False
    for char in reversed(number):
        digit = int(char)
        if double:
            digit *= 2
            if digit > 9:
                digit -= 9
        checksum += digit
        double = not double
    return checksum % 10 == 0


def _valid_ip(value: str) -> bool:
    parts = value.split(".")
    if len(parts) != 4:
        return False
    try:
        return all(0 <= int(part) <= 255 for part in parts)
    except ValueError:
        return False


def _sub_all(pattern: re.Pattern[str], text: str, kind: str, drop: DropFn) -> str:
    def repl(match: re.Match[str]) -> str:
        return drop(kind, match.group(0))

    return pattern.sub(repl, text)


def _sub_valid(
    pattern: re.Pattern[str],
    text: str,
    kind: str,
    drop: DropFn,
    ok,
) -> str:
    def repl(match: re.Match[str]) -> str:
        found = match.group(0)
        if not ok(found):
            return found
        return drop(kind, found)

    return pattern.sub(repl, text)


def mask_sensitive(text: str, drop: DropFn) -> str:
    """打码坐标、证件、联系方式和凭据。检测数值、日期、GB/HJ 标准号保持原样。"""
    text = _sub_all(_DMS, text, "坐标", drop)
    text = _sub_all(_LABELED_COORD, text, "坐标", drop)
    text = _sub_all(_NE_COORD, text, "坐标", drop)
    text = _sub_all(_EW_COORD, text, "坐标", drop)
    text = _sub_all(_EMAIL, text, "邮箱", drop)
    text = _sub_all(_URL, text, "链接", drop)
    text = _sub_valid(_ID_CARD, text, "身份证", drop, is_valid_id_card)
    text = _sub_valid(_USCC, text, "信用代码", drop, is_valid_uscc)
    text = _sub_all(_MOBILE, text, "电话", drop)
    text = _sub_all(_LANDLINE, text, "电话", drop)
    text = _sub_valid(_LONG_DIGITS, text, "银行卡", drop, luhn_ok)
    text = _sub_all(_SECRET_ASSIGN, text, "凭据", drop)
    text = _sub_all(_SK_KEY, text, "凭据", drop)
    text = _sub_valid(_PRIVATE_IP, text, "内网地址", drop, _valid_ip)
    return text
