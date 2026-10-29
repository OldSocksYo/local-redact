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
_LATIN_NAME = re.compile(r"^[A-Za-z][A-Za-z·•.\-]{1,19}$")
_CN_NAME = re.compile(r"^[\u4e00-\u9fff]{2,4}$")
_CN_NAME_RUN = re.compile(r"^(?:[\u4e00-\u9fff]{2}){2,12}$")
# 说明句/列举句特征：有这些就不当人员格处理
_PROSE_PUNCT = re.compile(r"[：:。；;（）()《》【】]")
# 《百家姓》单字姓。复姓走 _COMPOUND_SURNAMES，不把复姓拆字扩进来，以免误伤普通词。
_SURNAMES = frozenset(
    "赵钱孙李周吴郑王冯陈褚卫蒋沈韩杨朱秦尤许何吕施张孔曹严华金魏陶姜戚谢邹喻柏水窦章"
    "云苏潘葛奚范彭郎鲁韦昌马苗凤花方俞任袁柳酆鲍史唐费廉岑薛雷贺倪汤滕殷罗毕郝邬安常"
    "乐于时傅皮卞齐康伍余元卜顾孟平黄和穆萧尹姚邵湛汪祁毛禹狄米贝明臧计伏成戴谈宋茅"
    "庞熊纪舒屈项祝董梁杜阮蓝闵席季麻强贾路娄危江童颜郭梅盛林刁钟徐邱骆高夏蔡田樊胡"
    "凌霍虞万支柯昝管卢莫经房裘缪干解应宗丁宣贲邓郁单杭洪包诸左石崔吉钮龚程嵇邢滑裴"
    "陆荣翁荀羊於惠甄麴家封芮羿储靳汲邴糜松井段富巫乌焦巴弓牧隗山谷车侯宓蓬全郗班仰"
    "秋仲伊宫宁仇栾暴甘钭厉戎祖武符刘景詹束龙叶幸司韶郜黎蓟薄印宿白怀蒲邰从鄂索咸籍"
    "赖卓蔺屠蒙池乔阴鬱胥能苍双闻莘党翟谭贡劳逄姬申扶堵冉宰郦雍舄璩桑桂濮牛寿通边扈"
    "燕冀郏浦尚农温别庄晏柴瞿阎充慕连茹习宦艾鱼容向古易慎戈廖庾终暨居衡步都耿满弘匡"
    "国文寇广禄阙东欧殳沃利蔚越夔隆师巩厍聂晁勾敖融冷訾辛阚那简饶空曾毋沙乜养鞠须丰"
    "巢关蒯相查後荆红游竺权逯盖益桓公"
)
# 证件上常见、但《百家姓》写成另一个字的姓：肖/萧、闫/阎、付/傅。
_SURNAMES = _SURNAMES | frozenset("肖闫付覃丘涂朴麦佟冼")
_COMPOUND_SURNAMES = frozenset(
    {
        "欧阳",
        "太史",
        "端木",
        "上官",
        "司马",
        "东方",
        "独孤",
        "南宫",
        "万俟",
        "闻人",
        "夏侯",
        "诸葛",
        "尉迟",
        "公羊",
        "赫连",
        "澹台",
        "皇甫",
        "宗政",
        "濮阳",
        "公冶",
        "太叔",
        "申屠",
        "公孙",
        "慕容",
        "仲孙",
        "钟离",
        "长孙",
        "宇文",
        "司徒",
        "鲜于",
        "司空",
        "闾丘",
        "子车",
        "亓官",
        "司寇",
        "巫马",
        "公西",
        "颛孙",
        "壤驷",
        "公良",
        "漆雕",
        "乐正",
        "宰父",
        "谷梁",
        "拓跋",
        "夹谷",
        "轩辕",
        "令狐",
        "段干",
        "百里",
        "东郭",
        "南门",
        "呼延",
        "羊舌",
        "微生",
        "梁丘",
        "左丘",
        "东门",
        "西门",
        "南宫",
        "第五",
    }
)


def _is_prose_cell(value: str) -> bool:
    """整格像说明/句子，而不是人名或人名列表。"""
    text = value.strip()
    if _PROSE_PUNCT.search(text):
        return True
    if "备注" in text or "注明" in text or text.endswith("等"):
        return True
    return False


def looks_like_person_name(value: str) -> bool:
    """像中国姓名（首字为常见姓）或拉丁姓名；不维护业务词黑名单。

    人员格里常见「李伟田超吴广曾跃杰…」二字/三字名无分隔粘连，无法可靠切分时
    仍把整段当人员打码（一个代号），避免真名残留。
    """
    text = value.strip()
    if not text or len(text) > 40:
        return False
    if _LATIN_NAME.fullmatch(text):
        return True
    if _CN_NAME.fullmatch(text):
        if text[:2] in _COMPOUND_SURNAMES:
            return True
        return text[0] in _SURNAMES
    if _CN_NAME_RUN.fullmatch(text):
        # 全是二字名粘连：李伟田超吴广
        return all(text[i] in _SURNAMES for i in range(0, len(text), 2))
    # 二字/三字名混粘（长度为奇数或无法均匀双字切）：整段打码
    if re.fullmatch(r"[\u4e00-\u9fff]{4,40}", text):
        if text[:2] in _COMPOUND_SURNAMES or text[0] in _SURNAMES:
            return True
    return False


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
        # 人员列：先排除说明句，再按人名启发式打码（换行/顿号分隔的多人）
        if rule.value_mode == "person":
            if _is_prose_cell(value):
                return value
            if re.search(r"[\n\r、，,;；/／]", value):
                return self._apply_person_list(value, rule)
            if not looks_like_person_name(value):
                return value
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
            if len(person) < 2 or not looks_like_person_name(person):
                out.append(chunk)
                continue
            out.append(self._apply_one(person, rule))
        return "".join(out)

    def redact_string(self, text: str) -> str:
        if not text:
            return text
        updated = fold_digits(text)
        updated = self._inline(updated)
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
