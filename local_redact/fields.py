"""检测报告里要打码的栏目。

表格和正文都认这些标签。要增减栏目，改这里即可。
prefix 非空：换成稳定代号（同一原文始终同一个代号）。
drop 非空：换成「[某类已移除]」，原文只进本地对照表。
inline=False：只在表格标签格生效，不在句子里匹配，避免「审核：符合」被误伤。
"""

from __future__ import annotations

from dataclasses import dataclass
import re


@dataclass(frozen=True)
class Rule:
    prefix: str = ""
    drop: str = ""
    value_mode: str = "rest"
    inline: bool = True


def _org(inline: bool = True) -> Rule:
    return Rule(prefix="机构", inline=inline)


def _person(inline: bool = True) -> Rule:
    return Rule(prefix="人员", value_mode="person", inline=inline)


def _addr() -> Rule:
    return Rule(drop="地址")


def _phone() -> Rule:
    return Rule(drop="电话", value_mode="phone")


def _coord() -> Rule:
    return Rule(drop="坐标", value_mode="coord")


FORM_LABELS: dict[str, Rule] = {
    "委托单位名称": _org(),
    "委托单位": _org(),
    "委托方": _org(),
    "甲方": _org(),
    "受检单位名称": _org(),
    "受检单位": _org(),
    "受测单位": _org(),
    "被测单位": _org(),
    "被检单位": _org(),
    "被监测单位": _org(),
    "排污单位": _org(),
    "建设单位": _org(),
    "施工单位": _org(),
    "监理单位": _org(),
    "分包单位": _org(),
    "送样单位": _org(),
    "送检单位": _org(),
    "客户名称": _org(),
    "客户单位": _org(),
    "企业名称": _org(),
    "项目名称": Rule(prefix="项目"),
    "企业地址": _addr(),
    "单位地址": _addr(),
    "项目地址": _addr(),
    "采样地址": _addr(),
    "监测地址": _addr(),
    "委托地址": _addr(),
    "受检地址": _addr(),
    "项目地点": _addr(),
    # 机构通讯栏常见写法；须带冒号，避免叙述句里的「地址」被误伤
    "地址": _addr(),
    "通讯地址": _addr(),

    "联系人": _person(),
    "采样人": _person(),
    "采样人员": _person(),
    "检测人员": _person(),
    "监测人员": _person(),
    "采样/检测人员": _person(),
    "采样/监测人员": _person(),
    "编制人": _person(),
    "审核人": _person(),
    "签发人": _person(),
    "批准人": _person(),
    "授权签字人": _person(),
    "法定代表人": _person(),
    "法人代表": _person(),
    "项目负责人": _person(),
    "现场负责人": _person(),
    "企业陪同人": _person(),
    "企业陪同人员": _person(),
    "陪同人员": _person(),
    "陪同人": _person(),
    # 太短，只认表格里单独一格，不在正文句子里匹配
    "编制": _person(inline=False),
    "审核": _person(inline=False),
    "签发": _person(inline=False),
    "批准": _person(inline=False),
    "人员": _person(inline=False),
    "联系电话": _phone(),
    "联系人电话": _phone(),
    "联系方式": _phone(),
    "手机": _phone(),
    "电话": _phone(),
}

COLUMN_RULES: dict[str, Rule] = {
    "项目编号": Rule(prefix="项目号"),
    "工程编号": Rule(prefix="项目号"),
    "合同编号": Rule(prefix="合同"),
    "排污许可证编号": Rule(prefix="许可证"),
    "排污许可证号": Rule(prefix="许可证"),
    "统一社会信用代码": Rule(drop="信用代码"),
    "组织机构代码": Rule(drop="信用代码"),
    # 结果表人员列（表头在 HEADER_WORDS 里时按列打码）
    "采样人员": _person(),
    "检测人员": _person(),
    "监测人员": _person(),
    "采样/检测人员": _person(),
    "采样/监测人员": _person(),
    "采样人": _person(),
    "检测人": _person(),
    "监测人": _person(),
    "人员": _person(inline=False),
    "经纬度": _coord(),
    "地理坐标": _coord(),
    "点位坐标": _coord(),
    "坐标": _coord(),
    "经度": _coord(),
    "纬度": _coord(),
    # 纯小数平面坐标只在这些列去掉，不在正文里扫长数字
    "东坐标": _coord(),
    "北坐标": _coord(),
    "横坐标": _coord(),
    "纵坐标": _coord(),
    "X坐标": _coord(),
    "Y坐标": _coord(),
    "x坐标": _coord(),
    "y坐标": _coord(),
    "高斯坐标": _coord(),
    "平面坐标": _coord(),
    "大地坐标": _coord(),
}

LABELS: dict[str, Rule] = {**FORM_LABELS, **COLUMN_RULES}

# 结果表的表头。整行每个非空格都是表头时，才按列处理，避免把「委托单位 | 公司名」当成表头。
# 「序号」「样品编号」也算表头，否则人员列只会打到表头正下方那一格。
RESULT_HEADERS = frozenset(
    {
        "序号",
        "样品编号",
        "样品名称",
        "检测项目",
        "监测项目",
        "分析项目",
        "监测因子",
        "检测因子",
        "结果",
        "检测结果",
        "监测结果",
        "测定值",
        "实测浓度",
        "浓度",
        "排放浓度",
        "单位",
        "标准限值",
        "限值",
        "标准值",
        "检出限",
        "最低检出限",
        "方法依据",
        "方法标准",
        "检测方法",
        "仪器",
        "仪器名称",
        "仪器型号",
        "仪器编号",
        "分析标准",
        "评价标准",
        "执行标准",
        "采样时间",
        "采样日期",
        "分析日期",
        "收样日期",
        "检测时间",
        "监测时间",
        "采样/检测时间",
        "采样/监测时间",
        "点位",
        "点位名称",
        "监测点位",
        "采样点位",
        "点位编号",
        "频次",
        "检测频次",
        "监测频次",
        "样品状态",
        "样品介质",
        "样品介质/状态",
        "检测类别",
        "监测类别",
        "采样位置",
        "检测位置",
        "监测位置",
        "采样/检测位置",
        "采样/监测位置",
        "备注",
        "是否达标",
        "达标情况",
    }
)

HEADER_WORDS = frozenset(COLUMN_RULES) | RESULT_HEADERS

TOKEN_RE = re.compile(
    r"(?:机构|人员|项目|项目号|报告|样品|合同|许可证|名称)_[0-9a-f]{8}|\[[^\]]+已移除\]"
)


def normalize_label(text: str) -> str:
    compacted = re.sub(r"[\s\u3000]+", "", text or "")
    compacted = re.sub(r"[*＊]+", "", compacted)
    compacted = re.sub(r"[（(][^）)]{0,12}[）)]$", "", compacted)
    return compacted.strip("：:;；")
