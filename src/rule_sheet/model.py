"""数据结构与全局常量。"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

TRUE_WORDS = {"是", "true", "True", "TRUE", "yes"}
FALSE_WORDS = {"否", "false", "False", "FALSE", "no"}
BLANK_WORDS = {"", "-", "—", "任意"}
NULL_WORD = "空"  # 空值关键字: 条件里与 = / != 搭配判空 (本期.x == 空), 输出格 / 期望格 / 测试输入里就是空值
CLEAR_WORDS = {"清空", "置空"}  # 输出格: 把目标列写成空值 (区别于 - 的「不改」)
EXPECT_NULL_WORDS = CLEAR_WORDS | {"为空", NULL_WORD}  # 测试用例期望格: 期望该列为空 (留空仍是「不变」)
HIT_POLICIES = ("优先级", "唯一", "首次命中", "多条拼接")
JOIN_POLICY = "多条拼接"  # DMN 的 Collect: 每个输出列把所有命中规则的值按优先级顺序用分隔符拼起来 (多标签分类)
JOIN_SEP_KEY = "拼接分隔符"  # 「设置」里给 多条拼接 用的分隔符
DEFAULT_JOIN_SEP = "/"
DUP_POLICIES = ("取第一条", "报错", "拼接")  # 参照表匹配键重复时; 拼接 = 同键的行按排序接成一行 (每列一个文本, 0.7)
CONCAT_POLICY = "拼接"
# 重复键: 拼接 的默认分隔符。与 多条拼接 的 / 是两个概念: 一个接同一个键的多条记录, 一个接多条规则的输出
DEFAULT_CONCAT_SEP = ";"
CONCAT_OPTS = ("排序", "分隔符", "空值占位", "带序号")  # 输入声明里只在 重复键: 拼接 时允许写的四个可选项
APPEND_KEY = "独有行"  # 输入声明: 参照表里主表没有的行怎么办: 忽略 (默认) / 追加 (并入输出的行集末尾, 0.7, issue #19)
APPEND_WORDS = {"忽略": False, "追加": True}
PRIMARY_WORD = "主表"  # 「输入」里主表那一行的「匹配」写它: 输出的行就是主表的行, 其它输入都对齐到它
SETTING_KEYS = ("基准日期", "输出", JOIN_SEP_KEY)
FIELD_GROUP_KEY = "字段组"  # Excel sheet 名 / YAML 顶层键: 给一组列起个名字, 决策表里一格就能整组复制 / 清空
EXIST_VAR = "存在"  # 参照表是否匹配到, 实际变量名是 <输入名>.存在
BASE_DATE_VAR = "基准日期"

Namespace = dict[str, Any]
Condition = Callable[[Namespace], Any]


class _Clear:
    """resolve_output 的返回哨兵: 命中后把该列清空。None 仍表示「不改」。"""

    def __repr__(self) -> str:
        return "清空"


CLEAR = _Clear()


class _Null:
    """表达式里 空 关键字求值的结果。它不是一个值: 只能和 == / != 搭配 (就是判空), 别的运算一律报错。"""

    def __repr__(self) -> str:
        return "空"

    def __bool__(self) -> bool:
        return False


NULL = _Null()


@dataclass
class Rule:
    rule_id: str
    priority: float
    conditions: list[Condition]
    outputs: dict[str, Any]  # 输出列 -> 原始单元格值 / YAML 值
    desc: str = ""


@dataclass
class TestCase:
    __test__ = False  # 不是 pytest 用例, 防止被收集

    name: str
    inputs: dict[str, Any]
    expected: dict[str, Any]  # 期望列 -> 值; None 表示「不变」


@dataclass
class InputSpec:
    """规则声明的一个输入 (形参): 名字是规则里引用列的前缀 (<名字>.<列>)。

    文件是可选的默认绑定 (相对规则文件目录), 命令行 -i 名字=路径 可覆盖。主表之外的输入各自按 keys 对齐到主表。
    """

    name: str
    primary: bool = False
    keys: list[tuple[str, str]] = field(default_factory=list)  # (本输入的列, 主表的列); 主表为空
    dup_policy: str = "取第一条"  # 见 DUP_POLICIES
    path: str | None = None  # 可带 #sheet名
    # 以下只在 dup_policy == 拼接 时有意义: 同键的行先按 sort 排 (空值一律排末尾), 每列各自接成一个文本
    sort: list[tuple[str, bool]] = field(default_factory=list)  # [(列, 升序?)]
    sep: str = DEFAULT_CONCAT_SEP
    placeholder: str = ""  # 空值占位: 空格子换成它, 位置保留 (11 列之间按位置对齐); 留空 = 空文本
    numbered: bool = False  # 带序号: 1.a;2.b (占位符也带序号)
    # 独有行: 追加 — 这个参照表里键在主表找不到的行追加到输出行集末尾 (同名列带入, 键列按匹配映射填), 主表 .存在 为 否
    append: bool = False


@dataclass
class RuleSet:
    inputs: list[InputSpec]  # 恰有一个 primary
    settings: dict[str, Any]  # 基准日期 / 输出 (默认输出路径)
    derived: list[tuple[str, str]]  # (变量名, 表达式), 按定义顺序求值
    hit_policy: str  # 见 HIT_POLICIES
    rules: list[Rule]
    outputs: list[str]
    tests: list[TestCase] = field(default_factory=list)
    source: str = ""
    variables: list[str] = field(default_factory=list)  # 规则引用到的变量, 测试用例没给的视为空值
    # 字段组: 组名 -> 列名清单。加载时已把规则里写在组名下的输出格展开成逐列 (Rule.outputs 里没有组名),
    # 这里留着给导出 / 比对 / 运行时「组名与真实列撞名」的检查用
    field_groups: dict[str, list[str]] = field(default_factory=dict)

    @property
    def primary(self) -> InputSpec:
        return next(s for s in self.inputs if s.primary)

    @property
    def lookups(self) -> list[InputSpec]:
        return [s for s in self.inputs if not s.primary]

    @property
    def names(self) -> list[str]:
        return [s.name for s in self.inputs]

    def is_input_derived(self, name: str) -> bool:
        """带输入前缀的派生变量 (`报表7.现岗`): 在那张表自己的行上算, 发生在匹配 / 聚合之前; 用例里直接给值不推导。"""
        prefix, dot, rest = name.partition(".")
        return bool(dot and rest) and prefix in self.names

    @property
    def input_derived(self) -> list[tuple[str, str]]:
        return [(k, v) for k, v in self.derived if self.is_input_derived(k)]

    @property
    def plain_derived(self) -> list[tuple[str, str]]:
        return [(k, v) for k, v in self.derived if not self.is_input_derived(k)]
