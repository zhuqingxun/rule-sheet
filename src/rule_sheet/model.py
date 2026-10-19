"""数据结构与全局常量。"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

TRUE_WORDS = {"是", "true", "True", "TRUE", "yes"}
FALSE_WORDS = {"否", "false", "False", "FALSE", "no"}
BLANK_WORDS = {"", "-", "—", "任意"}
EXIST_VAR = "存在"  # 右表是否匹配到, 实际变量名是 <右表别名>.存在
BASE_DATE_VAR = "基准日期"

Namespace = dict[str, Any]
Condition = Callable[[Namespace], Any]


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
class RuleSet:
    match: dict[str, Any]
    derived: list[tuple[str, str]]  # (变量名, 表达式), 按定义顺序求值
    hit_policy: str  # 优先级 / 唯一
    rules: list[Rule]
    outputs: list[str]
    tests: list[TestCase] = field(default_factory=list)
    source: str = ""
    variables: list[str] = field(default_factory=list)  # 规则引用到的变量, 测试用例没给的视为空值

    @property
    def left_alias(self) -> str:
        return str(self.match.get("左表别名", "主表"))

    @property
    def right_alias(self) -> str:
        return str(self.match.get("右表别名", "本月"))
