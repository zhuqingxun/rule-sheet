"""条件与输出格的解析。

- 单元格条件: FEEL 一元测试子集 (决策表前端)
- 条件组合: 全部满足 / 任一满足 / 不满足 (YAML 前端)
- 输出格: 不改 / 字面值 / 复制变量 / =表达式

空值约定: 空值参与任何比较都不成立 (SQL 风格); 判空必须显式写 `为空` / `非空`。
"""

from __future__ import annotations

import numbers
import re
from collections.abc import Callable
from datetime import date, datetime
from typing import Any

import pandas as pd

from .expressions import eval_expr, literal, split_top_level
from .model import BLANK_WORDS, Condition, Namespace

Mask = pd.Series

_RANGE_RE = re.compile(r"^([\[\(])\s*(.+?)\s*\.\.\s*(.+?)\s*([\]\)])$")
_CMP_RE = re.compile(r"^(>=|<=|!=|<>|==|=|>|<)\s*(.+)$")
_NULL_TESTS = ("为空", "非空")


def _both_present(var: pd.Series, rhs: Any) -> Mask:
    ok = var.notna()
    if isinstance(rhs, pd.Series):
        ok = ok & rhs.notna()
    return ok


def unary_test(cell: Any, var: pd.Series, ns: Namespace) -> Mask:
    """对变量 var 应用一个单元格条件, 返回布尔 Series。"""
    all_true = pd.Series(True, index=var.index)
    if cell is None:
        return all_true
    if isinstance(cell, bool) or isinstance(cell, numbers.Real):
        return (var == cell) & var.notna()
    if isinstance(cell, (datetime, date)):
        return (pd.to_datetime(var) == pd.Timestamp(cell)) & var.notna()

    s = str(cell).strip()
    if s in BLANK_WORDS:
        return all_true
    if s.startswith("not(") and s.endswith(")"):
        inner = s[4:-1].strip()
        neg = ~unary_test(inner, var, ns)
        # not(...) 只对有值的格成立; 对 为空/非空 本身取反时保持纯否定
        return neg if inner in _NULL_TESTS else neg & var.notna()
    parts = split_top_level(s)
    if len(parts) > 1:  # 枚举: 任一满足
        m = pd.Series(False, index=var.index)
        for p in parts:
            m = m | unary_test(p, var, ns)
        return m
    if s == "非空":
        return var.notna()
    if s == "为空":
        return var.isna()
    if m := _RANGE_RE.match(s):
        lo, hi = literal(m.group(2), ns, var), literal(m.group(3), ns, var)
        lo_ok = (var >= lo) if m.group(1) == "[" else (var > lo)
        hi_ok = (var <= hi) if m.group(4) == "]" else (var < hi)
        return lo_ok & hi_ok & var.notna()
    if m := _CMP_RE.match(s):
        op, rhs = m.group(1), literal(m.group(2), ns, var)
        ok = _both_present(var, rhs)
        if op in ("=", "=="):
            return (var == rhs) & ok
        if op in ("!=", "<>"):
            return (var != rhs) & ok
        if op == ">":
            return (var > rhs) & ok
        if op == ">=":
            return (var >= rhs) & ok
        if op == "<":
            return (var < rhs) & ok
        return (var <= rhs) & ok
    rhs = literal(s, ns, var)
    return (var == rhs) & _both_present(var, rhs)


def cell_condition(var_name: str, cell: Any) -> Condition:
    def cond(ns: Namespace) -> Mask:
        if var_name not in ns:
            raise KeyError(f"规则引用了未定义的变量「{var_name}」")
        return unary_test(cell, ns[var_name], ns)

    return cond


def expr_condition(expr: str) -> Condition:
    return lambda ns: eval_expr(expr, ns)


def to_mask(res: Any, ns: Namespace) -> Mask:
    index = next(iter(ns.values())).index
    if not isinstance(res, pd.Series):
        res = pd.Series(bool(res), index=index)
    return res.fillna(False).astype(bool)


def all_of(conds: list[Condition]) -> Condition:
    def cond(ns: Namespace) -> Mask:
        m = to_mask(True, ns)
        for c in conds:
            m = m & to_mask(c(ns), ns)
        return m

    return cond


def any_of(conds: list[Condition]) -> Condition:
    def cond(ns: Namespace) -> Mask:
        m = to_mask(False, ns)
        for c in conds:
            m = m | to_mask(c(ns), ns)
        return m

    return cond


def not_of(inner: Condition) -> Condition:
    return lambda ns: ~to_mask(inner(ns), ns)


COMBINATORS: dict[str, Callable[[list[Condition]], Condition]] = {"全部满足": all_of, "任一满足": any_of}


def build_condition(node: Any) -> Condition:
    """YAML 条件树: 字符串 = 一条表达式; 列表 = 全部满足; 字典 = 全部满足 / 任一满足 / 不满足, 可嵌套。"""
    if isinstance(node, str):
        return expr_condition(node)
    if isinstance(node, list):
        return all_of([build_condition(n) for n in node])
    if isinstance(node, dict) and len(node) == 1:
        key, val = next(iter(node.items()))
        if key in COMBINATORS:
            items = val if isinstance(val, list) else [val]
            return COMBINATORS[key]([build_condition(n) for n in items])
        if key == "不满足":
            return not_of(build_condition(val))
    raise ValueError(f"无法识别的条件写法: {node!r} (只认 表达式 / 全部满足 / 任一满足 / 不满足)")


def strings_in(node: Any) -> list[str]:
    """递归收集条件树里的所有表达式字符串, 用于扫描引用到的变量名。"""
    if isinstance(node, str):
        return [node]
    if isinstance(node, list):
        return [s for n in node for s in strings_in(n)]
    if isinstance(node, dict):
        return [s for v in node.values() for s in strings_in(v)]
    return []


def resolve_output(spec: Any, ns: Namespace) -> Any:
    """解析输出格。返回 None 表示不改。"""
    if spec is None:
        return None
    if isinstance(spec, str):
        s = spec.strip()
        if s in BLANK_WORDS:
            return None
        if s.startswith("="):
            return eval_expr(s[1:], ns)
        if s in ns:
            return ns[s]
        if len(s) >= 2 and s[0] == s[-1] and s[0] in "\"'":
            return s[1:-1]
        return s
    return spec
