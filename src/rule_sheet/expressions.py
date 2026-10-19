"""表达式求值与字面量解析。

中文变量名 (如 `主表.状态`) 不是合法的 Python 标识符, 求值前先按名字长度降序替换成安全标识符,
再交给受限命名空间的 eval。规则文件视为可信输入。
"""

from __future__ import annotations

import re
from typing import Any

import pandas as pd

from .model import FALSE_WORDS, TRUE_WORDS, Namespace


def _days_between(later: Any, earlier: Any) -> Any:
    delta = pd.to_datetime(later) - pd.to_datetime(earlier)
    return delta.dt.days if isinstance(delta, pd.Series) else delta.days


def _notna(x: Any) -> Any:
    return x.notna() if isinstance(x, pd.Series) else pd.notna(x)


def _isna(x: Any) -> Any:
    return x.isna() if isinstance(x, pd.Series) else pd.isna(x)


def _isin(x: Any, *values: Any) -> Any:
    return x.isin(values) if isinstance(x, pd.Series) else x in values


FUNCS: dict[str, Any] = {
    "abs": abs,
    "round": round,
    "min": min,
    "max": max,
    "天数差": _days_between,
    "days_between": _days_between,
    "非空": _notna,
    "为空": _isna,
    "属于": _isin,
    "是": True,
    "否": False,
    "true": True,
    "false": False,
}

_IDENTS: dict[str, str] = {}


def _ident(name: str) -> str:
    if name not in _IDENTS:
        _IDENTS[name] = f"_v{len(_IDENTS)}"
    return _IDENTS[name]


def eval_expr(expr: str, ns: Namespace) -> Any:
    """求值表达式。变量名按长度降序替换, 防止 `金额` 吃掉 `主表.金额`。"""
    text = expr
    local: dict[str, Any] = {}
    for name in sorted(ns, key=len, reverse=True):
        pattern = r"(?<![\w.])" + re.escape(name) + r"(?![\w.])"
        text, n = re.subn(pattern, _ident(name), text)
        if n:
            local[_ident(name)] = ns[name]
    local.update(FUNCS)
    try:
        return eval(text, {"__builtins__": {}}, local)  # noqa: S307
    except Exception as exc:
        raise ValueError(f"表达式求值失败: {expr!r} ({exc})") from exc


_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def split_top_level(s: str, sep: str = ",") -> list[str]:
    """按分隔符切分, 但跳过引号内和括号内的分隔符。"""
    parts: list[str] = []
    buf: list[str] = []
    depth = 0
    quote: str | None = None
    for ch in s:
        if quote:
            buf.append(ch)
            if ch == quote:
                quote = None
            continue
        if ch in "\"'":
            quote = ch
            buf.append(ch)
            continue
        if ch in "([":
            depth += 1
        elif ch in ")]":
            depth -= 1
        if ch == sep and depth == 0:
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    parts.append("".join(buf))
    return [p.strip() for p in parts if p.strip()]


def literal(token: str, ns: Namespace, var: pd.Series | None = None) -> Any:
    """把单元格里的一个 token 解析成值: 引号字符串 / 变量引用 / 布尔词 / 数字 / 日期 / 裸字符串。

    布尔词 (是/否) 只在被比较的变量本身是布尔列时才转成 True/False, 否则按普通文本处理。
    """
    t = token.strip()
    if len(t) >= 2 and t[0] == t[-1] and t[0] in "\"'":
        return t[1:-1]
    if t in ns:
        return ns[t]
    if var is not None and var.dtype == bool:
        if t in TRUE_WORDS:
            return True
        if t in FALSE_WORDS:
            return False
    try:
        f = float(t)
        return int(f) if f.is_integer() and "." not in t else f
    except ValueError:
        pass
    if _DATE_RE.match(t):
        return pd.Timestamp(t)
    return t
