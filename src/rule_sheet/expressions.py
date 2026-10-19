"""表达式求值与字面量解析。

中文变量名 (如 `主表.状态`、`本期.类别-会议`) 不是合法的 Python 标识符, 求值前先按名字长度降序替换成安全标识符,
再交给受限命名空间的 eval。替换发生在 Python 解析之前, 所以名字里含 `-`、括号等任何字符都行, 只要它在命名空间里。
规则文件视为可信输入。
"""

from __future__ import annotations

import re
from collections.abc import Iterable
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

# 变量名里允许的字符: 除空白和 ASCII 标点 (Python 的全部语法符号) 之外的任何字符, 外加下划线。
# 中文表头里常见的全角括号、全角横线等都算名字的一部分, 它们永远不会是 Python 语法。
NAME_CHAR = r"[^\s!-/:-@\[-^`{-~]"
_BACKTICK_RE = re.compile(r"`([^`]+)`")


def _ident(name: str) -> str:
    if name not in _IDENTS:
        _IDENTS[name] = f"_v{len(_IDENTS)}"
    return _IDENTS[name]


def eval_expr(expr: str, ns: Namespace) -> Any:
    """求值表达式。变量名按长度降序替换, 防止 `金额` 吃掉 `主表.金额`。

    反引号包裹的名字 (如 `本期.金额(万元)`) 原样整体查找, 给含特殊符号、启发式认不出的字段名用。
    """
    local: dict[str, Any] = {}

    def quoted(m: re.Match[str]) -> str:
        name = m.group(1).strip()
        if name not in ns:
            raise ValueError(f"表达式求值失败: {expr!r} (反引号里的变量「{name}」未定义)")
        local[_ident(name)] = ns[name]
        return _ident(name)

    text = _BACKTICK_RE.sub(quoted, expr)
    for name in sorted(ns, key=len, reverse=True):
        pattern = r"(?<![\w.])" + re.escape(name) + r"(?![\w.])"
        text, n = re.subn(pattern, _ident(name), text)
        if n:
            local[_ident(name)] = ns[name]
    local.update(FUNCS)
    try:
        return eval(text, {"__builtins__": {}}, local)  # noqa: S307
    except NameError as exc:
        raise ValueError(
            f"表达式求值失败: {expr!r} (「{exc.name}」既不是变量也不是可用函数; "
            f"字段名含特殊符号时可用反引号包裹整个名字, 如 `本期.类别-会议`)"
        ) from exc
    except Exception as exc:
        raise ValueError(f"表达式求值失败: {expr!r} ({exc})") from exc


def find_names(text: str, aliases: Iterable[str], known: Iterable[str] = ()) -> list[str]:
    """从表达式文本里找出引用到的 `<别名>.<列名>` 变量名 (YAML 前端用: 测试用例没给的变量据此置空)。

    列名允许含 `-` 与 `/` (中文表头常见, 如 类别-会议、岗位/职级)。两者后面紧跟 `别名.`、数字或已知名字
    (派生变量、函数) 时当作运算符, 否则当作名字的一部分; 宁可少认也不多认, 多认会把一个减法表达式整体当成空变量。
    反引号包裹的内容原样作为一个名字, 不受启发式限制。
    """
    names: list[str] = []

    def take_quoted(m: re.Match[str]) -> str:
        names.append(m.group(1).strip())
        return " "

    rest = _BACKTICK_RE.sub(take_quoted, text)
    alias_alt = "|".join(re.escape(a) for a in aliases)
    if not alias_alt:
        return names
    stops = [rf"(?:{alias_alt})\.", r"\d", *(re.escape(k) + rf"(?!{NAME_CHAR})" for k in known)]
    seg = NAME_CHAR + "+"
    rx = re.compile(rf"(?<![\w.])(?:{alias_alt})\.{seg}(?:[-/](?!{'|'.join(stops)}){seg})*")
    names.extend(m.group(0) for m in rx.finditer(rest))
    return names


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
