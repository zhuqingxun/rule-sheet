"""表达式求值与字面量解析。

中文变量名 (如 `主表.状态`、`本期.类别-会议`) 不是合法的 Python 标识符, 求值前先按名字长度降序替换成安全标识符,
再交给受限命名空间的 eval。替换发生在 Python 解析之前, 所以名字里含 `-`、括号等任何字符都行, 只要它在命名空间里。

求值前还有一道 AST 改写, 让「Python 语法、整列运算」对写规则的人不再有坑:
    - `且` / `或` / `非` (及 and / or / not) 按元素逐行运算, 不会撞上 pandas 的「整列真值不明确」
    - 比较 (`== != > >= < <=` / `in` / `not in`) 任一侧为空就不成立, 与决策表单元格的空值约定一致
      (pandas 原生的 `!=` 对空值返回 True, Python 标量 `None == None` 为 True)
    - 链式比较 `0.02 <= x <= 0.05` 拆成两段 AND
规则文件视为可信输入。
"""

from __future__ import annotations

import ast
import copy
import operator
import re
from collections.abc import Callable, Iterable
from functools import lru_cache
from types import CodeType
from typing import Any

import pandas as pd

from .model import FALSE_WORDS, TRUE_WORDS, Namespace

# ---------------------------------------------------------------------------
# 空值与真值
# ---------------------------------------------------------------------------


def is_na(x: Any) -> bool:
    """标量是否为空值 (None / NaN / NaT); 列表、元组等容器不算空。"""
    if x is None:
        return True
    try:
        return bool(pd.isna(x))
    except (TypeError, ValueError):
        return False


def truth(x: Any) -> Any:
    """真值归一: Series -> 布尔 Series (空值算 False); 标量 -> bool。

    不能直接 `astype(bool)`: NaN 转 bool 是 True。
    """
    if isinstance(x, pd.Series):
        if x.dtype == bool:
            return x
        present = x.notna()
        return x[present].astype(bool).reindex(x.index, fill_value=False)
    return False if is_na(x) else bool(x)


def _present(x: Any) -> Any:
    return x.notna() if isinstance(x, pd.Series) else not is_na(x)


# ---------------------------------------------------------------------------
# 可用函数
# ---------------------------------------------------------------------------


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

# ---------------------------------------------------------------------------
# AST 改写: 逻辑连接词按元素运算, 比较空值不成立
# ---------------------------------------------------------------------------


def _all(*xs: Any) -> Any:
    out: Any = True
    for x in xs:
        out = out & truth(x)
    return out


def _any(*xs: Any) -> Any:
    out: Any = False
    for x in xs:
        out = out | truth(x)
    return out


def _not(x: Any) -> Any:
    t = truth(x)
    return ~t if isinstance(t, pd.Series) else not t


def _cmp(op: Callable[[Any, Any], Any]) -> Callable[[Any, Any], Any]:
    """比较: 任一侧为空就不成立 (SQL 风格, 与单元格一致)。

    pandas 原生只有 != 对空值返回 True; Python 标量 None == None 为 True、None > 1 直接报错。
    """

    def f(a: Any, b: Any) -> Any:
        return truth(op(a, b)) & _present(a) & _present(b)

    return f


def _values(b: Any) -> list[Any]:
    if isinstance(b, pd.Series):
        return b.tolist()
    if isinstance(b, (list, tuple, set, frozenset)):
        return list(b)
    return [b]


def _in(a: Any, b: Any) -> Any:
    vals = _values(b)
    if isinstance(a, pd.Series):
        return a.isin(vals) & a.notna()
    return (not is_na(a)) and a in vals


def _not_in(a: Any, b: Any) -> Any:
    vals = _values(b)
    if isinstance(a, pd.Series):
        return ~a.isin(vals) & a.notna()
    return (not is_na(a)) and a not in vals


_HELPERS: dict[str, Any] = {
    "_all": _all,
    "_any": _any,
    "_not": _not,
    "_in": _in,
    "_not_in": _not_in,
    "_eq": _cmp(operator.eq),
    "_ne": _cmp(operator.ne),
    "_lt": _cmp(operator.lt),
    "_le": _cmp(operator.le),
    "_gt": _cmp(operator.gt),
    "_ge": _cmp(operator.ge),
}
_CMP_CALLS: dict[type, str] = {
    ast.Eq: "_eq",
    ast.NotEq: "_ne",
    ast.Lt: "_lt",
    ast.LtE: "_le",
    ast.Gt: "_gt",
    ast.GtE: "_ge",
    ast.In: "_in",
    ast.NotIn: "_not_in",
}


def _call(name: str, args: list[ast.expr], at: ast.AST) -> ast.expr:
    return ast.copy_location(ast.Call(ast.Name(name, ast.Load()), args, []), at)


class _Rewrite(ast.NodeTransformer):
    def visit_BoolOp(self, node: ast.BoolOp) -> ast.expr:
        self.generic_visit(node)
        return _call("_all" if isinstance(node.op, ast.And) else "_any", node.values, node)

    def visit_UnaryOp(self, node: ast.UnaryOp) -> ast.expr:
        self.generic_visit(node)
        if isinstance(node.op, ast.Not):
            return _call("_not", [node.operand], node)
        return node

    def visit_Compare(self, node: ast.Compare) -> ast.expr:
        self.generic_visit(node)
        chained = len(node.ops) > 1
        parts: list[ast.expr] = []
        left = node.left
        for op, right in zip(node.ops, node.comparators, strict=True):
            lhs, rhs = (copy.deepcopy(left), copy.deepcopy(right)) if chained else (left, right)
            fn = _CMP_CALLS.get(type(op))
            if fn:
                parts.append(_call(fn, [lhs, rhs], node))
            else:
                parts.append(ast.copy_location(ast.Compare(lhs, [op], [rhs]), node))
            left = right
        return parts[0] if not chained else _call("_all", parts, node)


@lru_cache(maxsize=512)
def _compile(text: str) -> CodeType:
    tree = ast.parse(text, mode="eval")
    tree = ast.fix_missing_locations(_Rewrite().visit(tree))
    return compile(tree, "<rule-sheet>", "eval")


# ---------------------------------------------------------------------------
# 名字替换与求值
# ---------------------------------------------------------------------------

_IDENTS: dict[str, str] = {}

# 变量名里允许的字符: 除空白和 ASCII 标点 (Python 的全部语法符号) 之外的任何字符, 外加下划线。
# 中文表头里常见的全角括号、全角横线等都算名字的一部分, 它们永远不会是 Python 语法。
NAME_CHAR = r"[^\s!-/:-@\[-^`{-~]"
_BACKTICK_RE = re.compile(r"`([^`]+)`")
_STRING_RE = re.compile(r'"(?:[^"\\]|\\.)*"|\'(?:[^\'\\]|\\.)*\'')
# 中文连接词: 独立成词 (两侧不是名字字符) 才算, `非空(` 里的 非 不受影响
_KEYWORD_RE = re.compile(r"(?<!\w)(且|或|非)(?!\w)")
_KEYWORDS = {"且": " and ", "或": " or ", "非": " not "}


def _ident(name: str) -> str:
    if name not in _IDENTS:
        _IDENTS[name] = f"_v{len(_IDENTS)}"
    return _IDENTS[name]


def outside_strings(text: str, fn: Callable[[str], str]) -> str:
    """对引号字符串之外的片段逐段应用 fn, 字符串字面量原样保留。"""
    out: list[str] = []
    pos = 0
    for m in _STRING_RE.finditer(text):
        out.append(fn(text[pos : m.start()]))
        out.append(m.group(0))
        pos = m.end()
    out.append(fn(text[pos:]))
    return "".join(out)


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

    def substitute(seg: str) -> str:
        seg = _BACKTICK_RE.sub(quoted, seg)
        for name in sorted(ns, key=len, reverse=True):
            pattern = r"(?<![\w.])" + re.escape(name) + r"(?![\w.])"
            seg, n = re.subn(pattern, _ident(name), seg)
            if n:
                local[_ident(name)] = ns[name]
        return _KEYWORD_RE.sub(lambda m: _KEYWORDS[m.group(1)], seg)

    text = outside_strings(expr, substitute).strip()  # 开头的 非 换成 not 后会带出前导空格, ast 会当成缩进
    local.update(FUNCS)
    local.update(_HELPERS)
    try:
        code = _compile(text)
    except SyntaxError as exc:
        raise ValueError(f"表达式语法错误: {expr!r} ({exc.msg})") from exc
    try:
        return eval(code, {"__builtins__": {}}, local)  # noqa: S307
    except NameError as exc:
        raise ValueError(
            f"表达式求值失败: {expr!r} (「{exc.name}」既不是变量也不是可用函数; "
            f"字段名含特殊符号时可用反引号包裹整个名字, 如 `本期.金额(万元)`)"
        ) from exc
    except ValueError as exc:
        if "ambiguous" in str(exc):
            raise ValueError(
                f"表达式求值失败: {expr!r} (对整列做了真假判断; 条件组合用 且 / 或 / 非, 不支持 if-else 三元写法)"
            ) from exc
        raise ValueError(f"表达式求值失败: {expr!r} ({exc})") from exc
    except Exception as exc:
        raise ValueError(f"表达式求值失败: {expr!r} ({exc})") from exc


def find_names(text: str, aliases: Iterable[str], known: Iterable[str] = ()) -> list[str]:
    """从表达式文本里找出引用到的 `<别名>.<列名>` 变量名 (YAML 前端用: 测试用例没给的变量据此置空)。

    列名允许含 `-` 与 `/` (中文表头常见, 如 类别-会议、岗位/职级)。两者后面紧跟 `别名.`、数字或已知名字
    (派生变量、函数) 时当作运算符, 否则当作名字的一部分; 宁可少认也不多认, 多认会把一个减法表达式整体当成空变量。
    反引号包裹的内容原样作为一个名字, 不受启发式限制。引号字符串里的内容不扫描。
    """
    names: list[str] = []

    def take_quoted(m: re.Match[str]) -> str:
        names.append(m.group(1).strip())
        return " "

    rest = outside_strings(text, lambda seg: _BACKTICK_RE.sub(take_quoted, seg))
    rest = _STRING_RE.sub(" ", rest)
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
