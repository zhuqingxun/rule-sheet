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
import numbers
import operator
import re
from collections.abc import Callable, Iterable
from datetime import date, datetime
from functools import lru_cache
from types import CodeType
from typing import Any

import pandas as pd

from .model import FALSE_WORDS, NULL, NULL_WORD, TRUE_WORDS, Namespace

# ---------------------------------------------------------------------------
# 空值与真值
# ---------------------------------------------------------------------------


def is_na(x: Any) -> bool:
    """标量是否为空值 (None / NaN / NaT / 关键字 空); 列表、元组等容器不算空。"""
    if x is None or x is NULL:
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
        if not x.index.is_unique:  # 引擎里索引总是唯一的 (RangeIndex); Python API 自带的 ns 可能不是, reindex 会抛
            return present & x.map(lambda v: not is_na(v) and bool(v))
        return x[present].astype(bool).reindex(x.index, fill_value=False)
    return False if is_na(x) else bool(x)


def _present(x: Any) -> Any:
    return x.notna() if isinstance(x, pd.Series) else not is_na(x)


def text_of(x: Any) -> str:
    """非空值的文本形态, 拼接 / 子串匹配 / 审计列共用: 整数不带 .0, 日期 yyyy-mm-dd, 布尔 是/否。"""
    if isinstance(x, bool):
        return "是" if x else "否"
    if isinstance(x, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(x).strftime("%Y-%m-%d")
    if isinstance(x, numbers.Real) and float(x).is_integer():
        return str(int(x))
    return str(x)


def _texts(x: Any) -> Any:
    """Series -> 文本 Series (空值保持 None); 标量 -> str 或 None。

    不能用 Series.map: 它会把返回的 None 变成 NaN, 后面的 `is None` 判断就失效。
    """
    if isinstance(x, pd.Series):
        return pd.Series([None if is_na(v) else text_of(v) for v in x.tolist()], index=x.index, dtype=object)
    return None if is_na(x) else text_of(x)


def _rowwise(fn: Callable[..., Any], *args: Any) -> Any:
    """把 fn 逐行应用到若干参数上: Series 按行取值, 标量广播; 全是标量时直接调用。"""
    series = [a for a in args if isinstance(a, pd.Series)]
    if not series:
        return fn(*args)
    index = series[0].index
    cols = [a.tolist() if isinstance(a, pd.Series) else [a] * len(index) for a in args]
    return pd.Series([fn(*row) for row in zip(*cols, strict=True)], index=index)


def _reject_null(fname: str, *args: Any) -> None:
    if any(a is NULL for a in args):
        raise ValueError(f"{fname}() 的参数不能是 空 (空 只能和 == / != 搭配)")


# ---------------------------------------------------------------------------
# 可用函数
# ---------------------------------------------------------------------------


def _days_between(later: Any, earlier: Any) -> Any:
    delta = pd.to_datetime(later) - pd.to_datetime(earlier)
    return delta.dt.days if isinstance(delta, pd.Series) else delta.days


def _days(n: Any) -> pd.Timedelta:
    """时长字面量 `365天` 改写成 _days(365): 基准日期 - 365天 / 本期.推荐时间 + 30天。"""
    return pd.Timedelta(days=float(n))


def _text_test(fname: str, pred: Callable[[str, str], bool], *, negate: bool = False) -> Callable[..., Any]:
    """子串类函数: 包含(x, 文本, ...) 任一命中即真; 否定形式 (不包含) 要求 x 有值且全部不命中。

    空值参与不成立 (与比较一致): x 为空, 或否定形式下任一文本为空, 结果都是 False。文本按字面匹配, 不是正则。
    非文本列先转成文本 (整数不带 .0, 日期 yyyy-mm-dd), 所以 包含(编码, "00") 对数值列也能用。
    """

    def one(v: str | None, *needles: str | None) -> bool:
        if v is None or (negate and any(n is None for n in needles)):
            return False
        hit = any(n is not None and pred(v, n) for n in needles)
        return (not hit) if negate else hit

    def f(x: Any, *needles: Any) -> Any:
        if not needles:
            raise ValueError(f"{fname}(x, 文本, ...) 至少要给一个文本")
        _reject_null(fname, x, *needles)
        return _rowwise(one, _texts(x), *[_texts(n) for n in needles])

    return f


def _concat(*parts: Any) -> Any:
    """拼接(a, b, ...): 空值当作空字符串; 全部参数都为空时结果是空值 (而不是空串, 否则审计列与写出的文件对不上)。"""
    if not parts:
        raise ValueError("拼接() 至少要给一个参数")
    _reject_null("拼接", *parts)

    def one(*vals: str | None) -> str | None:
        texts = [v for v in vals if v is not None]
        return "".join(texts) if texts else None

    return _rowwise(one, *[_texts(p) for p in parts])


_NUMBER_RE = re.compile(r"-?\d+(?:\.\d+)?")


def _number(x: Any) -> Any:
    """数字(x): 文本里的第一个数 (22级 / P23 / 第3档 -> 22 / 23 / 3), 取不到为空; 数值原样返回。

    逐行生效, 是不开放的 float() / int() 的替代。
    """
    _reject_null("数字", x)

    def one(v: Any) -> Any:
        if is_na(v):
            return None
        if isinstance(v, numbers.Real) and not isinstance(v, bool):
            return v
        m = _NUMBER_RE.search(str(v))
        if not m:
            return None
        f = float(m.group(0))
        return int(f) if f.is_integer() else f

    res = _rowwise(one, x)
    return pd.to_numeric(res) if isinstance(res, pd.Series) else res


def _if(cond: Any, a: Any, b: Any) -> Any:
    """如果(条件, 真值, 假值): 向量化三元。条件走真值归一 (空值算假), 分支可以是标量 / 列 / 空。"""
    t = truth(cond)
    series = [v for v in (a, b) if isinstance(v, pd.Series)]
    if not isinstance(t, pd.Series):
        if not series:
            return a if t else b
        t = pd.Series(t, index=series[0].index)
    index = t.index

    def column(v: Any) -> pd.Series:
        if v is NULL:
            return pd.Series([None] * len(index), index=index, dtype=object)
        if isinstance(v, pd.Series):
            return v
        return pd.Series([v] * len(index), index=index)

    return column(a).where(t, column(b))


def _notna(x: Any) -> Any:
    return x.notna() if isinstance(x, pd.Series) else not is_na(x)


def _isna(x: Any) -> Any:
    return x.isna() if isinstance(x, pd.Series) else is_na(x)


def _values(b: Any) -> tuple[list[Any], bool]:
    """枚举值列表, 以及其中是否含关键字 空 (属于(x, "a", 空) = x 是 a 或 x 为空)。"""
    if isinstance(b, pd.Series):
        vals = b.tolist()
    elif isinstance(b, (list, tuple, set, frozenset)):
        vals = list(b)
    else:
        vals = [b]
    return [v for v in vals if v is not NULL], any(v is NULL for v in vals)


def _in(a: Any, b: Any) -> Any:
    vals, has_null = _values(b)
    if isinstance(a, pd.Series):
        m = a.isin(vals) & a.notna()
        return m | a.isna() if has_null else m
    return has_null if is_na(a) else a in vals


def _not_in(a: Any, b: Any) -> Any:
    vals, _ = _values(b)  # 空值的 not in 永远不成立: 列表里有 空 时它在列表里, 没有时空值比较不成立
    if isinstance(a, pd.Series):
        return ~a.isin(vals) & a.notna()
    return (not is_na(a)) and a not in vals


def _isin(x: Any, *values: Any) -> Any:
    return _in(x, values)


# 决策表单元格也认这几个关键字: `包含 终端, 法务` = 包含(列, "终端", "法务")
TEXT_FUNCS: dict[str, Callable[..., Any]] = {
    "包含": _text_test("包含", lambda v, n: n in v),
    "不包含": _text_test("不包含", lambda v, n: n in v, negate=True),
    "开头是": _text_test("开头是", lambda v, n: v.startswith(n)),
    "开头不是": _text_test("开头不是", lambda v, n: v.startswith(n), negate=True),
    "结尾是": _text_test("结尾是", lambda v, n: v.endswith(n)),
    "结尾不是": _text_test("结尾不是", lambda v, n: v.endswith(n), negate=True),
}

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
    **TEXT_FUNCS,
    "拼接": _concat,
    "数字": _number,
    "如果": _if,
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


def _cmp(op: Callable[[Any, Any], Any], symbol: str) -> Callable[[Any, Any], Any]:
    """比较: 任一侧为空就不成立 (SQL 风格, 与单元格一致)。

    pandas 原生只有 != 对空值返回 True; Python 标量 None == None 为 True、None > 1 直接报错。
    关键字 空 是唯一的例外: `x == 空` / `x != 空` 就是判空 (FEEL 对 null 的相等也是这么特例的), 其它运算符碰到它报错。
    """

    def f(a: Any, b: Any) -> Any:
        if a is NULL or b is NULL:
            if op not in (operator.eq, operator.ne):
                raise ValueError(f"空 只能和 == / != 搭配, 不能用 {symbol}")
            other = b if a is NULL else a
            return _isna(other) if op is operator.eq else _notna(other)
        a_s, b_s = isinstance(a, pd.Series), isinstance(b, pd.Series)
        if not a_s and not b_s:  # 标量: 先查空再比较, None > 1 在 Python 里直接报错
            return not is_na(a) and not is_na(b) and bool(op(a, b))
        if (not a_s and is_na(a)) or (not b_s and is_na(b)):
            return pd.Series(False, index=(a if a_s else b).index)
        return truth(op(a, b)) & _present(a) & _present(b)

    return f


_HELPERS: dict[str, Any] = {
    "_all": _all,
    "_any": _any,
    "_not": _not,
    "_in": _in,
    "_not_in": _not_in,
    "_days": _days,
    "_eq": _cmp(operator.eq, "=="),
    "_ne": _cmp(operator.ne, "!="),
    "_lt": _cmp(operator.lt, "<"),
    "_le": _cmp(operator.le, "<="),
    "_gt": _cmp(operator.gt, ">"),
    "_ge": _cmp(operator.ge, ">="),
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
# 时长字面量 `365天` / `1.5天`: 数字紧跟 天, 且 天 后面不再是名字字符 (`天数差(` 不受影响);
# 变量名已先换成 _vN, 列名含「天」不受影响
_DAYS_RE = re.compile(r"(?<![\w.])(\d+(?:\.\d+)?)\s*天(?!\w)")


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
        seg = _DAYS_RE.sub(r"_days(\1)", seg)
        return _KEYWORD_RE.sub(lambda m: _KEYWORDS[m.group(1)], seg)

    text = outside_strings(expr, substitute).strip()  # 开头的 非 换成 not 后会带出前导空格, ast 会当成缩进
    local.update(FUNCS)
    local.update(_HELPERS)
    local[NULL_WORD] = NULL
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
                f"表达式求值失败: {expr!r} (对整列做了真假判断; 条件组合用 且 / 或 / 非, "
                f"三元写法用 如果(条件, 真值, 假值) 而不是 if-else)"
            ) from exc
        raise ValueError(f"表达式求值失败: {expr!r} ({exc})") from exc
    except TypeError as exc:
        if "_Null" in str(exc):  # 空 + 1 / abs(空) 之类: 空 不是值
            raise ValueError(f"表达式求值失败: {expr!r} (空 只能和 == / != 搭配, 不能参与运算或传给函数)") from exc
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


def dangling_name(token: str, ns: Namespace) -> bool:
    """长得像变量名 (<已知别名>.<列>) 却不在命名空间里: 多半是列名写错或表里没这一列。

    单元格 / 输出格里这种 token 不能当成字面文本静默参与比较或写出。
    """
    if "." not in token or token in ns:
        return False
    prefix = token.split(".", 1)[0] + "."
    return any(k.startswith(prefix) for k in ns)


def undefined_variable(token: str) -> KeyError:
    return KeyError(f"规则引用了未定义的变量「{token}」(两表列名和变量定义里都没有它; 想当成文字用请加引号)")


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
    if t == NULL_WORD:
        return NULL
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
    if dangling_name(t, ns):
        raise undefined_variable(t)
    return t
