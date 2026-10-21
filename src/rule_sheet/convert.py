"""两种写法之间的转换与比对。

- Excel 决策表 -> YAML (export_workbook): 无损。单元格文法的每一种形态都对应一条表达式, 翻译函数对照
  conditions.unary_test 逐分支写, 续行翻成 任一满足。导出的 YAML 是生成物: 源仍是 Excel, 改了 Excel 重新导出。
- 两份规则文件的语义比对 (diff_rules): 结构逐项比 (输入 / 设置 / 变量 / 命中策略 / 规则 / 输出列 / 用例);
  条件是编译好的函数比不了文本, 改为把两边的测试用例合并, 每条在两套规则上各跑一次, 比实际输出和命中 ID
  (按用例交叉验证, 是抽样等价, 不是证明)。
"""

from __future__ import annotations

import numbers
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

import yaml

from .conditions import _CMP_RE, _NULL_TESTS, _RANGE_RE, _TEXT_TEST_RE
from .engine import apply_rules, case_namespace, fmt, same
from .expressions import FUNCS, split_top_level
from .loaders import RawRule, load_excel_raw, load_rules
from .model import (
    BASE_DATE_VAR,
    BLANK_WORDS,
    EXIST_VAR,
    FALSE_WORDS,
    NULL_WORD,
    PRIMARY_WORD,
    TRUE_WORDS,
    RuleSet,
    TestCase,
)

_DATE_TEXT_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
_CMP_OPS = {"=": "==", "==": "==", "<>": "!=", "!=": "!=", ">": ">", ">=": ">=", "<": "<", "<=": "<="}
# 名字里出现这些就得用反引号包起来: ASCII 标点 (除 . - /)、空白, 以及 - 或 / 后面紧跟数字 (会被当成减法 / 除法)
_NEEDS_BACKTICKS_RE = re.compile(r"[\s!-,:-@\[-^`{-~]|[-/]\d")


def _quote(s: str) -> str:
    if '"' not in s:
        return f'"{s}"'
    if "'" not in s:
        return f"'{s}'"
    return '"' + s.replace('"', '\\"') + '"'


def _number_text(x: Any) -> str:
    if isinstance(x, bool):
        return "是" if x else "否"
    if isinstance(x, numbers.Real) and float(x).is_integer():
        return str(int(x))
    return repr(x)


def _date_text(x: Any) -> str:
    ts = datetime(x.year, x.month, x.day) if not isinstance(x, datetime) else x
    text = ts.strftime("%Y-%m-%d") if (ts.hour, ts.minute, ts.second) == (0, 0, 0) else ts.strftime("%Y-%m-%d %H:%M:%S")
    return f'日期("{text}")'


class CellTranslator:
    """单元格条件 -> 表达式文本。知道的名字: 输入名前缀、派生变量、<参照表>.存在、基准日期。"""

    def __init__(self, rs: RuleSet) -> None:
        self.inputs = set(rs.names)
        self.derived = {k for k, _ in rs.derived}
        self.exists = {f"{s.name}.{EXIST_VAR}" for s in rs.lookups}
        self.known = list(self.derived) + list(FUNCS) + [BASE_DATE_VAR]

    def is_var(self, token: str) -> bool:
        if token in self.derived or token == BASE_DATE_VAR or token in self.exists:
            return True
        prefix, dot, rest = token.partition(".")
        return bool(dot and rest) and prefix in self.inputs

    def is_boolish(self, var: str) -> bool:
        """这些列上的 是 / 否 是布尔 (运行时 literal() 看的是列的 dtype, 导出时没有数据, 按这个约定走)。"""
        return var in self.exists or var in self.derived

    def name(self, var: str) -> str:
        if _NEEDS_BACKTICKS_RE.search(var) or any(
            re.search(rf"[-/]{re.escape(k)}(?![^\s!-/:-@\[-^`{{-~])", var) for k in self.known
        ):
            return f"`{var}`"
        return var

    def value(self, token: str, var: str) -> str:
        """比较右侧 / 枚举项 / 等于的字面值, 按 expressions.literal 的判定顺序翻译。"""
        t = token.strip()
        if len(t) >= 2 and t[0] == t[-1] and t[0] in "\"'":
            return _quote(t[1:-1])
        if t == NULL_WORD:
            return NULL_WORD
        if self.is_var(t):
            return self.name(t)
        if t in TRUE_WORDS or t in FALSE_WORDS:
            if self.is_boolish(var):
                return "是" if t in TRUE_WORDS else "否"
            return _quote(t)
        try:
            f = float(t)
            return _number_text(int(f) if f.is_integer() and "." not in t else f)
        except ValueError:
            pass
        if _DATE_TEXT_RE.match(t):
            return f'日期("{t}")'
        return _quote(t)

    def _is_plain(self, part: str) -> bool:
        """枚举项是不是单纯的字面值 (不是比较 / 区间 / 文本关键字 / 判空 / not)。"""
        p = part.strip()
        if p == NULL_WORD:
            return True  # 属于(x, "a", 空) / x not in ("a", 空) 都认 空
        return not (
            p in _NULL_TESTS or p.startswith("not(") or _RANGE_RE.match(p) or _CMP_RE.match(p) or _TEXT_TEST_RE.match(p)
        )

    def cell(self, var: str, cell: Any) -> str | None:
        """None = 不关心。分支顺序与 conditions.unary_test 一致。"""
        v = self.name(var)
        if cell is None:
            return None
        if isinstance(cell, bool) or isinstance(cell, numbers.Real):
            return f"{v} == {_number_text(cell)}"
        if isinstance(cell, (datetime, date)):
            return f"{v} == {_date_text(cell)}"
        s = str(cell).strip()
        if s in BLANK_WORDS:
            return None
        if s.startswith("not(") and s.endswith(")"):
            inner = s[4:-1].strip()
            if inner in (NULL_WORD, "为空"):
                return f"{v} != {NULL_WORD}"
            if inner == "非空":
                return f"{v} == {NULL_WORD}"
            parts = split_top_level(inner)
            if parts and all(self._is_plain(p) for p in parts):  # != / not in 本身就是空值不成立, 与 not(...) 一致
                if len(parts) == 1:
                    return f"{v} != {self.value(parts[0], var)}"
                return f"{v} not in ({', '.join(self.value(p, var) for p in parts)})"
            return f"{v} != {NULL_WORD} 且 非 ({self.cell(var, inner)})"
        if m := _TEXT_TEST_RE.match(s):
            args = [self.value(p, var) for p in split_top_level(m.group(2) or m.group(3))]
            return f"{m.group(1)}({v}, {', '.join(args)})"
        parts = split_top_level(s)
        if len(parts) > 1:
            if any(p in BLANK_WORDS for p in parts):
                return None  # 枚举里有「不关心」, 整格就是不关心
            if all(self._is_plain(p) for p in parts):
                return f"属于({v}, {', '.join(self.value(p, var) for p in parts)})"
            return " 或 ".join(str(self.cell(var, p)) for p in parts)
        if s == "非空":
            return f"{v} != {NULL_WORD}"
        if s in ("为空", NULL_WORD):
            return f"{v} == {NULL_WORD}"
        if m := _RANGE_RE.match(s):
            lo, hi = self.value(m.group(2), var), self.value(m.group(3), var)
            return f"{lo} {'<=' if m.group(1) == '[' else '<'} {v} {'<=' if m.group(4) == ']' else '<'} {hi}"
        if m := _CMP_RE.match(s):
            return f"{v} {_CMP_OPS[m.group(1)]} {self.value(m.group(2), var)}"
        return f"{v} == {self.value(s, var)}"


# ---------------------------------------------------------------------------
# Excel -> YAML
# ---------------------------------------------------------------------------


class FlowDict(dict):
    """dump 成 {a: 1, b: 2} 的一行写法 (输入声明、用例的输入 / 期望)。"""


class _Dumper(yaml.SafeDumper):
    pass


_Dumper.add_representer(FlowDict, lambda d, data: d.represent_mapping("tag:yaml.org,2002:map", data, flow_style=True))


def _plain(v: Any) -> Any:
    """单元格值 -> YAML 值: 日期去掉 0 点的时间, 其余原样。"""
    if isinstance(v, datetime):
        return v.date() if (v.hour, v.minute, v.second) == (0, 0, 0) else v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def _given(d: dict[str, Any]) -> FlowDict:
    """用例的输入 / 期望: 留空的键省略 (YAML 里不写 = 空值 / 不变, 与 Excel 留空同义)。"""
    return FlowDict({k: _plain(v) for k, v in d.items() if not _blank_value(v)})


def _blank_value(v: Any) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def _match_text(spec: Any, primary: str) -> str:
    if spec.primary:
        return PRIMARY_WORD
    return ", ".join(m if m == t else f"{m} = {primary}.{t}" for m, t in spec.keys)


def _rule_doc(tr: CellTranslator, r: RawRule) -> dict[str, Any]:
    doc: dict[str, Any] = {"编号": r.rule_id}
    if r.desc:
        doc["名称"] = r.desc
    doc["优先级"] = int(r.priority) if float(r.priority).is_integer() else r.priority
    branches = []
    for b in r.branches:
        exprs = [e for e in (tr.cell(var, cell) for var, cell in b.items()) if e is not None]
        branches.append(exprs)
    if len(branches) == 1:
        if branches[0]:
            doc["当"] = branches[0]
    else:
        doc["当"] = [{"任一满足": [b[0] if len(b) == 1 else b for b in branches]}]
    then = {col: _plain(v) for col, v in r.outputs.items() if v is not None and str(v).strip() not in BLANK_WORDS}
    if then:
        doc["则"] = then
    return doc


def excel_to_yaml_doc(path: str | Path) -> tuple[dict[str, Any], RuleSet]:
    rs, raw = load_excel_raw(Path(path))
    tr = CellTranslator(rs)
    primary = rs.primary.name
    doc: dict[str, Any] = {"输入": {}}
    for s in rs.inputs:
        entry: dict[str, Any] = {}
        if s.path:
            entry["文件"] = s.path
        entry["匹配"] = _match_text(s, primary)
        if not s.primary and s.dup_policy != "取第一条":
            entry["重复键"] = s.dup_policy
        doc["输入"][s.name] = FlowDict(entry)
    if rs.settings:
        doc["设置"] = {k: _plain(v) for k, v in rs.settings.items()}
    if rs.derived:
        doc["变量"] = dict(rs.derived)
    doc["命中策略"] = rs.hit_policy
    doc["规则"] = [_rule_doc(tr, r) for r in raw]
    if rs.tests:
        doc["测试"] = [{"用例": t.name, "输入": _given(t.inputs), "期望": _given(t.expected)} for t in rs.tests]
    return doc, rs


def excel_to_yaml(path: str | Path) -> str:
    doc, rs = excel_to_yaml_doc(path)
    body = yaml.dump(doc, Dumper=_Dumper, allow_unicode=True, sort_keys=False, width=200)
    head = f"# 由 rule-sheet export 从 {Path(path).name} 生成, 请勿手改: 源是 Excel, 改了 Excel 后重新导出。\n\n"
    return head + body


def export_workbook(path: str | Path, out: str | Path | None = None) -> Path:
    """Excel 规则工作簿 -> YAML 文件; 默认写到同目录同名 .yaml。"""
    src = Path(path)
    if src.suffix.lower() != ".xlsx":
        raise ValueError(f"export 现在只支持 Excel 规则工作簿 (.xlsx) 导出成 YAML: {src.name}")
    target = Path(out) if out else src.with_suffix(".yaml")
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(excel_to_yaml(src), encoding="utf-8")
    return target


# ---------------------------------------------------------------------------
# 两份规则文件的语义比对
# ---------------------------------------------------------------------------


@dataclass
class DiffResult:
    differences: list[str] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)  # 路径之类不算差异的提示
    cases: int = 0  # 交叉验证用了多少条用例

    @property
    def same(self) -> bool:
        return not self.differences


def _norm_cell(v: Any) -> Any:
    if v is None:
        return None
    if isinstance(v, str):
        s = v.strip()
        return None if s in BLANK_WORDS else s
    if isinstance(v, datetime):
        return _plain(v)
    if isinstance(v, float) and v.is_integer():
        return int(v)
    return v


def _norm_given(d: dict[str, Any]) -> dict[str, Any]:
    out = {}
    for k, v in d.items():
        if v is None or (isinstance(v, str) and not v.strip()):
            continue
        out[k] = _plain(v) if isinstance(v, (datetime, float)) else v
    return out


def _case_key(name: str) -> str:
    """用例按首个词 (用例 ID) 配对: Excel 的「用例」列只写 T1, YAML 习惯写 "T1 偏差 8% 触发 R1"。"""
    return name.strip().split()[0] if name.strip() else name


def _structure_diff(a: RuleSet, b: RuleSet, res: DiffResult) -> None:
    ia = {s.name: s for s in a.inputs}
    ib = {s.name: s for s in b.inputs}
    for n in sorted(set(ia) | set(ib)):
        if n not in ia or n not in ib:
            res.differences.append(f"输入「{n}」只在 {'B' if n not in ia else 'A'} 里声明")
            continue
        sa, sb = ia[n], ib[n]
        if (sa.primary, sa.keys, sa.dup_policy) != (sb.primary, sb.keys, sb.dup_policy):
            ka, kb = sa.keys or PRIMARY_WORD, sb.keys or PRIMARY_WORD
            res.differences.append(f"输入「{n}」: A 匹配 {ka} / {sa.dup_policy}, B {kb} / {sb.dup_policy}")
        if sa.path != sb.path:
            res.notes.append(f"输入「{n}」的默认文件不同: A {sa.path}, B {sb.path}")
    for k in sorted(set(a.settings) | set(b.settings)):
        va, vb = _norm_cell(a.settings.get(k)), _norm_cell(b.settings.get(k))
        if va != vb:
            (res.notes if k == "输出" else res.differences).append(f"设置「{k}」: A {va}, B {vb}")
    if a.derived != b.derived:
        res.differences.append(f"变量定义不同: A {a.derived}, B {b.derived}")
    if a.hit_policy != b.hit_policy:
        res.differences.append(f"命中策略: A {a.hit_policy}, B {b.hit_policy}")
    if set(a.outputs) != set(b.outputs):
        res.differences.append(f"输出列: A {a.outputs}, B {b.outputs}")
    ra = {r.rule_id: r for r in a.rules}
    rb = {r.rule_id: r for r in b.rules}
    for rid in [r.rule_id for r in a.rules] + [r.rule_id for r in b.rules if r.rule_id not in ra]:
        if rid not in ra or rid not in rb:
            res.differences.append(f"规则 {rid} 只在 {'B' if rid not in ra else 'A'} 里")
            continue
        if ra[rid].priority != rb[rid].priority:
            res.differences.append(f"规则 {rid} 优先级: A {ra[rid].priority:g}, B {rb[rid].priority:g}")
        oa = {c: _norm_cell(v) for c, v in ra[rid].outputs.items() if _norm_cell(v) is not None}
        ob = {c: _norm_cell(v) for c, v in rb[rid].outputs.items() if _norm_cell(v) is not None}
        if oa != ob:
            res.differences.append(f"规则 {rid} 输出: A {oa}, B {ob}")
    ta = {_case_key(t.name): t for t in a.tests}
    tb = {_case_key(t.name): t for t in b.tests}
    keys = [_case_key(t.name) for t in a.tests] + [k for k in (_case_key(t.name) for t in b.tests) if k not in ta]
    for key in keys:
        if key not in ta or key not in tb:
            res.differences.append(f"用例「{key}」只在 {'B' if key not in ta else 'A'} 里")
            continue
        ga, gb = _norm_given(ta[key].inputs), _norm_given(tb[key].inputs)
        if ga != gb:
            res.differences.append(f"用例「{key}」输入不同: A {ga}, B {gb}")
        ea, eb = _norm_given(ta[key].expected), _norm_given(tb[key].expected)
        if ea != eb:
            res.differences.append(f"用例「{key}」期望不同: A {ea}, B {eb}")


def _run_case(rs: RuleSet, tc: TestCase) -> tuple[dict[str, Any], str] | str:
    ns, originals = case_namespace(rs, tc)
    try:
        results, hits = apply_rules(rs, ns, originals)
    except Exception as exc:  # noqa: BLE001
        return str(exc)
    return {col: results[col].iloc[0] for col in rs.outputs}, hits.iloc[0]


def _behavior_diff(a: RuleSet, b: RuleSet, res: DiffResult) -> None:
    """两边的用例合并, 每条在两套规则上各跑一次, 比实际输出与命中 ID (不是比通过与否)。"""
    seen: set[tuple[str, str]] = set()
    cols = [c for c in a.outputs if c in b.outputs]
    for tc in a.tests + b.tests:
        key = (_case_key(tc.name), repr(sorted(_norm_given(tc.inputs).items())))
        if key in seen:
            continue  # 两边同名同输入的用例只跑一次
        seen.add(key)
        res.cases += 1
        oa, ob = _run_case(a, tc), _run_case(b, tc)
        name = _case_key(tc.name)
        if isinstance(oa, str) or isinstance(ob, str):
            ea = oa if isinstance(oa, str) else "正常"
            eb = ob if isinstance(ob, str) else "正常"
            res.differences.append(f"用例「{name}」出错: A {ea}; B {eb}")
            continue
        (va, ha), (vb, hb) = oa, ob
        for col in cols:
            if not same(va[col], vb[col]):
                res.differences.append(f"用例「{name}」输出「{col}」: A {fmt(va[col])}, B {fmt(vb[col])}")
        if ha != hb:
            res.differences.append(f"用例「{name}」命中规则: A {ha or '无'}, B {hb or '无'}")


def diff_rules(path_a: str | Path, path_b: str | Path) -> DiffResult:
    a, b = load_rules(path_a), load_rules(path_b)
    res = DiffResult()
    _structure_diff(a, b, res)
    _behavior_diff(a, b, res)
    return res
