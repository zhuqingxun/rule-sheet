"""执行核心: 合并两表 -> 派生变量 -> 逐条规则求 mask -> 按命中策略合并输出 -> 审计列。"""

from __future__ import annotations

import numbers
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.utils import get_column_letter

from .conditions import ns_index, resolve_output, to_mask
from .expressions import eval_expr
from .loaders import load_rules
from .model import BASE_DATE_VAR, CLEAR, EXIST_VAR, EXPECT_NULL_WORDS, TRUE_WORDS, Namespace, Rule, RuleSet

Mask = pd.Series

# ---------------------------------------------------------------------------
# 读表
# ---------------------------------------------------------------------------


def read_table(path: str | Path) -> pd.DataFrame:
    """读 Excel, 保住文本格式的数字串。

    pandas 3 的 read_excel 默认会把文本单元格里的 "007" / "13800000001" 转成数字 (前导零丢失, 电话变浮点)。
    先按 object 原样读, 再让 pandas 只对整列同类型的数据做推断: 数值列、日期列正常, 文本列保持文本。
    """
    return pd.read_excel(path, dtype=object).infer_objects()


# ---------------------------------------------------------------------------
# 命名空间: 把两张表合并成 <左别名>.* / <右别名>.* 变量
# ---------------------------------------------------------------------------


def build_namespace(
    rs: RuleSet,
    left: pd.DataFrame,
    right: pd.DataFrame,
    keys: list[str],
    notes: list[str] | None = None,
) -> Namespace:
    for k in keys:
        if k not in left.columns or k not in right.columns:
            raise ValueError(f"匹配键「{k}」必须同时存在于左右两表")
    dup = right.duplicated(subset=keys, keep="first")
    if dup.any():
        policy = str(rs.match.get("右表重复键", "取第一条"))
        if policy != "取第一条":
            raise ValueError(f"右表有 {int(dup.sum())} 行重复键, 而「右表重复键」设置为 {policy}")
        if notes is not None:
            notes.append(f"右表有 {int(dup.sum())} 行重复键, 按「取第一条」处理")
        right = right[~dup]

    merged = left.merge(right, on=keys, how="left", suffixes=("", "__r"), indicator=True)
    merged.index = left.index
    la, ra = rs.left_alias, rs.right_alias
    ns: Namespace = {}
    for col in left.columns:
        ns[f"{la}.{col}"] = merged[col]
    for col in right.columns:
        if col in keys:
            continue
        src = f"{col}__r" if f"{col}__r" in merged.columns else col
        ns[f"{ra}.{col}"] = merged[src]
    ns[f"{ra}.{EXIST_VAR}"] = merged["_merge"].eq("both")
    base = rs.match.get(BASE_DATE_VAR)
    ns[BASE_DATE_VAR] = pd.Timestamp(base) if base else pd.Timestamp(date.today())
    return ns


def add_derived(rs: RuleSet, ns: Namespace, index: pd.Index, *, lenient: bool) -> None:
    """按定义顺序求派生变量。测试用例里直接给了值的不再推导; lenient 下求值失败置空。"""
    for name, expr in rs.derived:
        if name in ns:
            continue
        try:
            val = eval_expr(expr, ns)
        except ValueError:
            if not lenient:
                raise
            val = float("nan")
        if not isinstance(val, pd.Series):
            val = pd.Series([val] * len(index), index=index)
        ns[name] = val


# ---------------------------------------------------------------------------
# 规则执行
# ---------------------------------------------------------------------------


def rule_mask(rule: Rule, ns: Namespace, index: pd.Index) -> Mask:
    m = pd.Series(True, index=index)
    for cond in rule.conditions:
        try:
            res = cond(ns)
        except KeyError as exc:
            raise KeyError(f"规则 {rule.rule_id}: {exc.args[0]}") from exc
        m = m & to_mask(res, ns)
    return m


def apply_rules(
    rs: RuleSet,
    ns: Namespace,
    originals: dict[str, pd.Series],
    labels: pd.Series | None = None,
) -> tuple[dict[str, pd.Series], pd.Series]:
    """返回 (每个输出列的新值, 每行命中的规则 ID 串)。

    命中策略 优先级: 每个输出列取优先级数字最小的命中规则的值, 未写该列的规则不参与。
    命中策略 唯一: 任何行命中多条即报错。
    命中策略 首次命中: 每行只算优先级数字最小的那条命中 (同优先级按书写顺序), 其余规则对该行视为未命中;
        之后走与 优先级 相同的逐列合并, 此时每行至多一条规则, 结果就是该规则的输出 (- 的列保持原值)。
    """
    index = ns_index(ns)
    masks = {r.rule_id: rule_mask(r, ns, index) for r in rs.rules}
    ordered = sorted(rs.rules, key=lambda r: r.priority)  # 稳定排序: 同优先级保持书写顺序

    if rs.hit_policy == "唯一":
        hits = sum((masks[r.rule_id].astype(int) for r in rs.rules), pd.Series(0, index=index))
        bad = hits[hits > 1]
        if not bad.empty:
            detail = {
                str(labels.loc[i]) if labels is not None else f"第{int(i) + 2}行": [
                    r.rule_id for r in ordered if masks[r.rule_id].loc[i]
                ]
                for i in bad.index
            }
            raise ValueError(f"命中策略为「唯一」, 但以下行命中多条规则: {detail}")
    elif rs.hit_policy == "首次命中":
        taken = pd.Series(False, index=index)
        for rule in ordered:
            effective = masks[rule.rule_id] & ~taken
            taken = taken | effective
            masks[rule.rule_id] = effective

    results: dict[str, pd.Series] = {}
    for col in rs.outputs:
        cur = originals.get(col)
        if cur is None:
            cur = pd.Series([None] * len(index), index=index, dtype=object)
        cur = cur.astype(object)
        for rule in reversed(ordered):  # 优先级数字大的先写, 小的最后写并覆盖
            val = resolve_output(rule.outputs.get(col), ns)
            if val is None:
                continue
            if val is CLEAR:
                val = None
            elif isinstance(val, pd.Series):
                val = val.astype(object)
            cur = cur.mask(masks[rule.rule_id], val)
        results[col] = cur

    hit_ids = pd.Series(
        ["; ".join(r.rule_id for r in ordered if masks[r.rule_id].loc[i]) for i in index],
        index=index,
    )
    return results, hit_ids


# ---------------------------------------------------------------------------
# 比较与格式化
# ---------------------------------------------------------------------------


def is_na(x: Any) -> bool:
    if x is None:
        return True
    try:
        return bool(pd.isna(x))
    except (TypeError, ValueError):
        return False


def same(a: Any, b: Any) -> bool:
    if is_na(a) or is_na(b):
        return is_na(a) and is_na(b)
    if isinstance(a, bool) or isinstance(b, bool):
        return bool(a) == bool(b)
    if isinstance(a, numbers.Real) and isinstance(b, numbers.Real):
        return abs(float(a) - float(b)) < 1e-9
    if isinstance(a, (pd.Timestamp, datetime, date)) or isinstance(b, (pd.Timestamp, datetime, date)):
        try:
            return pd.Timestamp(a) == pd.Timestamp(b)
        except (TypeError, ValueError):
            return False
    return str(a) == str(b)


def fmt(x: Any) -> str:
    if is_na(x):
        return "空"
    if isinstance(x, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(x).strftime("%Y-%m-%d")
    if isinstance(x, bool):
        return "是" if x else "否"
    if isinstance(x, numbers.Real) and float(x).is_integer():
        return str(int(x))
    return str(x)


# ---------------------------------------------------------------------------
# 规则级测试用例
# ---------------------------------------------------------------------------


@dataclass
class TestOutcome:
    name: str
    passed: bool
    hit: str = ""
    problems: list[str] = field(default_factory=list)
    error: str = ""


def run_tests(rs: RuleSet) -> list[TestOutcome]:
    exist_var = f"{rs.right_alias}.{EXIST_VAR}"
    outcomes: list[TestOutcome] = []
    for tc in rs.tests:
        index = pd.RangeIndex(1)
        ns: Namespace = {}
        for name, val in tc.inputs.items():
            if name == exist_var and isinstance(val, str):
                val = val.strip() in TRUE_WORDS
            if isinstance(val, str) and val.strip() == "":
                val = None
            dtype = object if isinstance(val, str) or val is None else None
            ns[name] = pd.Series([val], index=index, dtype=dtype)
        ns.setdefault(BASE_DATE_VAR, pd.Timestamp(rs.match.get(BASE_DATE_VAR) or date.today()))
        derived_names = {d for d, _ in rs.derived}
        for name in rs.variables:  # 用例没提到的变量 = 空值
            if name not in ns and name not in derived_names:
                ns[name] = pd.Series([None], index=index, dtype=object)
        add_derived(rs, ns, index, lenient=True)
        originals = {
            col: ns.get(f"{rs.left_alias}.{col}", pd.Series([None], index=index, dtype=object)) for col in rs.outputs
        }
        try:
            results, hit_ids = apply_rules(rs, ns, originals)
        except Exception as exc:  # noqa: BLE001
            outcomes.append(TestOutcome(tc.name, False, error=str(exc)))
            continue
        problems = []
        for col in rs.outputs:
            exp = tc.expected.get(col)
            got = results[col].iloc[0]
            if exp is None or (isinstance(exp, str) and exp.strip() == ""):
                want = originals[col].iloc[0]  # 期望留空 = 该列不变
            elif isinstance(exp, str) and exp.strip() in EXPECT_NULL_WORDS:
                want = None  # 期望写 清空 / 为空 = 该列应为空值
            else:
                want = exp
            if not same(got, want):
                problems.append(f"{col}: 期望 {fmt(want)}, 实际 {fmt(got)}")
        outcomes.append(TestOutcome(tc.name, not problems, hit=hit_ids.iloc[0], problems=problems))
    return outcomes


# ---------------------------------------------------------------------------
# 输出与整体流程
# ---------------------------------------------------------------------------


def write_xlsx_safely(df: pd.DataFrame, path: Path) -> Path:
    """目标被 Excel 占用时自动换 _v2/_v3 后缀, 不中断。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    candidates = [path] + [path.with_name(f"{path.stem}_v{i}{path.suffix}") for i in range(2, 10)]
    for p in candidates:
        try:
            with pd.ExcelWriter(p, engine="openpyxl") as xw:
                df.to_excel(xw, index=False)
                ws = xw.sheets[next(iter(xw.sheets))]
                for i, col in enumerate(df.columns, start=1):
                    width = max([len(str(col))] + [len(str(v)) for v in df[col].head(50)]) * 1.6 + 2
                    ws.column_dimensions[get_column_letter(i)].width = min(max(width, 8), 60)
                    if pd.api.types.is_datetime64_any_dtype(df[col]):
                        for r in range(2, len(df) + 2):
                            ws.cell(row=r, column=i).number_format = "yyyy-mm-dd"
                ws.freeze_panes = "B2"
            return p
        except PermissionError:
            continue
    raise PermissionError(f"无法写入 {path} 及其备选文件名, 请关闭占用的 Excel 窗口")


def resolve_path(p: str | Path, base_dir: Path) -> Path:
    path = Path(str(p))
    return path if path.is_absolute() else base_dir / path


@dataclass
class RunResult:
    ruleset: RuleSet
    tests: list[TestOutcome]
    keys: list[str] = field(default_factory=list)
    out: pd.DataFrame | None = None
    changed: pd.DataFrame | None = None
    written: Path | None = None
    stopped_by_tests: bool = False
    notes: list[str] = field(default_factory=list)

    @property
    def tests_passed(self) -> int:
        return sum(1 for t in self.tests if t.passed)


def run(
    rules_path: str | Path,
    *,
    base_dir: str | Path | None = None,
    ignore_tests: bool = False,
    dry_run: bool = False,
    output: str | Path | None = None,
) -> RunResult:
    """完整流程: 加载规则 -> 跑内嵌测试 -> 读两表 -> 执行 -> 写出。

    规则文件里的相对路径 (左表 / 右表 / 输出) 相对 base_dir, 默认为规则文件所在目录。
    """
    rules_path = Path(rules_path)
    base = Path(base_dir) if base_dir else rules_path.parent
    rs = load_rules(rules_path)
    outcomes = run_tests(rs) if rs.tests else []
    if any(not o.passed for o in outcomes) and not ignore_tests:
        return RunResult(rs, outcomes, stopped_by_tests=True)

    for key in ("左表", "右表", "匹配键"):
        if key not in rs.match:
            raise ValueError(f"匹配设置缺少「{key}」")
    left = read_table(resolve_path(rs.match["左表"], base))
    right = read_table(resolve_path(rs.match["右表"], base))
    keys = [k.strip() for k in str(rs.match["匹配键"]).split(",")]

    notes: list[str] = []
    ns = build_namespace(rs, left, right, keys, notes)
    add_derived(rs, ns, left.index, lenient=False)
    originals = {col: left[col] for col in rs.outputs if col in left.columns}
    results, hit_ids = apply_rules(rs, ns, originals, labels=left[keys[0]])

    out = left.copy()
    details: list[str] = []
    for i in left.index:
        parts = []
        for col in rs.outputs:
            old = left[col].loc[i] if col in left.columns else None
            new = results[col].loc[i]
            if not same(old, new):
                parts.append(f"{col}: {fmt(old)} -> {fmt(new)}")
        details.append("; ".join(parts))
    for col in rs.outputs:
        out[col] = results[col]
    out["命中规则"] = hit_ids
    out["变更明细"] = details
    changed = out[out["变更明细"] != ""]

    written = None
    if not dry_run:
        out_path = Path(output) if output else resolve_path(rs.match.get("输出", "output/结果.xlsx"), base)
        written = write_xlsx_safely(out, out_path)
    return RunResult(rs, outcomes, keys, out, changed, written, notes=notes)
