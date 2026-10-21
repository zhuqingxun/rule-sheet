"""执行核心: 读入各输入表 -> 以主表为锚逐个合并参照表 -> 派生变量 -> 逐条规则求 mask -> 按命中策略合并 -> 审计列。"""

from __future__ import annotations

import numbers
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .conditions import ns_index, resolve_output, to_mask
from .expressions import eval_expr, is_na, text_of
from .loaders import load_rules
from .model import (
    BASE_DATE_VAR,
    CLEAR,
    DEFAULT_JOIN_SEP,
    EXIST_VAR,
    EXPECT_NULL_WORDS,
    JOIN_POLICY,
    JOIN_SEP_KEY,
    NULL,
    NULL_WORD,
    TRUE_WORDS,
    Namespace,
    Rule,
    RuleSet,
)
from .style import display_width

Mask = pd.Series

# ---------------------------------------------------------------------------
# 读表
# ---------------------------------------------------------------------------


def resolve_path(p: str | Path, base_dir: Path | None) -> Path:
    path = Path(str(p))
    return path if path.is_absolute() or base_dir is None else base_dir / path


def split_sheet(spec: str | Path, base_dir: Path | None = None) -> tuple[Path, str | None]:
    """`路径#sheet名` 拆成 (路径, sheet名)。文件名本身也可以含 #: 整个字符串就是一个存在的文件时不拆。"""
    text = str(spec)
    whole = resolve_path(text, base_dir)
    if whole.exists() or "#" not in text:
        return whole, None
    base, sheet = text.rsplit("#", 1)
    return resolve_path(base, base_dir), sheet or None


def read_table(path: str | Path, sheet: str | None = None) -> pd.DataFrame:
    """读 Excel, 保住文本格式的数字串。

    pandas 3 的 read_excel 默认会把文本单元格里的 "007" / "13800000001" 转成数字 (前导零丢失, 电话变浮点)。
    先按 object 原样读, 再让 pandas 只对整列同类型的数据做推断: 数值列、日期列正常, 文本列保持文本。
    """
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"数据表不存在: {p}")
    if sheet is not None:
        wb = load_workbook(p, read_only=True)
        names = wb.sheetnames
        wb.close()
        if sheet not in names:
            raise ValueError(f"数据表 {p.name} 里没有 sheet「{sheet}」, 有: {', '.join(names)}")
    return pd.read_excel(p, sheet_name=sheet if sheet is not None else 0, dtype=object).infer_objects()


# ---------------------------------------------------------------------------
# 命名空间: 以主表为锚, 把每个参照表按各自的匹配键合并进来, 变量名 <输入名>.<列>
# ---------------------------------------------------------------------------


def build_namespace(rs: RuleSet, tables: dict[str, pd.DataFrame], notes: list[str] | None = None) -> Namespace:
    primary = rs.primary
    base = tables[primary.name]
    ns: Namespace = {f"{primary.name}.{col}": base[col] for col in base.columns}
    for spec in rs.lookups:
        df = tables[spec.name]
        if not spec.keys:
            raise ValueError(f"输入「{spec.name}」没有匹配键, 不知道怎么对齐到主表「{primary.name}」")
        mine = [m for m, _ in spec.keys]
        theirs = [t for _, t in spec.keys]
        for c in theirs:
            if c not in base.columns:
                raise ValueError(f"主表「{primary.name}」没有列「{c}」, 而输入「{spec.name}」要按它对齐")
        for c in mine:
            if c not in df.columns:
                raise ValueError(f"输入「{spec.name}」没有匹配键列「{c}」")
        df = df[df[mine].notna().all(axis=1)]  # 键为空的行不参与匹配: pandas merge 会把两边的 NaN 配成一对
        dup = df.duplicated(subset=mine, keep="first")
        if dup.any():
            if spec.dup_policy != "取第一条":
                raise ValueError(
                    f"输入「{spec.name}」有 {int(dup.sum())} 行重复键, 而它的「重复键」设置为 {spec.dup_policy}"
                )
            if notes is not None:
                notes.append(f"输入「{spec.name}」有 {int(dup.sum())} 行重复键, 按「取第一条」处理")
            df = df[~dup]
        # 左侧只带主表的键列, 参照表之间同名列互不干扰; 参照表里与主表键列同名的非键列保留参照表的 (主表的加后缀)
        merged = base[theirs].merge(df, left_on=theirs, right_on=mine, how="left", suffixes=("__p", ""), indicator=True)
        merged.index = base.index
        for col in df.columns:
            if col not in mine:
                ns[f"{spec.name}.{col}"] = merged[col]
        ns[f"{spec.name}.{EXIST_VAR}"] = merged["_merge"].eq("both")
    base_date = rs.settings.get(BASE_DATE_VAR)
    ns[BASE_DATE_VAR] = pd.Timestamp(base_date) if base_date else pd.Timestamp(date.today())
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
        if val is NULL:
            val = None
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


def _join_hits(
    rs: RuleSet,
    col: str,
    ordered: list[Rule],
    masks: dict[str, Mask],
    cur: pd.Series,
    output_of: Any,
) -> pd.Series:
    """多条拼接: 按优先级顺序把每条命中规则对该列写的值接到后面; 清空 / 算出空值的规则不贡献。"""
    sep = str(rs.settings.get(JOIN_SEP_KEY, DEFAULT_JOIN_SEP))
    acc = pd.Series([None] * len(cur), index=cur.index, dtype=object)
    for rule in ordered:
        val = output_of(rule, col)
        if val is None or val is CLEAR:
            continue
        if not isinstance(val, pd.Series):
            val = pd.Series([val] * len(cur), index=cur.index, dtype=object)
        texts = pd.Series([None if is_na(v) else text_of(v) for v in val.tolist()], index=cur.index, dtype=object)
        add = masks[rule.rule_id] & texts.notna()
        prefix = acc.map(lambda s: "" if s is None else s + sep)
        acc = acc.mask(add, prefix + texts.fillna(""))
    return cur.mask(acc.notna(), acc)


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
    命中策略 多条拼接: 每个输出列把所有命中规则写的值按优先级顺序 (同优先级按书写顺序) 用分隔符拼成一个文本,
        替换原值; 没写该列的规则不参与, 没有任何命中的行不动。
    """
    index = ns_index(ns)
    dup = sorted({r.rule_id for r in rs.rules if sum(x.rule_id == r.rule_id for x in rs.rules) > 1})
    if dup:  # 两个加载器都已拦截, 这里兜底: masks 按 ID 建字典, 重复 ID 会静默丢掉一条规则
        raise ValueError(f"规则 ID 重复: {', '.join(dup)}")
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

    def output_of(rule: Rule, col: str) -> Any:
        try:
            return resolve_output(rule.outputs.get(col), ns)
        except KeyError as exc:
            raise KeyError(f"规则 {rule.rule_id} 输出「{col}」: {exc.args[0]}") from exc

    results: dict[str, pd.Series] = {}
    for col in rs.outputs:
        cur = originals.get(col)
        if cur is None:
            cur = pd.Series([None] * len(index), index=index, dtype=object)
        cur = cur.astype(object)
        if rs.hit_policy == JOIN_POLICY:
            results[col] = _join_hits(rs, col, ordered, masks, cur, output_of)
            continue
        for rule in reversed(ordered):  # 优先级数字大的先写, 小的最后写并覆盖
            val = output_of(rule, col)
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
    return "空" if is_na(x) else text_of(x)


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
    exist_vars = {f"{s.name}.{EXIST_VAR}" for s in rs.lookups}
    derived_names = {d for d, _ in rs.derived}
    outcomes: list[TestOutcome] = []
    for tc in rs.tests:
        index = pd.RangeIndex(1)
        ns: Namespace = {}
        for name, val in tc.inputs.items():
            if name in exist_vars and isinstance(val, str):
                val = val.strip() in TRUE_WORDS
            if isinstance(val, str) and val.strip() in ("", NULL_WORD):  # 留空或写 空 = 空值
                val = None
            if val is None and name in derived_names:
                continue  # 派生变量留空 = 没给, 按表达式推导 (Excel 测试 sheet 每列每行都在, 留空不能算「给了空值」)
            if isinstance(val, (date, datetime)):  # Excel 日期格 / YAML 日期 -> 与真实数据的日期列同型, 才能做大小比较
                val = pd.Timestamp(val)
            dtype = object if isinstance(val, str) or val is None else None
            ns[name] = pd.Series([val], index=index, dtype=dtype)
        ns.setdefault(BASE_DATE_VAR, pd.Timestamp(rs.settings.get(BASE_DATE_VAR) or date.today()))
        for name in rs.variables:  # 用例没提到的变量 = 空值
            if name not in ns and name not in derived_names:
                ns[name] = pd.Series([None], index=index, dtype=object)
        add_derived(rs, ns, index, lenient=True)
        originals = {
            col: ns.get(f"{rs.primary.name}.{col}", pd.Series([None], index=index, dtype=object)) for col in rs.outputs
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
                want = None  # 期望写 空 / 清空 / 为空 = 该列应为空值
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
                    width = max([display_width(col)] + [display_width(v) for v in df[col].head(50)]) + 2
                    ws.column_dimensions[get_column_letter(i)].width = min(max(width, 8), 60)
                    if pd.api.types.is_datetime64_any_dtype(df[col]):
                        for r in range(2, len(df) + 2):
                            ws.cell(row=r, column=i).number_format = "yyyy-mm-dd"
                ws.freeze_panes = "B2"
            return p
        except PermissionError:
            continue
    raise PermissionError(f"无法写入 {path} 及其备选文件名, 请关闭占用的 Excel 窗口")


def bind_inputs(rs: RuleSet, base_dir: Path, overrides: dict[str, str] | None) -> dict[str, tuple[Path, str | None]]:
    """每个声明的输入对应一个文件: 命令行给的 (相对当前目录) 优先, 否则规则文件里的默认文件 (相对规则文件目录)。"""
    overrides = dict(overrides or {})
    unknown = [n for n in overrides if n not in rs.names]
    if unknown:
        raise ValueError(f"命令行给的输入「{unknown[0]}」不是规则声明的输入 (规则声明了: {', '.join(rs.names)})")
    bound: dict[str, tuple[Path, str | None]] = {}
    missing: list[str] = []
    for spec in rs.inputs:
        if spec.name in overrides:
            bound[spec.name] = split_sheet(overrides[spec.name])
        elif spec.path:
            bound[spec.name] = split_sheet(spec.path, base_dir)
        else:
            missing.append(spec.name)
    if missing:
        raise ValueError(
            f"输入「{'」「'.join(missing)}」没有绑定文件: 规则文件里没写默认文件, 命令行也没给 (-i 名称=路径)"
        )
    return bound


@dataclass
class RunResult:
    ruleset: RuleSet
    tests: list[TestOutcome]
    keys: list[str] = field(default_factory=list)  # 主表上用来标识一行的列 (第一个参照表的匹配键, 没有就是第一列)
    out: pd.DataFrame | None = None
    changed: pd.DataFrame | None = None
    written: Path | None = None
    stopped_by_tests: bool = False
    notes: list[str] = field(default_factory=list)
    inputs: dict[str, tuple[Path, int]] = field(default_factory=dict)  # 输入名 -> (实际读的文件, 行数)

    @property
    def tests_passed(self) -> int:
        return sum(1 for t in self.tests if t.passed)


def run(
    rules_path: str | Path,
    *,
    base_dir: str | Path | None = None,
    inputs: dict[str, str] | None = None,
    output: str | Path | None = None,
    ignore_tests: bool = False,
    dry_run: bool = False,
) -> RunResult:
    """完整流程: 加载规则 -> 跑内嵌测试 -> 绑定并读入各输入表 -> 执行 -> 写出。

    inputs: {输入名: 文件路径 (可带 #sheet名)}, 覆盖规则文件里的默认文件, 相对当前目录。
    规则文件里的相对路径 (默认文件 / 输出) 相对 base_dir, 默认为规则文件所在目录。
    """
    rules_path = Path(rules_path)
    base = Path(base_dir) if base_dir else rules_path.parent
    rs = load_rules(rules_path)
    outcomes = run_tests(rs) if rs.tests else []
    if any(not o.passed for o in outcomes) and not ignore_tests:
        return RunResult(rs, outcomes, stopped_by_tests=True)

    bound = bind_inputs(rs, base, inputs)
    tables = {name: read_table(p, sheet) for name, (p, sheet) in bound.items()}
    primary_df = tables[rs.primary.name]

    notes: list[str] = []
    ns = build_namespace(rs, tables, notes)
    add_derived(rs, ns, primary_df.index, lenient=False)
    originals = {col: primary_df[col] for col in rs.outputs if col in primary_df.columns}
    first = rs.lookups[0].keys if rs.lookups else []
    keys = [t for _, t in first] or [str(primary_df.columns[0])]
    results, hit_ids = apply_rules(rs, ns, originals, labels=primary_df[keys[0]])

    out = primary_df.copy()
    details: list[str] = []
    for i in primary_df.index:
        parts = []
        for col in rs.outputs:
            old = primary_df[col].loc[i] if col in primary_df.columns else None
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
        if output:
            out_path = Path(output)
        else:
            default = rs.settings.get("输出") or f"output/{bound[rs.primary.name][0].stem}_更新后.xlsx"
            out_path = resolve_path(default, base)
        written = write_xlsx_safely(out, out_path)
    read_info = {name: (p, len(tables[name])) for name, (p, _) in bound.items()}
    return RunResult(rs, outcomes, keys, out, changed, written, notes=notes, inputs=read_info)
