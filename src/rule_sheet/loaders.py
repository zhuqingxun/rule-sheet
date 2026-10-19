"""规则文件加载: Excel 决策表 与 YAML 两种前端, 都产出 RuleSet。"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml
from openpyxl import load_workbook

from .conditions import build_condition, cell_condition, strings_in
from .expressions import FUNCS, find_names
from .model import BASE_DATE_VAR, HIT_POLICIES, Rule, RuleSet, TestCase


def _check_policy(policy: Any) -> str:
    name = str(policy or "优先级").strip()
    if name not in HIT_POLICIES:
        raise ValueError(f"未知的命中策略「{name}」, 只认: {' / '.join(HIT_POLICIES)}")
    return name


# ---------------------------------------------------------------------------
# Excel 决策表
# ---------------------------------------------------------------------------


def _rows(ws: Any) -> list[list[Any]]:
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _read_kv(ws: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for row in _rows(ws)[1:]:
        if row and row[0] not in (None, "") and len(row) > 1 and row[1] not in (None, ""):
            out[str(row[0]).strip()] = row[1]
    return out


def _parse_header(grid: list[list[Any]]) -> tuple[list[tuple[int, str, str]], int]:
    """表头固定在第 2、3 行: 第 2 行列类型 (输入/输出/...), 第 3 行变量名。返回 (列索引, 类型, 名字) 与正文起始行。"""
    kinds, names = grid[1], grid[2]
    cols = [
        (i, str(k).strip(), str(n).strip())
        for i, (k, n) in enumerate(zip(kinds, names, strict=False))
        if k not in (None, "") and n not in (None, "")
    ]
    return cols, 3


def load_excel_rules(path: Path) -> RuleSet:
    # data_only=False: 以 = 开头的格 (= 活跃 / =主表.金额 * 1.1) 在 Excel 里会被当成公式存起来,
    # data_only=True 读到的是公式的缓存结果 (没算过是 None, 算过是 #NAME?), 条件就静默变成「不关心」。
    # 这里按原文读, 规则工作簿里的 = 一律是 rule-sheet 表达式, 不支持写 Excel 公式。
    wb = load_workbook(path, data_only=False)
    for sheet in ("匹配设置", "决策表"):
        if sheet not in wb.sheetnames:
            raise ValueError(f"规则文件缺少「{sheet}」sheet: {path.name}")
    match = _read_kv(wb["匹配设置"])
    derived: list[tuple[str, str]] = []
    if "变量定义" in wb.sheetnames:
        derived = [
            (str(r[0]).strip(), str(r[1]).strip())
            for r in _rows(wb["变量定义"])[1:]
            if r and r[0] not in (None, "") and len(r) > 1 and r[1] not in (None, "")
        ]

    grid = _rows(wb["决策表"])
    hit_policy = _check_policy(grid[0][1])
    cols, body_start = _parse_header(grid)
    id_col = next(i for i, k, _ in cols if k == "规则ID")
    pri_col = next((i for i, k, _ in cols if k == "优先级"), None)
    desc_col = next((i for i, k, _ in cols if k == "说明"), None)
    inputs = [(i, n) for i, k, n in cols if k == "输入"]
    outputs = [(i, n) for i, k, n in cols if k == "输出"]

    rules: list[Rule] = []
    for order, row in enumerate(grid[body_start:]):
        if len(row) <= id_col or row[id_col] in (None, ""):
            continue
        pri = float(row[pri_col]) if pri_col is not None and row[pri_col] not in (None, "") else 1000.0 + order
        rules.append(
            Rule(
                rule_id=str(row[id_col]).strip(),
                priority=pri,
                conditions=[cell_condition(name, row[i]) for i, name in inputs],
                outputs={name: row[i] for i, name in outputs},
                desc=str(row[desc_col] or "") if desc_col is not None else "",
            )
        )

    tests: list[TestCase] = []
    if "测试" in wb.sheetnames:
        tgrid = _rows(wb["测试"])
        tcols, tstart = _parse_header(tgrid)
        case_col = next(i for i, k, _ in tcols if k == "用例")
        t_inputs = [(i, n) for i, k, n in tcols if k == "输入"]
        t_expect = [(i, n) for i, k, n in tcols if k == "期望"]
        for row in tgrid[tstart:]:
            if len(row) <= case_col or row[case_col] in (None, ""):
                continue
            tests.append(
                TestCase(
                    name=str(row[case_col]),
                    inputs={n: row[i] for i, n in t_inputs},
                    expected={n: row[i] for i, n in t_expect},
                )
            )

    return RuleSet(
        match,
        derived,
        hit_policy,
        rules,
        [n for _, n in outputs],
        tests,
        source=path.name,
        variables=[n for _, n in inputs],
    )


# ---------------------------------------------------------------------------
# YAML
# ---------------------------------------------------------------------------


def load_yaml_rules(path: Path) -> RuleSet:
    doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    rules: list[Rule] = []
    outputs: list[str] = []
    for order, r in enumerate(doc.get("规则") or []):
        conds = r.get("当") or []
        then = r.get("则") or {}  # 则: 省略或为空 = 命中但什么都不改 (首次命中 下用来拦截一行)
        rules.append(
            Rule(
                rule_id=str(r.get("编号", f"R{order + 1}")),
                priority=float(r.get("优先级", 1000 + order)),
                conditions=[build_condition(conds)] if conds else [],
                outputs=dict(then),
                desc=str(r.get("名称", "")),
            )
        )
        for col in then:
            if col not in outputs:
                outputs.append(col)
    tests = [
        TestCase(str(t.get("用例", f"T{i + 1}")), dict(t.get("输入") or {}), dict(t.get("期望") or {}))
        for i, t in enumerate(doc.get("测试") or [])
    ]
    derived = [(str(k), str(v)) for k, v in (doc.get("变量") or {}).items()]
    match = dict(doc.get("匹配") or {})
    la, ra = str(match.get("左表别名", "主表")), str(match.get("右表别名", "本月"))
    # 扫描 当 / 则 / 变量 里引用到的 <别名>.<列> 变量名; YAML 里新增写变量名的位置要同步加进 texts
    texts = [s for r in doc.get("规则") or [] for s in strings_in(r.get("当"))]
    texts += [str(v) for r in doc.get("规则") or [] for v in (r.get("则") or {}).values()]
    texts += [v for _, v in derived]
    known = [k for k, _ in derived] + list(FUNCS) + [BASE_DATE_VAR]
    variables = sorted({n for t in texts for n in find_names(t, (la, ra), known)} | {k for k, _ in derived})
    return RuleSet(
        match,
        derived,
        _check_policy(doc.get("命中策略")),
        rules,
        outputs,
        tests,
        source=path.name,
        variables=variables,
    )


def load_rules(path: str | Path) -> RuleSet:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"规则文件不存在: {p}")
    if p.suffix.lower() in (".yaml", ".yml"):
        return load_yaml_rules(p)
    return load_excel_rules(p)
