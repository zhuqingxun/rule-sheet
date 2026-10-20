"""规则文件加载: Excel 决策表 与 YAML 两种前端, 都产出 RuleSet。

加载阶段尽量把「写错了但不报错、测试假通过」的情况拦下来: 表头位置、列类型拼写、列名缺失、
测试用例的输入列 / 期望列对不上输出列或变量、重复的规则 ID、空的测试 sheet, 都在这里报错。
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .conditions import all_of, any_of, build_condition, cell_condition, strings_in
from .expressions import FUNCS, find_names
from .model import BASE_DATE_VAR, BLANK_WORDS, EXIST_VAR, HIT_POLICIES, Condition, Rule, RuleSet, TestCase

RULE_KINDS = ("规则ID", "优先级", "输入", "输出", "说明")
TEST_KINDS = ("用例", "输入", "期望", "说明")
POLICY_LABEL = "命中策略"


def _check_policy(policy: Any) -> str:
    name = str(policy or "优先级").strip()
    if name not in HIT_POLICIES:
        raise ValueError(f"未知的命中策略「{name}」, 只认: {' / '.join(HIT_POLICIES)}")
    return name


def _text(v: Any) -> str:
    return "" if v is None else str(v).strip()


def _blank(v: Any) -> bool:
    return _text(v) in BLANK_WORDS


def _same_cell(a: Any, b: Any) -> bool:
    if isinstance(a, str) or isinstance(b, str):
        return _text(a) == _text(b)
    return a == b


def _at(row: list[Any], i: int | None) -> Any:
    return row[i] if i is not None and i < len(row) else None


def _allowed_test_inputs(
    variables: list[str], derived: list[tuple[str, str]], la: str, ra: str, outputs: list[str]
) -> set[str]:
    """测试用例里可以给值的名字: 规则引用的变量、派生变量、内置变量、输出列的原值 (<左别名>.<输出列>)。"""
    names = set(variables) | {k for k, _ in derived} | {f"{ra}.{EXIST_VAR}", BASE_DATE_VAR}
    return names | {f"{la}.{o}" for o in outputs}


def _check_test_columns(
    where: str, inputs: list[str], expects: list[str], allowed: set[str], outputs: list[str]
) -> None:
    """期望列名不是输出列时整列会被忽略、输入列名不是变量时填了也没用: 两种情况都会让用例假通过, 直接报错。"""
    bad = [n for n in expects if n not in outputs]
    if bad:
        raise ValueError(
            f"{where}: 期望列「{bad[0]}」不是输出列 (输出列: {', '.join(outputs)}); "
            f"期望列对不上时会被忽略, 用例就会假通过"
        )
    bad = [n for n in inputs if n not in allowed]
    if bad:
        raise ValueError(
            f"{where}: 输入列「{bad[0]}」不是规则引用到的变量, 填了也不起作用 (可用: {', '.join(sorted(allowed))})"
        )


# ---------------------------------------------------------------------------
# Excel 决策表
# ---------------------------------------------------------------------------


def _rows(ws: Any) -> list[list[Any]]:
    """读整张 sheet。合并单元格只有左上格有值, 这里把它展开到区域内每一格 (续行的规则 ID / 输出格常被合并)。"""
    grid = [list(r) for r in ws.iter_rows(values_only=True)]
    for rng in ws.merged_cells.ranges:
        val = _at(grid[rng.min_row - 1], rng.min_col - 1) if rng.min_row - 1 < len(grid) else None
        for r in range(rng.min_row - 1, min(rng.max_row, len(grid))):
            for c in range(rng.min_col - 1, min(rng.max_col, len(grid[r]))):
                grid[r][c] = val
    return grid


def _read_kv(ws: Any) -> dict[str, Any]:
    out: dict[str, Any] = {}
    for row in _rows(ws)[1:]:
        if row and row[0] not in (None, "") and len(row) > 1 and row[1] not in (None, ""):
            out[str(row[0]).strip()] = row[1]
    return out


def _parse_header(
    grid: list[list[Any]], sheet: str, key_kind: str, allowed: tuple[str, ...]
) -> tuple[list[tuple[int, str, str]], int, int]:
    """定位表头: 第一个含 key_kind (决策表「规则ID」/ 测试「用例」) 的行是列类型行, 下一行是列名行, 再下一行起是正文。

    之前的行是元信息 (命中策略 / 说明), 可有可无。
    返回 (列清单 (列索引, 类型, 名字), 列类型行号, 正文起始行号), 行号 0 起。
    """
    kind_row = next((i for i, row in enumerate(grid) if any(_text(v) == key_kind for v in row)), None)
    kinds_doc = " / ".join(allowed)
    if kind_row is None:
        raise ValueError(
            f"「{sheet}」sheet 找不到列类型行: 需要有一行标出每列的类型 ({kinds_doc}), "
            f"其中一格写「{key_kind}」, 下一行写列名"
        )
    if kind_row + 1 >= len(grid):
        raise ValueError(f"「{sheet}」sheet 第 {kind_row + 1} 行是列类型行, 但下面没有列名行")
    kinds, names = grid[kind_row], grid[kind_row + 1]
    cols: list[tuple[int, str, str]] = []
    for i, k in enumerate(kinds):
        kind = _text(k)
        if not kind:
            continue  # 没标类型的列不参与解析, 可以放备注
        at = f"{get_column_letter(i + 1)}{kind_row + 1}"
        if kind not in allowed:
            raise ValueError(f"「{sheet}」sheet {at} 格的列类型「{kind}」不认识, 只认: {kinds_doc}")
        name = _text(_at(names, i))
        if not name:
            below = f"{get_column_letter(i + 1)}{kind_row + 2}"
            raise ValueError(f"「{sheet}」sheet {at} 格标了「{kind}」列, 但 {below} 格没有写列名")
        cols.append((i, kind, name))
    keys = [i for i, k, _ in cols if k == key_kind]
    if len(keys) != 1:
        raise ValueError(f"「{sheet}」sheet 必须恰有一列类型为「{key_kind}」, 现在有 {len(keys)} 列")
    return cols, kind_row, kind_row + 2


def _find_policy(grid: list[list[Any]], kind_row: int) -> Any:
    """列类型行之上是元信息行: 找「命中策略」标签, 取它右边第一个非空格; 没有标签时认一个单独写着策略名的格。"""
    fallback = None
    for row in grid[:kind_row]:
        cells = [_text(v).rstrip(":：") for v in row]
        if POLICY_LABEL in cells:
            j = cells.index(POLICY_LABEL)
            return next((c for c in cells[j + 1 :] if c), "")
        fallback = fallback or next((c for c in cells if c in HIT_POLICIES), None)
    return fallback


def load_excel_rules(path: Path) -> RuleSet:
    # data_only=False: 以 = 开头的格 (= 活跃 / =主表.金额 * 1.1) 在 Excel 里会被当成公式存起来,
    # data_only=True 读到的是公式的缓存结果 (没算过是 None, 算过是 #NAME?), 条件就静默变成「不关心」。
    # 这里按原文读, 规则工作簿里的 = 一律是 rule-sheet 表达式, 不支持写 Excel 公式。
    wb = load_workbook(path, data_only=False)
    for sheet in ("匹配设置", "决策表"):
        if sheet not in wb.sheetnames:
            raise ValueError(f"规则文件缺少「{sheet}」sheet: {path.name}")
    match = _read_kv(wb["匹配设置"])
    la, ra = str(match.get("左表别名", "主表")), str(match.get("右表别名", "本月"))
    derived: list[tuple[str, str]] = []
    if "变量定义" in wb.sheetnames:
        derived = [
            (str(r[0]).strip(), str(r[1]).strip())
            for r in _rows(wb["变量定义"])[1:]
            if r and r[0] not in (None, "") and len(r) > 1 and r[1] not in (None, "")
        ]

    grid = _rows(wb["决策表"])
    cols, kind_row, body_start = _parse_header(grid, "决策表", "规则ID", RULE_KINDS)
    hit_policy = _check_policy(_find_policy(grid, kind_row))
    id_col = next(i for i, k, _ in cols if k == "规则ID")
    pri_col = next((i for i, k, _ in cols if k == "优先级"), None)
    desc_col = next((i for i, k, _ in cols if k == "说明"), None)
    inputs = [(i, n) for i, k, n in cols if k == "输入"]
    outputs = [(i, n) for i, k, n in cols if k == "输出"]

    # 相邻的同 ID 行是同一条规则的多个「或」分支: 每行内各输入格 AND, 行间 OR; 优先级 / 输出 / 说明以首行为准
    rules: list[Rule] = []
    branches: list[list[list[Condition]]] = []
    first_row: list[int] = []
    seen: dict[str, int] = {}
    for offset, row in enumerate(grid[body_start:]):
        excel_row = body_start + offset + 1
        rid = _text(_at(row, id_col))
        if not rid:
            continue
        pri_raw = _at(row, pri_col)
        outs = {name: _at(row, i) for i, name in outputs}
        conds = [cell_condition(name, _at(row, i)) for i, name in inputs]
        if rules and rid == rules[-1].rule_id:
            rule = rules[-1]
            where = f"规则 {rid} 的续行 (第 {excel_row} 行)"
            head = f"首行 (第 {first_row[-1]} 行)"
            if not _blank(pri_raw) and float(pri_raw) != rule.priority:
                raise ValueError(f"{where} 优先级与{head}不一致; 续行的优先级留空或与首行相同")
            for name, v in outs.items():
                if not _blank(v) and not _same_cell(v, rule.outputs.get(name)):
                    first = _text(rule.outputs.get(name)) or "-"
                    raise ValueError(
                        f"{where} 输出「{name}」= {_text(v)}, 与{head}的 {first} 不一致; "
                        f"同一规则的续行只写条件, 输出留空或与首行相同"
                    )
            branches[-1].append(conds)
            continue
        if rid in seen:
            raise ValueError(
                f"规则ID「{rid}」在第 {seen[rid]} 行和第 {excel_row} 行重复出现; "
                f"同一规则的多个条件分支 (或) 必须写在相邻的行"
            )
        seen[rid] = excel_row
        rules.append(
            Rule(
                rule_id=rid,
                priority=1000.0 + offset if _blank(pri_raw) else float(pri_raw),
                conditions=conds,
                outputs=outs,
                desc=_text(_at(row, desc_col)),
            )
        )
        branches.append([conds])
        first_row.append(excel_row)
    for rule, brs in zip(rules, branches, strict=True):
        if len(brs) > 1:
            rule.conditions = [any_of([all_of(b) for b in brs])]

    known = [k for k, _ in derived] + list(FUNCS) + [BASE_DATE_VAR]
    input_names = [n for _, n in inputs]
    extra = sorted({n for _, e in derived for n in find_names(e, (la, ra), known)} - set(input_names))
    variables = input_names + extra  # 表头输入列 + 派生变量表达式里引用的列, 用例没给的都视为空
    output_names = [n for _, n in outputs]

    tests: list[TestCase] = []
    if "测试" in wb.sheetnames:
        tgrid = _rows(wb["测试"])
        tcols, _, tstart = _parse_header(tgrid, "测试", "用例", TEST_KINDS)
        case_col = next(i for i, k, _ in tcols if k == "用例")
        t_inputs = [(i, n) for i, k, n in tcols if k == "输入"]
        t_expect = [(i, n) for i, k, n in tcols if k == "期望"]
        allowed = _allowed_test_inputs(variables, derived, la, ra, output_names)
        _check_test_columns("「测试」sheet", [n for _, n in t_inputs], [n for _, n in t_expect], allowed, output_names)
        for row in tgrid[tstart:]:
            if not _text(_at(row, case_col)):
                continue
            tests.append(
                TestCase(
                    name=_text(_at(row, case_col)),
                    inputs={n: _at(row, i) for i, n in t_inputs},
                    expected={n: _at(row, i) for i, n in t_expect},
                )
            )
        if not tests:
            raise ValueError("「测试」sheet 里没有任何用例 (「用例」列为空的行会跳过); 写上用例, 或删掉这张 sheet")

    return RuleSet(match, derived, hit_policy, rules, output_names, tests, source=path.name, variables=variables)


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
        rid = str(r.get("编号", f"R{order + 1}"))
        if any(x.rule_id == rid for x in rules):
            raise ValueError(f"规则编号「{rid}」重复; 一条规则的多个条件分支请写在它的 当: 里 (任一满足:)")
        rules.append(
            Rule(
                rule_id=rid,
                priority=float(r.get("优先级", 1000 + order)),
                conditions=[build_condition(conds)] if conds else [],
                outputs=dict(then),
                desc=str(r.get("名称", "")),
            )
        )
        for col in then:
            if col not in outputs:
                outputs.append(col)
    derived = [(str(k), str(v)) for k, v in (doc.get("变量") or {}).items()]
    match = dict(doc.get("匹配") or {})
    la, ra = str(match.get("左表别名", "主表")), str(match.get("右表别名", "本月"))
    # 扫描 当 / 则 / 变量 里引用到的 <别名>.<列> 变量名; YAML 里新增写变量名的位置要同步加进 texts
    texts = [s for r in doc.get("规则") or [] for s in strings_in(r.get("当"))]
    texts += [str(v) for r in doc.get("规则") or [] for v in (r.get("则") or {}).values()]
    texts += [v for _, v in derived]
    known = [k for k, _ in derived] + list(FUNCS) + [BASE_DATE_VAR]
    variables = sorted({n for t in texts for n in find_names(t, (la, ra), known)} | {k for k, _ in derived})

    tests: list[TestCase] = []
    allowed = _allowed_test_inputs(variables, derived, la, ra, outputs)
    for i, t in enumerate(doc.get("测试") or []):
        name = str(t.get("用例", f"T{i + 1}"))
        inputs = {str(k): v for k, v in (t.get("输入") or {}).items()}
        expected = {str(k): v for k, v in (t.get("期望") or {}).items()}
        _check_test_columns(f"测试用例「{name}」", list(inputs), list(expected), allowed, outputs)
        tests.append(TestCase(name, inputs, expected))
    if "测试" in doc and not tests:
        raise ValueError("YAML 的 测试: 下没有任何用例; 写上用例, 或删掉这个键")

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
