"""两种前端的加载: 结构、变量扫描、条件组合。"""

from datetime import date, datetime

import pandas as pd
import pytest

from rule_sheet.conditions import build_condition
from rule_sheet.loaders import load_rules, load_yaml_rules

from .conftest import EXCEL_RULES, YAML_RULES


def test_excel_structure():
    rs = load_rules(EXCEL_RULES)
    assert [r.rule_id for r in rs.rules] == ["R1", "R2", "R3", "R4", "R5a", "R5b"]
    assert rs.outputs == ["状态", "电话", "复核标记", "备注", "提醒"]
    assert rs.hit_policy == "优先级"
    assert len(rs.tests) == 8 and rs.tests[0].name == "T1"
    assert rs.variables == ["主表.状态", "主表.等级", "偏差率", "本月.存在", "距上次活跃天数", "本月.电话", "主表.电话"]
    assert [d for d, _ in rs.derived] == ["偏差率", "距上次活跃天数"]
    assert isinstance(rs.match["基准日期"], (date, datetime))
    assert rs.match["匹配键"] == "客户编号"


def test_yaml_structure_and_variable_scan():
    rs = load_rules(YAML_RULES)
    assert [r.rule_id for r in rs.rules] == ["R1", "R2", "R3", "R4", "R5"]
    assert rs.outputs == ["复核标记", "备注", "状态", "电话", "提醒"]
    assert len(rs.tests) == 8
    # 嵌套在 任一满足 里的变量也要被扫到, 测试用例没给时才会被置空而不是报未定义
    assert {"本月.电话", "主表.电话", "本月.存在", "偏差率", "距上次活跃天数"} <= set(rs.variables)


def test_nested_combinators_evaluate(ns_two_rows):
    cond = build_condition(
        [
            "本月.存在 == 是",
            {"任一满足": ["为空(本月.电话)", "为空(主表.电话)"]},
        ]
    )
    assert cond(ns_two_rows).tolist() == [False, False, True]
    cond = build_condition({"不满足": {"任一满足": ['主表.状态 == "活跃"', '主表.状态 == "休眠"']}})
    assert cond(ns_two_rows).tolist() == [False, False, True]
    cond = build_condition({"全部满足": "本月.存在 == 是"})  # 单个值也接受
    assert cond(ns_two_rows).tolist() == [True, False, True]


def test_unknown_combinator_rejected():
    with pytest.raises(ValueError, match="无法识别"):
        build_condition({"或者": ["a == 1"]})
    with pytest.raises(ValueError, match="无法识别"):
        build_condition(42)


def test_yaml_minimal_document(tmp_path):
    p = tmp_path / "r.yaml"
    p.write_text(
        "规则:\n  - 当:\n      - x > 1\n    则:\n      y: 2\n",
        encoding="utf-8",
    )
    rs = load_yaml_rules(p)
    assert rs.rules[0].rule_id == "R1" and rs.rules[0].priority == 1000
    assert rs.outputs == ["y"] and rs.hit_policy == "优先级"
    mask = rs.rules[0].conditions[0]({"x": pd.Series([0, 2])})
    assert mask.tolist() == [False, True]


def test_yaml_hyphen_names_and_unspaced_subtraction(tmp_path):
    """issue #1 的现场: 字段名含 -, 用例没给该变量时要被扫描到并置空, 而不是被截成 上期.类别 留下裸的 会议。
    同时无空格的减法不能被扫描器整体当成一个变量名吞掉。"""
    p = tmp_path / "r.yaml"
    p.write_text(
        "匹配: {左表别名: 本期, 右表别名: 上期}\n"
        "变量:\n  差额: 本期.金额-上期.金额\n"
        "规则:\n"
        '  - 编号: S1\n    当:\n      - 上期.类别-会议 == "变动通报"\n    则:\n      状态-会议: 02-进展更新\n'
        "  - 编号: S2\n    当:\n      - 差额 > 0\n    则:\n      备注: 涨了\n"
        "测试:\n"
        "  - 用例: 用例没给 上期.类别-会议, 视为空, 不命中\n"
        "    输入: {本期.金额: 10, 上期.金额: 4}\n"
        "    期望: {备注: 涨了}\n"
        "  - 用例: 给了就命中\n"
        "    输入: {上期.类别-会议: 变动通报, 本期.金额: 1, 上期.金额: 4}\n"
        "    期望: {状态-会议: 02-进展更新}\n",
        encoding="utf-8",
    )
    rs = load_yaml_rules(p)
    assert rs.variables == ["上期.类别-会议", "上期.金额", "差额", "本期.金额"]
    from rule_sheet.engine import run_tests

    outcomes = run_tests(rs)
    assert [(o.passed, o.hit, o.error) for o in outcomes] == [(True, "S2", ""), (True, "S1", "")]


def test_yaml_rule_without_then_and_policy_validation(tmp_path):
    p = tmp_path / "r.yaml"
    p.write_text(
        "命中策略: 首次命中\n规则:\n  - 编号: G\n    当:\n      - x > 1\n    则:\n  - 则: {y: 2}\n",
        encoding="utf-8",
    )
    rs = load_yaml_rules(p)
    assert rs.hit_policy == "首次命中"
    assert rs.rules[0].outputs == {} and rs.rules[1].outputs == {"y": 2} and rs.outputs == ["y"]
    p.write_text("命中策略: 首个\n规则: []\n", encoding="utf-8")
    with pytest.raises(ValueError, match="未知的命中策略「首个」.*首次命中"):
        load_yaml_rules(p)


def test_excel_policy_validation(tmp_path):
    from openpyxl import Workbook

    wb = Workbook()
    wb.active.title = "匹配设置"
    ws = wb.create_sheet("决策表")
    ws.append(["命中策略", "first"])
    ws.append(["规则ID"])
    ws.append(["规则ID"])
    wb.save(tmp_path / "bad.xlsx")
    with pytest.raises(ValueError, match="未知的命中策略「first」"):
        load_rules(tmp_path / "bad.xlsx")


def test_missing_file_and_missing_sheet(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_rules(tmp_path / "nope.xlsx")
    from openpyxl import Workbook

    wb = Workbook()
    wb.active.title = "随便"
    wb.save(tmp_path / "bad.xlsx")
    with pytest.raises(ValueError, match="匹配设置"):
        load_rules(tmp_path / "bad.xlsx")
