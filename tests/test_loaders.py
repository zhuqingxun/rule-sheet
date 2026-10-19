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


def test_missing_file_and_missing_sheet(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_rules(tmp_path / "nope.xlsx")
    from openpyxl import Workbook

    wb = Workbook()
    wb.active.title = "随便"
    wb.save(tmp_path / "bad.xlsx")
    with pytest.raises(ValueError, match="匹配设置"):
        load_rules(tmp_path / "bad.xlsx")
