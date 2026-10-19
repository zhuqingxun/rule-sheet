"""表达式求值、字面量解析、顶层切分。"""

import pandas as pd
import pytest

from rule_sheet.expressions import eval_expr, literal, split_top_level


def test_longest_name_wins(ns_two_rows):
    # 命名空间里同时有 金额 / 主表.金额 / 本月.金额, 不能把 主表.金额 里的 金额 替换掉
    out = eval_expr("abs(本月.金额 - 主表.金额) / 主表.金额", ns_two_rows)
    assert out.round(3).tolist()[0] == 0.08
    assert pd.isna(out.iloc[1])  # 本月缺失 -> 空
    assert eval_expr("金额 * 2", ns_two_rows).tolist() == [2.0, 4.0, 6.0]


def test_days_between_and_functions(ns_two_rows):
    days = eval_expr("天数差(基准日期, 主表.上次活跃)", ns_two_rows)
    assert days.tolist()[:2] == [10, 291]
    assert pd.isna(days.iloc[2])
    assert eval_expr("非空(本月.电话)", ns_two_rows).tolist() == [True, True, False]
    assert eval_expr("为空(主表.电话)", ns_two_rows).tolist() == [False, True, False]
    assert eval_expr('属于(主表.状态, "活跃", "试用")', ns_two_rows).tolist() == [True, False, False]
    assert eval_expr("本月.存在 == 是", ns_two_rows).tolist() == [True, False, True]


def test_scalar_namespace_works_too():
    assert eval_expr("偏差率 > 0.05", {"偏差率": 0.08}) is True
    assert eval_expr('属于(等级, "VIP")', {"等级": "VIP"}) is True


def test_failure_mentions_expression():
    with pytest.raises(ValueError, match="未定义变量"):
        eval_expr("未定义变量 > 1", {"x": 1})


def test_split_top_level_respects_quotes_and_brackets():
    assert split_top_level("VIP, 金卡") == ["VIP", "金卡"]
    assert split_top_level('"a, b", c') == ['"a, b"', "c"]
    assert split_top_level("not(a, b), c") == ["not(a, b)", "c"]
    assert split_top_level("[1..2], [3..4]") == ["[1..2]", "[3..4]"]
    assert split_top_level(" , ") == []


def test_literal_kinds():
    ns = {"主表.电话": pd.Series(["1"])}
    assert literal('"活跃"', ns) == "活跃"
    assert literal("活跃", ns) == "活跃"
    assert literal("180", ns) == 180
    assert literal("0.05", ns) == 0.05
    assert literal("2026-10-19", ns) == pd.Timestamp("2026-10-19")
    assert literal("主表.电话", ns) is ns["主表.电话"]
    assert literal("是", ns) == "是"  # 没有布尔列上下文时保持文本
    assert literal("是", ns, pd.Series([True])) is True
