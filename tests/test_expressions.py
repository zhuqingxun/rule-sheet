"""表达式求值、字面量解析、顶层切分、变量名扫描。"""

import pandas as pd
import pytest

from rule_sheet.expressions import eval_expr, find_names, literal, split_top_level


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


# ---------------------------------------------------------------------------
# 字段名含 - 等特殊字符 (issue #1)
# ---------------------------------------------------------------------------


def test_hyphen_names_resolve_as_a_whole():
    ns = {
        "本期.类别-会议": pd.Series(["变动通报", None, "酝酿"], dtype=object),
        "本期.类别": pd.Series(["x", "y", "z"], dtype=object),  # 同时存在更短的前缀名也不能被抢
        "会议": pd.Series([1, 2, 3]),
    }
    assert eval_expr('本期.类别-会议 == "变动通报"', ns).tolist() == [True, False, False]
    assert eval_expr("为空(本期.类别-会议)", ns).tolist() == [False, True, False]
    assert eval_expr("本期.类别-会议 == 本期.类别-会议", ns).tolist() == [True, False, True]


def test_subtraction_without_spaces_still_works():
    ns = {"本期.金额": pd.Series([10.0, 5.0]), "上期.金额": pd.Series([4.0, 5.0]), "偏差率": pd.Series([1.0, 1.0])}
    assert eval_expr("本期.金额-上期.金额", ns).tolist() == [6.0, 0.0]
    assert eval_expr("本期.金额-偏差率", ns).tolist() == [9.0, 4.0]
    assert eval_expr("本期.金额-1", ns).tolist() == [9.0, 4.0]
    assert eval_expr("-本期.金额", ns).tolist() == [-10.0, -5.0]


def test_backtick_quotes_any_name():
    ns = {"本期.金额(万元)": pd.Series([150, 50]), "本期.岗位 名称": pd.Series(["a", "b"], dtype=object)}
    assert eval_expr("`本期.金额(万元)` > 100", ns).tolist() == [True, False]
    assert eval_expr('`本期.岗位 名称` == "a"', ns).tolist() == [True, False]
    with pytest.raises(ValueError, match="反引号里的变量「没有这个」未定义"):
        eval_expr("`没有这个` > 1", ns)


def test_name_error_hints_backticks():
    # 名字不在命名空间里时, 短前缀被替换、剩下的碎片触发 NameError, 提示要指出碎片并给出反引号写法
    with pytest.raises(ValueError, match="「会议」既不是变量也不是可用函数.*反引号"):
        eval_expr("本期.类别-会议 == 1", {"本期.类别": pd.Series([1])})
    with pytest.raises(ValueError, match="「float」既不是变量"):
        eval_expr("float('nan')", {"x": 1})


@pytest.mark.parametrize(
    ("text", "want"),
    [
        ('上期.类别-会议 == "变动通报"', ["上期.类别-会议"]),
        ("非空(本期.岗动-拟任命人-流程节点)", ["本期.岗动-拟任命人-流程节点"]),
        ("上期.类别-会议 == 本期.类别-会议", ["上期.类别-会议", "本期.类别-会议"]),
        ("本期.金额-上期.金额", ["本期.金额", "上期.金额"]),  # - 后紧跟 别名. = 减号, 不能把整个减法当成一个空变量
        ("本期.金额 - 上期.金额", ["本期.金额", "上期.金额"]),
        ("本期.金额-100", ["本期.金额"]),  # - 后紧跟数字 = 减号
        ("本期.金额-偏差率", ["本期.金额"]),  # - 后紧跟已知名字 (派生变量) = 减号
        ("本期.金额-abs(上期.金额)", ["本期.金额", "上期.金额"]),  # - 后紧跟函数名 = 减号
        ("本期.金额/上期.金额", ["本期.金额", "上期.金额"]),
        ("本期.岗位/职级 == '高级'", ["本期.岗位/职级"]),
        ("本期.金额（万元）> 100", ["本期.金额（万元）"]),  # 全角括号永远不是 Python 语法, 算名字
        ("`本期.金额(万元)` > 100 + 本期.x", ["本期.金额(万元)", "本期.x"]),  # 反引号内容原样作为名字
        ("偏差率 > 0.05", []),
    ],
)
def test_find_names(text, want):
    assert find_names(text, ("本期", "上期"), known=["偏差率", "abs"]) == want


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
