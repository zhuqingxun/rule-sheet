"""表达式求值、字面量解析、顶层切分、变量名扫描。"""

import pandas as pd
import pytest

from rule_sheet.expressions import eval_expr, find_names, literal, split_top_level, truth


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
# AST 改写: 逻辑连接词按元素运算 (issue #7), != / in 空值不成立 (issue #5), 链式比较
# ---------------------------------------------------------------------------


def test_logic_words_work_on_whole_columns(ns_two_rows):
    want = [True, False, False]  # 行 1 主表电话空, 行 2 本月电话空
    assert eval_expr("非空(本月.电话) 且 非空(主表.电话) 且 本月.电话 != 主表.电话", ns_two_rows).tolist() == want
    assert eval_expr("非空(本月.电话) and 非空(主表.电话) and (本月.电话 != 主表.电话)", ns_two_rows).tolist() == want
    assert eval_expr('主表.状态 == "活跃" 或 本月.存在 == 否', ns_two_rows).tolist() == [True, True, False]
    assert eval_expr('主表.状态 == "活跃" or 本月.存在 == 否', ns_two_rows).tolist() == [True, True, False]
    assert eval_expr('非 (主表.状态 == "活跃")', ns_two_rows).tolist() == [False, True, True]  # 纯取反, 空值行也为真
    assert eval_expr("not 本月.存在", ns_two_rows).tolist() == [False, True, False]
    assert eval_expr("非空(本月.电话) 且 非空(主表.电话)", {"本月.电话": "a", "主表.电话": None}) is False  # 标量也行


def test_not_equal_is_null_safe_like_cells(ns_two_rows):
    # pandas 原生 != 对空值返回 True; 改写后任一侧为空就不成立, 与决策表单元格的 != 一致
    assert eval_expr("本月.电话 != 主表.电话", ns_two_rows).tolist() == [True, False, False]
    assert eval_expr('主表.状态 != "活跃"', ns_two_rows).tolist() == [False, True, False]
    assert eval_expr("本月.金额 != 1", ns_two_rows).tolist() == [True, False, True]  # NaN != 1 不成立
    assert eval_expr("a != 1", {"a": None}) is False and eval_expr("a != 1", {"a": 2}) is True
    assert eval_expr("a == b", {"a": None, "b": None}) is False
    assert eval_expr("a > 1", {"a": None}) is False  # 标量 None > 1 在 Python 里直接报错, 这里是「不成立」
    assert eval_expr("a > b", {"a": pd.Series([1, 2]), "b": None}).tolist() == [False, False]


def test_chained_comparison_and_in(ns_two_rows):
    assert eval_expr("5000 <= 主表.金额 <= 10000", ns_two_rows).tolist() == [True, False, True]
    assert eval_expr("0 < 本月.金额 < 10500", ns_two_rows).tolist() == [False, False, True]
    assert eval_expr('主表.状态 in ("活跃", "试用")', ns_two_rows).tolist() == [True, False, False]
    assert eval_expr('主表.状态 not in ("活跃", "试用")', ns_two_rows).tolist() == [False, True, False]  # 空值不成立
    assert eval_expr('s in ("a",)', {"s": None}) is False


def test_null_keyword_is_a_null_test(ns_two_rows):
    """关键字 空: == 空 / != 空 就是判空 (与 为空 / 非空 同义); 它不是值, 其它运算一律报错。"""
    assert eval_expr("主表.电话 == 空", ns_two_rows).tolist() == [False, True, False]
    assert eval_expr("空 == 主表.电话", ns_two_rows).tolist() == [False, True, False]
    assert eval_expr("本月.电话 != 空", ns_two_rows).tolist() == [True, True, False]
    assert eval_expr("上期 == 空", {"上期": None}) is True and eval_expr("上期 != 空", {"上期": 1}) is True
    assert eval_expr("空 == 空", {"x": 1}) is True
    assert eval_expr('属于(主表.状态, "活跃", 空)', ns_two_rows).tolist() == [True, False, True]
    assert eval_expr('主表.状态 in ("活跃", 空)', ns_two_rows).tolist() == [True, False, True]
    assert eval_expr('主表.状态 not in ("活跃", 空)', ns_two_rows).tolist() == [False, True, False]
    assert eval_expr("主表.电话 == 本月.电话", {"主表.电话": None, "本月.电话": None}) is False  # 两个空变量仍不相等
    for bad in ("主表.金额 > 空", "空 + 1", "abs(空)"):
        with pytest.raises(ValueError, match="空 只能和 == / != 搭配"):
            eval_expr(bad, ns_two_rows)


def test_string_literals_are_left_alone():
    ns = {"主表.状态": pd.Series(["主表.状态", "或", "x"], dtype=object)}
    assert eval_expr('主表.状态 == "主表.状态"', ns).tolist() == [True, False, False]  # 字符串里的变量名不替换
    assert eval_expr('主表.状态 == "或"', ns).tolist() == [False, True, False]  # 字符串里的连接词不替换
    assert eval_expr("主表.状态 == '或' 或 主表.状态 == 'x'", ns).tolist() == [False, True, True]


def test_syntax_and_truth_value_errors_are_readable(ns_two_rows):
    with pytest.raises(ValueError, match="表达式语法错误"):
        eval_expr("本月.金额 >", ns_two_rows)
    with pytest.raises(ValueError, match="真假判断.*且 / 或 / 非"):
        eval_expr("1 if 本月.存在 else 0", ns_two_rows)


# ---------------------------------------------------------------------------
# 0.6: 文本函数 / 拼接 / 如果 / 数字 / N天 (issue #9 #10 #11 #14)
# ---------------------------------------------------------------------------


@pytest.fixture
def ns_text():
    idx = pd.RangeIndex(5)
    return {
        "本期.岗位名称": pd.Series(["集团CEO", "终端BG-总裁助理", None, "法务部", 7], index=idx, dtype=object),
        "本期.一层组织": pd.Series(["1-集团", "2-终端", "3-华东", None, "11-控制"], index=idx, dtype=object),
        "本期.职级": pd.Series(["22级", "23", "P24", None, 25], index=idx, dtype=object),
        "本期.推荐时间": pd.Series(
            pd.to_datetime(["2026-06-01", "2025-10-01", "2025-09-30", None, "2026-10-01"]), index=idx
        ),
        "本期.工号": pd.Series(["007", None, "009", "010", "011"], index=idx, dtype=object),
        "基准日期": pd.Timestamp("2026-10-01"),
    }


def test_text_functions_are_literal_and_null_safe(ns_text):
    """包含 / 开头是 / 结尾是 按字面匹配 (不是正则), 任一子串命中即真; 空值不成立, 否定形式也不成立。"""
    assert eval_expr('包含(本期.岗位名称, "CEO", "总裁")', ns_text).tolist() == [True, True, False, False, False]
    assert eval_expr('不包含(本期.岗位名称, "助理")', ns_text).tolist() == [True, False, False, True, True]
    assert eval_expr('开头是(本期.一层组织, "2-", "3-")', ns_text).tolist() == [False, True, True, False, False]
    assert eval_expr('开头不是(本期.一层组织, "1")', ns_text).tolist() == [False, True, True, False, False]
    assert eval_expr('结尾是(本期.一层组织, "东")', ns_text).tolist() == [False, False, True, False, False]
    assert eval_expr('结尾不是(本期.一层组织, "东")', ns_text).tolist() == [True, True, False, False, True]
    assert eval_expr('非 包含(本期.岗位名称, "CEO")', ns_text).tolist() == [False, True, True, True, True]  # 纯取反
    assert eval_expr('包含(本期.工号, "00")', ns_text).tolist() == [True, False, True, False, False]  # 文本列保住前导零
    by_column = eval_expr("包含(本期.工号, 本期.工号)", ns_text)  # 子串也可以是列
    assert by_column.tolist() == [True, False, True, True, True]
    assert eval_expr('包含(x, ".")', {"x": "a.b"}) is True and eval_expr('包含(x, ".")', {"x": "ab"}) is False
    assert eval_expr('包含(x, "a")', {"x": None}) is False and eval_expr('不包含(x, "a")', {"x": None}) is False
    assert eval_expr("不包含(x, y)", {"x": "abc", "y": None}) is False  # 否定形式下子串为空也不成立
    assert eval_expr('包含(x, "7")', {"x": 7.0}) is True  # 非文本值按 7 而不是 7.0 匹配
    with pytest.raises(ValueError, match="至少要给一个文本"):
        eval_expr("包含(本期.岗位名称)", ns_text)
    with pytest.raises(ValueError, match="参数不能是 空"):
        eval_expr('包含(空, "a")', ns_text)


def test_concat_treats_null_as_empty_and_all_null_as_null(ns_text):
    out = eval_expr('拼接(本期.一层组织, "/", 本期.岗位名称)', ns_text)
    assert out.tolist()[:2] == ["1-集团/集团CEO", "2-终端/终端BG-总裁助理"]
    assert out.tolist()[2:] == ["3-华东/", "/法务部", "11-控制/7"]
    out = eval_expr("拼接(本期.工号, 本期.工号)", ns_text)
    assert out.iloc[0] == "007007" and pd.isna(out.iloc[1])  # 全部参数为空 -> 空值, 不是空串
    assert eval_expr('拼接("x", 1.0, 日期)', {"日期": pd.Timestamp("2026-10-01")}) == "x12026-10-01"
    assert eval_expr('拼接("x", 本期.名称) == "xa"', {"本期.名称": pd.Series(["a", None], dtype=object)}).tolist() == [
        True,
        False,
    ]
    with pytest.raises(ValueError, match="参数不能是 空"):
        eval_expr("拼接(本期.工号, 空)", ns_text)


def test_if_is_vectorized_and_null_goes_to_else(ns_text):
    out = eval_expr('如果(本期.一层组织 != 空, 拼接(本期.一层组织, "/", 本期.岗位名称), 本期.岗位名称)', ns_text)
    assert out.tolist() == ["1-集团/集团CEO", "2-终端/终端BG-总裁助理", "3-华东/", "法务部", "11-控制/7"]
    nested = eval_expr('如果(包含(本期.岗位名称, "CEO"), "高管", "普通")', ns_text)
    assert nested.tolist()[:3] == ["高管", "普通", "普通"]
    out = eval_expr("如果(数字(本期.职级) >= 23, 数字(本期.职级), 空)", ns_text)
    assert out.tolist()[1:3] == [23.0, 24.0] and out.iloc[0] is None and out.iloc[3] is None  # 分支可以写 空
    assert eval_expr('如果(本期.岗位名称 == "法务部", 1, 2.5)', ns_text).tolist() == [2.5, 2.5, 2.5, 1.0, 2.5]
    out = eval_expr('如果(本期.岗位名称 == "集团CEO", 本期.推荐时间, 基准日期)', ns_text)  # 日期列 / 日期标量混用
    assert out.tolist()[:2] == [pd.Timestamp("2026-06-01"), pd.Timestamp("2026-10-01")]
    assert eval_expr("如果(是, 1, 2)", {"x": 1}) == 1 and eval_expr("如果(否, 1, 2)", {"x": 1}) == 2  # 标量
    assert eval_expr("如果(x > 1, 'a', 'b')", {"x": None}) == "b"  # 条件为空 = 假
    assert eval_expr("如果(x > 1, 'a', 'b')", {"x": pd.Series([None, 2], dtype=object)}).tolist() == ["b", "a"]
    with pytest.raises(ValueError, match="三元写法用 如果"):
        eval_expr("1 if 本期.推荐时间 else 0", ns_text)


def test_number_extracts_first_number_from_text(ns_text):
    out = eval_expr("数字(本期.职级)", ns_text)
    assert out.tolist()[:3] == [22.0, 23.0, 24.0] and pd.isna(out.iloc[3]) and out.iloc[4] == 25
    assert eval_expr("数字(本期.职级) >= 23", ns_text).tolist() == [False, True, True, False, True]
    assert eval_expr("数字(x)", {"x": "第3档-2024"}) == 3 and eval_expr("数字(x)", {"x": "P1.5"}) == 1.5
    assert eval_expr("数字(x)", {"x": "无"}) is None and eval_expr("数字(x)", {"x": 4}) == 4
    assert eval_expr("数字(x)", {"x": pd.Series([1.5, None])}).tolist()[0] == 1.5


def test_days_literal(ns_text):
    assert eval_expr("本期.推荐时间 >= 基准日期 - 365天", ns_text).tolist() == [True, True, False, False, True]
    assert eval_expr("基准日期 - 1.5天", ns_text) == pd.Timestamp("2026-09-29 12:00:00")
    assert eval_expr("天数差(基准日期 + 30天, 基准日期)", ns_text) == 30
    assert eval_expr("天数差(基准日期, 本期.推荐时间) <= 365", ns_text).tolist() == [True, True, False, False, True]
    ns = {"本期.30天内": pd.Series([True, False]), "基准日期": pd.Timestamp("2026-10-01")}
    assert eval_expr("本期.30天内", ns).tolist() == [True, False]  # 列名里的 天 不受影响
    assert eval_expr('x == "3天"', {"x": "3天"}) is True  # 字符串里的不改写


def test_truth_treats_null_as_false():
    assert truth(pd.Series([float("nan"), 1.0, 0.0])).tolist() == [False, True, False]  # NaN.astype(bool) 本来是 True
    assert truth(pd.Series(["a", None, ""], dtype=object)).tolist() == [True, False, False]
    assert truth(None) is False and truth(float("nan")) is False and truth("x") is True


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
        ('上期.x == "本期.y"', ["上期.x"]),  # 引号字符串里的不算
        (
            "本期.拟任人 != 上期.拟任人 或 (为空(上期.拟任人) 且 非空(本期.拟任人))",
            ["本期.拟任人", "上期.拟任人", "上期.拟任人", "本期.拟任人"],
        ),
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
