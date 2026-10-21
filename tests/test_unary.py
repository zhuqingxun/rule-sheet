"""单元格条件语法 (FEEL 一元测试子集) 的逐项覆盖。"""

from datetime import datetime

import pandas as pd
import pytest

from rule_sheet.conditions import unary_test


def s(*vals):
    has_str = any(isinstance(v, str) for v in vals)
    return pd.Series(list(vals), dtype=object if has_str or all(v is None for v in vals) else None)


def hits(cell, var, ns=None):
    return unary_test(cell, var, ns or {}).tolist()


@pytest.mark.parametrize("cell", [None, "", "-", "—", "任意", "  -  "])
def test_blank_means_dont_care(cell):
    assert hits(cell, s("活跃", None, 1)) == [True, True, True]


def test_literal_equality_and_null():
    assert hits("活跃", s("活跃", "试用", None)) == [True, False, False]


def test_numeric_cell_from_excel():
    assert hits(0.05, s(0.05, 0.1, None)) == [True, False, False]
    assert hits("180", s(180, 181)) == [True, False]


def test_comparisons_are_null_safe():
    assert hits("> 0.05", s(0.08, 0.03, None)) == [True, False, False]
    assert hits(">= 180", s(180, 179, None)) == [True, False, False]
    assert hits("!= 休眠", s("活跃", "休眠", None)) == [True, False, False]
    assert hits("<> 休眠", s("活跃", "休眠")) == [True, False]
    assert hits("= 活跃", s("活跃", "休眠")) == [True, False]


def test_comparison_against_another_variable():
    var = s("a", "b", None, "c")
    ns = {"主表.电话": s("x", "b", "c", None)}
    assert hits("!= 主表.电话", var, ns) == [True, False, False, False]
    assert hits("= 主表.电话", var, ns) == [False, True, False, False]


def test_dangling_variable_name_is_an_error_not_text():
    # 右侧长得像 <别名>.<列> 却不在命名空间里 (列名写错 / 表里没这列), 不能拿字面 "主表.电洁" 去比较然后静默不命中
    ns = {"主表.电话": s("x", "b")}
    with pytest.raises(KeyError, match="未定义的变量「主表.电洁」"):
        hits("!= 主表.电洁", s("a", "b"), ns)
    with pytest.raises(KeyError, match="未定义的变量「主表.上限」"):
        hits("[0..主表.上限]", s(1, 2), ns)
    assert hits("v1.2", s("v1.2", "v2"), ns) == [True, False]  # 没有叫 v1 的别名, 仍是普通文字
    assert hits('"主表.电洁"', s("主表.电洁", "b"), ns) == [True, False]  # 加引号 = 明确要当文字


def test_ranges_closed_and_open():
    var = s(0.02, 0.03, 0.05, 0.051, None)
    assert hits("[0.02..0.05]", var) == [True, True, True, False, False]
    assert hits("(0.02..0.05)", var) == [False, True, False, False, False]
    assert hits("[0.02..0.05)", var) == [True, True, False, False, False]


def test_enumeration_is_or():
    assert hits("VIP, 金卡", s("VIP", "金卡", "普通", None)) == [True, True, False, False]
    assert hits("> 10, < 0", s(11, -1, 5)) == [True, True, False]


def test_quoted_literal_keeps_comma():
    assert hits('"VIP 小幅偏差, 关注"', s("VIP 小幅偏差, 关注", "VIP")) == [True, False]


def test_null_tests():
    assert hits("非空", s("a", None, "")) == [True, False, True]
    assert hits("为空", s("a", None)) == [False, True]


def test_null_keyword_in_cells():
    """单元格里 空 / = 空 / != 空 就是判空 (与 为空 / 非空 同义)。
    空 不是值: 大小比较和区间端点报错; 要匹配文字「空」加引号。"""
    var = s("a", None, "")
    assert hits("空", var) == [False, True, False]
    assert hits("= 空", var) == [False, True, False]
    assert hits("== 空", var) == [False, True, False]
    assert hits("!= 空", var) == [True, False, True]
    assert hits("<> 空", var) == [True, False, True]
    assert hits("not(空)", var) == [True, False, True]
    assert hits("活跃, 空", s("活跃", None, "x")) == [True, True, False]
    assert hits('"空"', s("空", None)) == [True, False]
    with pytest.raises(ValueError, match="空 只能和 = / != 搭配"):
        hits("> 空", s(1, 2))
    with pytest.raises(ValueError, match="区间端点不能是 空"):
        hits("[空..5]", s(1, 2))


def test_not_excludes_nulls_except_for_null_tests():
    assert hits("not(活跃)", s("活跃", "试用", None)) == [False, True, False]
    assert hits("not(VIP, 金卡)", s("VIP", "普通", None)) == [False, True, False]
    assert hits("not(为空)", s("a", None)) == [True, False]
    assert hits("not(非空)", s("a", None)) == [False, True]


def test_text_keywords_in_cells():
    """0.6: 包含 / 不包含 / 开头是 / 开头不是 / 结尾是 / 结尾不是, 后面跟空格或括号; 逗号分隔的几项任一命中。
    参数取原文 (01 不会变成数字 1), 命中变量名的换成那一列; 空值不成立; not(...) 照常只对有值的格成立。"""
    var = s("集团CEO", "终端BG-总裁助理", None, "法务部")
    assert hits("包含 CEO, 总裁", var) == [True, True, False, False]
    assert hits("包含(CEO, 总裁)", var) == [True, True, False, False]
    assert hits("不包含 助理", var) == [True, False, False, True]
    assert hits("not(包含 CEO)", var) == [False, True, False, True]
    assert hits("开头是 集团", var) == [True, False, False, False]
    assert hits("开头不是 集团, 终端", var) == [False, False, False, True]
    assert hits("结尾是 部", var) == [False, False, False, True]
    assert hits("结尾不是 部", var) == [True, True, False, False]
    assert hits('包含 "CEO, 总裁"', var) == [False, False, False, False]  # 引号里的逗号不拆
    assert hits("包含测试", s("包含测试", "x")) == [True, False]  # 没有空格 / 括号就是普通文字
    assert hits('"包含 CEO"', s("包含 CEO", "x")) == [True, False]
    codes = s("007", None, "010", "1")
    assert hits("开头是 01", codes) == [False, False, True, False]  # 01 保持文本, 不会变成数字 1
    assert hits("包含 主表.工号", codes, {"主表.工号": s("7", "x", "1", None)}) == [True, False, True, False]
    with pytest.raises(KeyError, match="未定义的变量「主表.工资」"):
        hits("包含 主表.工资", codes, {"主表.工号": s("7", "x", "1", None)})
    with pytest.raises(ValueError, match="包含 后面要写至少一个文本"):
        hits("包含 ,", var)


def test_bool_column_accepts_yes_no_words():
    var = pd.Series([True, False, True])
    assert hits("是", var) == [True, False, True]
    assert hits("否", var) == [False, True, False]


def test_yes_no_on_text_column_is_plain_text():
    assert hits("是", s("是", "否")) == [True, False]


def test_date_cell_and_date_literal():
    var = pd.Series(pd.to_datetime(["2026-10-08", "2026-10-09", None]))
    assert hits(datetime(2026, 10, 8), var) == [True, False, False]
    assert hits("2026-10-09", var) == [False, True, False]
    assert hits("> 2026-10-08", var) == [False, True, False]
