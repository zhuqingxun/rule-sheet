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


def test_not_excludes_nulls_except_for_null_tests():
    assert hits("not(活跃)", s("活跃", "试用", None)) == [False, True, False]
    assert hits("not(VIP, 金卡)", s("VIP", "普通", None)) == [False, True, False]
    assert hits("not(为空)", s("a", None)) == [True, False]
    assert hits("not(非空)", s("a", None)) == [False, True]


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
