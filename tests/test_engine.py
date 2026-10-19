"""执行核心: 命名空间构建、命中策略合并、输出格解析、测试用例运行。"""

import pandas as pd
import pytest
from openpyxl import Workbook

from rule_sheet.conditions import cell_condition, expr_condition
from rule_sheet.engine import add_derived, apply_rules, build_namespace, read_table, run_tests
from rule_sheet.model import Rule, RuleSet, TestCase

# ---------------------------------------------------------------------------
# read_table
# ---------------------------------------------------------------------------


def test_read_table_keeps_text_cells_as_text(tmp_path):
    """pandas 3 默认会把文本格式的 "007" 读成 7、电话读成浮点; 业务数据不能接受。"""
    wb = Workbook()
    ws = wb.active
    ws.append(["编号", "电话", "金额", "日期"])
    ws.append(["007", "13800000001", 100, pd.Timestamp("2026-10-01").to_pydatetime()])
    ws.append(["008", None, 250, None])
    for row in ws.iter_rows(min_row=2, max_col=2):
        for c in row:
            c.number_format = "@"
    p = tmp_path / "t.xlsx"
    wb.save(p)
    df = read_table(p)
    assert df["编号"].tolist() == ["007", "008"]
    assert df["电话"].iloc[0] == "13800000001" and pd.isna(df["电话"].iloc[1])
    assert pd.api.types.is_integer_dtype(df["金额"]) and df["金额"].tolist() == [100, 250]
    assert pd.api.types.is_datetime64_any_dtype(df["日期"])


def make_rs(rules, outputs, hit_policy="优先级", **kw):
    return RuleSet(
        match=kw.pop("match", {}),
        derived=kw.pop("derived", []),
        hit_policy=hit_policy,
        rules=rules,
        outputs=outputs,
        **kw,
    )


# ---------------------------------------------------------------------------
# build_namespace
# ---------------------------------------------------------------------------


def test_namespace_prefixes_exist_flag_and_collision():
    left = pd.DataFrame({"编号": ["A", "B", "C"], "金额": [1, 2, 3], "备注": ["x", "y", "z"]})
    right = pd.DataFrame({"编号": ["A", "C"], "金额": [10, 30], "电话": ["t1", None]})
    rs = make_rs([], [], match={"基准日期": "2026-10-19"})
    ns = build_namespace(rs, left, right, ["编号"])
    assert ns["主表.金额"].tolist() == [1, 2, 3]
    assert ns["本月.金额"].tolist()[0] == 10 and pd.isna(ns["本月.金额"].iloc[1])
    assert ns["本月.存在"].tolist() == [True, False, True]
    assert "本月.编号" not in ns  # 匹配键只保留左侧
    assert ns["基准日期"] == pd.Timestamp("2026-10-19")


def test_namespace_aliases_from_match():
    left = pd.DataFrame({"k": [1]})
    right = pd.DataFrame({"k": [1], "v": [2]})
    rs = make_rs([], [], match={"左表别名": "L", "右表别名": "R"})
    ns = build_namespace(rs, left, right, ["k"])
    assert set(ns) >= {"L.k", "R.v", "R.存在", "基准日期"}


def test_duplicate_right_keys_policy():
    left = pd.DataFrame({"k": [1]})
    right = pd.DataFrame({"k": [1, 1], "v": ["first", "second"]})
    notes: list[str] = []
    ns = build_namespace(make_rs([], []), left, right, ["k"], notes)
    assert ns["本月.v"].tolist() == ["first"] and notes
    with pytest.raises(ValueError, match="重复键"):
        build_namespace(make_rs([], [], match={"右表重复键": "报错"}), left, right, ["k"])


def test_missing_key_column():
    with pytest.raises(ValueError, match="匹配键"):
        build_namespace(make_rs([], []), pd.DataFrame({"a": [1]}), pd.DataFrame({"b": [1]}), ["a"])


# ---------------------------------------------------------------------------
# apply_rules
# ---------------------------------------------------------------------------


def test_priority_merges_per_output_column(ns_two_rows):
    r1 = Rule("R1", 1, [cell_condition("主表.状态", "活跃")], {"备注": "来自R1", "电话": "-"})
    r2 = Rule("R2", 2, [cell_condition("主表.状态", "活跃")], {"备注": "来自R2", "电话": "本月.电话"})
    rs = make_rs([r2, r1], ["备注", "电话"])
    results, hit = apply_rules(rs, ns_two_rows, {"电话": ns_two_rows["主表.电话"]})
    assert results["备注"].tolist()[0] == "来自R1"  # 两条都写了, 优先级小的赢
    assert results["电话"].tolist()[0] == "139"  # 只有 R2 写了, 取 R2 (复制变量)
    assert results["电话"].tolist()[1] is None  # 未命中行保持原值 (原值为空)
    assert hit.tolist() == ["R1; R2", "", ""]


def test_output_kinds(ns_two_rows):
    r = Rule(
        "R",
        1,
        [cell_condition("本月.存在", "是")],
        {
            "a": "-",
            "b": '"引号字面值"',
            "c": "=主表.金额 * 2",
            "d": 42,
            "e": "本月.金额",
        },
    )
    rs = make_rs([r], ["a", "b", "c", "d", "e"])
    res, _ = apply_rules(rs, ns_two_rows, {"a": pd.Series(["keep"] * 3)})
    assert res["a"].tolist() == ["keep"] * 3
    assert res["b"].iloc[0] == "引号字面值" and res["b"].iloc[1] is None
    assert res["c"].iloc[0] == 20000.0
    assert res["d"].iloc[2] == 42
    assert res["e"].iloc[2] == 5100.0


def test_unique_policy_reports_label(ns_two_rows):
    r1 = Rule("R1", 1, [cell_condition("本月.存在", "是")], {"x": "1"})
    r2 = Rule("R2", 2, [expr_condition("主表.金额 > 100")], {"x": "2"})
    rs = make_rs([r1, r2], ["x"], hit_policy="唯一")
    with pytest.raises(ValueError, match="C001"):
        apply_rules(rs, ns_two_rows, {}, labels=pd.Series(["C001", "C002", "C003"]))


def test_unknown_variable_names_rule(ns_two_rows):
    r = Rule("R9", 1, [cell_condition("不存在的变量", "x")], {"x": "1"})
    with pytest.raises(KeyError, match="R9"):
        apply_rules(make_rs([r], ["x"]), ns_two_rows, {})


# ---------------------------------------------------------------------------
# add_derived / run_tests
# ---------------------------------------------------------------------------


def test_derived_respects_given_value_and_lenient(ns_two_rows):
    rs = make_rs([], [], derived=[("偏差率", "abs(本月.金额 - 主表.金额) / 主表.金额"), ("坏的", "没有这个变量 + 1")])
    ns = dict(ns_two_rows)
    ns["偏差率"] = pd.Series([9.9, 9.9, 9.9])
    add_derived(rs, ns, ns["主表.金额"].index, lenient=True)
    assert ns["偏差率"].tolist() == [9.9, 9.9, 9.9]
    assert ns["坏的"].isna().all()
    with pytest.raises(ValueError):
        add_derived(rs, dict(ns_two_rows), ns_two_rows["主表.金额"].index, lenient=False)


def test_run_tests_blank_expected_means_unchanged_and_missing_vars_are_null():
    r1 = Rule("R1", 1, [cell_condition("主表.状态", "活跃"), cell_condition("本月.电话", "非空")], {"备注": "有电话"})
    rs = make_rs(
        [r1],
        ["备注", "状态"],
        tests=[
            TestCase("命中", {"主表.状态": "活跃", "本月.电话": "1"}, {"备注": "有电话"}),
            TestCase("电话没给=空, 不命中", {"主表.状态": "活跃"}, {}),
            TestCase("故意写错", {"主表.状态": "活跃", "本月.电话": "1"}, {"备注": "别的"}),
            TestCase("状态保持原值", {"主表.状态": "休眠"}, {"状态": "休眠"}),
        ],
        variables=["主表.状态", "本月.电话"],
    )
    outcomes = run_tests(rs)
    assert [o.passed for o in outcomes] == [True, True, False, True]
    assert outcomes[0].hit == "R1" and outcomes[1].hit == ""
    assert "备注" in outcomes[2].problems[0]
