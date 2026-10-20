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


def test_output_dangling_variable_is_an_error(ns_two_rows):
    # 想复制 本月.电话 却写成 本月.电洁: 不能把字面 "本月.电洁" 写进结果; 报错带规则 ID 和输出列
    r = Rule("R7", 1, [cell_condition("本月.存在", "是")], {"电话": "本月.电洁", "备注": '"本月.电洁"'})
    with pytest.raises(KeyError, match="规则 R7 输出「电话」: 规则引用了未定义的变量「本月.电洁」"):
        apply_rules(make_rs([r], ["电话", "备注"]), ns_two_rows, {})
    r.outputs = {"备注": '"本月.电洁"'}  # 加引号 = 明确要当文字
    res, _ = apply_rules(make_rs([r], ["备注"]), ns_two_rows, {})
    assert res["备注"].iloc[0] == "本月.电洁"


def test_clear_output_writes_null(ns_two_rows):
    """输出格 清空 把目标列写成空值 (issue #2); - 仍是不改; 加引号的 "清空" 是普通文字。"""
    r = Rule("R", 1, [cell_condition("本月.存在", "是")], {"a": "清空", "b": "置空", "c": "-", "d": '"清空"'})
    rs = make_rs([r], ["a", "b", "c", "d"])
    keep = {k: pd.Series(["keep"] * 3) for k in "abcd"}
    res, _ = apply_rules(rs, ns_two_rows, keep)
    assert res["a"].tolist() == [None, "keep", None]
    assert res["b"].tolist() == [None, "keep", None]
    assert res["c"].tolist() == ["keep"] * 3
    assert res["d"].tolist() == ["清空", "keep", "清空"]


def test_clear_takes_part_in_priority_merge(ns_two_rows):
    # 清空 和普通输出一样参与逐列合并: 优先级数字小的 清空 赢过数字大的写值, 反过来也一样
    r1 = Rule("R1", 1, [cell_condition("本月.存在", "是")], {"x": "清空"})
    r2 = Rule("R2", 2, [cell_condition("本月.存在", "是")], {"x": "写了"})
    res, _ = apply_rules(make_rs([r1, r2], ["x"]), ns_two_rows, {"x": pd.Series(["old"] * 3)})
    assert res["x"].tolist() == [None, "old", None]
    swapped = [Rule("R1", 2, r1.conditions, r1.outputs), Rule("R2", 1, r2.conditions, r2.outputs)]
    res, _ = apply_rules(make_rs(swapped, ["x"]), ns_two_rows, {"x": pd.Series(["old"] * 3)})
    assert res["x"].tolist() == ["写了", "old", "写了"]


def test_first_match_policy_stops_at_first_hit(ns_two_rows):
    """首次命中 (issue #3): 每行只听优先级数字最小的那条; 没有输出的规则也能拦住一行。"""
    guard = Rule("G", 1, [cell_condition("主表.状态", "活跃")], {})  # 只拦截, 什么都不改
    r2 = Rule("R2", 2, [cell_condition("本月.存在", "是")], {"x": "来自R2", "y": "来自R2"})
    r3 = Rule("R3", 3, [], {"x": "兜底"})  # 无条件
    originals = {"x": pd.Series(["old"] * 3), "y": pd.Series(["old"] * 3)}
    res, hit = apply_rules(make_rs([r3, r2, guard], ["x", "y"], hit_policy="首次命中"), ns_two_rows, originals)
    # 行 0: 活跃 -> G 拦住, 不改; 行 1: 休眠, 本月不存在 -> R3 兜底; 行 2: 状态空, 本月存在 -> R2
    assert res["x"].tolist() == ["old", "兜底", "来自R2"]
    assert res["y"].tolist() == ["old", "old", "来自R2"]
    assert hit.tolist() == ["G", "R3", "R2"]  # 只记录赢的那条
    # 同一组规则换成 优先级 策略: G 不占任何列, 行 0 的 x 被 R2 写、y 也被 R2 写
    res, hit = apply_rules(make_rs([r3, r2, guard], ["x", "y"]), ns_two_rows, originals)
    assert res["x"].tolist() == ["来自R2", "兜底", "来自R2"]
    assert hit.tolist() == ["G; R2; R3", "R3", "R2; R3"]


def test_first_match_ties_follow_definition_order(ns_two_rows):
    a = Rule("A", 5, [], {"x": "A"})
    b = Rule("B", 5, [], {"x": "B"})
    res, hit = apply_rules(make_rs([b, a], ["x"], hit_policy="首次命中"), ns_two_rows, {})
    assert res["x"].tolist() == ["B"] * 3 and hit.tolist() == ["B"] * 3


def test_unique_policy_reports_label(ns_two_rows):
    r1 = Rule("R1", 1, [cell_condition("本月.存在", "是")], {"x": "1"})
    r2 = Rule("R2", 2, [expr_condition("主表.金额 > 100")], {"x": "2"})
    rs = make_rs([r1, r2], ["x"], hit_policy="唯一")
    with pytest.raises(ValueError, match="C001"):
        apply_rules(rs, ns_two_rows, {}, labels=pd.Series(["C001", "C002", "C003"]))


def test_duplicate_rule_ids_rejected(ns_two_rows):
    # masks 按 ID 建字典, 重复 ID 会静默丢掉一条规则; 两个加载器都已拦截, 引擎再兜一层
    rules = [Rule("R", 1, [], {"x": "a"}), Rule("R", 2, [], {"x": "b"})]
    with pytest.raises(ValueError, match="规则 ID 重复: R"):
        apply_rules(make_rs(rules, ["x"]), ns_two_rows, {})


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


def test_run_tests_expected_cleared():
    """期望格写 清空 / 为空 = 期望该列为空值; 留空仍是「不变」, 所以能区分「被清空」和「没动」。"""
    r1 = Rule("R1", 1, [cell_condition("主表.状态", "活跃")], {"备注": "清空"})
    rs = make_rs(
        [r1],
        ["备注"],
        tests=[
            TestCase("清空成功", {"主表.状态": "活跃", "主表.备注": "旧"}, {"备注": "清空"}),
            TestCase("也认 为空", {"主表.状态": "活跃", "主表.备注": "旧"}, {"备注": "为空"}),
            TestCase("留空=不变, 抓到被清空", {"主表.状态": "活跃", "主表.备注": "旧"}, {"备注": None}),
            TestCase("没命中, 期望空则失败", {"主表.状态": "休眠", "主表.备注": "旧"}, {"备注": "清空"}),
        ],
        variables=["主表.状态"],
    )
    outcomes = run_tests(rs)
    assert [o.passed for o in outcomes] == [True, True, False, False]
    assert outcomes[2].problems == ["备注: 期望 旧, 实际 空"]
    assert outcomes[3].problems == ["备注: 期望 空, 实际 旧"]
