"""执行核心: 命名空间构建、命中策略合并、输出格解析、测试用例运行。"""

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from rule_sheet.conditions import cell_condition, expr_condition
from rule_sheet.engine import (
    _is_date_column,
    add_derived,
    apply_rules,
    build_namespace,
    read_table,
    run_tests,
    split_sheet,
    write_xlsx_safely,
)
from rule_sheet.model import InputSpec, Rule, RuleSet, TestCase

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


def inputs_of(*specs: tuple):
    """(名称, 匹配键列表 | None) ...; 第一个是主表。键写成 "本输入的列=主表的列" 或 "同名列"。"""
    out = []
    for i, (name, keys) in enumerate(specs):
        pairs = [tuple(k.split("=")) if "=" in k else (k, k) for k in (keys or [])]
        out.append(InputSpec(name, primary=(i == 0), keys=pairs))
    return out


def make_rs(rules, outputs, hit_policy="优先级", **kw):
    return RuleSet(
        inputs=kw.pop("inputs", inputs_of(("主表", None), ("本月", ["编号"]))),
        settings=kw.pop("settings", {}),
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
    rs = make_rs([], [], settings={"基准日期": "2026-10-19"})
    ns = build_namespace(rs, {"主表": left, "本月": right})
    assert ns["主表.金额"].tolist() == [1, 2, 3]
    assert ns["本月.金额"].tolist()[0] == 10 and pd.isna(ns["本月.金额"].iloc[1])
    assert ns["本月.存在"].tolist() == [True, False, True]
    assert "本月.编号" not in ns  # 参照表的匹配键列不暴露, 值就是主表的
    assert ns["基准日期"] == pd.Timestamp("2026-10-19")


def test_namespace_three_inputs_with_own_keys_and_same_named_columns():
    """N 个参照表各自按自己的键对齐到主表; 参照表之间同名列 (姓名) 互不干扰; 键名可以不同 (工号 = 主表.拟任人工号)。"""
    base = pd.DataFrame({"岗位": ["P1", "P2", "P3"], "拟任人工号": ["E1", "E2", None], "姓名": ["a", "b", "c"]})
    last = pd.DataFrame({"岗位": ["P1", "P3"], "姓名": ["a0", "c0"], "状态": ["新增", "进展"]})
    roster = pd.DataFrame({"工号": ["E1", "E2"], "姓名": ["张三", "李四"], "部门": ["财务", "人事"]})
    rs = make_rs([], [], inputs=inputs_of(("本期", None), ("上期", ["岗位"]), ("花名册", ["工号=拟任人工号"])))
    ns = build_namespace(rs, {"本期": base, "上期": last, "花名册": roster})
    assert ns["本期.姓名"].tolist() == ["a", "b", "c"]
    assert ns["上期.姓名"].tolist()[0] == "a0" and pd.isna(ns["上期.姓名"].iloc[1])
    assert ns["花名册.姓名"].tolist()[:2] == ["张三", "李四"] and pd.isna(ns["花名册.姓名"].iloc[2])
    assert ns["上期.存在"].tolist() == [True, False, True]
    assert ns["花名册.存在"].tolist() == [True, True, False]  # 主表键为空的行匹配不到
    assert "花名册.工号" not in ns and "上期.岗位" not in ns


def test_namespace_single_input_needs_no_join():
    base = pd.DataFrame({"k": [1, 2], "v": ["a", "b"]})
    ns = build_namespace(make_rs([], [], inputs=inputs_of(("表", None))), {"表": base})
    assert set(ns) == {"表.k", "表.v", "基准日期"}


def test_lookup_rows_with_null_keys_never_match():
    # pandas merge 会把两边的 NaN 键配成一对, 参照表里键为空的行必须先剔掉, 否则 存在 会假阳性
    base = pd.DataFrame({"k": [1, None], "v": ["a", "b"]})
    lookup = pd.DataFrame({"k": [1, None], "w": ["x", "y"]})
    ns = build_namespace(
        make_rs([], [], inputs=inputs_of(("主表", None), ("本月", ["k"]))), {"主表": base, "本月": lookup}
    )
    assert ns["本月.存在"].tolist() == [True, False]
    assert ns["本月.w"].tolist()[0] == "x" and pd.isna(ns["本月.w"].iloc[1])


def test_duplicate_lookup_keys_policy():
    left = pd.DataFrame({"编号": [1]})
    right = pd.DataFrame({"编号": [1, 1], "v": ["first", "second"]})
    notes: list[str] = []
    ns = build_namespace(make_rs([], []), {"主表": left, "本月": right}, notes)
    assert ns["本月.v"].tolist() == ["first"] and notes
    strict = make_rs([], [])
    strict.lookups[0].dup_policy = "报错"
    with pytest.raises(ValueError, match="输入「本月」有 1 行重复键"):
        build_namespace(strict, {"主表": left, "本月": right})


def test_missing_key_column():
    with pytest.raises(ValueError, match="主表「主表」没有列「编号」"):
        build_namespace(make_rs([], []), {"主表": pd.DataFrame({"a": [1]}), "本月": pd.DataFrame({"编号": [1]})})
    with pytest.raises(ValueError, match="输入「本月」没有匹配键列「编号」"):
        build_namespace(make_rs([], []), {"主表": pd.DataFrame({"编号": [1]}), "本月": pd.DataFrame({"b": [1]})})


def test_split_sheet(tmp_path):
    plain = tmp_path / "a.xlsx"
    plain.write_bytes(b"")
    hashed = tmp_path / "报表#1.xlsx"
    hashed.write_bytes(b"")
    assert split_sheet(str(plain)) == (plain, None)
    assert split_sheet(f"{plain}#明细") == (plain, "明细")
    assert split_sheet(str(hashed)) == (hashed, None)  # 文件名本身含 #, 整个字符串是存在的文件就不拆
    assert split_sheet("a.xlsx#s", tmp_path) == (plain, "s")  # 相对路径按 base_dir 解析


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


def test_null_keyword_in_outputs_tests_and_derived(ns_two_rows):
    """关键字 空: 输出格 (空 / =空) 等价于 清空; 期望格写 空 = 期望为空; 输入写 空 = 空值; 派生变量算出 空 = 空列。"""
    r = Rule("R", 1, [cell_condition("本月.存在", "是")], {"a": "空", "b": "=空", "c": "= 空"})
    keep = {k: pd.Series(["keep"] * 3) for k in "abc"}
    res, _ = apply_rules(make_rs([r], ["a", "b", "c"]), ns_two_rows, keep)
    assert res["a"].tolist() == res["b"].tolist() == res["c"].tolist() == [None, "keep", None]

    r1 = Rule("R1", 1, [cell_condition("主表.状态", "空")], {"备注": "没状态"})
    r2 = Rule("R2", 2, [cell_condition("主表.状态", "活跃")], {"备注": "空"})
    rs = make_rs(
        [r1, r2],
        ["备注"],
        tests=[
            TestCase("输入写 空 = 空值", {"主表.状态": "空", "主表.备注": "旧"}, {"备注": "没状态"}),
            TestCase("期望写 空 = 期望为空", {"主表.状态": "活跃", "主表.备注": "旧"}, {"备注": "空"}),
            TestCase("没清就失败", {"主表.状态": "休眠", "主表.备注": "旧"}, {"备注": "空"}),
        ],
        variables=["主表.状态"],
    )
    outcomes = run_tests(rs)
    assert [o.passed for o in outcomes] == [True, True, False]
    assert outcomes[2].problems == ["备注: 期望 空, 实际 旧"]

    ns = dict(ns_two_rows)
    add_derived(make_rs([], [], derived=[("占位", "空")]), ns, ns["主表.金额"].index, lenient=False)
    assert ns["占位"].isna().all()


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


def test_join_policy_concatenates_all_hits_in_priority_order(ns_two_rows):
    """多条拼接 (issue #16, DMN 的 Collect): 每个输出列把所有命中规则的值按优先级顺序用分隔符接起来并替换原值;
    没写该列的规则不参与, 算出空值的规则不贡献, 没命中的行不动; 审计列记全部命中。"""
    r1 = Rule("R1", 1, [cell_condition("主表.状态", "活跃")], {"标签": "高管", "备注": "a"})
    r2 = Rule("R2", 3, [cell_condition("本月.存在", "是")], {"标签": "直管"})
    big = '=如果(主表.金额 > 6000, "大客户", 空)'
    r3 = Rule("R3", 2, [cell_condition("本月.存在", "是")], {"标签": big, "备注": "-"})
    originals = {"标签": pd.Series(["旧"] * 3), "备注": pd.Series(["旧"] * 3)}
    rs = make_rs([r1, r2, r3], ["标签", "备注"], hit_policy="多条拼接")
    res, hit = apply_rules(rs, ns_two_rows, originals)
    assert res["标签"].tolist() == ["高管/大客户/直管", "旧", "直管"]  # 默认分隔符 /; 行 2 的 R3 算出空, 不贡献
    assert res["备注"].tolist() == ["a", "旧", "旧"]
    assert hit.tolist() == ["R1; R3; R2", "", "R3; R2"]
    rs = make_rs([r1, r2, r3], ["标签", "备注"], hit_policy="多条拼接", settings={"拼接分隔符": "; "})
    res, _ = apply_rules(rs, ns_two_rows, originals)
    assert res["标签"].tolist() == ["高管; 大客户; 直管", "旧", "直管"]
    r4 = Rule("R4", 0, [], {"标签": 7.0, "备注": "本月.电话"})  # 数值 / 复制列 一样进拼接, 7.0 写成 7, 空的列值不贡献
    res, _ = apply_rules(make_rs([r1, r4], ["标签", "备注"], hit_policy="多条拼接"), ns_two_rows, originals)
    assert res["标签"].tolist() == ["7/高管", "7", "7"] and res["备注"].tolist() == ["139/a", "138", "旧"]


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


def test_date_columns_written_with_date_format(tmp_path):
    """规则写过之后变成 object 的日期列也按 yyyy-mm-dd 写出; 混着文本的列和全空列走原来的路径 (不当日期列)。"""
    idx = pd.RangeIndex(3)
    df = pd.DataFrame(
        {
            "全日期": pd.Series(
                [pd.Timestamp("2026-10-20"), None, pd.Timestamp("2026-09-20")], index=idx, dtype=object
            ),
            "混合": pd.Series([pd.Timestamp("2026-10-20"), "待定", None], index=idx, dtype=object),
            "全空": pd.Series([None, None, None], index=idx, dtype=object),
            "原生": pd.to_datetime(pd.Series(["2026-10-20", None, "2026-09-20"], index=idx)),
        }
    )
    assert [_is_date_column(df[c]) for c in df.columns] == [True, False, False, True]
    out = write_xlsx_safely(df, tmp_path / "o.xlsx")
    ws = load_workbook(out)["Sheet1"]
    formats = {ws.cell(1, c).value: ws.cell(2, c).number_format for c in range(1, 5)}
    assert formats["全日期"] == "yyyy-mm-dd" and formats["原生"] == "yyyy-mm-dd"
    assert formats["混合"] != "yyyy-mm-dd" and formats["全空"] != "yyyy-mm-dd"
