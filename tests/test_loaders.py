"""两种前端的加载: 结构、变量扫描、条件组合。"""

from datetime import date, datetime

import pandas as pd
import pytest

from rule_sheet.conditions import build_condition
from rule_sheet.loaders import load_rules, load_yaml_rules

from .conftest import EXCEL_RULES, YAML_RULES


def test_excel_structure():
    rs = load_rules(EXCEL_RULES)
    assert [r.rule_id for r in rs.rules] == ["R1", "R2", "R3", "R4", "R5"]  # R5 占两行 (或 分支), 仍是一条
    assert rs.outputs == ["状态", "电话", "复核标记", "备注", "提醒"]
    assert rs.hit_policy == "优先级"
    assert len(rs.tests) == 8 and rs.tests[0].name == "T1"
    # 表头输入列在前, 之后是派生变量表达式里引用到的列 (用例没给时置空, 派生变量才能算出来)
    assert rs.variables == [
        "客户.状态", "客户.等级", "偏差率", "本月.存在", "距上次活跃天数", "本月.电话", "客户.电话",
        "客户.上次活跃", "客户.金额", "本月.金额",
    ]  # fmt: skip
    assert [d for d, _ in rs.derived] == ["偏差率", "距上次活跃天数"]
    assert isinstance(rs.settings["基准日期"], (date, datetime))
    assert [(s.name, s.primary, s.keys, s.path) for s in rs.inputs] == [
        ("客户", True, [], "客户主表.xlsx"),
        ("本月", False, [("客户编号", "客户编号")], "本月流水.xlsx"),
    ]


def test_yaml_structure_and_variable_scan():
    rs = load_rules(YAML_RULES)
    assert [r.rule_id for r in rs.rules] == ["R1", "R2", "R3", "R4", "R5"]
    assert rs.outputs == ["复核标记", "备注", "状态", "电话", "提醒"]
    assert len(rs.tests) == 8
    # 嵌套在 任一满足 里的变量也要被扫到, 测试用例没给时才会被置空而不是报未定义
    assert {"本月.电话", "客户.电话", "本月.存在", "偏差率", "距上次活跃天数"} <= set(rs.variables)


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


def test_yaml_backtick_name_end_to_end(tmp_path):
    """反引号名字从 YAML 文件走完整路径: 以反引号开头的表达式整行要加引号 (YAML 的保留字符),
    扫描到的名字在用例没给时置空、给了就参与求值, 两种情况都不报错。"""
    p = tmp_path / "r.yaml"
    p.write_text(
        "匹配: {左表别名: 本期, 右表别名: 上期}\n"
        "规则:\n"
        '  - 编号: R1\n    当:\n      - "`本期.金额(万元)` > 100"\n    则:\n      标记: 大额\n'
        "测试:\n"
        "  - 用例: 没给, 视为空, 不命中\n    输入: {}\n    期望: {}\n"
        "  - 用例: 给了就命中\n    输入: {本期.金额(万元): 150}\n    期望: {标记: 大额}\n",
        encoding="utf-8",
    )
    rs = load_yaml_rules(p)
    assert rs.variables == ["本期.金额(万元)"]
    from rule_sheet.engine import run_tests

    assert [(o.passed, o.hit, o.error) for o in run_tests(rs)] == [(True, "", ""), (True, "R1", "")]


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
    with pytest.raises(ValueError, match="未知的命中策略「first」.*优先级 / 唯一 / 首次命中 / 多条拼接"):
        load_rules(make_workbook(tmp_path / "bad.xlsx", [["命中策略", "first"], ["规则ID"], ["规则ID"]]))


def test_join_policy_validation(tmp_path):
    """多条拼接 下输出格写 清空 没有意义, 报错; 没用这个策略却设了「拼接分隔符」多半是策略名写错, 也报错。"""
    cleared = [*HEADER, ["R1", 1, "x", "-", "清空"]]
    with pytest.raises(ValueError, match="规则 R1 输出「标记」= 清空: 命中策略 多条拼接 只拼接有值的输出"):
        load_rules(make_workbook(tmp_path / "a.xlsx", [["命中策略", "多条拼接"], *cleared]))
    with pytest.raises(ValueError, match="设置了「拼接分隔符」, 但命中策略是「优先级」"):
        load_rules(make_workbook(tmp_path / "b.xlsx", RULE, settings=[["拼接分隔符", ";"]]))
    ok = make_workbook(tmp_path / "c.xlsx", [["命中策略", "多条拼接"], *RULE], settings=[["拼接分隔符", ";"]])
    rs = load_rules(ok)
    assert rs.hit_policy == "多条拼接" and rs.settings["拼接分隔符"] == ";"
    p = tmp_path / "r.yaml"
    p.write_text("命中策略: 多条拼接\n规则:\n  - 编号: A\n    则: {y: =空}\n", encoding="utf-8")
    with pytest.raises(ValueError, match="规则 A 输出「y」= =空: 命中策略 多条拼接"):
        load_yaml_rules(p)
    p.write_text("设置: {拼接分隔符: '/'}\n规则:\n  - 则: {y: 1}\n", encoding="utf-8")
    with pytest.raises(ValueError, match="设置了「拼接分隔符」"):
        load_yaml_rules(p)


INPUTS = [["名称", "匹配", "文件"], ["主表", "主表", "a.xlsx"], ["本月", "k", "b.xlsx"]]


def make_workbook(
    path, decision, tests=None, derived=None, merges=(), inputs=None, settings=None, groups=None, groups_header=None
):
    """最小规则工作簿: decision / tests / inputs 是整张 sheet 的行 (含表头), groups 不含表头, merges 是要合并的区域。"""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "输入"
    for row in inputs or INPUTS:
        ws.append(row)
    if settings:
        ws = wb.create_sheet("设置")
        ws.append(["设置项", "值"])
        for row in settings:
            ws.append(row)
    if derived:
        ws = wb.create_sheet("变量定义")
        ws.append(["变量", "定义"])
        for row in derived:
            ws.append(list(row))
    if groups is not None:
        ws = wb.create_sheet("字段组")
        ws.append(list(groups_header or ["名称", "列名"]))
        for row in groups:
            ws.append(list(row))
    ws = wb.create_sheet("决策表")
    for row in decision:
        ws.append(row)
    for rng in merges:
        ws.merge_cells(rng)
    if tests is not None:
        ws = wb.create_sheet("测试")
        for row in tests:
            ws.append(row)
    wb.save(path)
    return path


HEADER = [["规则ID", "优先级", "输入", "输入", "输出"], ["规则ID", "优先级", "主表.a", "主表.b", "标记"]]
TEST_HEADER = [["用例", "输入", "输入", "期望"], ["用例", "主表.a", "主表.b", "标记"]]


def test_excel_header_found_by_content_not_row_number(tmp_path):
    """issue #6: 表头写在第 1、2 行 (没有元信息行) 也能识别。
    以前第 2 行被当列类型、第一条用例被当列名, 剩下的用例空输入空期望假通过。"""
    from rule_sheet.engine import run_tests

    p = make_workbook(
        tmp_path / "r.xlsx",
        [*HEADER, ["R1", 1, "x", "-", "命中"]],
        [*TEST_HEADER, ["T1", "x", None, "命中"], ["T2", "q", None, "命中"]],
    )
    rs = load_rules(p)
    assert rs.hit_policy == "优先级" and [r.rule_id for r in rs.rules] == ["R1"]
    assert [(t.name, t.inputs["主表.a"]) for t in rs.tests] == [("T1", "x"), ("T2", "q")]
    assert [o.passed for o in run_tests(rs)] == [True, False]  # T2 不该命中, 现在能抓到


def test_excel_policy_label_anywhere_above_header(tmp_path):
    rule = ["R1", 1, "x", "-", "命中"]
    p = make_workbook(tmp_path / "r.xlsx", [["说明", "随便写"], [None, "命中策略:", "首次命中"], *HEADER, rule])
    assert load_rules(p).hit_policy == "首次命中"
    p = make_workbook(tmp_path / "r2.xlsx", [[None, "唯一"], *HEADER, rule])  # 没有标签, 单独一格策略名也认
    assert load_rules(p).hit_policy == "唯一"


@pytest.mark.parametrize(
    ("decision", "msg"),
    [
        ([["规则", "输入"], ["编号", "主表.a"]], "找不到列类型行"),
        ([["规则ID", "条件", "输出"], ["规则ID", "主表.a", "标记"]], "列类型「条件」不认识"),
        ([["规则ID", "输入", "输出"], ["规则ID", None, "标记"]], "B2 格没有写列名"),
        ([["规则ID", "规则ID"], ["规则ID", "又一个"]], "恰有一列类型为「规则ID」"),
        ([["规则ID", "输入"]], "下面没有列名行"),
    ],
)
def test_excel_header_errors(tmp_path, decision, msg):
    with pytest.raises(ValueError, match=msg):
        load_rules(make_workbook(tmp_path / "r.xlsx", decision))


def test_excel_test_sheet_guards(tmp_path):
    """期望列拼错 = 整列被忽略, 输入列拼错 = 填了不起作用, 空 sheet = 什么都没验证: 三种假通过都在加载时报错。"""
    rule = [*HEADER, ["R1", 1, "x", "-", "命中"]]
    kinds = ["用例", "输入", "期望"]
    with pytest.raises(ValueError, match="「测试」sheet: 期望列「标计」不是输出列"):
        load_rules(make_workbook(tmp_path / "a.xlsx", rule, [kinds, ["用例", "主表.a", "标计"], ["T1", "x", "命中"]]))
    with pytest.raises(ValueError, match="输入列「主表.c」不是规则引用到的变量"):
        load_rules(make_workbook(tmp_path / "b.xlsx", rule, [kinds, ["用例", "主表.c", "标记"], ["T1", "x", "命中"]]))
    with pytest.raises(ValueError, match="没有任何用例"):
        load_rules(make_workbook(tmp_path / "c.xlsx", rule, TEST_HEADER))
    # 输出列的原值 (主表.标记)、派生变量及其引用的列、内置变量 都允许作为输入; 派生变量留空 = 按表达式推导
    p = make_workbook(
        tmp_path / "d.xlsx",
        [["规则ID", "输入", "输出"], ["规则ID", "翻倍", "标记"], ["R1", "> 10", "大"]],
        [
            ["用例", "输入", "输入", "输入", "输入", "期望"],
            ["用例", "主表.标记", "主表.n", "翻倍", "本月.存在", "标记"],
            ["T1", "旧", 6, None, "是", "大"],
            ["T2", "旧", 6, 1, "是", None],  # 给了派生变量的值就不再推导: 翻倍 = 1 不命中
        ],
        derived=[("翻倍", "主表.n * 2")],
    )
    rs = load_rules(p)
    assert rs.variables == ["翻倍", "主表.n"]
    from rule_sheet.engine import run_tests

    assert [o.passed for o in run_tests(rs)] == [True, True]


def test_excel_scans_variables_referenced_inside_cells(tmp_path):
    """输入格右侧 (!= 上期.b / [0..主表.上限]) 和输出格 (=主表.c * 2 / 本月.d) 引用的变量不一定是表头输入列,
    它们也要进 variables: 测试用例才能给值, 没给的才会置空而不是被当成字面文本比较。"""
    from rule_sheet.engine import run_tests

    p = make_workbook(
        tmp_path / "r.xlsx",
        [
            ["规则ID", "输入", "输入", "输出", "输出"],
            ["规则ID", "主表.a", "主表.n", "标记", "电话"],
            ["R1", "!= 本月.b", "[0..主表.上限]", "=主表.c * 2", "本月.d"],
        ],
        [
            ["用例", "输入", "输入", "输入", "输入", "输入", "输入", "期望", "期望"],
            ["用例", "主表.a", "本月.b", "主表.n", "主表.上限", "主表.c", "本月.d", "标记", "电话"],
            ["T1", "x", "y", 5, 10, 21, "139", 42, "139"],
            ["T2", "x", "x", 5, 10, 21, "139", None, None],  # a == b 不命中
            ["T3", "x", None, 5, 10, 21, "139", None, None],  # 本月.b 没给 = 空, != 不成立, 不能拿字面 "本月.b" 去比
        ],
    )
    rs = load_rules(p)
    assert rs.variables == ["主表.a", "主表.n", "主表.c", "主表.上限", "本月.b", "本月.d"]
    assert [(o.passed, o.hit) for o in run_tests(rs)] == [(True, "R1"), (True, ""), (True, "")]


def test_excel_continuation_rows_are_or_branches(tmp_path):
    """issue #8: 相邻的同 ID 行是同一条规则的「或」分支 (每行各格 AND), 规则 ID 可以合并单元格, 输出以首行为准。"""
    from rule_sheet.engine import apply_rules

    rows = [["R1", 1, "x", "-", "命中"], [None, None, "-", "y", None], ["R2", 2, "-", "-", "兜底"]]
    p = make_workbook(
        tmp_path / "r.xlsx",
        [["命中策略", "首次命中"], *HEADER, *rows],
        merges=("A4:A5",),  # R1 的规则 ID 合并两行, 第二行读出来是 None, 展开后就是 R1
    )
    rs = load_rules(p)
    assert [r.rule_id for r in rs.rules] == ["R1", "R2"] and rs.rules[0].priority == 1
    ns = {"主表.a": pd.Series(["x", "q", "q"], dtype=object), "主表.b": pd.Series([None, "y", "n"], dtype=object)}
    res, hit = apply_rules(rs, ns, {})
    assert res["标记"].tolist() == ["命中", "命中", "兜底"] and hit.tolist() == ["R1", "R1", "R2"]


@pytest.mark.parametrize(
    ("rows", "msg"),
    [
        ([["R1", 1, "x", "-", "命中"], ["R1", None, "-", "y", "别的"]], "输出「标记」= 别的, 与首行"),
        ([["R1", 1, "x", "-", "命中"], ["R1", 9, "-", "y", None]], "优先级与首行"),
        ([["R1", 1, "x", "-", "命中"], ["R2", 2, "-", "-", "兜底"], ["R1", 3, "-", "y", "命中"]], "重复出现.*相邻"),
    ],
)
def test_excel_continuation_row_errors(tmp_path, rows, msg):
    with pytest.raises(ValueError, match=msg):
        load_rules(make_workbook(tmp_path / "r.xlsx", [*HEADER, *rows]))


def test_yaml_duplicate_id_and_test_guards(tmp_path):
    p = tmp_path / "r.yaml"
    p.write_text("规则:\n  - 编号: A\n    则: {y: 1}\n  - 编号: A\n    则: {y: 2}\n", encoding="utf-8")
    with pytest.raises(ValueError, match="规则编号「A」重复"):
        load_yaml_rules(p)
    base = "规则:\n  - 编号: A\n    当: [x > 1]\n    则: {y: 1}\n测试:\n"
    p.write_text(base, encoding="utf-8")
    with pytest.raises(ValueError, match="没有任何用例"):
        load_yaml_rules(p)
    p.write_text(base + "  - 用例: T1\n    输入: {x: 2}\n    期望: {z: 1}\n", encoding="utf-8")
    with pytest.raises(ValueError, match="测试用例「T1」: 期望列「z」不是输出列"):
        load_yaml_rules(p)
    p.write_text(base + "  - 用例: T1\n    输入: {本月.x: 2}\n    期望: {y: 1}\n", encoding="utf-8")
    with pytest.raises(ValueError, match="测试用例「T1」: 输入列「本月.x」不是规则引用到的变量"):
        load_yaml_rules(p)


# ---------------------------------------------------------------------------
# 字段组: 组名当输出列, 一格表达整组复制 / 整组清空; 加载时展开成逐列, 引擎不知道有字段组
# ---------------------------------------------------------------------------

GROUP_DECISION = [
    ["规则ID", "输入", "输出", "输出"],
    ["规则ID", "主表.a", "带入", "y"],
    ["R1", "x", "本月", "-"],  # 整组从 本月 复制
    ["R2", "q", "本月", "=主表.a"],  # 显式写的 y 优先于字段组里的 y
    ["R3", "z", "清空", "-"],  # 整组清空
    ["R4", "w", "-", 1],  # 字段组格留空 = 整组不改
]


def test_excel_field_groups_expand_to_columns(tmp_path):
    from rule_sheet.engine import apply_rules

    rs = load_rules(make_workbook(tmp_path / "r.xlsx", GROUP_DECISION, groups=[["带入", "x, y, z"]]))
    assert rs.field_groups == {"带入": ["x", "y", "z"]} and rs.outputs == ["x", "y", "z"]
    assert [r.outputs for r in rs.rules] == [
        {"x": "本月.x", "y": "本月.y", "z": "本月.z"},
        {"x": "本月.x", "y": "=主表.a", "z": "本月.z"},
        {"x": "清空", "y": "清空", "z": "清空"},
        {"y": 1},
    ]
    assert rs.variables == ["主表.a", "本月.x", "本月.y", "本月.z"]  # 展开出来的 本月.x 也进变量清单, 用例能给值
    idx = pd.RangeIndex(4)
    ns = {
        "主表.a": pd.Series(["x", "q", "z", "w"], index=idx, dtype=object),
        "本月.x": pd.Series([1, 2, 3, 4], index=idx),
        "本月.y": pd.Series([5, 6, 7, 8], index=idx),
        "本月.z": pd.Series([9, 10, 11, 12], index=idx),
    }
    originals = {c: pd.Series([0] * 4, index=idx, dtype=object) for c in rs.outputs}
    res, hits = apply_rules(rs, ns, originals)
    assert hits.tolist() == ["R1", "R2", "R3", "R4"]
    assert res["x"].tolist() == [1, 2, None, 0] and res["z"].tolist() == [9, 10, None, 0]
    assert res["y"].tolist() == [5, "q", None, 1]


def test_yaml_field_groups(tmp_path):
    p = tmp_path / "r.yaml"
    p.write_text(
        "输入:\n  主表: {匹配: 主表}\n  本月: {匹配: k}\n"
        "字段组:\n  带入: [x, y]\n  另一组: a, b\n"  # 列表或逗号分隔的字符串都认
        "规则:\n  - 编号: R1\n    则: {带入: 本月, y: 2}\n  - 编号: R2\n    则: {另一组: 清空, 带入: '-'}\n",
        encoding="utf-8",
    )
    rs = load_yaml_rules(p)
    assert rs.field_groups == {"带入": ["x", "y"], "另一组": ["a", "b"]}
    assert rs.rules[0].outputs == {"x": "本月.x", "y": 2} and rs.rules[1].outputs == {"a": "清空", "b": "清空"}
    assert rs.outputs == ["x", "y", "a", "b"] and "本月.x" in rs.variables
    p.write_text("字段组: {带入: 5}\n规则: []\n", encoding="utf-8")
    with pytest.raises(ValueError, match="「带入」的值要是列名清单"):
        load_yaml_rules(p)


BAD_GROUP_CELL = [*GROUP_DECISION[:2], ["R1", "x", "上月", "-"]]  # 上月 不是输入名
JOIN_GROUP = [["命中策略", "多条拼接"], *GROUP_DECISION]  # R3 整组清空, 在 多条拼接 下没有意义


@pytest.mark.parametrize(
    ("groups", "header", "decision", "msg"),
    [
        (
            [["带入", "x, y"]],
            None,
            BAD_GROUP_CELL,
            "规则 R1 \\(第 3 行\\) 输出「带入」= 上月: 字段组格只能写一个输入名",
        ),
        ([["带入", ""]], None, GROUP_DECISION, "字段组「带入」没有列"),
        ([["带入", "x, x"]], None, GROUP_DECISION, "字段组「带入」里的列「x」重复"),
        ([["带入", "x"], ["带入", "y"]], None, GROUP_DECISION, "字段组「带入」重复"),
        ([["本月", "x"]], None, GROUP_DECISION, "字段组「本月」与输入名重复"),
        ([["带入", "x, 带入"]], None, GROUP_DECISION, "字段组名「带入」同时是某个字段组里的列名"),
        ([["带入", "x"]], ["名称", "字段"], GROUP_DECISION, "「字段组」sheet 第 1 行缺少「列名」列"),
        ([["带入", "x, y"]], None, JOIN_GROUP, "规则 R3 输出「x」= 清空: 命中策略 多条拼接"),
    ],
)
def test_field_group_errors(tmp_path, groups, header, decision, msg):
    with pytest.raises(ValueError, match=msg):
        load_rules(make_workbook(tmp_path / "r.xlsx", decision, groups=groups, groups_header=header))


def test_field_group_copy_is_fine_under_join_policy(tmp_path):
    """多条拼接 只拦 清空; 字段组写输入名展开成的 上期.x 逐列参与拼接, 和手写六列一样。"""
    decision = [["命中策略", "多条拼接"], *GROUP_DECISION[:2], ["R1", "x", "本月", "-"], ["R2", "x", "-", "附加"]]
    rs = load_rules(
        make_workbook(tmp_path / "r.xlsx", decision, groups=[["带入", "x, y"]], settings=[["拼接分隔符", ";"]])
    )
    assert rs.hit_policy == "多条拼接" and [r.outputs for r in rs.rules] == [
        {"x": "本月.x", "y": "本月.y"},
        {"y": "附加"},
    ]
    from rule_sheet.engine import apply_rules

    idx = pd.RangeIndex(1)
    ns = {
        "主表.a": pd.Series(["x"], index=idx, dtype=object),
        "本月.x": pd.Series([1], index=idx),
        "本月.y": pd.Series(["b"], index=idx),
    }
    res, hits = apply_rules(rs, ns, {c: pd.Series([None], index=idx, dtype=object) for c in rs.outputs})
    assert hits.tolist() == ["R1; R2"] and res["x"].tolist() == ["1"] and res["y"].tolist() == ["b;附加"]


# ---------------------------------------------------------------------------
# 输入声明: 规则文件是形参 (名字 + 怎么对齐到主表), 文件只是可选的默认绑定
# ---------------------------------------------------------------------------

RULE = [*HEADER, ["R1", 1, "x", "-", "命中"]]


def test_excel_inputs_sheet(tmp_path):
    rows = [
        ["说明", "名称", "文件", "匹配", "重复键"],  # 列顺序随意, 多余的列忽略
        ["被更新的表", "主表", "a.xlsx", "主表", None],
        ["参照", "本月", None, "k", "报错"],
        ["另一张", "花名册", "c.xlsx#人员", "工号 = 主表.拟任人工号, 年月", None],
    ]
    rs = load_rules(make_workbook(tmp_path / "r.xlsx", RULE, inputs=rows, settings=[["基准日期", "2026-10-19"]]))
    assert [(s.name, s.primary, s.keys, s.dup_policy, s.path) for s in rs.inputs] == [
        ("主表", True, [], "取第一条", "a.xlsx"),
        ("本月", False, [("k", "k")], "报错", None),
        ("花名册", False, [("工号", "拟任人工号"), ("年月", "年月")], "取第一条", "c.xlsx#人员"),
    ]
    assert rs.names == ["主表", "本月", "花名册"] and rs.settings == {"基准日期": "2026-10-19"}


def test_single_input_may_omit_primary_word(tmp_path):
    rs = load_rules(make_workbook(tmp_path / "r.xlsx", RULE, inputs=[["名称", "匹配"], ["主表", None]]))
    assert rs.primary.name == "主表" and rs.lookups == []


@pytest.mark.parametrize(
    ("inputs", "msg"),
    [
        ([["名称", "匹配"], ["主表", None], ["本月", "k"]], "必须恰有一个输入的「匹配」写 主表 .* 现在有 0 个"),
        ([["名称", "匹配"], ["主表", "主表"], ["本月", "主表"]], "现在有 2 个 \\(主表, 本月\\)"),
        ([["名称", "匹配"], ["主表", "主表"], ["本月", None]], "输入「本月」没有写「匹配」"),
        ([["名称", "匹配"], ["主表", "主表"], ["主表", "k"]], "输入名「主表」重复"),
        ([["名称", "匹配", "重复键"], ["主表", "主表", None], ["本月", "k", "去重"]], "「重复键」只认 取第一条 / 报错"),
        ([["名称", "匹配"], ["主表", "主表"], ["本月", "k"], ["花名册", "工号 = 本月.工号"]], "只能对齐到主表「主表」"),
        ([["名字", "匹配"], ["主表", "主表"]], "第 1 行缺少「名称」列"),
        ([["名称", "匹配"]], "没有声明任何输入"),
    ],
)
def test_excel_inputs_errors(tmp_path, inputs, msg):
    with pytest.raises(ValueError, match=msg):
        load_rules(make_workbook(tmp_path / "r.xlsx", RULE, inputs=inputs))


def test_settings_sheet_rejects_unknown_keys(tmp_path):
    with pytest.raises(ValueError, match="设置项「輸出」不认识"):
        load_rules(make_workbook(tmp_path / "r.xlsx", RULE, settings=[["輸出", "x.xlsx"]]))


def test_legacy_match_sheet_is_converted(tmp_path):
    """0.4 以前的「匹配设置」sheet (左表 / 右表 / 别名 / 匹配键 / 右表重复键) 折算成输入声明, 老文件不用改。"""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "匹配设置"
    for row in (
        ["设置项", "值"],
        ["左表", "a.xlsx"],
        ["右表", "b.xlsx"],
        ["左表别名", "本期"],
        ["右表别名", "上期"],
        ["匹配键", "岗位编码, 年月"],
        ["右表重复键", "报错"],
        ["基准日期", "2026-10-19"],
        ["输出", "out.xlsx"],
    ):
        ws.append(row)
    ws = wb.create_sheet("决策表")
    for row in [["规则ID", "输入", "输出"], ["规则ID", "本期.a", "标记"], ["R1", "x", "y"]]:
        ws.append(row)
    wb.save(tmp_path / "old.xlsx")
    rs = load_rules(tmp_path / "old.xlsx")
    assert [(s.name, s.primary, s.keys, s.dup_policy, s.path) for s in rs.inputs] == [
        ("本期", True, [], "取第一条", "a.xlsx"),
        ("上期", False, [("岗位编码", "岗位编码"), ("年月", "年月")], "报错", "b.xlsx"),
    ]
    assert rs.settings == {"基准日期": "2026-10-19", "输出": "out.xlsx"}


def test_yaml_inputs_and_settings(tmp_path):
    p = tmp_path / "r.yaml"
    p.write_text(
        "输入:\n"
        "  本期: {文件: a.xlsx, 匹配: 主表}\n"
        "  上期: {文件: b.xlsx, 匹配: 岗位编码, 重复键: 报错}\n"
        "  花名册: {匹配: 工号 = 本期.拟任人工号}\n"
        "设置: {输出: out.xlsx}\n"
        "规则:\n  - 当: [本期.x > 1]\n    则: {y: 2}\n",
        encoding="utf-8",
    )
    rs = load_yaml_rules(p)
    assert [(s.name, s.primary, s.keys, s.dup_policy, s.path) for s in rs.inputs] == [
        ("本期", True, [], "取第一条", "a.xlsx"),
        ("上期", False, [("岗位编码", "岗位编码")], "报错", "b.xlsx"),
        ("花名册", False, [("工号", "拟任人工号")], "取第一条", None),
    ]
    assert rs.settings == {"输出": "out.xlsx"} and rs.variables == ["本期.x"]
    p.write_text("输入:\n  本期: a.xlsx\n规则: []\n", encoding="utf-8")  # 单输入: 只写文件名即可
    assert load_yaml_rules(p).primary.name == "本期"
    p.write_text("输入: {本期: {文件: a.xlsx, 匹配: 主表}}\n匹配: {左表: a.xlsx}\n规则: []\n", encoding="utf-8")
    with pytest.raises(ValueError, match="不能同时出现"):
        load_yaml_rules(p)
    p.write_text("输入: {本期: {档案: a.xlsx}}\n规则: []\n", encoding="utf-8")
    with pytest.raises(ValueError, match="「档案」不认识"):
        load_yaml_rules(p)


def test_missing_file_and_missing_sheet(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_rules(tmp_path / "nope.xlsx")
    from openpyxl import Workbook

    wb = Workbook()
    wb.active.title = "决策表"
    wb.save(tmp_path / "bad.xlsx")
    with pytest.raises(ValueError, match="缺少「输入」sheet"):
        load_rules(tmp_path / "bad.xlsx")
    wb = Workbook()
    wb.active.title = "随便"
    wb.save(tmp_path / "bad2.xlsx")
    with pytest.raises(ValueError, match="缺少「决策表」sheet"):
        load_rules(tmp_path / "bad2.xlsx")
