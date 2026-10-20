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
        "主表.状态", "主表.等级", "偏差率", "本月.存在", "距上次活跃天数", "本月.电话", "主表.电话",
        "主表.上次活跃", "主表.金额", "本月.金额",
    ]  # fmt: skip
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
    from openpyxl import Workbook

    wb = Workbook()
    wb.active.title = "匹配设置"
    ws = wb.create_sheet("决策表")
    ws.append(["命中策略", "first"])
    ws.append(["规则ID"])
    ws.append(["规则ID"])
    wb.save(tmp_path / "bad.xlsx")
    with pytest.raises(ValueError, match="未知的命中策略「first」"):
        load_rules(tmp_path / "bad.xlsx")


def make_workbook(path, decision, tests=None, derived=None, merges=()):
    """最小规则工作簿: decision / tests 是整张 sheet 的行 (含表头), merges 是决策表里要合并的区域。"""
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "匹配设置"
    for row in (["设置项", "值"], ["左表", "a.xlsx"], ["右表", "b.xlsx"], ["匹配键", "k"]):
        ws.append(row)
    if derived:
        ws = wb.create_sheet("变量定义")
        ws.append(["变量", "定义"])
        for row in derived:
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


def test_missing_file_and_missing_sheet(tmp_path):
    with pytest.raises(FileNotFoundError):
        load_rules(tmp_path / "nope.xlsx")
    from openpyxl import Workbook

    wb = Workbook()
    wb.active.title = "随便"
    wb.save(tmp_path / "bad.xlsx")
    with pytest.raises(ValueError, match="匹配设置"):
        load_rules(tmp_path / "bad.xlsx")
