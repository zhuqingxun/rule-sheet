"""端到端: 自带样例在两种前端下结果一致, 且与设计表一致。"""

import pandas as pd
import pytest
from openpyxl import load_workbook

from rule_sheet import cli
from rule_sheet.engine import run, run_tests
from rule_sheet.loaders import load_rules

from .conftest import (
    CAND_EXCEL,
    CAND_YAML,
    CARRY_DIR,
    CARRY_EXCEL,
    CARRY_YAML,
    EXCEL_RULES,
    STATUS_DIR,
    STATUS_EXCEL,
    STATUS_YAML,
    TAGS_EXCEL,
    TAGS_YAML,
    YAML_RULES,
)

FRONTENDS = pytest.mark.parametrize("rules_file", [EXCEL_RULES, YAML_RULES], ids=["excel", "yaml"])
STATUS_FRONTENDS = pytest.mark.parametrize("rules_file", [STATUS_EXCEL, STATUS_YAML], ids=["excel", "yaml"])
TAGS_FRONTENDS = pytest.mark.parametrize("rules_file", [TAGS_EXCEL, TAGS_YAML], ids=["excel", "yaml"])
CARRY_FRONTENDS = pytest.mark.parametrize("rules_file", [CARRY_EXCEL, CARRY_YAML], ids=["excel", "yaml"])
CAND_FRONTENDS = pytest.mark.parametrize("rules_file", [CAND_EXCEL, CAND_YAML], ids=["excel", "yaml"])

EXPECTED_HITS = {
    "C001": "R1",
    "C002": "",
    "C003": "R2",
    "C004": "",
    "C005": "R3",
    "C006": "R1; R3",
    "C007": "",
}


@FRONTENDS
def test_embedded_cases_all_pass(rules_file):
    outcomes = run_tests(load_rules(rules_file))
    assert len(outcomes) == 8
    assert all(o.passed for o in outcomes), [(o.name, o.problems, o.error) for o in outcomes if not o.passed]


@FRONTENDS
def test_real_data_changes(rules_file, tmp_path):
    res = run(rules_file, output=tmp_path / "out.xlsx")
    assert not res.stopped_by_tests and res.out is not None
    hits = dict(zip(res.out["客户编号"], res.out["命中规则"], strict=True))
    for cid, want in EXPECTED_HITS.items():
        assert hits[cid] == want, cid
    assert hits["C008"] == "R4; R5"  # Excel 里 R5 是两行同 ID, 审计列按一条规则记
    assert len(res.changed) == 5
    assert res.written is not None and res.written.exists()
    row = res.out.set_index("客户编号")
    assert row.loc["C003", "状态"] == "休眠"
    assert row.loc["C006", "电话"] == "13900000006"
    assert row.loc["C006", "备注"] == "金额偏差超 5%"
    assert row.loc["C008", "提醒"] == "电话有缺失"
    assert pd.isna(row.loc["C002", "复核标记"])


def test_both_frontends_agree(tmp_path):
    a = run(EXCEL_RULES, output=tmp_path / "a.xlsx").out
    b = run(YAML_RULES, output=tmp_path / "b.xlsx").out
    # 不比 变更明细: 它的字段顺序跟着输出列顺序走, Excel 按表头、YAML 按 则: 首次出现, 两边不同
    cols = ["客户编号", "状态", "电话", "复核标记", "备注", "提醒", "命中规则"]
    pd.testing.assert_frame_equal(a[cols], b[cols], check_dtype=False)


def test_inputs_bound_from_cli_override_rule_file_defaults(tmp_path, capsys):
    """规则文件声明输入 (形参), 命令行 -i 名称=路径 绑定文件 (实参, 相对当前目录); 没给的用规则文件里的默认文件。"""
    data = STATUS_DIR
    res = run(
        STATUS_YAML, inputs={"本期": str(data / "本期岗位.xlsx"), "上期": str(data / "上期岗位.xlsx")}, dry_run=True
    )
    assert res.inputs["本期"] == (data / "本期岗位.xlsx", 9) and res.inputs["上期"][1] == 6
    assert len(res.changed) == 7
    res = run(
        STATUS_YAML, inputs={"上期": str(data / "本期岗位.xlsx#本期岗位")}, dry_run=True
    )  # 带 sheet 名; 本期 用默认文件
    assert res.inputs["上期"][1] == 9 and res.inputs["本期"][1] == 9
    with pytest.raises(ValueError, match="输入「花名册」不是规则声明的输入"):
        run(STATUS_YAML, inputs={"花名册": "x.xlsx"}, dry_run=True)
    with pytest.raises(ValueError, match="没有 sheet「明细」, 有: 本期岗位"):
        run(STATUS_YAML, inputs={"本期": str(data / "本期岗位.xlsx#明细")}, dry_run=True)
    with pytest.raises(FileNotFoundError, match="数据表不存在"):
        run(STATUS_YAML, inputs={"本期": str(tmp_path / "没有.xlsx")}, dry_run=True)
    # 命令行同样的事; 输出路径 -o 覆盖, 默认文件名按主表文件派生
    out = tmp_path / "r.xlsx"
    code = cli.main(["run", str(STATUS_EXCEL), "-i", f"上期={data / '上期岗位.xlsx'}", "-o", str(out)])
    text = capsys.readouterr().out
    assert code == 0 and out.exists() and "输入: 本期 (主表) = " in text and "上期 = " in text
    assert cli.main(["run", str(STATUS_EXCEL), "-i", "上期", "--dry-run"]) == 2
    assert "-i 要写成 名称=路径" in capsys.readouterr().out


def test_unbound_input_is_an_error(tmp_path):
    p = tmp_path / "r.yaml"
    p.write_text("输入:\n  本期: {匹配: 主表}\n  上期: {匹配: k}\n规则:\n  - 则: {y: 1}\n", encoding="utf-8")
    with pytest.raises(ValueError, match="输入「本期」「上期」没有绑定文件"):
        run(p, dry_run=True)


def test_dry_run_writes_nothing(tmp_path):
    res = run(EXCEL_RULES, dry_run=True, output=tmp_path / "never.xlsx")
    assert res.written is None and not (tmp_path / "never.xlsx").exists()


def test_written_file_has_audit_columns(tmp_path):
    res = run(YAML_RULES, output=tmp_path / "out.xlsx")
    back = pd.read_excel(res.written)
    assert list(back.columns)[-2:] == ["命中规则", "变更明细"]
    assert len(back) == 8


# ---------------------------------------------------------------------------
# 第二个样例 status-machine: 字段名含 - / 输出格 清空 / 命中策略 首次命中
# ---------------------------------------------------------------------------

STATUS_HITS = {
    "P001": "S1",
    "P002": "S2",
    "P003": "S3",  # 拟任人变了 (S3 第一个分支)
    "P004": "S5",
    "P005": "S1",
    "P006": "S5",
    "P007": "S3",  # 流程节点变了 (S3 第二个分支: Excel 的续行 / YAML 的 任一满足)
    "P008": "S4",
    "P009": "S3",  # 上期拟任人为空、本期有: 变量定义里显式写的那一支
}


@STATUS_FRONTENDS
def test_status_embedded_cases_all_pass(rules_file):
    rs = load_rules(rules_file)
    assert rs.hit_policy == "首次命中"
    assert [r.rule_id for r in rs.rules] == ["S1", "S2", "S3", "S4", "S5"]  # Excel 里 S3 占两行, 仍是一条规则
    outcomes = run_tests(rs)
    assert len(outcomes) == 10
    assert all(o.passed for o in outcomes), [(o.name, o.problems, o.error) for o in outcomes if not o.passed]


@STATUS_FRONTENDS
def test_status_real_data_first_match(rules_file, tmp_path):
    res = run(rules_file, output=tmp_path / "out.xlsx")
    assert not res.stopped_by_tests and res.out is not None
    assert dict(zip(res.out["岗位编码"], res.out["命中规则"], strict=True)) == STATUS_HITS  # 每行只记赢的那条
    row = res.out.set_index("岗位编码")
    assert pd.isna(row.loc["P001", "状态-会议"])  # S1 清空
    assert pd.isna(row.loc["P002", "状态-会议"])  # S2 拦截: 优先级 策略下会被兜底 S4 写成 01-新增
    assert row.loc["P003", "状态-会议"] == "02-进展更新"
    assert row.loc["P004", "状态-会议"] == "01-新增"
    assert row.loc["P006", "状态-会议"] == "01-新增"
    assert row.loc["P007", "状态-会议"] == "02-进展更新"
    assert row.loc["P008", "状态-会议"] == "03-按计划进行"
    assert row.loc["P009", "状态-会议"] == "02-进展更新"
    assert row.loc["P001", "变更明细"] == "状态-会议: 01-新增 -> 空"
    assert len(res.changed) == 7
    back = pd.read_excel(res.written)
    assert pd.isna(back.set_index("岗位编码").loc["P001", "状态-会议"])  # 写出的 Excel 里是真空白


def test_status_frontends_agree(tmp_path):
    a = run(STATUS_EXCEL, output=tmp_path / "a.xlsx").out
    b = run(STATUS_YAML, output=tmp_path / "b.xlsx").out
    cols = ["岗位编码", "状态-会议", "命中规则", "变更明细"]
    pd.testing.assert_frame_equal(a[cols], b[cols], check_dtype=False)


def test_status_needs_first_match_not_priority():
    """同一组规则换成 优先级 策略, S2 不占任何列, 用例 T2 (酝酿中应不动) 就会被兜底 S4 改掉。"""
    rs = load_rules(STATUS_YAML)
    rs.hit_policy = "优先级"
    outcomes = {o.name: o for o in run_tests(rs)}
    t2 = next(o for name, o in outcomes.items() if name.startswith("T2"))
    assert not t2.passed and t2.hit == "S2; S5" and t2.problems == ["状态-会议: 期望 空, 实际 01-新增"]


# ---------------------------------------------------------------------------
# 第三个样例 position-tags: 多条拼接 / 文本关键字 / 清单当参照表 / 数字() / N天 / 如果() / 拼接()
# ---------------------------------------------------------------------------

TAGS_EXPECTED = {  # 岗位编码 -> (命中规则, 标签, 显示名)
    "J001": ("G1; G5; G9", "高管/新推荐", "1-集团总部/集团CEO"),
    "J002": ("G9", "普通", "3-华东代表处/华东代表处-行政助理"),  # 名称含 助理, G2 不命中; 标签保持原值
    "J003": ("G2; G3; G4; G9", "直管/关键岗位/高职级", "3-华东代表处/华东代表处-代表"),  # 原值 待定 被整体替换
    "J004": ("G1; G2; G4; G5; G9", "高管/直管/高职级/新推荐", "2-终端BG/终端BG-产品总裁"),  # 职级格是数字 24
    "J005": ("G5; G9", "新推荐", "法务经理"),  # 一层组织为空: 如果() 的假分支
    "J006": ("G3; G4; G9", "关键岗位/高职级", "4-研发体系/研发-SPDT主管"),  # P23 -> 23; 推荐超过一年
    "J007": ("G5; G9", "新推荐", "5-职能平台/财务部-会计"),  # 职级为空, 数字() 取不到
    "J008": ("G5; G9", "新推荐", "6-供应链/供应链-采购经理"),  # 正好 365 天
}


@TAGS_FRONTENDS
def test_tags_embedded_cases_all_pass(rules_file):
    rs = load_rules(rules_file)
    assert rs.hit_policy == "多条拼接" and rs.settings["拼接分隔符"] == "/"
    assert [r.rule_id for r in rs.rules] == ["G1", "G2", "G3", "G4", "G5", "G9"]
    assert {"岗位.职级", "岗位.推荐时间", "关键岗位清单.存在"} <= set(rs.variables)  # 派生变量引用的列 + 清单的 存在
    outcomes = run_tests(rs)
    assert len(outcomes) == 8
    assert all(o.passed for o in outcomes), [(o.name, o.problems, o.error) for o in outcomes if not o.passed]


@TAGS_FRONTENDS
def test_tags_real_data_collects_all_hits(rules_file, tmp_path):
    res = run(rules_file, output=tmp_path / "out.xlsx")
    assert not res.stopped_by_tests and res.out is not None
    row = res.out.set_index("岗位编码")
    for code, (hits, tag, display) in TAGS_EXPECTED.items():
        assert (row.loc[code, "命中规则"], row.loc[code, "标签"], row.loc[code, "显示名"]) == (hits, tag, display), code
    assert len(res.changed) == 8  # G9 无条件写 显示名, 每一行都变
    assert row.loc["J002", "变更明细"] == "显示名: 空 -> 3-华东代表处/华东代表处-行政助理"
    back = pd.read_excel(res.written).set_index("岗位编码")
    assert back.loc["J004", "标签"] == "高管/直管/高职级/新推荐"


def test_tags_frontends_agree(tmp_path):
    a = run(TAGS_EXCEL, output=tmp_path / "a.xlsx").out
    b = run(TAGS_YAML, output=tmp_path / "b.xlsx").out
    cols = ["岗位编码", "标签", "显示名", "命中规则", "变更明细"]
    pd.testing.assert_frame_equal(a[cols], b[cols], check_dtype=False)


# ---------------------------------------------------------------------------
# 第四个样例 carry-forward: 字段组 (整组从上期复制 / 显式列优先 / 整组同步把空值也带过来)
# ---------------------------------------------------------------------------

CARRY_OUTPUTS = ["例会时间", "类别-会议", "状态-会议", "例会结论", "备注", "待跟进事宜"]
CARRY_HITS = {"K001": "R1", "K002": "R2", "K003": "R3", "K004": "", "K005": "R1", "K006": "", "K007": "R1"}


@CARRY_FRONTENDS
def test_carry_field_groups_expand_at_load(rules_file):
    rs = load_rules(rules_file)
    assert rs.field_groups == {"会议记录": CARRY_OUTPUTS}
    assert rs.outputs == CARRY_OUTPUTS  # 字段组在原位置展开; 例会时间 / 状态-会议 既在组里又是显式列, 只出现一次
    r1, r2, r3 = rs.rules
    copied = {c: f"上期.{c}" for c in CARRY_OUTPUTS}
    assert r1.outputs == {**copied, "例会时间": "=基准日期"}  # 显式写的列优先于字段组里的同名列
    assert r2.outputs == {"例会时间": "=基准日期", "状态-会议": "01-新增"}  # 字段组格留空 = 整组不改
    assert r3.outputs == {**copied, "状态-会议": "06-本批未冒泡"}
    assert "会议记录" not in r1.outputs and {f"上期.{c}" for c in CARRY_OUTPUTS} <= set(rs.variables)
    outcomes = run_tests(rs)
    assert len(outcomes) == 6
    assert all(o.passed for o in outcomes), [(o.name, o.problems, o.error) for o in outcomes if not o.passed]


@CARRY_FRONTENDS
def test_carry_real_data_copies_whole_group(rules_file, tmp_path):
    res = run(rules_file, output=tmp_path / "out.xlsx")
    assert not res.stopped_by_tests and res.out is not None
    assert dict(zip(res.out["岗位编码"], res.out["命中规则"], strict=True)) == CARRY_HITS
    row = res.out.set_index("岗位编码")
    assert row.loc["K001", "例会时间"] == pd.Timestamp("2026-10-20") and row.loc["K001", "待跟进事宜"] == "确认任命时间"
    assert row.loc["K003", "例会时间"] == pd.Timestamp("2026-09-20") and row.loc["K003", "状态-会议"] == "06-本批未冒泡"
    assert (
        pd.isna(row.loc["K005", "备注"]) and row.loc["K005", "状态-会议"] == "03-按计划进行"
    )  # 整组同步: 上期为空也带过来
    assert pd.isna(row.loc["K004", "状态-会议"]) and pd.isna(row.loc["K006", "状态-会议"])
    assert row.loc["K005", "变更明细"] == (
        "例会时间: 2026-09-20 -> 2026-10-20; 状态-会议: 02-进展更新 -> 03-按计划进行; "
        "例会结论: 候选人面谈 -> 面谈通过; 备注: 临时备注 -> 空; 待跟进事宜: 空 -> 走审批"
    )
    assert len(res.changed) == 5
    back = pd.read_excel(res.written).set_index("岗位编码")
    assert back.loc["K001", "例会时间"] == pd.Timestamp("2026-10-20")
    ws = load_workbook(res.written)["Sheet1"]
    col = next(c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value == "例会时间")
    assert ws.cell(2, col).number_format == "yyyy-mm-dd"  # 规则写过之后变成 object 的日期列, 写出仍是日期格


def test_carry_frontends_agree(tmp_path):
    a = run(CARRY_EXCEL, output=tmp_path / "a.xlsx").out
    b = run(CARRY_YAML, output=tmp_path / "b.xlsx").out
    cols = ["岗位编码", *CARRY_OUTPUTS, "命中规则", "变更明细"]
    pd.testing.assert_frame_equal(a[cols], b[cols], check_dtype=False)


def test_field_group_named_like_a_real_column_is_an_error(tmp_path):
    """组名在决策表里占一个输出列的位置; 主表真有同名列的话那一列永远写不到, 读到数据时报错而不是静默。"""
    src = CARRY_YAML.read_text(encoding="utf-8")
    p = tmp_path / "r.yaml"
    renamed = src.replace("会议记录: [例会时间", "岗位名称: [例会时间").replace("会议记录: 上期", "岗位名称: 上期")
    assert renamed != src
    p.write_text(renamed, encoding="utf-8")
    rs = load_rules(p)
    assert "岗位名称" in rs.field_groups  # 加载时没有数据, 查不出撞名
    with pytest.raises(ValueError, match="字段组「岗位名称」与主表「本期」的列同名"):
        run(p, base_dir=CARRY_DIR, dry_run=True)


# ---------------------------------------------------------------------------
# 第五个样例 candidate-rollup: 重复键 拼接 (排序 / 空值占位 / 带序号) + 带输入前缀的派生变量
# ---------------------------------------------------------------------------

CAND_OUTPUTS = ["拟任人工号", "拟任人姓名", "拟任人现岗", "推荐时间", "冻结期间", "冻结依据"]
CAND_EXPECTED = {  # 岗位编码 -> (命中规则, 拟任人工号, 拟任人现岗, 推荐时间, 冻结期间, 冻结依据)
    "P1": (
        "P1; F1",
        "E2;E1;E3",  # 推荐时间 降序, 没有推荐时间的排末尾
        "-经理-23级;财务部-会计-22级;人事部-主管-21级",  # 聚合前逐行算的 推荐.现岗; 现部门为空出 -经理-23级
        "2026-10-01;2026-09-01;NA",
        "1.2026-09-01~2026-09-30;2.2026-05-01~2026-05-31",  # 带序号
        "1.审计;2.纪律",
    ),
    "P2": ("P1", "E4;E7", "市场部-经理-;销售部-主管-22级", "NA;NA", None, None),  # 整组全空仍占位, 顺序按原文件
    "P3": ("F1", None, None, None, "1.2026-08-10~", "1.休假"),  # 冻结结束为空: 拼接() 部分空
    "P4": ("P1", "E5;E6", "NA;研发部-工程师-20级", "2026-07-01;2026-06-01", None, None),  # 三列全空 -> 现岗 空 -> NA
}


@CAND_FRONTENDS
def test_cand_concat_inputs_load(rules_file):
    rs = load_rules(rules_file)
    rec, frz = rs.lookups
    assert (rec.dup_policy, rec.sort, rec.sep, rec.placeholder, rec.numbered) == (
        "拼接",
        [("推荐时间", False)],
        ";",
        "NA",
        False,
    )
    assert (frz.keys, frz.dup_policy, frz.sort, frz.numbered) == (
        [("工号", "现任工号")],
        "拼接",
        [("冻结开始", False)],
        True,
    )
    assert [k for k, _ in rs.input_derived] == ["推荐.现岗", "冻结.期间"] and rs.plain_derived == []
    assert {"推荐.现岗", "冻结.期间", "推荐.现部门", "冻结.冻结结束"} <= set(rs.variables)
    outcomes = run_tests(rs)
    assert len(outcomes) == 4
    assert all(o.passed for o in outcomes), [(o.name, o.problems, o.error) for o in outcomes if not o.passed]


@CAND_FRONTENDS
def test_cand_real_data_aggregates_in_matching_layer(rules_file, tmp_path):
    res = run(rules_file, output=tmp_path / "out.xlsx")
    assert not res.stopped_by_tests and res.out is not None
    assert any("输入「推荐」有 4 行重复键, 按「拼接」接成一行" in n for n in res.notes)
    row = res.out.set_index("岗位编码")
    for code, (hits, emp, pos, when, period, basis) in CAND_EXPECTED.items():
        got = tuple(
            None if pd.isna(v) else v for v in row.loc[code, ["命中规则", *CAND_OUTPUTS[0:1], *CAND_OUTPUTS[2:]]]
        )
        assert got == (hits, emp, pos, when, period, basis), code
    assert row.loc["P1", "拟任人姓名"] == "李四;张三;王五"
    assert len(res.changed) == 4
    back = pd.read_excel(res.written).set_index("岗位编码")
    assert back.loc["P2", "推荐时间"] == "NA;NA" and pd.isna(back.loc["P3", "拟任人工号"])


def test_cand_frontends_agree(tmp_path):
    a = run(CAND_EXCEL, output=tmp_path / "a.xlsx").out
    b = run(CAND_YAML, output=tmp_path / "b.xlsx").out
    cols = ["岗位编码", *CAND_OUTPUTS, "命中规则", "变更明细"]
    pd.testing.assert_frame_equal(a[cols], b[cols], check_dtype=False)


def test_cli_run_and_test(tmp_path, capsys):
    assert cli.main(["test", str(YAML_RULES)]) == 0
    assert "8/8 通过" in capsys.readouterr().out
    no_tests = tmp_path / "no_tests.yaml"
    no_tests.write_text("规则:\n  - 则: {y: 1}\n", encoding="utf-8")
    assert cli.main(["test", str(no_tests)]) == 1  # 没有用例 = 什么都没验证, 不能算通过
    assert "没有测试用例" in capsys.readouterr().out
    no_tests.write_text(
        "输入:\n  客户: {文件: 客户主表.xlsx, 匹配: 主表}\n  本月: {文件: 本月流水.xlsx, 匹配: 客户编号}\n"
        "规则:\n  - 则: {提醒: 看一眼}\n",
        encoding="utf-8",
    )
    assert cli.main(["run", str(no_tests), "--base-dir", str(YAML_RULES.parent), "--dry-run"]) == 0
    assert "提示: 规则文件没有内嵌测试用例" in capsys.readouterr().out  # run 不拦, 但要说一声
    assert cli.main([str(EXCEL_RULES), "--output", str(tmp_path / "o.xlsx")]) == 0  # 省略 run 子命令
    out = capsys.readouterr().out
    assert "5 行发生变更" in out and "已写出" in out
    assert cli.main(["run", str(tmp_path / "missing.yaml")]) == 2
