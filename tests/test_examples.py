"""端到端: 自带样例在两种前端下结果一致, 且与设计表一致。"""

import pandas as pd
import pytest

from rule_sheet import cli
from rule_sheet.engine import run, run_tests
from rule_sheet.loaders import load_rules

from .conftest import EXCEL_RULES, STATUS_EXCEL, STATUS_YAML, YAML_RULES

FRONTENDS = pytest.mark.parametrize("rules_file", [EXCEL_RULES, YAML_RULES], ids=["excel", "yaml"])
STATUS_FRONTENDS = pytest.mark.parametrize("rules_file", [STATUS_EXCEL, STATUS_YAML], ids=["excel", "yaml"])

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
    assert hits["C008"] in ("R4; R5a", "R4; R5")
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
    cols = ["客户编号", "状态", "电话", "复核标记", "备注", "提醒"]
    pd.testing.assert_frame_equal(a[cols], b[cols], check_dtype=False)


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

STATUS_HITS = {"P001": "S1", "P002": "S2", "P003": "S3", "P004": "S4", "P005": "S1", "P006": "S4"}


@STATUS_FRONTENDS
def test_status_embedded_cases_all_pass(rules_file):
    rs = load_rules(rules_file)
    assert rs.hit_policy == "首次命中"
    outcomes = run_tests(rs)
    assert len(outcomes) == 6
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
    assert row.loc["P001", "变更明细"] == "状态-会议: 01-新增 -> 空"
    assert len(res.changed) == 4
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
    assert not t2.passed and t2.hit == "S2; S4" and t2.problems == ["状态-会议: 期望 空, 实际 01-新增"]


def test_cli_run_and_test(tmp_path, capsys):
    assert cli.main(["test", str(YAML_RULES)]) == 0
    assert "8/8 通过" in capsys.readouterr().out
    assert cli.main([str(EXCEL_RULES), "--output", str(tmp_path / "o.xlsx")]) == 0  # 省略 run 子命令
    out = capsys.readouterr().out
    assert "5 行发生变更" in out and "已写出" in out
    assert cli.main(["run", str(tmp_path / "missing.yaml")]) == 2
