"""Excel -> YAML 导出的无损性 (逐形态行为等价 + 三个样例端到端), 以及两份规则文件的语义比对。"""

import shutil
from datetime import datetime

import pandas as pd
import pytest

from rule_sheet import cli
from rule_sheet.conditions import to_mask, unary_test
from rule_sheet.convert import CellTranslator, diff_rules, excel_to_yaml, export_workbook
from rule_sheet.engine import run, run_tests
from rule_sheet.expressions import eval_expr
from rule_sheet.loaders import load_rules
from rule_sheet.model import InputSpec, RuleSet

from .conftest import EXCEL_RULES, STATUS_DIR, STATUS_EXCEL, STATUS_YAML, TAGS_DIR, TAGS_EXCEL, TAGS_YAML, YAML_RULES

SAMPLES = pytest.mark.parametrize(
    ("excel", "yaml_file", "data_dir"),
    [
        (EXCEL_RULES, YAML_RULES, EXCEL_RULES.parent),
        (STATUS_EXCEL, STATUS_YAML, STATUS_DIR),
        (TAGS_EXCEL, TAGS_YAML, TAGS_DIR),
    ],
    ids=["customer", "status", "tags"],
)


@pytest.fixture
def translator():
    rs = RuleSet(
        inputs=[InputSpec("客户", primary=True), InputSpec("本月", keys=[("客户编号", "客户编号")])],
        settings={},
        derived=[("拟任人有变化", "客户.电话 != 本月.电话")],
        hit_policy="优先级",
        rules=[],
        outputs=[],
    )
    return CellTranslator(rs)


@pytest.fixture
def ns():
    idx = pd.RangeIndex(5)
    return {
        "客户.状态": pd.Series(["活跃", "休眠", None, "VIP", "是"], index=idx, dtype=object),
        "偏差率": pd.Series([0.08, 0.03, None, 11.0, 0.05], index=idx),
        "本月.存在": pd.Series([True, False, True, False, True], index=idx),
        "客户.电话": pd.Series(["138", None, "139", "007", "x"], index=idx, dtype=object),
        "本月.电话": pd.Series(["139", "138", None, "007", "x"], index=idx, dtype=object),
        "客户.上次活跃": pd.Series(
            pd.to_datetime(["2026-10-09", "2026-10-08", None, "2026-01-01", "2026-10-08"]), index=idx
        ),
        "拟任人有变化": pd.Series([True, False, True, False, True], index=idx),
        "基准日期": pd.Timestamp("2026-10-19"),
    }


@pytest.mark.parametrize(
    ("var", "cell", "expr"),
    [
        ("客户.状态", "活跃", '客户.状态 == "活跃"'),
        ("客户.状态", "!= 休眠", '客户.状态 != "休眠"'),
        ("客户.状态", "<> 休眠", '客户.状态 != "休眠"'),
        ("客户.状态", "= 活跃", '客户.状态 == "活跃"'),
        ("客户.状态", "VIP, 活跃", '属于(客户.状态, "VIP", "活跃")'),
        ("客户.状态", "活跃, 空", '属于(客户.状态, "活跃", 空)'),
        ("客户.状态", '"VIP 小幅偏差, 关注"', '客户.状态 == "VIP 小幅偏差, 关注"'),
        ("客户.状态", "= 空", "客户.状态 == 空"),
        ("客户.状态", "!= 空", "客户.状态 != 空"),
        ("客户.状态", "空", "客户.状态 == 空"),
        ("客户.状态", "为空", "客户.状态 == 空"),
        ("客户.状态", "非空", "客户.状态 != 空"),
        ("客户.状态", "not(活跃)", '客户.状态 != "活跃"'),
        ("客户.状态", "not(VIP, 活跃)", '客户.状态 not in ("VIP", "活跃")'),
        ("客户.状态", "not(空)", "客户.状态 != 空"),
        ("客户.状态", "not(为空)", "客户.状态 != 空"),
        ("客户.状态", "not(非空)", "客户.状态 == 空"),
        ("客户.状态", "not(包含 活)", '客户.状态 != 空 且 非 (包含(客户.状态, "活"))'),
        ("客户.状态", "包含 活, 休", '包含(客户.状态, "活", "休")'),
        ("客户.状态", "开头是(活)", '开头是(客户.状态, "活")'),
        ("客户.状态", "不包含 眠", '不包含(客户.状态, "眠")'),
        ("客户.状态", "是", '客户.状态 == "是"'),  # 文本列上的 是 就是文字
        ("偏差率", "> 0.05", "偏差率 > 0.05"),
        ("偏差率", ">= 0.05", "偏差率 >= 0.05"),
        ("偏差率", "[0.02..0.05]", "0.02 <= 偏差率 <= 0.05"),
        ("偏差率", "(0.02..0.05]", "0.02 < 偏差率 <= 0.05"),
        ("偏差率", "[0.02..0.05)", "0.02 <= 偏差率 < 0.05"),
        ("偏差率", "> 10, 为空", "偏差率 > 10 或 偏差率 == 空"),
        ("偏差率", "not(> 0.05)", "偏差率 != 空 且 非 (偏差率 > 0.05)"),
        ("偏差率", 0.05, "偏差率 == 0.05"),
        ("偏差率", 11, "偏差率 == 11"),
        ("偏差率", "11", "偏差率 == 11"),
        ("本月.存在", "是", "本月.存在 == 是"),
        ("本月.存在", "否", "本月.存在 == 否"),
        ("本月.存在", True, "本月.存在 == 是"),
        ("拟任人有变化", "是", "拟任人有变化 == 是"),
        ("本月.电话", "!= 客户.电话", "本月.电话 != 客户.电话"),
        ("本月.电话", "= 客户.电话", "本月.电话 == 客户.电话"),
        ("本月.电话", "客户.电话", "本月.电话 == 客户.电话"),
        ("客户.电话", "007", "客户.电话 == 7"),  # 和单元格一样: 裸的 007 经 literal 变成 7, 要保文本就加引号
        ("客户.电话", '"007"', '客户.电话 == "007"'),
        ("客户.电话", "v1.2", '客户.电话 == "v1.2"'),
        ("客户.上次活跃", "> 2026-10-08", '客户.上次活跃 > 日期("2026-10-08")'),
        ("客户.上次活跃", "2026-10-08", '客户.上次活跃 == 日期("2026-10-08")'),
        ("客户.上次活跃", datetime(2026, 10, 8), '客户.上次活跃 == 日期("2026-10-08")'),
        ("客户.状态", "-", None),
        ("客户.状态", None, None),
        ("客户.状态", "任意", None),
        ("客户.状态", "活跃, -", None),
    ],
)
def test_cell_translation_matches_unary_test(translator, ns, var, cell, expr):
    """每种单元格形态翻出来的表达式, 在含空值的数据上与 unary_test 的 mask 完全一致 (这是「无损」的定义)。"""
    assert translator.cell(var, cell) == expr
    want = unary_test(cell, ns[var], ns).tolist()
    got = to_mask(eval_expr(expr, ns), ns).tolist() if expr is not None else [True] * 5
    assert got == want, (cell, expr)


def test_names_with_special_chars_get_backticks(translator):
    assert translator.cell("客户.金额(万元)", "> 100") == "`客户.金额(万元)` > 100"
    assert translator.cell("客户.类别-2024", "x") == '`客户.类别-2024` == "x"'
    assert translator.cell("客户.类别-会议", "x") == '客户.类别-会议 == "x"'
    assert translator.cell("客户.岗位/职级", "x") == '客户.岗位/职级 == "x"'
    assert translator.cell("客户.电话", "!= 客户.金额(万元)") == "客户.电话 != `客户.金额(万元)`"
    assert translator.cell("客户.电话", '= "a\'b"') == '客户.电话 == "a\'b"'


@SAMPLES
def test_export_is_lossless_end_to_end(excel, yaml_file, data_dir, tmp_path):
    """导出件自己的用例全过; 导出件跑真实数据与 Excel 结果相同; 与手写 YAML 语义一致 (只有输出路径这种提示)。"""
    out = export_workbook(excel, tmp_path / "exported.yaml")
    text = out.read_text(encoding="utf-8")
    assert text.startswith("# 由 rule-sheet export 从") and "输入:" in text and "规则:" in text
    rs = load_rules(out)
    outcomes = run_tests(rs)
    failed = [(o.name, o.problems, o.error) for o in outcomes if not o.passed]
    assert outcomes and not failed, failed
    a = run(excel, dry_run=True).out
    b = run(out, base_dir=data_dir, dry_run=True).out
    cols = list(load_rules(excel).outputs) + ["命中规则"]
    pd.testing.assert_frame_equal(a[cols], b[cols], check_dtype=False)
    res = diff_rules(excel, out)
    assert res.same and res.cases == len(rs.tests), res.differences
    res = diff_rules(out, yaml_file)  # 手写的 YAML 样例带注释, 不由导出覆盖; 靠 diff 守住两者一致
    assert res.same, res.differences
    assert all("输出" in n or "默认文件" in n for n in res.notes)


@SAMPLES
def test_hand_written_yaml_matches_excel(excel, yaml_file, data_dir):
    """这就是主张的同步机制: Excel 是源, 手写 / 生成的 YAML 用 diff 对账。"""
    res = diff_rules(excel, yaml_file)
    assert res.same, res.differences


def test_diff_reports_each_kind_of_change(tmp_path):
    src = STATUS_YAML.read_text(encoding="utf-8")
    p = tmp_path / "changed.yaml"
    s3_cond = "- 本期.类别-会议 == 上期.类别-会议      # 上期没有"
    changed = (
        src.replace("优先级: 4\n", "优先级: 40\n", 1)
        .replace("状态-会议: 03-按计划进行\n", "状态-会议: 03-按计划\n", 1)
        .replace(s3_cond, s3_cond.replace("==", "!="), 1)
    )
    assert changed != src and s3_cond in src
    p.write_text(changed, encoding="utf-8")
    res = diff_rules(STATUS_EXCEL, p)
    text = "\n".join(res.differences)
    assert "规则 S4 优先级: A 4, B 40" in text
    assert "规则 S4 输出: A {'状态-会议': '03-按计划进行'}, B {'状态-会议': '03-按计划'}" in text
    assert "用例「T3」命中规则: A S3, B" in text  # S3 的条件改了, 结构比不出来, 交叉验证能抓到
    assert any("输出「状态-会议」" in d for d in res.differences)
    p.write_text(src.replace("  - 编号: S5\n", "  - 编号: S6\n", 1), encoding="utf-8")
    res = diff_rules(STATUS_EXCEL, p)
    assert "规则 S5 只在 A 里" in res.differences and "规则 S6 只在 B 里" in res.differences


def test_export_rejects_yaml_and_cli(tmp_path, capsys):
    with pytest.raises(ValueError, match="只支持 Excel"):
        export_workbook(YAML_RULES)
    src = tmp_path / "r.xlsx"
    shutil.copy(TAGS_EXCEL, src)
    assert cli.main(["export", str(src)]) == 0
    out = capsys.readouterr().out
    assert "已导出" in out and "(6 条规则, 8 条用例)" in out and (tmp_path / "r.yaml").exists()
    assert cli.main(["diff", str(src), str(tmp_path / "r.yaml")]) == 0
    assert "语义一致" in capsys.readouterr().out
    assert cli.main(["diff", str(src), str(STATUS_YAML)]) == 1
    assert "发现" in capsys.readouterr().out
    assert cli.main(["export", str(YAML_RULES)]) == 2


def test_exported_yaml_round_trips_through_loader(tmp_path):
    """导出的 YAML 再被加载时, 规则 ID / 优先级 / 输出 / 用例与 Excel 一致 (条件靠上面的端到端测试)。"""
    out = export_workbook(STATUS_EXCEL, tmp_path / "s.yaml")
    a, b = load_rules(STATUS_EXCEL), load_rules(out)
    assert [(r.rule_id, r.priority, r.desc) for r in a.rules] == [(r.rule_id, r.priority, r.desc) for r in b.rules]
    assert [t.name for t in a.tests] == [t.name for t in b.tests]
    assert b.hit_policy == "首次命中" and b.outputs == ["状态-会议"]
    text = out.read_text(encoding="utf-8")
    assert "任一满足:" in text  # S3 的两行续行
    assert excel_to_yaml(STATUS_EXCEL) == excel_to_yaml(STATUS_EXCEL)  # 确定性
