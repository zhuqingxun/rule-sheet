"""版式规范: 样例工作簿的颜色 / 列宽 / 冻结 / 文本格, `rule-sheet format` 的幂等、不改语义与报错。"""

import shutil

import pytest
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from rule_sheet import cli
from rule_sheet.loaders import load_rules
from rule_sheet.style import FILLS, HEADER_FILL, ID_KINDS, KV_WIDTH_BOUNDS, WIDTH_BOUNDS, display_width, format_workbook

from .conftest import CAND_EXCEL, CARRY_EXCEL, DIFF_EXCEL, EXCEL_RULES, STATUS_EXCEL, TAGS_EXCEL, YAML_RULES
from .test_loaders import HEADER, RULE, TEST_HEADER, make_workbook

SAMPLES = pytest.mark.parametrize(
    "path",
    [EXCEL_RULES, STATUS_EXCEL, TAGS_EXCEL, CARRY_EXCEL, CAND_EXCEL, DIFF_EXCEL],
    ids=["customer", "status", "tags", "carry", "cand", "diff"],
)


def color(cell):
    return cell.fill.fgColor.rgb[-6:] if cell.fill.fill_type == "solid" else None


def width(ws, col):
    return ws.column_dimensions[get_column_letter(col)].width


def header_row(ws, key):
    return next(r for r in range(1, 6) if ws.cell(r, 1).value == key)


def test_display_width_counts_wide_chars_double():
    assert display_width("abc") == 3 and display_width("中文") == 4 and display_width("a中") == 3
    assert display_width(None) == 0 and display_width(7.0) == 1 and display_width("ab\n中文字") == 6
    assert display_width("（全角）") == 8


@SAMPLES
def test_sample_workbooks_follow_the_spec(path):
    """三个样例都由同一套规范生成: 表头与正文按列类型配色, 列宽在区间内, 冻结到 ID 列之右, = 开头的格是文本。"""
    wb = load_workbook(path)
    for sheet, key in (("决策表", "规则ID"), ("测试", "用例")):
        ws = wb[sheet]
        kinds_row = header_row(ws, key)
        names_row, body = kinds_row + 1, kinds_row + 2
        kinds = [ws.cell(kinds_row, c).value for c in range(1, ws.max_column + 1)]
        for c, kind in enumerate(kinds, start=1):
            if kind not in FILLS:
                continue
            assert color(ws.cell(kinds_row, c)) == FILLS[kind] and color(ws.cell(names_row, c)) == FILLS[kind]
            assert ws.cell(kinds_row, c).font.bold and ws.cell(names_row, c).font.bold
            assert color(ws.cell(body, c)) == (None if kind in ID_KINDS else FILLS[kind]), (sheet, kind)
            lo, hi = WIDTH_BOUNDS[kind]
            assert lo <= width(ws, c) <= hi, (sheet, kind, width(ws, c))
            assert ws.cell(body, c).alignment.horizontal == ("center" if kind in ("输入", *ID_KINDS) else "left")
        frozen = sum(1 for k in kinds if k in ID_KINDS)  # 规则ID + 优先级 / 用例
        assert ws.freeze_panes == f"{get_column_letter(frozen + 1)}{body}"
        for row in ws.iter_rows(min_row=body):
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    assert cell.data_type == "s" and cell.number_format == "@", (sheet, cell.coordinate)
    ws = wb["决策表"]
    assert ws.cell(1, 1).value == "命中策略" and ws.cell(1, 1).font.bold and color(ws.cell(1, 3)) == FILLS["说明"]
    kv_sheets = [s for s in ("输入", "设置", "变量定义", "字段组") if s in wb.sheetnames]
    assert len(kv_sheets) >= 2  # 设置 / 变量定义 / 字段组 都是可选的, 但每个样例至少有一个
    for sheet in kv_sheets:
        ws = wb[sheet]
        header = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
        assert all(color(ws.cell(1, c)) == HEADER_FILL and ws.cell(1, c).font.bold for c in range(1, len(header) + 1))
        note = header.index("说明") + 1
        assert color(ws.cell(2, note)) == FILLS["说明"] and ws.cell(2, 1).font.bold
        for c, name in enumerate(header, start=1):
            lo, hi = KV_WIDTH_BOUNDS[name]
            assert lo <= width(ws, c) <= hi, (sheet, name)
        assert ws.freeze_panes == "A2"


def test_long_cells_wrap_and_rows_grow():
    wb = load_workbook(TAGS_EXCEL)
    ws = wb["决策表"]
    body = header_row(ws, "规则ID") + 2
    note_col = next(c for c in range(1, ws.max_column + 1) if ws.cell(body - 2, c).value == "说明")
    display_col = next(c for c in range(1, ws.max_column + 1) if ws.cell(body - 1, c).value == "显示名")
    assert ws.cell(body, note_col).alignment.wrap_text and ws.row_dimensions[body].height > 15
    g9 = next(r for r in range(body, ws.max_row + 1) if str(ws.cell(r, display_col).value).startswith("="))
    assert ws.cell(g9, display_col).alignment.wrap_text  # G9 的 =如果(...) 超过输出列上限, 换行而不是继续加宽
    assert width(ws, display_col) == WIDTH_BOUNDS["输出"][1]


def _semantics(path):
    rs = load_rules(path)
    return (
        rs.hit_policy,
        rs.settings,
        rs.derived,
        [(r.rule_id, r.priority, r.outputs, r.desc) for r in rs.rules],
        rs.outputs,
        [(t.name, t.inputs, t.expected) for t in rs.tests],
        rs.variables,
        [(s.name, s.primary, s.keys, s.dup_policy, s.path) for s in rs.inputs],
    )


def _styles(path):
    wb = load_workbook(path)
    out = {}
    for ws in wb.worksheets:
        cells = {}
        for row in ws.iter_rows():
            for cell in row:
                cells[cell.coordinate] = (
                    color(cell),
                    cell.font.bold,
                    cell.alignment.horizontal,
                    cell.alignment.wrap_text,
                    cell.number_format,
                    cell.data_type,
                    cell.border.left.style,
                )
        out[ws.title] = (
            ws.freeze_panes,
            {k: v.width for k, v in ws.column_dimensions.items()},
            {k: v.height for k, v in ws.row_dimensions.items() if v.height},
            cells,
        )
    return out


def test_format_is_idempotent_and_keeps_semantics(tmp_path):
    """format 原地改写: 套两次与套一次相同; 加载出来的规则 / 测试 / 变量 / 设置不变 (status-machine 含合并单元格)。"""
    src = tmp_path / "r.xlsx"
    shutil.copy(STATUS_EXCEL, src)
    before = _semantics(src)
    assert format_workbook(src) == src
    once = _styles(src)
    assert format_workbook(src) == src
    assert _styles(src) == once
    assert _semantics(src) == before
    other = format_workbook(src, tmp_path / "sub" / "o.xlsx")
    assert other.exists() and _styles(other) == once


def test_format_applies_spec_to_a_plain_workbook(tmp_path):
    """业务人员从零写的、没有任何样式的工作簿, format 之后和样例一个规范。"""
    p = make_workbook(
        tmp_path / "plain.xlsx",
        [["命中策略", "首次命中"], *RULE],
        [*TEST_HEADER, ["T1", "x", None, "命中"]],
        derived=[("翻倍", "主表.a * 2")],
        settings=[["基准日期", "2026-10-19"]],
    )
    format_workbook(p)
    wb = load_workbook(p)
    ws = wb["决策表"]
    assert color(ws.cell(2, 3)) == FILLS["输入"] and color(ws.cell(4, 3)) == FILLS["输入"]
    assert color(ws.cell(2, 5)) == FILLS["输出"] and color(ws.cell(4, 1)) is None
    assert ws.freeze_panes == "C4" and ws.cell(1, 1).font.bold
    assert wb["测试"].freeze_panes == "B3"
    assert color(wb["输入"].cell(1, 1)) == HEADER_FILL and wb["变量定义"].cell(2, 1).font.bold
    assert WIDTH_BOUNDS["输入"][0] <= width(ws, 3) <= WIDTH_BOUNDS["输入"][1]


def test_format_rejects_bad_structure_and_non_excel(tmp_path):
    with pytest.raises(ValueError, match="format 只处理 Excel 规则工作簿"):
        format_workbook(YAML_RULES)
    bad = make_workbook(tmp_path / "bad.xlsx", [["规则", "输入"], ["编号", "主表.a"]])
    with pytest.raises(ValueError, match="找不到列类型行"):  # 加载器的原话, 结构错就不碰文件
        format_workbook(bad)
    assert bad.stat().st_size > 0


def test_cli_format(tmp_path, capsys):
    src = tmp_path / "r.xlsx"
    shutil.copy(TAGS_EXCEL, src)
    assert cli.main(["format", str(src)]) == 0
    assert "已格式化" in capsys.readouterr().out
    out = tmp_path / "o.xlsx"
    assert cli.main(["format", str(src), "-o", str(out)]) == 0 and out.exists()
    assert cli.main(["format", str(YAML_RULES)]) == 2
    assert "format 只处理" in capsys.readouterr().out
    assert _semantics(src) == _semantics(TAGS_EXCEL)


def test_header_unchanged_by_format(tmp_path):
    """HEADER 常量在 test_loaders 里是列类型 / 列名两行; format 后两行仍被加载器识别为表头。"""
    p = make_workbook(tmp_path / "h.xlsx", [*HEADER, ["R1", 1, "x", "-", "命中"]])
    format_workbook(p)
    assert [r.rule_id for r in load_rules(p).rules] == ["R1"]
