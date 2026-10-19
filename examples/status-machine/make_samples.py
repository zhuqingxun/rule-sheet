# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl>=3.1"]
# ///
"""生成「岗位上会状态」样例: 两份数据 + Excel 决策表, 全部写到本脚本所在目录。

这个样例展示三件 customer-sync 里没有的事:
    - 字段名含 `-` (类别-会议 / 状态-会议), 直接来自业务方表头
    - 输出格 `清空`: 命中后把目标列写成空值 (区别于 `-` 的不改)
    - 命中策略 `首次命中`: 每行只听优先级数字最小的那条, 命中即停, 后面的规则不再看这一行;
      S2 没有任何输出, 只是把「酝酿中」的岗位拦下来不让兜底规则 S4 改它

运行 (任选其一):
    uv run --script examples/status-machine/make_samples.py
    uv run python examples/status-machine/make_samples.py

产物:
    本期岗位.xlsx      被更新的左表
    上期岗位.xlsx      参照用的右表
    规则_决策表.xlsx   三张 sheet: 匹配设置 / 决策表 / 测试 (没有派生变量, 省掉 变量定义)
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

ROOT = Path(__file__).resolve().parent

FILL_BY_KIND = {
    "输入": "DDEBF7",   # 浅蓝: 业务人员改的条件格
    "输出": "E2EFDA",   # 浅绿: 命中后写回的值
    "期望": "E2EFDA",
    "说明": "FFF2CC",   # 浅黄: 纯注释
    "规则ID": "EDEDED",
    "优先级": "EDEDED",
    "用例": "EDEDED",
}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD = Font(bold=True)


def put(ws: Worksheet, row: int, col: int, value: Any, *, bold: bool = False, fill: str | None = None) -> None:
    cell = ws.cell(row=row, column=col, value=value)
    if isinstance(value, str) and value.startswith("="):
        cell.data_type = "s"  # 以 = 开头的 rule-sheet 表达式按文本写, 否则 openpyxl / Excel 会当成公式
        cell.number_format = "@"
    cell.border = BORDER
    cell.alignment = Alignment(vertical="center")
    if bold:
        cell.font = BOLD
    if fill:
        cell.fill = PatternFill("solid", fgColor=fill)


def set_widths(ws: Worksheet, widths: list[int]) -> None:
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def write_data_sheet(ws: Worksheet, header: list[str], rows: list[list[Any]], widths: list[int]) -> None:
    for c, name in enumerate(header, start=1):
        put(ws, 1, c, name, bold=True, fill="D9D9D9")
    for r, row in enumerate(rows, start=2):
        for c, val in enumerate(row, start=1):
            put(ws, r, c, val)
    ws.freeze_panes = "A2"
    set_widths(ws, widths)


# ---------------------------------------------------------------------------
# 样例数据: 每一行对应一种情形, 便于对照决策表看命中结果
# ---------------------------------------------------------------------------
PLAN = "近2个月酝酿及上会计划"
NOTICE = "变动通报"

CURRENT_HEADER = ["岗位编码", "岗位名称", "类别-会议", "状态-会议"]
CURRENT_ROWS = [
    # 编号    名称            类别-会议   状态-会议       上期情况            预期
    ["P001", "财务部-部长",   None,      "01-新增"],    # 上期 变动通报       -> S1 清空状态
    ["P002", "人事部-副部长", PLAN,      None],         # 上期无             -> S2 拦截不动 (优先级 策略下会被 S4 写成 01-新增)
    ["P003", "研发中心-主任", NOTICE,    "01-新增"],    # 上期 变动通报 (相同) -> S3 02-进展更新
    ["P004", "市场部-经理",   NOTICE,    None],         # 上期无             -> S4 01-新增
    ["P005", "法务部-主管",   None,      None],         # 上期无             -> S1 清空空值, 无变更
    ["P006", "采购部-经理",   NOTICE,    "02-进展更新"],  # 上期 酝酿 (不同)    -> S4 01-新增 (类别变了, 重新计为新增)
]

PREVIOUS_HEADER = ["岗位编码", "类别-会议", "状态-会议"]
PREVIOUS_ROWS = [
    ["P001", NOTICE, "01-新增"],
    ["P003", NOTICE, "01-新增"],
    ["P006", PLAN,   None],
]


def build_data() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "本期岗位"
    write_data_sheet(ws, CURRENT_HEADER, CURRENT_ROWS, widths=[10, 16, 24, 14])
    wb.save(ROOT / "本期岗位.xlsx")

    wb = Workbook()
    ws = wb.active
    ws.title = "上期岗位"
    write_data_sheet(ws, PREVIOUS_HEADER, PREVIOUS_ROWS, widths=[10, 24, 14])
    wb.save(ROOT / "上期岗位.xlsx")


# ---------------------------------------------------------------------------
# 规则工作簿
# ---------------------------------------------------------------------------
def write_kv_sheet(ws: Worksheet, title: tuple[str, str, str], rows: list[tuple[Any, Any, Any]],
                   widths: list[int]) -> None:
    for c, name in enumerate(title, start=1):
        put(ws, 1, c, name, bold=True, fill="D9D9D9")
    for r, (k, v, note) in enumerate(rows, start=2):
        put(ws, r, 1, k, bold=True)
        put(ws, r, 2, v)
        put(ws, r, 3, note, fill=FILL_BY_KIND["说明"])
    set_widths(ws, widths)


def write_rule_table(ws: Worksheet, top_note: list[Any], kinds: list[str], names: list[str],
                     rows: list[list[Any]], widths: list[int]) -> None:
    """三行表头: 第 1 行元信息, 第 2 行列类型(输入/输出), 第 3 行变量名; 第 4 行起每行一条规则。"""
    for c, val in enumerate(top_note, start=1):
        put(ws, 1, c, val, bold=(c == 1), fill=FILL_BY_KIND["说明"] if c == 3 else None)
    for c, (kind, name) in enumerate(zip(kinds, names, strict=True), start=1):
        fill = FILL_BY_KIND.get(kind)
        put(ws, 2, c, kind, bold=True, fill=fill)
        put(ws, 3, c, name, bold=True, fill=fill)
    for r, row in enumerate(rows, start=4):
        for c, val in enumerate(row, start=1):
            kind = kinds[c - 1]
            fill = FILL_BY_KIND.get(kind) if kind in ("输入", "输出", "期望") else None
            put(ws, r, c, val, fill=fill)
            horizontal = "center" if kind in ("输入", "规则ID", "优先级", "用例") else "left"
            ws.cell(row=r, column=c).alignment = Alignment(horizontal=horizontal, vertical="center")
    ws.freeze_panes = "C4"
    set_widths(ws, widths)


def build_rules() -> None:
    wb = Workbook()

    # ---- sheet 1: 匹配设置 --------------------------------------------------
    ws = wb.active
    ws.title = "匹配设置"
    write_kv_sheet(ws, ("设置项", "值", "说明"), [
        ("左表", "本期岗位.xlsx", "被更新的表, 相对路径以本规则文件所在目录为基准"),
        ("右表", "上期岗位.xlsx", "参照表, 只读"),
        ("左表别名", "本期", "规则里引用左表字段的前缀, 如 本期.类别-会议"),
        ("右表别名", "上期", "规则里引用右表字段的前缀, 如 上期.类别-会议"),
        ("匹配键", "岗位编码", "两表按此列对行"),
        ("输出", "output/本期岗位_更新后.xlsx", "结果文件, 原表不动"),
    ], widths=[14, 32, 52])

    # ---- sheet 2: 决策表 ----------------------------------------------------
    ws = wb.create_sheet("决策表")
    kinds = ["规则ID", "优先级", "输入", "输入", "输出", "说明"]
    names = ["规则ID", "优先级", "本期.类别-会议", "上期.类别-会议", "状态-会议", "说明"]
    rules = [
        ["S1", 1, "为空", "-",              "清空",       "本期不再上会, 清掉状态 (清空 = 写成空值; - 才是不改)"],
        ["S2", 2, PLAN,   "-",              "-",          "酝酿中的岗位由人工维护, 不处理。首次命中 下命中即停, S3 / S4 不再看这一行 (优先级 策略下会被 S4 改成 01-新增)"],
        ["S3", 3, "-",    "= 本期.类别-会议", "02-进展更新", "类别与上期相同, 记为进展更新"],
        ["S4", 4, "-",    "-",              "01-新增",    "兜底: 前面都没命中的行 (有类别、不是酝酿、与上期不同) 都是新增"],
    ]
    write_rule_table(
        ws,
        top_note=["命中策略", "首次命中", "每行只听优先级数字最小的那条命中规则, 命中即停; 同优先级按行序。输出格 清空 = 写成空值, - = 不改"],
        kinds=kinds, names=names, rows=rules,
        widths=[8, 8, 22, 18, 14, 60],
    )

    # ---- sheet 3: 测试 ------------------------------------------------------
    ws = wb.create_sheet("测试")
    kinds = ["用例", "输入", "输入", "输入", "期望", "说明"]
    names = ["用例", "本期.类别-会议", "本期.状态-会议", "上期.类别-会议", "状态-会议", "说明"]
    tests = [
        ["T1", None,   "01-新增",    NOTICE, "清空",       "类别为空, 清掉状态。期望格写 清空 = 期望该列为空 (留空是期望不变)"],
        ["T2", PLAN,   None,         None,   None,         "S2 拦截, 状态保持空; 优先级 策略下会被 S4 改成 01-新增, 这条用例会失败"],
        ["T3", NOTICE, "01-新增",    NOTICE, "02-进展更新", "类别与上期相同 -> 进展更新"],
        ["T4", NOTICE, None,         None,   "01-新增",    "上期没有这个岗位 -> 兜底新增"],
        ["T5", None,   None,         None,   None,         "清空本来就空的格, 无变化"],
        ["T6", NOTICE, "02-进展更新", PLAN,   "01-新增",    "类别变了 -> 重新计为新增"],
    ]
    write_rule_table(
        ws,
        top_note=["说明", None, "每行一个用例: 填输入变量的值, 期望列留空表示「该列不变」, 写 清空 表示「该列应为空」"],
        kinds=kinds, names=names, rows=tests,
        widths=[6, 22, 14, 22, 14, 60],
    )

    wb.save(ROOT / "规则_决策表.xlsx")


if __name__ == "__main__":
    build_data()
    build_rules()
    print(f"已生成到 {ROOT}: 本期岗位.xlsx  上期岗位.xlsx  规则_决策表.xlsx")
