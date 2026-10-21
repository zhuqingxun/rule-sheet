"""规则工作簿的版式规范: 颜色、列宽、换行、冻结, 一处定义, 样例生成脚本与 `rule-sheet format` 共用。

规范 (使用指南 3.6 有同样的表):
- 列类型配色: 输入 浅蓝 / 输出 与 期望 浅绿 / 说明 浅黄 / 规则ID 优先级 用例 浅灰 (只在表头); 键值 sheet 表头 灰
- 表头加粗; 表区域细灰边框; 垂直居中; 输入 与 ID 类列水平居中, 输出 / 说明 左对齐
- 列宽按内容算 (中文算 2 个字符宽), 每类列有上下限; 超过上限的格换行, 行高按行数估算后显式写入
- 冻结: 决策表 / 测试 冻结到列名行之下、最左边连续 ID 类列之右; 其它 sheet 冻结首行
- 以 = 开头的格按文本存 (否则 Excel 当成公式, 显示 #NAME?); 日期格 yyyy-mm-dd
表头在哪一行由 loaders 的同一套逻辑判断, 所以 format 与加载器对「哪里是表头」永远一个答案。
"""

from __future__ import annotations

import math
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .loaders import POLICY_LABEL, RULE_KINDS, TEST_KINDS, _parse_header, _rows, load_rules

# ---------------------------------------------------------------------------
# 规范常量
# ---------------------------------------------------------------------------

FILLS = {  # 列类型 -> 底色 (表头与正文同色)
    "输入": "DDEBF7",  # 浅蓝: 业务人员改的条件格
    "输出": "E2EFDA",  # 浅绿: 命中后写回的值
    "期望": "E2EFDA",
    "说明": "FFF2CC",  # 浅黄: 纯注释, 不参与解析
    "规则ID": "EDEDED",  # 浅灰: 只在表头
    "优先级": "EDEDED",
    "用例": "EDEDED",
}
HEADER_FILL = "D9D9D9"  # 键值 sheet 与数据表的表头
ID_KINDS = ("规则ID", "优先级", "用例")
CENTERED_KINDS = ("输入", *ID_KINDS)
NOTE = "说明"
# 列类型 -> (宽度下限, 上限); 内容超过上限的格换行而不是继续加宽
WIDTH_BOUNDS = {
    "规则ID": (8, 12),
    "优先级": (8, 10),
    "用例": (8, 12),
    "输入": (10, 28),
    "输出": (10, 40),
    "期望": (10, 28),
    NOTE: (24, 60),
}
# 键值 sheet (输入 / 设置 / 变量定义 / 旧 匹配设置) 按列名定宽度区间
KV_WIDTH_BOUNDS = {
    "名称": (10, 16),
    "匹配": (10, 28),
    "重复键": (8, 12),
    "文件": (14, 36),
    "设置项": (12, 16),
    "值": (16, 40),
    "变量": (12, 24),
    "定义": (24, 60),
    NOTE: (30, 70),
}
DEFAULT_BOUNDS = (8, 40)
PADDING = 2  # 内容宽度之外的留白 (字符)
LINE_HEIGHT = 15.0  # Excel 默认行高 (磅), 换行格的行高 = 行数 * 它
DATE_FORMAT = "yyyy-mm-dd"
KV_SHEETS = ("输入", "设置", "变量定义", "匹配设置")

_THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
BOLD = Font(bold=True)
REGULAR = Font(bold=False)
NO_FILL = PatternFill(fill_type=None)


def fill(color: str) -> PatternFill:
    return PatternFill("solid", fgColor=color)


# ---------------------------------------------------------------------------
# 宽度与文本
# ---------------------------------------------------------------------------


def cell_text(value: Any) -> str:
    """单元格显示出来的文本 (估宽用): 日期 yyyy-mm-dd, 整数不带 .0。"""
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return "2026-01-01"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def display_width(value: Any) -> int:
    """显示宽度 (字符数): 中文等全角 / 宽字符算 2, 多行取最宽的一行。"""
    lines = cell_text(value).split("\n")
    return max(sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in line) for line in lines)


def _clamp(width: int, bounds: tuple[int, int]) -> int:
    lo, hi = bounds
    return max(lo, min(hi, width + PADDING))


def _lines_needed(value: Any, width: int) -> int:
    """按列宽估算一个格要占几行 (换行格的行高用)。"""
    usable = max(width - PADDING, 1)
    total = 0
    for line in cell_text(value).split("\n"):
        w = display_width(line)
        total += max(1, math.ceil(w / usable))
    return total


def _style_cell(
    cell: Any, *, bold: bool | None = None, color: str | None = None, horizontal: str = "left", wrap: bool = False
) -> None:
    try:
        cell.border = BORDER
        cell.alignment = Alignment(horizontal=horizontal, vertical="center", wrap_text=wrap)
        if bold is not None:
            cell.font = BOLD if bold else REGULAR
        cell.fill = fill(color) if color else NO_FILL
    except AttributeError:  # 合并区域里的从属格 (MergedCell) 在旧版 openpyxl 里不能设样式, 跳过即可
        return


def _fix_value_format(cell: Any) -> None:
    """= 开头的文本按文本存 (不是公式); 日期格统一格式。"""
    v = cell.value
    if isinstance(v, str) and v.startswith("="):
        cell.data_type = "s"
        cell.number_format = "@"
    elif isinstance(v, (datetime, date)):
        cell.number_format = DATE_FORMAT


def _set_widths_and_wrap(
    ws: Worksheet, columns: dict[int, tuple[int, int]], rows: range, *, wrap_rows: range | None = None
) -> dict[int, int]:
    """按内容给每列定宽 (在区间内), 超宽的格换行并估算行高。columns: 列号 (1 起) -> 宽度区间。返回每列最终宽度。"""
    widths: dict[int, int] = {}
    for col, bounds in columns.items():
        content = max((display_width(ws.cell(row=r, column=col).value) for r in rows), default=0)
        widths[col] = _clamp(content, bounds)
        ws.column_dimensions[get_column_letter(col)].width = widths[col]
    for r in wrap_rows if wrap_rows is not None else rows:
        lines = 1
        for col, width in widths.items():
            cell = ws.cell(row=r, column=col)
            if cell.value is None:
                continue
            need = _lines_needed(cell.value, width)
            if need > 1 or "\n" in cell_text(cell.value):
                try:
                    cell.alignment = Alignment(horizontal=cell.alignment.horizontal, vertical="center", wrap_text=True)
                except AttributeError:
                    continue
                lines = max(lines, need)
        ws.row_dimensions[r].height = LINE_HEIGHT * lines + 2 if lines > 1 else None
    return widths


def _last_row(ws: Worksheet) -> int:
    grid = _rows(ws)
    for i in range(len(grid), 0, -1):
        if any(v not in (None, "") for v in grid[i - 1]):
            return i
    return 0


# ---------------------------------------------------------------------------
# 各类 sheet
# ---------------------------------------------------------------------------


def style_rule_table(ws: Worksheet, key_kind: str, allowed: tuple[str, ...]) -> None:
    """决策表 / 测试: 表头按内容定位 (与加载器同一套逻辑), 之上是元信息行, 之下是正文。"""
    grid = _rows(ws)
    cols, kind_row, body_start = _parse_header(grid, ws.title, key_kind, allowed)
    kind_by_col = {i + 1: kind for i, kind, _ in cols}
    header_rows = (kind_row + 1, kind_row + 2)
    last = max(_last_row(ws), header_rows[1])
    body = range(body_start + 1, last + 1)
    ncols = max(kind_by_col)

    for r in range(1, kind_row + 1):  # 元信息行: 标签加粗, 其余有字的格当说明 (不参与定宽, 让文字溢出到右边的空格)
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            if cell.value in (None, ""):
                continue
            label = str(cell.value).strip().rstrip(":：") == POLICY_LABEL or c == 1
            _style_cell(cell, bold=label, color=None if label else FILLS[NOTE])
    for r in header_rows:
        for c in range(1, ncols + 1):
            kind = kind_by_col.get(c)
            _style_cell(ws.cell(row=r, column=c), bold=True, color=FILLS.get(kind or ""), horizontal="center")
    for r in body:
        for c in range(1, ncols + 1):
            kind = kind_by_col.get(c)
            cell = ws.cell(row=r, column=c)
            color = FILLS.get(kind) if kind and kind not in ID_KINDS else None
            _style_cell(cell, bold=False, color=color, horizontal="center" if kind in CENTERED_KINDS else "left")
            _fix_value_format(cell)

    columns = {c: WIDTH_BOUNDS.get(kind_by_col.get(c, ""), DEFAULT_BOUNDS) for c in range(1, ncols + 1)}
    _set_widths_and_wrap(ws, columns, range(header_rows[0], last + 1), wrap_rows=body)
    frozen = 0
    for c in range(1, ncols + 1):  # 最左边连续的 ID 类列冻结住
        if kind_by_col.get(c) in ID_KINDS:
            frozen = c
        else:
            break
    ws.freeze_panes = f"{get_column_letter(frozen + 1)}{body_start + 1}"


def style_kv_sheet(ws: Worksheet) -> None:
    """输入 / 设置 / 变量定义 / 旧 匹配设置: 第 1 行表头, A 列加粗, 「说明」列浅黄。"""
    last = _last_row(ws)
    if last == 0:
        return
    header = [cell_text(v).strip() for v in _rows(ws)[0]]
    ncols = max((i + 1 for i, h in enumerate(header) if h), default=0)
    if ncols == 0:
        return
    for c in range(1, ncols + 1):
        _style_cell(ws.cell(row=1, column=c), bold=True, color=HEADER_FILL, horizontal="center")
    for r in range(2, last + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            is_note = header[c - 1] == NOTE
            _style_cell(cell, bold=(c == 1), color=FILLS[NOTE] if is_note else None)
            _fix_value_format(cell)
    columns = {c: KV_WIDTH_BOUNDS.get(header[c - 1], DEFAULT_BOUNDS) for c in range(1, ncols + 1)}
    _set_widths_and_wrap(ws, columns, range(1, last + 1), wrap_rows=range(2, last + 1))
    ws.freeze_panes = "A2"


def style_data_sheet(ws: Worksheet, *, text_cols: tuple[int, ...] = ()) -> None:
    """样例数据表: 表头灰底加粗, 细边框, 按内容定宽, 冻结首行; text_cols 里的列设成文本格式 (电话 / 带前导零的编号)。"""
    last = _last_row(ws)
    if last == 0:
        return
    ncols = ws.max_column
    for c in range(1, ncols + 1):
        _style_cell(ws.cell(row=1, column=c), bold=True, color=HEADER_FILL, horizontal="center")
    for r in range(2, last + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            _style_cell(cell, bold=False)
            if c in text_cols:
                cell.number_format = "@"
            elif isinstance(cell.value, (datetime, date)):
                cell.number_format = DATE_FORMAT
    _set_widths_and_wrap(ws, dict.fromkeys(range(1, ncols + 1), DEFAULT_BOUNDS), range(1, last + 1), wrap_rows=range(0))
    ws.freeze_panes = "A2"


def style_rules_workbook(wb: Workbook) -> None:
    """按 sheet 名套用版式; 不认识的 sheet 不动。"""
    for ws in wb.worksheets:
        if ws.title == "决策表":
            style_rule_table(ws, "规则ID", RULE_KINDS)
        elif ws.title == "测试":
            style_rule_table(ws, "用例", TEST_KINDS)
        elif ws.title in KV_SHEETS:
            style_kv_sheet(ws)


def format_workbook(path: str | Path, out: str | Path | None = None) -> Path:
    """先按加载器校验结构 (结构错就报加载器的原话), 再套版式。默认原地改写。

    openpyxl 往返会保留批注、数据验证、条件格式, 但会丢图片和图表。
    """
    src = Path(path)
    if src.suffix.lower() != ".xlsx":
        raise ValueError(f"format 只处理 Excel 规则工作簿 (.xlsx): {src.name}")
    load_rules(src)
    wb = load_workbook(src, data_only=False)
    style_rules_workbook(wb)
    target = Path(out) if out else src
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        wb.save(target)
    except PermissionError as exc:
        raise PermissionError(f"无法写入 {target}: 文件可能正被 Excel 打开, 关闭后再试") from exc
    return target
