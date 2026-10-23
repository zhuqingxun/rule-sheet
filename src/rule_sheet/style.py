"""规则工作簿的版式规范: 颜色、列宽、换行、冻结, 一处定义, 样例生成脚本与 `rule-sheet format` 共用。

规范 (使用指南 3.6 有同样的表):
- 列类型配色: 输入 浅蓝 / 输出 与 期望 浅绿 / 说明 浅黄 / 规则ID 优先级 用例 浅灰 (只在表头); 键值 sheet 表头 灰
- 表头加粗; 表区域细灰边框; 垂直居中; 输入 与 ID 类列水平居中, 输出 / 说明 左对齐
- 列宽按内容算 (中文算 2 个字符宽), 每类列有上下限; 超过上限的格换行, 行高按行数估算后显式写入
- 冻结: 决策表 / 测试 冻结到列名行之下、最左边连续 ID 类列之右; 「设置」冻结到输入表列名之下; 其它 sheet 冻结首行
- 元信息行 (决策表的 命中策略 / 拼接分隔符, 「设置」的设置项): 标签加粗, 值不上色, 说明浅黄、合并到表宽换行
- 「设置」的分段标题行 (设置项 / 输入表): 深蓝底白字横跨表宽
- sheet 顺序固定: 设置 / 变量定义 / 字段组 / 决策表 / 测试, 不认识的排最后
- 以 = 开头的格按文本存 (否则 Excel 当成公式, 显示 #NAME?); 日期格 yyyy-mm-dd
表头在哪一行由 loaders 的同一套逻辑判断, 所以 format 与加载器对「哪里是表头」永远一个答案。
"""

from __future__ import annotations

import contextlib
import math
import unicodedata
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .loaders import POLICY_LABEL, RULE_KINDS, TEST_KINDS, _parse_header, _rows, input_header_row, load_rules
from .model import FIELD_GROUP_KEY, JOIN_SEP_KEY, LEGACY_INPUTS_SHEET, SETTING_KEYS, SETTINGS_SHEET

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
# 键值 sheet (「设置」的输入表 / 变量定义 / 字段组 / 旧 输入 / 旧 匹配设置) 按列名定宽度区间
KV_WIDTH_BOUNDS = {
    "名称": (10, 16),
    "匹配": (10, 28),
    "筛选": (16, 40),
    "重复键": (8, 12),
    "独有行": (8, 10),
    "排序": (10, 24),
    "分隔符": (8, 10),
    "空值占位": (8, 12),
    "带序号": (8, 10),
    "文件": (14, 36),
    "设置项": (12, 16),
    "值": (16, 40),
    "变量": (12, 24),
    "定义": (24, 60),
    "列名": (24, 70),
    NOTE: (30, 70),
}
DEFAULT_BOUNDS = (8, 40)
PADDING = 2  # 内容宽度之外的留白 (字符); Excel 显示的列宽比写入值少约 0.6, 实际留白约 1.4
# 一个中文字在 Excel 默认字体 (等线 11) 下约占 2.05 (常规) ~ 2.15 (加粗) 个列宽单位 (2026-10-23 用 Excel AutoFit 实测);
# 按 2 算会让长度接近列宽的格被右边的格挡住一截
WIDE_CHAR = 2.2
LINE_HEIGHT = 15.0  # Excel 默认行高 (磅), 换行格的行高 = 行数 * 它
DATE_FORMAT = "yyyy-mm-dd"
KV_SHEETS = ("输入", "设置", "变量定义", "字段组", "匹配设置")
# sheet 的固定顺序 (format 时重排): 先「设置」(输入从哪来), 再 变量定义 / 字段组 (规则里用到的名字), 再决策表与测试;
# 旧的 输入 / 匹配设置 紧跟「设置」; 不认识的 sheet 排在最后, 相互顺序不变
SHEET_ORDER = (SETTINGS_SHEET, LEGACY_INPUTS_SHEET, "匹配设置", "变量定义", FIELD_GROUP_KEY, "决策表", "测试")
# 「设置」里 设置项 / 输入表 两块的分段标题行: 深蓝底白字, 横跨整张表, 说明合并到表宽并换行 (不能溢出到色带外: 白字白底)
SECTION_FILL = "305496"
SECTION_LABEL_FONT = Font(bold=True, color="FFFFFF")
SECTION_NOTE_FONT = Font(bold=False, color="FFFFFF")
SECTION_MIN_HEIGHT = 20.0
# 「设置」sheet 两块内容上方的标题行: B 列留空, 加载器按「值为空」跳过 (设置项 这个词也被显式跳过);
# 样例生成脚本与 YAML -> Excel 共用
SETTINGS_TITLE = (
    "设置项",
    None,
    "每行一项: A 列名字 / B 列值 / 再往右是说明; 只认 基准日期 / 输出, 写别的加载时报错; B 列留空 = 取默认值",
)
INPUTS_TITLE = (
    "输入表",
    None,
    "下一行是列名, 之后每行声明一个输入: 名称 = 规则里的前缀; 匹配 = 主表写 主表, 其它写对齐到主表的键; "
    "文件 = 默认文件 (命令行 -i 可覆盖)。输入表下面不要再写别的",
)

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


def display_width(value: Any) -> float:
    """显示宽度 (Excel 列宽单位): 中文等全角 / 宽字符算 WIDE_CHAR, 其余算 1, 多行取最宽的一行。"""
    lines = cell_text(value).split("\n")
    return max(sum(WIDE_CHAR if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in line) for line in lines)


def _clamp(width: float, bounds: tuple[int, int]) -> float:
    lo, hi = bounds
    return max(lo, min(hi, width + PADDING))


def _widen(ws: Worksheet, col: int, value: Any, bounds: tuple[int, int]) -> None:
    """按一个格的内容把列放宽 (只放宽不收窄, 仍在区间内): 元信息行的标签 / 值不参与正文定宽, 但要显示全。"""
    dim = ws.column_dimensions[get_column_letter(col)]
    dim.width = max(dim.width or 0, _clamp(display_width(value), bounds))


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


def _meta_label(value: Any) -> bool:
    """决策表元信息行的标签格 (命中策略 / 拼接分隔符, 可带冒号)。"""
    return str(value or "").strip().rstrip(":：") in (POLICY_LABEL, JOIN_SEP_KEY)


def _span_note(ws: Worksheet, row: int, start: int, end: int) -> int:
    """把一行的说明格从 start 合并到 end (表的右边界), 换行显示, 不再溢出到表外。返回说明实际占到的列。

    右边还有别的字就不合并 (合并会把它们清掉), 只在 start 一格里换行; 已有的合并先拆开再按现在的表宽合并 (format 幂等)。
    """
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row <= row <= rng.max_row and rng.max_col >= start and rng.min_col <= max(start, end):
            if (rng.min_row, rng.max_row, rng.min_col, rng.max_col) == (row, row, start, end):
                return end
            ws.unmerge_cells(str(rng))
    if end <= start or any(ws.cell(row=row, column=c).value not in (None, "") for c in range(start + 1, end + 1)):
        return start
    ws.merge_cells(start_row=row, start_column=start, end_row=row, end_column=end)
    return end


def _fit_note_height(ws: Worksheet, row: int, start: int, end: int, min_height: float | None = None) -> None:
    """合并后的说明按合并区的总宽估行数定行高 (列宽定完之后再调)。"""
    width = sum(ws.column_dimensions[get_column_letter(c)].width or 8.43 for c in range(start, end + 1))
    width += 0.6 * (end - start)  # 合并区还含列间的网格线 (每条约 0.6 个列宽单位, Excel 实测)
    lines = _lines_needed(ws.cell(row=row, column=start).value, int(width))
    height = LINE_HEIGHT * lines + 2 if lines > 1 else None
    if min_height is not None:
        height = max(height or 0, min_height)
    ws.row_dimensions[row].height = height


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

    notes: dict[int, tuple[int, int]] = {}  # 元信息行 -> 说明合并区 (起, 止)
    for r in range(1, kind_row + 1):  # 元信息行: 标签加粗, 紧挨着的值不上色, 再往右是说明 (浅黄, 合并到表宽换行)
        filled = [c for c in range(1, ncols + 1) if ws.cell(row=r, column=c).value not in (None, "")]
        label = next((c for c in filled if _meta_label(ws.cell(row=r, column=c).value)), None)
        for c in filled:
            cell = ws.cell(row=r, column=c)
            if c == label or (label is None and c == 1):
                _style_cell(cell, bold=True)
            elif label is not None and c == label + 1:
                _style_cell(cell, bold=False)
            else:
                _style_cell(cell, bold=False, color=FILLS[NOTE])
        rest = [c for c in filled if label is not None and c > label + 1]
        if rest:  # 只动有 命中策略 / 拼接分隔符 标签的行; 单独写一格策略名的旧写法不碰
            notes[r] = (rest[0], _span_note(ws, r, rest[0], ncols))
            ws.cell(row=r, column=rest[0]).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
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
    for r in range(1, kind_row + 1):  # 元信息行的标签 (命中策略 / 拼接分隔符) 和紧挨着的值要显示全
        for c in range(1, ncols + 1):
            value = ws.cell(row=r, column=c).value
            if _meta_label(value):
                _widen(ws, c, str(value).strip(), columns[c])
                if c + 1 in columns:
                    _widen(ws, c + 1, ws.cell(row=r, column=c + 1).value, columns[c + 1])
    for r, (start, end) in notes.items():
        _fit_note_height(ws, r, start, end)
    frozen = 0
    for c in range(1, ncols + 1):  # 最左边连续的 ID 类列冻结住
        if kind_by_col.get(c) in ID_KINDS:
            frozen = c
        else:
            break
    ws.freeze_panes = f"{get_column_letter(frozen + 1)}{body_start + 1}"


def style_kv_sheet(ws: Worksheet, header_row: int = 1) -> None:
    """变量定义 / 字段组 / 旧 输入 / 旧 匹配设置 / 「设置」的输入表: header_row 行是表头, A 列加粗, 「说明」列浅黄。"""
    last = _last_row(ws)
    if last < header_row:
        return
    header = [cell_text(v).strip() for v in _rows(ws)[header_row - 1]]
    ncols = max((i + 1 for i, h in enumerate(header) if h), default=0)
    if ncols == 0:
        return
    for c in range(1, ncols + 1):
        _style_cell(ws.cell(row=header_row, column=c), bold=True, color=HEADER_FILL, horizontal="center")
    body = range(header_row + 1, last + 1)
    for r in body:
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            is_note = header[c - 1] == NOTE
            _style_cell(cell, bold=(c == 1), color=FILLS[NOTE] if is_note else None)
            _fix_value_format(cell)
    columns = {c: KV_WIDTH_BOUNDS.get(header[c - 1], DEFAULT_BOUNDS) for c in range(1, ncols + 1)}
    _set_widths_and_wrap(ws, columns, range(header_row, last + 1), wrap_rows=body)
    ws.freeze_panes = f"A{header_row + 1}"


def style_settings_sheet(ws: Worksheet) -> None:
    """「设置」: 上方设置项行 (设置项加粗 / 值 / 说明浅黄, 说明不参与定宽, 溢出到右边), 下方输入表按键值 sheet 的版式;
    冻结到输入表列名之下。没有输入表的是旧版只有设置项的「设置」, 按第 1 行表头的键值 sheet 处理。"""
    grid = _rows(ws)
    header = input_header_row(grid)
    if header is None:
        style_kv_sheet(ws)
        return
    ncols = max((i + 1 for i, h in enumerate(grid[header]) if cell_text(h).strip()), default=3)
    items: list[int] = []  # 设置项行的设置项所在列, 放宽用
    notes: dict[int, tuple[int, int, float | None]] = {}  # 行 -> 说明合并区 (起, 止, 最小行高)
    for r in range(1, header + 1):  # 与 loaders.read_setting_rows 同一读法: 第一个有字的格是设置项, 紧挨着的是值
        cells = [c for c in range(1, ws.max_column + 1) if ws.cell(row=r, column=c).value not in (None, "")]
        if not cells:
            continue
        j, end = cells[0], max(ncols, cells[-1])
        label = cell_text(ws.cell(row=r, column=j).value).strip().rstrip(":：")
        # 值为空又不是设置项 (输出 留空 = 取默认值, 仍是设置项) 的行加载器会跳过: 当分段标题, 深蓝底白字横跨表宽
        title = ws.cell(row=r, column=j + 1).value in (None, "") and label not in SETTING_KEYS
        if title:
            for c in range(1, end + 1):
                cell = ws.cell(row=r, column=c)
                _style_cell(cell, color=SECTION_FILL)
                with contextlib.suppress(AttributeError):
                    cell.font = SECTION_LABEL_FONT if c == j else SECTION_NOTE_FONT
        else:
            items.append(r)
            for c in cells:
                cell = ws.cell(row=r, column=c)
                _style_cell(cell, bold=(c == j), color=FILLS[NOTE] if c > j + 1 else None)
                _fix_value_format(cell)
        rest = [c for c in cells if c > j + 1]
        if rest:
            notes[r] = (rest[0], _span_note(ws, r, rest[0], end), SECTION_MIN_HEIGHT if title else None)
            ws.cell(row=r, column=rest[0]).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        elif title:
            notes[r] = (j, j, SECTION_MIN_HEIGHT)
    style_kv_sheet(ws, header_row=header + 1)
    # 设置项与值要显示全 (右边有说明挡着, 不能溢出): 在输入表算出的宽度上按它们放宽, 区间取 设置项 / 值 的
    for r in items:
        j = next(c for c in range(1, ws.max_column + 1) if ws.cell(row=r, column=c).value not in (None, ""))
        _widen(ws, j, ws.cell(row=r, column=j).value, KV_WIDTH_BOUNDS["设置项"])
        _widen(ws, j + 1, ws.cell(row=r, column=j + 1).value, KV_WIDTH_BOUNDS["值"])
    for r, (start, end, min_height) in notes.items():  # 列宽定完再按合并区总宽估行高
        _fit_note_height(ws, r, start, end, min_height)


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


def order_sheets(wb: Workbook) -> None:
    """按 SHEET_ORDER 重排 sheet, 不认识的排最后 (相互顺序不变); 打开时停在第一张, 且只选中它
    (move_sheet 不动选中状态, 留着多张 tabSelected, Excel 打开就是「组合工作表」, 一改改多张)。"""
    rank = {name: i for i, name in enumerate(SHEET_ORDER)}
    for i, name in enumerate(sorted(wb.sheetnames, key=lambda n: rank.get(n, len(rank)))):
        wb.move_sheet(wb[name], i - wb.sheetnames.index(name))
    wb.active = 0
    for i, name in enumerate(wb.sheetnames):
        wb[name].sheet_view.tabSelected = i == 0


def style_rules_workbook(wb: Workbook) -> None:
    """按 sheet 名套用版式并按固定顺序重排; 不认识的 sheet 只挪到最后, 内容不动。"""
    order_sheets(wb)
    for ws in wb.worksheets:
        if ws.title == "决策表":
            style_rule_table(ws, "规则ID", RULE_KINDS)
        elif ws.title == "测试":
            style_rule_table(ws, "用例", TEST_KINDS)
        elif ws.title == SETTINGS_SHEET:
            style_settings_sheet(ws)
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
