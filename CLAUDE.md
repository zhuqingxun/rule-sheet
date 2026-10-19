# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目定位

rule-sheet: 把「两份 Excel 互相比较, 按规则更新字段」的业务规则从代码里抽出来, 写成非编码人员能读、能改的规则文件 (Excel 决策表 / YAML), 由 Python 解释器执行。

业界对标: DMN 决策表 (Camunda / Drools / GoRules) 的思路, 简化为 Excel 原生布局 + 中文关键词。设计取舍见 README「条件之间的 AND / OR」与「命中策略」两节。

## 状态

uv 项目, src 布局, `rule-sheet` 命令, pytest 套件。2026-10-19 从 `D:/CODE/scratch/excel-rules-samples` 的单文件原型毕业迁入并完成打包。

| 路径 | 作用 |
|---|---|
| `src/rule_sheet/model.py` | 数据结构与关键词常量 (是/否/空白词、内置变量名) |
| `src/rule_sheet/expressions.py` | 表达式求值 (中文变量名替换 + 受限 eval)、`literal`、`split_top_level` |
| `src/rule_sheet/conditions.py` | 单元格一元测试 `unary_test`、YAML 组合词 `build_condition`、输出格 `resolve_output` |
| `src/rule_sheet/loaders.py` | `load_excel_rules` / `load_yaml_rules` -> `RuleSet` |
| `src/rule_sheet/engine.py` | `build_namespace` / `apply_rules` / `run_tests` / `run` 整体流程 / 写出 |
| `src/rule_sheet/cli.py` | `rule-sheet run|test` |
| `examples/customer-sync/` | 自带样例: 两份数据 + 决策表 + YAML + Gherkin + 生成脚本 |
| `tests/` | 单元格语法 / 表达式 / 引擎 / 加载器 / 端到端 |

## 运行与验证

```powershell
uv sync
uv run rule-sheet run  examples/customer-sync/规则_决策表.xlsx
uv run rule-sheet run  examples/customer-sync/rules.yaml
uv run pytest
uv run pytest tests/test_examples.py::test_both_frontends_agree   # 单个用例
uv run pytest -k excel                                            # 按名字筛选 (参数化 id 为 excel / yaml)
uv run ruff check . ; uv run ruff format --check .
uv run --script examples/customer-sync/make_samples.py     # 改了样例设计后重新生成 xlsx
```

验收标准: `uv run pytest` 全绿。其中 `tests/test_examples.py` 保证两种前端对样例数据输出一致 (5 行变更) 且各自 8 条内嵌用例通过。改解释器语义后必须跑全量, 不要只跑一个前端。

## 执行管线 (跨文件, 改引擎前先读)

`load_rules` 按后缀分派到 `load_excel_rules` / `load_yaml_rules`, 两者汇合成同一个 `RuleSet`; 之后 `engine.run` 依次: `run_tests` -> `read_table` 读左右表 -> `build_namespace` 合并 -> `add_derived` 派生变量 -> `apply_rules` -> 拼审计列 -> `write_xlsx_safely` (被 Excel 占用时自动 `_v2`/`_v3`)。

- **一切向量化**: `Namespace` 是 `{变量名: pd.Series}` (整列), `Condition = Callable[[Namespace], Series]` 返回布尔 mask, `apply_rules` 用 `Series.mask` 按优先级倒序覆盖写入。不要在规则层写逐行循环
- **内嵌测试复用引擎**: `run_tests` 把每条用例构造成 1 行的 `Namespace`, 走同一套 `add_derived(lenient=True)` + `apply_rules`。改 `apply_rules` 语义会同时影响真实数据与用例
- **两前端的分叉点** (改一边必须查另一边):
  - 条件: Excel 每个输入格一个 `cell_condition` (走 `unary_test`); YAML 整棵 `当:` 树一个 `build_condition` (走 `eval_expr`)
  - `rs.variables` (用例未给即视为空的变量清单): Excel 直接取表头输入列; YAML 靠 `<别名>.\w+` 正则扫描 `当` / `则` / `变量` 的文本。YAML 里新增写变量名的位置要同步进这段扫描
  - `rs.outputs` 顺序: Excel 按表头输出列; YAML 按 `则:` 首次出现
- **表达式求值**: `eval_expr` 把中文变量名按长度降序替换为 `_vN` 再 `eval`, 可用函数只有 `expressions.FUNCS` 里那几个; 加内置函数就加在这个 dict
- **Excel 版式契约** (`loaders._parse_header` 与 `make_samples.py` 必须一致): sheet 名固定为 匹配设置 / 变量定义 / 决策表 / 测试; 决策表 B1 为命中策略, 第 2 行是列类型 (规则ID / 优先级 / 说明 / 输入 / 输出; 测试 sheet 为 用例 / 输入 / 期望), 第 3 行是变量名, 第 4 行起正文; 优先级留空取 `1000 + 行序`

## 核心语义约定 (改代码前先读)

- **读表**: 一律走 `engine.read_table` (`dtype=object` + `infer_objects`)。pandas 3 的 `read_excel` 默认把文本单元格的 "007" 读成 7、电话读成浮点, 直接用会丢数据 (2026-10-19 实测 pandas 3.0.6)
- **匹配**: 左表 left join 右表, 变量命名 `<左表别名>.<列>` / `<右表别名>.<列>`, 内置 `<右表别名>.存在`、`基准日期`
- **路径**: 规则文件里的相对路径以规则文件所在目录为基准 (`--base-dir` 可覆盖)
- **单元格条件** (FEEL 一元测试子集): `-` 不关心 / 字面量等于 / 比较 (右侧可为变量名) / `[a..b]` 区间 / 逗号枚举 (OR) / `非空` `为空` / `not(...)`
- **空值语义**: 空值参与任何比较都不成立 (SQL 风格); `not(...)` 同样只对有值的格成立, 仅 `not(为空)` / `not(非空)` 是纯否定。YAML 表达式是 Python 语义, 需要自己写 `非空()`
- **布尔词**: `是` / `否` 只在被比较的列本身是布尔列时转成 True/False, 文本列里的「是」是普通文本
- **YAML 条件组合**: 列表默认 `全部满足`, `任一满足` / `不满足` 可嵌套
- **命中策略**: `优先级` 按输出列逐列合并 (数字小者胜, 未写该列的规则不参与); `唯一` 多命中即报错并指出匹配键值
- **输出格**: `-` 不改 / 字面值 / 变量名复制 / `=表达式`
- **审计列**: 输出文件追加 `命中规则`、`变更明细`
- **测试先行**: 规则文件内的测试用例不过, `run` 不处理真实数据; 用例没提到的变量视为空值

## 约定

- 中文 DSL 关键词是产品的一部分, 改名需同步 README 速查表、样例文件、`tests/test_unary.py`
- 样例数据的每一行对应一种规则情形 (见 README「样例数据设计」), 加规则时同步加数据行、Excel 测试 sheet 用例、YAML 测试用例, 并重新生成 xlsx
- 生成的 xlsx 进 git (体积小, 便于直接打开看), `examples/*/output/` 不进
- 版本号在 `pyproject.toml` 与 `src/rule_sheet/__init__.py` 两处, 同步改
- `examples/**` 被 ruff format 排除并忽略 E501: `make_samples.py` 里的规则表是手工按列对齐的, 不要去「修」它的格式
- `model.TestCase` 靠 `__test__ = False` 防止被 pytest 当用例收集; 新增 `Test` 开头的 dataclass 同样要加

## 已知局限 (候选 roadmap, 未排序)

- 表达式走 Python `eval`, 规则文件视为可信输入; 多人场景需换受限求值器
- 一对多匹配只支持「取第一条」
- 无静态校验: 重叠规则、永不命中的规则、未覆盖的输入组合
- Gherkin 无执行器
