# CLAUDE.md

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
uv run ruff check . ; uv run ruff format --check .
uv run --script examples/customer-sync/make_samples.py     # 改了样例设计后重新生成 xlsx
```

验收标准: `uv run pytest` 全绿。其中 `tests/test_examples.py` 保证两种前端对样例数据输出一致 (5 行变更) 且各自 8 条内嵌用例通过。改解释器语义后必须跑全量, 不要只跑一个前端。

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

## 已知局限 (候选 roadmap, 未排序)

- 表达式走 Python `eval`, 规则文件视为可信输入; 多人场景需换受限求值器
- 一对多匹配只支持「取第一条」
- 无静态校验: 重叠规则、永不命中的规则、未覆盖的输入组合
- Gherkin 无执行器
