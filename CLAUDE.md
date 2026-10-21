# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目定位

rule-sheet: 把「若干份 Excel 按键对上, 按规则更新其中一份」的业务规则从代码里抽出来, 写成非编码人员能读、能改的规则文件 (Excel 决策表 / YAML), 由 Python 解释器执行。模型是 规则文件 (形参: 输入声明 + 规则 + 测试) + 输入文件绑定 (实参: 规则里的默认文件或命令行 `-i 名称=路径`) → 输出文件 (主表副本 + 审计列)。

业界对标: DMN 决策表 (Camunda / Drools / GoRules) 的思路, 简化为 Excel 原生布局 + 中文关键词。设计取舍见 README「条件之间的 AND / OR」与「命中策略」两节。

## 状态

uv 项目, src 布局, `rule-sheet` 命令, pytest 套件。2026-10-19 从 `D:/CODE/scratch/excel-rules-samples` 的单文件原型毕业迁入并完成打包。

| 路径 | 作用 |
|---|---|
| `src/rule_sheet/model.py` | 数据结构与关键词常量 (是/否/空白词/清空/空、`CLEAR` / `NULL` 哨兵、内置变量名) |
| `src/rule_sheet/expressions.py` | 表达式求值 (中文变量名替换 + 受限 eval)、`literal`、`split_top_level` |
| `src/rule_sheet/conditions.py` | 单元格一元测试 `unary_test`、YAML 组合词 `build_condition`、输出格 `resolve_output` |
| `src/rule_sheet/loaders.py` | `load_excel_rules` / `load_yaml_rules` -> `RuleSet` |
| `src/rule_sheet/engine.py` | `build_namespace` / `apply_rules` / `run_tests` / `run` 整体流程 / 写出 |
| `src/rule_sheet/style.py` | 规则工作簿的版式规范 (列类型配色 / 列宽区间 / 换行行高 / 冻结 / `=` 格存文本), `style_rules_workbook` 给样例生成脚本用, `format_workbook` 给 `rule-sheet format` 用; 表头定位复用 `loaders._parse_header` |
| `src/rule_sheet/convert.py` | Excel -> YAML 无损导出 (`CellTranslator.cell` 对照 `unary_test` 逐形态翻译, 续行 -> `任一满足`); YAML -> Excel 只对表格子集 (`ExprToCell` 在 `substitute_names` 之后对 AST 做模式匹配: 比较 / 链式区间 / in / 属于 / 文本函数 / 判空 / 裸布尔名, 顶层 `任一满足` -> 续行, 同一变量的上下界合成区间, 其它报错指向变量定义); 两份规则文件的语义比对 `diff_rules` (结构逐项比 + 用例合并后在两套规则上各跑一次比输出与命中, 抽样等价不是证明) |
| `src/rule_sheet/cli.py` | `rule-sheet run|test|format|export|diff` |
| `examples/customer-sync/` | 自带样例: 两份数据 + 决策表 + YAML + Gherkin + 生成脚本 (逐列合并的 `优先级` 策略) |
| `examples/status-machine/` | 第二个样例: 字段名含 `-` / 输出格 `清空` / 命中策略 `首次命中` (2026-10-19 为 issue #1-#4 加); 0.3 (issue #5-#8) 扩成: 决策表同 ID 续行表达跨变量 OR、变量定义里 `且` / `或`、测试 sheet 表头从第 1 行开始 |
| `examples/position-tags/` | 第三个样例 (0.6, issue #9-#16): 多标签分类, 命中策略 `多条拼接`、单元格文本关键字 `包含` / `不包含` / `开头是`、清单文件当参照表 (`关键岗位清单.存在`)、`数字()` / `N天` / `如果()` / `拼接()` |
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
uv run python examples/customer-sync/make_samples.py       # 改了样例设计后重新生成 xlsx (脚本 import rule_sheet.style, 不能用 --script 隔离环境跑)
uv run python examples/status-machine/make_samples.py      # 第二个样例同理
uv run python examples/position-tags/make_samples.py       # 第三个样例同理
uv run rule-sheet format examples/customer-sync/规则_决策表.xlsx   # 按版式规范重排任意规则工作簿 (原地)
uv run rule-sheet export examples/customer-sync/规则_决策表.xlsx -o /tmp/r.yaml   # Excel -> YAML 无损导出
uv run rule-sheet diff   examples/customer-sync/规则_决策表.xlsx examples/customer-sync/rules.yaml   # 语义比对, 一致返回 0
```

验收标准: `uv run pytest` 全绿。其中 `tests/test_examples.py` 保证两种前端对样例数据输出一致 (customer-sync 5 行变更 / 8 条内嵌用例; status-machine 7 行变更 / 10 条内嵌用例, 且换成 `优先级` 策略时 T2 必须失败; position-tags 8 行变更 / 8 条内嵌用例, 每行的 命中规则 / 标签 / 显示名 都有预期值)。改解释器语义后必须跑全量, 不要只跑一个前端。

## 执行管线 (跨文件, 改引擎前先读)

`load_rules` 按后缀分派到 `load_excel_rules` / `load_yaml_rules`, 两者汇合成同一个 `RuleSet`; 之后 `engine.run` 依次: `run_tests` -> `bind_inputs` (命令行 `-i` 相对当前目录优先, 否则规则里的默认文件相对规则目录, 都没有报错) -> `read_table` 读各输入表 (`split_sheet` 拆 `#sheet名`) -> `build_namespace` 以主表为锚逐个 left join 参照表 -> `add_derived` 派生变量 -> `apply_rules` -> 拼审计列 -> `write_xlsx_safely` (被 Excel 占用时自动 `_v2`/`_v3`)。

- **一切向量化**: `Namespace` 是 `{变量名: pd.Series}` (整列); `Condition = Callable[[Namespace], Any]` 可返回含 NaN 的 Series 或标量, 由 `conditions.to_mask` 统一归一成布尔 mask (NaN 视为不命中); `apply_rules` 用 `Series.mask` 按优先级倒序覆盖写入。不要在规则层写逐行循环
- **内嵌测试复用引擎**: `run_tests` 把每条用例构造成 1 行的 `Namespace`, 走同一套 `add_derived(lenient=True)` + `apply_rules`。改 `apply_rules` 语义会同时影响真实数据与用例
- **两前端的分叉点** (改一边必须查另一边):
  - 条件: Excel 每个输入格一个 `cell_condition` (走 `unary_test`), 相邻同 ID 行合成 `any_of([all_of(行)...])`; YAML 整棵 `当:` 树一个 `build_condition` (走 `eval_expr`)。决策表正文先由 `loaders.read_decision_rows` 读成 `RawRule` (原始格值), 加载器与导出器共用; **新增单元格形态要同步 `convert.CellTranslator.cell` 和 `tests/test_convert.py` 的逐形态对照表** (翻出来的表达式必须在含空值的数据上与 `unary_test` 的 mask 一致)
  - 两种写法的同步: Excel 是源, `rule-sheet export` 生成 YAML (生成物, 文件头标「请勿手改」), `rule-sheet diff` 对账; YAML -> Excel 只对表格子集, 新增表达式函数或单元格形态时 `ExprToCell` 与 `CellTranslator` 两个方向都要加 (测试里两个方向各有一张逐形态对照表, 都要求翻过去的结果在含空值的数据上 mask 一致); 三个样例的手写 `rules.yaml` 不由导出覆盖, 由 `test_convert.py` 用 diff 守住与 Excel 一致, 并要求它们都在表格子集内 (YAML -> Excel 往返一致)。设计见 `docs/设计架构.md` 5.15
  - `rs.variables` (用例未给即视为空的变量清单): Excel 取表头输入列 + `find_names` 扫描输入格 / 输出格 / 派生变量表达式得到的列 (`..` 区间语法前先补空格, 扫描器的边界不认 `.` 前缀); YAML 靠 `expressions.find_names` 扫描 `当` / `则` / `变量` 的文本 (名字可含 `-` `/` 全角括号, `-`/`/` 后紧跟 `别名.`/数字/已知名字才算运算符; 反引号包裹的整体算一个名字; 引号字符串不扫)。YAML 里新增写变量名的位置要同步进这段扫描。扫描器宁可少认不可多认: 多认会把 `本期.金额-上期.金额` 整体当成空变量, 测试静默算错
  - `rs.outputs` 顺序: Excel 按表头输出列; YAML 按 `则:` 首次出现
  - 测试列名校验 (`loaders._check_test_columns`) 两边共用: 期望列 ⊆ 输出列, 输入列 ⊆ 变量 ∪ 派生 ∪ `<参照表名>.存在` ∪ `基准日期` ∪ `<主表名>.<输出列>`
  - 输入声明 (`loaders._build_inputs`) 两边共用: Excel「输入」sheet 一行一个 (名称 / 匹配 / 重复键 / 文件), YAML `输入: {名: {文件, 匹配, 重复键}}`; 主表那行「匹配」必须写 `主表` (`model.PRIMARY_WORD`, 单输入可省), 其它写键 (`k` 或 `本表列 = 主表.列`, 逗号分隔, 只能对齐到主表); 「设置」sheet / `设置:` 只认 基准日期 / 输出。旧「匹配设置」/ `匹配:` 由 `_legacy_inputs` 折算 (保留一个版本), YAML 两者都没有也走旧默认 主表 / 本月
- **表达式求值**: `eval_expr` 把中文变量名按长度降序替换为 `_vN` (跳过引号字符串; 反引号 `` `名字` `` 整体查), `且`/`或`/`非` 换成 `and`/`or`/`not`, 然后 `ast.parse` → `expressions._Rewrite` 改写 → `compile` (按文本 `lru_cache`) → 受限 `eval`。改写把 `BoolOp`/`Not` 换成逐元素的 `_all`/`_any`/`_not`, 六个比较和 `in`/`not in` 换成「任一侧为空就不成立」的 helper, 链式比较拆段; 算术 / 函数调用不碰 (三元用函数 `如果()`, 不改写 `IfExp`)。时长字面量 `365天` 在名字替换之后、字符串之外改写成 `_days(365)`。可用函数只有 `expressions.FUNCS` 里那几个; 加内置函数就加在这个 dict, 且必须对 Series 逐行生效 (逐行的用 `_rowwise` 包, 参数里的 `空` 哨兵用 `_reject_null` 拒; 不开放 `float` / `str` / `int`, 它们对整列不是逐行的, `数字()` / `拼接()` 是替代)。`TEXT_FUNCS` 里的文本函数同时是决策表单元格的关键字 (`包含 a, b`, 走 `conditions._TEXT_TEST_RE`, 参数取原文不过 `literal`)。文本形态统一走 `expressions.text_of` (整数不带 .0, 日期 yyyy-mm-dd), 审计列的 `engine.fmt` 也用它。真值归一统一走 `expressions.truth` (NaN 算 False; `NaN.astype(bool)` 是 True, 别直接用)
- **Excel 版式契约** (`loaders._parse_header` 与两个 `make_samples.py` 必须一致): sheet 名固定为 输入 / 设置 / 变量定义 / 决策表 / 测试 (输入 与 决策表 缺失即报错, 其余可选; 旧的 匹配设置 仍认); 表头按内容定位: 写着 `规则ID` (测试 sheet: `用例`) 的第一行是列类型行 (`RULE_KINDS` / `TEST_KINDS` 白名单, 拼错报错; 标了类型的列必须有列名), 下一行是列名行, 再下一行起正文; 之前的行是元信息, 命中策略靠 `命中策略` 标签右边一格 (没有标签就认一个单独写着策略名的格, 按 `model.HIT_POLICIES` 校验); 优先级留空取 `1000 + 行序`; 相邻同 ID 行是续行 (优先级 / 输出留空或与首行相同, 否则报错; 不相邻的重复 ID 报错); 合并单元格由 `_rows` 展开成左上格的值。工作簿用 `data_only=False` 读: 以 `=` 开头的格 (如 `= 本期.类别-会议`) 会被 Excel / openpyxl 存成公式, `data_only=True` 读到 None 条件就静默变成不关心 (2026-10-19 实测); 生成脚本写这类格要 `cell.data_type = "s"`

## 核心语义约定 (改代码前先读)

- **读表**: 一律走 `engine.read_table` (`dtype=object` + `infer_objects`)。pandas 3 的 `read_excel` 默认把文本单元格的 "007" 读成 7、电话读成浮点, 直接用会丢数据 (2026-10-19 实测 pandas 3.0.6)
- **匹配**: 输入 ≥ 1, 恰一个主表 (输出的行 / 写回的列都是它的); 每个参照表单独按自己的键 left join 到主表 (左侧只带主表键列, 参照表之间同名列互不干扰; 参照表键为空的行先剔掉, 否则 pandas merge 把 NaN 配 NaN 让 `存在` 假阳性), 变量命名 `<输入名>.<列>`, 内置 `<参照表名>.存在`、`基准日期`; 参照表的键列不暴露。输入名对引擎是不透明字符串, 不要在引擎里写死 本期 / 上期 / 客户
- **路径**: 规则文件里的相对路径 (默认文件 / 输出) 以规则文件所在目录为基准 (`--base-dir` 可覆盖); 命令行 `-i` 给的相对当前目录。默认输出 `output/<主表文件名>_更新后.xlsx`
- **单元格条件** (FEEL 一元测试子集): `-` 不关心 / 字面量等于 / 比较 (右侧可为变量名, 不是表达式; 要算的东西放变量定义) / `[a..b]` 区间 / 逗号枚举 (OR) / `= 空` `!= 空` `空` (同义: `为空` `非空`) / `not(...)` / 文本关键字 `包含 a, b` `不包含` `开头是` `开头不是` `结尾是` `结尾不是` (关键字后要有空格或括号, 逗号项任一命中, 空值不成立)
- **空值语义**: 空值参与任何比较都不成立 (SQL 风格), 单元格和表达式一样 (0.3 起表达式经 AST 改写; 0.2 的表达式 `!=` 对空值为真); 单元格 `not(...)` 只对有值的格成立, 仅 `not(空)` / `not(为空)` / `not(非空)` 是纯否定; 表达式的 `非` 和 YAML 的 `不满足:` 是纯取反 (空值行也为真), 这是保留的不对称, 见 `docs/设计架构.md` 5.3
- **关键字 `空`** (`model.NULL_WORD` / 哨兵 `model.NULL`, 0.4 起): 判空的主推写法是 `x == 空` / `x != 空` (单元格 `= 空` / `!= 空` / 单独 `空`), `为空(x)` / `非空(x)` 是同义。它不是值: `_cmp` 与 `unary_test` 的比较分支特判哨兵, 其它运算符 / 算术 / 函数参数报错; `属于` 与 `in` 的枚举里认它; 输出格 `空` / `=空` 等价 `清空`; 测试期望格与输入格写 `空` = 空值; 派生变量算出 `NULL` 置成空列。不加 `null` 别名 (导出数据里常有文本 "null")
- **布尔词**: `是` / `否` 只在被比较的列本身是布尔列时转成 True/False, 文本列里的「是」是普通文本
- **条件组合**: YAML 列表默认 `全部满足`, `任一满足` / `不满足` 可嵌套; 表达式内 `且` / `或` / `非` (也认 and / or / not, 两侧留空格); 决策表跨变量 OR 用相邻同 ID 行 (每行内 AND, 行间 OR)
- **命中策略**: `优先级` 按输出列逐列合并 (数字小者胜, 未写该列的规则不参与); `唯一` 多命中即报错并指出匹配键值; `首次命中` 每行只算优先级数字最小的一条 (同优先级按书写顺序), 实现是在逐列合并前把 masks 按优先级做成互斥, 没有输出的规则也能拦住一行; `多条拼接` (DMN Collect, 0.6) 每个输出列把所有命中规则的值按优先级顺序用「设置」里的 `拼接分隔符` (默认 `/`) 接起来替换原值 (`engine._join_hits`; 空值不贡献, 没命中不动), 加载时 `loaders._check_join_policy` 拦「输出格写 清空」和「设了分隔符却不是这个策略」。设计取舍见 `docs/设计架构.md` 5.2 与 5.13
- **输出格**: `-` 不改 / `清空` (或 `空` / `=空`) 写空值 (`resolve_output` 返回 `model.CLEAR` 哨兵, 与 `-` 的 None 必须区分) / 字面值 / 变量名复制 / `=表达式`。单元格右侧与输出格里长得像 `<别名>.<列>` 却不在 ns 里的名字报 KeyError (`expressions.dangling_name`), 不当文字; 加引号才是文字
- **审计列**: 输出文件追加 `命中规则` (`首次命中` 下只记赢的那条)、`变更明细` (清空显示 `旧值 -> 空`)
- **测试先行**: 规则文件内的测试用例不过, `run` 不处理真实数据; 用例没提到的变量视为空值, 输入写 `空` 也是空值; 派生变量留空 = 按表达式推导 (给了值才不推导); 期望格留空 = 不变, 写 `空` / `清空` / `为空` = 应为空值 (不用 YAML `null`, 否则与 Excel 空格分叉); 测试 sheet / `测试:` 存在但为空报错, `test` 子命令对没有用例的文件返回 1
- **假通过零容忍**: 加载时能查出来的「写错了但什么都没发生」一律报错 (表头错位、列类型拼错、期望列 / 输入列拼错、重复 ID), 不要为了宽容而猜。新增校验时同步 `docs/使用指南.md` 第 9 节的报错表

## 约定

- 中文 DSL 关键词是产品的一部分, 改名需同步 README 速查表、样例文件、`tests/test_unary.py`; 加内置函数要同步三个样例「变量定义」sheet 的「可用函数」提示行并重新生成 xlsx
- 样例数据的每一行对应一种规则情形 (见 README「样例数据设计」), 加规则时同步加数据行、Excel 测试 sheet 用例、YAML 测试用例, 并重新生成 xlsx
- 版式只在 `style.py` 一处定义 (`FILLS` / `WIDTH_BOUNDS` / `KV_WIDTH_BOUNDS` 等), 生成脚本里不写颜色和列宽; 改规范要同步 `docs/使用指南.md` 3.6 的表并重新生成三个样例 (`tests/test_style.py` 对照常量校验样例)
- 生成的 xlsx 进 git (体积小, 便于直接打开看), `examples/*/output/` 不进
- 版本号在 `pyproject.toml` 与 `src/rule_sheet/__init__.py` 两处, 同步改
- `examples/**` 被 ruff format 排除并忽略 E501: `make_samples.py` 里的规则表是手工按列对齐的, 不要去「修」它的格式
- `model.TestCase` 靠 `__test__ = False` 防止被 pytest 当用例收集; 新增 `Test` 开头的 dataclass 同样要加

## 已知局限 (候选 roadmap, 未排序)

- 表达式走 Python `eval`, 规则文件视为可信输入; 多人场景需换受限求值器
- 一切逐行, 边界见 `docs/设计架构.md` 5.13: 一对多匹配只支持「取第一条」/ 报错, 无聚合 (分组排序拼接若做, 形态是「重复键」第三个选项加排序列, 不是新规则类型); 参照表只对齐到主表 (星型), 无链式连接; 输出只有主表一张, 不新增行 (「上期有本期无」另写一份以上期为主表的规则); 无跨列循环。issue 里再来这类诉求先对着 5.13 裁
- 无静态校验: 重叠规则、永不命中的规则、未覆盖的输入组合
- Gherkin 无执行器
