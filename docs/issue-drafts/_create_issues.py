"""把本次实验发现的 rule-sheet 缺口批量提 issue 到 zhuqingxun/rule-sheet。

每条 issue 的标题/标签/正文在此集中维护，便于复用与追溯。

PAT 从 git credential-manager 取（不硬编码）：
    echo -e "protocol=https\\nhost=github.com\\n" | git credential-manager get
也可用环境变量 GH_TOKEN 覆盖。
"""
import json
import os
import subprocess
import urllib.parse
import urllib.request


def get_pat() -> str:
    """从 git credential-manager 取 GitHub PAT。"""
    out = subprocess.run(
        ["git", "credential-manager", "get"],
        input="protocol=https\nhost=github.com\n",
        capture_output=True, text=True, check=True,
    ).stdout
    for line in out.splitlines():
        if line.startswith("password="):
            return line.split("=", 1)[1]
    raise RuntimeError("credential-manager 未找到 github.com 的 PAT")


PAT = os.environ.get("GH_TOKEN") or get_pat()
OWNER, REPO = "zhuqingxun", "rule-sheet"
PROXY = "http://z00821210:Zqx%402602@proxysg.huawei.com:8080"

OPENER = urllib.request.build_opener(
    urllib.request.ProxyHandler({"http": PROXY, "https": PROXY})
)


def gh(method, path, body=None):
    url = f"https://api.github.com{path}"
    data = json.dumps(body).encode("utf-8") if body is not None else None
    req = urllib.request.Request(
        url,
        data=data,
        method=method,
        headers={
            "Authorization": f"Bearer {PAT}",
            "Accept": "application/vnd.github+json",
            "User-Agent": "zhuqingxun",
            "Content-Type": "application/json",
        },
    )
    with OPENER.open(req, timeout=60) as resp:
        return resp.status, json.loads(resp.read().decode("utf-8"))


ISSUES = [
    {
        "title": "表达式求值：变量名含 `-` 时被拆开（如 `本期.类别-会议`）",
        "labels": ["bug", "i18n"],
        "body": """## 现象

字段名含连字符 `-` 时（中文业务表常见，如 `类别-会议`、`状态-会议`、`岗动-拟任命人-流程节点`），YAML `当:` 条件与变量定义表达式都求值失败：

```
[错误] T1: 表达式求值失败: '上期.类别-会议 == "变动通报"' (name '会议' is not defined)
[错误] T7: 表达式求值失败: '非空(本期.岗动-拟任命人-流程节点)' (name '拟任命人' is not defined)
```

## 复现

```yaml
匹配:
  左表: a.xlsx
  右表: b.xlsx
  左表别名: 本期
  右表别名: 上期
  匹配键: 本岗编码

规则:
  - 编号: S1
    当:
      - 为空(本期.类别-会议)
    则:
      状态-会议: 02-进展更新
```

运行 `uv run rule-sheet test xxx.yaml` 即报 `name '会议' is not defined`。

## 根因（初步判断）

`expressions.py` 的 `eval_expr` 用正则 `(?<![\\w.])` + `(?!\\w.])` 做变量名替换。`\\w` 在 Unicode 模式下匹配中文，但不匹配 `-`。所以 `本期.类别-会议` 这个变量名整体不会出现在替换候选里——它含 `-`，被正则的 `(?!\\w.])` 在 `-` 处截断，导致解析时把 `类别-会议` 拆成 `类别` 和 `会议` 两个 token。

变量名长度降序排序也救不了，因为含 `-` 的名字本身就不匹配 `[\\w.]+` 的边界约定。

## 期望

支持字段名含 `-`（以及其它中文业务表中常见的非标识符字符）。两个方向任选：

1. **改变量名识别边界**：把 `-` 也当作变量名合法字符，正则改为 `(?<![\\w.\\-])...(?!\\w.\\-])`。需要确认不会与 `> 0.05`、`!= 主表.电话` 这类负号场景冲突。
2. **允许 YAML 里用引号或反引号包裹变量名**：如 `` `本期.类别-会议` == "变动通报" ``，更显式。

方向 1 改动小但边界复杂；方向 2 显式但增加语法负担。

## 上下文

在真实业务规则（干部岗位变动分析 06 步骤状态机）上做验证时发现。所有字段名都来自业务方既有 Excel 表头，无法回避 `-`。
""",
    },
    {
        "title": "输出格缺少「清空」语义（写入空值），`-` 是「不改」",
        "labels": ["enhancement"],
        "body": """## 背景

当前输出格语义：

| 写法 | 含义 |
|---|---|
| `-` 或留空 | 不改（保持原值）|
| 字面值 / 变量名 / `=表达式` | 写入新值 |

没有「把目标列清空」的写法。

## 需求场景

状态机规则里常见「条件 A 成立则清空该列」。例如 06-目标状态规则 S1：

> 当「类别-会议」列为空时，清空「状态-会议」列。

当前只能用 `=float('nan')` 绕过：

```yaml
则:
  状态-会议: "=float('nan')"
```

但 `float` 内置被 `__builtins__: {}` 屏蔽，会报 `name 'float' is not defined`（见另一条 issue）。

## 期望

新增显式关键词，如 `清空` / `(清空)` / `(空)`：

```yaml
则:
  状态-会议: 清空
```

Excel 决策表输出格写 `清空` 或 `(空)`。

实现上：在 `resolve_output` 里识别该关键词，返回一个 sentinel（如 `CLEAR = object()`），`apply_rules` 里遇到 sentinel 写入 `None`/`NaN`，并在变更明细里显示 `旧值 -> 空`。

## 备注

- 与 `-`(不改) 必须明确区分，不能复用
- 期望清空后 Excel 单元格是真空白（NaN），不是字符串 `"清空"`
- 测试用例里 `期望: {状态-会议: null}` 应能匹配清空后的值
""",
    },
    {
        "title": "命中策略缺少 first-match（按优先级命中第一个即停止）",
        "labels": ["enhancement"],
        "body": """## 背景

当前命中策略只有两种：

| 策略 | 行为 |
|---|---|
| `优先级` | 所有命中规则都参与，每个输出列取优先级数字最小的那条 |
| `唯一` | 任何行命中多条即报错 |

## 需求场景

状态机规则天然是 first-match 语义：**按优先级顺序匹配，命中第一个就停止，后续规则不再评估**。

例如 06-目标状态规则（按优先级排序）：

| 规则 | 优先级 | 条件 | 输出 D |
|---|---|---|---|
| S1 | 1 | 类别为空 | 清空 |
| S2 | 2 | 类别 = "近2个月酝酿及上会计划" | **不处理**（不改 D）|
| S3 | 3 | 类别 = 变动通报 且 D 非空 | 02-进展更新 |
| S4 | 4 | 类别非空 且 D 为空 | 01-新增 |

某行 `类别=近2个月酝酿` 且 `D=空` 时：
- **first-match**：S2 命中即停，D 保持空 ✅ 符合业务预期
- **priority-merge**：S2 没写 D → 跳过；S4 也命中 → D=01-新增 ❌ 错误

## 当前绕过方案

把 S2 的「不处理」显式写成「`=本期.状态-会议`」（覆盖为自身 = 等价不改），让 S2 在 priority-merge 下也「占住」D 列，阻断 S4 覆盖。

```yaml
- 编号: S2
  当:
    - 本期.类别-会议 == "近2个月酝酿及上会计划"
  则:
    状态-会议: "=本期.状态-会议"   # 占位
```

能用，但语义不直观，业务人员看不懂「为什么 S2 输出自己」。

## 期望

新增 `命中策略: 首次命中`（或 `first-match` / `终止`）：

- 按优先级升序评估规则
- 每行命中第一条即停止，后续规则不再评估该行
- 该行输出 = 命中规则的输出（含 `-`/不改 列保持原值）

实现要点：
- `apply_rules` 中按行计算命中：找到第一个 mask=True 的规则即 break
- 性能上比 priority-merge 略好（少评估规则）

## 备注

DMN 标准里这叫 "First Match" hit policy，是和 "Priority"（你现在的 `优先级`）、"Unique"（你的 `唯一`）并列的常见策略。建议命名对齐 DMN 术语，便于业务方理解。
""",
    },
    {
        "title": "YAML 表达式无法调用 Python 内置函数（如 `float()`、`str()`）",
        "labels": ["bug"],
        "body": """## 现象

`expressions.py` 用 `eval(text, {"__builtins__": {}}, local)` 屏蔽所有内置，导致 `float('nan')` / `str(x)` / `int(x)` 等都不可用：

```
[错误] T5: 表达式求值失败: "float('nan')" (name 'float' is not defined)
```

## 复现

```yaml
规则:
  - 编号: R1
    当:
      - 为空(本期.类别)
    则:
      状态: "=float('nan')"
```

## 根因

`eval_expr` 出于安全考虑把 `__builtins__` 清空，只暴露 `FUNCS` 字典里的几个函数。`float` 不在白名单里。

## 期望

两个方向任选：

1. **白名单扩充**：把 `float` / `int` / `str` / `bool` 等纯函数加入 `FUNCS`
2. **提供内置 `空` 常量**：`FUNCS["空"] = float("nan")` 或专门 sentinel，用户写 `=空` 即可写入空值

方向 2 更贴合中文规则文件的整体风格，且能顺带解决「清空语义」issue（在 `resolve_output` 里识别 `空` 关键词）。

## 上下文

发现于「清空 D 列」需求：原规则 S1 要求清空，目前用 `=float('nan')` 绕过，但 `float` 被屏蔽，连绕过都做不到。
""",
    },
]


def main():
    # 先列已有 issue，避免重复
    existing = []
    page = 1
    while True:
        st, data = gh("GET", f"/repos/{OWNER}/{REPO}/issues?state=open&per_page=100&page={page}")
        if st != 200 or not data:
            break
        existing.extend(data)
        if len(data) < 100:
            break
        page += 1
    existing_titles = {i["title"] for i in existing}
    print(f"已存在 open issue: {len(existing)} 条")
    for t in sorted(existing_titles):
        print(f"  - {t}")

    created = []
    for issue in ISSUES:
        if issue["title"] in existing_titles:
            print(f"[跳过] 已存在: {issue['title']}")
            continue
        # 先确保 labels 存在
        for lbl in issue.get("labels", []):
            try:
                gh("POST", f"/repos/{OWNER}/{REPO}/labels",
                   {"name": lbl, "color": "ededed"})
            except Exception:
                pass
        st, data = gh("POST", f"/repos/{OWNER}/{REPO}/issues", issue)
        if st == 201:
            created.append(data)
            print(f"[创建] #{data['number']} {data['title']}")
            print(f"       {data['html_url']}")
        else:
            print(f"[失败] {st} {issue['title']}: {data}")

    print(f"\n本次创建 {len(created)} 条 issue")


if __name__ == "__main__":
    main()
