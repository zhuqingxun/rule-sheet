from pathlib import Path

import pandas as pd
import pytest

EXAMPLES = Path(__file__).resolve().parents[1] / "examples"
EXAMPLE_DIR = EXAMPLES / "customer-sync"
EXCEL_RULES = EXAMPLE_DIR / "规则_决策表.xlsx"
YAML_RULES = EXAMPLE_DIR / "rules.yaml"
STATUS_DIR = EXAMPLES / "status-machine"
STATUS_EXCEL = STATUS_DIR / "规则_决策表.xlsx"
STATUS_YAML = STATUS_DIR / "rules.yaml"
TAGS_DIR = EXAMPLES / "position-tags"
TAGS_EXCEL = TAGS_DIR / "规则_决策表.xlsx"
TAGS_YAML = TAGS_DIR / "rules.yaml"
CARRY_DIR = EXAMPLES / "carry-forward"
CARRY_EXCEL = CARRY_DIR / "规则_决策表.xlsx"
CARRY_YAML = CARRY_DIR / "rules.yaml"


@pytest.fixture
def ns_two_rows():
    """两行的小命名空间, 覆盖字符串 / 数值 / 空值 / 布尔。"""
    idx = pd.RangeIndex(3)
    return {
        "主表.状态": pd.Series(["活跃", "休眠", None], index=idx, dtype=object),
        "主表.金额": pd.Series([10000.0, 0.0, 5000.0], index=idx),
        "本月.金额": pd.Series([10800.0, float("nan"), 5100.0], index=idx),
        "金额": pd.Series([1.0, 2.0, 3.0], index=idx),
        "本月.存在": pd.Series([True, False, True], index=idx),
        "主表.电话": pd.Series(["138", None, "139"], index=idx, dtype=object),
        "本月.电话": pd.Series(["139", "138", None], index=idx, dtype=object),
        "基准日期": pd.Timestamp("2026-10-19"),
        "主表.上次活跃": pd.Series(pd.to_datetime(["2026-10-09", "2026-01-01", None]), index=idx),
    }
