"""rule-sheet: 把 Excel 比对更新规则写成决策表 / YAML, 由同一个解释器执行。"""

from .engine import RunResult, TestOutcome, apply_rules, build_namespace, run, run_tests
from .loaders import load_excel_rules, load_rules, load_yaml_rules
from .model import Rule, RuleSet, TestCase

__version__ = "0.3.0"

__all__ = [
    "Rule",
    "RuleSet",
    "RunResult",
    "TestCase",
    "TestOutcome",
    "__version__",
    "apply_rules",
    "build_namespace",
    "load_excel_rules",
    "load_rules",
    "load_yaml_rules",
    "run",
    "run_tests",
]
