"""命令行入口。

rule-sheet run  <规则文件> [--base-dir DIR] [--output PATH] [--ignore-tests] [--dry-run]
rule-sheet test <规则文件>
rule-sheet <规则文件>            等价于 run
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .engine import RunResult, TestOutcome, run, run_tests
from .loaders import load_rules

SUBCOMMANDS = {"run", "test"}


def _print_tests(outcomes: list[TestOutcome]) -> None:
    for o in outcomes:
        if o.error:
            print(f"  [错误] {o.name}: {o.error}")
        elif o.passed:
            print(f"  [通过] {o.name} (命中 {o.hit or '无'})")
        else:
            print(f"  [失败] {o.name} (命中 {o.hit or '无'}): " + "; ".join(o.problems))
    passed = sum(1 for o in outcomes if o.passed)
    print(f"测试结果: {passed}/{len(outcomes)} 通过")


def _print_run(res: RunResult) -> None:
    rs = res.ruleset
    print(f"规则文件: {rs.source}   规则数: {len(rs.rules)}   命中策略: {rs.hit_policy}")
    print(f"输出列: {', '.join(rs.outputs)}")
    if res.tests:
        print(f"\n== 先跑测试用例 ({len(res.tests)} 条) ==")
        _print_tests(res.tests)
    else:
        print("提示: 规则文件没有内嵌测试用例, 直接处理真实数据")
    if res.stopped_by_tests:
        print("有测试未通过, 停止处理真实数据 (加 --ignore-tests 可强制继续)")
        return
    assert res.out is not None and res.changed is not None
    print(f"\n== 处理真实数据 ==\n左表 {len(res.out)} 行, 匹配键 {res.keys}")
    for note in res.notes:
        print(f"提示: {note}")
    show_cols = res.keys + ["命中规则", "变更明细"]
    print(f"\n{len(res.changed)} 行发生变更:")
    print(res.changed[show_cols].to_string(index=False) if not res.changed.empty else "(无)")
    if res.written:
        print(f"\n已写出: {res.written}")
    else:
        print("\n--dry-run: 未写文件")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="rule-sheet", description="把 Excel 比对更新规则写成决策表 / YAML 并执行")
    parser.add_argument("--version", action="version", version=f"rule-sheet {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    p_run = sub.add_parser("run", help="跑内嵌测试, 通过后处理真实数据并写出结果")
    p_run.add_argument("rules", type=Path, help="规则文件: .xlsx 决策表 或 .yaml")
    p_run.add_argument("--base-dir", type=Path, help="规则文件内相对路径的基准目录 (默认: 规则文件所在目录)")
    p_run.add_argument("--output", type=Path, help="覆盖规则文件里的输出路径")
    p_run.add_argument("--ignore-tests", action="store_true", help="测试失败也继续")
    p_run.add_argument("--dry-run", action="store_true", help="只打印, 不写文件")

    p_test = sub.add_parser("test", help="只跑规则文件内嵌的测试用例")
    p_test.add_argument("rules", type=Path, help="规则文件: .xlsx 决策表 或 .yaml")
    return parser


def main(argv: list[str] | None = None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[union-attr]
    args = list(sys.argv[1:] if argv is None else argv)
    if args and args[0] not in SUBCOMMANDS and not args[0].startswith("-"):
        args.insert(0, "run")  # rule-sheet <规则文件> 等价于 run
    ns = build_parser().parse_args(args)
    try:
        if ns.command == "test":
            rs = load_rules(ns.rules)
            if not rs.tests:
                print("规则文件没有测试用例 (Excel 加「测试」sheet / YAML 加 测试:), 什么都没验证")
                return 1
            outcomes = run_tests(rs)
            _print_tests(outcomes)
            return 0 if all(o.passed for o in outcomes) else 1
        res = run(
            ns.rules,
            base_dir=ns.base_dir,
            ignore_tests=ns.ignore_tests,
            dry_run=ns.dry_run,
            output=ns.output,
        )
        _print_run(res)
        return 1 if res.stopped_by_tests else 0
    except (ValueError, KeyError, FileNotFoundError, PermissionError) as exc:
        print(f"\n错误: {exc.args[0] if exc.args else exc}")
        return 2


if __name__ == "__main__":
    sys.exit(main())
