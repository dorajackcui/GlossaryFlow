#!/usr/bin/env python3
"""Friendly terminal assistant for the term workflow."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from term_workflow import (
    check_terms_workbook,
    fill_demand_workbook,
    prefill_terms_workbook,
    split_demand_workbook,
)


OUTPUT_DIR_NAME = "outputs"


@dataclass(frozen=True)
class NewBatchResult:
    batch_dir: Path
    batch_terms: Path
    prefilled_terms: Path
    next_step: str


@dataclass(frozen=True)
class CheckResult:
    batch_dir: Path
    report: Path
    conflict_count: int
    next_step: str


@dataclass(frozen=True)
class FillResult:
    batch_dir: Path
    filled_demand: Path
    next_step: str


@dataclass(frozen=True)
class DemoResult:
    batch_dir: Path
    batch_terms: Path
    prefilled_terms: Path
    report: Path
    filled_demand: Path
    conflict_count: int
    next_step: str


def safe_stem(path: Path) -> str:
    return path.stem.replace(" ", "_")


def default_output_root_for_demand(demand_path: Path) -> Path:
    return demand_path.resolve().parent / OUTPUT_DIR_NAME


def make_batch_dir(anchor_path: Path, output_root: Path) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    batch_dir = output_root / f"{timestamp}_{safe_stem(anchor_path)}"
    batch_dir.mkdir(parents=True, exist_ok=False)
    return batch_dir


def sibling_batch_dir(input_path: Path, output_root: Path | None = None) -> Path:
    if output_root is not None:
        output_root.mkdir(parents=True, exist_ok=True)
        resolved = input_path.resolve()
        for candidate in output_root.iterdir():
            if candidate.is_dir() and resolved in {item.resolve() for item in candidate.glob("*.xlsx")}:
                return candidate
        return make_batch_dir(input_path, output_root)

    if input_path.parent.name:
        return input_path.resolve().parent

    return default_output_root_for_demand(input_path)


def ensure_xlsx(path: Path, label: str) -> Path:
    if not path.exists():
        raise FileNotFoundError(f"{label} 不存在：{path}")
    if path.suffix.lower() != ".xlsx":
        raise ValueError(f"{label} 需要是 .xlsx 文件：{path}")
    return path


def run_new_batch(
    demand_path: Path,
    previous_tb_path: Path,
    output_root: Path | None = None,
) -> NewBatchResult:
    demand_path = ensure_xlsx(demand_path, "需求 Excel")
    previous_tb_path = ensure_xlsx(previous_tb_path, "过往 TB")
    output_root = output_root or default_output_root_for_demand(demand_path)
    batch_dir = make_batch_dir(demand_path, output_root)

    batch_terms = batch_dir / "01_本批次术语表.xlsx"
    prefilled_terms = batch_dir / "02_本批次术语表_预填.xlsx"

    split_demand_workbook(demand_path, batch_terms)
    prefill_terms_workbook(batch_terms, previous_tb_path, prefilled_terms)

    return NewBatchResult(
        batch_dir=batch_dir,
        batch_terms=batch_terms,
        prefilled_terms=prefilled_terms,
        next_step=(
            f"请打开 {prefilled_terms}，把 status = needs_translation 的 target 补上。"
            "补完后另存为 02_本批次术语表_已翻译.xlsx，再回到这里选择 2。"
        ),
    )


def run_check(
    translated_terms_path: Path,
    previous_tb_path: Path,
    output_root: Path | None = None,
) -> CheckResult:
    translated_terms_path = ensure_xlsx(translated_terms_path, "已翻译术语表")
    previous_tb_path = ensure_xlsx(previous_tb_path, "过往 TB")
    batch_dir = sibling_batch_dir(translated_terms_path, output_root)
    report = batch_dir / "03_检查报告.xlsx"

    conflict_count = check_terms_workbook(translated_terms_path, previous_tb_path, report)
    if conflict_count:
        next_step = f"发现 {conflict_count} 个冲突。请打开 {report}，修正术语表后重新选择 2。"
    else:
        next_step = f"检查通过。下一步选择 3，使用 {translated_terms_path} 回填需求 Excel。"

    return CheckResult(
        batch_dir=batch_dir,
        report=report,
        conflict_count=conflict_count,
        next_step=next_step,
    )


def run_fill(
    demand_path: Path,
    translated_terms_path: Path,
    output_root: Path | None = None,
) -> FillResult:
    demand_path = ensure_xlsx(demand_path, "需求 Excel")
    translated_terms_path = ensure_xlsx(translated_terms_path, "已翻译术语表")
    batch_dir = sibling_batch_dir(translated_terms_path, output_root)
    filled_demand = batch_dir / "04_需求_已回填.xlsx"

    fill_demand_workbook(demand_path, translated_terms_path, filled_demand)
    return FillResult(
        batch_dir=batch_dir,
        filled_demand=filled_demand,
        next_step=f"完成。请打开 {filled_demand} 查看最终 target。",
    )


def run_demo(
    demand_path: Path = Path("test.xlsx"),
    previous_tb_path: Path = Path("test_glossaries.xlsx"),
    output_root: Path | None = None,
) -> DemoResult:
    batch = run_new_batch(demand_path, previous_tb_path, output_root)
    check = run_check(batch.prefilled_terms, previous_tb_path, output_root)
    fill = run_fill(demand_path, batch.prefilled_terms, output_root)
    return DemoResult(
        batch_dir=batch.batch_dir,
        batch_terms=batch.batch_terms,
        prefilled_terms=batch.prefilled_terms,
        report=check.report,
        filled_demand=fill.filled_demand,
        conflict_count=check.conflict_count,
        next_step=f"演示完成。请打开文件夹 {batch.batch_dir} 查看 01-04 四个文件。",
    )


def ask_path(prompt: str) -> Path:
    return Path(input(prompt).strip().strip('"').strip("'"))


def print_header() -> None:
    print()
    print("术语处理小助手")
    print("=" * 28)
    print("1. 开始新批次：拆分需求 Excel + 用过往 TB 预填")
    print("2. 我已经翻译完了：检查冲突")
    print("3. 检查通过了：回填需求 Excel")
    print("4. 用 test.xlsx 跑一遍演示")
    print("q. 退出")
    print()


def print_result(title: str, lines: list[str]) -> None:
    print()
    print(title)
    print("-" * 28)
    for line in lines:
        print(line)
    print()


def interactive_main() -> None:
    while True:
        print_header()
        choice = input("请选择：").strip().lower()

        try:
            if choice == "1":
                demand = ask_path("请输入需求 Excel 路径：")
                tb = ask_path("请输入过往 TB 路径：")
                result = run_new_batch(demand, tb)
                print_result(
                    "新批次已生成",
                    [
                        f"输出文件夹：{result.batch_dir}",
                        f"拆分术语表：{result.batch_terms}",
                        f"预填术语表：{result.prefilled_terms}",
                        f"下一步：{result.next_step}",
                    ],
                )
            elif choice == "2":
                translated = ask_path("请输入已翻译术语表路径：")
                tb = ask_path("请输入过往 TB 路径：")
                result = run_check(translated, tb)
                print_result(
                    "检查完成",
                    [
                        f"检查报告：{result.report}",
                        f"冲突数量：{result.conflict_count}",
                        f"下一步：{result.next_step}",
                    ],
                )
            elif choice == "3":
                demand = ask_path("请输入需求 Excel 路径：")
                translated = ask_path("请输入已翻译术语表路径：")
                result = run_fill(demand, translated)
                print_result(
                    "回填完成",
                    [
                        f"最终文件：{result.filled_demand}",
                        f"下一步：{result.next_step}",
                    ],
                )
            elif choice == "4":
                result = run_demo()
                print_result(
                    "演示完成",
                    [
                        f"输出文件夹：{result.batch_dir}",
                        f"01：{result.batch_terms}",
                        f"02：{result.prefilled_terms}",
                        f"03：{result.report}",
                        f"04：{result.filled_demand}",
                        f"冲突数量：{result.conflict_count}",
                        f"下一步：{result.next_step}",
                    ],
                )
            elif choice in {"q", "quit", "exit"}:
                print("已退出。")
                return
            else:
                print("没看懂这个选项，请输入 1、2、3、4 或 q。")
        except Exception as error:
            print()
            print(f"出错了：{error}")
            print("请检查文件路径和 Excel 格式后再试。")


if __name__ == "__main__":
    interactive_main()
