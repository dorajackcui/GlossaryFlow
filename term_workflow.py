#!/usr/bin/env python3
"""Workflow for splitting, pre-filling, checking, and filling term Excel files."""

from __future__ import annotations

import argparse
import re
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook, load_workbook


STEM_SHEET = "主干术语表"
SUFFIX_SHEET = "后缀术语表"
REPORT_SHEET = "检查报告"
SOURCE_SUFFIX_MARKER = "·"
TARGET_SUFFIX_RE = re.compile(r"^(?P<stem>.*?)(?P<suffix>\s*[:：].*)$")


def clean_cell(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\u00a0", " ").strip()


def clean_target_cell(value: object) -> str:
    text = "" if value is None else str(value).replace("\u00a0", " ")
    stripped = text.strip()
    if stripped.startswith((":", "：")):
        return f" {stripped}"
    return stripped


def normalize_header(value: object) -> str:
    return clean_cell(value).lower().replace(" ", "").replace("_", "")


def split_source(value: str) -> tuple[str, str]:
    if SOURCE_SUFFIX_MARKER not in value:
        return value.strip(), ""

    stem, suffix = value.split(SOURCE_SUFFIX_MARKER, 1)
    return stem.strip(), f"{SOURCE_SUFFIX_MARKER}{suffix.strip()}"


def split_target(value: str) -> tuple[str, str]:
    if not value:
        return "", ""

    match = TARGET_SUFFIX_RE.match(value)
    if not match:
        return value.strip(), ""

    return match.group("stem").strip(), match.group("suffix")


def first_value_wins_add(glossary: dict[str, str], source: str, target: str = "") -> None:
    if source and source not in glossary:
        glossary[source] = target


def find_header_columns(rows: list[tuple[object, ...]]) -> tuple[int, int, int | None]:
    if not rows:
        raise ValueError("Input sheet is empty.")

    header = [normalize_header(cell) for cell in rows[0]]
    source_names = {"source", "src", "源文", "原文", "源语言", "源术语"}
    target_names = {"target", "tgt", "译文", "目标", "目标语言", "目标术语"}

    source_col = next((index for index, name in enumerate(header) if name in source_names), None)
    target_col = next((index for index, name in enumerate(header) if name in target_names), None)
    if source_col is not None:
        return 1, source_col, target_col

    return 0, 0, 1 if len(rows[0]) > 1 else None


def find_sheet_with_source(workbook) -> str:
    for sheet in workbook.worksheets:
        first_row = list(sheet.iter_rows(values_only=True, max_row=1))
        if not first_row:
            continue
        headers = {normalize_header(cell) for cell in first_row[0]}
        if headers & {"source", "src", "源文", "原文", "源语言", "源术语"}:
            return sheet.title
    return workbook.worksheets[0].title


def read_source_target_rows(path: Path, sheet_name: str | None = None) -> tuple[str, list[tuple[str, str]]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    actual_sheet = sheet_name or find_sheet_with_source(workbook)
    if actual_sheet not in workbook.sheetnames:
        raise ValueError(f"Sheet {actual_sheet!r} not found in {path}.")

    rows = list(workbook[actual_sheet].iter_rows(values_only=True))
    start_row, source_col, target_col = find_header_columns(rows)
    result: list[tuple[str, str]] = []
    for row in rows[start_row:]:
        source = clean_cell(row[source_col] if source_col < len(row) else "")
        target = clean_cell(row[target_col] if target_col is not None and target_col < len(row) else "")
        if source:
            result.append((source, target))
    return actual_sheet, result


def split_terms_from_rows(rows: list[tuple[str, str]]) -> tuple[dict[str, str], dict[str, str]]:
    stems: dict[str, str] = {}
    suffixes: dict[str, str] = {}

    for source, target in rows:
        source_stem, source_suffix = split_source(source)
        target_stem, target_suffix = split_target(target)

        first_value_wins_add(stems, source_stem, target_stem)
        if source_suffix:
            first_value_wins_add(suffixes, source_suffix, target_suffix)

    return stems, suffixes


def write_terms_workbook(
    output_path: Path,
    stems: dict[str, str],
    suffixes: dict[str, str],
    include_status: bool = True,
) -> None:
    workbook = Workbook()
    stem_sheet = workbook.active
    stem_sheet.title = STEM_SHEET
    suffix_sheet = workbook.create_sheet(SUFFIX_SHEET)

    for sheet, glossary in ((stem_sheet, stems), (suffix_sheet, suffixes)):
        sheet.append(["source", "target", "status"] if include_status else ["source", "target"])
        for source, target in glossary.items():
            if include_status:
                sheet.append([source, target or None, "matched" if target else "needs_translation"])
            else:
                sheet.append([source, target or None])
        sheet.freeze_panes = "A2"
        sheet.column_dimensions["A"].width = 32
        sheet.column_dimensions["B"].width = 42
        if include_status:
            sheet.column_dimensions["C"].width = 20

    workbook.save(output_path)


def split_demand_workbook(input_path: Path, output_path: Path, sheet_name: str | None = None) -> None:
    _, rows = read_source_target_rows(input_path, sheet_name)
    stems, suffixes = split_terms_from_rows([(source, "") for source, _ in rows])
    write_terms_workbook(output_path, stems, suffixes, include_status=True)


def read_terms_workbook(path: Path) -> dict[str, dict[str, str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    terms: dict[str, dict[str, str]] = {}

    for sheet_name in (STEM_SHEET, SUFFIX_SHEET):
        if sheet_name not in workbook.sheetnames:
            terms[sheet_name] = {}
            continue

        rows = list(workbook[sheet_name].iter_rows(values_only=True))
        if not rows:
            terms[sheet_name] = {}
            continue

        start_row, source_col, target_col = find_header_columns(rows)
        sheet_terms: dict[str, str] = {}
        for row in rows[start_row:]:
            source = clean_cell(row[source_col] if source_col < len(row) else "")
            target = clean_target_cell(row[target_col] if target_col is not None and target_col < len(row) else "")
            first_value_wins_add(sheet_terms, source, target)
        terms[sheet_name] = sheet_terms

    return terms


def read_terms_entries(path: Path) -> dict[str, list[tuple[str, str]]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    entries: dict[str, list[tuple[str, str]]] = {}

    for sheet_name in (STEM_SHEET, SUFFIX_SHEET):
        if sheet_name not in workbook.sheetnames:
            entries[sheet_name] = []
            continue

        rows = list(workbook[sheet_name].iter_rows(values_only=True))
        if not rows:
            entries[sheet_name] = []
            continue

        start_row, source_col, target_col = find_header_columns(rows)
        sheet_entries: list[tuple[str, str]] = []
        for row in rows[start_row:]:
            source = clean_cell(row[source_col] if source_col < len(row) else "")
            target = clean_target_cell(row[target_col] if target_col is not None and target_col < len(row) else "")
            if source:
                sheet_entries.append((source, target))
        entries[sheet_name] = sheet_entries

    return entries


def prefill_terms_workbook(batch_path: Path, previous_tb_path: Path, output_path: Path) -> None:
    batch_terms = read_terms_workbook(batch_path)
    previous_terms = read_terms_workbook(previous_tb_path)

    prefilled: dict[str, dict[str, str]] = {}
    for sheet_name in (STEM_SHEET, SUFFIX_SHEET):
        prefilled[sheet_name] = {}
        for source, current_target in batch_terms[sheet_name].items():
            target = current_target or previous_terms[sheet_name].get(source, "")
            first_value_wins_add(prefilled[sheet_name], source, target)

    write_terms_workbook(output_path, prefilled[STEM_SHEET], prefilled[SUFFIX_SHEET], include_status=True)


def check_conflicts(
    batch_terms: dict[str, dict[str, str]],
    previous_terms: dict[str, dict[str, str]],
) -> list[dict[str, str]]:
    batch_entries = {
        sheet_name: list(terms.items())
        for sheet_name, terms in batch_terms.items()
    }
    previous_entries = {
        sheet_name: list(terms.items())
        for sheet_name, terms in previous_terms.items()
    }
    return check_conflict_entries(batch_entries, previous_entries)


def check_conflict_entries(
    batch_entries: dict[str, list[tuple[str, str]]],
    previous_entries: dict[str, list[tuple[str, str]]],
) -> list[dict[str, str]]:
    conflicts: list[dict[str, str]] = []

    for sheet_name in (STEM_SHEET, SUFFIX_SHEET):
        by_source: dict[str, set[str]] = defaultdict(set)
        by_target: dict[str, set[str]] = defaultdict(set)

        for term_entries in (previous_entries.get(sheet_name, []), batch_entries.get(sheet_name, [])):
            for source, target in term_entries:
                source = clean_cell(source)
                target = clean_target_cell(target)
                if not source or not target:
                    continue
                by_source[source].add(target)
                by_target[target].add(source)

        for source, targets in by_source.items():
            if len(targets) > 1:
                conflicts.append(
                    {
                        "issue_type": "source_conflict",
                        "term_type": sheet_name,
                        "source": source,
                        "target": " | ".join(sorted(targets)),
                        "details": "Same source has multiple targets.",
                    }
                )

        for target, sources in by_target.items():
            if len(sources) > 1:
                conflicts.append(
                    {
                        "issue_type": "target_conflict",
                        "term_type": sheet_name,
                        "source": " | ".join(sorted(sources)),
                        "target": target,
                        "details": "Same target is used by multiple sources.",
                    }
                )

    return conflicts


def write_check_report(output_path: Path, conflicts: list[dict[str, str]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = REPORT_SHEET
    sheet.append(["issue_type", "term_type", "source", "target", "details"])
    for conflict in conflicts:
        sheet.append(
            [
                conflict["issue_type"],
                conflict["term_type"],
                conflict["source"],
                conflict["target"],
                conflict["details"],
            ]
        )
    for column, width in {"A": 20, "B": 16, "C": 44, "D": 44, "E": 36}.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    workbook.save(output_path)


def check_terms_workbook(batch_path: Path, previous_tb_path: Path, output_path: Path) -> int:
    conflicts = check_conflict_entries(read_terms_entries(batch_path), read_terms_entries(previous_tb_path))
    write_check_report(output_path, conflicts)
    return len(conflicts)


def compose_target(source: str, terms: dict[str, dict[str, str]]) -> str:
    source_stem, source_suffix = split_source(source)
    stem_target = terms[STEM_SHEET].get(source_stem, "")
    if not stem_target:
        return ""
    if not source_suffix:
        return stem_target

    suffix_target = terms[SUFFIX_SHEET].get(source_suffix, "")
    if not suffix_target:
        return ""
    return f"{stem_target}{suffix_target}"


def fill_demand_workbook(
    demand_path: Path,
    terms_path: Path,
    output_path: Path,
    sheet_name: str | None = None,
) -> None:
    workbook = load_workbook(demand_path)
    actual_sheet = sheet_name or find_sheet_with_source(workbook)
    if actual_sheet not in workbook.sheetnames:
        raise ValueError(f"Sheet {actual_sheet!r} not found in {demand_path}.")

    sheet = workbook[actual_sheet]
    rows = list(sheet.iter_rows(values_only=True))
    start_row, source_col, target_col = find_header_columns(rows)
    if target_col is None:
        target_col = sheet.max_column
        sheet.cell(row=1, column=target_col + 1, value="target")

    terms = read_terms_workbook(terms_path)
    for row_index in range(start_row + 1, sheet.max_row + 1):
        source = clean_cell(sheet.cell(row=row_index, column=source_col + 1).value)
        if not source:
            continue
        target = compose_target(source, terms)
        sheet.cell(row=row_index, column=target_col + 1, value=target or None)

    workbook.save(output_path)


def default_output_path(input_path: Path, suffix: str) -> Path:
    return input_path.with_name(f"{input_path.stem}_{suffix}.xlsx")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the source/target term workflow.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    split_parser = subparsers.add_parser("split", help="Split demand Excel into batch stem/suffix term sheets.")
    split_parser.add_argument("demand", type=Path)
    split_parser.add_argument("--sheet")
    split_parser.add_argument("--output", type=Path)

    prefill_parser = subparsers.add_parser("prefill", help="Fill batch terms from previous TB where possible.")
    prefill_parser.add_argument("batch_terms", type=Path)
    prefill_parser.add_argument("previous_tb", type=Path)
    prefill_parser.add_argument("--output", type=Path)

    check_parser = subparsers.add_parser("check", help="Check source/target conflicts across batch terms and previous TB.")
    check_parser.add_argument("batch_terms", type=Path)
    check_parser.add_argument("previous_tb", type=Path)
    check_parser.add_argument("--output", type=Path)

    fill_parser = subparsers.add_parser("fill", help="Fill demand Excel target column from translated terms.")
    fill_parser.add_argument("demand", type=Path)
    fill_parser.add_argument("translated_terms", type=Path)
    fill_parser.add_argument("--sheet")
    fill_parser.add_argument("--output", type=Path)

    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.command == "split":
        output_path = args.output or default_output_path(args.demand, "batch_terms")
        split_demand_workbook(args.demand, output_path, args.sheet)
        print(f"Output: {output_path}")
        return

    if args.command == "prefill":
        output_path = args.output or default_output_path(args.batch_terms, "prefilled")
        prefill_terms_workbook(args.batch_terms, args.previous_tb, output_path)
        print(f"Output: {output_path}")
        return

    if args.command == "check":
        output_path = args.output or default_output_path(args.batch_terms, "check_report")
        count = check_terms_workbook(args.batch_terms, args.previous_tb, output_path)
        print(f"Conflicts: {count}")
        print(f"Output: {output_path}")
        return

    if args.command == "fill":
        output_path = args.output or default_output_path(args.demand, "filled")
        fill_demand_workbook(args.demand, args.translated_terms, output_path, args.sheet)
        print(f"Output: {output_path}")


if __name__ == "__main__":
    main()
