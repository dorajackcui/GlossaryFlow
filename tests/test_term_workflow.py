import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from term_workflow import (
    check_conflicts,
    check_terms_workbook,
    fill_demand_workbook,
    prefill_terms_workbook,
    split_demand_workbook,
)


def write_rows(path: Path, sheet_name: str, rows: list[list[str]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    for row in rows:
        sheet.append(row)
    workbook.save(path)


def read_sheet(path: Path, sheet_name: str) -> list[tuple[object, ...]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    return list(workbook[sheet_name].iter_rows(values_only=True))


class TermWorkflowTests(unittest.TestCase):
    def test_split_demand_workbook_writes_unique_source_only_terms(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            demand = Path(tmp) / "demand.xlsx"
            output = Path(tmp) / "batch_terms.xlsx"
            write_rows(
                demand,
                "需求",
                [
                    ["source"],
                    ["命运刻面·笼中梦"],
                    ["灵魂洄游·笼中梦"],
                    ["夜息垂芒"],
                    ["命运刻面·希望"],
                ],
            )

            split_demand_workbook(demand, output)

            self.assertEqual(
                read_sheet(output, "主干术语表"),
                [
                    ("source", "target", "status"),
                    ("命运刻面", None, "needs_translation"),
                    ("灵魂洄游", None, "needs_translation"),
                    ("夜息垂芒", None, "needs_translation"),
                ],
            )
            self.assertEqual(
                read_sheet(output, "后缀术语表"),
                [
                    ("source", "target", "status"),
                    ("·笼中梦", None, "needs_translation"),
                    ("·希望", None, "needs_translation"),
                ],
            )

    def test_prefill_terms_workbook_uses_previous_tb_first_value_by_source(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            batch = Path(tmp) / "batch_terms.xlsx"
            previous = Path(tmp) / "previous_tb.xlsx"
            output = Path(tmp) / "prefilled.xlsx"

            workbook = Workbook()
            stem_sheet = workbook.active
            stem_sheet.title = "主干术语表"
            stem_sheet.append(["source", "target", "status"])
            stem_sheet.append(["命运刻面", None, "needs_translation"])
            stem_sheet.append(["夜息垂芒", None, "needs_translation"])
            suffix_sheet = workbook.create_sheet("后缀术语表")
            suffix_sheet.append(["source", "target", "status"])
            suffix_sheet.append(["·笼中梦", None, "needs_translation"])
            workbook.save(batch)

            workbook = Workbook()
            stem_sheet = workbook.active
            stem_sheet.title = "主干术语表"
            stem_sheet.append(["source", "target"])
            stem_sheet.append(["命运刻面", "Coupes du destin"])
            suffix_sheet = workbook.create_sheet("后缀术语表")
            suffix_sheet.append(["source", "target"])
            suffix_sheet.append(["·笼中梦", " : Rêve en cage"])
            workbook.save(previous)

            prefill_terms_workbook(batch, previous, output)

            self.assertEqual(
                read_sheet(output, "主干术语表"),
                [
                    ("source", "target", "status"),
                    ("命运刻面", "Coupes du destin", "matched"),
                    ("夜息垂芒", None, "needs_translation"),
                ],
            )
            self.assertEqual(
                read_sheet(output, "后缀术语表"),
                [
                    ("source", "target", "status"),
                    ("·笼中梦", " : Rêve en cage", "matched"),
                ],
            )

    def test_check_conflicts_reports_source_and_target_conflicts(self) -> None:
        conflicts = check_conflicts(
            {
                "主干术语表": {
                    "命运刻面": "Coupes du destin",
                    "另一个刻面": "Coupes du destin",
                    "夜息垂芒": "Nuit calme",
                },
            },
            {
                "主干术语表": {
                    "命运刻面": "Destin coupe",
                    "灵魂洄游": "Âmes errantes",
                },
            },
        )

        self.assertEqual(len(conflicts), 2)
        self.assertEqual(conflicts[0]["issue_type"], "source_conflict")
        self.assertEqual(conflicts[0]["source"], "命运刻面")
        self.assertEqual(conflicts[1]["issue_type"], "target_conflict")
        self.assertEqual(conflicts[1]["target"], "Coupes du destin")

    def test_check_terms_workbook_reports_duplicate_rows_before_deduping(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            batch = Path(tmp) / "batch_terms.xlsx"
            previous = Path(tmp) / "previous_tb.xlsx"
            report = Path(tmp) / "report.xlsx"

            workbook = Workbook()
            stem_sheet = workbook.active
            stem_sheet.title = "主干术语表"
            stem_sheet.append(["source", "target"])
            stem_sheet.append(["命运刻面", "Coupes du destin"])
            stem_sheet.append(["命运刻面", "Destin coupe"])
            stem_sheet.append(["另一个刻面", "Coupes du destin"])
            workbook.create_sheet("后缀术语表").append(["source", "target"])
            workbook.save(batch)

            workbook = Workbook()
            workbook.active.title = "主干术语表"
            workbook.active.append(["source", "target"])
            workbook.create_sheet("后缀术语表").append(["source", "target"])
            workbook.save(previous)

            conflict_count = check_terms_workbook(batch, previous, report)
            report_rows = read_sheet(report, "检查报告")

            self.assertEqual(conflict_count, 2)
            self.assertEqual(report_rows[1][0], "source_conflict")
            self.assertEqual(report_rows[2][0], "target_conflict")

    def test_fill_demand_workbook_combines_stem_and_suffix_targets(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            demand = Path(tmp) / "demand.xlsx"
            terms = Path(tmp) / "translated_terms.xlsx"
            output = Path(tmp) / "filled.xlsx"

            write_rows(
                demand,
                "需求",
                [
                    ["source"],
                    ["命运刻面·笼中梦"],
                    ["夜息垂芒"],
                    ["未知·笼中梦"],
                ],
            )

            workbook = Workbook()
            stem_sheet = workbook.active
            stem_sheet.title = "主干术语表"
            stem_sheet.append(["source", "target"])
            stem_sheet.append(["命运刻面", "Coupes du destin"])
            stem_sheet.append(["夜息垂芒", "Nuit berçante"])
            suffix_sheet = workbook.create_sheet("后缀术语表")
            suffix_sheet.append(["source", "target"])
            suffix_sheet.append(["·笼中梦", " : Rêve en cage"])
            workbook.save(terms)

            fill_demand_workbook(demand, terms, output)

            self.assertEqual(
                read_sheet(output, "需求"),
                [
                    ("source", "target"),
                    ("命运刻面·笼中梦", "Coupes du destin : Rêve en cage"),
                    ("夜息垂芒", "Nuit berçante"),
                    ("未知·笼中梦", None),
                ],
            )


if __name__ == "__main__":
    unittest.main()
