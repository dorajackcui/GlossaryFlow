import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from term_assistant import run_check, run_demo, run_fill, run_new_batch


def write_demand(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "需求"
    sheet.append(["source"])
    sheet.append(["命运刻面·笼中梦"])
    sheet.append(["夜息垂芒"])
    workbook.save(path)


def write_tb(path: Path) -> None:
    workbook = Workbook()
    stem_sheet = workbook.active
    stem_sheet.title = "主干术语表"
    stem_sheet.append(["source", "target"])
    stem_sheet.append(["命运刻面", "Coupes du destin"])
    stem_sheet.append(["夜息垂芒", "Nuit berçante"])
    suffix_sheet = workbook.create_sheet("后缀术语表")
    suffix_sheet.append(["source", "target"])
    suffix_sheet.append(["·笼中梦", " : Rêve en cage"])
    workbook.save(path)


def read_rows(path: Path, sheet_name: str) -> list[tuple[object, ...]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    return list(workbook[sheet_name].iter_rows(values_only=True))


class TermAssistantTests(unittest.TestCase):
    def test_run_new_batch_creates_split_and_prefilled_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            demand = tmp_path / "需求.xlsx"
            tb = tmp_path / "过往TB.xlsx"
            output_root = tmp_path / "outputs"
            write_demand(demand)
            write_tb(tb)

            result = run_new_batch(demand, tb, output_root)

            self.assertTrue(result.batch_terms.exists())
            self.assertTrue(result.prefilled_terms.exists())
            self.assertIn("请打开", result.next_step)
            rows = read_rows(result.prefilled_terms, "主干术语表")
            self.assertEqual(rows[1], ("命运刻面", "Coupes du destin", "matched"))

    def test_run_new_batch_defaults_outputs_next_to_demand_excel(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            demand_dir = tmp_path / "需求文件夹"
            demand_dir.mkdir()
            demand = demand_dir / "需求.xlsx"
            tb = tmp_path / "过往TB.xlsx"
            write_demand(demand)
            write_tb(tb)

            result = run_new_batch(demand, tb)

            self.assertEqual(result.batch_dir.parent.resolve(), (demand_dir / "outputs").resolve())
            self.assertTrue(result.prefilled_terms.exists())

    def test_run_check_creates_report(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            demand = tmp_path / "需求.xlsx"
            tb = tmp_path / "过往TB.xlsx"
            output_root = tmp_path / "outputs"
            write_demand(demand)
            write_tb(tb)
            batch = run_new_batch(demand, tb, output_root)

            result = run_check(batch.prefilled_terms, tb, output_root)

            self.assertEqual(result.conflict_count, 0)
            self.assertTrue(result.report.exists())
            self.assertEqual(len(read_rows(result.report, "检查报告")), 1)

    def test_run_fill_creates_filled_demand(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            demand = tmp_path / "需求.xlsx"
            tb = tmp_path / "过往TB.xlsx"
            output_root = tmp_path / "outputs"
            write_demand(demand)
            write_tb(tb)
            batch = run_new_batch(demand, tb, output_root)

            result = run_fill(demand, batch.prefilled_terms, output_root)

            self.assertTrue(result.filled_demand.exists())
            rows = read_rows(result.filled_demand, "需求")
            self.assertEqual(rows[1], ("命运刻面·笼中梦", "Coupes du destin : Rêve en cage"))

    def test_run_demo_creates_all_demo_files(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            demand = tmp_path / "test.xlsx"
            tb = tmp_path / "test_glossaries.xlsx"
            output_root = tmp_path / "outputs"
            write_demand(demand)
            write_tb(tb)

            result = run_demo(demand, tb, output_root)

            self.assertTrue(result.batch_terms.exists())
            self.assertTrue(result.prefilled_terms.exists())
            self.assertTrue(result.report.exists())
            self.assertTrue(result.filled_demand.exists())


if __name__ == "__main__":
    unittest.main()
