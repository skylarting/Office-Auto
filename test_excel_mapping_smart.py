import os
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook, load_workbook
from PySide6.QtWidgets import QApplication

from excel_mapping_smart import (
    SmartMappingWindow, build_plan, candidate_matches, find_plan_files,
    rebind_plan,
)
from smart_template import execute_smart_plan, load_plan, save_plan


def make_book(path: Path, sheet_name: str, shifted: bool = False) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = sheet_name
    ro, co = (2, 1) if shifted else (0, 0)
    sheet.cell(1 + ro, 2 + co, "余额")
    sheet.cell(2 + ro, 2 + co, "人民币")
    sheet.cell(2 + ro, 3 + co, "外币")
    sheet.cell(3 + ro, 1 + co, "Ⅰ.资产")
    sheet.cell(4 + ro, 1 + co, "1.现金")
    sheet.cell(4 + ro, 2 + co, 10)
    sheet.cell(4 + ro, 3 + co, 20)
    sheet.cell(5 + ro, 1 + co, "2.存款")
    sheet.cell(5 + ro, 2 + co, 30)
    sheet.cell(5 + ro, 3 + co, "=B4+C4")
    book.save(path)


class SmartWizardTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_window_starts_on_beginner_home(self):
        window = SmartMappingWindow()
        self.assertIs(window.stack.currentWidget(), window.home)
        self.assertIn("智能报表助手", window.windowTitle())
        window.close()

    def test_pending_cells_are_grouped_by_sheet_pair(self):
        window = SmartMappingWindow()
        from smart_template import SmartMatch, SmartTemplatePlan
        window.plan = SmartTemplatePlan(matches=[
            SmartMatch("a.xlsx", "数据", "A1", (), (), "b.xlsx", "报表", "B1", (), (), .7, "待确认"),
            SmartMatch("a.xlsx", "数据", "A2", (), (), "b.xlsx", "报表", "B2", (), (), .7, "待确认"),
            SmartMatch("c.xlsx", "其他", "A1", (), (), "b.xlsx", "其他", "B1", (), (), .7, "待确认"),
        ])
        groups = window._pending_groups()
        self.assertEqual(len(groups), 2)
        self.assertEqual(sorted(map(len, groups)), [1, 2])
        window.close()

    def test_learning_plan_can_save_load_and_execute_copy(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "source.xlsx"
            target = root / "target.xlsx"
            make_book(source, "数据")
            make_book(target, "报表", shifted=True)
            plan = build_plan([source], target)
            self.assertTrue(plan.pairs)
            self.assertGreater(candidate_matches(plan)[0], 0)
            scheme = root / "月报.smartmap.json"
            save_plan(scheme, plan)
            self.assertEqual(find_plan_files(root), [scheme])
            loaded = load_plan(scheme)
            outputs, _ = execute_smart_plan(loaded, root / "映射结果")
            self.assertEqual(len(outputs), 1)
            result = load_workbook(outputs[0], data_only=False)
            try:
                self.assertEqual(result["报表"]["C6"].value, 10)
                self.assertEqual(result["报表"]["D7"].value, "=B4+C4")
            finally:
                result.close()

    def test_monthly_rebind_uses_current_folder(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            old_folder = root / "old"; new_folder = root / "new"
            old_folder.mkdir(); new_folder.mkdir()
            old_source = old_folder / "支行数据202605.xlsx"
            new_source = new_folder / "支行数据202606.xlsx"
            old_target = root / "历史报表.xlsx"
            new_target = root / "本月模板.xlsx"
            make_book(old_source, "数据")
            make_book(new_source, "数据")
            make_book(old_target, "报表", shifted=True)
            make_book(new_target, "报表", shifted=True)
            saved = build_plan([old_source], old_target)
            rebound = rebind_plan(saved, new_folder, new_target)
            self.assertEqual(Path(rebound.pairs[0].source_file), new_source.resolve())
            self.assertEqual(Path(rebound.target_file), new_target.resolve())


if __name__ == "__main__":
    unittest.main()
