import os
import subprocess
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from openpyxl import Workbook
from PySide6.QtWidgets import QApplication

from excel_mapping_studio import (
    StudioWindow, WorksheetMapping, mapping_rules_from_rows,
    preserve_existing_pair_order, suggest_studio_mappings,
    validate_studio_duplicate_targets,
)


def make_book(path: Path, sheet_name: str, shifted: bool = False) -> None:
    book = Workbook(); sheet = book.active; sheet.title = sheet_name
    ro, co = (2, 1) if shifted else (0, 0)
    sheet.cell(1 + ro, 2 + co, "余额")
    sheet.cell(2 + ro, 2 + co, "人民币")
    sheet.cell(2 + ro, 3 + co, "外币")
    sheet.cell(3 + ro, 1 + co, "Ⅰ.资产")
    for offset, (label, value) in enumerate((("1.现金", 10), ("2.存款", 20)), 4):
        sheet.cell(offset + ro, 1 + co, label)
        sheet.cell(offset + ro, 2 + co, value)
        sheet.cell(offset + ro, 3 + co, value * 2)
    book.save(path)


class StudioTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def test_window_starts_on_file_selection(self):
        window = StudioWindow()
        self.assertIs(window.stack.currentWidget(), window.files_page)
        window.close()

    def test_module_imports_when_tkinter_is_unavailable(self):
        script = r'''\
import importlib.abc
import sys
class BlockTk(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path=None, target=None):
        if fullname == "tkinter" or fullname.startswith("tkinter."):
            raise ModuleNotFoundError("blocked for packaged-build test")
        return None
sys.meta_path.insert(0, BlockTk())
import excel_mapping_studio
assert "tkinter" in sys.modules
'''
        environment = dict(os.environ)
        environment["QT_QPA_PLATFORM"] = "offscreen"
        completed = subprocess.run(
            [sys.executable, "-c", script],
            cwd=Path(__file__).parent,
            env=environment,
            text=True,
            capture_output=True,
            timeout=30,
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)

    def test_sheet_list_chooses_best_target_without_selecting_cells(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "来源.xlsx"; target = root / "目标.xlsx"
            make_book(source, "GFX010")
            make_book(target, "GFX010_报表", True)
            rows = suggest_studio_mappings([source], [target])
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0].target_sheet, "GFX010_报表")
            self.assertEqual(rows[0].source_cells, [])
            self.assertEqual(rows[0].target_cells, [])

    def test_manual_selection_is_the_only_saved_rule(self):
        row = WorksheetMapping(
            "source.xlsx", "数据", "target.xlsx", "报表",
            ["A2", "A1"], ["B1", "B2"],
        )
        rules = mapping_rules_from_rows([row])
        self.assertEqual(rules[0].source_cells, ["A2", "A1"])
        self.assertEqual(rules[0].target_cells, ["B1", "B2"])

    def test_mismatched_manual_counts_are_rejected(self):
        row = WorksheetMapping(
            "source.xlsx", "数据", "target.xlsx", "报表",
            ["A1", "A2"], ["B1"],
        )
        with self.assertRaisesRegex(ValueError, "数量必须相同"):
            mapping_rules_from_rows([row])

    def test_existing_cross_order_survives_viewer_save(self):
        row = WorksheetMapping(
            "source.xlsx", "数据", "target.xlsx", "报表",
            ["A2", "A1"], ["B1", "B2"],
        )
        sources, targets = preserve_existing_pair_order(
            row, ["A1", "A2", "C1"], ["B1", "B2", "D1"]
        )
        self.assertEqual(sources, ["A2", "A1", "C1"])
        self.assertEqual(targets, ["B1", "B2", "D1"])

    def test_duplicate_smart_addresses_do_not_break_equal_visible_counts(self):
        row = WorksheetMapping(
            "source.xlsx", "数据", "target.xlsx", "报表",
            ["A1", "A1", "A2"], ["B1", "B2", "B2"],
        )
        sources, targets = preserve_existing_pair_order(
            row, ["A1", "A2", "A3"], ["B1", "B2", "B3"]
        )
        self.assertEqual(len(sources), 3)
        self.assertEqual(len(targets), 3)
        self.assertEqual(len(set(sources)), 3)
        self.assertEqual(len(set(targets)), 3)

    def test_duplicate_target_error_names_both_mapping_relations(self):
        first = WorksheetMapping(
            "/tmp/source1.xlsx", "数据1", "/tmp/target.xlsx", "报表",
            ["A1", "A2"], ["C8", "C9"],
        )
        second = WorksheetMapping(
            "/tmp/source2.xlsx", "数据2", "/tmp/target.xlsx", "报表",
            ["B1", "B2"], ["C8", "D9"],
        )
        with self.assertRaises(ValueError) as caught:
            validate_studio_duplicate_targets([first, second])
        message = str(caught.exception)
        self.assertIn("第 1 组", message)
        self.assertIn("第 2 组", message)
        self.assertIn("source1.xlsx/数据1/A1-A2", message)
        self.assertIn("source2.xlsx/数据2/B1-B2", message)
        self.assertIn("target.xlsx/报表/C8", message)


if __name__ == "__main__":
    unittest.main()
