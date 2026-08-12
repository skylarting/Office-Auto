import unittest
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace

from excel_mapper import MappingRule, MODE_MANUAL
from excel_mapping_qt import ExcelMappingQtWindow


class ExcelMappingQtTests(unittest.TestCase):
    def make_rule(self) -> MappingRule:
        return MappingRule(
            "source.xlsx",
            "江苏分行统计",
            ["A1"],
            "target.xlsx",
            "",
            ["A1"],
            MODE_MANUAL,
        )

    def test_target_sheet_uses_best_source_sheet_name_match(self) -> None:
        rule = self.make_rule()

        ExcelMappingQtWindow._auto_match_target_sheet(
            object(),
            rule,
            ["浙江分行统计", "江苏分行汇总", "其他"],
        )

        self.assertEqual(rule.target_sheet, "江苏分行汇总")

    def test_target_sheet_falls_back_to_first_sheet(self) -> None:
        rule = self.make_rule()

        ExcelMappingQtWindow._auto_match_target_sheet(
            object(), rule, ["甲", "乙"]
        )

        self.assertEqual(rule.target_sheet, "甲")

    def test_manual_target_sheet_is_not_overwritten(self) -> None:
        rule = self.make_rule()
        rule.target_sheet = "用户选择"
        rule._target_sheet_manually_selected = True

        ExcelMappingQtWindow._auto_match_target_sheet(
            object(), rule, ["江苏分行统计", "用户选择"]
        )

        self.assertEqual(rule.target_sheet, "用户选择")

    def test_default_output_folder_is_deferred_until_mapping(self) -> None:
        class FakeEdit:
            def __init__(self) -> None:
                self.text = ""

            def setText(self, value: str) -> None:
                self.text = value

            def clear(self) -> None:
                self.text = ""

        with TemporaryDirectory() as folder:
            window = SimpleNamespace(
                base_folder=Path(folder),
                output_folder=None,
                output_edit=FakeEdit(),
            )

            ExcelMappingQtWindow._set_default_output_folder(window)

            expected = Path(folder) / "映射结果"
            self.assertEqual(window.output_folder, expected)
            self.assertFalse(expected.exists())
            self.assertTrue(window.output_edit.text)


if __name__ == "__main__":
    unittest.main()
