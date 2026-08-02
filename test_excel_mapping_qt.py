import unittest

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


if __name__ == "__main__":
    unittest.main()
