from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook
import xlwt

import excel_mapper


def make_xlsx(path: Path, sheet_name: str, values: dict[str, object]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name
    for address, value in values.items():
        sheet[address] = value
    workbook.save(path)
    workbook.close()


def make_xls(path: Path, sheet_name: str) -> None:
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet(sheet_name)
    style = xlwt.easyxf(num_format_str="0")
    sheet.write(1, 0, 12.6, style)
    sheet.write(1, 2, 34.4, style)
    workbook.save(str(path))


class ExcelMapperTests(unittest.TestCase):
    def test_workbook_files_in_folder(self) -> None:
        with TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            (folder / "甲.xlsx").touch()
            (folder / "乙.xls").touch()
            (folder / "~$临时.xlsx").touch()
            (folder / "说明.txt").touch()
            self.assertEqual(
                [path.name for path in excel_mapper.workbook_files_in_folder(folder)],
                ["乙.xls", "甲.xlsx"],
            )

    def test_expand_sequence_and_one_to_many(self) -> None:
        sequence = excel_mapper.MappingRule(
            "source.xlsx",
            "数据",
            ["A2", "C2"],
            "target.xlsx",
            "模板",
            ["B5", "D5"],
            excel_mapper.MODE_SEQUENCE,
        )
        expanded = excel_mapper.expand_rule(sequence)
        self.assertEqual(
            [(item.source_cell, item.target_cell) for item in expanded],
            [("A2", "B5"), ("C2", "D5")],
        )

        one_to_many = excel_mapper.MappingRule(
            "source.xlsx",
            "数据",
            ["A2"],
            "target.xlsx",
            "模板",
            ["B5", "D5", "F5"],
            excel_mapper.MODE_ONE_TO_MANY,
        )
        expanded = excel_mapper.expand_rule(one_to_many)
        self.assertEqual(
            [item.target_cell for item in expanded],
            ["B5", "D5", "F5"],
        )

    def test_invalid_mapping_counts(self) -> None:
        rule = excel_mapper.MappingRule(
            "source.xlsx",
            "数据",
            ["A2", "A3"],
            "target.xlsx",
            "模板",
            ["B5"],
            excel_mapper.MODE_SEQUENCE,
        )
        with self.assertRaises(ValueError):
            excel_mapper.expand_rule(rule)

    def test_execute_multiple_targets_without_modifying_originals(self) -> None:
        with TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            source = folder / "来源.xlsx"
            target_one = folder / "目标一.xlsx"
            target_two = folder / "目标二.xlsx"
            output = folder / "输出"
            make_xlsx(source, "数据", {"A2": 12.6, "C2": 34.4})
            make_xlsx(target_one, "模板", {"B5": "原值"})
            make_xlsx(target_two, "报表", {})

            source_book = load_workbook(source)
            source_book["数据"]["A2"].number_format = "0"
            source_book["数据"]["C2"].number_format = "0"
            source_book.save(source)
            source_book.close()

            rules = [
                excel_mapper.MappingRule(
                    str(source),
                    "数据",
                    ["A2", "C2"],
                    str(target_one),
                    "模板",
                    ["B5", "D5"],
                    excel_mapper.MODE_SEQUENCE,
                ),
                excel_mapper.MappingRule(
                    str(source),
                    "数据",
                    ["A2"],
                    str(target_two),
                    "报表",
                    ["C3", "E3"],
                    excel_mapper.MODE_ONE_TO_MANY,
                ),
            ]
            outputs = excel_mapper.execute_mapping_plan(
                [source],
                [target_one, target_two],
                rules,
                output,
            )
            self.assertEqual(len(outputs), 2)

            original = load_workbook(target_one)
            self.assertEqual(original["模板"]["B5"].value, "原值")
            original.close()

            mapped_one = load_workbook(output / "目标一_已映射.xlsx")
            self.assertEqual(mapped_one["模板"]["B5"].value, 12.6)
            self.assertEqual(mapped_one["模板"]["D5"].value, 34.4)
            self.assertEqual(mapped_one["模板"]["B5"].number_format, "0")
            mapped_one.close()

            mapped_two = load_workbook(output / "目标二_已映射.xlsx")
            self.assertEqual(mapped_two["报表"]["C3"].value, 12.6)
            self.assertEqual(mapped_two["报表"]["E3"].value, 12.6)
            mapped_two.close()

    def test_xls_source_and_target_are_supported(self) -> None:
        with TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            source = folder / "来源.xls"
            target = folder / "目标.xls"
            output = folder / "输出"
            make_xls(source, "数据")
            make_xls(target, "模板")
            rule = excel_mapper.MappingRule(
                str(source),
                "数据",
                ["A2", "C2"],
                str(target),
                "模板",
                ["D4", "E4"],
                excel_mapper.MODE_SEQUENCE,
            )

            outputs = excel_mapper.execute_mapping_plan(
                [source],
                [target],
                [rule],
                output,
            )
            self.assertEqual(outputs[0].suffix, ".xlsx")
            copied = load_workbook(outputs[0])
            self.assertEqual(copied["模板"]["D4"].value, 12.6)
            self.assertEqual(copied["模板"]["E4"].value, 34.4)
            copied.close()

    def test_project_round_trip(self) -> None:
        with TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            project = folder / "方案.txt"
            source = folder / "来源.xlsx"
            target = folder / "目标.xlsx"
            rule = excel_mapper.MappingRule(
                str(source),
                "数据",
                ["A2"],
                str(target),
                "模板",
                ["B5"],
                excel_mapper.MODE_MANUAL,
            )
            excel_mapper.save_mapping_project(
                project,
                [rule],
            )
            sources, targets, rules, output = (
                excel_mapper.load_mapping_project(project)
            )
            self.assertEqual(sources, [source.resolve()])
            self.assertEqual(targets, [target.resolve()])
            self.assertEqual(Path(rules[0].source_file), source.resolve())
            self.assertEqual(Path(rules[0].target_file), target.resolve())
            self.assertEqual(rules[0].source_cells, ["A2"])
            self.assertEqual(rules[0].target_cells, ["B5"])
            self.assertEqual(output, (folder / "映射结果").resolve())

            text = project.read_text(encoding="utf-8-sig")
            self.assertEqual(
                text,
                "【来源.xlsx；数据；A2 → 目标.xlsx；模板；B5】\n",
            )

    def test_txt_project_can_be_written_by_hand(self) -> None:
        with TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            project = folder / "手工方案.txt"
            project.write_text(
                "{来源.xlsx;数据;A1 => 目标.xlsx;模板;B2,C2}\n"
                "（来源.xlsx；数据；C3-C5 >> 目标.xlsx；模板；同位置）\n",
                encoding="utf-8-sig",
            )

            sources, targets, rules, output = (
                excel_mapper.load_mapping_project(project)
            )

            self.assertEqual(sources, [(folder / "来源.xlsx").resolve()])
            self.assertEqual(targets, [(folder / "目标.xlsx").resolve()])
            self.assertEqual(rules[0].mode, excel_mapper.MODE_ONE_TO_MANY)
            self.assertEqual(rules[1].source_cells, ["C3", "C4", "C5"])
            self.assertEqual(rules[1].target_cells, ["C3", "C4", "C5"])
            self.assertEqual(output, (folder / "映射结果").resolve())

    def test_txt_project_rejects_unmatched_brackets(self) -> None:
        with TemporaryDirectory() as temp_dir:
            project = Path(temp_dir) / "错误方案.txt"
            project.write_text(
                "【来源.xlsx；数据；A1 → 目标.xlsx；模板；A1]\n",
                encoding="utf-8-sig",
            )
            with self.assertRaisesRegex(ValueError, "括号不匹配"):
                excel_mapper.load_mapping_project(project)

    def test_mapping_mode_is_inferred(self) -> None:
        self.assertEqual(
            excel_mapper.infer_mapping_mode(["A1"], ["B1"]),
            excel_mapper.MODE_MANUAL,
        )
        self.assertEqual(
            excel_mapper.infer_mapping_mode(["A1"], ["B1", "C1"]),
            excel_mapper.MODE_ONE_TO_MANY,
        )
        self.assertEqual(
            excel_mapper.infer_mapping_mode(["A1", "A2"], ["B1", "B2"]),
            excel_mapper.MODE_SEQUENCE,
        )
        with self.assertRaises(ValueError):
            excel_mapper.infer_mapping_mode(
                ["A1", "A2"],
                ["B1", "B2", "B3"],
            )


if __name__ == "__main__":
    unittest.main()
