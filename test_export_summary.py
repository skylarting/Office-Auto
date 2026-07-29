from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook, load_workbook
import xlwt

import export_summary


def make_source(path: Path, marker: int) -> None:
    workbook = Workbook()
    first = workbook.active
    first.title = "数据1"
    second = workbook.create_sheet("数据2")
    for index, sheet in enumerate((first, second), start=1):
        sheet["I8"] = marker + index + 0.25
        sheet["I8"].number_format = "0"
        sheet["J8"] = marker + index + 0.75
        sheet["J8"].number_format = "0"
    workbook.save(path)


def make_xls_source(path: Path) -> None:
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet("旧版数据")
    integer_style = xlwt.easyxf(num_format_str="0")
    sheet.write(1, 0, 12.6, integer_style)
    sheet.write(1, 2, 34.4, integer_style)
    sheet.write(2, 0, 56.5, integer_style)
    sheet.write(2, 2, 78.2, integer_style)
    workbook.save(str(path))


class ExportSummaryTests(unittest.TestCase):
    def test_parse_cell_addresses(self) -> None:
        self.assertEqual(
            export_summary.parse_cell_addresses("i8， J8\nL8;I8"),
            ["I8", "J8", "L8"],
        )
        with self.assertRaises(ValueError):
            export_summary.parse_cell_addresses("I0, BAD")

    def test_cell_grid_address_helpers(self) -> None:
        self.assertEqual(export_summary.column_letters_to_number("A"), 1)
        self.assertEqual(export_summary.column_letters_to_number("AZ"), 52)
        self.assertEqual(export_summary.column_number_to_letters(52), "AZ")
        self.assertEqual(
            sorted(["L8", "I17", "J8", "I8"], key=export_summary.cell_sort_key),
            ["I8", "J8", "L8", "I17"],
        )

    def test_create_combined_summary(self) -> None:
        with TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            first = folder / "甲.xlsx"
            second = folder / "乙.xlsx"
            output = folder / "批量汇总表.xlsx"
            make_source(first, 100)
            make_source(second, 200)

            export_summary.create_new_summary_workbook(
                [first, second],
                ["I8", "J8"],
                output,
                "汇总",
            )

            workbook = load_workbook(output, data_only=False)
            sheet = workbook["汇总"]
            self.assertEqual(sheet.max_row, 5)
            self.assertEqual(sheet["A1"].value, "文件名称")
            self.assertEqual(sheet["B1"].value, "Sheet名称")
            self.assertEqual(sheet["C2"].value, 101.25)
            self.assertEqual(sheet["C2"].number_format, "0")
            self.assertEqual(sheet["A1"].border.left.style, "thin")
            self.assertEqual(sheet["D5"].border.bottom.style, "thin")
            workbook.close()

    def test_create_copy_with_summary_as_first_sheet(self) -> None:
        with TemporaryDirectory() as temp_dir:
            source = Path(temp_dir) / "源.xlsx"
            output = Path(temp_dir) / "源_已汇总.xlsx"
            make_source(source, 300)

            export_summary.create_summary_copy(
                source,
                ["I8", "J8"],
                "汇总",
                output,
            )

            original = load_workbook(source, data_only=False)
            self.assertNotIn("汇总", original.sheetnames)
            original.close()

            copied = load_workbook(output, data_only=False)
            self.assertEqual(copied.sheetnames[0], "汇总")
            self.assertEqual(copied["汇总"]["B2"].value, 301.25)
            self.assertEqual(copied["汇总"]["B2"].number_format, "0")
            copied.close()

    def test_default_cell_addresses(self) -> None:
        self.assertEqual(
            export_summary.DEFAULT_CELL_ADDRESSES,
            ["A2", "C2", "A3", "C3"],
        )

    def test_xls_source_and_xlsx_copy(self) -> None:
        with TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            source = folder / "旧版.xls"
            copy_path = folder / "旧版_已汇总.xlsx"
            make_xls_source(source)

            records = export_summary.read_records(
                source,
                export_summary.DEFAULT_CELL_ADDRESSES,
                "汇总",
            )
            self.assertEqual(records[0].values, [12.6, 34.4, 56.5, 78.2])
            self.assertEqual(records[0].number_formats, ["0", "0", "0", "0"])

            export_summary.create_summary_copy(
                source,
                export_summary.DEFAULT_CELL_ADDRESSES,
                "汇总",
                copy_path,
            )
            self.assertTrue(source.exists())
            copied = load_workbook(copy_path)
            self.assertEqual(copied.sheetnames[0], "汇总")
            self.assertEqual(copied["汇总"]["B2"].value, 12.6)
            self.assertEqual(copied["汇总"]["B2"].number_format, "0")
            copied.close()


if __name__ == "__main__":
    unittest.main()

