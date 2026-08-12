import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from openpyxl import Workbook

from smart_template import SheetPair, match_pair


def make_ambiguous_book(path: Path, sheet_name: str, target: bool = False) -> None:
    book = Workbook(); sheet = book.active; sheet.title = sheet_name
    sheet["B1"] = "人民币"; sheet["C1"] = "外币"
    sheet["A2"] = "贷款"; sheet["A3"] = "贷款合计"
    sheet["B2"] = 10; sheet["C2"] = 20
    sheet["B3"] = 30; sheet["C3"] = 40
    if target:
        sheet["A2"] = "各项贷款"
        sheet["A3"] = "贷款总计"
    book.save(path)


class UniqueSmartMatchingTests(unittest.TestCase):
    def test_one_target_cell_is_never_assigned_twice(self):
        with TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "source.xlsx"; target = root / "target.xlsx"
            make_ambiguous_book(source, "数据")
            make_ambiguous_book(target, "报表", True)
            matches = match_pair(
                SheetPair(str(source), "数据", str(target), "报表", 1.0)
            )
            targets = [item.target_address for item in matches if item.target_address]
            self.assertEqual(len(targets), len(set(targets)))


if __name__ == "__main__":
    unittest.main()
