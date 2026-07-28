"""Export selected cells from every worksheet into one summary workbook."""

from argparse import ArgumentParser
from pathlib import Path
import sys
from tkinter import Tk, messagebox

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


CELL_ADDRESSES = [
    "I8", "J8", "L8",
    "I17", "J17", "L17",
    "I24", "J24", "L24",
    "I25", "J25", "L25",
    "I26", "J26", "L26",
]


def export_summary(source_path: Path, output_path: Path) -> None:
    # data_only=True reads the last calculated result instead of formula text.
    value_book = load_workbook(source_path, data_only=True)
    # A second read preserves the original cell number formats.
    format_book = load_workbook(source_path, data_only=False)

    output_book = Workbook()
    output_sheet = output_book.active
    output_sheet.title = "汇总"

    headers = ["Sheet名称", *CELL_ADDRESSES]
    output_sheet.append(headers)

    for row_number, sheet_name in enumerate(value_book.sheetnames, start=2):
        value_sheet = value_book[sheet_name]
        format_sheet = format_book[sheet_name]

        output_sheet.cell(row=row_number, column=1, value=sheet_name)

        for column_number, address in enumerate(CELL_ADDRESSES, start=2):
            source_value_cell = value_sheet[address]
            source_format_cell = format_sheet[address]
            output_cell = output_sheet.cell(
                row=row_number,
                column=column_number,
                value=source_value_cell.value,
            )

            # Copy display format only; the complete underlying value is retained.
            output_cell.number_format = source_format_cell.number_format

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in output_sheet[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    output_sheet.freeze_panes = "A2"
    output_sheet.column_dimensions["A"].width = 18
    for column in range(2, len(headers) + 1):
        output_sheet.column_dimensions[
            output_sheet.cell(row=1, column=column).column_letter
        ].width = 14

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_book.save(output_path)


def application_folder() -> Path:
    """Return the EXE folder, or the script folder when running as Python."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def show_message(title: str, message: str, error: bool = False) -> None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    if error:
        messagebox.showerror(title, message, parent=root)
    else:
        messagebox.showinfo(title, message, parent=root)
    root.destroy()


def is_source_workbook(path: Path) -> bool:
    name = path.name
    lower_name = name.lower()
    if name.startswith("~$") or name.startswith(".~"):
        return False
    if lower_name == "sheetsummary.xlsx":
        return False
    if path.stem.endswith("_汇总表") or path.stem.endswith("_汇总表_Python"):
        return False
    return path.suffix.lower() in {".xlsx", ".xlsm"}


def run_by_double_click() -> None:
    folder = application_folder()
    source_files = sorted(
        path for path in folder.iterdir()
        if path.is_file() and is_source_workbook(path)
    )

    if not source_files:
        show_message(
            "没有找到Excel文件",
            "请把ExcelSummary.exe放到含有.xlsx或.xlsm文件的文件夹中，"
            "然后重新双击运行。",
            error=True,
        )
        return

    completed = []
    failed = []

    for source_path in source_files:
        output_path = source_path.with_name(
            f"{source_path.stem}_汇总表.xlsx"
        )
        try:
            export_summary(source_path, output_path)
            completed.append(output_path.name)
        except Exception as exc:
            failed.append(f"{source_path.name}: {exc}")

    message_parts = []
    if completed:
        message_parts.append(
            "已生成：\n" + "\n".join(completed)
        )
    if failed:
        message_parts.append(
            "处理失败：\n" + "\n".join(failed)
        )

    show_message(
        "汇总完成" if not failed else "汇总结果",
        "\n\n".join(message_parts),
        error=bool(failed),
    )


def main() -> None:
    # No command-line arguments means the program was normally double-clicked.
    if len(sys.argv) == 1:
        run_by_double_click()
        return

    parser = ArgumentParser(
        description="汇总每个 Sheet 中指定单元格的值。"
    )
    parser.add_argument(
        "source",
        type=Path,
        help="源工作簿路径",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="输出路径；不指定时在源文件旁生成“原文件名_汇总表.xlsx”",
    )
    args = parser.parse_args()

    if not args.source.exists():
        parser.error(f"找不到源文件：{args.source}")

    output_path = args.output or args.source.with_name(
        f"{args.source.stem}_汇总表.xlsx"
    )
    export_summary(args.source, output_path)
    print(f"汇总表已生成：{output_path.resolve()}")


if __name__ == "__main__":
    main()
