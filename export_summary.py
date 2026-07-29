"""GUI tool for extracting configurable Excel cells into summary sheets."""

from __future__ import annotations

from argparse import ArgumentParser
from dataclasses import dataclass
from pathlib import Path
import re
import sys
from tkinter import (
    BOTH,
    END,
    LEFT,
    W,
    X,
    BooleanVar,
    StringVar,
    Text,
    Tk,
    filedialog,
    messagebox,
    ttk,
)
from typing import Callable

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill


DEFAULT_CELL_ADDRESSES = [
    "I8", "J8", "L8",
    "I17", "J17", "L17",
    "I24", "J24", "L24",
    "I25", "J25", "L25",
    "I26", "J26", "L26",
]
VALID_CELL_RE = re.compile(r"^[A-Z]{1,3}[1-9][0-9]*$")
WORKBOOK_SUFFIXES = {".xlsx", ".xlsm"}


@dataclass
class SummaryRecord:
    file_name: str
    sheet_name: str
    values: list[object]
    number_formats: list[str]


def parse_cell_addresses(raw_text: str) -> list[str]:
    """Parse comma, semicolon, whitespace, or Chinese punctuation separators."""
    parts = re.split(r"[\s,，;；]+", raw_text.strip().upper())
    addresses: list[str] = []
    invalid: list[str] = []

    for part in parts:
        if not part:
            continue
        if not VALID_CELL_RE.fullmatch(part):
            invalid.append(part)
        elif part not in addresses:
            addresses.append(part)

    if invalid:
        raise ValueError("以下单元格地址不合法：" + "、".join(invalid))
    if not addresses:
        raise ValueError("请至少填写一个需要提取的单元格地址。")
    return addresses


def is_source_workbook(path: Path, summary_name: str = "汇总") -> bool:
    name = path.name
    if name.startswith("~$") or name.startswith(".~"):
        return False
    if path.suffix.lower() not in WORKBOOK_SUFFIXES:
        return False
    excluded_stems = {
        "SheetSummary",
        "批量汇总表",
        f"{summary_name}",
    }
    if path.stem in excluded_stems:
        return False
    if path.stem.endswith("_汇总表") or path.stem.endswith("_汇总表_Python"):
        return False
    return True


def find_workbooks(path: Path, source_mode: str, summary_name: str) -> list[Path]:
    if source_mode == "file":
        if not path.is_file() or not is_source_workbook(path, summary_name):
            raise ValueError("请选择有效的 .xlsx 或 .xlsm 文件。")
        return [path]

    if not path.is_dir():
        raise ValueError("请选择一个有效的文件夹。")

    files = sorted(
        item for item in path.iterdir()
        if item.is_file() and is_source_workbook(item, summary_name)
    )
    if not files:
        raise ValueError("所选文件夹中没有可处理的 .xlsx 或 .xlsm 文件。")
    return files


def read_records(
    source_path: Path,
    addresses: list[str],
    summary_name: str,
) -> list[SummaryRecord]:
    keep_vba = source_path.suffix.lower() == ".xlsm"
    value_book = load_workbook(
        source_path,
        data_only=True,
        keep_vba=keep_vba,
    )
    format_book = load_workbook(
        source_path,
        data_only=False,
        keep_vba=keep_vba,
    )
    records: list[SummaryRecord] = []

    try:
        for sheet_name in value_book.sheetnames:
            if sheet_name == summary_name:
                continue
            value_sheet = value_book[sheet_name]
            format_sheet = format_book[sheet_name]
            records.append(
                SummaryRecord(
                    file_name=source_path.name,
                    sheet_name=sheet_name,
                    values=[value_sheet[address].value for address in addresses],
                    number_formats=[
                        format_sheet[address].number_format
                        for address in addresses
                    ],
                )
            )
    finally:
        value_book.close()
        format_book.close()

    return records


def format_summary_sheet(
    sheet,
    addresses: list[str],
    records: list[SummaryRecord],
    include_file_name: bool,
) -> None:
    headers = (
        ["文件名称", "Sheet名称", *addresses]
        if include_file_name
        else ["Sheet名称", *addresses]
    )
    sheet.append(headers)

    value_start_column = 3 if include_file_name else 2
    for row_number, record in enumerate(records, start=2):
        if include_file_name:
            sheet.cell(row=row_number, column=1, value=record.file_name)
            sheet.cell(row=row_number, column=2, value=record.sheet_name)
        else:
            sheet.cell(row=row_number, column=1, value=record.sheet_name)

        for offset, (value, number_format) in enumerate(
            zip(record.values, record.number_formats),
        ):
            target = sheet.cell(
                row=row_number,
                column=value_start_column + offset,
                value=value,
            )
            target.number_format = number_format

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sheet.freeze_panes = "A2"
    sheet.sheet_view.showGridLines = False
    sheet.column_dimensions["A"].width = 24 if include_file_name else 18
    if include_file_name:
        sheet.column_dimensions["B"].width = 18
    for column in range(value_start_column, len(headers) + 1):
        sheet.column_dimensions[
            sheet.cell(row=1, column=column).column_letter
        ].width = 14


def create_new_summary_workbook(
    source_files: list[Path],
    addresses: list[str],
    output_path: Path,
    summary_name: str,
    progress: Callable[[int, int, str], None] | None = None,
) -> None:
    all_records: list[SummaryRecord] = []
    total = len(source_files)

    for index, source_path in enumerate(source_files, start=1):
        if progress:
            progress(index - 1, total, f"正在读取：{source_path.name}")
        all_records.extend(read_records(source_path, addresses, summary_name))

    output_book = Workbook()
    output_sheet = output_book.active
    output_sheet.title = summary_name
    format_summary_sheet(
        output_sheet,
        addresses,
        all_records,
        include_file_name=len(source_files) > 1,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_book.save(output_path)
    output_book.close()

    if progress:
        progress(total, total, f"已生成：{output_path.name}")


def insert_summary_into_workbook(
    source_path: Path,
    addresses: list[str],
    summary_name: str,
) -> None:
    records = read_records(source_path, addresses, summary_name)
    keep_vba = source_path.suffix.lower() == ".xlsm"
    workbook = load_workbook(source_path, keep_vba=keep_vba)

    if summary_name in workbook.sheetnames:
        workbook.remove(workbook[summary_name])
    summary_sheet = workbook.create_sheet(summary_name, 0)
    format_summary_sheet(
        summary_sheet,
        addresses,
        records,
        include_file_name=False,
    )
    workbook.save(source_path)
    workbook.close()


def insert_summaries(
    source_files: list[Path],
    addresses: list[str],
    summary_name: str,
    progress: Callable[[int, int, str], None] | None = None,
) -> None:
    total = len(source_files)
    for index, source_path in enumerate(source_files, start=1):
        if progress:
            progress(index - 1, total, f"正在处理：{source_path.name}")
        insert_summary_into_workbook(source_path, addresses, summary_name)
        if progress:
            progress(index, total, f"已完成：{source_path.name}")


def default_output_path(source_path: Path, source_mode: str) -> Path:
    if source_mode == "file":
        return source_path.with_name(f"{source_path.stem}_汇总表.xlsx")
    return source_path / "批量汇总表.xlsx"


class SummaryApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("Excel 单元格汇总工具")
        self.root.geometry("780x610")
        self.root.minsize(720, 560)

        self.source_mode = StringVar(value="file")
        self.source_path = StringVar()
        self.output_mode = StringVar(value="new")
        self.output_path = StringVar()
        self.summary_name = StringVar(value="汇总")
        self.keep_on_top = BooleanVar(value=False)
        self.status_text = StringVar(value="请选择一个 Excel 文件或文件夹。")

        self._build_ui()
        self._update_output_state()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=18)
        container.pack(fill=BOTH, expand=True)

        source_frame = ttk.LabelFrame(container, text="1. 选择数据来源", padding=12)
        source_frame.pack(fill=X, pady=(0, 12))

        ttk.Radiobutton(
            source_frame,
            text="选择一个表格文件",
            variable=self.source_mode,
            value="file",
            command=self._source_mode_changed,
        ).pack(side=LEFT, padx=(0, 20))
        ttk.Radiobutton(
            source_frame,
            text="选择文件夹内的所有表格",
            variable=self.source_mode,
            value="folder",
            command=self._source_mode_changed,
        ).pack(side=LEFT)

        source_path_row = ttk.Frame(source_frame)
        source_path_row.pack(fill=X, pady=(12, 0))
        ttk.Entry(source_path_row, textvariable=self.source_path).pack(
            side=LEFT,
            fill=X,
            expand=True,
        )
        ttk.Button(
            source_path_row,
            text="浏览...",
            command=self._browse_source,
        ).pack(side=LEFT, padx=(8, 0))

        cell_frame = ttk.LabelFrame(container, text="2. 填写提取单元格", padding=12)
        cell_frame.pack(fill=X, pady=(0, 12))
        ttk.Label(
            cell_frame,
            text="可用逗号、空格或换行分隔，例如：I8, J8, L8, I17",
        ).pack(anchor=W)
        self.cell_text = Text(cell_frame, height=4, wrap="word")
        self.cell_text.pack(fill=X, pady=(8, 0))
        self.cell_text.insert("1.0", ", ".join(DEFAULT_CELL_ADDRESSES))

        output_frame = ttk.LabelFrame(container, text="3. 选择输出方式", padding=12)
        output_frame.pack(fill=X, pady=(0, 12))
        ttk.Radiobutton(
            output_frame,
            text="新建一个汇总文件",
            variable=self.output_mode,
            value="new",
            command=self._update_output_state,
        ).pack(anchor=W)
        ttk.Radiobutton(
            output_frame,
            text="在每个源工作簿最前面插入汇总表",
            variable=self.output_mode,
            value="insert",
            command=self._update_output_state,
        ).pack(anchor=W, pady=(6, 0))

        name_row = ttk.Frame(output_frame)
        name_row.pack(fill=X, pady=(10, 0))
        ttk.Label(name_row, text="汇总表名称：").pack(side=LEFT)
        ttk.Entry(
            name_row,
            textvariable=self.summary_name,
            width=18,
        ).pack(side=LEFT)

        output_path_row = ttk.Frame(output_frame)
        output_path_row.pack(fill=X, pady=(10, 0))
        ttk.Label(output_path_row, text="输出文件：").pack(side=LEFT)
        self.output_entry = ttk.Entry(
            output_path_row,
            textvariable=self.output_path,
        )
        self.output_entry.pack(side=LEFT, fill=X, expand=True)
        self.output_button = ttk.Button(
            output_path_row,
            text="浏览...",
            command=self._browse_output,
        )
        self.output_button.pack(side=LEFT, padx=(8, 0))

        progress_frame = ttk.Frame(container)
        progress_frame.pack(fill=X, pady=(4, 10))
        self.progress = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress.pack(fill=X)
        ttk.Label(
            progress_frame,
            textvariable=self.status_text,
        ).pack(anchor=W, pady=(6, 0))

        ttk.Checkbutton(
            container,
            text="窗口保持在最前面",
            variable=self.keep_on_top,
            command=lambda: self.root.attributes(
                "-topmost",
                self.keep_on_top.get(),
            ),
        ).pack(anchor=W)

        action_row = ttk.Frame(container)
        action_row.pack(fill=X, pady=(14, 0))
        ttk.Button(
            action_row,
            text="开始提取",
            command=self._run,
        ).pack(side=LEFT)
        ttk.Button(
            action_row,
            text="退出",
            command=self.root.destroy,
        ).pack(side=LEFT, padx=(8, 0))

    def _source_mode_changed(self) -> None:
        self.source_path.set("")
        self.output_path.set("")
        self._update_output_state()

    def _browse_source(self) -> None:
        if self.source_mode.get() == "file":
            selected = filedialog.askopenfilename(
                title="选择 Excel 文件",
                filetypes=[
                    ("Excel 文件", "*.xlsx *.xlsm"),
                    ("所有文件", "*.*"),
                ],
            )
        else:
            selected = filedialog.askdirectory(title="选择包含 Excel 文件的文件夹")

        if selected:
            path = Path(selected)
            self.source_path.set(str(path))
            if self.output_mode.get() == "new":
                self.output_path.set(
                    str(default_output_path(path, self.source_mode.get()))
                )

    def _browse_output(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="选择汇总文件保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
            initialfile=Path(self.output_path.get() or "汇总表.xlsx").name,
        )
        if selected:
            self.output_path.set(selected)

    def _update_output_state(self) -> None:
        state = "normal" if self.output_mode.get() == "new" else "disabled"
        self.output_entry.configure(state=state)
        self.output_button.configure(state=state)

        source_text = self.source_path.get().strip()
        if state == "normal" and source_text and not self.output_path.get():
            source_path = Path(source_text)
            self.output_path.set(
                str(default_output_path(source_path, self.source_mode.get()))
            )

    def _progress(self, current: int, total: int, message: str) -> None:
        self.progress.configure(maximum=max(total, 1), value=current)
        self.status_text.set(message)
        self.root.update_idletasks()

    def _run(self) -> None:
        try:
            summary_name = self.summary_name.get().strip()
            if not summary_name:
                raise ValueError("请填写汇总表名称。")
            if any(char in summary_name for char in r"[]:*?/\\"):
                raise ValueError("汇总表名称包含 Excel 不允许的字符。")
            if len(summary_name) > 31:
                raise ValueError("汇总表名称不能超过 31 个字符。")

            addresses = parse_cell_addresses(self.cell_text.get("1.0", END))
            source_path = Path(self.source_path.get().strip())
            source_files = find_workbooks(
                source_path,
                self.source_mode.get(),
                summary_name,
            )

            if self.output_mode.get() == "new":
                output_text = self.output_path.get().strip()
                output_path = (
                    Path(output_text)
                    if output_text
                    else default_output_path(source_path, self.source_mode.get())
                )
                if output_path.resolve() in {
                    path.resolve() for path in source_files
                }:
                    raise ValueError("输出文件不能覆盖源文件。")
                create_new_summary_workbook(
                    source_files,
                    addresses,
                    output_path,
                    summary_name,
                    self._progress,
                )
                result_message = f"汇总文件已生成：\n{output_path}"
            else:
                confirmed = messagebox.askyesno(
                    "确认写入源文件",
                    "此操作会修改所选源工作簿，并在最前面插入或替换"
                    f"“{summary_name}”工作表。\n\n建议提前备份源文件。是否继续？",
                    parent=self.root,
                )
                if not confirmed:
                    return
                insert_summaries(
                    source_files,
                    addresses,
                    summary_name,
                    self._progress,
                )
                result_message = f"已更新 {len(source_files)} 个工作簿。"

            messagebox.showinfo("处理完成", result_message, parent=self.root)
        except Exception as exc:
            self.status_text.set("处理失败。")
            messagebox.showerror("处理失败", str(exc), parent=self.root)


def run_command_line() -> None:
    parser = ArgumentParser(description="汇总 Excel 工作簿中的指定单元格。")
    parser.add_argument("source", type=Path, help="源工作簿路径")
    parser.add_argument("-o", "--output", type=Path, help="输出工作簿路径")
    parser.add_argument(
        "-c",
        "--cells",
        default=",".join(DEFAULT_CELL_ADDRESSES),
        help="需要提取的单元格，以逗号分隔",
    )
    args = parser.parse_args()

    addresses = parse_cell_addresses(args.cells)
    output_path = args.output or args.source.with_name(
        f"{args.source.stem}_汇总表.xlsx"
    )
    create_new_summary_workbook(
        [args.source],
        addresses,
        output_path,
        "汇总",
    )
    print(f"汇总表已生成：{output_path.resolve()}")


def main() -> None:
    if len(sys.argv) > 1:
        run_command_line()
        return
    root = Tk()
    SummaryApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
