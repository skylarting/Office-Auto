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
    RIGHT,
    W,
    X,
    Canvas,
    IntVar,
    StringVar,
    Text,
    Tk,
    filedialog,
    messagebox,
    ttk,
)
from typing import Callable, Iterable

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


def column_letters_to_number(letters: str) -> int:
    value = 0
    for character in letters.upper():
        if character < "A" or character > "Z":
            raise ValueError("列名必须由英文字母组成，例如 Z 或 AZ。")
        value = value * 26 + ord(character) - ord("A") + 1
    return value


def column_number_to_letters(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def cell_sort_key(address: str) -> tuple[int, int]:
    match = re.fullmatch(r"([A-Z]{1,3})([1-9][0-9]*)", address)
    if not match:
        return (sys.maxsize, sys.maxsize)
    return (int(match.group(2)), column_letters_to_number(match.group(1)))


class CellGridPicker(ttk.Frame):
    """Scrollable canvas-based cell picker kept in sync with manual input."""

    CELL_WIDTH = 48
    CELL_HEIGHT = 28
    ROW_HEADER_WIDTH = 48
    COLUMN_HEADER_HEIGHT = 28
    MAX_ROWS = 200
    MAX_COLUMNS = 52

    def __init__(
        self,
        parent,
        on_change: Callable[[list[str]], None],
    ) -> None:
        super().__init__(parent)
        self.on_change = on_change
        self.row_count = 30
        self.column_count = 26
        self.selected: set[str] = set()

        controls = ttk.Frame(self)
        controls.pack(fill=X, pady=(0, 8))
        ttk.Label(controls, text="显示行数：").pack(side=LEFT)
        self.rows_var = IntVar(value=self.row_count)
        ttk.Spinbox(
            controls,
            from_=1,
            to=self.MAX_ROWS,
            width=6,
            textvariable=self.rows_var,
        ).pack(side=LEFT)
        ttk.Label(controls, text="显示至列：").pack(side=LEFT, padx=(14, 0))
        self.end_column_var = StringVar(value="Z")
        ttk.Entry(
            controls,
            width=6,
            textvariable=self.end_column_var,
        ).pack(side=LEFT)
        ttk.Button(
            controls,
            text="更新网格",
            command=self.update_dimensions,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Label(
            controls,
            text="点击单元格可选中或取消",
            foreground="#666666",
        ).pack(side=LEFT, padx=(16, 0))

        grid_frame = ttk.Frame(self)
        grid_frame.pack(fill=BOTH, expand=True)
        self.canvas = Canvas(
            grid_frame,
            height=220,
            background="#FFFFFF",
            highlightthickness=1,
            highlightbackground="#B8B8B8",
        )
        x_scroll = ttk.Scrollbar(
            grid_frame,
            orient="horizontal",
            command=self.canvas.xview,
        )
        y_scroll = ttk.Scrollbar(
            grid_frame,
            orient="vertical",
            command=self.canvas.yview,
        )
        self.canvas.configure(
            xscrollcommand=x_scroll.set,
            yscrollcommand=y_scroll.set,
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        grid_frame.rowconfigure(0, weight=1)
        grid_frame.columnconfigure(0, weight=1)
        self.canvas.bind("<Button-1>", self._handle_click)
        self._draw()

    def update_dimensions(self) -> None:
        try:
            rows = int(self.rows_var.get())
            columns = column_letters_to_number(
                self.end_column_var.get().strip()
            )
            if not 1 <= rows <= self.MAX_ROWS:
                raise ValueError(
                    f"显示行数应在 1 至 {self.MAX_ROWS} 之间。"
                )
            if not 1 <= columns <= self.MAX_COLUMNS:
                raise ValueError(
                    "可视选择器最多显示至 AZ 列；更远地址可用手动输入。"
                )
            self.row_count = rows
            self.column_count = columns
            self.end_column_var.set(column_number_to_letters(columns))
            self._draw()
        except (TypeError, ValueError) as exc:
            messagebox.showerror("网格范围错误", str(exc), parent=self)

    def set_selected(self, addresses: Iterable[str]) -> None:
        self.selected = set(addresses)
        self._draw()

    def _handle_click(self, event) -> None:
        x = self.canvas.canvasx(event.x)
        y = self.canvas.canvasy(event.y)
        if x < self.ROW_HEADER_WIDTH or y < self.COLUMN_HEADER_HEIGHT:
            return

        column = int((x - self.ROW_HEADER_WIDTH) // self.CELL_WIDTH) + 1
        row = int((y - self.COLUMN_HEADER_HEIGHT) // self.CELL_HEIGHT) + 1
        if not 1 <= column <= self.column_count or not 1 <= row <= self.row_count:
            return

        address = f"{column_number_to_letters(column)}{row}"
        if address in self.selected:
            self.selected.remove(address)
        else:
            self.selected.add(address)
        ordered = sorted(self.selected, key=cell_sort_key)
        self._draw()
        self.on_change(ordered)

    def _draw(self) -> None:
        self.canvas.delete("all")
        total_width = (
            self.ROW_HEADER_WIDTH + self.column_count * self.CELL_WIDTH
        )
        total_height = (
            self.COLUMN_HEADER_HEIGHT + self.row_count * self.CELL_HEIGHT
        )
        self.canvas.configure(scrollregion=(0, 0, total_width, total_height))

        self.canvas.create_rectangle(
            0,
            0,
            self.ROW_HEADER_WIDTH,
            self.COLUMN_HEADER_HEIGHT,
            fill="#E9EDF3",
            outline="#C5CBD3",
        )
        for column in range(1, self.column_count + 1):
            x1 = self.ROW_HEADER_WIDTH + (column - 1) * self.CELL_WIDTH
            x2 = x1 + self.CELL_WIDTH
            self.canvas.create_rectangle(
                x1,
                0,
                x2,
                self.COLUMN_HEADER_HEIGHT,
                fill="#E9EDF3",
                outline="#C5CBD3",
            )
            self.canvas.create_text(
                (x1 + x2) / 2,
                self.COLUMN_HEADER_HEIGHT / 2,
                text=column_number_to_letters(column),
                fill="#303640",
            )

        for row in range(1, self.row_count + 1):
            y1 = self.COLUMN_HEADER_HEIGHT + (row - 1) * self.CELL_HEIGHT
            y2 = y1 + self.CELL_HEIGHT
            self.canvas.create_rectangle(
                0,
                y1,
                self.ROW_HEADER_WIDTH,
                y2,
                fill="#E9EDF3",
                outline="#C5CBD3",
            )
            self.canvas.create_text(
                self.ROW_HEADER_WIDTH / 2,
                (y1 + y2) / 2,
                text=str(row),
                fill="#303640",
            )

            for column in range(1, self.column_count + 1):
                x1 = self.ROW_HEADER_WIDTH + (column - 1) * self.CELL_WIDTH
                x2 = x1 + self.CELL_WIDTH
                address = f"{column_number_to_letters(column)}{row}"
                selected = address in self.selected
                self.canvas.create_rectangle(
                    x1,
                    y1,
                    x2,
                    y2,
                    fill="#2F75B5" if selected else "#FFFFFF",
                    outline="#D8DDE5",
                )
                if selected:
                    self.canvas.create_text(
                        (x1 + x2) / 2,
                        (y1 + y2) / 2,
                        text=address,
                        fill="#FFFFFF",
                    )


class SummaryApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("Excel 单元格汇总工具")
        self.root.geometry("820x720")
        self.root.minsize(760, 650)

        self.source_mode = ""
        self.source_path = StringVar()
        self.source_summary = StringVar(value="尚未选择数据来源")
        self.output_choice = StringVar(value="新建表格文件")
        self.output_path = StringVar()
        self.summary_name = StringVar(value="汇总")
        self.selection_summary = StringVar()
        self.validation_text = StringVar()
        self.status_text = StringVar(value="等待开始")
        self.progress_percent = StringVar(value="0%")
        self.selected_addresses = list(DEFAULT_CELL_ADDRESSES)
        self.cells_valid = True
        self._syncing_cells = False

        self._build_ui()
        self._set_selected_addresses(DEFAULT_CELL_ADDRESSES)
        self._show_output_fields()
        self._reset_progress()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=18)
        container.pack(fill=BOTH, expand=True)

        source_frame = ttk.LabelFrame(container, text="1. 选择数据来源", padding=12)
        source_frame.pack(fill=X, pady=(0, 12))

        source_buttons = ttk.Frame(source_frame)
        source_buttons.pack(fill=X)
        ttk.Button(
            source_buttons,
            text="选择表格文件",
            command=self._choose_file,
        ).pack(side=LEFT)
        ttk.Button(
            source_buttons,
            text="选择文件夹",
            command=self._choose_folder,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Label(
            source_frame,
            textvariable=self.source_summary,
            foreground="#3A3A3A",
        ).pack(anchor=W, pady=(10, 0))

        cell_frame = ttk.LabelFrame(
            container,
            text="2. 选择需要提取的单元格",
            padding=12,
        )
        cell_frame.pack(fill=BOTH, expand=True, pady=(0, 12))
        self.cell_notebook = ttk.Notebook(cell_frame)
        self.cell_notebook.pack(fill=BOTH, expand=True)

        manual_tab = ttk.Frame(self.cell_notebook, padding=10)
        visual_tab = ttk.Frame(self.cell_notebook, padding=10)
        self.cell_notebook.add(manual_tab, text="手动输入")
        self.cell_notebook.add(visual_tab, text="点击选择")

        ttk.Label(
            manual_tab,
            text="可用逗号、空格或换行分隔，例如：I8, J8, L8, I17",
        ).pack(anchor=W)
        self.cell_text = Text(manual_tab, height=4, wrap="word")
        self.cell_text.pack(fill=X, pady=(8, 6))
        self.cell_text.bind("<KeyRelease>", self._manual_cells_changed)
        ttk.Label(
            manual_tab,
            textvariable=self.validation_text,
            foreground="#C62828",
        ).pack(anchor=W)
        manual_actions = ttk.Frame(manual_tab)
        manual_actions.pack(fill=X, pady=(8, 0))
        ttk.Button(
            manual_actions,
            text="恢复默认",
            command=self._restore_default_cells,
        ).pack(side=LEFT)
        ttk.Button(
            manual_actions,
            text="清空",
            command=self._clear_cells,
        ).pack(side=LEFT, padx=(8, 0))

        self.cell_grid = CellGridPicker(
            visual_tab,
            self._grid_cells_changed,
        )
        self.cell_grid.pack(fill=BOTH, expand=True)
        visual_actions = ttk.Frame(visual_tab)
        visual_actions.pack(fill=X, pady=(8, 0))
        ttk.Button(
            visual_actions,
            text="恢复默认",
            command=self._restore_default_cells,
        ).pack(side=LEFT)
        ttk.Button(
            visual_actions,
            text="清空选择",
            command=self._clear_cells,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Label(
            visual_actions,
            textvariable=self.selection_summary,
        ).pack(side=LEFT, padx=(16, 0))

        output_frame = ttk.LabelFrame(container, text="3. 选择输出方式", padding=12)
        output_frame.pack(fill=X, pady=(0, 12))
        output_selector = ttk.Frame(output_frame)
        output_selector.pack(fill=X)
        ttk.Label(output_selector, text="输出方式：").pack(side=LEFT)
        output_combo = ttk.Combobox(
            output_selector,
            textvariable=self.output_choice,
            values=("新建表格文件", "插入表单"),
            state="readonly",
            width=20,
        )
        output_combo.pack(side=LEFT)
        output_combo.bind("<<ComboboxSelected>>", self._output_changed)
        self.output_fields = ttk.Frame(output_frame)
        self.output_fields.pack(fill=X, pady=(10, 0))

        progress_frame = ttk.LabelFrame(
            container,
            text="4. 处理进度",
            padding=12,
        )
        progress_frame.pack(fill=X, pady=(0, 12))
        progress_title = ttk.Frame(progress_frame)
        progress_title.pack(fill=X, pady=(0, 6))
        ttk.Label(progress_title, text="处理进度").pack(side=LEFT)
        ttk.Label(
            progress_title,
            textvariable=self.progress_percent,
        ).pack(side=RIGHT)
        self.progress = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress.pack(fill=X)
        ttk.Label(
            progress_frame,
            textvariable=self.status_text,
        ).pack(anchor=W, pady=(6, 0))

        action_row = ttk.Frame(container)
        action_row.pack(fill=X)
        self.start_button = ttk.Button(
            action_row,
            text="开始提取",
            command=self._run,
            state="disabled",
        )
        self.start_button.pack(side=LEFT)
        ttk.Button(
            action_row,
            text="退出",
            command=self.root.destroy,
        ).pack(side=LEFT, padx=(8, 0))

    def _choose_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        if selected:
            path = Path(selected)
            self.source_mode = "file"
            self.source_path.set(str(path))
            self.source_summary.set(f"已选择文件：{path}")
            self._source_changed()

    def _choose_folder(self) -> None:
        selected = filedialog.askdirectory(
            title="选择包含 Excel 文件的文件夹"
        )
        if not selected:
            return

        path = Path(selected)
        try:
            files = find_workbooks(path, "folder", self.summary_name.get())
        except ValueError as exc:
            messagebox.showerror("无法选择文件夹", str(exc), parent=self.root)
            return

        confirmed = messagebox.askyesno(
            "确认批量提取",
            f"该文件夹中发现 {len(files)} 个可处理的 Excel 文件。\n\n"
            "是否提取该文件夹下的所有表格文件？",
            parent=self.root,
        )
        if not confirmed:
            return

        self.source_mode = "folder"
        self.source_path.set(str(path))
        self.source_summary.set(
            f"已选择文件夹：{path}（共 {len(files)} 个 Excel 文件）"
        )
        self._source_changed()

    def _source_changed(self) -> None:
        source_path = Path(self.source_path.get())
        if self.output_choice.get() == "新建表格文件":
            self.output_path.set(
                str(default_output_path(source_path, self.source_mode))
            )
            self._show_output_fields()
        self._configuration_changed()

    def _browse_output(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="选择汇总文件保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
            initialfile=Path(self.output_path.get() or "汇总表.xlsx").name,
        )
        if selected:
            self.output_path.set(selected)
            self._configuration_changed()

    def _output_changed(self, _event=None) -> None:
        if (
            self.output_choice.get() == "新建表格文件"
            and self.source_path.get()
        ):
            self.output_path.set(
                str(
                    default_output_path(
                        Path(self.source_path.get()),
                        self.source_mode,
                    )
                )
            )
        self._show_output_fields()
        self._configuration_changed()

    def _show_output_fields(self) -> None:
        for child in self.output_fields.winfo_children():
            child.destroy()

        if self.output_choice.get() == "新建表格文件":
            ttk.Label(self.output_fields, text="输出文件：").pack(side=LEFT)
            output_entry = ttk.Entry(
                self.output_fields,
                textvariable=self.output_path,
            )
            output_entry.pack(side=LEFT, fill=X, expand=True)
            output_entry.bind("<KeyRelease>", self._configuration_changed)
            ttk.Button(
                self.output_fields,
                text="浏览...",
                command=self._browse_output,
            ).pack(side=LEFT, padx=(8, 0))
        else:
            ttk.Label(self.output_fields, text="表单名称：").pack(side=LEFT)
            name_entry = ttk.Entry(
                self.output_fields,
                textvariable=self.summary_name,
                width=20,
            )
            name_entry.pack(side=LEFT)
            name_entry.bind("<KeyRelease>", self._configuration_changed)
            ttk.Label(
                self.output_fields,
                text="将在每个源工作簿最前面插入或替换该表单",
                foreground="#666666",
            ).pack(side=LEFT, padx=(12, 0))

    def _manual_cells_changed(self, _event=None) -> None:
        if self._syncing_cells:
            return
        try:
            addresses = parse_cell_addresses(self.cell_text.get("1.0", END))
            self.validation_text.set("")
            self._set_selected_addresses(addresses, update_text=False)
        except ValueError as exc:
            self.cells_valid = False
            self.validation_text.set(str(exc))
            self.start_button.configure(state="disabled")
            self._reset_progress()

    def _grid_cells_changed(self, addresses: list[str]) -> None:
        self._set_selected_addresses(addresses)

    def _set_selected_addresses(
        self,
        addresses: Iterable[str],
        update_text: bool = True,
    ) -> None:
        ordered = sorted(set(addresses), key=cell_sort_key)
        self.selected_addresses = ordered
        self.cells_valid = bool(ordered)
        self._syncing_cells = True
        try:
            if update_text:
                self.cell_text.delete("1.0", END)
                self.cell_text.insert("1.0", ", ".join(ordered))
            self.cell_grid.set_selected(ordered)
        finally:
            self._syncing_cells = False

        preview = ", ".join(ordered[:8])
        if len(ordered) > 8:
            preview += "..."
        self.selection_summary.set(
            f"已选择 {len(ordered)} 个单元格"
            + (f"：{preview}" if preview else "")
        )
        self.validation_text.set(
            "" if ordered else "请至少选择一个需要提取的单元格。"
        )
        self._configuration_changed()

    def _restore_default_cells(self) -> None:
        self._set_selected_addresses(DEFAULT_CELL_ADDRESSES)

    def _clear_cells(self) -> None:
        self._set_selected_addresses([])

    def _configuration_changed(self, _event=None) -> None:
        self._reset_progress()
        valid = bool(
            self.source_path.get()
            and self.selected_addresses
            and self.cells_valid
        )
        if self.output_choice.get() == "新建表格文件":
            valid = valid and bool(self.output_path.get().strip())
        else:
            valid = valid and bool(self.summary_name.get().strip())
        self.start_button.configure(state="normal" if valid else "disabled")

    def _reset_progress(self) -> None:
        self.progress.configure(maximum=100, value=0, mode="determinate")
        self.progress_percent.set("0%")
        self.status_text.set("等待开始")

    def _progress(self, current: int, total: int, message: str) -> None:
        maximum = max(total, 1)
        self.progress.configure(maximum=maximum, value=current)
        percentage = int(current / maximum * 100)
        self.progress_percent.set(f"{percentage}%")
        self.status_text.set(message)
        self.root.update_idletasks()

    def _run(self) -> None:
        try:
            self._reset_progress()
            summary_name = (
                "汇总"
                if self.output_choice.get() == "新建表格文件"
                else self.summary_name.get().strip()
            )
            if not summary_name:
                raise ValueError("请填写汇总表名称。")
            if any(char in summary_name for char in r"[]:*?/\\"):
                raise ValueError("汇总表名称包含 Excel 不允许的字符。")
            if len(summary_name) > 31:
                raise ValueError("汇总表名称不能超过 31 个字符。")

            addresses = list(self.selected_addresses)
            source_path = Path(self.source_path.get().strip())
            source_files = find_workbooks(
                source_path,
                self.source_mode,
                summary_name,
            )

            if self.output_choice.get() == "新建表格文件":
                output_text = self.output_path.get().strip()
                output_path = (
                    Path(output_text)
                    if output_text
                    else default_output_path(source_path, self.source_mode)
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

            self.progress.configure(value=self.progress.cget("maximum"))
            self.progress_percent.set("100%")
            self.status_text.set("处理完成")
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
