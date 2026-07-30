"""GUI for mapping cells between multiple Excel workbooks."""

from __future__ import annotations

from dataclasses import asdict, dataclass
import json
from pathlib import Path
import shutil
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    W,
    X,
    StringVar,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
    ttk,
)
from typing import Callable

import xlrd
from openpyxl import load_workbook

from export_summary import (
    WORKBOOK_SUFFIXES,
    column_letters_to_number,
    copy_xls_to_xlsx,
    parse_cell_addresses,
    unique_output_path,
    xls_cell_value,
    xls_number_format,
)


MODE_SEQUENCE = "按顺序一一对应"
MODE_ONE_TO_MANY = "一个来源写入多个目标"
MODE_MANUAL = "逐条手动设置"
MAPPING_MODES = (MODE_SEQUENCE, MODE_ONE_TO_MANY, MODE_MANUAL)


@dataclass
class CellValue:
    value: object
    number_format: str


@dataclass
class MappingRule:
    source_file: str
    source_sheet: str
    source_cells: list[str]
    target_file: str
    target_sheet: str
    target_cells: list[str]
    mode: str


@dataclass
class ExpandedMapping:
    source_file: str
    source_sheet: str
    source_cell: str
    target_file: str
    target_sheet: str
    target_cell: str


def workbook_files_in_folder(folder: Path) -> list[Path]:
    return sorted(
        path
        for path in folder.iterdir()
        if (
            path.is_file()
            and path.suffix.lower() in WORKBOOK_SUFFIXES
            and not path.name.startswith(("~$", ".~"))
        )
    )


def workbook_sheet_names(path: Path) -> list[str]:
    if path.suffix.lower() == ".xls":
        book = xlrd.open_workbook(path, on_demand=True)
        try:
            return book.sheet_names()
        finally:
            book.release_resources()
    book = load_workbook(
        path,
        read_only=True,
        keep_vba=path.suffix.lower() == ".xlsm",
    )
    try:
        return list(book.sheetnames)
    finally:
        book.close()


def expand_rule(rule: MappingRule) -> list[ExpandedMapping]:
    sources = rule.source_cells
    targets = rule.target_cells
    if not sources or not targets:
        raise ValueError("来源单元格和目标单元格都不能为空。")

    if rule.mode == MODE_ONE_TO_MANY:
        if len(sources) != 1:
            raise ValueError("“一个来源写入多个目标”只能填写一个来源单元格。")
        pairs = [(sources[0], target) for target in targets]
    elif rule.mode == MODE_MANUAL:
        if len(sources) != 1 or len(targets) != 1:
            raise ValueError("“逐条手动设置”每条规则只能填写一个来源和一个目标。")
        pairs = [(sources[0], targets[0])]
    elif rule.mode == MODE_SEQUENCE:
        if len(sources) != len(targets):
            raise ValueError(
                "“按顺序一一对应”的来源和目标单元格数量必须相同。"
            )
        pairs = list(zip(sources, targets))
    else:
        raise ValueError(f"不支持的映射方式：{rule.mode}")

    return [
        ExpandedMapping(
            source_file=rule.source_file,
            source_sheet=rule.source_sheet,
            source_cell=source,
            target_file=rule.target_file,
            target_sheet=rule.target_sheet,
            target_cell=target,
        )
        for source, target in pairs
    ]


def expand_rules(rules: list[MappingRule]) -> list[ExpandedMapping]:
    expanded: list[ExpandedMapping] = []
    for index, rule in enumerate(rules, start=1):
        try:
            expanded.extend(expand_rule(rule))
        except ValueError as exc:
            raise ValueError(f"第 {index} 条映射：{exc}") from exc
    return expanded


class WorkbookReader:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.is_xls = path.suffix.lower() == ".xls"
        if self.is_xls:
            self.xls_book = xlrd.open_workbook(
                path,
                formatting_info=True,
                on_demand=True,
            )
            self.value_book = None
            self.format_book = None
        else:
            keep_vba = path.suffix.lower() == ".xlsm"
            self.xls_book = None
            self.value_book = load_workbook(
                path,
                data_only=True,
                read_only=False,
                keep_vba=keep_vba,
            )
            self.format_book = load_workbook(
                path,
                data_only=False,
                read_only=False,
                keep_vba=keep_vba,
            )

    def read(self, sheet_name: str, address: str) -> CellValue:
        if self.is_xls:
            if sheet_name not in self.xls_book.sheet_names():
                raise ValueError(f"{self.path.name} 中不存在工作表“{sheet_name}”。")
            sheet = self.xls_book.sheet_by_name(sheet_name)
            row, column = split_address(address)
            if row >= sheet.nrows or column >= sheet.ncols:
                return CellValue(None, "General")
            cell = sheet.cell(row, column)
            return CellValue(
                xls_cell_value(self.xls_book, cell),
                xls_number_format(self.xls_book, cell),
            )

        if sheet_name not in self.value_book.sheetnames:
            raise ValueError(f"{self.path.name} 中不存在工作表“{sheet_name}”。")
        return CellValue(
            self.value_book[sheet_name][address].value,
            self.format_book[sheet_name][address].number_format,
        )

    def close(self) -> None:
        if self.xls_book is not None:
            self.xls_book.release_resources()
        if self.value_book is not None:
            self.value_book.close()
        if self.format_book is not None:
            self.format_book.close()


def split_address(address: str) -> tuple[int, int]:
    addresses = parse_cell_addresses(address)
    if len(addresses) != 1:
        raise ValueError(f"单元格地址不合法：{address}")
    normalized = addresses[0]
    column_text = "".join(char for char in normalized if char.isalpha())
    row_text = "".join(char for char in normalized if char.isdigit())
    return int(row_text) - 1, column_letters_to_number(column_text) - 1


def validate_mapping_plan(
    source_files: list[Path],
    target_files: list[Path],
    rules: list[MappingRule],
) -> tuple[list[ExpandedMapping], list[str]]:
    if not source_files:
        raise ValueError("请至少添加一个来源工作簿。")
    if not target_files:
        raise ValueError("请至少添加一个目标工作簿。")
    if not rules:
        raise ValueError("请至少添加一条映射规则。")

    source_set = {str(path.resolve()) for path in source_files}
    target_set = {str(path.resolve()) for path in target_files}
    for path in [*source_files, *target_files]:
        if not path.is_file() or path.suffix.lower() not in WORKBOOK_SUFFIXES:
            raise ValueError(f"文件不存在或格式不支持：{path}")

    expanded = expand_rules(rules)
    seen_targets: dict[tuple[str, str, str], int] = {}
    warnings: list[str] = []
    readers: dict[str, WorkbookReader] = {}
    try:
        for index, mapping in enumerate(expanded, start=1):
            source_key = str(Path(mapping.source_file).resolve())
            target_key = str(Path(mapping.target_file).resolve())
            if source_key not in source_set:
                raise ValueError(f"第 {index} 项使用了未添加的来源文件。")
            if target_key not in target_set:
                raise ValueError(f"第 {index} 项使用了未添加的目标文件。")

            source_reader = readers.get(source_key)
            if source_reader is None:
                source_reader = WorkbookReader(Path(mapping.source_file))
                readers[source_key] = source_reader
            source_reader.read(mapping.source_sheet, mapping.source_cell)

            target_reader = readers.get(target_key)
            if target_reader is None:
                target_reader = WorkbookReader(Path(mapping.target_file))
                readers[target_key] = target_reader
            existing = target_reader.read(
                mapping.target_sheet,
                mapping.target_cell,
            ).value
            if existing not in (None, ""):
                warnings.append(
                    f"{Path(mapping.target_file).name}/"
                    f"{mapping.target_sheet}/{mapping.target_cell} 已有内容"
                )

            destination = (
                target_key,
                mapping.target_sheet,
                mapping.target_cell,
            )
            if destination in seen_targets:
                previous = seen_targets[destination]
                raise ValueError(
                    f"第 {previous} 项和第 {index} 项写入同一个目标单元格："
                    f"{Path(mapping.target_file).name}/"
                    f"{mapping.target_sheet}/{mapping.target_cell}"
                )
            seen_targets[destination] = index
    finally:
        for reader in readers.values():
            reader.close()
    return expanded, warnings


def output_name_for_target(path: Path) -> str:
    suffix = ".xlsx" if path.suffix.lower() == ".xls" else path.suffix
    return f"{path.stem}_已映射{suffix}"


def execute_mapping_plan(
    source_files: list[Path],
    target_files: list[Path],
    rules: list[MappingRule],
    output_folder: Path,
    progress: Callable[[int, int, str], None] | None = None,
) -> list[Path]:
    expanded, _warnings = validate_mapping_plan(
        source_files,
        target_files,
        rules,
    )
    output_folder.mkdir(parents=True, exist_ok=True)
    readers: dict[str, WorkbookReader] = {}
    target_books: dict[str, object] = {}
    output_paths: dict[str, Path] = {}
    try:
        for target_path in target_files:
            key = str(target_path.resolve())
            output_path = unique_output_path(
                output_folder / output_name_for_target(target_path)
            )
            if target_path.suffix.lower() == ".xls":
                workbook = copy_xls_to_xlsx(target_path)
            else:
                shutil.copy2(target_path, output_path)
                workbook = load_workbook(
                    output_path,
                    keep_vba=target_path.suffix.lower() == ".xlsm",
                )
            target_books[key] = workbook
            output_paths[key] = output_path

        total = len(expanded)
        for index, mapping in enumerate(expanded, start=1):
            if progress:
                progress(
                    index - 1,
                    total,
                    f"正在映射：{Path(mapping.source_file).name}/"
                    f"{mapping.source_cell} → "
                    f"{Path(mapping.target_file).name}/{mapping.target_cell}",
                )
            source_key = str(Path(mapping.source_file).resolve())
            reader = readers.get(source_key)
            if reader is None:
                reader = WorkbookReader(Path(mapping.source_file))
                readers[source_key] = reader
            cell_data = reader.read(mapping.source_sheet, mapping.source_cell)

            target_key = str(Path(mapping.target_file).resolve())
            target_book = target_books[target_key]
            if mapping.target_sheet not in target_book.sheetnames:
                raise ValueError(
                    f"{Path(mapping.target_file).name} 中不存在工作表"
                    f"“{mapping.target_sheet}”。"
                )
            target_cell = target_book[mapping.target_sheet][mapping.target_cell]
            target_cell.value = cell_data.value
            target_cell.number_format = cell_data.number_format
            if progress:
                progress(index, total, f"已完成 {index}/{total} 项映射")

        for key, workbook in target_books.items():
            workbook.save(output_paths[key])
            workbook.close()
        target_books.clear()
        return list(output_paths.values())
    finally:
        for reader in readers.values():
            reader.close()
        for workbook in target_books.values():
            workbook.close()


def save_mapping_project(
    path: Path,
    source_files: list[Path],
    target_files: list[Path],
    rules: list[MappingRule],
    output_folder: Path | None,
) -> None:
    payload = {
        "version": 1,
        "source_files": [str(item) for item in source_files],
        "target_files": [str(item) for item in target_files],
        "rules": [asdict(rule) for rule in rules],
        "output_folder": str(output_folder) if output_folder else "",
    }
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_mapping_project(
    path: Path,
) -> tuple[list[Path], list[Path], list[MappingRule], Path | None]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("version") != 1:
        raise ValueError("不支持的映射方案版本。")
    return (
        [Path(item) for item in payload.get("source_files", [])],
        [Path(item) for item in payload.get("target_files", [])],
        [MappingRule(**item) for item in payload.get("rules", [])],
        Path(payload["output_folder"])
        if payload.get("output_folder")
        else None,
    )


class MappingDialog:
    def __init__(
        self,
        parent,
        source_files: list[Path],
        target_files: list[Path],
        initial: MappingRule | None = None,
    ) -> None:
        self.result: MappingRule | None = None
        self.source_files = source_files
        self.target_files = target_files
        self.window = Toplevel(parent)
        self.window.title("设置映射关系")
        self.window.geometry("650x430")
        self.window.resizable(False, False)
        self.window.transient(parent)
        self.window.grab_set()

        self.source_file = StringVar(
            value=initial.source_file if initial else str(source_files[0])
        )
        self.source_sheet = StringVar(
            value=initial.source_sheet if initial else ""
        )
        self.source_cells = StringVar(
            value=", ".join(initial.source_cells) if initial else ""
        )
        self.target_file = StringVar(
            value=initial.target_file if initial else str(target_files[0])
        )
        self.target_sheet = StringVar(
            value=initial.target_sheet if initial else ""
        )
        self.target_cells = StringVar(
            value=", ".join(initial.target_cells) if initial else ""
        )
        self.mode = StringVar(value=initial.mode if initial else MODE_SEQUENCE)

        container = ttk.Frame(self.window, padding=16)
        container.pack(fill=BOTH, expand=True)
        self.source_sheet_combo = self._location_group(
            container,
            "读取位置",
            self.source_file,
            self.source_sheet,
            self.source_cells,
            source_files,
            0,
        )
        self.target_sheet_combo = self._location_group(
            container,
            "写入位置",
            self.target_file,
            self.target_sheet,
            self.target_cells,
            target_files,
            1,
        )

        mode_frame = ttk.LabelFrame(container, text="映射方式", padding=10)
        mode_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Combobox(
            mode_frame,
            textvariable=self.mode,
            values=MAPPING_MODES,
            state="readonly",
        ).pack(fill=X)

        actions = ttk.Frame(container)
        actions.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(14, 0))
        ttk.Button(actions, text="保存映射", command=self._save).pack(side=LEFT)
        ttk.Button(
            actions,
            text="取消",
            command=self.window.destroy,
        ).pack(side=LEFT, padx=(8, 0))
        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)

        self._refresh_source_sheets()
        self._refresh_target_sheets()
        parent.wait_window(self.window)

    def _location_group(
        self,
        parent,
        title: str,
        file_var: StringVar,
        sheet_var: StringVar,
        cells_var: StringVar,
        files: list[Path],
        column: int,
    ):
        frame = ttk.LabelFrame(parent, text=title, padding=10)
        frame.grid(
            row=0,
            column=column,
            sticky="nsew",
            padx=(0, 5) if column == 0 else (5, 0),
        )
        ttk.Label(frame, text="工作簿：").pack(anchor=W)
        file_combo = ttk.Combobox(
            frame,
            textvariable=file_var,
            values=[str(path) for path in files],
            state="readonly",
        )
        file_combo.pack(fill=X, pady=(2, 8))
        ttk.Label(frame, text="工作表：").pack(anchor=W)
        sheet_combo = ttk.Combobox(
            frame,
            textvariable=sheet_var,
            state="readonly",
        )
        sheet_combo.pack(fill=X, pady=(2, 8))
        ttk.Label(frame, text="单元格（逗号分隔）：").pack(anchor=W)
        ttk.Entry(frame, textvariable=cells_var).pack(fill=X, pady=(2, 0))
        if column == 0:
            file_combo.bind(
                "<<ComboboxSelected>>",
                lambda _event: self._refresh_source_sheets(),
            )
        else:
            file_combo.bind(
                "<<ComboboxSelected>>",
                lambda _event: self._refresh_target_sheets(),
            )
        return sheet_combo

    def _refresh_source_sheets(self) -> None:
        self._refresh_sheets(
            Path(self.source_file.get()),
            self.source_sheet,
            self.source_sheet_combo,
        )

    def _refresh_target_sheets(self) -> None:
        self._refresh_sheets(
            Path(self.target_file.get()),
            self.target_sheet,
            self.target_sheet_combo,
        )

    @staticmethod
    def _refresh_sheets(path: Path, variable: StringVar, combo) -> None:
        try:
            names = workbook_sheet_names(path)
            combo.configure(values=names)
            if variable.get() not in names:
                variable.set(names[0] if names else "")
        except Exception as exc:
            messagebox.showerror("无法读取工作表", str(exc), parent=combo)

    def _save(self) -> None:
        try:
            rule = MappingRule(
                source_file=self.source_file.get(),
                source_sheet=self.source_sheet.get(),
                source_cells=parse_cell_addresses(self.source_cells.get()),
                target_file=self.target_file.get(),
                target_sheet=self.target_sheet.get(),
                target_cells=parse_cell_addresses(self.target_cells.get()),
                mode=self.mode.get(),
            )
            expand_rule(rule)
            self.result = rule
            self.window.destroy()
        except ValueError as exc:
            messagebox.showerror("映射设置错误", str(exc), parent=self.window)


class MapperApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("Excel 单元格映射工具")
        self.root.geometry("1100x760")
        self.root.minsize(900, 650)
        self.source_files: list[Path] = []
        self.target_files: list[Path] = []
        self.rules: list[MappingRule] = []
        self.output_folder = StringVar()
        self.status = StringVar(value="请添加来源工作簿和目标工作簿。")
        self.progress_text = StringVar(value="0%")
        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill=BOTH, expand=True)

        file_row = ttk.Frame(container)
        file_row.pack(fill=X)
        file_row.columnconfigure(0, weight=1, uniform="files")
        file_row.columnconfigure(1, weight=1, uniform="files")
        self.source_tree = self._file_panel(
            file_row,
            "1. 来源工作簿",
            0,
            self._add_sources,
            self._add_source_folder,
            self._remove_sources,
        )
        self.target_tree = self._file_panel(
            file_row,
            "2. 目标工作簿",
            1,
            self._add_targets,
            self._add_target_folder,
            self._remove_targets,
        )

        mapping_frame = ttk.LabelFrame(
            container,
            text="3. 映射关系",
            padding=10,
        )
        mapping_frame.pack(fill=BOTH, expand=True, pady=(10, 0))
        self.mapping_tree = ttk.Treeview(
            mapping_frame,
            columns=("source", "target", "mode"),
            show="headings",
            height=10,
        )
        self.mapping_tree.heading("source", text="来源位置")
        self.mapping_tree.heading("target", text="目标位置")
        self.mapping_tree.heading("mode", text="映射方式")
        self.mapping_tree.column("source", width=370)
        self.mapping_tree.column("target", width=370)
        self.mapping_tree.column("mode", width=150)
        self.mapping_tree.pack(fill=BOTH, expand=True)
        mapping_actions = ttk.Frame(mapping_frame)
        mapping_actions.pack(fill=X, pady=(8, 0))
        for text, command in (
            ("添加映射", self._add_rule),
            ("编辑选中", self._edit_rule),
            ("复制选中", self._duplicate_rule),
            ("删除选中", self._delete_rule),
            ("清空映射", self._clear_rules),
            ("展开预览", self._preview_expanded),
        ):
            ttk.Button(
                mapping_actions,
                text=text,
                command=command,
            ).pack(side=LEFT, padx=(0, 6))

        output_frame = ttk.LabelFrame(
            container,
            text="4. 输出与执行",
            padding=10,
        )
        output_frame.pack(fill=X, pady=(10, 0))
        location_row = ttk.Frame(output_frame)
        location_row.pack(fill=X)
        ttk.Label(location_row, text="副本输出文件夹：").pack(side=LEFT)
        ttk.Entry(
            location_row,
            textvariable=self.output_folder,
        ).pack(side=LEFT, fill=X, expand=True)
        ttk.Button(
            location_row,
            text="浏览...",
            command=self._browse_output,
        ).pack(side=LEFT, padx=(8, 0))

        progress_row = ttk.Frame(output_frame)
        progress_row.pack(fill=X, pady=(8, 0))
        self.progress = ttk.Progressbar(progress_row, mode="determinate")
        self.progress.pack(side=LEFT, fill=X, expand=True)
        ttk.Label(
            progress_row,
            textvariable=self.progress_text,
        ).pack(side=RIGHT, padx=(8, 0))
        ttk.Label(output_frame, textvariable=self.status).pack(
            anchor=W,
            pady=(5, 0),
        )

        actions = ttk.Frame(container)
        actions.pack(fill=X, pady=(10, 0))
        for text, command in (
            ("保存方案", self._save_project),
            ("载入方案", self._load_project),
            ("预检查", self._precheck),
            ("开始映射", self._run),
            ("退出", self.root.destroy),
        ):
            ttk.Button(actions, text=text, command=command).pack(
                side=LEFT,
                padx=(0, 8),
            )

    def _file_panel(
        self,
        parent,
        title: str,
        column: int,
        add_command,
        folder_command,
        remove_command,
    ):
        frame = ttk.LabelFrame(parent, text=title, padding=10)
        frame.grid(
            row=0,
            column=column,
            sticky="nsew",
            padx=(0, 5) if column == 0 else (5, 0),
        )
        tree = ttk.Treeview(
            frame,
            columns=("path",),
            show="headings",
            height=5,
        )
        tree.heading("path", text="文件路径")
        tree.column("path", width=430)
        tree.pack(fill=BOTH, expand=True)
        actions = ttk.Frame(frame)
        actions.pack(fill=X, pady=(6, 0))
        ttk.Button(actions, text="添加文件", command=add_command).pack(side=LEFT)
        ttk.Button(
            actions,
            text="添加文件夹",
            command=folder_command,
        ).pack(side=LEFT, padx=(6, 0))
        ttk.Button(
            actions,
            text="删除选中",
            command=remove_command,
        ).pack(side=LEFT, padx=(6, 0))
        return tree

    @staticmethod
    def _choose_workbooks() -> list[Path]:
        selected = filedialog.askopenfilenames(
            title="选择 Excel 工作簿",
            filetypes=[
                ("Excel 文件", "*.xls *.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        return [Path(item) for item in selected]

    def _add_sources(self) -> None:
        self._add_files(self.source_files, self.source_tree)

    def _add_targets(self) -> None:
        self._add_files(self.target_files, self.target_tree)

    def _add_source_folder(self) -> None:
        self._add_folder(self.source_files, self.source_tree)

    def _add_target_folder(self) -> None:
        self._add_folder(self.target_files, self.target_tree)

    def _add_files(self, collection: list[Path], tree) -> None:
        self._append_files(collection, tree, self._choose_workbooks())

    def _add_folder(self, collection: list[Path], tree) -> None:
        selected = filedialog.askdirectory(title="选择包含 Excel 文件的文件夹")
        if not selected:
            return
        files = workbook_files_in_folder(Path(selected))
        if not files:
            messagebox.showinfo(
                "没有可添加的文件",
                "所选文件夹中没有 .xls、.xlsx 或 .xlsm 文件。",
                parent=self.root,
            )
            return
        if not messagebox.askyesno(
            "确认批量添加",
            f"发现 {len(files)} 个 Excel 文件，是否全部添加？",
            parent=self.root,
        ):
            return
        self._append_files(collection, tree, files)

    def _append_files(self, collection: list[Path], tree, files) -> None:
        for path in files:
            if path not in collection:
                collection.append(path)
                tree.insert("", END, values=(str(path),))
        if not self.output_folder.get() and self.target_files:
            self.output_folder.set(str(self.target_files[0].parent / "映射结果"))

    def _remove_sources(self) -> None:
        self._remove_files(self.source_files, self.source_tree)

    def _remove_targets(self) -> None:
        self._remove_files(self.target_files, self.target_tree)

    @staticmethod
    def _remove_files(collection: list[Path], tree) -> None:
        selected = tree.selection()
        paths = {Path(tree.item(item, "values")[0]) for item in selected}
        collection[:] = [path for path in collection if path not in paths]
        for item in selected:
            tree.delete(item)

    def _selected_rule_index(self) -> int | None:
        selection = self.mapping_tree.selection()
        if not selection:
            return None
        return self.mapping_tree.index(selection[0])

    def _add_rule(self) -> None:
        if not self.source_files or not self.target_files:
            messagebox.showinfo(
                "请先添加文件",
                "请先添加至少一个来源工作簿和一个目标工作簿。",
                parent=self.root,
            )
            return
        dialog = MappingDialog(
            self.root,
            self.source_files,
            self.target_files,
        )
        if dialog.result:
            self.rules.append(dialog.result)
            self._refresh_rules()

    def _edit_rule(self) -> None:
        index = self._selected_rule_index()
        if index is None:
            return
        dialog = MappingDialog(
            self.root,
            self.source_files,
            self.target_files,
            self.rules[index],
        )
        if dialog.result:
            self.rules[index] = dialog.result
            self._refresh_rules()

    def _duplicate_rule(self) -> None:
        index = self._selected_rule_index()
        if index is None:
            return
        self.rules.insert(index + 1, MappingRule(**asdict(self.rules[index])))
        self._refresh_rules()

    def _delete_rule(self) -> None:
        index = self._selected_rule_index()
        if index is not None:
            self.rules.pop(index)
            self._refresh_rules()

    def _clear_rules(self) -> None:
        if self.rules and messagebox.askyesno(
            "清空映射",
            "确定清空全部映射规则吗？",
            parent=self.root,
        ):
            self.rules.clear()
            self._refresh_rules()

    def _refresh_rules(self) -> None:
        self.mapping_tree.delete(*self.mapping_tree.get_children())
        for rule in self.rules:
            source = (
                f"{Path(rule.source_file).name} / {rule.source_sheet} / "
                f"{', '.join(rule.source_cells)}"
            )
            target = (
                f"{Path(rule.target_file).name} / {rule.target_sheet} / "
                f"{', '.join(rule.target_cells)}"
            )
            self.mapping_tree.insert(
                "",
                END,
                values=(source, target, rule.mode),
            )

    def _preview_expanded(self) -> None:
        try:
            expanded = expand_rules(self.rules)
        except ValueError as exc:
            messagebox.showerror("无法展开映射", str(exc), parent=self.root)
            return
        if not expanded:
            messagebox.showinfo(
                "没有映射",
                "请先添加映射规则。",
                parent=self.root,
            )
            return

        window = Toplevel(self.root)
        window.title(f"映射展开预览（共 {len(expanded)} 项）")
        window.geometry("900x520")
        window.transient(self.root)
        container = ttk.Frame(window, padding=12)
        container.pack(fill=BOTH, expand=True)
        tree = ttk.Treeview(
            container,
            columns=("number", "source", "target"),
            show="headings",
        )
        tree.heading("number", text="序号")
        tree.heading("source", text="来源位置")
        tree.heading("target", text="目标位置")
        tree.column("number", width=60, anchor="center")
        tree.column("source", width=390)
        tree.column("target", width=390)
        scroll = ttk.Scrollbar(
            container,
            orient="vertical",
            command=tree.yview,
        )
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.pack(side=RIGHT, fill="y")
        for index, mapping in enumerate(expanded, start=1):
            source = (
                f"{Path(mapping.source_file).name} / "
                f"{mapping.source_sheet} / {mapping.source_cell}"
            )
            target = (
                f"{Path(mapping.target_file).name} / "
                f"{mapping.target_sheet} / {mapping.target_cell}"
            )
            tree.insert("", END, values=(index, source, target))

    def _browse_output(self) -> None:
        selected = filedialog.askdirectory(title="选择副本输出文件夹")
        if selected:
            self.output_folder.set(selected)

    def _save_project(self) -> None:
        selected = filedialog.asksaveasfilename(
            title="保存映射方案",
            defaultextension=".json",
            filetypes=[("映射方案", "*.json")],
        )
        if selected:
            save_mapping_project(
                Path(selected),
                self.source_files,
                self.target_files,
                self.rules,
                Path(self.output_folder.get())
                if self.output_folder.get()
                else None,
            )

    def _load_project(self) -> None:
        selected = filedialog.askopenfilename(
            title="载入映射方案",
            filetypes=[("映射方案", "*.json")],
        )
        if not selected:
            return
        try:
            (
                self.source_files,
                self.target_files,
                self.rules,
                output_folder,
            ) = load_mapping_project(Path(selected))
            self.output_folder.set(str(output_folder) if output_folder else "")
            self._refresh_file_tree(self.source_tree, self.source_files)
            self._refresh_file_tree(self.target_tree, self.target_files)
            self._refresh_rules()
        except Exception as exc:
            messagebox.showerror("载入失败", str(exc), parent=self.root)

    @staticmethod
    def _refresh_file_tree(tree, files: list[Path]) -> None:
        tree.delete(*tree.get_children())
        for path in files:
            tree.insert("", END, values=(str(path),))

    def _precheck(self) -> bool:
        try:
            expanded, warnings = validate_mapping_plan(
                self.source_files,
                self.target_files,
                self.rules,
            )
            message = (
                f"来源工作簿：{len(self.source_files)} 个\n"
                f"目标工作簿：{len(self.target_files)} 个\n"
                f"映射规则：{len(self.rules)} 条\n"
                f"展开后写入：{len(expanded)} 项"
            )
            if warnings:
                preview = "\n".join(warnings[:8])
                message += (
                    f"\n\n有 {len(warnings)} 个目标单元格已有内容，"
                    "将在副本中覆盖：\n"
                    f"{preview}"
                )
                if len(warnings) > 8:
                    message += "\n..."
            messagebox.showinfo("预检查通过", message, parent=self.root)
            return True
        except Exception as exc:
            messagebox.showerror("预检查失败", str(exc), parent=self.root)
            return False

    def _progress(self, current: int, total: int, message: str) -> None:
        maximum = max(total, 1)
        self.progress.configure(maximum=maximum, value=current)
        self.progress_text.set(f"{int(current / maximum * 100)}%")
        self.status.set(message)
        self.root.update_idletasks()

    def _run(self) -> None:
        if not self.output_folder.get().strip():
            messagebox.showerror(
                "缺少输出位置",
                "请选择副本输出文件夹。",
                parent=self.root,
            )
            return
        try:
            expanded, warnings = validate_mapping_plan(
                self.source_files,
                self.target_files,
                self.rules,
            )
            if warnings and not messagebox.askyesno(
                "确认覆盖副本中的内容",
                f"有 {len(warnings)} 个目标单元格已有内容，"
                "将在新副本中覆盖。\n\n原文件不会修改。是否继续？",
                parent=self.root,
            ):
                return
            self.progress.configure(maximum=max(len(expanded), 1), value=0)
            self.progress_text.set("0%")
            outputs = execute_mapping_plan(
                self.source_files,
                self.target_files,
                self.rules,
                Path(self.output_folder.get()),
                self._progress,
            )
            self.progress.configure(value=self.progress.cget("maximum"))
            self.progress_text.set("100%")
            self.status.set("映射完成；原文件未修改。")
            messagebox.showinfo(
                "映射完成",
                f"已生成 {len(outputs)} 个工作簿副本：\n"
                + "\n".join(str(path) for path in outputs),
                parent=self.root,
            )
        except Exception as exc:
            self.status.set("映射失败。")
            messagebox.showerror("映射失败", str(exc), parent=self.root)


def main() -> None:
    root = Tk()
    MapperApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

