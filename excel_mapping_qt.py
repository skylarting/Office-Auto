"""Modern PySide6 interface for the Excel cell mapping engine."""

from __future__ import annotations

from dataclasses import asdict
from pathlib import Path
import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QColor, QKeySequence, QShortcut
from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from excel_mapper import (
    MODE_MANUAL,
    MappingRule,
    WorkbookReader,
    execute_mapping_plan,
    expand_rules,
    format_cell_addresses,
    format_project_path,
    infer_mapping_mode,
    load_excel_mapping_scheme,
    mapping_rule_is_blank,
    parse_cell_addresses,
    resolve_project_path,
    save_excel_mapping_scheme,
    spreadsheet_automation_provider,
    validate_mapping_plan,
    workbook_files_in_folder,
    workbook_sheet_names,
    xls_formula_risk,
)


ROLE_GROUP = Qt.ItemDataRole.UserRole
ROLE_KIND = Qt.ItemDataRole.UserRole + 1
ROLE_DRAFT = Qt.ItemDataRole.UserRole + 2


class QtCellPickerDialog(QDialog):
    """Spreadsheet-like picker that preserves selection order."""

    def __init__(
        self,
        parent,
        workbook_path: Path,
        sheet_name: str,
        initial_cells: list[str],
        same_position_cells: list[str] | None = None,
    ) -> None:
        super().__init__(parent)
        self.reader = WorkbookReader(workbook_path)
        self.sheet_name = sheet_name
        self.result: list[str] | None = None
        self.selection_order = list(dict.fromkeys(initial_cells))
        self.previous_selection = set(self.selection_order)
        self.same_position_cells = same_position_cells
        rows, columns = self.reader.used_dimensions(sheet_name)
        self.rows = max(rows, 30)
        self.columns = max(columns, 12)

        self.setWindowTitle(
            f"选择单元格 — {workbook_path.name} / {sheet_name}"
        )
        self.resize(1080, 720)
        layout = QVBoxLayout(self)

        top = QHBoxLayout()
        top.addWidget(QLabel("已选择（按顺序）："))
        self.selected_label = QLabel()
        self.selected_label.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse
        )
        top.addWidget(self.selected_label, 1)
        self.jump_edit = QLineEdit()
        self.jump_edit.setPlaceholderText("例如 C21")
        self.jump_edit.setMaximumWidth(120)
        top.addWidget(QLabel("定位到："))
        top.addWidget(self.jump_edit)
        jump = QPushButton("跳转")
        jump.clicked.connect(self.jump_to_cell)
        top.addWidget(jump)
        layout.addLayout(top)

        batch = QHBoxLayout()
        batch.addWidget(QLabel("批量选择："))
        for text, callback in (
            ("清空", self.clear_selection),
            ("全选", self.select_all_nonempty),
            ("反选", self.invert_nonempty),
        ):
            button = QPushButton(text)
            button.clicked.connect(callback)
            batch.addWidget(button)
        batch.addStretch(1)
        for text, trait in (
            ("公式型", "formula"),
            ("数值（非公式）", "number"),
            ("带底色", "fill"),
            ("文字型", "text"),
        ):
            button = QPushButton(text)
            button.clicked.connect(
                lambda _checked=False, value=trait: self.toggle_trait(value)
            )
            batch.addWidget(button)
        layout.addLayout(batch)

        self.table = QTableWidget(self.rows, self.columns)
        self.table.setSelectionMode(
            QTableWidget.SelectionMode.MultiSelection
        )
        self.table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectItems
        )
        self.table.setStyleSheet(
            "QTableWidget::item:selected {"
            "  border: 2px solid #D93025;"
            "  background-color: transparent;"
            "  color: palette(text);"
            "}"
        )
        self.table.horizontalHeader().setDefaultSectionSize(125)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.itemSelectionChanged.connect(self.selection_changed)
        self._populate()
        layout.addWidget(self.table, 1)

        actions = QHBoxLayout()
        confirm = QPushButton("确定选择")
        confirm.setDefault(True)
        confirm.clicked.connect(self.confirm)
        actions.addWidget(confirm)
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        actions.addWidget(cancel)
        actions.addStretch(1)
        if same_position_cells is not None:
            same = QPushButton("与来源同位置")
            same.clicked.connect(self.use_same_position)
            actions.addWidget(same)
        layout.addLayout(actions)

        self._restore_initial_selection()
        self.update_selected_label()

    def _populate(self) -> None:
        self.table.blockSignals(True)
        try:
            for row in range(self.rows):
                for column in range(self.columns):
                    address = self._address(row, column)
                    try:
                        cell = self.reader.read(self.sheet_name, address)
                        value = cell.value
                        fill = self.reader.fill_color(
                            self.sheet_name, address
                        )
                    except Exception:
                        value, fill = "", None
                    item = QTableWidgetItem(
                        "" if value is None else str(value)
                    )
                    item.setToolTip(f"{address}: {item.text()}")
                    if fill:
                        item.setBackground(QColor(fill))
                    self.table.setItem(row, column, item)
        finally:
            self.table.blockSignals(False)

    def _restore_initial_selection(self) -> None:
        self.table.blockSignals(True)
        try:
            for address in self.selection_order:
                row, column = self._indices(address)
                if row < self.rows and column < self.columns:
                    self.table.item(row, column).setSelected(True)
        finally:
            self.table.blockSignals(False)

    @staticmethod
    def _address(row: int, column: int) -> str:
        result = ""
        number = column + 1
        while number:
            number, remainder = divmod(number - 1, 26)
            result = chr(65 + remainder) + result
        return f"{result}{row + 1}"

    @staticmethod
    def _indices(address: str) -> tuple[int, int]:
        letters = "".join(char for char in address if char.isalpha())
        digits = "".join(char for char in address if char.isdigit())
        column = 0
        for char in letters.upper():
            column = column * 26 + ord(char) - 64
        return int(digits) - 1, column - 1

    def selected_addresses(self) -> set[str]:
        return {
            self._address(item.row(), item.column())
            for item in self.table.selectedItems()
        }

    def selection_changed(self) -> None:
        current = self.selected_addresses()
        added = current - self.previous_selection
        removed = self.previous_selection - current
        if removed:
            self.selection_order = [
                address
                for address in self.selection_order
                if address not in removed
            ]
        if added:
            for item in self.table.selectedItems():
                address = self._address(item.row(), item.column())
                if address in added and address not in self.selection_order:
                    self.selection_order.append(address)
        self.previous_selection = current
        self.update_selected_label()

    def update_selected_label(self) -> None:
        if not self.selection_order:
            text = "尚未选择"
        elif len(self.selection_order) <= 40:
            text = format_cell_addresses(self.selection_order)
        else:
            text = (
                format_cell_addresses(self.selection_order[:12])
                + f"……（共 {len(self.selection_order)} 个）"
            )
        self.selected_label.setText(text)

    def _addresses_with_trait(self, trait: str) -> list[str]:
        result = []
        rows, columns = self.reader.used_dimensions(self.sheet_name)
        for column in range(columns):
            for row in range(rows):
                address = self._address(row, column)
                if trait in self.reader.cell_traits(
                    self.sheet_name, address
                ):
                    result.append(address)
        return result

    def _toggle_addresses(self, addresses: list[str]) -> None:
        if not addresses:
            return
        current = self.selected_addresses()
        remove = all(address in current for address in addresses)
        self.table.blockSignals(True)
        try:
            for address in addresses:
                row, column = self._indices(address)
                if row >= self.rows or column >= self.columns:
                    continue
                self.table.item(row, column).setSelected(not remove)
                if remove:
                    if address in self.selection_order:
                        self.selection_order.remove(address)
                elif address not in self.selection_order:
                    self.selection_order.append(address)
        finally:
            self.table.blockSignals(False)
        self.previous_selection = self.selected_addresses()
        self.update_selected_label()

    def clear_selection(self) -> None:
        self.table.clearSelection()

    def select_all_nonempty(self) -> None:
        self._toggle_addresses(self._addresses_with_trait("nonempty"))

    def invert_nonempty(self) -> None:
        addresses = self._addresses_with_trait("nonempty")
        current = self.selected_addresses()
        self.table.blockSignals(True)
        try:
            for address in addresses:
                row, column = self._indices(address)
                selecting = address not in current
                self.table.item(row, column).setSelected(selecting)
                if selecting and address not in self.selection_order:
                    self.selection_order.append(address)
                elif not selecting and address in self.selection_order:
                    self.selection_order.remove(address)
        finally:
            self.table.blockSignals(False)
        self.previous_selection = self.selected_addresses()
        self.update_selected_label()

    def toggle_trait(self, trait: str) -> None:
        self._toggle_addresses(self._addresses_with_trait(trait))

    def jump_to_cell(self) -> None:
        try:
            address = parse_cell_addresses(self.jump_edit.text())[0]
            row, column = self._indices(address)
            self.table.scrollToItem(self.table.item(row, column))
            self.table.setCurrentCell(row, column)
        except Exception as exc:
            QMessageBox.warning(self, "无法跳转", str(exc))

    def use_same_position(self) -> None:
        if not self.same_position_cells:
            QMessageBox.information(
                self, "尚未设置来源", "请先设置本组的来源单元格。"
            )
            return
        self.result = list(self.same_position_cells)
        self.accept()

    def confirm(self) -> None:
        if not self.selection_order:
            QMessageBox.information(
                self, "尚未选择", "请至少选择一个单元格。"
            )
            return
        self.result = list(self.selection_order)
        self.accept()

    def done(self, result: int) -> None:
        self.reader.close()
        super().done(result)


class ExcelMappingQtWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.rules: list[MappingRule] = []
        self.copied_groups: list[MappingRule] = []
        self.copied_rows: list[tuple[str, str, list[str]]] = []
        self.clipboard_scope = ""
        self.base_folder: Path | None = None
        self.output_folder: Path | None = None
        self._refreshing = False

        self.setWindowTitle("Excel 单元格映射工具")
        self.resize(1280, 800)
        self.setMinimumSize(920, 620)
        self._build_ui()
        self.refresh_table()

    def _build_ui(self) -> None:
        toolbar = QToolBar("方案操作", self)
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        for text, callback in (
            ("导入映射表", self.import_scheme),
            ("导出映射表", self.export_scheme),
            ("清空所有映射", self.clear_all),
            ("操作说明", self.show_help),
        ):
            action = QAction(text, self)
            action.triggered.connect(callback)
            toolbar.addAction(action)
            if text == "导出映射表":
                toolbar.addSeparator()

        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(16, 12, 16, 12)

        folders = QGridLayout()
        self.base_edit = QLineEdit()
        self.base_edit.setReadOnly(True)
        self.base_edit.setPlaceholderText("点击选择默认文件夹")
        self.base_edit.mousePressEvent = (
            lambda event: self.choose_base_folder()
        )
        self.output_edit = QLineEdit()
        self.output_edit.setReadOnly(True)
        self.output_edit.setPlaceholderText("点击选择输出文件夹")
        self.output_edit.mousePressEvent = (
            lambda event: self.choose_output_folder()
        )
        folders.addWidget(QLabel("默认文件夹："), 0, 0)
        folders.addWidget(self.base_edit, 0, 1)
        folders.addWidget(QLabel("输出文件夹："), 1, 0)
        folders.addWidget(self.output_edit, 1, 1)
        layout.addLayout(folders)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(
            ["映射组", "类型", "工作簿", "工作表", "单元格"]
        )
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)
        header.setMinimumSectionSize(60)
        for column, width in enumerate((90, 90, 420, 300, 420)):
            self.table.setColumnWidth(column, width)
        self.table.setSelectionMode(
            QTableWidget.SelectionMode.ExtendedSelection
        )
        self.table.setSelectionBehavior(
            QTableWidget.SelectionBehavior.SelectItems
        )
        self.table.setAlternatingRowColors(False)
        self.table.cellClicked.connect(self.cell_clicked)
        self.table.itemChanged.connect(self.item_changed)
        self.table.setContextMenuPolicy(
            Qt.ContextMenuPolicy.CustomContextMenu
        )
        self.table.customContextMenuRequested.connect(
            self.show_context_menu
        )
        layout.addWidget(self.table, 1)

        self.hint = QLabel(
            "单击“映射组”列选中整组，单击“类型”列选中单行；"
            "按住 Ctrl 可继续多选。Delete 删除选中组。\n"
            "Ctrl+C / Ctrl+V 会根据当前选中的是单行还是整组自动复制粘贴；"
            "最后一组为默认新增行。"
        )
        self.hint.setStyleSheet("color: #59636e;")
        layout.addWidget(self.hint)

        status_line = QHBoxLayout()
        self.status_label = QLabel("尚未检查。")
        status_line.addWidget(self.status_label, 1)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setMaximumWidth(280)
        status_line.addWidget(self.progress)
        layout.addLayout(status_line)

        actions = QHBoxLayout()
        preview = QPushButton("展开预览")
        preview.clicked.connect(self.preview)
        actions.addWidget(preview)
        check = QPushButton("预检查")
        check.clicked.connect(self.precheck)
        actions.addWidget(check)
        actions.addStretch(1)
        run = QPushButton("开始映射")
        run.setDefault(True)
        run.clicked.connect(self.run_mapping)
        actions.addWidget(run)
        close = QPushButton("退出")
        close.clicked.connect(self.close)
        actions.addWidget(close)
        layout.addLayout(actions)

        delete_shortcut = QShortcut(QKeySequence.Delete, self)
        delete_shortcut.setContext(Qt.ShortcutContext.WindowShortcut)
        delete_shortcut.activated.connect(
            self.delete_selected_groups
        )
        copy_shortcut = QShortcut(QKeySequence.Copy, self)
        copy_shortcut.setContext(Qt.ShortcutContext.WindowShortcut)
        copy_shortcut.activated.connect(self.copy_selection)
        paste_shortcut = QShortcut(QKeySequence.Paste, self)
        paste_shortcut.setContext(Qt.ShortcutContext.WindowShortcut)
        paste_shortcut.activated.connect(self.paste_selection)

    def active_rules(self) -> list[MappingRule]:
        return [rule for rule in self.rules if not mapping_rule_is_blank(rule)]

    def refresh_table(self) -> None:
        self._refreshing = True
        self.table.blockSignals(True)
        try:
            self.table.setRowCount((len(self.rules) + 1) * 2)
            for index in range(len(self.rules) + 1):
                rule = (
                    self.rules[index]
                    if index < len(self.rules)
                    else MappingRule("", "", [], "", "", [], MODE_MANUAL)
                )
                draft = index == len(self.rules)
                for offset, kind in enumerate(("source", "target")):
                    row = index * 2 + offset
                    values = self._row_values(index, rule, kind, draft)
                    background = QColor(
                        "#F2F2F2" if (index + 1) % 2 else "#FFFFFF"
                    )
                    for column, text in enumerate(values):
                        item = QTableWidgetItem(text)
                        item.setData(ROLE_GROUP, index)
                        item.setData(ROLE_KIND, kind)
                        item.setData(ROLE_DRAFT, draft)
                        item.setBackground(background)
                        if column < 2:
                            item.setFlags(
                                item.flags()
                                & ~Qt.ItemFlag.ItemIsEditable
                            )
                        self.table.setItem(row, column, item)
        finally:
            self.table.blockSignals(False)
            self._refreshing = False

    def _row_values(
        self, index: int, rule: MappingRule, kind: str, draft: bool
    ) -> list[str]:
        source = kind == "source"
        path = rule.source_file if source else rule.target_file
        sheet = rule.source_sheet if source else rule.target_sheet
        cells = rule.source_cells if source else rule.target_cells
        workbook_text = (
            format_project_path(Path(path), self.base_folder)
            if path
            else "单击选择"
        )
        sheet_text = sheet or "单击选择或输入"
        if source:
            cell_text = (
                format_cell_addresses(cells)
                if cells
                else "单击选择或输入"
            )
        elif not rule.source_cells and not cells:
            cell_text = "默认同位置；点击可修改"
        elif rule.source_cells and cells == rule.source_cells:
            cell_text = "同位置"
        else:
            cell_text = (
                format_cell_addresses(cells)
                if cells
                else "单击选择或输入"
            )
        return [
            str(index + 1),
            "来源" if source else "目标",
            workbook_text,
            sheet_text,
            cell_text,
        ]

    def ensure_rule(self, index: int) -> MappingRule:
        while len(self.rules) <= index:
            self.rules.append(
                MappingRule("", "", [], "", "", [], MODE_MANUAL)
            )
        return self.rules[index]

    def item_changed(self, item: QTableWidgetItem) -> None:
        if self._refreshing or item.column() < 2:
            return
        index = int(item.data(ROLE_GROUP))
        kind = str(item.data(ROLE_KIND))
        rule = self.ensure_rule(index)
        value = item.text().strip()
        try:
            self._set_rule_value(rule, kind, item.column(), value)
            self.refresh_table()
        except Exception as exc:
            QMessageBox.warning(self, "无法保存修改", str(exc))
            self.refresh_table()

    def _set_rule_value(
        self, rule: MappingRule, kind: str, column: int, value: str
    ) -> None:
        source = kind == "source"
        if column == 2:
            path = (
                ""
                if not value or value.startswith(("单击", "（未选择"))
                else str(
                    resolve_project_path(
                        value, self.base_folder or Path.cwd()
                    )
                )
            )
            if source:
                rule.source_file = path
            else:
                rule.target_file = path
        elif column == 3:
            sheet = (
                ""
                if not value or value.startswith(("单击", "（未选择"))
                else value
            )
            if source:
                rule.source_sheet = sheet
            else:
                rule.target_sheet = sheet
        elif column == 4:
            old_same = rule.target_cells == rule.source_cells
            if not source and value in (
                "同位置",
                "默认同位置；点击可修改",
            ):
                rule.target_cells = list(rule.source_cells)
            else:
                cells = parse_cell_addresses(value)
                if source:
                    rule.source_cells = cells
                    if old_same:
                        rule.target_cells = list(cells)
                else:
                    rule.target_cells = cells
            if rule.source_cells and rule.target_cells:
                rule.mode = infer_mapping_mode(
                    rule.source_cells, rule.target_cells
                )

    def cell_clicked(self, row: int, column: int) -> None:
        item = self.table.item(row, column)
        if item is None:
            return
        if column == 0:
            self.select_group(int(item.data(ROLE_GROUP)))
        elif column == 1:
            preserve = bool(
                QApplication.keyboardModifiers()
                & Qt.KeyboardModifier.ControlModifier
            )
            if not preserve:
                self.table.clearSelection()
            for cell_column in range(self.table.columnCount()):
                cell = self.table.item(row, cell_column)
                if cell is not None:
                    cell.setSelected(True)
        elif column == 2:
            self.choose_workbook(row)
        elif column == 3:
            self.choose_sheet(row)
        elif column == 4:
            self.choose_cells(row)

    def select_group(self, index: int) -> None:
        preserve = bool(
            QApplication.keyboardModifiers()
            & Qt.KeyboardModifier.ControlModifier
        )
        if not preserve:
            self.table.clearSelection()
        for row in (index * 2, index * 2 + 1):
            if row < self.table.rowCount():
                for column in range(self.table.columnCount()):
                    item = self.table.item(row, column)
                    if item is not None:
                        item.setSelected(True)

    def selected_group_indices(self) -> list[int]:
        indices = {
            int(item.data(ROLE_GROUP))
            for item in self.table.selectedItems()
            if not bool(item.data(ROLE_DRAFT))
        }
        return sorted(index for index in indices if index < len(self.rules))

    def choose_workbook(self, row: int) -> None:
        item = self.table.item(row, 2)
        index, kind = int(item.data(ROLE_GROUP)), str(item.data(ROLE_KIND))
        menu = QMenu(self)
        none = menu.addAction("（未选择工作簿）")
        candidates = (
            workbook_files_in_folder(self.base_folder)
            if self.base_folder and self.base_folder.is_dir()
            else []
        )
        candidate_actions = {}
        for path in candidates:
            candidate_actions[menu.addAction(path.name)] = path
        menu.addSeparator()
        browse = menu.addAction("浏览其他工作簿…")
        chosen = menu.exec(
            self.table.viewport().mapToGlobal(
                self.table.visualItemRect(item).bottomLeft()
            )
        )
        if chosen is None:
            return
        if chosen == browse:
            self.browse_workbook(row)
            return
        path = candidate_actions.get(chosen)
        value = "" if chosen == none else str(path.resolve())
        rule = self.ensure_rule(index)
        if kind == "source":
            rule.source_file = value
        else:
            rule.target_file = value
        self._auto_match_sheet(rule, kind)
        self.refresh_table()

    def browse_workbook(self, row: int) -> None:
        item = self.table.item(row, 2)
        index, kind = int(item.data(ROLE_GROUP)), str(item.data(ROLE_KIND))
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "选择 Excel 工作簿",
            str(self.base_folder or Path.cwd()),
            "Excel 工作簿 (*.xls *.xlsx *.xlsm)",
        )
        if not selected:
            return
        rule = self.ensure_rule(index)
        if kind == "source":
            rule.source_file = selected
        else:
            rule.target_file = selected
        self._auto_match_sheet(rule, kind)
        self.refresh_table()

    def _auto_match_sheet(self, rule: MappingRule, kind: str) -> None:
        source = kind == "source"
        path_text = rule.source_file if source else rule.target_file
        if not path_text:
            return
        names = workbook_sheet_names(Path(path_text))
        current = rule.source_sheet if source else rule.target_sheet
        if current not in names and names:
            if source:
                rule.source_sheet = names[0]
            else:
                rule.target_sheet = names[0]

    def choose_sheet(self, row: int) -> None:
        item = self.table.item(row, 3)
        index, kind = int(item.data(ROLE_GROUP)), str(item.data(ROLE_KIND))
        rule = self.ensure_rule(index)
        path_text = (
            rule.source_file if kind == "source" else rule.target_file
        )
        menu = QMenu(self)
        none = menu.addAction("（未选择工作表）")
        actions = {}
        if path_text and Path(path_text).is_file():
            for name in workbook_sheet_names(Path(path_text)):
                actions[menu.addAction(name)] = name
        menu.addSeparator()
        manual = menu.addAction("手动输入…")
        chosen = menu.exec(
            self.table.viewport().mapToGlobal(
                self.table.visualItemRect(item).bottomLeft()
            )
        )
        if chosen is None:
            return
        if chosen == manual:
            current = (
                rule.source_sheet
                if kind == "source"
                else rule.target_sheet
            )
            value, accepted = QInputDialog.getText(
                self, "手动输入工作表", "工作表名称：", text=current
            )
            if not accepted:
                return
            if kind == "source":
                rule.source_sheet = value.strip()
            else:
                rule.target_sheet = value.strip()
            self.refresh_table()
            return
        value = "" if chosen == none else actions.get(chosen, "")
        if kind == "source":
            rule.source_sheet = value
        else:
            rule.target_sheet = value
        self.refresh_table()

    def choose_cells(self, row: int) -> None:
        item = self.table.item(row, 4)
        index, kind = int(item.data(ROLE_GROUP)), str(item.data(ROLE_KIND))
        rule = self.ensure_rule(index)
        source = kind == "source"
        menu = QMenu(self)
        choose = menu.addAction("打开表格选择单元格…")
        manual = menu.addAction("手动输入单元格…")
        same = None
        if not source:
            menu.addSeparator()
            same = menu.addAction("与来源同位置")
        chosen = menu.exec(
            self.table.viewport().mapToGlobal(
                self.table.visualItemRect(item).bottomLeft()
            )
        )
        if chosen is None:
            return
        if chosen == same:
            if not rule.source_cells:
                QMessageBox.information(
                    self, "尚未设置来源", "请先设置本组的来源单元格。"
                )
                return
            rule.target_cells = list(rule.source_cells)
            rule.mode = infer_mapping_mode(
                rule.source_cells, rule.target_cells
            )
            self.refresh_table()
            return
        if chosen == manual:
            current_cells = (
                rule.source_cells if source else rule.target_cells
            )
            value, accepted = QInputDialog.getText(
                self,
                "手动输入单元格",
                "单元格地址（逗号分隔）：",
                text=format_cell_addresses(current_cells),
            )
            if not accepted:
                return
            try:
                cells = parse_cell_addresses(value)
            except Exception as exc:
                QMessageBox.warning(self, "单元格格式错误", str(exc))
                return
            old_same = rule.target_cells == rule.source_cells
            if source:
                rule.source_cells = cells
                if old_same:
                    rule.target_cells = list(cells)
            else:
                rule.target_cells = cells
            if rule.source_cells and rule.target_cells:
                rule.mode = infer_mapping_mode(
                    rule.source_cells, rule.target_cells
                )
            self.refresh_table()
            return
        if chosen != choose:
            return
        path_text = rule.source_file if source else rule.target_file
        sheet = rule.source_sheet if source else rule.target_sheet
        if not path_text or not sheet:
            QMessageBox.information(
                self,
                "请先选择位置",
                "请先选择本行的工作簿和工作表。",
            )
            return
        dialog = QtCellPickerDialog(
            self,
            Path(path_text),
            sheet,
            list(rule.source_cells if source else rule.target_cells),
            None if source else list(rule.source_cells),
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        old_same = rule.target_cells == rule.source_cells
        if source:
            rule.source_cells = list(dialog.result or [])
            if old_same:
                rule.target_cells = list(rule.source_cells)
        else:
            rule.target_cells = list(dialog.result or [])
        if rule.source_cells and rule.target_cells:
            rule.mode = infer_mapping_mode(
                rule.source_cells, rule.target_cells
            )
        self.refresh_table()

    def show_context_menu(self, position) -> None:
        menu = QMenu(self)
        menu.addAction("复制", self.copy_selection)
        menu.addAction("粘贴", self.paste_selection)
        menu.addSeparator()
        menu.addAction("在上方插入映射组", self.insert_above)
        menu.addAction("在下方插入映射组", self.insert_below)
        menu.addAction("删除选中映射组", self.delete_selected_groups)
        menu.exec(self.table.viewport().mapToGlobal(position))

    def delete_selected_groups(self) -> None:
        indices = self.selected_group_indices()
        if not indices:
            self.refresh_table()
            self.status_label.setText(
                "默认新增行已清空，并已自动补充新的空白组。"
            )
            return
        groups = "、".join(str(index + 1) for index in indices)
        if QMessageBox.question(
            self,
            "删除选中映射组",
            f"确定删除第 {groups} 组吗？",
        ) != QMessageBox.StandardButton.Yes:
            return
        for index in reversed(indices):
            self.rules.pop(index)
        self.refresh_table()
        self.status_label.setText(f"已删除 {len(indices)} 个映射组。")

    def copy_groups(self) -> None:
        indices = self.selected_group_indices()
        self.copied_groups = [
            MappingRule(**asdict(self.rules[index]))
            for index in indices
        ]
        self.copied_rows = []
        self.clipboard_scope = "group"
        self.status_label.setText(
            f"已复制 {len(self.copied_groups)} 个映射组。"
        )

    def paste_groups(self) -> None:
        if not self.copied_groups:
            self.status_label.setText("尚未复制映射组。")
            return
        indices = self.selected_group_indices()
        insertion = (max(indices) + 1) if indices else len(self.rules)
        copies = [
            MappingRule(**asdict(rule)) for rule in self.copied_groups
        ]
        self.rules[insertion:insertion] = copies
        self.refresh_table()
        self.status_label.setText(f"已粘贴 {len(copies)} 个映射组。")

    def copy_rows(self) -> None:
        rows = sorted({item.row() for item in self.table.selectedItems()})
        flattened = self._flatten_rows()
        self.copied_rows = [
            (
                flattened[row][0],
                flattened[row][1],
                list(flattened[row][2]),
            )
            for row in rows
            if row < len(flattened)
        ]
        self.copied_groups = []
        self.clipboard_scope = "row"
        lines = [
            "\t".join(
                (
                    format_project_path(Path(book), self.base_folder)
                    if book else "",
                    sheet,
                    format_cell_addresses(cells) if cells else "",
                )
            )
            for book, sheet, cells in self.copied_rows
        ]
        QApplication.clipboard().setText("\n".join(lines))
        self.status_label.setText(f"已复制 {len(lines)} 行内容。")

    def paste_rows(self) -> None:
        start = self.table.currentRow()
        if start < 0:
            return
        rows = self._flatten_rows()
        incoming = [
            (book, sheet, list(cells))
            for book, sheet, cells in self.copied_rows
        ]
        if not incoming:
            QMessageBox.information(self, "无可粘贴内容", "请先复制单行。")
            return
        while len(rows) < start + len(incoming):
            rows.extend([("", "", []), ("", "", [])])
        rows[start:start + len(incoming)] = incoming
        self._replace_from_rows(rows)
        self.refresh_table()

    def selection_is_complete_groups(self) -> bool:
        selected_rows = {
            item.row() for item in self.table.selectedItems()
            if not bool(item.data(ROLE_DRAFT))
        }
        if not selected_rows:
            return False
        return all(
            index * 2 in selected_rows and index * 2 + 1 in selected_rows
            for index in {row // 2 for row in selected_rows}
        )

    def copy_selection(self) -> None:
        if self.selection_is_complete_groups():
            self.copy_groups()
        else:
            self.copy_rows()

    def paste_selection(self) -> None:
        if self.clipboard_scope == "group":
            self.paste_groups()
        elif self.clipboard_scope == "row":
            self.paste_rows()
        else:
            QMessageBox.information(self, "无可粘贴内容", "请先使用 Ctrl+C 复制内容。")

    def _flatten_rows(self):
        rows = []
        for rule in self.rules:
            rows.extend(
                (
                    (
                        rule.source_file,
                        rule.source_sheet,
                        list(rule.source_cells),
                    ),
                    (
                        rule.target_file,
                        rule.target_sheet,
                        list(rule.target_cells),
                    ),
                )
            )
        return rows

    def _replace_from_rows(self, rows) -> None:
        if len(rows) % 2:
            rows.append(("", "", []))
        rebuilt = []
        for index in range(0, len(rows), 2):
            source = rows[index]
            target = rows[index + 1]
            mode = (
                infer_mapping_mode(source[2], target[2])
                if source[2] and target[2]
                else MODE_MANUAL
            )
            rebuilt.append(
                MappingRule(
                    source[0], source[1], source[2],
                    target[0], target[1], target[2], mode,
                )
            )
        self.rules = rebuilt

    def insert_above(self) -> None:
        indices = self.selected_group_indices()
        index = min(indices) if indices else len(self.rules)
        self.rules.insert(
            index, MappingRule("", "", [], "", "", [], MODE_MANUAL)
        )
        self.refresh_table()

    def insert_below(self) -> None:
        indices = self.selected_group_indices()
        index = max(indices) + 1 if indices else len(self.rules)
        self.rules.insert(
            index, MappingRule("", "", [], "", "", [], MODE_MANUAL)
        )
        self.refresh_table()

    def choose_base_folder(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self, "选择默认文件夹", str(self.base_folder or Path.cwd())
        )
        if selected:
            self.base_folder = Path(selected)
            self.base_edit.setText(selected)
            if self.output_folder is None:
                self.output_folder = self.base_folder / "映射结果"
                self.output_edit.setText(str(self.output_folder))
            self.refresh_table()

    def choose_output_folder(self) -> None:
        selected = QFileDialog.getExistingDirectory(
            self,
            "选择输出文件夹",
            str(self.output_folder or self.base_folder or Path.cwd()),
        )
        if selected:
            self.output_folder = Path(selected)
            self.output_edit.setText(selected)

    def import_scheme(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(
            self,
            "导入 Excel 映射表",
            str(self.base_folder or Path.cwd()),
            "Excel 映射表 (*.xlsx *.xlsm)",
        )
        if not selected:
            return
        try:
            _sources, _targets, imported = load_excel_mapping_scheme(
                Path(selected)
            )
        except Exception as exc:
            QMessageBox.critical(self, "导入失败", str(exc))
            return
        box = QMessageBox(self)
        box.setWindowTitle("选择导入方式")
        box.setText(f"映射表包含 {len(imported)} 组映射，请选择导入方式。")
        replace = box.addButton(
            "清空并覆盖", QMessageBox.ButtonRole.AcceptRole
        )
        append = box.addButton(
            "新增到现有映射", QMessageBox.ButtonRole.ActionRole
        )
        box.addButton("取消", QMessageBox.ButtonRole.RejectRole)
        box.exec()
        if box.clickedButton() == replace:
            self.rules = imported
        elif box.clickedButton() == append:
            self.rules.extend(imported)
        else:
            return
        self.base_folder = Path(selected).parent
        self.base_edit.setText(str(self.base_folder))
        self.refresh_table()

    def export_scheme(self) -> None:
        rules = self.active_rules()
        if not rules:
            QMessageBox.information(self, "方案为空", "没有可导出的映射。")
            return
        selected, _ = QFileDialog.getSaveFileName(
            self,
            "导出映射表",
            str((self.base_folder or Path.cwd()) / "映射方案.xlsx"),
            "Excel 映射表 (*.xlsx)",
        )
        if selected:
            path = Path(selected)
            if path.suffix.lower() != ".xlsx":
                path = path.with_suffix(".xlsx")
            save_excel_mapping_scheme(path, rules, path.parent)
            self.status_label.setText(f"映射表已导出：{path.name}")

    def clear_all(self) -> None:
        if QMessageBox.question(
            self, "清空所有映射", "确定清空全部映射关系吗？"
        ) == QMessageBox.StandardButton.Yes:
            self.rules.clear()
            self.refresh_table()

    def current_plan(self):
        rules = self.active_rules()
        if not rules:
            raise ValueError("请至少填写一组完整映射。")
        sources = list(
            dict.fromkeys(Path(rule.source_file) for rule in rules)
        )
        targets = list(
            dict.fromkeys(Path(rule.target_file) for rule in rules)
        )
        return sources, targets, rules

    def precheck(self) -> bool:
        try:
            sources, targets, rules = self.current_plan()
            expanded, warnings = validate_mapping_plan(
                sources, targets, rules
            )
            text = (
                f"来源工作簿：{len(sources)} 个\n"
                f"目标工作簿：{len(targets)} 个\n"
                f"映射组：{len(rules)} 组\n"
                f"展开写入：{len(expanded)} 项"
            )
            if warnings:
                text += f"\n\n{len(warnings)} 个目标单元格已有内容。"
            QMessageBox.information(self, "预检查通过", text)
            return True
        except Exception as exc:
            QMessageBox.critical(self, "预检查失败", str(exc))
            return False

    def preview(self) -> None:
        try:
            expanded = expand_rules(self.current_plan()[2])
        except Exception as exc:
            QMessageBox.warning(self, "无法展开映射", str(exc))
            return
        dialog = QDialog(self)
        dialog.setWindowTitle(f"映射预览（{len(expanded)} 项）")
        dialog.resize(900, 560)
        layout = QVBoxLayout(dialog)
        table = QTableWidget(len(expanded), 3)
        table.setHorizontalHeaderLabels(["序号", "来源", "目标"])
        table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        for index, mapping in enumerate(expanded):
            table.setItem(index, 0, QTableWidgetItem(str(index + 1)))
            table.setItem(
                index,
                1,
                QTableWidgetItem(
                    f"{Path(mapping.source_file).name} / "
                    f"{mapping.source_sheet} / {mapping.source_cell}"
                ),
            )
            table.setItem(
                index,
                2,
                QTableWidgetItem(
                    f"{Path(mapping.target_file).name} / "
                    f"{mapping.target_sheet} / {mapping.target_cell}"
                ),
            )
        layout.addWidget(table)
        close = QPushButton("关闭")
        close.clicked.connect(dialog.accept)
        layout.addWidget(close)
        dialog.exec()

    def run_mapping(self) -> None:
        if self.output_folder is None:
            QMessageBox.warning(self, "缺少输出位置", "请选择输出文件夹。")
            return
        try:
            sources, targets, rules = self.current_plan()
            expanded, warnings = validate_mapping_plan(
                sources, targets, rules
            )
            if warnings and QMessageBox.question(
                self,
                "确认覆盖副本中的内容",
                f"{len(warnings)} 个目标单元格已有内容，是否继续？",
            ) != QMessageBox.StandardButton.Yes:
                return
            if spreadsheet_automation_provider() is None:
                source_formulas, target_formulas = xls_formula_risk(
                    expanded, targets
                )
                if (source_formulas or target_formulas) and (
                    QMessageBox.question(
                        self,
                        "纯 Python 模式：公式将写成数值",
                        "未检测到 Excel/WPS。\n"
                        f"来源公式：{source_formulas}；"
                        f"目标公式：{target_formulas}。\n"
                        "公式会按当前计算结果写成数值，是否继续？",
                    )
                    != QMessageBox.StandardButton.Yes
                ):
                    return

            def progress(current: int, total: int, message: str) -> None:
                self.progress.setValue(
                    int(current / max(total, 1) * 100)
                )
                self.status_label.setText(message)
                QApplication.processEvents()

            outputs = execute_mapping_plan(
                sources, targets, rules, self.output_folder, progress
            )
            self.progress.setValue(100)
            QMessageBox.information(
                self,
                "映射完成",
                "已生成以下副本：\n" + "\n".join(map(str, outputs)),
            )
            self.status_label.setText("映射完成；原文件未修改。")
        except Exception as exc:
            QMessageBox.critical(self, "映射失败", str(exc))
            self.status_label.setText("映射失败。")
        finally:
            self.progress.setValue(0)

    def show_help(self) -> None:
        QMessageBox.information(
            self,
            "操作说明",
            "单击“映射组”列：选中整组\n"
            "单击“类型”列：选中单行\n"
            "单击工作簿/工作表/单元格：从菜单选择或手动输入\n\n"
            "Delete：删除选中映射组\n"
            "Ctrl+C / Ctrl+V：根据选中的单行或整组自动复制粘贴\n\n"
            "导入映射表支持：\n"
            "三列：工作簿、工作表、单元格（来源/目标交替）\n"
            "四列：类型、工作簿、工作表、单元格\n"
            "五列：映射组、类型、工作簿、工作表、单元格",
        )


def main() -> None:
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = ExcelMappingQtWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
