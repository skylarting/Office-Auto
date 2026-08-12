"""Unified smart-and-manual Excel worksheet mapping studio."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QDialog, QFileDialog, QFrame,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QListWidget, QMainWindow,
    QMenu, QMessageBox, QProgressBar, QPushButton, QInputDialog, QStackedWidget,
    QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

# The Qt build intentionally excludes Tk/Tcl. Import the Qt compatibility
# bootstrap first so excel_mapper sees the lightweight tkinter placeholders.
from excel_mapping_qt import bundled_asset, install_tkinter_build_stub

from excel_mapper import (
    MODE_SEQUENCE, MappingRule, best_name_match, execute_mapping_plan,
    format_cell_addresses, infer_mapping_mode, spreadsheet_automation_provider,
    validate_mapping_plan, workbook_files_in_folder, workbook_sheet_names,
    xls_formula_risk,
)
from excel_sheet_viewer import DualSheetViewer
from smart_template import SheetPair, SmartMatch, match_pair, suggest_sheet_pairs


WORKBOOK_FILTER = "Excel 文件 (*.xls *.xlsx *.xlsm)"


@dataclass
class WorksheetMapping:
    source_file: str
    source_sheet: str
    target_file: str
    target_sheet: str
    source_cells: list[str] = field(default_factory=list)
    target_cells: list[str] = field(default_factory=list)
    smart_source_cells: list[str] = field(default_factory=list)
    smart_target_cells: list[str] = field(default_factory=list)
    score: float = 0.0
    enabled: bool = True

    @property
    def status(self) -> str:
        if not self.target_file or not self.target_sheet:
            return "尚未选择目标工作表"
        if not self.source_cells:
            return "尚未选择单元格"
        if len(self.source_cells) != len(self.target_cells):
            return f"数量不一致（{len(self.source_cells)} / {len(self.target_cells)}）"
        return f"已设置 {len(self.source_cells)} 项"


def suggest_studio_mappings(
    source_files: list[Path], target_files: list[Path]
) -> list[WorksheetMapping]:
    """Choose the best target workbook/sheet for each source worksheet."""
    best: dict[tuple[str, str], SheetPair] = {}
    for target in target_files:
        for pair in suggest_sheet_pairs(source_files, target):
            key = (pair.source_file, pair.source_sheet)
            if key not in best or pair.score > best[key].score:
                best[key] = pair
    result: list[WorksheetMapping] = []
    for pair in best.values():
        matches = [item for item in match_pair(pair) if item.target_address]
        sources = [item.source_address for item in matches]
        targets = [item.target_address for item in matches]
        result.append(
            WorksheetMapping(
                pair.source_file, pair.source_sheet,
                pair.target_file, pair.target_sheet,
                list(sources), list(targets), list(sources), list(targets),
                pair.score,
            )
        )
    return result


def mapping_rules_from_rows(rows: list[WorksheetMapping]) -> list[MappingRule]:
    rules: list[MappingRule] = []
    for index, row in enumerate(rows, 1):
        if not row.enabled:
            continue
        if not row.source_cells or not row.target_cells:
            raise ValueError(f"第 {index} 行尚未选择需要映射的单元格。")
        if len(row.source_cells) != len(row.target_cells):
            raise ValueError(
                f"第 {index} 行来源选择了 {len(row.source_cells)} 个单元格，"
                f"目标选择了 {len(row.target_cells)} 个单元格。两边数量必须相同。"
            )
        rules.append(
            MappingRule(
                row.source_file, row.source_sheet, list(row.source_cells),
                row.target_file, row.target_sheet, list(row.target_cells),
                infer_mapping_mode(row.source_cells, row.target_cells),
            )
        )
    if not rules:
        raise ValueError("请至少设置一组工作表对应关系。")
    return rules


def preserve_existing_pair_order(
    mapping: WorksheetMapping,
    selected_sources: list[str],
    selected_targets: list[str],
) -> tuple[list[str], list[str]]:
    """Keep smart/manual pair order, then append newly selected cells in view order."""
    source_set, target_set = set(selected_sources), set(selected_targets)
    retained = [
        (source, target)
        for source, target in zip(mapping.source_cells, mapping.target_cells)
        if source in source_set and target in target_set
    ]
    used_sources = {source for source, _ in retained}
    used_targets = {target for _, target in retained}
    new_sources = [cell for cell in selected_sources if cell not in used_sources]
    new_targets = [cell for cell in selected_targets if cell not in used_targets]
    if len(new_sources) != len(new_targets):
        raise ValueError(
            "保留原有对应关系后，新增的来源和目标单元格数量不一致。"
        )
    pairs = retained + list(zip(new_sources, new_targets))
    return [source for source, _ in pairs], [target for _, target in pairs]


class FileListPanel(QFrame):
    def __init__(self, title: str, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("filePanel")
        box = QVBoxLayout(self)
        heading = QLabel(title); heading.setObjectName("sectionTitle")
        box.addWidget(heading)
        self.list = QListWidget()
        self.list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        box.addWidget(self.list, 1)
        buttons = QHBoxLayout()
        add_files = QPushButton("添加文件")
        add_files.clicked.connect(self.choose_files)
        buttons.addWidget(add_files)
        add_folder = QPushButton("添加文件夹")
        add_folder.clicked.connect(self.choose_folder)
        buttons.addWidget(add_folder)
        remove = QPushButton("移除选中")
        remove.clicked.connect(self.remove_selected)
        buttons.addWidget(remove)
        buttons.addStretch(1)
        box.addLayout(buttons)

    def paths(self) -> list[Path]:
        return [Path(self.list.item(index).data(Qt.ItemDataRole.UserRole)) for index in range(self.list.count())]

    def selected_paths(self) -> list[Path]:
        return [Path(item.data(Qt.ItemDataRole.UserRole)) for item in self.list.selectedItems()]

    def add_paths(self, paths: list[Path]) -> None:
        existing = {str(path.resolve()) for path in self.paths()}
        for path in paths:
            resolved = str(path.resolve())
            if resolved in existing:
                continue
            self.list.addItem(path.name)
            added = self.list.item(self.list.count() - 1)
            added.setData(Qt.ItemDataRole.UserRole, resolved)
            added.setToolTip(resolved)
            existing.add(resolved)

    def choose_files(self) -> None:
        values, _ = QFileDialog.getOpenFileNames(self, "选择 Excel 文件", "", WORKBOOK_FILTER)
        self.add_paths([Path(value) for value in values])

    def choose_folder(self) -> None:
        value = QFileDialog.getExistingDirectory(self, "选择包含 Excel 文件的文件夹")
        if value:
            self.add_paths(workbook_files_in_folder(Path(value)))

    def remove_selected(self) -> None:
        for item in reversed(self.list.selectedItems()):
            self.list.takeItem(self.list.row(item))

    def take_selected(self) -> list[Path]:
        paths = self.selected_paths()
        self.remove_selected()
        return paths


class SheetMappingDialog(QDialog):
    def __init__(self, parent, mapping: WorksheetMapping) -> None:
        super().__init__(parent)
        self.mapping = mapping
        self.saved = False
        self.setWindowTitle("设置工作表的数据对应关系")
        self.resize(1450, 900)
        self.setMinimumSize(1000, 700)
        root = QVBoxLayout(self)
        root.setContentsMargins(16, 14, 16, 14)
        title = QLabel("检查并修改需要映射的单元格")
        title.setObjectName("dialogTitle")
        root.addWidget(title)
        note = QLabel(
            "红框是程序提供的初始建议。您可以直接单击、拖动或使用条件按钮修改；"
            "保存时只采用当前实际选中的单元格。"
        )
        note.setObjectName("secondaryText")
        root.addWidget(note)
        self.viewer = DualSheetViewer()
        self.viewer.load_pair(
            Path(mapping.source_file), mapping.source_sheet,
            Path(mapping.target_file), mapping.target_sheet,
        )
        self.viewer.source.select_addresses(mapping.source_cells)
        self.viewer.target.select_addresses(mapping.target_cells)
        root.addWidget(self.viewer, 1)

        summary = QFrame(); summary.setObjectName("summaryPanel")
        summary_box = QVBoxLayout(summary)
        self.source_summary = QLabel(); self.source_summary.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.target_summary = QLabel(); self.target_summary.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        summary_box.addWidget(self.source_summary)
        summary_box.addWidget(self.target_summary)
        root.addWidget(summary)
        self.viewer.source.selection_changed.connect(lambda _items: self.refresh_summary())
        self.viewer.target.selection_changed.connect(lambda _items: self.refresh_summary())
        self.refresh_summary()

        actions = QHBoxLayout()
        restore = QPushButton("恢复智能建议")
        restore.clicked.connect(self.restore_suggestion)
        actions.addWidget(restore)
        actions.addStretch(1)
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        actions.addWidget(cancel)
        save = QPushButton("保存这组对应关系")
        save.setObjectName("primaryButton")
        save.clicked.connect(self.save_mapping)
        actions.addWidget(save)
        root.addLayout(actions)

    def refresh_summary(self) -> None:
        source = self.viewer.source.selected_addresses()
        target = self.viewer.target.selected_addresses()
        self.source_summary.setText(
            f"来源（{len(source)}）：{format_cell_addresses(source) or '尚未选择'}"
        )
        self.target_summary.setText(
            f"目标（{len(target)}）：{format_cell_addresses(target) or '尚未选择'}"
        )

    def restore_suggestion(self) -> None:
        self.viewer.source.select_addresses(self.mapping.smart_source_cells)
        self.viewer.target.select_addresses(self.mapping.smart_target_cells)
        self.refresh_summary()

    def save_mapping(self) -> None:
        source = self.viewer.source.selected_addresses()
        target = self.viewer.target.selected_addresses()
        if len(source) != len(target):
            QMessageBox.warning(
                self, "两边数量不一致",
                f"来源选择了 {len(source)} 个单元格，目标选择了 {len(target)} 个单元格。\n"
                "请继续选择或取消部分单元格，使两边数量相同。",
            )
            return
        if not source:
            QMessageBox.warning(self, "尚未选择", "请至少选择一组来源和目标单元格。")
            return
        try:
            source, target = preserve_existing_pair_order(
                self.mapping, source, target
            )
        except ValueError as exc:
            QMessageBox.warning(self, "无法保存对应关系", str(exc))
            return
        self.mapping.source_cells = source
        self.mapping.target_cells = target
        self.saved = True
        self.accept()

    def closeEvent(self, event) -> None:
        self.viewer.close()
        super().closeEvent(event)


class StudioWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Excel 映射工作台")
        self.resize(1280, 820)
        self.setMinimumSize(940, 650)
        self.mappings: list[WorksheetMapping] = []
        self.output_folder: Path | None = None
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)
        self.files_page = self.build_files_page()
        self.mapping_page = self.build_mapping_page()
        self.stack.addWidget(self.files_page)
        self.stack.addWidget(self.mapping_page)

    def page_header(self, title: str, subtitle: str, layout: QVBoxLayout) -> None:
        heading = QLabel(title); heading.setObjectName("pageTitle")
        layout.addWidget(heading)
        note = QLabel(subtitle); note.setObjectName("pageSubtitle"); note.setWordWrap(True)
        layout.addWidget(note)

    def build_files_page(self) -> QWidget:
        page = QWidget(); root = QVBoxLayout(page); root.setContentsMargins(24, 20, 24, 22)
        self.page_header("1. 选择来源文件和目标文件", "可以一次添加多个文件或整个文件夹。选错身份时，选中文件后点击中间的转换按钮。", root)
        lists = QHBoxLayout()
        self.sources = FileListPanel("来源文件")
        self.targets = FileListPanel("目标文件")
        lists.addWidget(self.sources, 1)
        middle = QVBoxLayout(); middle.addStretch(1)
        transfer = QPushButton("⇄  转换为另一侧")
        transfer.clicked.connect(self.transfer_files)
        middle.addWidget(transfer)
        middle.addStretch(1)
        lists.addLayout(middle)
        lists.addWidget(self.targets, 1)
        root.addLayout(lists, 1)
        actions = QHBoxLayout(); actions.addStretch(1)
        generate = QPushButton("智能生成工作表对应列表  →")
        generate.setObjectName("primaryButton")
        generate.clicked.connect(self.generate_mappings)
        actions.addWidget(generate)
        root.addLayout(actions)
        return page

    def build_mapping_page(self) -> QWidget:
        page = QWidget(); root = QVBoxLayout(page); root.setContentsMargins(24, 20, 24, 22)
        top = QHBoxLayout()
        back = QPushButton("← 返回文件选择")
        back.clicked.connect(lambda: self.stack.setCurrentWidget(self.files_page))
        top.addWidget(back); top.addStretch(1)
        root.addLayout(top)
        self.page_header(
            "2. 检查工作表对应关系",
            "每行是一组来源工作表和目标工作表。双击任意一行打开双表查看器；"
            "右键可以修改工作簿、工作表或启用状态。",
            root,
        )
        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(
            ["使用", "来源工作簿", "来源工作表", "→", "目标工作簿", "目标工作表", "数据对应"]
        )
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.setColumnWidth(0, 54); self.table.setColumnWidth(1, 260)
        self.table.setColumnWidth(2, 150); self.table.setColumnWidth(3, 36)
        self.table.setColumnWidth(4, 260); self.table.setColumnWidth(5, 180)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.cellDoubleClicked.connect(lambda row, _column: self.open_mapping(row))
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.mapping_menu)
        root.addWidget(self.table, 1)
        tools = QHBoxLayout()
        add = QPushButton("添加工作表对应")
        add.clicked.connect(self.add_mapping)
        tools.addWidget(add)
        remove = QPushButton("删除选中")
        remove.clicked.connect(self.remove_mappings)
        tools.addWidget(remove)
        refresh = QPushButton("重新智能识别")
        refresh.clicked.connect(self.generate_mappings)
        tools.addWidget(refresh)
        tools.addStretch(1)
        root.addLayout(tools)
        execute_box = QFrame(); execute_box.setObjectName("summaryPanel")
        execute = QHBoxLayout(execute_box)
        self.output_edit = QLineEdit(); self.output_edit.setReadOnly(True)
        self.output_edit.setPlaceholderText("点击选择输出文件夹；不选择时使用目标文件旁的“映射结果”")
        self.output_edit.mousePressEvent = lambda _event: self.choose_output()
        execute.addWidget(QLabel("输出文件夹：")); execute.addWidget(self.output_edit, 1)
        self.progress = QProgressBar(); self.progress.setMaximumWidth(190); self.progress.setValue(0)
        execute.addWidget(self.progress)
        run = QPushButton("开始生成报表副本")
        run.setObjectName("primaryButton"); run.clicked.connect(self.run_mapping)
        execute.addWidget(run)
        root.addWidget(execute_box)
        return page

    def transfer_files(self) -> None:
        source_selected = self.sources.selected_paths()
        target_selected = self.targets.selected_paths()
        if source_selected and target_selected:
            QMessageBox.information(self, "请选择一侧", "请只在来源或目标列表中选择需要转换的文件。")
            return
        if source_selected:
            self.targets.add_paths(self.sources.take_selected())
        elif target_selected:
            self.sources.add_paths(self.targets.take_selected())
        else:
            QMessageBox.information(self, "尚未选择文件", "请先选择需要转换身份的文件。")

    def generate_mappings(self) -> None:
        sources, targets = self.sources.paths(), self.targets.paths()
        if not sources or not targets:
            QMessageBox.warning(self, "文件不完整", "请至少添加一个来源文件和一个目标文件。")
            return
        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            self.mappings = suggest_studio_mappings(sources, targets)
            self.refresh_table()
            self.stack.setCurrentWidget(self.mapping_page)
        except Exception as exc:
            QMessageBox.critical(self, "智能识别失败", str(exc))
        finally:
            QApplication.restoreOverrideCursor()

    def refresh_table(self) -> None:
        self.table.setRowCount(len(self.mappings))
        for row, mapping in enumerate(self.mappings):
            use = QTableWidgetItem("使用" if mapping.enabled else "忽略")
            use.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            values = (
                use, QTableWidgetItem(Path(mapping.source_file).name),
                QTableWidgetItem(mapping.source_sheet), QTableWidgetItem("→"),
                QTableWidgetItem(Path(mapping.target_file).name if mapping.target_file else "尚未选择"),
                QTableWidgetItem(mapping.target_sheet or "尚未选择"),
                QTableWidgetItem(mapping.status),
            )
            for column, item in enumerate(values):
                item.setToolTip(
                    mapping.source_file if column == 1 else
                    mapping.target_file if column == 4 else item.text()
                )
                self.table.setItem(row, column, item)

    def open_mapping(self, row: int) -> None:
        mapping = self.mappings[row]
        if not mapping.target_file or not mapping.target_sheet:
            QMessageBox.warning(self, "目标工作表不完整", "请先为这一行选择目标工作簿和目标工作表。")
            return
        dialog = SheetMappingDialog(self, mapping)
        if dialog.exec() and dialog.saved:
            self.refresh_table()

    def edit_mapping_field(self, row: int, column: int) -> None:
        mapping = self.mappings[row]
        if column in (1, 4):
            paths = self.sources.paths() if column == 1 else self.targets.paths()
            current = mapping.source_file if column == 1 else mapping.target_file
            labels = [str(path) for path in paths]
            value, ok = QInputDialog.getItem(
                self, "选择工作簿", "工作簿：", labels,
                max(0, labels.index(current) if current in labels else 0), False,
            )
            if not ok or not value:
                return
            if column == 1:
                mapping.source_file = value
                sheets = workbook_sheet_names(Path(value))
                if mapping.source_sheet not in sheets:
                    mapping.source_sheet = sheets[0]
            else:
                mapping.target_file = value
                sheets = workbook_sheet_names(Path(value))
                if mapping.target_sheet not in sheets:
                    mapping.target_sheet = sheets[0]
        else:
            path = Path(mapping.source_file if column == 2 else mapping.target_file)
            sheets = workbook_sheet_names(path)
            current = mapping.source_sheet if column == 2 else mapping.target_sheet
            value, ok = QInputDialog.getItem(
                self, "选择工作表", "工作表：", sheets,
                max(0, sheets.index(current) if current in sheets else 0), False,
            )
            if not ok or not value:
                return
            if column == 2:
                mapping.source_sheet = value
            else:
                mapping.target_sheet = value
        self.refresh_suggestion(mapping)
        self.refresh_table()

    @staticmethod
    def refresh_suggestion(mapping: WorksheetMapping) -> None:
        pair = SheetPair(
            mapping.source_file, mapping.source_sheet,
            mapping.target_file, mapping.target_sheet, mapping.score,
        )
        matches = [item for item in match_pair(pair) if item.target_address]
        mapping.source_cells = [item.source_address for item in matches]
        mapping.target_cells = [item.target_address for item in matches]
        mapping.smart_source_cells = list(mapping.source_cells)
        mapping.smart_target_cells = list(mapping.target_cells)

    def mapping_menu(self, position) -> None:
        row = self.table.rowAt(position.y())
        if row < 0:
            return
        menu = QMenu(self)
        edit = menu.addAction("打开双表完整查看器")
        menu.addSeparator()
        edit_source_book = menu.addAction("修改来源工作簿")
        edit_source_sheet = menu.addAction("修改来源工作表")
        edit_target_book = menu.addAction("修改目标工作簿")
        edit_target_sheet = menu.addAction("修改目标工作表")
        menu.addSeparator()
        toggle = menu.addAction("忽略此行" if self.mappings[row].enabled else "使用此行")
        delete = menu.addAction("删除此行")
        chosen = menu.exec(self.table.viewport().mapToGlobal(position))
        if chosen == edit:
            self.open_mapping(row)
        elif chosen == edit_source_book:
            self.edit_mapping_field(row, 1)
        elif chosen == edit_source_sheet:
            self.edit_mapping_field(row, 2)
        elif chosen == edit_target_book:
            self.edit_mapping_field(row, 4)
        elif chosen == edit_target_sheet:
            self.edit_mapping_field(row, 5)
        elif chosen == toggle:
            self.mappings[row].enabled = not self.mappings[row].enabled
            self.refresh_table()
        elif chosen == delete:
            self.mappings.pop(row); self.refresh_table()

    def add_mapping(self) -> None:
        sources, targets = self.sources.paths(), self.targets.paths()
        if not sources or not targets:
            return
        source = sources[0]; target = targets[0]
        source_sheet = workbook_sheet_names(source)[0]
        target_sheet = best_name_match(source_sheet, workbook_sheet_names(target)) or workbook_sheet_names(target)[0]
        self.mappings.append(WorksheetMapping(str(source), source_sheet, str(target), target_sheet))
        self.refresh_table()

    def remove_mappings(self) -> None:
        rows = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.mappings.pop(row)
        self.refresh_table()

    def choose_output(self) -> None:
        value = QFileDialog.getExistingDirectory(self, "选择输出文件夹", self.output_edit.text())
        if value:
            self.output_folder = Path(value); self.output_edit.setText(value)

    def run_mapping(self) -> None:
        try:
            rules = mapping_rules_from_rows(self.mappings)
            source_files = list(dict.fromkeys(Path(item.source_file) for item in self.mappings if item.enabled))
            target_files = list(dict.fromkeys(Path(item.target_file) for item in self.mappings if item.enabled))
            expanded, warnings = validate_mapping_plan(source_files, target_files, rules)
            if warnings and QMessageBox.question(
                self, "确认写入副本", f"目标副本中有 {len(warnings)} 个位置已有内容，是否继续？"
            ) != QMessageBox.StandardButton.Yes:
                return
            if spreadsheet_automation_provider() is None:
                source_formulas, target_formulas = xls_formula_risk(expanded, target_files)
                if (source_formulas or target_formulas) and QMessageBox.question(
                    self, "公式处理提示",
                    "未检测到 Excel/WPS。旧版 XLS 中的公式可能按当前计算结果写成数值，是否继续？"
                ) != QMessageBox.StandardButton.Yes:
                    return
            output = self.output_folder or target_files[0].parent / "映射结果"
            output.mkdir(parents=True, exist_ok=True)
            def progress(current, total, _message):
                self.progress.setValue(int(current / max(total, 1) * 100)); QApplication.processEvents()
            outputs = execute_mapping_plan(source_files, target_files, rules, output, progress)
            QMessageBox.information(self, "生成完成", "已生成报表副本：\n" + "\n".join(map(str, outputs)))
        except Exception as exc:
            QMessageBox.critical(self, "生成失败", str(exc))
        finally:
            self.progress.setValue(0)


def application_style() -> str:
    return """
    QMainWindow, QWidget { background: #F5F7FA; color: #1F2937; font-family: "Microsoft YaHei UI"; font-size: 14px; }
    QLabel#pageTitle { font-size: 26px; font-weight: 600; color: #172B4D; }
    QLabel#pageSubtitle, QLabel#secondaryText { color: #667085; }
    QLabel#sectionTitle, QLabel#dialogTitle, QLabel#sheetPaneTitle { font-size: 17px; font-weight: 600; color: #172B4D; }
    QFrame#filePanel, QFrame#summaryPanel, QFrame#sheetPane { background: white; border: 1px solid #DDE3EC; border-radius: 9px; }
    QLineEdit, QComboBox, QListWidget, QTableWidget, QTableView { background: white; border: 1px solid #CBD5E1; border-radius: 5px; selection-background-color: #DCEAFF; selection-color: #172B4D; }
    QPushButton { background: white; border: 1px solid #B8C2D1; border-radius: 6px; padding: 8px 14px; }
    QPushButton:hover { background: #F0F5FF; border-color: #7EA6E0; }
    QPushButton:checked { background: #EAF2FF; border-color: #2563EB; color: #175CD3; }
    QPushButton#primaryButton { background: #2563EB; color: white; border-color: #2563EB; font-weight: 600; }
    QPushButton#primaryButton:hover { background: #1D4ED8; }
    QProgressBar { background: #E5EAF1; border: none; border-radius: 4px; text-align: center; }
    QProgressBar::chunk { background: #2563EB; }
    """


def main() -> None:
    install_tkinter_build_stub()
    app = QApplication(sys.argv); app.setStyle("Fusion"); app.setStyleSheet(application_style())
    icon = bundled_asset("assets/excel-mapper.ico")
    if icon.exists(): app.setWindowIcon(QIcon(str(icon)))
    window = StudioWindow(); window.show(); sys.exit(app.exec())


if __name__ == "__main__":
    main()
