"""Qt window for extracting selected cells from one or many Excel workbooks."""

from __future__ import annotations

from collections import defaultdict
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QComboBox, QFileDialog, QFrame, QGridLayout,
    QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMainWindow,
    QMenu, QMessageBox, QProgressBar, QPushButton, QSplitter, QToolButton,
    QVBoxLayout, QWidget,
)

from excel_mapper import WorkbookReader, format_cell_addresses, split_address, workbook_files_in_folder, workbook_sheet_names
from excel_sheet_viewer import SheetViewPane
from export_summary import create_new_summary_workbook, create_summary_copies


WORKBOOK_FILTER = "Excel 文件 (*.xls *.xlsx *.xlsm)"


class ExcelSummaryWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Excel 单元格汇总工具")
        self.resize(1180, 800); self.setMinimumSize(900, 640)
        self.source_files: list[Path] = []
        self.preview_pair: tuple[Path, str] | None = None

        central = QWidget(); self.setCentralWidget(central)
        root = QVBoxLayout(central); root.setContentsMargins(14, 10, 14, 10); root.setSpacing(7)
        title = QLabel("Excel 单元格汇总工具"); title.setObjectName("pageTitle"); root.addWidget(title)

        source_panel = QFrame(); source_panel.setObjectName("summaryPanel")
        source_layout = QGridLayout(source_panel); source_layout.setContentsMargins(10, 8, 10, 8)
        source_layout.addWidget(QLabel("数据来源："), 0, 0)
        self.mode = QComboBox(); self.mode.addItems(("单个工作簿", "文件夹内全部工作簿"))
        self.mode.currentIndexChanged.connect(self.clear_source); source_layout.addWidget(self.mode, 0, 1)
        self.source_edit = QLineEdit(); self.source_edit.setReadOnly(True)
        self.source_edit.setPlaceholderText("点击此处选择工作簿")
        self.source_edit.mousePressEvent = lambda _event: self.choose_source()
        source_layout.addWidget(self.source_edit, 0, 2, 1, 3)
        source_layout.addWidget(QLabel("预览工作表："), 1, 0)
        self.preview_button = QToolButton(); self.preview_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.preview_button.setText("请先选择数据来源"); self.preview_menu = QMenu(self.preview_button)
        self.preview_button.setMenu(self.preview_menu); source_layout.addWidget(self.preview_button, 1, 1, 1, 4)
        source_layout.setColumnStretch(2, 1); root.addWidget(source_panel)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.viewer = SheetViewPane("选择需要提取的单元格")
        self.viewer.selection_changed.connect(lambda _items: self.update_selection_summary())
        splitter.addWidget(self.viewer)
        options = QFrame(); options.setObjectName("summaryPanel"); options.setMaximumWidth(360)
        layout = QVBoxLayout(options); layout.setContentsMargins(10, 8, 10, 8); layout.setSpacing(6)
        heading = QHBoxLayout(); section = QLabel("处理工作表"); section.setObjectName("sectionTitle"); heading.addWidget(section); heading.addStretch(1)
        select_all = QPushButton("全选"); select_all.setObjectName("compactButton"); select_all.clicked.connect(self.select_all_sheets); heading.addWidget(select_all)
        clear = QPushButton("清空"); clear.setObjectName("compactButton"); clear.clicked.connect(self.sheet_list_clear); heading.addWidget(clear); layout.addLayout(heading)
        self.sheet_list = QListWidget(); self.sheet_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        layout.addWidget(self.sheet_list, 1)
        layout.addWidget(QLabel("文件夹模式按“工作簿 / 工作表”显示；默认全部选中。"))
        section = QLabel("输出设置"); section.setObjectName("sectionTitle"); layout.addWidget(section)
        self.output_mode = QComboBox(); self.output_mode.addItems(("新建一个汇总工作簿", "另存工作簿副本，并插入汇总工作表"))
        self.output_mode.currentIndexChanged.connect(self.output_mode_changed); layout.addWidget(self.output_mode)
        row = QHBoxLayout(); row.addWidget(QLabel("汇总表名称")); self.summary_name = QLineEdit("汇总"); row.addWidget(self.summary_name, 1); layout.addLayout(row)
        row = QHBoxLayout(); row.addWidget(QLabel("列名")); self.header_mode = QComboBox(); self.header_mode.addItems(("使用单元格地址（A1）", "自动使用行名 + 列名")); row.addWidget(self.header_mode, 1); layout.addLayout(row)
        self.output_edit = QLineEdit(); self.output_edit.setReadOnly(True); self.output_edit.setPlaceholderText("点击选择输出位置")
        self.output_edit.mousePressEvent = lambda _event: self.choose_output(); layout.addWidget(self.output_edit)
        splitter.addWidget(options); splitter.setStretchFactor(0, 1); splitter.setStretchFactor(1, 0); root.addWidget(splitter, 1)

        self.selection_summary = QLabel("已选择 0 项：尚未选择"); self.selection_summary.setObjectName("secondaryText"); root.addWidget(self.selection_summary)
        bottom = QGridLayout(); bottom.setVerticalSpacing(3)
        self.status = QLabel("尚未开始"); bottom.addWidget(self.status, 0, 0)
        self.progress = QProgressBar(); self.progress.setValue(0); bottom.addWidget(self.progress, 1, 0)
        run = QPushButton("生成汇总文件"); run.setObjectName("primaryButton"); run.clicked.connect(self.run_summary); bottom.addWidget(run, 0, 1, 2, 1)
        root.addLayout(bottom)

    def clear_source(self) -> None:
        self.source_files = []; self.preview_pair = None; self.source_edit.clear(); self.preview_menu.clear(); self.preview_button.setText("请先选择数据来源")
        self.sheet_list.clear(); self.viewer.clear_sheet(); self.status.setText("尚未开始"); self.progress.setValue(0)
        self.source_edit.setPlaceholderText("点击此处选择工作簿" if self.mode.currentIndex() == 0 else "点击此处选择文件夹")

    def choose_source(self) -> None:
        if self.mode.currentIndex() == 0:
            value, _ = QFileDialog.getOpenFileName(self, "选择 Excel 工作簿", str(Path.home()), WORKBOOK_FILTER)
            if not value: return
            self.source_files = [Path(value)]; self.source_edit.setText(value)
        else:
            value = QFileDialog.getExistingDirectory(self, "选择包含 Excel 工作簿的文件夹", str(Path.home()))
            if not value: return
            self.source_files = workbook_files_in_folder(Path(value))
            if not self.source_files:
                QMessageBox.warning(self, "没有表格", "所选文件夹中没有可处理的 Excel 工作簿。"); return
            self.source_edit.setText(f"{value}（{len(self.source_files)} 个工作簿）")
        self.rebuild_sheet_choices()
        base = self.source_files[0].parent
        self.output_edit.setText(str(base / "单元格汇总.xlsx"))

    def rebuild_sheet_choices(self) -> None:
        self.preview_menu.clear(); self.sheet_list.clear(); first: tuple[Path, str] | None = None
        for path in self.source_files:
            submenu = self.preview_menu.addMenu(path.name)
            for sheet in workbook_sheet_names(path):
                if first is None: first = (path, sheet)
                action = submenu.addAction(sheet); action.triggered.connect(lambda _checked=False, p=path, s=sheet: self.select_preview(p, s))
                label = sheet if len(self.source_files) == 1 else f"{path.stem} / {sheet}"
                item = QListWidgetItem(label); item.setData(Qt.ItemDataRole.UserRole, (str(path.resolve()), sheet)); self.sheet_list.addItem(item); item.setSelected(True)
        if first: self.select_preview(*first)

    def select_preview(self, path: Path, sheet: str) -> None:
        self.preview_pair = (path, sheet); self.preview_button.setText(f"{path.name} / {sheet}")
        try: self.viewer.load_sheet(path, sheet)
        except Exception as exc: QMessageBox.warning(self, "无法预览工作表", str(exc))

    def select_all_sheets(self) -> None:
        self.sheet_list.selectAll()

    def sheet_list_clear(self) -> None:
        self.sheet_list.clearSelection()

    def output_mode_changed(self) -> None:
        self.output_edit.clear()
        self.output_edit.setPlaceholderText("点击选择汇总文件" if self.output_mode.currentIndex() == 0 else "点击选择副本输出文件夹")

    def choose_output(self) -> None:
        if self.output_mode.currentIndex() == 0:
            value, _ = QFileDialog.getSaveFileName(self, "保存汇总文件", self.output_edit.text() or str(Path.cwd() / "单元格汇总.xlsx"), "Excel 工作簿 (*.xlsx)")
            if value:
                path = Path(value); self.output_edit.setText(str(path if path.suffix.lower() == ".xlsx" else path.with_suffix(".xlsx")))
        else:
            value = QFileDialog.getExistingDirectory(self, "选择副本输出文件夹", self.output_edit.text() or str(Path.cwd()))
            if value: self.output_edit.setText(value)

    def update_selection_summary(self) -> None:
        addresses = self.viewer.selected_addresses(); self.selection_summary.setText(f"已选择 {len(addresses)} 项：{format_cell_addresses(addresses) or '尚未选择'}")

    def included_sheets(self) -> dict[str, set[str]]:
        result: dict[str, set[str]] = defaultdict(set)
        for item in self.sheet_list.selectedItems():
            path, sheet = item.data(Qt.ItemDataRole.UserRole); result[path].add(sheet)
        return dict(result)

    def automatic_headers(self, addresses: list[str]) -> list[str]:
        if not self.preview_pair: return addresses
        path, sheet = self.preview_pair; reader = WorkbookReader(path); headers: list[str] = []
        try:
            for address in addresses:
                row, column = split_address(address); row_label = column_label = ""
                for c in range(column - 1, -1, -1):
                    from excel_sheet_viewer import address_for
                    value = reader.read(sheet, address_for(row, c)).value
                    if isinstance(value, str) and value.strip(): row_label = value.strip(); break
                for r in range(row - 1, -1, -1):
                    value = reader.read(sheet, address_for(r, column)).value
                    if isinstance(value, str) and value.strip(): column_label = value.strip(); break
                headers.append(" / ".join(part for part in (row_label, column_label) if part) or address)
        finally: reader.close()
        return headers

    def run_summary(self) -> None:
        self.progress.setValue(0); self.status.setText("正在准备…")
        try:
            if not self.source_files: raise ValueError("请先选择数据来源。")
            addresses = self.viewer.selected_addresses()
            if not addresses: raise ValueError("请在工作表预览中选择需要提取的单元格。")
            included = self.included_sheets()
            if not included: raise ValueError("请至少选择一个需要处理的工作表。")
            output_text = self.output_edit.text().strip()
            if not output_text: raise ValueError("请选择输出位置。")
            summary_name = self.summary_name.text().strip() or "汇总"
            headers = addresses if self.header_mode.currentIndex() == 0 else self.automatic_headers(addresses)
            def progress(current, total, message):
                self.status.setText(message); self.progress.setValue(int(current / max(total, 1) * 100)); QApplication.processEvents()
            if self.output_mode.currentIndex() == 0:
                create_new_summary_workbook(self.source_files, addresses, Path(output_text), summary_name, progress, included_sheets_by_file=included, value_headers=headers)
                outputs = [Path(output_text)]
            else:
                output_folder = Path(output_text); output_folder.mkdir(parents=True, exist_ok=True)
                outputs = create_summary_copies(self.source_files, addresses, summary_name, output_folder, progress, included, headers)
            self.status.setText("汇总完成"); self.progress.setValue(100)
            QMessageBox.information(self, "汇总完成", "已生成：\n" + "\n".join(str(path) for path in outputs))
        except Exception as exc:
            self.status.setText("汇总失败"); self.progress.setValue(0); QMessageBox.critical(self, "汇总失败", str(exc))

    def closeEvent(self, event) -> None:
        self.viewer.close(); super().closeEvent(event)
