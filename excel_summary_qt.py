"""Qt window for extracting selected cells from one or many Excel workbooks."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView, QApplication, QCheckBox, QComboBox, QFileDialog, QFrame,
    QGridLayout, QHBoxLayout, QLabel, QLineEdit, QListWidget, QMainWindow,
    QMessageBox, QProgressBar, QPushButton, QSplitter, QVBoxLayout, QWidget,
)

from excel_mapper import format_cell_addresses, workbook_files_in_folder, workbook_sheet_names
from excel_sheet_viewer import SheetViewPane
from export_summary import create_new_summary_workbook


WORKBOOK_FILTER = "Excel 文件 (*.xls *.xlsx *.xlsm)"


class ExcelSummaryWindow(QMainWindow):
    """A compact visual workflow for creating a new consolidated workbook."""

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Excel 单元格汇总工具")
        self.resize(1180, 820); self.setMinimumSize(900, 650)
        self.source_files: list[Path] = []

        central = QWidget(); self.setCentralWidget(central)
        root = QVBoxLayout(central); root.setContentsMargins(18, 14, 18, 14); root.setSpacing(10)
        heading = QHBoxLayout()
        title = QLabel("Excel 单元格汇总工具"); title.setObjectName("pageTitle")
        heading.addWidget(title); heading.addStretch(1)
        self.mode = QComboBox(); self.mode.addItems(("单个工作簿", "文件夹内全部工作簿"))
        self.mode.currentIndexChanged.connect(self.clear_source)
        heading.addWidget(self.mode)
        choose = QPushButton("选择数据来源…"); choose.setObjectName("primaryButton")
        choose.clicked.connect(self.choose_source); heading.addWidget(choose)
        root.addLayout(heading)

        source_panel = QFrame(); source_panel.setObjectName("summaryPanel")
        source_layout = QGridLayout(source_panel)
        source_layout.addWidget(QLabel("数据来源："), 0, 0)
        self.source_edit = QLineEdit(); self.source_edit.setReadOnly(True)
        self.source_edit.setPlaceholderText("请选择一个工作簿或包含工作簿的文件夹")
        source_layout.addWidget(self.source_edit, 0, 1)
        source_layout.addWidget(QLabel("预览工作簿："), 1, 0)
        self.preview_book = QComboBox(); self.preview_book.currentIndexChanged.connect(self.book_changed)
        source_layout.addWidget(self.preview_book, 1, 1)
        source_layout.addWidget(QLabel("预览工作表："), 2, 0)
        self.preview_sheet = QComboBox(); self.preview_sheet.currentIndexChanged.connect(self.sheet_changed)
        source_layout.addWidget(self.preview_sheet, 2, 1)
        root.addWidget(source_panel)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        self.viewer = SheetViewPane("选择需要提取的单元格")
        self.viewer.selection_changed.connect(lambda _items: self.update_selection_summary())
        splitter.addWidget(self.viewer)
        options = QFrame(); options.setObjectName("summaryPanel"); options.setMaximumWidth(330)
        options_layout = QVBoxLayout(options)
        section = QLabel("处理工作表"); section.setObjectName("sectionTitle"); options_layout.addWidget(section)
        self.all_sheets = QCheckBox("处理每个工作簿中的全部工作表")
        self.all_sheets.setChecked(True); self.all_sheets.toggled.connect(self.sheet_scope_changed)
        options_layout.addWidget(self.all_sheets)
        self.sheet_list = QListWidget(); self.sheet_list.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.sheet_list.setEnabled(False); options_layout.addWidget(self.sheet_list, 1)
        options_layout.addWidget(QLabel("未勾选全部工作表时，可在上方多选需要处理的工作表名称。"))
        options_layout.addSpacing(10)
        section = QLabel("汇总文件"); section.setObjectName("sectionTitle"); options_layout.addWidget(section)
        options_layout.addWidget(QLabel("汇总工作表名称"))
        self.summary_name = QLineEdit("汇总"); options_layout.addWidget(self.summary_name)
        options_layout.addWidget(QLabel("输出文件"))
        self.output_edit = QLineEdit(); self.output_edit.setReadOnly(True)
        self.output_edit.setPlaceholderText("点击选择保存位置")
        self.output_edit.mousePressEvent = lambda _event: self.choose_output()
        options_layout.addWidget(self.output_edit)
        options_layout.addStretch(1)
        splitter.addWidget(options); splitter.setStretchFactor(0, 1); splitter.setStretchFactor(1, 0)
        root.addWidget(splitter, 1)

        self.selection_summary = QLabel("已选择 0 项：尚未选择")
        self.selection_summary.setObjectName("secondaryText"); root.addWidget(self.selection_summary)
        bottom = QGridLayout()
        self.status = QLabel("尚未开始"); bottom.addWidget(self.status, 0, 0)
        self.progress = QProgressBar(); self.progress.setValue(0); bottom.addWidget(self.progress, 1, 0)
        run = QPushButton("生成汇总文件"); run.setObjectName("primaryButton")
        run.clicked.connect(self.run_summary); bottom.addWidget(run, 0, 1, 2, 1)
        root.addLayout(bottom)

    def clear_source(self) -> None:
        self.source_files = []; self.source_edit.clear(); self.preview_book.clear(); self.preview_sheet.clear()

    def choose_source(self) -> None:
        if self.mode.currentIndex() == 0:
            value, _ = QFileDialog.getOpenFileName(self, "选择 Excel 工作簿", str(Path.home()), WORKBOOK_FILTER)
            if not value: return
            self.source_files = [Path(value)]
            self.source_edit.setText(value)
        else:
            value = QFileDialog.getExistingDirectory(self, "选择包含 Excel 工作簿的文件夹", str(Path.home()))
            if not value: return
            self.source_files = workbook_files_in_folder(Path(value))
            if not self.source_files:
                QMessageBox.warning(self, "没有表格", "所选文件夹中没有可处理的 Excel 工作簿。"); return
            self.source_edit.setText(f"{value}（{len(self.source_files)} 个工作簿）")
        self.preview_book.blockSignals(True); self.preview_book.clear()
        for path in self.source_files: self.preview_book.addItem(path.name, str(path))
        self.preview_book.blockSignals(False); self.preview_book.setCurrentIndex(0); self.book_changed()
        if not self.output_edit.text():
            base = self.source_files[0].parent
            self.output_edit.setText(str(base / "单元格汇总.xlsx"))

    def book_changed(self) -> None:
        path_text = self.preview_book.currentData()
        self.preview_sheet.blockSignals(True); self.preview_sheet.clear()
        if path_text:
            self.preview_sheet.addItems(workbook_sheet_names(Path(path_text)))
        self.preview_sheet.blockSignals(False)
        self.refresh_sheet_names(); self.sheet_changed()

    def refresh_sheet_names(self) -> None:
        names: list[str] = []
        for path in self.source_files:
            for name in workbook_sheet_names(path):
                if name not in names: names.append(name)
        selected = {item.text() for item in self.sheet_list.selectedItems()}
        self.sheet_list.clear(); self.sheet_list.addItems(names)
        for index in range(self.sheet_list.count()):
            if self.sheet_list.item(index).text() in selected:
                self.sheet_list.item(index).setSelected(True)

    def sheet_changed(self) -> None:
        path_text = self.preview_book.currentData(); sheet = self.preview_sheet.currentText()
        if path_text and sheet:
            try: self.viewer.load_sheet(Path(path_text), sheet)
            except Exception as exc: QMessageBox.warning(self, "无法预览工作表", str(exc))

    def sheet_scope_changed(self, all_enabled: bool) -> None:
        self.sheet_list.setEnabled(not all_enabled)

    def choose_output(self) -> None:
        value, _ = QFileDialog.getSaveFileName(
            self, "保存汇总文件", self.output_edit.text() or str(Path.cwd() / "单元格汇总.xlsx"),
            "Excel 工作簿 (*.xlsx)",
        )
        if value:
            path = Path(value); self.output_edit.setText(str(path if path.suffix.lower() == ".xlsx" else path.with_suffix(".xlsx")))

    def update_selection_summary(self) -> None:
        addresses = self.viewer.selected_addresses()
        self.selection_summary.setText(f"已选择 {len(addresses)} 项：{format_cell_addresses(addresses) or '尚未选择'}")

    def run_summary(self) -> None:
        self.progress.setValue(0)
        try:
            if not self.source_files: raise ValueError("请先选择数据来源。")
            addresses = self.viewer.selected_addresses()
            if not addresses: raise ValueError("请在工作表预览中选择需要提取的单元格。")
            summary_name = self.summary_name.text().strip() or "汇总"
            output_text = self.output_edit.text().strip()
            if not output_text: raise ValueError("请选择汇总文件的保存位置。")
            included = None if self.all_sheets.isChecked() else {item.text() for item in self.sheet_list.selectedItems()}
            if included == set(): raise ValueError("请至少选择一个需要处理的工作表名称。")
            def progress(current, total, message):
                self.status.setText(message); self.progress.setValue(int(current / max(total, 1) * 100)); QApplication.processEvents()
            create_new_summary_workbook(self.source_files, addresses, Path(output_text), summary_name, progress, included)
            self.status.setText("汇总完成")
            QMessageBox.information(self, "汇总完成", f"已生成：\n{output_text}")
        except Exception as exc:
            self.status.setText("汇总失败"); QMessageBox.critical(self, "汇总失败", str(exc))
        finally:
            self.progress.setValue(0)

    def closeEvent(self, event) -> None:
        self.viewer.close(); super().closeEvent(event)
