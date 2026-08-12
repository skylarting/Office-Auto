"""Reusable, virtualized single/dual Excel worksheet viewer for Qt apps."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import (
    QAbstractTableModel, QItemSelectionModel, QModelIndex, QSignalBlocker,
    Qt, Signal,
)
from PySide6.QtGui import QBrush, QColor, QPainter, QPalette, QPen
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QFrame, QHBoxLayout, QLabel, QPushButton,
    QSlider, QStyle, QStyledItemDelegate, QStyleOptionViewItem, QTableView,
    QVBoxLayout, QWidget,
)

from excel_mapper import WorkbookReader, format_cell_addresses


def address_for(row: int, column: int) -> str:
    letters = ""
    number = column + 1
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return f"{letters}{row + 1}"


class WorksheetModel(QAbstractTableModel):
    """Read cells on demand so a complete worksheet need not be materialized."""

    def __init__(self, path: Path, sheet_name: str, parent=None) -> None:
        super().__init__(parent)
        self.path = path
        self.sheet_name = sheet_name
        self.reader = WorkbookReader(path)
        rows, columns = self.reader.used_dimensions(sheet_name)
        self.rows = max(rows, 1)
        self.columns = max(columns, 1)
        self._value_cache: dict[tuple[int, int], object] = {}
        self._fill_cache: dict[tuple[int, int], str | None] = {}
        self._trait_cache: dict[tuple[int, int], set[str]] = {}

    def close(self) -> None:
        self.reader.close()

    def rowCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else self.rows

    def columnCount(self, parent=QModelIndex()) -> int:
        return 0 if parent.isValid() else self.columns

    def value_at(self, row: int, column: int):
        key = (row, column)
        if key not in self._value_cache:
            self._value_cache[key] = self.reader.read(
                self.sheet_name, address_for(row, column)
            ).value
        return self._value_cache[key]

    def traits_at(self, row: int, column: int) -> set[str]:
        key = (row, column)
        if key not in self._trait_cache:
            self._trait_cache[key] = self.reader.cell_traits(
                self.sheet_name, address_for(row, column)
            )
        return self._trait_cache[key]

    def data(self, index: QModelIndex, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None
        row, column = index.row(), index.column()
        if role == Qt.ItemDataRole.DisplayRole:
            value = self.value_at(row, column)
            return "" if value is None else str(value)
        if role == Qt.ItemDataRole.BackgroundRole:
            key = (row, column)
            if key not in self._fill_cache:
                self._fill_cache[key] = self.reader.fill_color(
                    self.sheet_name, address_for(row, column)
                )
            color = self._fill_cache[key]
            return QBrush(QColor(color)) if color else None
        if role == Qt.ItemDataRole.TextAlignmentRole:
            traits = self.traits_at(row, column)
            return int(
                Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
                if "number" in traits or "formula" in traits
                else Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
            )
        if role == Qt.ItemDataRole.ToolTipRole:
            return f"{address_for(row, column)}  {self.data(index)}"
        return None

    def headerData(self, section: int, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        return address_for(0, section)[:-1] if orientation == Qt.Orientation.Horizontal else str(section + 1)


class RedOutlineDelegate(QStyledItemDelegate):
    """Keep workbook fill colors visible and mark selected cells with a red box."""

    def paint(self, painter: QPainter, option, index) -> None:
        selected = bool(option.state & QStyle.StateFlag.State_Selected)
        clean = QStyleOptionViewItem(option)
        clean.state &= ~QStyle.StateFlag.State_Selected
        # Some platform styles still use Highlight after State_Selected is
        # removed. Make both highlight brushes transparent while leaving the
        # model's original BackgroundRole untouched.
        clean.palette.setBrush(
            QPalette.ColorRole.Highlight, QBrush(QColor(0, 0, 0, 0))
        )
        clean.palette.setBrush(
            QPalette.ColorRole.HighlightedText,
            clean.palette.brush(QPalette.ColorRole.Text),
        )
        super().paint(painter, clean, index)
        if selected:
            painter.save()
            painter.setPen(QPen(QColor("#D92D20"), 2))
            painter.drawRect(option.rect.adjusted(1, 1, -2, -2))
            painter.restore()


class SheetViewPane(QFrame):
    selection_changed = Signal(list)

    def __init__(self, title: str, parent=None) -> None:
        super().__init__(parent)
        self.setObjectName("sheetPane")
        self.title_text = title
        self.model: WorksheetModel | None = None
        self.zoom = 100
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(7)
        self.heading = QLabel(title)
        self.heading.setObjectName("sheetPaneTitle")
        self.heading.setWordWrap(True)
        root.addWidget(self.heading)

        tools = QHBoxLayout()
        self.tools_layout = tools
        clear = QPushButton("取消全部")
        clear.clicked.connect(self.clear_selection)
        tools.addWidget(clear)
        nonempty = QPushButton("选择有值单元格")
        nonempty.clicked.connect(lambda: self.select_trait("nonempty", True))
        tools.addWidget(nonempty)
        tools.addStretch(1)
        self.trait_buttons: dict[str, QPushButton] = {}
        for text, trait in (
            ("公式", "formula"), ("数值（非公式）", "number"),
            ("文字", "text"), ("带底色", "fill"),
        ):
            button = QPushButton(text)
            button.setCheckable(True)
            button.toggled.connect(
                lambda checked, kind=trait: self.select_trait(kind, checked)
            )
            tools.addWidget(button)
            self.trait_buttons[trait] = button
        root.addLayout(tools)

        zoom_line = QHBoxLayout()
        fit = QPushButton("适合窗口")
        fit.clicked.connect(self.fit_window)
        zoom_line.addWidget(fit)
        zoom_line.addWidget(QLabel("缩放"))
        self.zoom_slider = QSlider(Qt.Orientation.Horizontal)
        self.zoom_slider.setRange(40, 180)
        self.zoom_slider.setValue(100)
        self.zoom_slider.valueChanged.connect(self.set_zoom)
        zoom_line.addWidget(self.zoom_slider, 1)
        self.zoom_label = QLabel("100%")
        zoom_line.addWidget(self.zoom_label)
        root.addLayout(zoom_line)

        self.table = QTableView()
        self.table.setItemDelegate(RedOutlineDelegate(self.table))
        # MultiSelection makes an ordinary click toggle one cell, matching the
        # behavior business users expect from the earlier mapping picker.
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(False)
        self.table.horizontalHeader().setDefaultSectionSize(105)
        self.table.verticalHeader().setDefaultSectionSize(28)
        root.addWidget(self.table, 1)

    def load_sheet(self, path: Path, sheet_name: str) -> None:
        if self.model:
            self.model.close()
        self.model = WorksheetModel(path, sheet_name, self)
        self.table.setModel(self.model)
        self.heading.setText(f"{self.title_text}\n{path.name} / {sheet_name}")
        self.table.selectionModel().selectionChanged.connect(
            lambda *_: self.selection_changed.emit(self.selected_addresses())
        )
        self.set_zoom(self.zoom_slider.value())

    def add_toolbar_button(self, text: str, callback) -> QPushButton:
        button = QPushButton(text)
        button.clicked.connect(callback)
        self.tools_layout.insertWidget(2, button)
        return button

    def selected_addresses(self) -> list[str]:
        if not self.table.selectionModel():
            return []
        indexes = sorted(
            self.table.selectionModel().selectedIndexes(),
            key=lambda item: (item.column(), item.row()),
        )
        return [address_for(item.row(), item.column()) for item in indexes]

    def selected_text(self) -> str:
        return format_cell_addresses(self.selected_addresses())

    def clear_selection(self) -> None:
        self.table.clearSelection()
        for button in self.trait_buttons.values():
            with QSignalBlocker(button):
                button.setChecked(False)

    def select_trait(self, trait: str, select: bool) -> None:
        if not self.model or not self.table.selectionModel():
            return
        selection = self.table.selectionModel()
        flag = (
            QItemSelectionModel.SelectionFlag.Select
            if select else QItemSelectionModel.SelectionFlag.Deselect
        )
        for row in range(self.model.rows):
            for column in range(self.model.columns):
                traits = self.model.traits_at(row, column)
                if trait in traits:
                    selection.select(self.model.index(row, column), flag)

    def select_addresses(self, addresses: list[str], clear: bool = True) -> None:
        if not self.model or not self.table.selectionModel():
            return
        from excel_mapper import split_address
        if clear:
            self.table.clearSelection()
        selection = self.table.selectionModel()
        for address in addresses:
            try:
                row, column = split_address(address)
            except ValueError:
                continue
            if row < self.model.rows and column < self.model.columns:
                selection.select(
                    self.model.index(row, column),
                    QItemSelectionModel.SelectionFlag.Select,
                )

    def scroll_to_address(self, address: str) -> None:
        if not self.model:
            return
        from excel_mapper import split_address
        row, column = split_address(address)
        if row < self.model.rows and column < self.model.columns:
            self.table.scrollTo(
                self.model.index(row, column),
                QAbstractItemView.ScrollHint.PositionAtCenter,
            )

    def set_zoom(self, value: int) -> None:
        self.zoom = value
        self.zoom_label.setText(f"{value}%")
        self.table.horizontalHeader().setDefaultSectionSize(max(42, int(105 * value / 100)))
        self.table.verticalHeader().setDefaultSectionSize(max(18, int(28 * value / 100)))

    def fit_window(self) -> None:
        if not self.model:
            return
        visible_columns = min(self.model.columns, 8)
        width = max(self.table.viewport().width(), 320)
        size = max(42, int(width / max(visible_columns, 1)))
        percent = max(40, min(180, round(size / 105 * 100)))
        self.zoom_slider.setValue(percent)

    def close(self) -> None:
        if self.model:
            self.model.close()
            self.model = None


class DualSheetViewer(QWidget):
    """Two complete worksheet viewers with optional linked scrolling."""

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        top = QHBoxLayout()
        self.link_scroll = QCheckBox("同步滚动")
        self.link_scroll.setChecked(False)
        top.addWidget(self.link_scroll)
        top.addStretch(1)
        root.addLayout(top)
        panes = QHBoxLayout()
        self.source = SheetViewPane("来源数据")
        self.target = SheetViewPane("要填写的报表")
        panes.addWidget(self.source, 1)
        panes.addWidget(self.target, 1)
        root.addLayout(panes, 1)
        self._syncing = False
        self.source.table.verticalScrollBar().valueChanged.connect(
            lambda value: self._sync_scroll(self.target, value, True)
        )
        self.target.table.verticalScrollBar().valueChanged.connect(
            lambda value: self._sync_scroll(self.source, value, True)
        )
        self.source.table.horizontalScrollBar().valueChanged.connect(
            lambda value: self._sync_scroll(self.target, value, False)
        )
        self.target.table.horizontalScrollBar().valueChanged.connect(
            lambda value: self._sync_scroll(self.source, value, False)
        )

    def load_pair(
        self, source_path: Path, source_sheet: str,
        target_path: Path, target_sheet: str,
    ) -> None:
        self.source.load_sheet(source_path, source_sheet)
        self.target.load_sheet(target_path, target_sheet)

    def _sync_scroll(self, pane: SheetViewPane, value: int, vertical: bool) -> None:
        if not self.link_scroll.isChecked() or self._syncing:
            return
        self._syncing = True
        try:
            bar = pane.table.verticalScrollBar() if vertical else pane.table.horizontalScrollBar()
            bar.setValue(value)
        finally:
            self._syncing = False

    def close(self) -> None:
        self.source.close()
        self.target.close()
