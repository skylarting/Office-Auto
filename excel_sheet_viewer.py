"""Reusable, virtualized single/dual Excel worksheet viewer for Qt apps."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import (
    QAbstractTableModel, QEvent, QItemSelectionModel, QModelIndex, QSignalBlocker,
    Qt, Signal,
)
from PySide6.QtGui import QAction, QBrush, QColor, QKeySequence, QPainter, QPalette, QPen, QShortcut
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QFrame, QHBoxLayout, QLabel, QMenu, QPushButton,
    QSlider, QStyle, QStyledItemDelegate, QStyleOptionViewItem, QTableView,
    QToolButton, QVBoxLayout, QWidget,
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
        rows, columns = self.reader.visible_dimensions(sheet_name)
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
        clean.state = QStyle.StateFlag.State_Enabled
        if option.state & QStyle.StateFlag.State_MouseOver:
            clean.state |= QStyle.StateFlag.State_MouseOver
        # Rebuild the non-selected appearance explicitly. Windows/Fusion can
        # keep painting a blue selection even after State_Selected is removed.
        original_background = index.data(Qt.ItemDataRole.BackgroundRole)
        if original_background:
            painter.fillRect(option.rect, original_background)
        else:
            painter.fillRect(option.rect, clean.palette.brush(QPalette.ColorRole.Base))
        clean.palette.setBrush(
            QPalette.ColorRole.Highlight, QBrush(QColor(0, 0, 0, 0))
        )
        clean.palette.setBrush(
            QPalette.ColorGroup.Inactive, QPalette.ColorRole.Highlight,
            QBrush(QColor(0, 0, 0, 0)),
        )
        clean.palette.setBrush(
            QPalette.ColorGroup.Disabled, QPalette.ColorRole.Highlight,
            QBrush(QColor(0, 0, 0, 0)),
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
        self.use_click_order = False
        self._manual_order: list[str] = []
        self._selection_history: list[list[str]] = [[]]
        self._restoring_selection = False
        self._batching_selection = False
        root = QVBoxLayout(self)
        root.setContentsMargins(10, 10, 10, 10)
        root.setSpacing(7)
        heading_line = QHBoxLayout()
        self.heading = QLabel(title)
        self.heading.setObjectName("sheetPaneTitle")
        self.heading.setWordWrap(True)
        heading_line.addWidget(self.heading)
        heading_line.addStretch(1)
        heading_line.addWidget(QLabel("缩放"))
        self.zoom_slider = QSlider(Qt.Orientation.Horizontal)
        self.zoom_slider.setRange(40, 180)
        self.zoom_slider.setValue(100)
        self.zoom_slider.setMinimumWidth(150)
        self.zoom_slider.setMaximumWidth(310)
        self.zoom_slider.valueChanged.connect(self.set_zoom)
        heading_line.addWidget(self.zoom_slider, 1)
        self.zoom_label = QLabel("100%")
        self.zoom_label.setMinimumWidth(42)
        heading_line.addWidget(self.zoom_label)
        root.addLayout(heading_line)

        tools = QHBoxLayout()
        self.tools_layout = tools
        select_button = QToolButton()
        select_button.setText("条件筛选 ▾")
        select_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        select_menu = QMenu(select_button)
        select_button.setMenu(select_menu)
        self.select_menu = select_menu
        self.trait_labels = {
            "nonempty": "所有包含内容的单元格",
            "formula": "公式",
            "number": "数值（非公式）",
            "text": "文字",
            "fill": "带底色",
        }
        nonempty = select_menu.addAction(self.trait_labels["nonempty"])
        nonempty.triggered.connect(lambda: self.toggle_trait("nonempty"))
        select_menu.addSeparator()
        tools.addStretch(1)
        self.trait_actions: dict[str, QAction] = {"nonempty": nonempty}
        for text, trait in (
            ("公式", "formula"), ("数值（非公式）", "number"),
            ("文字", "text"), ("带底色", "fill"),
        ):
            action = select_menu.addAction(text)
            action.triggered.connect(lambda _checked=False, kind=trait: self.toggle_trait(kind))
            self.trait_actions[trait] = action
        select_menu.aboutToShow.connect(self.refresh_trait_states)
        tools.insertWidget(0, select_button)
        clear = QPushButton("清空")
        clear.setObjectName("compactButton")
        clear.setToolTip("取消当前工作表中的全部红框选择")
        clear.clicked.connect(self.clear_selection)
        tools.addWidget(clear)
        root.addLayout(tools)

        self.table = QTableView()
        self.table.setItemDelegate(RedOutlineDelegate(self.table))
        # The delegate draws the only selection indicator. Prevent the native
        # style sheet/palette from adding a blue selection fill underneath it.
        palette = self.table.palette()
        palette.setBrush(QPalette.ColorRole.Highlight, QBrush(QColor(0, 0, 0, 0)))
        palette.setBrush(QPalette.ColorGroup.Inactive, QPalette.ColorRole.Highlight, QBrush(QColor(0, 0, 0, 0)))
        palette.setBrush(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Highlight, QBrush(QColor(0, 0, 0, 0)))
        palette.setBrush(QPalette.ColorRole.HighlightedText, palette.brush(QPalette.ColorRole.Text))
        self.table.setPalette(palette)
        self.table.setStyleSheet(
            "QTableView::item:selected { background: transparent; color: palette(text); }"
        )
        # MultiSelection makes an ordinary click toggle one cell, matching the
        # behavior business users expect from the earlier mapping picker.
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.MultiSelection)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(False)
        self.table.horizontalHeader().setDefaultSectionSize(105)
        self.table.verticalHeader().setDefaultSectionSize(28)
        self.table.clicked.connect(self._record_clicked_cell)
        self.table.viewport().installEventFilter(self)
        self.undo_shortcut = QShortcut(QKeySequence.StandardKey.Undo, self.table)
        self.undo_shortcut.activated.connect(self.undo_selection)
        root.addWidget(self.table, 1)

    def eventFilter(self, watched, event) -> bool:
        if watched is self.table.viewport() and event.type() == QEvent.Type.Wheel:
            if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
                delta = event.angleDelta().y()
                if delta:
                    self.zoom_slider.setValue(
                        self.zoom_slider.value() + (5 if delta > 0 else -5)
                    )
                event.accept()
                return True
        return super().eventFilter(watched, event)

    def load_sheet(self, path: Path, sheet_name: str) -> None:
        if self.model:
            self.model.close()
        self.model = WorksheetModel(path, sheet_name, self)
        self.table.setModel(self.model)
        self.heading.setText(self.title_text)
        self._selection_history = [[]]
        self.table.selectionModel().selectionChanged.connect(self._selection_updated)
        self.set_zoom(self.zoom_slider.value())

    def clear_sheet(self) -> None:
        """Remove the previous workbook view instead of leaving stale content visible."""
        if self.model:
            self.model.close()
            self.model = None
        self.table.setModel(None)
        self.heading.setText(self.title_text)
        self._manual_order.clear()
        self._selection_history = [[]]
        self.refresh_trait_states()

    def add_toolbar_button(self, text: str, callback) -> QPushButton:
        button = QPushButton(text)
        button.clicked.connect(callback)
        self.tools_layout.insertWidget(2, button)
        return button

    def selected_addresses(self) -> list[str]:
        if not self.table.selectionModel():
            return []
        indexes = self.table.selectionModel().selectedIndexes()
        selected = {address_for(item.row(), item.column()) for item in indexes}
        if self.use_click_order:
            ordered = [address for address in self._manual_order if address in selected]
            missing = sorted(selected - set(ordered), key=self._address_sort_key)
            return ordered + missing
        return sorted(selected, key=self._address_sort_key)

    @staticmethod
    def _address_sort_key(address: str) -> tuple[int, int]:
        from excel_mapper import split_address
        row, column = split_address(address)
        return row, column

    def _record_clicked_cell(self, index: QModelIndex) -> None:
        if not self.use_click_order or not self.table.selectionModel():
            return
        address = address_for(index.row(), index.column())
        selected = index in self.table.selectionModel().selectedIndexes()
        if selected and address not in self._manual_order:
            self._manual_order.append(address)
        elif not selected and address in self._manual_order:
            self._manual_order.remove(address)
        self.selection_changed.emit(self.selected_addresses())

    def set_click_order(self, enabled: bool) -> None:
        self.use_click_order = enabled
        self._manual_order = self.selected_addresses()

    def selected_text(self) -> str:
        return format_cell_addresses(self.selected_addresses())

    def clear_selection(self) -> None:
        self._batching_selection = True
        try:
            self.table.clearSelection()
            self._manual_order.clear()
        finally:
            self._batching_selection = False
        self._selection_updated()

    def _selection_updated(self, *_args) -> None:
        addresses = self.selected_addresses()
        if self._batching_selection:
            return
        if not self._restoring_selection and (
            not self._selection_history or addresses != self._selection_history[-1]
        ):
            self._selection_history.append(addresses)
            self._selection_history = self._selection_history[-100:]
        self.selection_changed.emit(addresses)
        self.refresh_trait_states()

    def undo_selection(self) -> None:
        """Undo the most recent cell-selection change (Ctrl+Z)."""
        if len(self._selection_history) < 2:
            return
        self._selection_history.pop()
        previous = list(self._selection_history[-1])
        self._restoring_selection = True
        try:
            self.select_addresses(previous, clear=True)
        finally:
            self._restoring_selection = False
        self.selection_changed.emit(self.selected_addresses())
        self.refresh_trait_states()

    def trait_state(self, trait: str) -> str:
        """Return none/partial/all based only on the actual red-box selection."""
        if not self.model or not self.table.selectionModel():
            return "none"
        matching = {
            (row, column)
            for row in range(self.model.rows)
            for column in range(self.model.columns)
            if trait in self.model.traits_at(row, column)
        }
        if not matching:
            return "none"
        selected = {
            (index.row(), index.column())
            for index in self.table.selectionModel().selectedIndexes()
        }
        count = len(matching & selected)
        if count == 0:
            return "none"
        return "all" if count == len(matching) else "partial"

    def refresh_trait_states(self) -> None:
        symbols = {"none": "☐", "partial": "—", "all": "✓"}
        for trait, action in self.trait_actions.items():
            state = self.trait_state(trait)
            action.setText(f"{symbols[state]}  {self.trait_labels[trait]}")
            action.setToolTip(
                "再次点击将取消此类全部单元格"
                if state == "all"
                else "点击后选中此类全部单元格"
            )

    def toggle_trait(self, trait: str) -> None:
        self.select_trait(trait, self.trait_state(trait) != "all")
        self.refresh_trait_states()

    def select_trait(self, trait: str, select: bool) -> None:
        if not self.model or not self.table.selectionModel():
            return
        selection = self.table.selectionModel()
        flag = (
            QItemSelectionModel.SelectionFlag.Select
            if select else QItemSelectionModel.SelectionFlag.Deselect
        )
        self._batching_selection = True
        try:
            for row in range(self.model.rows):
                for column in range(self.model.columns):
                    traits = self.model.traits_at(row, column)
                    if trait in traits:
                        selection.select(self.model.index(row, column), flag)
        finally:
            self._batching_selection = False
        self._selection_updated()

    def select_addresses(self, addresses: list[str], clear: bool = True) -> None:
        if not self.model or not self.table.selectionModel():
            return
        from excel_mapper import split_address
        self._batching_selection = True
        try:
            if clear:
                self.table.clearSelection()
                self._manual_order.clear()
            selection = self.table.selectionModel()
            for address in addresses:
                try:
                    row, column = split_address(address)
                except ValueError:
                    continue
                if row < self.model.rows and column < self.model.columns:
                    normalized = address_for(row, column)
                    if normalized not in self._manual_order:
                        self._manual_order.append(normalized)
                    selection.select(
                        self.model.index(row, column),
                        QItemSelectionModel.SelectionFlag.Select,
                    )
        finally:
            self._batching_selection = False
        if not self._restoring_selection:
            self._selection_updated()

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
        self.link_scroll.setChecked(True)
        top.addWidget(self.link_scroll)
        self.link_zoom = QCheckBox("同步缩放")
        self.link_zoom.setChecked(True)
        top.addWidget(self.link_zoom)
        fit = QPushButton("适应窗口")
        fit.clicked.connect(self.fit_both)
        top.addWidget(fit)
        top.addWidget(QLabel("缩放"))
        self.zoom_slider = QSlider(Qt.Orientation.Horizontal)
        self.zoom_slider.setRange(40, 180)
        self.zoom_slider.setValue(100)
        self.zoom_slider.setMaximumWidth(240)
        top.addWidget(self.zoom_slider)
        self.zoom_label = QLabel("100%")
        top.addWidget(self.zoom_label)
        top.addStretch(1)
        self.controls_widget = QWidget()
        self.controls_widget.setLayout(top)
        root.addWidget(self.controls_widget)
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
        self.source.zoom_slider.valueChanged.connect(
            lambda value: self._sync_zoom(self.target, value)
        )
        self.target.zoom_slider.valueChanged.connect(
            lambda value: self._sync_zoom(self.source, value)
        )
        self.zoom_slider.valueChanged.connect(self.set_zoom)

    def load_pair(
        self, source_path: Path, source_sheet: str,
        target_path: Path, target_sheet: str,
    ) -> None:
        self.source.load_sheet(source_path, source_sheet)
        self.target.load_sheet(target_path, target_sheet)

    def clear_pair(self) -> None:
        self.source.clear_sheet()
        self.target.clear_sheet()

    def _sync_scroll(self, pane: SheetViewPane, value: int, vertical: bool) -> None:
        if not self.link_scroll.isChecked() or self._syncing:
            return
        self._syncing = True
        try:
            bar = pane.table.verticalScrollBar() if vertical else pane.table.horizontalScrollBar()
            bar.setValue(value)
        finally:
            self._syncing = False

    def _sync_zoom(self, pane: SheetViewPane, value: int) -> None:
        if not self.link_zoom.isChecked() or self._syncing:
            return
        self._syncing = True
        try:
            pane.zoom_slider.setValue(value)
        finally:
            self._syncing = False

    def fit_both(self) -> None:
        self.source.fit_window()
        if self.link_zoom.isChecked():
            self.target.zoom_slider.setValue(self.source.zoom_slider.value())
        else:
            self.target.fit_window()

    def set_zoom(self, value: int) -> None:
        self.source.zoom_slider.setValue(value)
        if self.link_zoom.isChecked():
            self.target.zoom_slider.setValue(value)

    def close(self) -> None:
        self.source.close()
        self.target.close()
