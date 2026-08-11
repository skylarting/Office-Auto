"""Five-step Qt workflow for smart label-based template filling."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSpinBox,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from excel_mapper import workbook_sheet_names
from smart_template import (
    BlockRule,
    FormulaRule,
    SheetStructure,
    SmartTemplatePlan,
    detect_sheet_structure,
    execute_smart_plan,
    load_plan,
    match_pair,
    matches_to_mapping_rules,
    save_plan,
    suggest_sheet_pairs,
)


def _compact_path(path: str) -> str:
    return Path(path).name if path else ""


class StructureDialog(QDialog):
    def __init__(self, parent, title: str, structure: SheetStructure) -> None:
        super().__init__(parent)
        self.setWindowTitle(title)
        self.result_structure: SheetStructure | None = None
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.row_cols = QLineEdit(",".join(str(value + 1) for value in structure.row_label_columns))
        self.header_rows = QLineEdit(",".join(str(value + 1) for value in structure.column_header_rows))
        self.data_rows = QLineEdit(self._range_text(structure.data_rows))
        self.data_cols = QLineEdit(self._range_text(structure.data_columns))
        form.addRow("行标题列序号：", self.row_cols)
        form.addRow("列标题行序号：", self.header_rows)
        form.addRow("数据行范围：", self.data_rows)
        form.addRow("数据列范围：", self.data_cols)
        layout.addLayout(form)
        note = QLabel("多个序号用逗号分隔；范围可写 8-42。所有序号从 1 开始。")
        note.setObjectName("secondaryText")
        layout.addWidget(note)
        actions = QHBoxLayout()
        actions.addStretch(1)
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        actions.addWidget(cancel)
        confirm = QPushButton("保存结构")
        confirm.setObjectName("primaryButton")
        confirm.clicked.connect(self._accept)
        actions.addWidget(confirm)
        layout.addLayout(actions)

    @staticmethod
    def _range_text(values: list[int]) -> str:
        if not values:
            return ""
        return str(values[0] + 1) if len(values) == 1 else f"{values[0] + 1}-{values[-1] + 1}"

    @staticmethod
    def _parse(text: str) -> list[int]:
        result = []
        for part in text.replace("，", ",").split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                start, end = (int(value) for value in part.split("-", 1))
                result.extend(range(start - 1, end))
            else:
                result.append(int(part) - 1)
        if any(value < 0 for value in result):
            raise ValueError("行列序号必须大于 0。")
        return sorted(dict.fromkeys(result))

    def _accept(self) -> None:
        try:
            self.result_structure = SheetStructure(
                self._parse(self.row_cols.text()),
                self._parse(self.header_rows.text()),
                self._parse(self.data_rows.text()),
                self._parse(self.data_cols.text()),
                1.0,
            )
        except Exception as exc:
            QMessageBox.warning(self, "结构设置不正确", str(exc))
            return
        self.accept()


class RuleDialog(QDialog):
    def __init__(self, parent, kind: str, sheets: list[str]) -> None:
        super().__init__(parent)
        self.kind = kind
        self.rule = None
        self.setWindowTitle("添加公式规则" if kind == "formula" else "添加禁填规则")
        layout = QVBoxLayout(self)
        form = QFormLayout()
        self.sheet = QComboBox()
        self.sheet.addItem("全部工作表", "*")
        for name in sheets:
            self.sheet.addItem(name, name)
        self.row = QLineEdit()
        self.row.setPlaceholderText("留空表示任意行，例如：合计")
        self.column = QLineEdit()
        self.column.setPlaceholderText("留空表示任意列，例如：人民币")
        form.addRow("目标工作表：", self.sheet)
        form.addRow("行名包含：", self.row)
        form.addRow("列名包含：", self.column)
        if kind == "formula":
            self.formula = QLineEdit()
            self.formula.setPlaceholderText("例如 =C{row}+D{row}；支持 {row}、{col}、{cell}")
            self.policy = QComboBox()
            self.policy.addItems(["仅填空白", "覆盖数值但保留已有公式", "强制覆盖"])
            form.addRow("公式模板：", self.formula)
            form.addRow("覆盖方式：", self.policy)
        else:
            self.scope = QComboBox()
            self.scope.addItems(["交叉位置", "整行", "整列"])
            self.reason = QLineEdit("禁填区域")
            form.addRow("作用范围：", self.scope)
            form.addRow("原因：", self.reason)
        layout.addLayout(form)
        actions = QHBoxLayout()
        actions.addStretch(1)
        cancel = QPushButton("取消")
        cancel.clicked.connect(self.reject)
        actions.addWidget(cancel)
        save = QPushButton("保存规则")
        save.setObjectName("primaryButton")
        save.clicked.connect(self._save)
        actions.addWidget(save)
        layout.addLayout(actions)

    def _save(self) -> None:
        sheet = str(self.sheet.currentData())
        if self.kind == "formula":
            if not self.formula.text().strip():
                QMessageBox.warning(self, "缺少公式", "请填写公式模板。")
                return
            self.rule = FormulaRule(
                sheet,
                self.row.text().strip(),
                self.column.text().strip(),
                self.formula.text().strip(),
                self.policy.currentText(),
            )
        else:
            if not self.row.text().strip() and not self.column.text().strip():
                QMessageBox.warning(self, "缺少条件", "行名或列名至少填写一项。")
                return
            self.rule = BlockRule(
                sheet,
                self.row.text().strip(),
                self.column.text().strip(),
                self.scope.currentText(),
                self.reason.text().strip() or "禁填区域",
            )
        self.accept()


class SmartTemplateWindow(QMainWindow):
    def __init__(self, style_sheet: str = "") -> None:
        super().__init__()
        self.plan = SmartTemplatePlan()
        self.output_folder: Path | None = None
        self.setWindowTitle("智能模板填报")
        self.resize(1380, 850)
        self.setMinimumSize(980, 650)
        if style_sheet:
            self.setStyleSheet(style_sheet)
        self._build_ui()

    def _build_ui(self) -> None:
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(18, 14, 18, 14)
        header = QHBoxLayout()
        title_box = QVBoxLayout()
        title = QLabel("智能模板填报")
        title.setObjectName("pageTitle")
        title_box.addWidget(title)
        subtitle = QLabel("自动识别工作表、多级行列标题和单元格关系；确认后另存目标副本。")
        subtitle.setObjectName("pageSubtitle")
        title_box.addWidget(subtitle)
        header.addLayout(title_box, 1)
        load = QPushButton("载入方案")
        load.clicked.connect(self.load_project)
        header.addWidget(load)
        save = QPushButton("保存方案")
        save.clicked.connect(self.save_project)
        header.addWidget(save)
        root.addLayout(header)

        self.tabs = QTabWidget()
        self.tabs.addTab(self._files_page(), "1  文件与工作表配对")
        self.tabs.addTab(self._structure_page(), "2  表格结构识别")
        self.tabs.addTab(self._mapping_page(), "3  单元格映射确认")
        self.tabs.addTab(self._rules_page(), "4  公式与禁填规则")
        self.tabs.addTab(self._execute_page(), "5  预检查与执行")
        root.addWidget(self.tabs, 1)

    def _files_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        actions = QHBoxLayout()
        target = QPushButton("选择目标模板")
        target.clicked.connect(self.choose_target)
        actions.addWidget(target)
        source = QPushButton("添加来源文件")
        source.clicked.connect(self.add_sources)
        actions.addWidget(source)
        folder = QPushButton("添加来源文件夹")
        folder.clicked.connect(self.add_source_folder)
        actions.addWidget(folder)
        auto = QPushButton("自动配对")
        auto.setObjectName("primaryButton")
        auto.clicked.connect(self.auto_pair)
        actions.addWidget(auto)
        actions.addStretch(1)
        layout.addLayout(actions)
        self.file_summary = QLabel("尚未选择目标模板和来源文件。")
        layout.addWidget(self.file_summary)
        self.pair_table = QTableWidget(0, 5)
        self.pair_table.setHorizontalHeaderLabels(
            ["来源工作簿", "来源工作表", "目标工作表（可修改）", "匹配度", "确认"]
        )
        self.pair_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.pair_table.itemChanged.connect(self.pair_item_changed)
        layout.addWidget(self.pair_table, 1)
        return page

    def _structure_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        line = QHBoxLayout()
        note = QLabel("程序先识别标题区域；复杂表格可选择一行后手动修正。")
        note.setObjectName("secondaryText")
        line.addWidget(note, 1)
        refresh = QPushButton("重新识别结构")
        refresh.clicked.connect(self.detect_structures)
        line.addWidget(refresh)
        edit_source = QPushButton("修改来源结构")
        edit_source.clicked.connect(lambda: self.edit_structure(True))
        line.addWidget(edit_source)
        edit_target = QPushButton("修改目标结构")
        edit_target.clicked.connect(lambda: self.edit_structure(False))
        line.addWidget(edit_target)
        layout.addLayout(line)
        self.structure_table = QTableWidget(0, 5)
        self.structure_table.setHorizontalHeaderLabels(
            ["来源工作表", "来源识别结果", "目标工作表", "目标识别结果", "置信度"]
        )
        self.structure_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.structure_table, 1)
        return page

    def _mapping_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        actions = QHBoxLayout()
        generate = QPushButton("生成智能映射")
        generate.setObjectName("primaryButton")
        generate.clicked.connect(self.generate_matches)
        actions.addWidget(generate)
        pending = QPushButton("仅显示待确认")
        pending.clicked.connect(lambda: self.refresh_matches(True))
        actions.addWidget(pending)
        show_all = QPushButton("显示全部")
        show_all.clicked.connect(lambda: self.refresh_matches(False))
        actions.addWidget(show_all)
        actions.addStretch(1)
        self.match_summary = QLabel("尚未生成映射。")
        actions.addWidget(self.match_summary)
        layout.addLayout(actions)
        self.match_table = QTableWidget(0, 8)
        self.match_table.setHorizontalHeaderLabels(
            ["使用", "来源行路径", "来源列路径", "来源", "目标行路径", "目标列路径", "目标地址", "状态"]
        )
        self.match_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        for col, width in enumerate((55, 240, 210, 85, 240, 210, 90, 90)):
            self.match_table.setColumnWidth(col, width)
        self.match_table.itemChanged.connect(self.match_item_changed)
        layout.addWidget(self.match_table, 1)
        return page

    def _rules_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        self.overwrite_formula_check = QCheckBox(
            "允许普通数值映射覆盖目标已有公式（高风险，默认关闭）"
        )
        self.overwrite_formula_check.toggled.connect(
            lambda checked: setattr(
                self.plan, "overwrite_target_formulas", checked
            )
        )
        layout.addWidget(self.overwrite_formula_check)
        rule_tabs = QTabWidget()
        formula_page = QWidget()
        formula_layout = QVBoxLayout(formula_page)
        f_actions = QHBoxLayout()
        add_formula = QPushButton("添加公式规则")
        add_formula.clicked.connect(self.add_formula_rule)
        f_actions.addWidget(add_formula)
        delete_formula = QPushButton("删除选中")
        delete_formula.clicked.connect(lambda: self.delete_rule(True))
        f_actions.addWidget(delete_formula)
        f_actions.addStretch(1)
        formula_layout.addLayout(f_actions)
        self.formula_table = QTableWidget(0, 5)
        self.formula_table.setHorizontalHeaderLabels(["工作表", "行名包含", "列名包含", "公式模板", "覆盖方式"])
        self.formula_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        formula_layout.addWidget(self.formula_table)
        rule_tabs.addTab(formula_page, "公式规则")

        block_page = QWidget()
        block_layout = QVBoxLayout(block_page)
        b_actions = QHBoxLayout()
        add_block = QPushButton("添加禁填规则")
        add_block.clicked.connect(self.add_block_rule)
        b_actions.addWidget(add_block)
        delete_block = QPushButton("删除选中")
        delete_block.clicked.connect(lambda: self.delete_rule(False))
        b_actions.addWidget(delete_block)
        b_actions.addStretch(1)
        block_layout.addLayout(b_actions)
        self.block_table = QTableWidget(0, 5)
        self.block_table.setHorizontalHeaderLabels(["工作表", "行名包含", "列名包含", "范围", "原因"])
        self.block_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        block_layout.addWidget(self.block_table)
        rule_tabs.addTab(block_page, "禁填规则")
        layout.addWidget(rule_tabs)
        return page

    def _execute_page(self) -> QWidget:
        page = QWidget()
        layout = QVBoxLayout(page)
        actions = QHBoxLayout()
        precheck = QPushButton("预检查")
        precheck.clicked.connect(self.precheck)
        actions.addWidget(precheck)
        output = QPushButton("选择输出文件夹")
        output.clicked.connect(self.choose_output)
        actions.addWidget(output)
        actions.addStretch(1)
        run = QPushButton("开始填报")
        run.setObjectName("primaryButton")
        run.clicked.connect(self.run_plan)
        actions.addWidget(run)
        layout.addLayout(actions)
        self.output_label = QLabel("输出文件夹：尚未选择")
        layout.addWidget(self.output_label)
        self.check_text = QLabel("请先完成工作表配对和单元格映射。")
        self.check_text.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.check_text.setWordWrap(True)
        layout.addWidget(self.check_text, 1)
        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        layout.addWidget(self.progress)
        return page

    def choose_target(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "选择目标模板", "", "Excel 工作簿 (*.xls *.xlsx *.xlsm)")
        if selected:
            self.plan.target_file = selected
            if self.output_folder is None:
                self.output_folder = Path(selected).parent / "智能填报结果"
            self._refresh_file_summary()

    def add_sources(self) -> None:
        selected, _ = QFileDialog.getOpenFileNames(self, "添加来源工作簿", "", "Excel 工作簿 (*.xls *.xlsx *.xlsm)")
        self.plan.source_files = list(dict.fromkeys(self.plan.source_files + selected))
        self._refresh_file_summary()

    def add_source_folder(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择来源文件夹")
        if not selected:
            return
        files = [str(path) for path in Path(selected).iterdir() if path.suffix.lower() in (".xls", ".xlsx", ".xlsm") and not path.name.startswith("~$")]
        self.plan.source_files = list(dict.fromkeys(self.plan.source_files + files))
        self._refresh_file_summary()

    def _refresh_file_summary(self) -> None:
        target = _compact_path(self.plan.target_file) or "未选择"
        self.file_summary.setText(f"目标模板：{target}　来源工作簿：{len(self.plan.source_files)} 个")
        if self.output_folder:
            self.output_label.setText(f"输出文件夹：{self.output_folder}")

    def auto_pair(self) -> None:
        if not self.plan.target_file or not self.plan.source_files:
            QMessageBox.warning(self, "缺少文件", "请先选择目标模板和来源工作簿。")
            return
        try:
            self.plan.pairs = suggest_sheet_pairs([Path(path) for path in self.plan.source_files], Path(self.plan.target_file))
        except Exception as exc:
            QMessageBox.critical(self, "自动配对失败", str(exc))
            return
        self.refresh_pairs()
        self.detect_structures()

    def refresh_pairs(self) -> None:
        self.pair_table.blockSignals(True)
        self.pair_table.setRowCount(len(self.plan.pairs))
        for row, pair in enumerate(self.plan.pairs):
            values = [_compact_path(pair.source_file), pair.source_sheet, pair.target_sheet, f"{pair.score:.0%}"]
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                if col != 2:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.pair_table.setItem(row, col, item)
            check = QTableWidgetItem()
            check.setCheckState(Qt.CheckState.Checked if pair.confirmed else Qt.CheckState.Unchecked)
            self.pair_table.setItem(row, 4, check)
        self.pair_table.blockSignals(False)

    def pair_item_changed(self, item: QTableWidgetItem) -> None:
        if item.row() >= len(self.plan.pairs):
            return
        pair = self.plan.pairs[item.row()]
        if item.column() == 2:
            pair.target_sheet = item.text().strip()
            pair.confirmed = True
        elif item.column() == 4:
            pair.confirmed = item.checkState() == Qt.CheckState.Checked

    def detect_structures(self) -> None:
        try:
            for pair in self.plan.pairs:
                from excel_mapper import WorkbookReader
                source = WorkbookReader(Path(pair.source_file))
                target = WorkbookReader(Path(pair.target_file))
                try:
                    pair.source_structure = detect_sheet_structure(source, pair.source_sheet)
                    pair.target_structure = detect_sheet_structure(target, pair.target_sheet)
                finally:
                    source.close(); target.close()
        except Exception as exc:
            QMessageBox.critical(self, "结构识别失败", str(exc))
            return
        self.refresh_structures()

    def refresh_structures(self) -> None:
        self.structure_table.setRowCount(len(self.plan.pairs))
        for row, pair in enumerate(self.plan.pairs):
            values = [pair.source_sheet, pair.source_structure.summary if pair.source_structure else "未识别", pair.target_sheet, pair.target_structure.summary if pair.target_structure else "未识别", f"{min(pair.source_structure.confidence, pair.target_structure.confidence):.0%}" if pair.source_structure and pair.target_structure else "-"]
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.structure_table.setItem(row, col, item)

    def edit_structure(self, source: bool) -> None:
        row = self.structure_table.currentRow()
        if row < 0 or row >= len(self.plan.pairs):
            QMessageBox.information(self, "请选择工作表", "请先选择需要修改的一行。")
            return
        pair = self.plan.pairs[row]
        structure = pair.source_structure if source else pair.target_structure
        if structure is None:
            return
        dialog = StructureDialog(self, "修改来源结构" if source else "修改目标结构", structure)
        if dialog.exec() and dialog.result_structure:
            if source:
                pair.source_structure = dialog.result_structure
            else:
                pair.target_structure = dialog.result_structure
            self.refresh_structures()

    def generate_matches(self) -> None:
        self.plan.matches = []
        try:
            for pair in self.plan.pairs:
                self.plan.matches.extend(match_pair(pair, self.plan.aliases))
        except Exception as exc:
            QMessageBox.critical(self, "生成映射失败", str(exc))
            return
        self.refresh_matches(False)
        self.tabs.setCurrentIndex(2)

    def refresh_matches(self, pending_only: bool = False) -> None:
        visible = [(index, match) for index, match in enumerate(self.plan.matches) if not pending_only or match.status in ("待确认", "未匹配")]
        self.match_table.blockSignals(True)
        self.match_table.setRowCount(len(visible))
        for row, (index, match) in enumerate(visible):
            check = QTableWidgetItem()
            check.setData(Qt.ItemDataRole.UserRole, index)
            check.setCheckState(Qt.CheckState.Checked if match.enabled else Qt.CheckState.Unchecked)
            self.match_table.setItem(row, 0, check)
            values = [" / ".join(match.source_row_path), " / ".join(match.source_column_path), match.source_address, " / ".join(match.target_row_path), " / ".join(match.target_column_path), match.target_address, match.status]
            for offset, value in enumerate(values, start=1):
                item = QTableWidgetItem(value)
                item.setData(Qt.ItemDataRole.UserRole, index)
                if offset != 6:
                    item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.match_table.setItem(row, offset, item)
        self.match_table.blockSignals(False)
        counts = {name: sum(match.status == name for match in self.plan.matches) for name in ("自动匹配", "相似匹配", "待确认", "未匹配")}
        self.match_summary.setText("　".join(f"{name} {count}" for name, count in counts.items()))

    def match_item_changed(self, item: QTableWidgetItem) -> None:
        index = item.data(Qt.ItemDataRole.UserRole)
        if index is None or int(index) >= len(self.plan.matches):
            return
        match = self.plan.matches[int(index)]
        if item.column() == 0:
            match.enabled = item.checkState() == Qt.CheckState.Checked
        elif item.column() == 6:
            match.target_address = item.text().strip().upper()
            match.enabled = bool(match.target_address)
            match.status = "人工修改" if match.target_address else "未匹配"

    def target_sheets(self) -> list[str]:
        return workbook_sheet_names(Path(self.plan.target_file)) if self.plan.target_file else []

    def add_formula_rule(self) -> None:
        dialog = RuleDialog(self, "formula", self.target_sheets())
        if dialog.exec() and dialog.rule:
            self.plan.formula_rules.append(dialog.rule)
            self.refresh_rules()

    def add_block_rule(self) -> None:
        dialog = RuleDialog(self, "block", self.target_sheets())
        if dialog.exec() and dialog.rule:
            self.plan.block_rules.append(dialog.rule)
            self.refresh_rules()

    def refresh_rules(self) -> None:
        self.formula_table.setRowCount(len(self.plan.formula_rules))
        for row, rule in enumerate(self.plan.formula_rules):
            for col, value in enumerate((rule.target_sheet, rule.row_pattern, rule.column_pattern, rule.formula, rule.overwrite)):
                self.formula_table.setItem(row, col, QTableWidgetItem(value))
        self.block_table.setRowCount(len(self.plan.block_rules))
        for row, rule in enumerate(self.plan.block_rules):
            for col, value in enumerate((rule.target_sheet, rule.row_pattern, rule.column_pattern, rule.scope, rule.reason)):
                self.block_table.setItem(row, col, QTableWidgetItem(value))

    def delete_rule(self, formula: bool) -> None:
        table = self.formula_table if formula else self.block_table
        rules = self.plan.formula_rules if formula else self.plan.block_rules
        rows = sorted({item.row() for item in table.selectedItems()}, reverse=True)
        for row in rows:
            rules.pop(row)
        self.refresh_rules()

    def choose_output(self) -> None:
        selected = QFileDialog.getExistingDirectory(self, "选择输出文件夹", str(self.output_folder or Path.cwd()))
        if selected:
            self.output_folder = Path(selected)
            self.output_label.setText(f"输出文件夹：{self.output_folder}")

    def precheck(self) -> bool:
        if not self.plan.target_file or not self.plan.pairs or not self.plan.matches:
            QMessageBox.warning(self, "方案不完整", "请先完成文件配对、结构识别和智能映射。")
            return False
        rules, skipped = matches_to_mapping_rules(self.plan)
        pending = sum(match.status in ("待确认", "未匹配") for match in self.plan.matches)
        text = (
            f"工作表配对：{len(self.plan.pairs)} 组\n"
            f"可写入数值：{len(rules)} 个\n"
            f"公式规则：{len(self.plan.formula_rules)} 条\n"
            f"禁填规则：{len(self.plan.block_rules)} 条\n"
            f"保留公式或禁填跳过：{len(skipped)} 个\n"
            f"待确认或未匹配：{pending} 个"
        )
        self.check_text.setText(text)
        if pending:
            QMessageBox.warning(self, "预检查完成", text + "\n\n请在第3步确认黄色或红色项目。")
            return False
        QMessageBox.information(self, "预检查通过", text)
        return True

    def run_plan(self) -> None:
        if not self.precheck():
            return
        if self.output_folder is None:
            self.output_folder = Path(self.plan.target_file).parent / "智能填报结果"
        try:
            def progress(current, total, message):
                self.progress.setValue(round(current / max(total, 1) * 100))
                self.check_text.setText(message)
            outputs, skipped = execute_smart_plan(self.plan, self.output_folder, progress)
            QMessageBox.information(self, "填报完成", "已生成：\n" + "\n".join(str(path) for path in outputs) + f"\n\n跳过 {len(skipped)} 个受保护位置。")
        except Exception as exc:
            QMessageBox.critical(self, "填报失败", str(exc))
        finally:
            self.progress.setValue(0)

    def save_project(self) -> None:
        selected, _ = QFileDialog.getSaveFileName(self, "保存智能填报方案", "智能填报方案.smartplan", "智能填报方案 (*.smartplan)")
        if selected:
            path = Path(selected)
            if path.suffix.lower() != ".smartplan":
                path = path.with_suffix(".smartplan")
            save_plan(path, self.plan)

    def load_project(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "载入智能填报方案", "", "智能填报方案 (*.smartplan)")
        if not selected:
            return
        try:
            self.plan = load_plan(Path(selected))
        except Exception as exc:
            QMessageBox.critical(self, "载入失败", str(exc))
            return
        self._refresh_file_summary()
        self.refresh_pairs(); self.refresh_structures(); self.refresh_matches(False); self.refresh_rules()
        self.overwrite_formula_check.setChecked(
            self.plan.overwrite_target_formulas
        )
