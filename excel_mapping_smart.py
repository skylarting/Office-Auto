"""Beginner-friendly wizard for smart Excel report generation."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import json
import sys

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QIcon
from PySide6.QtWidgets import (
    QApplication, QComboBox, QDialog, QFileDialog, QFrame, QGridLayout,
    QHBoxLayout, QLabel, QLineEdit, QListWidget, QMainWindow, QMessageBox,
    QProgressBar, QPushButton, QScrollArea, QStackedWidget, QTableWidget,
    QTableWidgetItem, QVBoxLayout, QWidget,
)

from excel_mapping_qt import bundled_asset, install_tkinter_build_stub
from excel_mapper import WorkbookReader, best_name_match, workbook_files_in_folder
from smart_template import (
    SheetPair, SmartMatch, SmartTemplatePlan, execute_smart_plan, load_plan,
    match_pair, matches_to_mapping_rules, save_plan, suggest_sheet_pairs,
)


PLAN_SUFFIX = ".smartmap.json"


def display_name(value: str) -> str:
    return Path(value).name if value else ""


def find_plan_files(folder: Path) -> list[Path]:
    if not folder.is_dir():
        return []
    return sorted(folder.glob(f"*{PLAN_SUFFIX}"), key=lambda p: p.stat().st_mtime, reverse=True)


def candidate_matches(plan: SmartTemplatePlan) -> tuple[int, int, int, int]:
    auto = sum(
        bool(item.target_address) and item.status in ("自动匹配", "相似匹配")
        for item in plan.matches
    )
    pending = sum(item.status == "待确认" for item in plan.matches)
    missing = sum(not item.target_address for item in plan.matches)
    try:
        _rules, skipped = matches_to_mapping_rules(plan)
        protected = len(skipped)
    except Exception:
        protected = 0
    return int(auto), int(pending), int(missing), protected


def build_plan(source_files: list[Path], target_file: Path) -> SmartTemplatePlan:
    pairs = suggest_sheet_pairs(source_files, target_file)
    plan = SmartTemplatePlan(str(target_file.resolve()), [str(p.resolve()) for p in source_files], pairs)
    for pair in pairs:
        plan.matches.extend(match_pair(pair))
    return plan


def rebind_plan(saved: SmartTemplatePlan, source_folder: Path, target_file: Path) -> SmartTemplatePlan:
    """Use learned workbook/sheet names while reading current-month files."""
    files = workbook_files_in_folder(source_folder)
    if not files:
        raise ValueError("所选文件夹中没有找到 Excel 文件。")
    pairs: list[SheetPair] = []
    for old in saved.pairs:
        file_names = [path.name for path in files]
        chosen_name = best_name_match(Path(old.source_file).name, file_names) or ""
        chosen = next((path for path in files if path.name == chosen_name), None)
        if chosen is None:
            continue
        pairs.append(SheetPair(str(chosen.resolve()), old.source_sheet, str(target_file.resolve()), old.target_sheet, old.score))
    if not pairs:
        pairs = suggest_sheet_pairs(files, target_file)
    plan = SmartTemplatePlan(
        str(target_file.resolve()), [str(path.resolve()) for path in files], pairs,
        aliases=dict(saved.aliases), formula_rules=list(saved.formula_rules),
        block_rules=list(saved.block_rules),
        overwrite_target_formulas=saved.overwrite_target_formulas,
    )
    for pair in pairs:
        try:
            plan.matches.extend(match_pair(pair, plan.aliases))
        except Exception:
            continue
    return plan


class ChoiceCard(QFrame):
    def __init__(self, title: str, description: str, action: str, callback, primary: bool = False) -> None:
        super().__init__()
        self.setObjectName("choiceCard")
        self.setMaximumHeight(320)
        box = QVBoxLayout(self)
        box.setContentsMargins(24, 22, 24, 22)
        heading = QLabel(title)
        heading.setObjectName("cardTitle")
        box.addWidget(heading)
        note = QLabel(description)
        note.setWordWrap(True)
        note.setObjectName("secondaryText")
        box.addWidget(note)
        box.addStretch(1)
        button = QPushButton(action)
        if primary:
            button.setObjectName("primaryButton")
        button.clicked.connect(callback)
        box.addWidget(button, 0, Qt.AlignmentFlag.AlignRight)


class FileField(QWidget):
    def __init__(self, label: str, placeholder: str, folder: bool = False) -> None:
        super().__init__()
        self.folder = folder
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        title = QLabel(label)
        title.setObjectName("fieldLabel")
        layout.addWidget(title)
        row = QHBoxLayout()
        self.edit = QLineEdit()
        self.edit.setPlaceholderText(placeholder)
        self.edit.setReadOnly(True)
        row.addWidget(self.edit, 1)
        button = QPushButton("选择…")
        button.clicked.connect(self.choose)
        row.addWidget(button)
        layout.addLayout(row)

    def choose(self) -> None:
        if self.folder:
            value = QFileDialog.getExistingDirectory(self, "选择文件夹", self.edit.text())
        else:
            value, _ = QFileDialog.getOpenFileName(
                self, "选择 Excel 文件", self.edit.text(), "Excel 文件 (*.xls *.xlsx *.xlsm)"
            )
        if value:
            self.edit.setText(value)

    def path(self) -> Path | None:
        text = self.edit.text().strip()
        return Path(text) if text else None


class ReviewDialog(QDialog):
    """Ask one plain-language question while showing both cell neighborhoods."""
    def __init__(self, parent, match: SmartMatch, index: int, total: int) -> None:
        super().__init__(parent)
        self.match = match
        self.answer = "skip"
        self.setWindowTitle(f"确认数据对应关系（{index}/{total}）")
        self.resize(1050, 620)
        root = QVBoxLayout(self)
        root.setContentsMargins(20, 18, 20, 18)
        question = QLabel("请比较左右两个项目，它们是否表示同一项？")
        question.setObjectName("dialogTitle")
        root.addWidget(question)
        detail = QLabel(
            f"左侧：{match.row_text if hasattr(match, 'row_text') else ' / '.join(match.source_row_path)}  ·  "
            f"{' / '.join(match.source_column_path)}\n"
            f"右侧：{' / '.join(match.target_row_path) or '程序建议位置'}  ·  "
            f"{' / '.join(match.target_column_path)}"
        )
        detail.setObjectName("secondaryText")
        root.addWidget(detail)
        grids = QHBoxLayout()
        grids.addWidget(self._preview("来源数据", Path(match.source_file), match.source_sheet, match.source_address), 1)
        grids.addWidget(self._preview("要填写的报表", Path(match.target_file), match.target_sheet, match.target_address), 1)
        root.addLayout(grids, 1)
        actions = QHBoxLayout()
        skip = QPushButton("本次不填写")
        skip.clicked.connect(lambda: self.finish("skip"))
        actions.addWidget(skip)
        actions.addStretch(1)
        reject = QPushButton("不是，暂不对应")
        reject.clicked.connect(lambda: self.finish("reject"))
        actions.addWidget(reject)
        accept = QPushButton("是，确认对应")
        accept.setObjectName("primaryButton")
        accept.clicked.connect(lambda: self.finish("accept"))
        actions.addWidget(accept)
        root.addLayout(actions)

    def _preview(self, title: str, path: Path, sheet: str, address: str) -> QWidget:
        box = QFrame()
        box.setObjectName("previewBox")
        layout = QVBoxLayout(box)
        heading = QLabel(f"{title}\n{path.name} / {sheet}")
        heading.setObjectName("fieldLabel")
        layout.addWidget(heading)
        table = QTableWidget(7, 7)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.horizontalHeader().setVisible(False)
        table.verticalHeader().setVisible(False)
        reader = None
        try:
            reader = WorkbookReader(path)
            from excel_mapper import split_address
            row, col = split_address(address) if address else (0, 0)
            for rr in range(7):
                for cc in range(7):
                    real_row, real_col = max(0, row - 3) + rr, max(0, col - 3) + cc
                    letters = ""
                    n = real_col + 1
                    while n:
                        n, rem = divmod(n - 1, 26)
                        letters = chr(65 + rem) + letters
                    cell_address = f"{letters}{real_row + 1}"
                    value = reader.read(sheet, cell_address).value
                    item = QTableWidgetItem("" if value is None else str(value))
                    if cell_address == address:
                        item.setBackground(QColor("#FFF0EE"))
                        item.setForeground(QColor("#B42318"))
                    table.setItem(rr, cc, item)
            table.resizeColumnsToContents()
        except Exception as exc:
            table.setItem(0, 0, QTableWidgetItem(f"无法预览：{exc}"))
        finally:
            if reader:
                reader.close()
        layout.addWidget(table, 1)
        return box

    def finish(self, answer: str) -> None:
        self.answer = answer
        self.accept()


class SmartMappingWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("智能报表助手")
        self.resize(1120, 760)
        self.setMinimumSize(860, 620)
        self.plan: SmartTemplatePlan | None = None
        self.active_plan_file: Path | None = None
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)
        self.home = self._home_page()
        self.monthly = self._monthly_page()
        self.learn = self._learn_page()
        self.check = self._check_page()
        self.result = self._result_page()
        for page in (self.home, self.monthly, self.learn, self.check, self.result):
            self.stack.addWidget(page)
        self.show_home()

    def _shell(self, title: str, subtitle: str, back=True) -> tuple[QWidget, QVBoxLayout]:
        page = QWidget()
        root = QVBoxLayout(page)
        root.setContentsMargins(28, 22, 28, 24)
        top = QHBoxLayout()
        if back:
            button = QPushButton("← 返回首页")
            button.setObjectName("ghostButton")
            button.clicked.connect(self.show_home)
            top.addWidget(button)
        top.addStretch(1)
        root.addLayout(top)
        heading = QLabel(title)
        heading.setObjectName("pageTitle")
        root.addWidget(heading)
        note = QLabel(subtitle)
        note.setObjectName("pageSubtitle")
        note.setWordWrap(True)
        root.addWidget(note)
        return page, root

    def _home_page(self) -> QWidget:
        page, root = self._shell("智能报表助手", "只需选择文件，程序会自动查找数据并生成新的报表副本。", False)
        cards = QHBoxLayout()
        cards.addWidget(ChoiceCard("每月生成报表", "使用已经保存的方案，读取本月数据并填写报表。", "开始使用  →", self.open_monthly, True), 1)
        cards.addWidget(ChoiceCard("第一次建立方案", "提供一份已完成报表和对应的数据文件，让程序学习数据来源。", "开始建立  →", self.open_learn), 1)
        root.addLayout(cards, 1)
        bottom = QHBoxLayout()
        self.home_status = QLabel("尚未选择方案。")
        self.home_status.setObjectName("secondaryText")
        bottom.addWidget(self.home_status, 1)
        manage = QPushButton("打开方案文件夹")
        manage.clicked.connect(self.open_plan_folder)
        bottom.addWidget(manage)
        root.addLayout(bottom)
        return page

    def _monthly_page(self) -> QWidget:
        page, root = self._shell("每月生成报表", "选择已有方案、本月数据文件夹和空白报表模板。")
        steps = QLabel("1  选择文件     →     2  自动检查     →     3  回答问题     →     4  生成报表")
        steps.setObjectName("steps")
        root.addWidget(steps)
        form = QFrame(); form.setObjectName("formPanel")
        box = QVBoxLayout(form); box.setContentsMargins(24, 22, 24, 22); box.setSpacing(18)
        box.addWidget(QLabel("使用哪个报表方案？"))
        self.plan_combo = QComboBox(); box.addWidget(self.plan_combo)
        self.month_source = FileField("本月数据放在哪里？", "点击选择本月数据文件夹", True); box.addWidget(self.month_source)
        self.month_target = FileField("要填写哪张空白报表？", "点击选择空白报表模板"); box.addWidget(self.month_target)
        root.addWidget(form, 1)
        actions = QHBoxLayout(); actions.addStretch(1)
        go = QPushButton("开始检查"); go.setObjectName("primaryButton"); go.clicked.connect(self.run_monthly_check); actions.addWidget(go)
        root.addLayout(actions)
        return page

    def _learn_page(self) -> QWidget:
        page, root = self._shell("第一次建立方案", "程序会根据一份正确的历史报表，自动学习数据来自哪里。")
        form = QFrame(); form.setObjectName("formPanel")
        box = QVBoxLayout(form); box.setContentsMargins(24, 22, 24, 22); box.setSpacing(18)
        box.addWidget(QLabel("方案名称"))
        self.plan_name = QLineEdit(); self.plan_name.setPlaceholderText("例如：昆山支行月报"); box.addWidget(self.plan_name)
        self.learn_target = FileField("已经填写完成的报表", "选择一份内容正确的历史报表"); box.addWidget(self.learn_target)
        self.learn_source = FileField("这份报表使用的数据文件", "选择对应月份的数据文件夹", True); box.addWidget(self.learn_source)
        self.plan_folder = FileField("方案保存位置", "选择保存方案的文件夹", True); box.addWidget(self.plan_folder)
        root.addWidget(form, 1)
        actions = QHBoxLayout(); actions.addStretch(1)
        go = QPushButton("开始分析"); go.setObjectName("primaryButton"); go.clicked.connect(self.run_learning); actions.addWidget(go)
        root.addLayout(actions)
        return page

    def _check_page(self) -> QWidget:
        page, root = self._shell("检查完成", "程序已经完成自动识别。只有不能确定的项目才需要您确认。")
        self.check_title = QLabel("正在检查…"); self.check_title.setObjectName("resultTitle"); root.addWidget(self.check_title)
        metrics = QGridLayout()
        self.metric_labels = []
        for index, label in enumerate(("自动找到", "需要确认", "未找到来源", "保持不变")):
            panel = QFrame(); panel.setObjectName("metricPanel")
            v = QVBoxLayout(panel); number = QLabel("0"); number.setObjectName("metricNumber"); v.addWidget(number); v.addWidget(QLabel(label))
            metrics.addWidget(panel, 0, index); self.metric_labels.append(number)
        root.addLayout(metrics)
        self.check_detail = QListWidget(); root.addWidget(self.check_detail, 1)
        self.progress = QProgressBar(); self.progress.setValue(0); root.addWidget(self.progress)
        actions = QHBoxLayout()
        self.review_button = QPushButton("查看需要确认的项目")
        self.review_button.clicked.connect(self.review_pending); actions.addWidget(self.review_button)
        actions.addStretch(1)
        self.generate_button = QPushButton("生成报表副本"); self.generate_button.setObjectName("primaryButton"); self.generate_button.clicked.connect(self.generate_report); actions.addWidget(self.generate_button)
        root.addLayout(actions)
        return page

    def _result_page(self) -> QWidget:
        page, root = self._shell("报表已生成", "原始数据和报表模板都没有被修改。")
        self.result_title = QLabel("完成"); self.result_title.setObjectName("resultTitle"); root.addWidget(self.result_title)
        self.result_files = QListWidget(); root.addWidget(self.result_files, 1)
        actions = QHBoxLayout(); actions.addStretch(1)
        again = QPushButton("返回首页"); again.setObjectName("primaryButton"); again.clicked.connect(self.show_home); actions.addWidget(again)
        root.addLayout(actions)
        return page

    def show_home(self) -> None:
        self.stack.setCurrentWidget(self.home)

    def open_monthly(self) -> None:
        folder = self.plan_folder.path() or Path.cwd()
        self.plan_combo.clear()
        for path in find_plan_files(folder):
            self.plan_combo.addItem(path.stem.replace(".smartmap", ""), str(path))
        if not self.plan_combo.count():
            self.plan_combo.addItem("请先选择方案文件…", "")
        self.stack.setCurrentWidget(self.monthly)

    def open_learn(self) -> None:
        if not self.plan_folder.edit.text():
            self.plan_folder.edit.setText(str(Path.cwd()))
        self.stack.setCurrentWidget(self.learn)

    def open_plan_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "选择方案所在文件夹", str(Path.cwd()))
        if folder:
            self.plan_folder.edit.setText(folder)
            plans = find_plan_files(Path(folder))
            self.home_status.setText(f"找到 {len(plans)} 个已保存方案。")

    def _validate_files(self, source: FileField, target: FileField) -> tuple[Path, Path] | None:
        source_path, target_path = source.path(), target.path()
        if not source_path or not source_path.is_dir():
            QMessageBox.warning(self, "请选择数据文件夹", "请选择存放 Excel 数据文件的文件夹。")
            return None
        if not target_path or not target_path.is_file():
            QMessageBox.warning(self, "请选择报表", "请选择一份 Excel 报表文件。")
            return None
        return source_path, target_path

    def run_learning(self) -> None:
        checked = self._validate_files(self.learn_source, self.learn_target)
        if not checked:
            return
        name = self.plan_name.text().strip()
        if not name:
            QMessageBox.warning(self, "请填写方案名称", "请填写一个容易识别的方案名称，例如“昆山支行月报”。")
            return
        source_folder, target = checked
        files = workbook_files_in_folder(source_folder)
        if not files:
            QMessageBox.warning(self, "没有数据文件", "所选文件夹中没有找到 Excel 文件。")
            return
        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            self.plan = build_plan(files, target)
            folder = self.plan_folder.path() or source_folder
            folder.mkdir(parents=True, exist_ok=True)
            self.active_plan_file = folder / f"{name}{PLAN_SUFFIX}"
            save_plan(self.active_plan_file, self.plan)
            self.show_check("方案学习完成")
        except Exception as exc:
            QMessageBox.critical(self, "无法建立方案", str(exc))
        finally:
            QApplication.restoreOverrideCursor()

    def run_monthly_check(self) -> None:
        checked = self._validate_files(self.month_source, self.month_target)
        if not checked:
            return
        plan_path = str(self.plan_combo.currentData() or "")
        if not plan_path:
            plan_path, _ = QFileDialog.getOpenFileName(self, "选择报表方案", str(Path.cwd()), f"智能报表方案 (*{PLAN_SUFFIX})")
        if not plan_path:
            return
        try:
            QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            saved = load_plan(Path(plan_path))
            self.active_plan_file = Path(plan_path)
            self.plan = rebind_plan(saved, checked[0], checked[1])
            self.show_check("本月报表检查完成")
        except Exception as exc:
            QMessageBox.critical(self, "无法检查报表", str(exc))
        finally:
            QApplication.restoreOverrideCursor()

    def show_check(self, title: str) -> None:
        assert self.plan is not None
        auto, pending, missing, formulas = candidate_matches(self.plan)
        self.check_title.setText(title)
        for widget, value in zip(self.metric_labels, (auto, pending, missing, formulas)):
            widget.setText(str(value))
        self.check_detail.clear()
        self.check_detail.addItem(f"找到 {len(self.plan.source_files)} 个数据文件")
        self.check_detail.addItem(f"识别到 {len(self.plan.pairs)} 组工作表对应关系")
        self.check_detail.addItem(f"预计填写 {auto + pending} 个单元格")
        self.check_detail.addItem("目标报表中的已有公式默认保持不变")
        self.review_button.setEnabled(pending > 0)
        self.review_button.setText(f"查看需要确认的项目（{pending}）")
        self.generate_button.setEnabled(pending == 0 and auto > 0)
        self.progress.setValue(0)
        self.stack.setCurrentWidget(self.check)

    def review_pending(self) -> None:
        if not self.plan:
            return
        pending = [m for m in self.plan.matches if m.status == "待确认"]
        for index, match in enumerate(pending, 1):
            dialog = ReviewDialog(self, match, index, len(pending))
            dialog.exec()
            if dialog.answer == "accept":
                match.enabled = bool(match.target_address); match.status = "人工确认"
            else:
                match.enabled = False; match.status = "本次不填写"
        if self.active_plan_file:
            save_plan(self.active_plan_file, self.plan)
        self.show_check("确认完成")

    def generate_report(self) -> None:
        if not self.plan:
            return
        base = Path(self.plan.target_file).parent / "映射结果"
        folder = QFileDialog.getExistingDirectory(self, "选择报表保存位置", str(base))
        output = Path(folder) if folder else base
        try:
            output.mkdir(parents=True, exist_ok=True)
            self.progress.setValue(10)
            def progress(current, total, message):
                self.progress.setValue(int(current / max(total, 1) * 90) + 10)
                QApplication.processEvents()
            outputs, skipped = execute_smart_plan(self.plan, output, progress)
            self.progress.setValue(100)
            self.result_files.clear()
            for path in outputs:
                self.result_files.addItem(str(path))
            if skipped:
                self.result_files.addItem(f"另有 {len(skipped)} 项因公式或保护规则保持不变。")
            self.result_title.setText(f"已生成 {len(outputs)} 份报表副本")
            self.stack.setCurrentWidget(self.result)
        except Exception as exc:
            QMessageBox.critical(self, "生成报表失败", str(exc))
        finally:
            self.progress.setValue(0)


def application_style() -> str:
    return """
    QMainWindow, QWidget { background: #F5F7FA; color: #1F2937; font-family: "Microsoft YaHei UI"; font-size: 14px; }
    QLabel#pageTitle { font-size: 28px; font-weight: 600; color: #172B4D; }
    QLabel#pageSubtitle, QLabel#secondaryText { color: #667085; }
    QLabel#cardTitle, QLabel#dialogTitle, QLabel#resultTitle { font-size: 20px; font-weight: 600; color: #172B4D; }
    QLabel#fieldLabel { font-weight: 600; color: #344054; }
    QLabel#steps { padding: 12px 16px; background: #EAF2FF; color: #175CD3; border-radius: 8px; }
    QFrame#choiceCard, QFrame#formPanel, QFrame#previewBox, QFrame#metricPanel { background: white; border: 1px solid #DDE3EC; border-radius: 10px; }
    QFrame#choiceCard:hover { border-color: #84ADFF; }
    QLineEdit, QComboBox, QListWidget, QTableWidget { background: white; border: 1px solid #CBD5E1; border-radius: 6px; padding: 7px; selection-background-color: #DCEAFF; selection-color: #172B4D; }
    QPushButton { background: white; border: 1px solid #B8C2D1; border-radius: 6px; padding: 9px 16px; }
    QPushButton:hover { background: #F0F5FF; border-color: #7EA6E0; }
    QPushButton#primaryButton { background: #2563EB; color: white; border-color: #2563EB; font-weight: 600; }
    QPushButton#primaryButton:hover { background: #1D4ED8; }
    QPushButton#ghostButton { border: none; background: transparent; color: #475467; padding-left: 0; }
    QLabel#metricNumber { font-size: 26px; font-weight: 600; color: #175CD3; }
    QProgressBar { background: #E5EAF1; border: none; border-radius: 5px; text-align: center; min-height: 10px; }
    QProgressBar::chunk { background: #2563EB; border-radius: 5px; }
    """


def main() -> None:
    install_tkinter_build_stub()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(application_style())
    icon = bundled_asset("assets/excel-mapper.ico")
    if icon.exists():
        app.setWindowIcon(QIcon(str(icon)))
    window = SmartMappingWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
