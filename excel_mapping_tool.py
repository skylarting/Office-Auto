"""Single-page Excel mapping table application."""

from pathlib import Path
from tkinter import (
    BOTH,
    LEFT,
    RIGHT,
    W,
    X,
    Menu,
    StringVar,
    Tk,
    messagebox,
    ttk,
)

from excel_mapper import (
    MapperApp,
    MappingRule,
    format_cell_addresses,
    infer_mapping_mode,
    parse_cell_addresses,
    resolve_project_path,
)


class ExcelMappingToolApp(MapperApp):
    """A focused mapping editor without the legacy manual/TXT workflows."""

    def _configure_styles(self) -> None:
        super()._configure_styles()
        style = ttk.Style(self.root)
        style.configure(
            "Mapping.Treeview",
            rowheight=30,
            font=("Microsoft YaHei UI", 10),
        )
        style.configure(
            "Mapping.Treeview.Heading",
            font=("Microsoft YaHei UI", 10, "bold"),
        )
        style.map(
            "Mapping.Treeview",
            background=[("selected", "#CFE8FA")],
            foreground=[("selected", "#102A43")],
        )
        style.configure(
            "Title.TLabel",
            font=("Microsoft YaHei UI", 15, "bold"),
        )
        style.configure(
            "Status.TLabel",
            font=("Microsoft YaHei UI", 9),
            foreground="#4F5B66",
        )

    def _build_ui(self) -> None:
        self.root.title("Excel 单元格映射工具")
        self.root.geometry("1180x760")
        self.root.minsize(860, 560)
        self._active_workflow_tab = 1

        container = ttk.Frame(self.root, padding=16)
        container.pack(fill=BOTH, expand=True)

        ttk.Label(
            container,
            text="Excel 单元格映射工具",
            style="Title.TLabel",
        ).pack(anchor=W, pady=(0, 12))

        command_bar = ttk.Frame(container)
        command_bar.pack(fill=X, pady=(0, 10))
        for text, command in (
            ("导入映射表", self._load_excel_scheme),
            ("导出映射表", self._save_excel_scheme),
        ):
            ttk.Button(
                command_bar,
                text=text,
                command=command,
                style="Toolbar.TButton",
            ).pack(side=LEFT, padx=(0, 8))
        ttk.Button(
            command_bar,
            text="清空所有映射",
            command=self._clear_all_workflow,
            style="Toolbar.TButton",
        ).pack(side=RIGHT)

        default_row = ttk.Frame(container)
        default_row.pack(fill=X, pady=(0, 12))
        ttk.Label(default_row, text="默认文件夹：").pack(side=LEFT)
        self.default_folder_display = StringVar(
            value="点击此处选择默认文件夹"
        )
        self.scheme_base_folder.trace_add(
            "write",
            lambda *_args: self.default_folder_display.set(
                self.scheme_base_folder.get().strip()
                or "点击此处选择默认文件夹"
            ),
        )
        self.scheme_base_entry = ttk.Entry(
            default_row,
            textvariable=self.default_folder_display,
            state="readonly",
            cursor="hand2",
        )
        self.scheme_base_entry.pack(side=LEFT, fill=X, expand=True)
        self.scheme_base_entry.bind(
            "<Button-1>",
            lambda _event: self._browse_scheme_base(),
        )
        self.scheme_base_entry.bind(
            "<Return>",
            lambda _event: self._browse_scheme_base(),
        )

        table_surface = ttk.Frame(container)
        table_surface.pack(fill=BOTH, expand=True)
        table_surface.rowconfigure(0, weight=1)
        table_surface.columnconfigure(0, weight=1)

        self.scheme_tree = ttk.Treeview(
            table_surface,
            columns=("group", "kind", "workbook", "sheet", "cells"),
            show="headings",
            style="Mapping.Treeview",
            selectmode="extended",
        )
        for column, title, width, stretch in (
            ("group", "映射组", 70, False),
            ("kind", "类型", 70, False),
            ("workbook", "工作簿", 430, True),
            ("sheet", "工作表", 230, True),
            ("cells", "单元格", 260, True),
        ):
            self.scheme_tree.heading(column, text=title, anchor=W)
            self.scheme_tree.column(
                column,
                width=width,
                minwidth=60,
                anchor=W,
                stretch=stretch,
            )
        self.scheme_tree.tag_configure("source", background="#FBFAF7")
        self.scheme_tree.tag_configure("target", background="#F2F0EB")
        self.scheme_tree.grid(row=0, column=0, sticky="nsew")
        vertical = ttk.Scrollbar(
            table_surface,
            orient="vertical",
            command=self.scheme_tree.yview,
        )
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal = ttk.Scrollbar(
            table_surface,
            orient="horizontal",
            command=self.scheme_tree.xview,
        )
        horizontal.grid(row=1, column=0, sticky="ew")
        self.scheme_tree.configure(
            yscrollcommand=vertical.set,
            xscrollcommand=horizontal.set,
        )
        self.scheme_tree.bind("<Double-1>", self._begin_scheme_cell_edit)
        self.scheme_tree.bind(
            "<ButtonPress-1>",
            self._scheme_row_button_press,
            add="+",
        )
        self.scheme_tree.bind(
            "<B1-Motion>",
            self._scheme_row_drag_motion,
            add="+",
        )
        self.scheme_tree.bind(
            "<ButtonRelease-1>",
            self._scheme_row_button_release,
        )
        self.scheme_tree.bind(
            "<Return>",
            self._begin_selected_scheme_cell_edit,
        )
        self.scheme_tree.bind(
            "<Delete>",
            lambda _event: self._delete_scheme_rule(),
        )
        self.scheme_tree.bind("<Button-3>", self._show_scheme_context_menu)
        self.scheme_tree.bind("<Control-c>", self._copy_selected_row_contents)
        self.scheme_tree.bind("<Control-C>", self._copy_selected_row_contents)
        self.scheme_tree.bind("<Control-v>", self._paste_row_contents)
        self.scheme_tree.bind("<Control-V>", self._paste_row_contents)
        self._row_drag = None
        self.scheme_context_menu = Menu(self.root, tearoff=False)
        self.scheme_context_menu.add_command(
            label="复制选中行内容",
            command=self._copy_selected_row_contents,
        )
        self.scheme_context_menu.add_command(
            label="粘贴行内容",
            command=self._paste_row_contents,
        )
        self.scheme_context_menu.add_separator()
        self.scheme_context_menu.add_command(
            label="在上方插入映射组",
            command=lambda: self._insert_scheme_rule("before"),
        )
        self.scheme_context_menu.add_command(
            label="在下方插入映射组",
            command=lambda: self._insert_scheme_rule("after"),
        )
        self.scheme_context_menu.add_command(
            label="复制本组",
            command=self._duplicate_scheme_rule,
        )
        self.scheme_context_menu.add_separator()
        self.scheme_context_menu.add_command(
            label="删除本组",
            command=self._delete_scheme_rule,
        )

        ttk.Label(
            container,
            text=(
                "单击选择，双击浏览或输入；按住“类型”列可拖动行内容，"
                "Ctrl 可多选；右键可复制、粘贴或插入/删除映射组。"
            ),
            style="Status.TLabel",
        ).pack(anchor=W, pady=(8, 4))
        ttk.Label(
            container,
            textvariable=self.scheme_status,
            style="Status.TLabel",
        ).pack(anchor=W, pady=(0, 10))

        ttk.Separator(container).pack(fill=X, pady=(0, 10))

        output_row = ttk.Frame(container)
        output_row.pack(fill=X)
        ttk.Label(output_row, text="输出文件夹：").pack(side=LEFT)
        self.output_folder_display = StringVar(
            value="点击此处选择输出文件夹"
        )
        self.output_folder.trace_add(
            "write",
            lambda *_args: self.output_folder_display.set(
                self.output_folder.get().strip()
                or "点击此处选择输出文件夹"
            ),
        )
        self.output_entry = ttk.Entry(
            output_row,
            textvariable=self.output_folder_display,
            state="readonly",
            cursor="hand2",
        )
        self.output_entry.pack(side=LEFT, fill=X, expand=True)
        self.output_entry.bind(
            "<Button-1>",
            lambda _event: self._browse_output(),
        )
        self.output_entry.bind(
            "<Return>",
            lambda _event: self._browse_output(),
        )

        status_row = ttk.Frame(container)
        status_row.pack(fill=X, pady=(10, 4))
        ttk.Label(
            status_row,
            textvariable=self.status,
            style="Status.TLabel",
        ).pack(side=LEFT)
        ttk.Label(
            status_row,
            textvariable=self.progress_text,
            style="Status.TLabel",
        ).pack(side=RIGHT)

        self.progress = ttk.Progressbar(
            container,
            orient="horizontal",
            mode="determinate",
        )
        self.progress.pack(fill=X, pady=(0, 12))

        actions = ttk.Frame(container)
        actions.pack(fill=X)
        ttk.Button(
            actions,
            text="展开预览",
            command=self._preview_expanded,
        ).pack(side=LEFT)
        ttk.Button(
            actions,
            text="预检查",
            command=self._precheck,
        ).pack(side=LEFT, padx=(8, 0))
        ttk.Button(
            actions,
            text="退出",
            command=self.root.destroy,
        ).pack(side=RIGHT)
        ttk.Button(
            actions,
            text="开始映射",
            command=self._run,
            style="Primary.TButton",
        ).pack(side=RIGHT, padx=(0, 10))

        self.status.set("尚未检查。")
        self._refresh_scheme_tree()

    def _clear_all_workflow(self) -> None:
        if not self.scheme_rules:
            self.scheme_status.set("当前没有需要清空的映射。")
            return
        if not messagebox.askyesno(
            "清空所有映射",
            "确定清空全部映射关系吗？\n\n默认文件夹和输出文件夹会保留。",
            parent=self.root,
        ):
            return
        self.scheme_rules.clear()
        self._refresh_scheme_tree()
        self.scheme_status.set("映射已清空，可以从底部空白组重新填写。")
        self.status.set("尚未检查。")

    @staticmethod
    def _actual_row_index(iid: str) -> int | None:
        try:
            _, index_text, kind = iid.split(":")
            index = int(index_text)
            return index * 2 + (1 if kind == "target" else 0)
        except (ValueError, IndexError):
            return None

    def _selected_scheme_rule_index(self) -> int | None:
        iid = self.scheme_tree.focus()
        if not iid:
            selected = self.scheme_tree.selection()
            iid = selected[0] if selected else ""
        try:
            return int(iid.split(":")[1])
        except (IndexError, ValueError):
            return None

    def _scheme_row_button_press(self, event):
        iid = self.scheme_tree.identify_row(event.y)
        column = self.scheme_tree.identify_column(event.x)
        if column == "#1" and iid:
            try:
                _, index_text, _kind = iid.split(":")
                pair = (
                    f"scheme:{index_text}:source",
                    f"scheme:{index_text}:target",
                )
            except ValueError:
                pair = ()
            existing = set(self.scheme_tree.selection())
            if event.state & 0x0004:
                for row in pair:
                    if row in existing:
                        self.scheme_tree.selection_remove(row)
                    elif self.scheme_tree.exists(row):
                        self.scheme_tree.selection_add(row)
            else:
                self.scheme_tree.selection_set(pair)
            self.scheme_tree.focus(iid)
            return "break"
        if column == "#2" and iid and not iid.startswith("scheme:new:"):
            self._row_drag = {
                "iid": iid,
                "start_y": event.y,
                "active": False,
            }

    def _scheme_row_drag_motion(self, event):
        drag = self._row_drag
        if not drag:
            return
        if not drag["active"] and abs(event.y - drag["start_y"]) < 6:
            return
        drag["active"] = True
        self.scheme_tree.configure(cursor="fleur")
        destination = self.scheme_tree.identify_row(event.y)
        if destination:
            self.scheme_tree.see(destination)

    def _scheme_row_button_release(self, event):
        drag = self._row_drag
        self._row_drag = None
        self.scheme_tree.configure(cursor="")
        if drag and drag["active"]:
            destination = self.scheme_tree.identify_row(event.y)
            self._move_selected_row_contents(drag["iid"], destination)
            return "break"
        return self._schedule_scheme_single_click(event)

    def _flatten_row_contents(self) -> list[tuple[str, str, list[str]]]:
        rows = []
        for rule in self.scheme_rules:
            rows.append(
                (rule.source_file, rule.source_sheet, list(rule.source_cells))
            )
            rows.append(
                (rule.target_file, rule.target_sheet, list(rule.target_cells))
            )
        return rows

    def _replace_rules_from_row_contents(
        self,
        rows: list[tuple[str, str, list[str]]],
    ) -> None:
        if len(rows) % 2:
            rows.append(("", "", ["A1"]))
        rebuilt = []
        for offset in range(0, len(rows), 2):
            source_file, source_sheet, source_cells = rows[offset]
            target_file, target_sheet, target_cells = rows[offset + 1]
            source_cells = source_cells or ["A1"]
            target_cells = target_cells or list(source_cells)
            rebuilt.append(
                MappingRule(
                    source_file,
                    source_sheet,
                    list(source_cells),
                    target_file,
                    target_sheet,
                    list(target_cells),
                    infer_mapping_mode(source_cells, target_cells),
                )
            )
        self.scheme_rules[:] = rebuilt

    def _move_selected_row_contents(
        self,
        pressed_iid: str,
        destination_iid: str,
    ) -> None:
        pressed = self._actual_row_index(pressed_iid)
        if pressed is None:
            return
        selected = {
            index
            for iid in self.scheme_tree.selection()
            if (index := self._actual_row_index(iid)) is not None
        }
        if pressed not in selected:
            selected = {pressed}
        selected = sorted(selected)
        rows = self._flatten_row_contents()
        destination = self._actual_row_index(destination_iid)
        if destination is None:
            destination = len(rows)
        moving = [rows[index] for index in selected]
        remaining = [
            row for index, row in enumerate(rows) if index not in selected
        ]
        insertion = destination - sum(index < destination for index in selected)
        insertion = max(0, min(insertion, len(remaining)))
        new_rows = remaining[:insertion] + moving + remaining[insertion:]
        if new_rows == rows:
            return
        self._replace_rules_from_row_contents(new_rows)
        self._refresh_scheme_tree()
        moved_iids = []
        for index in range(insertion, insertion + len(moving)):
            kind = "source" if index % 2 == 0 else "target"
            moved_iids.append(f"scheme:{index // 2}:{kind}")
        self.scheme_tree.selection_set(moved_iids)
        if moved_iids:
            self.scheme_tree.focus(moved_iids[0])
            self.scheme_tree.see(moved_iids[0])
        self.scheme_status.set(
            "已移动行内容；映射组和来源/目标已按新位置自动重新编号。"
        )

    def _copy_selected_row_contents(self, _event=None):
        selected = [
            iid
            for iid in self.scheme_tree.get_children()
            if iid in self.scheme_tree.selection()
            and self._actual_row_index(iid) is not None
        ]
        if not selected:
            focused = self.scheme_tree.focus()
            selected = [focused] if self._actual_row_index(focused) is not None else []
        if not selected:
            return "break"
        lines = []
        for iid in selected:
            values = self.scheme_tree.item(iid, "values")
            lines.append("\t".join(str(value) for value in values[2:5]))
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(lines))
        self.scheme_status.set(f"已复制 {len(lines)} 行的后三列内容。")
        return "break"

    def _paste_row_contents(self, _event=None):
        focused = self.scheme_tree.focus()
        start = self._actual_row_index(focused)
        if start is None:
            return "break"
        try:
            text = self.root.clipboard_get()
            incoming = []
            base = self._default_scheme_base_folder() or Path.cwd()
            for line in text.splitlines():
                fields = line.split("\t")
                if len(fields) != 3:
                    raise ValueError("粘贴内容必须是工作簿、工作表、单元格三列。")
                incoming.append(
                    (
                        str(resolve_project_path(fields[0].strip(), base))
                        if fields[0].strip()
                        else "",
                        fields[1].strip(),
                        parse_cell_addresses(fields[2].strip())
                        if fields[2].strip() not in ("", "同位置")
                        else [],
                    )
                )
            rows = self._flatten_row_contents()
            while len(rows) < start + len(incoming):
                rows.extend([("", "", ["A1"]), ("", "", ["A1"])])
            for offset, content in enumerate(incoming):
                rows[start + offset] = content
            self._replace_rules_from_row_contents(rows)
            self._refresh_scheme_tree()
            self.scheme_status.set(f"已粘贴 {len(incoming)} 行的后三列内容。")
        except Exception as exc:
            messagebox.showerror("无法粘贴", str(exc), parent=self.root)
        return "break"

    def _show_scheme_context_menu(self, event) -> str:
        iid = self.scheme_tree.identify_row(event.y)
        if not iid:
            return "break"
        if iid not in self.scheme_tree.selection():
            self.scheme_tree.selection_set(iid)
        self.scheme_tree.focus(iid)
        is_blank = iid.startswith("scheme:new:")
        state = "disabled" if is_blank else "normal"
        self.scheme_context_menu.entryconfigure(0, state=state)
        self.scheme_context_menu.entryconfigure(1, state=state)
        self.scheme_context_menu.entryconfigure(5, state=state)
        self.scheme_context_menu.entryconfigure(7, state=state)
        try:
            self.scheme_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.scheme_context_menu.grab_release()
        return "break"


def main() -> None:
    root = Tk()
    ExcelMappingToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
