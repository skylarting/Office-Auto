"""Single-page Excel mapping table application."""

from tkinter import BOTH, LEFT, RIGHT, W, X, Menu, Tk, messagebox, ttk

from excel_mapper import MapperApp


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
            background=[("selected", "#D9EAF7")],
            foreground=[("selected", "#111111")],
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
        self.scheme_base_entry = ttk.Entry(
            default_row,
            textvariable=self.scheme_base_folder,
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
        self.scheme_tree.tag_configure("source", background="#EAF3FB")
        self.scheme_tree.tag_configure("target", background="#EDF7ED")
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
            "<ButtonRelease-1>",
            self._schedule_scheme_single_click,
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
        self.scheme_context_menu = Menu(self.root, tearoff=False)
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
                "填写底部空白组即可新增；单击选择，双击工作簿浏览文件，"
                "双击工作表或单元格可输入；Delete/右键管理整组。"
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
        self.output_entry = ttk.Entry(
            output_row,
            textvariable=self.output_folder,
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


def main() -> None:
    root = Tk()
    ExcelMappingToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
