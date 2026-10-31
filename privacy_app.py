import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import os
import subprocess
import sys
import threading
from typing import Any

from privacy_handler.gui_controller import (
    GUIWorkflowController,
    DEFAULT_SAFE_DIR,
    DEFAULT_RESTORED_DIR,
)
from privacy_handler.table_scope import validate_composite_key, extract_table_scope, iter_table_scope_rows
from privacy_handler.identity import compute_run_fingerprint
import openpyxl


def open_folder(folder_path: Path | str) -> None:
    folder = Path(folder_path)
    folder.mkdir(parents=True, exist_ok=True)
    if sys.platform == "darwin":
        subprocess.run(["open", str(folder)], check=False)
    elif os.name == "nt":
        os.startfile(str(folder))
    else:
        subprocess.run(["xdg-open", str(folder)], check=False)


class RecordKeyDialog(tk.Toplevel):
    """Modal dialog to select and validate record identification columns for active sheet."""

    def __init__(self, parent: "PrivacyApp", sheet_name: str):
        super().__init__(parent)
        self.title(f"设置记录识别列 — {sheet_name}")
        self.geometry("540x520")
        self.minsize(460, 400)
        self.transient(parent)
        self.grab_set()

        self.app = parent
        self.sheet_name = sheet_name
        self.controller = parent.controller
        self.cfg = self.controller.sheet_configs.get(sheet_name)

        self.col_vars: dict[str, tk.BooleanVar] = {}
        self.validation_result: dict[str, Any] = {}

        self._build_ui()
        self._validate_current_selection()

    def _build_ui(self):
        # Header description
        top_frame = ttk.Frame(self, padding=(16, 12, 16, 8))
        top_frame.pack(fill="x")

        ttk.Label(
            top_frame,
            text=f"工作表：{self.sheet_name}",
            font=("", 12, "bold"),
            foreground="#2c3e50",
        ).pack(anchor="w")

        ttk.Label(
            top_frame,
            text="💡 选最少的字段组合，只要能够唯一识别每条记录即可（如“员工工号”或“员工工号 + 年月”）。不要为了保险多选很多字段。",
            font=("", 9),
            foreground="#555555",
            wraplength=500,
        ).pack(anchor="w", pady=(4, 0))

        # Scrollable column list
        list_container = ttk.Frame(self, padding=(16, 4, 16, 4))
        list_container.pack(fill="both", expand=True)

        canvas = tk.Canvas(list_container, highlightthickness=0, bg="#ffffff")
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)

        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        cwin = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(cwin, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        if self.cfg:
            for col in self.cfg.columns:
                var = tk.BooleanVar(value=col.is_record_key)
                self.col_vars[col.column_letter] = var

                row = ttk.Frame(inner, padding=(4, 2))
                row.pack(fill="x")
                chk = ttk.Checkbutton(
                    row,
                    text=col.display_name,
                    variable=var,
                    command=self._on_check_toggle,
                )
                chk.pack(side="left")

        # Validation status box
        status_box = ttk.LabelFrame(self, text="【组合验证结果】", padding=10)
        status_box.pack(fill="x", padx=16, pady=6)

        self.dialog_status_lbl = ttk.Label(
            status_box,
            text="正在校验...",
            font=("", 10),
            wraplength=490,
        )
        self.dialog_status_lbl.pack(anchor="w")

        # Buttons footer
        btn_bar = ttk.Frame(self, padding=(16, 8, 16, 12))
        btn_bar.pack(fill="x")

        ttk.Button(btn_bar, text="清空所有选择", command=self._clear_all).pack(side="left")
        ttk.Button(btn_bar, text="取消", command=self.destroy).pack(side="right", padx=(6, 0))
        self.btn_apply = ttk.Button(btn_bar, text="保存并应用", command=self._apply_and_close)
        self.btn_apply.pack(side="right")

    def _on_check_toggle(self):
        self._validate_current_selection()

    def _clear_all(self):
        for var in self.col_vars.values():
            var.set(False)
        self._validate_current_selection()

    def _get_selected_letters(self) -> list[str]:
        return [l for l, v in self.col_vars.items() if v.get()]

    def _validate_current_selection(self):
        if not self.cfg:
            return
        selected_letters = self._get_selected_letters()
        if not selected_letters:
            self.dialog_status_lbl.config(
                text="⚠ 未选择任何字段：仍可脱敏，但 AI 修改或调换顺序后可能无法自动安全恢复。",
                foreground="#d35400",
            )
            return

        selected_headers: list[str] = []
        for l in selected_letters:
            c = self.cfg.get_column(l)
            if c:
                selected_headers.append(c.header)

        # Headless validation via openpyxl
        try:
            wb = openpyxl.load_workbook(self.controller.source_path, data_only=False)
            ws = wb[self.sheet_name]
            scope = extract_table_scope(ws, header_row=self.cfg.header_row)
            rows = iter_table_scope_rows(ws, scope)
            wb.close()

            res = validate_composite_key(rows, selected_headers)
            key_desc = " + ".join(selected_headers)
            if res["is_valid"]:
                if len(selected_headers) > 1:
                    msg = f"✓ 组合后可唯一识别每条记录：当前组合可以唯一识别每条记录。不要继续增加字段。[{key_desc}]，恢复时可自动匹配"
                else:
                    msg = f"✓ 可唯一识别每条记录：已经可以唯一识别每条记录，无需再增加其他字段。[{key_desc}]，恢复时可自动匹配"
                self.dialog_status_lbl.config(text=msg, foreground="#27ae60")
            else:
                msg = f"⚠ 不能可靠识别每条记录：存在重复或空值，建议再增加一个能够区分记录的字段，例如“年月”：[{key_desc}]"
                self.dialog_status_lbl.config(text=msg, foreground="#c0392b")
        except Exception as exc:
            self.dialog_status_lbl.config(text=f"校验出错：{exc}", foreground="#c0392b")

    def _apply_and_close(self):
        if not self.cfg:
            self.destroy()
            return
        selected_letters = set(self._get_selected_letters())
        for col in self.cfg.columns:
            col.is_record_key = (col.column_letter in selected_letters)
        self.cfg.is_confirmed = False
        self.controller.validate_sheet_record_identity(self.sheet_name)
        self.app._render_column_list()
        self.app._render_sheet_list()
        self.destroy()


class SharedGroupDialog(tk.Toplevel):
    """Modal dialog to create and manage cross-sheet shared numeric analysis groups."""

    def __init__(self, parent: "PrivacyApp"):
        super().__init__(parent)
        self.title("设置跨表共享组")
        self.geometry("640x540")
        self.minsize(540, 420)
        self.transient(parent)
        self.grab_set()

        self.app = parent
        self.controller = parent.controller
        self.group_vars: dict[tuple[str, str], tk.StringVar] = {}

        self._build_ui()

    def _build_ui(self):
        # Description header
        top_frame = ttk.Frame(self, padding=(16, 12, 16, 8))
        top_frame.pack(fill="x")

        ttk.Label(
            top_frame,
            text="跨工作表数值共享组",
            font=("", 12, "bold"),
            foreground="#2c3e50",
        ).pack(anchor="w")

        ttk.Label(
            top_frame,
            text="💡 放入同一共享组的数值字段在跨表时将使用相同的缩放比例脱敏，适合跨表比较（如员工基本工资与奖金明细）。默认各表独立使用本地 A 组，互不影响。",
            font=("", 9),
            foreground="#555555",
            wraplength=600,
        ).pack(anchor="w", pady=(4, 0))

        # Main table of numeric columns across all selected sheets
        table_container = ttk.Frame(self, padding=(16, 4, 16, 4))
        table_container.pack(fill="both", expand=True)

        canvas = tk.Canvas(table_container, highlightthickness=0, bg="#ffffff")
        scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)

        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        cwin = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(cwin, width=e.width))
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Table header
        hdr_row = ttk.Frame(inner, padding=(4, 4))
        hdr_row.pack(fill="x")
        ttk.Label(hdr_row, text="工作表", width=16, font=("", 10, "bold")).pack(side="left")
        ttk.Label(hdr_row, text="字段名称", width=22, font=("", 10, "bold")).pack(side="left")
        ttk.Label(hdr_row, text="所属组别", width=18, font=("", 10, "bold")).pack(side="left")

        ttk.Separator(inner, orient="horizontal").pack(fill="x", pady=2)

        has_numeric = False
        group_options = ["本地组（不共享）", "共享组 1", "共享组 2", "共享组 3"]

        for sname, scfg in self.controller.sheet_configs.items():
            if not scfg.is_selected:
                continue
            for col in scfg.columns:
                if col.protection_type == "numeric":
                    has_numeric = True
                    row = ttk.Frame(inner, padding=(4, 3))
                    row.pack(fill="x")

                    ttk.Label(row, text=sname, width=16).pack(side="left")
                    ttk.Label(row, text=col.display_name, width=22).pack(side="left")

                    cur_val = "本地组（不共享）"
                    if col.shared_group_id:
                        cur_val = col.shared_group_id

                    var = tk.StringVar(value=cur_val)
                    self.group_vars[(sname, col.column_letter)] = var
                    combo = ttk.Combobox(row, textvariable=var, values=group_options, state="readonly", width=16)
                    combo.pack(side="left")

        if not has_numeric:
            empty_lbl = ttk.Label(inner, text="当前所有选中的工作表中均未设置“数值保护”字段。", font=("", 10), foreground="#7f8c8d")
            empty_lbl.pack(pady=20)

        # Buttons footer
        btn_bar = ttk.Frame(self, padding=(16, 8, 16, 12))
        btn_bar.pack(fill="x")

        ttk.Button(btn_bar, text="全部恢复为本地组", command=self._reset_all_local).pack(side="left")
        ttk.Button(btn_bar, text="取消", command=self.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(btn_bar, text="保存并应用", command=self._apply_and_close).pack(side="right")

    def _reset_all_local(self):
        for var in self.group_vars.values():
            var.set("本地组（不共享）")

    def _apply_and_close(self):
        # Update columns in controller
        affected_sheets: set[str] = set()
        for (sname, col_letter), var in self.group_vars.items():
            val = var.get()
            scfg = self.controller.sheet_configs.get(sname)
            if scfg:
                col = scfg.get_column(col_letter)
                if col:
                    if val == "本地组（不共享）":
                        if col.shared_group_id is not None:
                            col.shared_group_id = None
                            col.local_group = "A"
                            affected_sheets.add(sname)
                    else:
                        if col.shared_group_id != val:
                            col.shared_group_id = val
                            affected_sheets.add(sname)

        for sname in affected_sheets:
            if sname in self.controller.sheet_configs:
                self.controller.sheet_configs[sname].is_confirmed = False

        self.app._render_column_list()
        self.app._render_sheet_list()
        self.destroy()


class PromptPreviewDialog(tk.Toplevel):
    """Modal dialog to view full AI analysis prompt instructions."""

    def __init__(self, parent: "PrivacyApp", prompt_text: str):
        super().__init__(parent)
        self.title("可复制给 AI 的安全处理说明")
        self.geometry("720x540")
        self.minsize(580, 420)
        self.transient(parent)
        self.grab_set()

        self.prompt_content = prompt_text

        # Header description
        top_frame = ttk.Frame(self, padding=(16, 12, 16, 8))
        top_frame.pack(fill="x")
        ttk.Label(
            top_frame,
            text="AI 处理说明（完整内容）",
            font=("", 12, "bold"),
            foreground="#2c3e50",
        ).pack(anchor="w")
        ttk.Label(
            top_frame,
            text="可将以下说明完整复制并提供给 AI 工具（包含字段保护规则、保比分析指引与公式规范）：",
            font=("", 9),
            foreground="#555555",
        ).pack(anchor="w", pady=(4, 0))

        # Text with scrollbar
        txt_container = ttk.Frame(self, padding=(16, 4, 16, 4))
        txt_container.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(txt_container, orient="vertical")
        text_widget = tk.Text(txt_container, font=("Courier", 10), wrap="word", yscrollcommand=scrollbar.set)
        scrollbar.config(command=text_widget.yview)

        text_widget.insert("1.0", prompt_text if prompt_text else "（暂无生成的 AI 说明，请先执行脱敏）")
        text_widget.config(state="disabled")

        text_widget.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Footer actions
        btn_bar = ttk.Frame(self, padding=(16, 8, 16, 12))
        btn_bar.pack(fill="x")

        def _do_copy():
            self.clipboard_clear()
            self.clipboard_append(self.prompt_content)
            messagebox.showinfo("复制成功", "AI 处理说明已成功复制到剪贴板！")

        ttk.Button(btn_bar, text=" 复制全部说明 ", command=_do_copy).pack(side="left")
        ttk.Button(btn_bar, text=" 关闭 ", command=self.destroy).pack(side="right")


class PrivacyApp(tk.Tk):
    """Excel Privacy Tool V1 Desktop GUI application."""

    def __init__(self):
        super().__init__()
        self.title("Excel 隐私脱敏工具 (Excel Privacy Tool V1)")
        self.geometry("1060x860")
        self.minsize(1060, 760)
        self.resizable(True, True)

        self.controller = GUIWorkflowController()

        # UI reactive string vars
        self.source_path_var = tk.StringVar(value="")
        self.safe_dir_var = tk.StringVar(value=str(self.controller.safe_dir))
        self.restored_dir_var = tk.StringVar(value=str(self.controller.restored_dir))
        self.status_var = tk.StringVar(value="请选择需要脱敏的 Excel 原始文件")
        self.returned_path_var = tk.StringVar(value="")

        # Internal widget tracking
        self.sheet_check_vars: dict[str, tk.BooleanVar] = {}
        self.active_sheet_name: str | None = None
        self.col_widgets: list[dict[str, Any]] = []
        self.last_mask_batch_tag: str = "尚未脱敏"

        self._build_ui()

    def _build_ui(self):
        # Top title bar
        header_frame = ttk.Frame(self, padding=(16, 12, 16, 6))
        header_frame.pack(fill="x")

        ttk.Label(
            header_frame,
            text="Excel 隐私脱敏工具",
            font=("", 18, "bold"),
        ).pack(side="left")

        ttk.Label(
            header_frame,
            text="本地处理 · 可逆脱敏 · AI分析安全闭环",
            font=("", 11),
            foreground="#555555",
        ).pack(side="left", padx=16, pady=4)

        # Main notebook with 2 Tabs: A. 脱敏处理 (Mask)  B. AI 成果还原 (Restore)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=12, pady=(4, 4))

        self.tab_mask = ttk.Frame(self.notebook, padding=(8, 6, 8, 6))
        self.tab_restore = ttk.Frame(self.notebook, padding=(8, 6, 8, 6))

        self.notebook.add(self.tab_mask, text="  1. 工作簿脱敏与说明生成  ")
        self.notebook.add(self.tab_restore, text="  2. AI成果检查与数据恢复  ")

        self._build_mask_tab()
        self._build_restore_tab()

        # Bottom global status & progress bar
        status_bar = ttk.Frame(self, padding=(12, 4, 12, 6))
        status_bar.pack(fill="x", side="bottom")

        self.progress = ttk.Progressbar(status_bar, mode="indeterminate")
        self.progress.pack(fill="x", side="top", pady=(0, 2))

        self.status_lbl = ttk.Label(
            status_bar,
            textvariable=self.status_var,
            font=("", 10, "bold"),
            foreground="#1a5276",
            wraplength=1020,
        )
        self.status_lbl.pack(side="left", fill="x", expand=True)

    # --------------------------------------------------------------------------
    # Tab 1: Mask Workflow
    # --------------------------------------------------------------------------

    def _build_mask_tab(self):
        # Section A: Source file & directories
        sec_a = ttk.LabelFrame(self.tab_mask, text="【文件与输出设置】", padding=(8, 4, 8, 4))
        sec_a.pack(fill="x", pady=(0, 4))

        # Row 1: Source excel
        row1 = ttk.Frame(sec_a)
        row1.pack(fill="x", pady=2)
        ttk.Label(row1, text="原始 Excel：", width=12).pack(side="left")
        ttk.Entry(row1, textvariable=self.source_path_var, state="readonly").pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row1, text="选择文件...", command=self._choose_source_file).pack(side="left")

        # Row 2: Masked output dir
        row2 = ttk.Frame(sec_a)
        row2.pack(fill="x", pady=4)
        ttk.Label(row2, text="脱敏保存到：", width=12).pack(side="left")
        ttk.Entry(row2, textvariable=self.safe_dir_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row2, text="浏览...", command=self._choose_safe_dir).pack(side="left", padx=2)
        ttk.Button(row2, text="打开文件夹", command=lambda: open_folder(self.safe_dir_var.get())).pack(side="left")

        # Section C: Execution and AI Prompt Area (Workbook Action Bar, packed at bottom of tab_mask FIRST)
        self.workbook_action_card = ttk.LabelFrame(self.tab_mask, text="【工作簿脱敏与成果执行】", padding=(8, 4, 8, 4))
        self.workbook_action_card.pack(fill="x", side="bottom", pady=(4, 0))

        # Row 1: Workbook status & primary actions
        btn_row = ttk.Frame(self.workbook_action_card)
        btn_row.pack(fill="x", pady=1)

        self.workbook_status_lbl = ttk.Label(
            btn_row,
            text="工作簿状态：请先选择并配置工作表",
            font=("", 10, "bold"),
            foreground="#2c3e50",
            wraplength=460,
        )
        self.workbook_status_lbl.pack(side="left", fill="x", expand=True)

        self.btn_mask = ttk.Button(btn_row, text=" 开始脱敏处理 ", command=self._start_mask_task, state="disabled")
        self.btn_mask.pack(side="right", padx=(6, 0), ipadx=10, ipady=3)

        self.btn_copy_prompt = ttk.Button(btn_row, text=" 复制给 AI 的说明 ", command=self._copy_prompt, state="disabled")
        self.btn_copy_prompt.pack(side="right", padx=4, ipadx=6, ipady=3)

        self.btn_open_masked = ttk.Button(
            btn_row,
            text="打开脱敏结果文件夹",
            command=lambda: open_folder(self.safe_dir_var.get()),
        )
        self.btn_open_masked.pack(side="right", padx=4)

        # Row 2: Secondary status / Last Mask Product line & Prompt compact preview
        sec_status_row = ttk.Frame(self.workbook_action_card)
        sec_status_row.pack(fill="x", pady=(2, 1))

        self.last_mask_product_lbl = ttk.Label(
            sec_status_row,
            text="【上次脱敏产物】尚未执行脱敏",
            font=("", 9),
            foreground="#7f8c8d",
            wraplength=520,
        )
        self.last_mask_product_lbl.pack(side="left", fill="x", expand=True)

        self.btn_view_full_prompt = ttk.Button(
            sec_status_row,
            text="查看完整说明…",
            command=self._view_full_prompt_dialog,
            state="disabled",
            width=14,
        )
        self.btn_view_full_prompt.pack(side="right", padx=4)

        # Hidden text widget to maintain 100% backward compatibility with test_100
        self.prompt_text = tk.Text(self.workbook_action_card, height=1)
        self.prompt_text.insert("1.0", "（脱敏完成后，将自动生成结构化 AI 提示词说明）")
        self.prompt_text.config(state="disabled")

        # Section B: Multi-Sheet & Column Configuration (Split left/right, takes remaining middle space)
        sec_b = ttk.LabelFrame(self.tab_mask, text="【工作表与字段保护配置】", padding=10)
        sec_b.pack(fill="both", expand=True, pady=4)

        paned = ttk.PanedWindow(sec_b, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # Left pane: Sheet list
        left_frame = ttk.Frame(paned, padding=4)
        paned.add(left_frame, weight=1)

        ttk.Label(left_frame, text="工作表列表（点击切换配置）：", font=("", 11, "bold")).pack(anchor="w", pady=(0, 2))
        self.sheet_sub_lbl = ttk.Label(
            left_frame,
            text="已选择 0 个工作表，每张表需要分别确认保护字段。",
            font=("", 9),
            foreground="#555555",
            wraplength=260,
        )
        self.sheet_sub_lbl.pack(anchor="w", pady=(0, 6))
        self.sheet_list_box = ttk.Frame(left_frame)
        self.sheet_list_box.pack(fill="both", expand=True)

        # Right pane: Active sheet column configuration
        self.right_frame = ttk.Frame(paned, padding=6)
        paned.add(self.right_frame, weight=3)

        # Section A: Header row bar in right pane
        hdr_bar = ttk.Frame(self.right_frame)
        hdr_bar.pack(fill="x", pady=(0, 4))

        self.cur_sheet_title_lbl = ttk.Label(hdr_bar, text="请先选择工作表", font=("", 12, "bold"), foreground="#2980b9")
        self.cur_sheet_title_lbl.pack(side="left")

        ttk.Label(hdr_bar, text="表头所在行：").pack(side="left", padx=(14, 4))
        self.header_row_spin = ttk.Spinbox(hdr_bar, from_=1, to=100, width=5)
        self.header_row_spin.set(1)
        self.header_row_spin.pack(side="left")
        ttk.Button(hdr_bar, text="重新读取字段", command=self._on_header_row_change).pack(side="left", padx=4)

        # Section B: ① Record Identifier Card (Fixed top card, compressed)
        self.key_card = ttk.LabelFrame(self.right_frame, text="【① 记录识别】", padding=(6, 3, 6, 3))
        self.key_card.pack(fill="x", pady=(1, 2))

        # Row 1: Prompt hint, recommendation action, modify action, and help popup icon
        self.rec_key_frame = ttk.Frame(self.key_card)
        self.rec_key_frame.pack(fill="x", pady=(0, 1))

        self.rec_key_prompt_lbl = ttk.Label(
            self.rec_key_frame,
            text="系统建议：无",
            font=("", 9),
            foreground="#2c3e50",
        )
        self.rec_key_prompt_lbl.pack(side="left")

        self.btn_apply_rec_key = ttk.Button(
            self.rec_key_frame,
            text="使用推荐",
            width=9,
            command=self._apply_recommended_key,
        )
        self.btn_apply_rec_key.pack(side="left", padx=(6, 3))

        self.btn_modify_key = ttk.Button(
            self.rec_key_frame,
            text="修改识别字段…",
            command=self._open_record_key_dialog,
        )
        self.btn_modify_key.pack(side="left", padx=3)

        self.key_help_lbl = ttk.Label(
            self.rec_key_frame,
            text=" ⓘ 说明 ",
            font=("", 9),
            foreground="#2980b9",
            cursor="hand2",
        )
        self.key_help_lbl.pack(side="right", padx=2)
        self.key_help_lbl.bind(
            "<Button-1>",
            lambda e: messagebox.showinfo(
                "什么是“记录识别”？",
                "【记录识别】用来唯一识别每一条记录。\n\n"
                "💡 建议选最少的字段组合，只要能够唯一识别每条记录即可（如“员工工号”或“工号 + 年月”）。\n"
                "不要为了保险多选很多字段。",
            ),
        )

        # Row 2: Selected key status
        key_status_row = ttk.Frame(self.key_card)
        key_status_row.pack(fill="x", pady=(1, 0))

        self.key_status_lbl = ttk.Label(
            key_status_row,
            text="当前选择：未设置",
            font=("", 9, "bold"),
            foreground="#34495e",
            wraplength=660,
        )
        self.key_status_lbl.pack(side="left", anchor="w")

        # Section E: ④ Bottom Confirmation Bar (Fixed at bottom of right pane)
        self.confirm_footer = ttk.Frame(self.right_frame, padding=(4, 4, 4, 3))
        self.confirm_footer.pack(fill="x", side="bottom")

        self.footer_summary_lbl = ttk.Label(
            self.confirm_footer,
            text="④ 本表状态：文本保护 0列 · 数值保护 0列 · 识别：未设置 · 未保护 0列",
            font=("", 9),
            foreground="#555555",
            wraplength=520,
        )
        self.footer_summary_lbl.pack(side="left", fill="x", expand=True)

        self.btn_confirm_sheet = ttk.Button(
            self.confirm_footer,
            text=" 确认本表配置 ",
            command=self._confirm_active_sheet,
        )
        self.btn_confirm_sheet.pack(side="right", padx=2)

        # Section D: ③ Advanced Disclosure Card: 数值可比组（高级设置） (Above footer)
        self.adv_group_card = ttk.LabelFrame(self.right_frame, text="【③ 数值可比组（高级设置）】", padding=(6, 2, 6, 3))
        self.adv_group_card.pack(fill="x", side="bottom", pady=(1, 2))

        adv_hdr_row = ttk.Frame(self.adv_group_card)
        adv_hdr_row.pack(fill="x")

        self.btn_adv_toggle = ttk.Button(
            adv_hdr_row,
            text="▸ 展开数值可比组设置",
            width=20,
            command=self._toggle_advanced_group,
        )
        self.btn_adv_toggle.pack(side="left")

        self.adv_summary_lbl = ttk.Label(
            adv_hdr_row,
            text="默认按工作表独立保护（本表 A 组），跨表互不影响。",
            font=("", 9),
            foreground="#555555",
        )
        self.adv_summary_lbl.pack(side="left", padx=6)

        # adv_content_frame with height limit inside Canvas
        self.adv_outer_container = ttk.Frame(self.adv_group_card)
        self.adv_canvas = tk.Canvas(
            self.adv_outer_container,
            highlightthickness=0,
            height=80,
            yscrollincrement=1,
        )
        self.adv_scrollbar = ttk.Scrollbar(self.adv_outer_container, orient="vertical", command=self.adv_canvas.yview)
        self.adv_content_frame = ttk.Frame(self.adv_canvas, padding=(0, 2, 0, 0))

        self.adv_content_frame.bind(
            "<Configure>",
            lambda e: self.adv_canvas.configure(scrollregion=self.adv_canvas.bbox("all")),
        )
        self.adv_canvas_win = self.adv_canvas.create_window((0, 0), window=self.adv_content_frame, anchor="nw")
        self.adv_canvas.bind(
            "<Configure>",
            lambda e: self.adv_canvas.itemconfigure(self.adv_canvas_win, width=e.width),
        )
        self.adv_canvas.configure(yscrollcommand=self.adv_scrollbar.set)
        self._bind_numeric_group_scroll(self.adv_canvas)
        self._bind_numeric_group_scroll(self.adv_content_frame)
        self.adv_expanded = False

        # Section C: ② Field Protection List: Exactly 2 columns: 字段名称 + 保护方式 (Fills remaining middle space)
        self.col_table_card = ttk.LabelFrame(self.right_frame, text="【② 字段保护】", padding=(6, 4, 6, 4))
        self.col_table_card.pack(fill="both", expand=True, pady=(2, 2))

        col_container = ttk.Frame(self.col_table_card)
        col_container.pack(fill="both", expand=True)

        # Fixed Table header above canvas: 75% 字段名称, 25% 保护方式
        fixed_hdr_row = ttk.Frame(col_container, padding=(6, 3))
        fixed_hdr_row.pack(fill="x", pady=(1, 2))
        ttk.Label(fixed_hdr_row, text="字段名称（建议保护 / 识别）", font=("", 10, "bold")).pack(side="left")
        ttk.Label(fixed_hdr_row, text="保护方式", width=14, font=("", 10, "bold")).pack(side="right", padx=(0, 24))
        ttk.Separator(col_container, orient="horizontal").pack(fill="x", pady=(0, 2))

        table_scroll_frame = ttk.Frame(col_container)
        table_scroll_frame.pack(fill="both", expand=True)

        self.col_canvas = tk.Canvas(
            table_scroll_frame,
            highlightthickness=0,
            bg="#ffffff",
            # One scroll unit equals one pixel. Tk 9 TouchpadScroll deltas are
            # high-resolution pixel-like values, including small +/-1 events.
            yscrollincrement=1,
        )
        col_scrollbar = ttk.Scrollbar(table_scroll_frame, orient="vertical", command=self.col_canvas.yview)
        self.col_inner_frame = ttk.Frame(self.col_canvas)

        self.col_inner_frame.bind(
            "<Configure>",
            lambda e: self.col_canvas.configure(scrollregion=self.col_canvas.bbox("all")),
        )
        self.col_canvas_win = self.col_canvas.create_window((0, 0), window=self.col_inner_frame, anchor="nw")
        self.col_canvas.bind(
            "<Configure>",
            lambda e: self.col_canvas.itemconfigure(self.col_canvas_win, width=e.width),
        )
        self.col_canvas.configure(yscrollcommand=col_scrollbar.set)
        self.col_canvas.pack(side="left", fill="both", expand=True)
        col_scrollbar.pack(side="right", fill="y")
        self._bind_field_list_scroll(self.col_canvas)
        self._bind_field_list_scroll(self.col_inner_frame)

        # Initial refresh of workbook action bar status
        self._update_workbook_action_bar()

    # --------------------------------------------------------------------------
    # Tab 2: AI Return & Restore Workflow
    # --------------------------------------------------------------------------

    def _build_restore_tab(self):
        sec_ret = ttk.LabelFrame(self.tab_restore, text="【AI 返回文件与恢复路径设置】", padding=10)
        sec_ret.pack(fill="x", pady=(0, 8))

        # Row 1: Select AI returned file
        r1 = ttk.Frame(sec_ret)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="AI 返回 Excel：", width=14).pack(side="left")
        ttk.Entry(r1, textvariable=self.returned_path_var, state="readonly").pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(r1, text="选择文件...", command=self._choose_returned_file).pack(side="left")

        # Row 1.5: Batch identity feedback banner
        self.batch_feedback_lbl = ttk.Label(
            sec_ret,
            text="",
            font=("", 9),
            foreground="#555555",
            wraplength=960,
        )
        self.batch_feedback_lbl.pack(fill="x", padx=(14, 0), pady=(1, 4))

        # Row 2: Restored output dir
        r2 = ttk.Frame(sec_ret)
        r2.pack(fill="x", pady=4)
        ttk.Label(r2, text="恢复文件保存到：", width=14).pack(side="left")
        ttk.Entry(r2, textvariable=self.restored_dir_var).pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(r2, text="浏览...", command=self._choose_restored_dir).pack(side="left", padx=2)
        ttk.Button(r2, text="打开文件夹", command=lambda: open_folder(self.restored_dir_var.get())).pack(side="left")

        # Section Compare & Conflict Summary
        sec_cmp = ttk.LabelFrame(self.tab_restore, text="【AI 修改检查与冲突摘要】", padding=10)
        sec_cmp.pack(fill="both", expand=True, pady=4)

        cmp_btn_row = ttk.Frame(sec_cmp)
        cmp_btn_row.pack(fill="x", pady=(0, 4))

        self.btn_compare = ttk.Button(cmp_btn_row, text=" 检查 AI 修改 ", command=self._start_compare_task, state="disabled")
        self.btn_compare.pack(side="left", ipadx=10, ipady=4)

        # Status info card for Unconfirmed items
        self.unconfirmed_card_lbl = ttk.Label(
            cmp_btn_row,
            text=" ⓘ 需要人工核对说明 ",
            font=("", 9),
            foreground="#2980b9",
            cursor="hand2",
        )
        self.unconfirmed_card_lbl.pack(side="left", padx=(12, 0))
        self.unconfirmed_card_lbl.bind(
            "<Button-1>",
            lambda e: messagebox.showinfo(
                "什么是“需要人工核对”？",
                "“需要人工核对”表示程序没有足够证据确认应该恢复哪个真实值（例如未设置唯一记录识别列，或行列变动导致无法确切定位）。\n\n为了避免破坏真实数据，程序不会自行猜测或覆盖，这些项目将保持 AI 修改值不变，需用户人工确认。",
            ),
        )

        # Summary line in plain language (三态指标卡)
        self.cmp_summary_lbl = ttk.Label(
            sec_cmp,
            text="未比对",
            font=("", 11, "bold"),
            foreground="#2c3e50",
            wraplength=980,
        )
        self.cmp_summary_lbl.pack(anchor="w", pady=(4, 2))

        # Structural changes display: AI 新增行: 0 | AI 删除行: 0 | AI 新增列: 0
        self.cmp_struct_lbl = ttk.Label(
            sec_cmp,
            text="AI 新增行: 0 | AI 删除行: 0 | AI 新增列: 0",
            font=("", 10),
            foreground="#555555",
            wraplength=980,
        )
        self.cmp_struct_lbl.pack(anchor="w", pady=(0, 2))

        # Primary unconfirmed reason banner
        self.cmp_reason_lbl = ttk.Label(
            sec_cmp,
            text="",
            font=("", 10),
            foreground="#c0392b",
            wraplength=980,
        )
        self.cmp_reason_lbl.pack(anchor="w", pady=(0, 4))

        # Problem summary / breakdown label
        self.cmp_breakdown_lbl = ttk.Label(
            sec_cmp,
            text="",
            font=("", 10, "bold"),
            foreground="#2980b9",
            wraplength=980,
        )
        self.cmp_breakdown_lbl.pack(anchor="w", pady=(0, 4))

        # Fixed explanation note
        ttk.Label(
            sec_cmp,
            text="注：“需要人工核对”表示程序没有足够证据确认应该恢复哪个真实值，因此不会自行猜测或覆盖数据；冲突单元格默认保留 AI 修改值。",
            font=("", 9),
            foreground="#7f8c8d",
        ).pack(anchor="w", pady=(2, 4))

        # Treeview: cell-level detail without exposing original real values.
        self.conflict_tree = ttk.Treeview(
            sec_cmp,
            columns=("sheet", "cell", "field", "reason", "policy", "ai_value"),
            show="headings",
            height=7,
        )
        self.conflict_tree.heading("sheet", text="工作表")
        self.conflict_tree.heading("cell", text="单元格")
        self.conflict_tree.heading("field", text="字段")
        self.conflict_tree.heading("reason", text="原因")
        self.conflict_tree.heading("policy", text="处理策略")
        self.conflict_tree.heading("ai_value", text="AI 当前值")
        self.conflict_tree.column("sheet", width=110)
        self.conflict_tree.column("cell", width=70, anchor="center")
        self.conflict_tree.column("field", width=130)
        self.conflict_tree.column("reason", width=180)
        self.conflict_tree.column("policy", width=190)
        self.conflict_tree.column("ai_value", width=140)
        self.conflict_tree.pack(fill="both", expand=True, pady=2)

        # Section Restore Actions
        sec_act = ttk.Frame(self.tab_restore, padding=(0, 8, 0, 0))
        sec_act.pack(fill="x")

        self.btn_restore = ttk.Button(sec_act, text=" 恢复真实数据 ", command=self._start_restore_task, state="disabled")
        self.btn_restore.pack(side="left", ipadx=14, ipady=6)

        self.btn_open_restored = ttk.Button(
            sec_act,
            text="打开恢复文件夹",
            command=lambda: open_folder(self.restored_dir_var.get()),
        )
        self.btn_open_restored.pack(side="right")

    # --------------------------------------------------------------------------
    # Handlers: File & Directory Choosers
    # --------------------------------------------------------------------------

    def _choose_source_file(self):
        f = filedialog.askopenfilename(
            title="选择需要脱敏的 Excel 原始文件",
            filetypes=[("Excel 文件（V1 仅支持 .xlsx）", "*.xlsx")],
        )
        if not f:
            return
        self.source_path_var.set(f)
        self._set_busy(True, "正在读取并检查工作簿...")

        def _worker():
            ok, msg = self.controller.load_source_workbook(f)
            self.after(0, lambda: self._on_source_loaded(ok, msg))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_source_loaded(self, ok: bool, msg: str):
        self._set_busy(False, msg)
        if not ok:
            messagebox.showerror("打开文件失败", msg)
            return
        self._render_sheet_list()
        if self.controller.available_sheets:
            self._switch_active_sheet(self.controller.available_sheets[0])

    def _choose_safe_dir(self):
        d = filedialog.askdirectory(title="选择脱敏文件保存目录", initialdir=self.safe_dir_var.get())
        if d:
            self.safe_dir_var.set(d)
            self.controller.safe_dir = Path(d).resolve()

    def _choose_restored_dir(self):
        d = filedialog.askdirectory(title="选择恢复文件保存目录", initialdir=self.restored_dir_var.get())
        if d:
            self.restored_dir_var.set(d)
            self.controller.restored_dir = Path(d).resolve()

    def _choose_returned_file(self):
        f = filedialog.askopenfilename(
            title="选择 AI 处理返回的 Excel 文件",
            filetypes=[("Excel 文件（V1 仅支持 .xlsx）", "*.xlsx")],
        )
        if not f:
            return
        self.returned_path_var.set(f)
        self._set_busy(True, "正在匹配脱敏任务账本...")

        def _worker():
            ok, msg = self.controller.load_returned_workbook(f)
            self.after(0, lambda: self._on_returned_loaded(ok, msg))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_returned_loaded(self, ok: bool, msg: str):
        self._set_busy(False, msg)
        if not ok:
            if hasattr(self, "batch_feedback_lbl"):
                self.batch_feedback_lbl.config(
                    text="⚠ 无法确认该文件对应的脱敏批次，已停止自动恢复。",
                    foreground="#c0392b",
                )
            messagebox.showerror("匹配失败", msg)
            self.btn_compare.config(state="disabled")
            self.btn_restore.config(state="disabled")
            return
        if hasattr(self, "batch_feedback_lbl"):
            tag = self.controller.resolved_batch_tag or "已确认"
            self.batch_feedback_lbl.config(
                text=f"✓ 已自动匹配该文件对应的本地恢复资料  |  脱敏批次：{tag} · 批次参考号：{tag}",
                foreground="#27ae60",
            )
        self.btn_compare.config(state="normal")
        self.btn_restore.config(state="normal")
        self.cmp_summary_lbl.config(text="已匹配任务，点击“检查 AI 修改”查看变动详情")

    def _confirm_active_sheet(self):
        if not self.active_sheet_name:
            return
        ok, msg = self.controller.confirm_sheet_configuration(self.active_sheet_name)
        if not ok:
            messagebox.showwarning("确认配置未通过", msg)
        else:
            messagebox.showinfo("配置已确认", msg)
        self._render_column_list()
        self._render_sheet_list()

    def _apply_recommended_key(self):
        if not self.active_sheet_name:
            return
        ok = self.controller.apply_recommended_key(self.active_sheet_name)
        if ok:
            self._render_column_list()
            self._render_sheet_list()

    def _open_record_key_dialog(self):
        if not self.active_sheet_name:
            messagebox.showinfo("提示", "请先选择工作表。")
            return
        RecordKeyDialog(self, self.active_sheet_name)

    def _open_shared_group_dialog(self):
        SharedGroupDialog(self)

    def _toggle_advanced_group(self):
        self.adv_expanded = not self.adv_expanded
        if self.adv_expanded:
            self.btn_adv_toggle.config(text="▾ 收起数值可比组设置")
            self.adv_outer_container.pack(fill="x", expand=True, pady=(2, 0))
            self.adv_canvas.pack(side="left", fill="both", expand=True)
            self.adv_scrollbar.pack(side="right", fill="y")
            self._render_advanced_group_content()
        else:
            self.btn_adv_toggle.config(text="▸ 展开数值可比组设置")
            self.adv_outer_container.pack_forget()

    def _render_advanced_group_content(self):
        for widget in self.adv_content_frame.winfo_children():
            widget.destroy()

        if not self.active_sheet_name:
            self._refresh_advanced_group_geometry()
            return

        cfg = self.controller.sheet_configs.get(self.active_sheet_name)
        if not cfg:
            self._refresh_advanced_group_geometry()
            return

        # Top action row inside expanded area: cross-sheet shared group button & explanatory note
        top_act_row = ttk.Frame(self.adv_content_frame, padding=(2, 2))
        top_act_row.pack(fill="x", pady=(0, 2))
        self.btn_shared_groups = ttk.Button(
            top_act_row,
            text="设置跨表共享组…",
            command=self._open_shared_group_dialog,
        )
        self.btn_shared_groups.pack(side="left")
        ttk.Label(
            top_act_row,
            text="（如需跨工作表保持相同比例比较，点击左侧按钮配置）",
            font=("", 8),
            foreground="#7f8c8d",
        ).pack(side="left", padx=6)

        num_cols = [c for c in cfg.columns if c.protection_type == "numeric"]
        if not num_cols:
            lbl = ttk.Label(
                self.adv_content_frame,
                text="当前表暂无“数值保护”字段。在上表将字段设为“数值保护”后可在此设置可比组。",
                font=("", 9),
                foreground="#7f8c8d",
            )
            lbl.pack(anchor="w", pady=4)
            self._refresh_advanced_group_geometry()
            return

        ttk.Label(
            self.adv_content_frame,
            text="本表数值字段组别分配（不同组使用独立比例脱敏）：",
            font=("", 9, "bold"),
            foreground="#2c3e50",
        ).pack(anchor="w", pady=(2, 4))

        ttk.Label(
            self.adv_content_frame,
            text="加入跨表共享组后，该字段将使用共享组的保护比例，不再使用原本表组比例；移出后恢复本表组。",
            font=("", 8),
            foreground="#7f8c8d",
            wraplength=650,
        ).pack(anchor="w", pady=(0, 4))

        for col in num_cols:
            row = ttk.Frame(self.adv_content_frame, padding=(4, 2))
            row.pack(fill="x")

            ttk.Label(row, text=col.display_name, width=24).pack(side="left")

            if col.shared_group_id:
                ttk.Label(
                    row,
                    text=f"已加入跨表共享组：{col.shared_group_id}",
                    font=("", 9),
                    foreground="#2980b9",
                ).pack(side="left", padx=4)
                btn_rem = ttk.Button(
                    row,
                    text="移回本地组",
                    width=10,
                    command=lambda c=col: self._remove_from_shared(c),
                )
                btn_rem.pack(side="left", padx=4)
            else:
                ttk.Label(row, text="本地组：", font=("", 9)).pack(side="left")
                g_var = tk.StringVar(value=col.local_group)
                g_combo = ttk.Combobox(row, textvariable=g_var, values=["A", "B", "C", "D"], width=5, state="readonly")
                g_combo.pack(side="left", padx=4)

                def _on_local_grp_change(event=None, c=col, gv=g_var):
                    self.controller.set_column_local_group(self.active_sheet_name, c.column_letter, gv.get())
                    self._update_footer_and_adv_summary()
                    self._render_sheet_list()

                g_combo.bind("<<ComboboxSelected>>", _on_local_grp_change)

        self._refresh_advanced_group_geometry()

    def _refresh_advanced_group_geometry(self):
        """Synchronize rebuilt numeric-group content before it becomes visible."""
        self.adv_content_frame.update_idletasks()
        bbox = self.adv_canvas.bbox("all")
        self.adv_canvas.configure(scrollregion=bbox or (0, 0, 0, 0))
        self.adv_canvas.yview_moveto(0.0)

        requested = max(1, self.adv_content_frame.winfo_reqheight())
        canvas_height = min(max(requested, 48), 150)
        self.adv_canvas.configure(height=canvas_height)
        if requested > canvas_height:
            if not self.adv_scrollbar.winfo_manager():
                self.adv_scrollbar.pack(side="right", fill="y")
        else:
            self.adv_scrollbar.pack_forget()

        # Rows and comboboxes are recreated on each refresh, so route events
        # after construction rather than relying on stale child bindings.
        self._bind_numeric_group_scroll(self.adv_content_frame, recursive=True)

    def _remove_from_shared(self, col):
        if not self.active_sheet_name:
            return
        self.controller.remove_from_shared_group(self.active_sheet_name, col.column_letter)
        self._render_column_list()
        self._render_sheet_list()

    # --------------------------------------------------------------------------
    # Renderers: Sheets and Columns
    # --------------------------------------------------------------------------

    def _render_sheet_list(self):
        for widget in self.sheet_list_box.winfo_children():
            widget.destroy()
        self.sheet_check_vars.clear()

        # Update sheet subtitle summary
        total_sheets = len(self.controller.available_sheets)
        sel_count = sum(1 for s in self.controller.available_sheets if self.controller.sheet_configs[s].is_selected)
        self.sheet_sub_lbl.config(text=f"已选择 {sel_count}/{total_sheets} 个工作表，每张表需要分别确认保护字段。")

        for sname in self.controller.available_sheets:
            cfg = self.controller.sheet_configs[sname]
            var = tk.BooleanVar(value=cfg.is_selected)
            self.sheet_check_vars[sname] = var

            is_active = (sname == self.active_sheet_name)
            bg_col = "#eaf2f8" if is_active else "#fdfefe"

            row = tk.Frame(self.sheet_list_box, bg=bg_col, padx=4, pady=3, highlightbackground="#bdc3c7", highlightthickness=1 if is_active else 0)
            row.pack(fill="x", pady=2)

            chk = ttk.Checkbutton(
                row,
                text=sname,
                variable=var,
                command=lambda name=sname: self._on_sheet_toggle(name),
            )
            chk.pack(side="left")

            # Status Badge: ✓ 已确认 vs ● 待确认
            badge_text = "✓ 已确认" if cfg.is_confirmed else "● 待确认"
            badge_fg = "#27ae60" if cfg.is_confirmed else "#d35400"
            badge_lbl = tk.Label(
                row,
                text=badge_text,
                font=("", 9, "bold" if cfg.is_confirmed else "normal"),
                fg=badge_fg,
                bg=bg_col,
            )
            badge_lbl.pack(side="left", padx=6)

            btn = ttk.Button(
                row,
                text="配置 >" if is_active else "选择 >",
                width=8,
                command=lambda name=sname: self._switch_active_sheet(name),
            )
            btn.pack(side="right")

            # Clicking the row background or badge also switches active sheet
            row.bind("<Button-1>", lambda e, name=sname: self._switch_active_sheet(name))
            badge_lbl.bind("<Button-1>", lambda e, name=sname: self._switch_active_sheet(name))

        # Synchronize Workbook Action Bar whenever sheet list or confirmation states are rendered
        self._update_workbook_action_bar()

    def _on_sheet_toggle(self, sheet_name: str):
        var = self.sheet_check_vars.get(sheet_name)
        if var:
            self.controller.select_sheet(sheet_name, var.get())
        total_sheets = len(self.controller.available_sheets)
        sel_count = sum(1 for s in self.controller.available_sheets if self.controller.sheet_configs[s].is_selected)
        self.sheet_sub_lbl.config(text=f"已选择 {sel_count}/{total_sheets} 个工作表，每张表需要分别确认保护字段。")
        self._update_workbook_action_bar()

    def _switch_active_sheet(self, sheet_name: str):
        self.active_sheet_name = sheet_name
        cfg = self.controller.sheet_configs.get(sheet_name)
        if not cfg:
            return

        self.cur_sheet_title_lbl.config(text=f"当前配置工作表：{sheet_name}")
        self.header_row_spin.set(cfg.header_row)
        self._render_column_list()
        self._render_sheet_list()

    def _on_header_row_change(self):
        if not self.active_sheet_name:
            return
        try:
            r = int(self.header_row_spin.get())
        except ValueError:
            messagebox.showerror("输入错误", "表头行必须是正整数")
            return

        ok, msg = self.controller.set_sheet_header_row(self.active_sheet_name, r)
        self.status_var.set(msg)
        if not ok:
            messagebox.showwarning("刷新失败", msg)
        self._render_column_list()
        self._render_sheet_list()

    def _update_footer_and_adv_summary(self):
        """Update bottom confirmation bar summary text and advanced group summary in-place."""
        if not self.active_sheet_name:
            self.footer_summary_lbl.config(text="文本保护 0列 · 数值保护 0列 · 识别：未设置 · 未保护 0列")
            self.adv_summary_lbl.config(text="默认按工作表独立保护（本表 A 组），跨表互不影响。")
            return

        cfg = self.controller.sheet_configs.get(self.active_sheet_name)
        if not cfg:
            return

        t_cnt = len(cfg.get_text_protection_fields())
        n_cnt = len(cfg.get_numeric_protection_fields())
        unprot_cnt = sum(1 for c in cfg.columns if c.protection_type == "none")

        key_fields = cfg.get_record_key_fields()
        if key_fields:
            key_str = " + ".join(key_fields)
            if len(key_str) > 20:
                key_str = key_str[:18] + "…"
        else:
            key_str = "未设置"

        self.footer_summary_lbl.config(
            text=f"文本保护 {t_cnt}列 · 数值保护 {n_cnt}列 · 识别：{key_str} · 未保护 {unprot_cnt}列"
        )

        if n_cnt > 0:
            self.adv_summary_lbl.config(text=f"当前 {n_cnt} 个数值保护字段，默认本表 A 组。")
        else:
            self.adv_summary_lbl.config(text="默认按工作表独立保护（本表 A 组），跨表互不影响。")

    def _update_workbook_action_bar(self):
        """Update Workbook Action Bar status message and mask button state based on selected/confirmed sheets."""
        if not hasattr(self, "workbook_status_lbl") or not hasattr(self, "btn_mask"):
            return

        # If busy running background task, do not alter mask button state here
        if self.controller.is_busy:
            return

        selected_sheets = [s for s in self.controller.sheet_configs.values() if s.is_selected]
        sel_count = len(selected_sheets)

        if sel_count == 0:
            self.workbook_status_lbl.config(
                text="工作簿：请至少勾选 1 个工作表",
                foreground="#7f8c8d",
            )
            self.btn_mask.config(state="disabled")
            return

        unconfirmed = [s.sheet_name for s in selected_sheets if not s.is_confirmed]
        confirmed_count = sel_count - len(unconfirmed)

        if unconfirmed:
            if len(unconfirmed) == 1:
                detail = f"还有 1 个待确认：{unconfirmed[0]}"
            else:
                detail = f"还有 {len(unconfirmed)} 个工作表待确认"
            self.workbook_status_lbl.config(
                text=f"工作簿状态：已确认 {confirmed_count} / {sel_count}（{detail}）",
                foreground="#d35400",
            )
            self.btn_mask.config(state="disabled")
            if hasattr(self, "last_mask_product_lbl") and self.controller.last_mask_result:
                self.last_mask_product_lbl.config(
                    text=f"【上次脱敏产物】批次参考号：{self.last_mask_batch_tag}（当前配置已变更，生成新文件需重新点击脱敏）",
                    foreground="#d35400",
                )
        else:
            self.workbook_status_lbl.config(
                text=f"工作簿状态：✓ 已确认 {confirmed_count} / {sel_count} 个工作表，可以开始脱敏",
                foreground="#27ae60",
            )
            self.btn_mask.config(state="normal")
            if hasattr(self, "last_mask_product_lbl") and self.controller.last_mask_result:
                self.last_mask_product_lbl.config(
                    text=f"【上次脱敏产物】✓ 已生成 · 批次参考号：{self.last_mask_batch_tag}",
                    foreground="#27ae60",
                )

    def _render_column_list(self):
        for widget in self.col_inner_frame.winfo_children():
            widget.destroy()
        self.col_widgets.clear()

        if not self.active_sheet_name:
            empty_card = ttk.Frame(self.col_inner_frame, padding=30)
            empty_card.pack(fill="both", expand=True)
            ttk.Label(empty_card, text="暂无选中的工作表，请在左侧列表中点击选择。", font=("", 11), foreground="#7f8c8d").pack(pady=10)
            self._update_footer_and_adv_summary()
            return

        cfg = self.controller.sheet_configs.get(self.active_sheet_name)
        if not cfg:
            empty_card = ttk.Frame(self.col_inner_frame, padding=30)
            empty_card.pack(fill="both", expand=True)
            ttk.Label(empty_card, text="未找到工作表配置。", font=("", 11), foreground="#7f8c8d").pack(pady=10)
            self._update_footer_and_adv_summary()
            return

        # Empty state: if sheet has no extracted columns
        if not cfg.columns:
            empty_card = ttk.Frame(self.col_inner_frame, padding=40)
            empty_card.pack(fill="both", expand=True)
            ttk.Label(
                empty_card,
                text="该工作表没有可配置的数据字段（可能为空表或表头行设置过大）。",
                font=("", 11, "bold"),
                foreground="#7f8c8d",
            ).pack(pady=(10, 4))
            ttk.Label(
                empty_card,
                text="请检查“表头所在行”设置是否正确，然后点击下方按钮重新读取。",
                font=("", 10),
                foreground="#95a5a6",
            ).pack(pady=(0, 12))
            ttk.Button(empty_card, text="重新读取字段", command=self._on_header_row_change).pack()
            self.key_status_lbl.config(text="当前状态：该表无有效数据字段")
            self.rec_key_prompt_lbl.config(text="系统建议：无")
            self.btn_apply_rec_key.config(state="disabled")
            self._update_footer_and_adv_summary()
            return

        # Update recommendation action row in Card ① (matches frozen spec: 系统建议 / 推荐记录识别列)
        if cfg.recommended_key_desc:
            self.rec_key_prompt_lbl.config(text=f"系统建议（推荐记录识别列）：{cfg.recommended_key_desc}")
            self.btn_apply_rec_key.config(state="normal")
        else:
            self.rec_key_prompt_lbl.config(text="系统建议（推荐记录识别列）：无强候选列")
            self.btn_apply_rec_key.config(state="disabled")

        # Update key status display in Card ①
        selected_key_names = cfg.get_record_key_fields()
        if selected_key_names:
            key_tag_str = " + ".join(selected_key_names)
            self.key_status_lbl.config(text=f"✓ 已选择：{key_tag_str}  |  {cfg.key_validation_status}")
        else:
            self.key_status_lbl.config(text=f"当前状态：{cfg.key_validation_status}")

        # Header placeholder row inside inner frame for structure consistency & test compatibility
        hdr_row = ttk.Frame(self.col_inner_frame, padding=(6, 2))
        hdr_row.pack(fill="x", pady=(0, 2))
        ttk.Label(hdr_row, text="字段名称", font=("", 9, "bold")).pack(side="left")
        ttk.Label(hdr_row, text="保护方式", font=("", 9, "bold")).pack(side="right", padx=12)
        ttk.Separator(self.col_inner_frame, orient="horizontal").pack(fill="x", pady=(0, 2))

        # Populate rows: Exactly 2 columns
        for col in cfg.columns:
            # Row background tint: white (none), light blue (text), light green (numeric)
            if col.protection_type == "text":
                row_bg = "#ebf5fb"
            elif col.protection_type == "numeric":
                row_bg = "#eafaf1"
            else:
                row_bg = "#ffffff"

            crow = tk.Frame(self.col_inner_frame, bg=row_bg, padx=6, pady=4, highlightbackground="#eaeded", highlightthickness=1)
            crow.pack(fill="x", pady=1)

            # Col 1: Label with optional badges 【建议保护】 【识别】
            lbl_frame = tk.Frame(crow, bg=row_bg)
            lbl_frame.pack(side="left", fill="x", expand=True)

            col_label_text = col.display_name
            name_lbl = tk.Label(lbl_frame, text=col_label_text, bg=row_bg, anchor="w", font=("", 9))
            name_lbl.pack(side="left")

            rec_badge_lbl = None
            if col.is_recommended:
                rec_badge_lbl = tk.Label(lbl_frame, text="【建议保护】", bg=row_bg, fg="#e67e22", font=("", 8, "bold"))
                rec_badge_lbl.pack(side="left", padx=2)

            key_badge_lbl = None
            if col.is_record_key:
                key_badge_lbl = tk.Label(lbl_frame, text="【识别】", bg=row_bg, fg="#2980b9", font=("", 8, "bold"))
                key_badge_lbl.pack(side="left", padx=2)

            # Col 2: Protection selector
            if col.is_numeric_candidate:
                combo_values = ["不保护", "文本保护", "数值保护"]
            else:
                combo_values = ["不保护", "文本保护"]

            prot_val_map = {"none": "不保护", "text": "文本保护", "numeric": "数值保护"}
            current_prot_display = prot_val_map.get(col.protection_type, "不保护")
            if current_prot_display not in combo_values:
                current_prot_display = "不保护"

            prot_var = tk.StringVar(value=current_prot_display)
            prot_combo = ttk.Combobox(crow, textvariable=prot_var, values=combo_values, state="readonly", width=11)
            prot_combo.pack(side="right", padx=10)

            # In-place update callback: modifies row styling in-place without rebuilding table or resetting scroll
            def _make_on_prot_change(c=col, pv=prot_var, rf=crow, lf=lbl_frame, nlbl=name_lbl, rbl=rec_badge_lbl, kbl=key_badge_lbl):
                def _handler(event=None):
                    ptype_map = {"不保护": "none", "文本保护": "text", "数值保护": "numeric"}
                    new_ptype = ptype_map.get(pv.get(), "none")
                    self.controller.set_column_protection(self.active_sheet_name, c.column_letter, new_ptype)

                    # In-place visual update
                    if new_ptype == "text":
                        new_bg = "#ebf5fb"
                    elif new_ptype == "numeric":
                        new_bg = "#eafaf1"
                    else:
                        new_bg = "#ffffff"

                    rf.configure(bg=new_bg)
                    lf.configure(bg=new_bg)
                    nlbl.configure(bg=new_bg)
                    if rbl:
                        rbl.configure(bg=new_bg)
                    if kbl:
                        kbl.configure(bg=new_bg)

                    # Update footer summary and advanced group summary
                    self._update_footer_and_adv_summary()
                    if self.adv_expanded:
                        self._render_advanced_group_content()

                    # Invalidate confirmed status in sheet list view
                    self._render_sheet_list()

                return _handler

            prot_combo.bind("<<ComboboxSelected>>", _make_on_prot_change())

        # Trackpad/mouse-wheel events are bound only to the field list subtree,
        # so scrolling elsewhere in the window is never intercepted.
        self._bind_field_list_scroll(self.col_inner_frame, recursive=True)

        # Update footer counts and advanced group
        self._update_footer_and_adv_summary()
        if self.adv_expanded:
            self._render_advanced_group_content()

        # Canvas scroll to top upon rendering new sheet
        self.col_canvas.yview_moveto(0.0)
        self.col_canvas.configure(scrollregion=self.col_canvas.bbox("all"))

    def _bind_field_list_scroll(self, widget, recursive: bool = False):
        """Bind vertical wheel/trackpad scrolling to the field-list canvas only."""
        if not getattr(widget, "_field_list_scroll_bound", False):
            widget.bind("<MouseWheel>", self._on_field_list_mousewheel, add="+")
            widget.bind("<Button-4>", self._on_field_list_mousewheel, add="+")
            widget.bind("<Button-5>", self._on_field_list_mousewheel, add="+")
            if self._supports_touchpad_scroll():
                widget.bind("<TouchpadScroll>", self._on_field_list_touchpad_scroll, add="+")
            widget._field_list_scroll_bound = True
        if recursive:
            for child in widget.winfo_children():
                self._bind_field_list_scroll(child, recursive=True)

    def _bind_numeric_group_scroll(self, widget, recursive: bool = False):
        """Bind wheel/trackpad events only inside the numeric-group subtree."""
        if not getattr(widget, "_numeric_group_scroll_bound", False):
            widget.bind("<MouseWheel>", self._on_numeric_group_mousewheel, add="+")
            widget.bind("<Button-4>", self._on_numeric_group_mousewheel, add="+")
            widget.bind("<Button-5>", self._on_numeric_group_mousewheel, add="+")
            if self._supports_touchpad_scroll():
                widget.bind("<TouchpadScroll>", self._on_numeric_group_touchpad_scroll, add="+")
            widget._numeric_group_scroll_bound = True
        if recursive:
            for child in widget.winfo_children():
                self._bind_numeric_group_scroll(child, recursive=True)

    def _supports_touchpad_scroll(self) -> bool:
        """Return whether this Tk build exposes the Tk 9 precise-scroll event."""
        return bool(self.tk.call("info", "commands", "tk::PreciseScrollDeltas"))

    @staticmethod
    def _unpack_touchpad_scroll_delta(packed_delta: int) -> tuple[int, int]:
        """Decode Tk 9's packed signed 16-bit (x, y) TouchpadScroll delta."""
        packed = int(packed_delta) & 0xFFFFFFFF
        delta_x = (packed >> 16) & 0xFFFF
        delta_y = packed & 0xFFFF
        if delta_x & 0x8000:
            delta_x -= 0x10000
        if delta_y & 0x8000:
            delta_y -= 0x10000
        return delta_x, delta_y

    def _on_field_list_touchpad_scroll(self, event):
        """Route Tk 9 high-resolution touchpad scrolling to the field canvas."""
        _delta_x, delta_y = self._unpack_touchpad_scroll_delta(event.delta)
        if delta_y:
            self.col_canvas.yview_scroll(-delta_y, "units")
        # Stop descendant class bindings (notably TCombobox) from consuming the
        # same gesture and changing a selection instead of scrolling the list.
        return "break"

    def _on_numeric_group_touchpad_scroll(self, event):
        """Route Tk 9 high-resolution gestures to the numeric-group canvas."""
        _delta_x, delta_y = self._unpack_touchpad_scroll_delta(event.delta)
        if delta_y:
            self.adv_canvas.yview_scroll(-delta_y, "units")
        return "break"

    def _on_field_list_mousewheel(self, event):
        """Handle macOS trackpad, Windows wheel, and Linux wheel-button events."""
        button = getattr(event, "num", None)
        if button == 4:
            units = -30
        elif button == 5:
            units = 30
        else:
            delta = getattr(event, "delta", 0)
            if not delta:
                return None
            wheel_steps = int(delta / 120)
            if wheel_steps == 0:
                wheel_steps = 1 if delta > 0 else -1
            units = -wheel_steps * 30
        self.col_canvas.yview_scroll(units, "units")
        return "break"

    def _on_numeric_group_mousewheel(self, event):
        """Keep legacy mouse-wheel paths local to the numeric-group canvas."""
        button = getattr(event, "num", None)
        if button == 4:
            units = -30
        elif button == 5:
            units = 30
        else:
            delta = getattr(event, "delta", 0)
            if not delta:
                return None
            wheel_steps = int(delta / 120)
            if wheel_steps == 0:
                wheel_steps = 1 if delta > 0 else -1
            units = -wheel_steps * 30
        self.adv_canvas.yview_scroll(units, "units")
        return "break"

    # --------------------------------------------------------------------------
    # Background Task Runner (Async threading without blocking UI)
    # --------------------------------------------------------------------------

    def _set_busy(self, busy: bool, message: str = ""):
        self.controller.is_busy = busy
        self.status_var.set(message)
        if busy:
            self.progress.start(10)
            self.btn_mask.config(state="disabled")
            if hasattr(self, "workbook_status_lbl"):
                self.workbook_status_lbl.config(
                    text="工作簿状态：正在脱敏…",
                    foreground="#2980b9",
                )
            self.btn_compare.config(state="disabled")
            self.btn_restore.config(state="disabled")
        else:
            self.progress.stop()
            self._update_workbook_action_bar()
            if self.controller.returned_path:
                self.btn_compare.config(state="normal")
                self.btn_restore.config(state="normal")

    def _start_mask_task(self):
        # Update output dir from string var
        self.controller.safe_dir = Path(self.safe_dir_var.get()).resolve()

        valid, err = self.controller.validate_mask_preconditions()
        if not valid:
            messagebox.showwarning("脱敏配置不完整", err)
            return

        self._set_busy(True, "正在执行本地脱敏与加密账本构建...")

        def _worker():
            try:
                res = self.controller.execute_mask()
                self.after(0, lambda: self._on_mask_success(res))
            except Exception as exc:
                self.after(0, lambda exc=exc: self._on_mask_error(exc))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_mask_error(self, exc: Exception):
        self._set_busy(False, f"脱敏失败：{exc}")
        self.btn_copy_prompt.config(state="disabled")
        self.btn_view_full_prompt.config(state="disabled")
        messagebox.showerror("脱敏失败", str(exc))

    def _on_mask_success(self, result: dict[str, Any]):
        self._set_busy(False, self.controller.status_message)
        self.btn_copy_prompt.config(state="normal")
        self.btn_view_full_prompt.config(state="normal")

        run_id = result.get("run_id", "")
        self.last_mask_batch_tag = compute_run_fingerprint(run_id) if run_id else "已生成"

        if hasattr(self, "workbook_status_lbl"):
            self.workbook_status_lbl.config(
                text=f"工作簿状态：✓ 脱敏完成 · 批次参考号：{self.last_mask_batch_tag}",
                foreground="#27ae60",
            )
        if hasattr(self, "last_mask_product_lbl"):
            self.last_mask_product_lbl.config(
                text=f"【上次脱敏产物】✓ 已生成 · 批次参考号：{self.last_mask_batch_tag}",
                foreground="#27ae60",
            )

        # Update prompt text box
        self.prompt_text.config(state="normal")
        self.prompt_text.delete("1.0", "end")
        self.prompt_text.insert("1.0", self.controller.ai_prompt)
        self.prompt_text.config(state="disabled")

        masked_name = Path(result["masked_workbook_path"]).name
        messagebox.showinfo(
            "脱敏成功",
            f"脱敏文件已安全生成：\n{masked_name}\n批次参考号：{self.last_mask_batch_tag}\n\n已自动提取 AI 处理说明，请点击【复制给 AI 的说明】或【查看完整说明】提供给 AI 工具分析。",
        )

    def _view_full_prompt_dialog(self):
        PromptPreviewDialog(self, self.controller.ai_prompt)

    def _copy_prompt(self):
        p = self.controller.ai_prompt
        if not p:
            messagebox.showinfo("无提示词", "请先执行脱敏生成提示词。")
            return
        self.clipboard_clear()
        self.clipboard_append(p)
        messagebox.showinfo("复制成功", "AI 分析与处理说明已成功复制到剪贴板！可直接粘贴给 AI。")

    def _start_compare_task(self):
        self._set_busy(True, "正在比对 AI 返回文件与脱敏账本...")

        def _worker():
            try:
                summary = self.controller.execute_compare()
                self.after(0, lambda: self._on_compare_success(summary))
            except Exception as exc:
                self.after(0, lambda exc=exc: self._on_compare_error(exc))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_compare_success(self, summary: Any):
        self._set_busy(False, self.controller.status_message)

        # Plain language top line (三态卡片式指标)
        info_txt = (
            f"【可自动恢复】{summary.can_safe_restore_count} 处  |  "
            f"【AI 修改了保护数据】{summary.protected_conflict_count} 处  |  "
            f"【需要人工核对】{summary.require_confirmation_count} 处"
        )
        self.cmp_summary_lbl.config(text=info_txt)

        # Plain language row/column delta stats
        struct_txt = (
            f"AI 新增行: {summary.added_rows_count} | "
            f"AI 删除行: {summary.deleted_rows_count} | "
            f"AI 新增列: {summary.added_columns_count}"
        )
        self.cmp_struct_lbl.config(text=struct_txt)

        # Primary reason banner & conclusions
        if summary.require_confirmation_count == 0 and summary.protected_conflict_count == 0:
            self.cmp_reason_lbl.config(
                text="✓ 未发现需要你处理的问题，可以安全恢复。",
                foreground="#27ae60",
            )
        elif summary.primary_unconfirmed_reason:
            self.cmp_reason_lbl.config(
                text=summary.primary_unconfirmed_reason,
                foreground="#c0392b",
            )
        else:
            self.cmp_reason_lbl.config(text="", foreground="#c0392b")

        # Problem summary / breakdown
        if summary.reason_breakdown:
            breakdown_str = "【问题摘要】： " + "； ".join(f"{k}：{v} 项" for k, v in summary.reason_breakdown.items())
            self.cmp_breakdown_lbl.config(text=breakdown_str)
        else:
            self.cmp_breakdown_lbl.config(text="【问题摘要】：无冲突或需人工核对项")

        # Refill grouped summary treeview (去重复轰炸)
        for item in self.conflict_tree.get_children():
            self.conflict_tree.delete(item)

        if summary.conflict_details:
            for detail in summary.conflict_details:
                self.conflict_tree.insert(
                    "",
                    "end",
                    values=(
                        detail.get("sheet", ""),
                        detail.get("cell", ""),
                        detail.get("field", ""),
                        detail.get("reason", ""),
                        detail.get("policy", ""),
                        detail.get("ai_value", ""),
                    ),
                )
        else:
            self.conflict_tree.insert(
                "",
                "end",
                values=(
                    "—",
                    "—",
                    "—",
                    "✓ 没有需要逐项查看的问题",
                    "—",
                    "—",
                ),
            )

        if summary.has_conflicts or summary.require_confirmation_count > 0:
            messagebox.showwarning(
                "发现数据变动或冲突",
                f"检查完成：\n"
                f"- 【可自动恢复】：{summary.can_safe_restore_count} 处\n"
                f"- 【AI 修改了保护数据】：{summary.protected_conflict_count} 处\n"
                f"- 【需要人工核对】：{summary.require_confirmation_count} 处\n\n"
                f"{summary.primary_unconfirmed_reason}\n\n"
                f"系统将对可安全恢复项还原真实数据；受保护冲突与需核对项目将保留 AI 修改值，不进行破坏性覆盖。",
            )
        else:
            messagebox.showinfo("检查通过", f"✓ 未发现需要你处理的问题，共 {summary.can_safe_restore_count} 处可直接安全恢复！")

    def _on_compare_error(self, exc: Exception):
        self._set_busy(False, f"比对失败：{exc}")
        messagebox.showerror("比对失败", str(exc))

    def _start_restore_task(self):
        # Step 10: Confirmation dialog before restore if unconfirmed items exist
        summary = self.controller.last_compare_summary
        if summary and summary.require_confirmation_count > 0:
            unconf_cnt = summary.require_confirmation_count
            confirm_msg = (
                f"当前有 {unconf_cnt} 项需要人工核对。\n\n"
                f"程序只会恢复已确认安全的项目，其余项目将保持 AI 返回值不变。\n\n"
                f"是否继续安全恢复？"
            )
            proceed = messagebox.askokcancel("恢复前确认", confirm_msg)
            if not proceed:
                return

        self.controller.restored_dir = Path(self.restored_dir_var.get()).resolve()
        self._set_busy(True, "正在执行精准点位还原与公式激活...")

        def _worker():
            try:
                res = self.controller.execute_restore()
                self.after(0, lambda: self._on_restore_success(res))
            except Exception as exc:
                self.after(0, lambda exc=exc: self._on_restore_error(exc))

        threading.Thread(target=_worker, daemon=True).start()

    def _on_restore_success(self, res: Any):
        self._set_busy(False, self.controller.status_message)
        restored_name = Path(res.restored_workbook_path).name

        if res.skipped_conflicts_count > 0 or res.skipped_unconfirmed_count > 0:
            affected_sheets = "、".join(
                f"{sheet} {count} 项"
                for sheet, count in sorted(res.unresolved_by_sheet.items())
                if count
            ) or "无"
            msg = (
                f"恢复文件已生成：\n{restored_name}\n\n"
                f"成功恢复：{res.restored_cells_count} / {res.total_protected_cells_count} 个受保护单元格\n"
                f"跳过冲突：{res.skipped_conflicts_count} 项（保留 AI 修改值）\n"
                f"需要人工核对：{res.skipped_unconfirmed_count} 项（保留 AI 返回值）\n\n"
                f"无法自动恢复的影响范围：{affected_sheets}\n\n"
                f"请人工复核上述工作表。"
            )
            messagebox.showwarning("部分恢复成功（请人工复核）", msg)
        else:
            msg = (
                f"恢复文件已生成：\n{restored_name}\n\n"
                f"成功恢复：{res.restored_cells_count} 个受保护单元格\n"
                f"请在 Excel 中重算公式并核对 AI 新增内容。"
            )
            messagebox.showinfo("恢复完成", msg)

    def _on_restore_error(self, exc: Exception):
        self._set_busy(False, f"恢复失败：{exc}")
        messagebox.showerror("恢复失败", str(exc))


def main():
    if sys.argv[1:2] == ["--internal-ui-smoke"] and len(sys.argv) == 3:
        from privacy_handler.packaging_smoke import run_ui_phase

        run_ui_phase(Path(sys.argv[2]))
        return
    if "--verify-self" in sys.argv[1:]:
        # Internal packaging gate: use the same verifier as every sensitive GUI
        # operation, then exit without constructing a Tk window.
        from privacy_handler.excel_privacy_tool import verify_self_before_sensitive

        verify_self_before_sensitive()
        return
    if sys.argv[1:2] == ["--internal-smoke-mask"] and len(sys.argv) == 5:
        from privacy_handler.packaging_smoke import run_mask_phase

        run_mask_phase(Path(sys.argv[2]), Path(sys.argv[3]), Path(sys.argv[4]))
        return
    if sys.argv[1:2] == ["--internal-smoke-resume"] and len(sys.argv) == 6:
        from privacy_handler.packaging_smoke import run_resume_phase

        run_resume_phase(
            Path(sys.argv[2]), Path(sys.argv[3]), Path(sys.argv[4]), Path(sys.argv[5])
        )
        return
    app = PrivacyApp()
    app.mainloop()


if __name__ == "__main__":
    main()
