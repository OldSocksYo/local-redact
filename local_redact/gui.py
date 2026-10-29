"""CustomTkinter 本机界面。复用 cli.main，不改脱敏逻辑。"""

from __future__ import annotations

import queue
import sys
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from local_redact.cli import ROOT, main as cli_main
from local_redact.mapping import MappingBook

_DEFAULT_LEXICON = ROOT / "lexicon.txt"
_DEFAULT_OUTPUT = ROOT / "output"
_DEFAULT_MAP = ROOT / "data" / "mapping.json"

_LEXICON_HEADER = """# 一行一个要在全文、文件名和子目录名里替换的词：客户全称、简称、厂区地址、两三字的人名。
# 指定固定代号：原词 => 代号
# 不写代号时，工具会按本机密钥生成稳定代号（同一词每次相同）。
# 以 # 开头的是注释。不要把「机构」「报告」这种代号前缀写进来。
#
# 示例环保科技有限公司
# 文三路1号
# 张三 => 人员甲

"""

_AUTO_TYPES_TEXT = """不依赖词表也会自动处理：
· 栏目后的值：委托/受检/建设/送样等单位，客户与企业名称，项目名称，地址，联系人、法定代表人、项目负责人、陪同人，采样/编制/审核/签发人，电话
· 结果表人员列：采样/检测人员等（表头含序号、样品编号时仍整列打码；格内多人按行拆分；肖、闫、付等常见姓也会打码）
· 编号类：项目/合同/许可证号、统一社会信用代码（邮编、报告编号/文号、点位编号、样品编号保留）
· 格式类：手机与固话（含空格、短横线、括号分组）、经纬度与东/北坐标等坐标列、身份证、银行卡、邮箱、网址、内网 IP、密钥类字符串
· 一律删除：图片、公章、签名、二维码、水印、背景图、首页缩略图、嵌入对象；批注和修订里的作者、模板本机路径同样处理
· Excel 同样处理工作表（含隐藏表）、工作表名、页眉页脚、批注和文档属性；图片、宏、数据透视缓存和外部链接会去掉。结果是 .xlsx
需写入词表：客户简称、不在上述栏位里的两三字人名、无标签的页脚地址/英文机构名、裸域名等"""

_HELP_TEXT = f"""会去掉什么

{_AUTO_TYPES_TEXT}

输出目录

开始脱敏会先清空输出目录里已有的文件。目录里已经有文件时，会再问一次。
对照表在「对照表」页查看，文件仍是 data\\mapping.json。里面是原文，不要和脱敏结果一起上传。清空对照表会再确认，只删对照内容，不改已经脱敏的文件。

处理过程中

不要关闭窗口。关闭会中断本次处理，正在转换的 Word、Excel 或 WPS 也可能停在半截。
"""

_STATUS_BUSY = ("#1A5276", "#85C1E9")
_STATUS_OK = ("#196F3D", "#7DCEA0")
_STATUS_BAD = ("#922B21", "#F5B7B1")

_LEXICON_HELP = """写法：一行一个词。只要原词则自动生成稳定代号；需要固定写法时用「原词 => 代号」。
# 开头是注释。保存后写入上方路径；开始脱敏前若有未保存修改会先自动保存。"""

_MAP_COLUMNS = ("kind", "masked", "original")
_MAP_HEADINGS = {"kind": "类型", "masked": "脱敏结果", "original": "原文"}


def _mapping_rows(book: MappingBook) -> list[tuple[str, str, str]]:
    rows: list[tuple[str, str, str]] = []
    for original, renamed in book.files.items():
        rows.append(("文件名", str(renamed), str(original)))
    for token, original in book.tokens.items():
        rows.append(("代号", str(token), str(original)))
    for kind, values in book.removed.items():
        shown = f"[{kind}已移除]"
        for value in values:
            rows.append(("已移除", shown, str(value)))
    return rows


def _read_mapping_rows(path: Path) -> tuple[list[tuple[str, str, str]], str | None]:
    if not path.exists():
        return [], None
    book = MappingBook(path)
    try:
        book.load()
    except (OSError, ValueError) as exc:
        return [], str(exc)
    return _mapping_rows(book), None


def _filter_mapping_rows(
    rows: list[tuple[str, str, str]],
    kind: str,
    query: str,
) -> list[tuple[str, str, str]]:
    picked = rows if kind in ("", "全部") else [row for row in rows if row[0] == kind]
    needle = query.strip().casefold()
    if not needle:
        return picked
    return [row for row in picked if any(needle in cell.casefold() for cell in row)]


class _QueueWriter:
    def __init__(self, target: queue.Queue[str]):
        self._target = target

    def write(self, text: str) -> int:
        if text:
            self._target.put(text)
        return len(text)

    def flush(self) -> None:
        return None


class App(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("local-redact")
        self.geometry("860x760")
        self.minsize(720, 640)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self._log_queue: queue.Queue[str] = queue.Queue()
        self._worker: threading.Thread | None = None
        self._busy = False
        self._lex_dirty = False
        self._lex_loaded_text = ""
        self._lock_widgets: list[ctk.CTkBaseClass] = []
        self._map_rows: list[tuple[str, str, str]] = []
        self._map_error: str | None = None
        self._map_sort: tuple[str | None, bool] = (None, False)
        self._map_note_token = 0
        self.map_query = ctk.StringVar()
        self.map_kind = ctk.StringVar(value="全部")

        self.src_var = ctk.StringVar()
        self.out_var = ctk.StringVar(value=str(_DEFAULT_OUTPUT))
        self.lex_var = ctk.StringVar(
            value=str(_DEFAULT_LEXICON) if _DEFAULT_LEXICON.exists() else str(_DEFAULT_LEXICON)
        )

        self._build()
        self._load_lexicon_into_editor(show_error=False)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(80, self._drain_log)

    def _build(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        tabs = ctk.CTkTabview(self, anchor="nw")
        tabs.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
        self.tabs = tabs
        run_tab = tabs.add("脱敏")
        lex_tab = tabs.add("词表")
        map_tab = tabs.add("对照表")
        help_tab = tabs.add("说明")
        self._build_run_tab(run_tab)
        self._build_lex_tab(lex_tab)
        self._build_map_tab(map_tab)
        self._build_help_tab(help_tab)
        self._reload_mapping()

    def _build_run_tab(self, tab: ctk.CTkFrame) -> None:
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(4, weight=1)

        tip = ctk.CTkLabel(
            tab,
            text=(
                "数据只在本机处理，原稿不会被修改。开始脱敏会先清空输出目录，目录里已有文件时会再确认。\n"
                "对照表含原文，在「对照表」页查看，不要和结果一起上传。会去掉的内容见「说明」页。"
            ),
            anchor="w",
            justify="left",
            wraplength=480,
            text_color=("gray30", "gray75"),
        )
        tip.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 8))
        self._track_wrap(tip)

        form = ctk.CTkFrame(tab, fg_color="transparent")
        form.grid(row=1, column=0, sticky="ew", padx=8, pady=4)
        form.grid_columnconfigure(1, weight=1)

        self._row(form, 0, "输入", self.src_var, self._pick_src_file, self._pick_src_dir)
        self._row(form, 1, "输出目录", self.out_var, self._pick_out_dir)
        self._row(form, 2, "词表文件", self.lex_var, self._pick_lexicon)

        actions = ctk.CTkFrame(tab, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=8, pady=(8, 4))

        self.run_btn = ctk.CTkButton(actions, text="开始脱敏", width=120, command=self._start)
        self.run_btn.grid(row=0, column=0, padx=(0, 8))
        ctk.CTkButton(actions, text="打开输出目录", width=110, command=self._open_output).grid(
            row=0, column=1, padx=4
        )
        ctk.CTkButton(actions, text="打开对照表", width=110, command=self._open_mapping).grid(
            row=0, column=2, padx=4
        )
        ctk.CTkButton(actions, text="清空日志", width=90, command=self._clear_log).grid(
            row=0, column=3, padx=4
        )

        self.status = ctk.CTkLabel(tab, text="就绪", anchor="w")
        self.status.grid(row=3, column=0, sticky="ew", padx=8, pady=(2, 4))

        self.log = ctk.CTkTextbox(tab, wrap="word", font=ctk.CTkFont(family="Consolas", size=13))
        self.log.grid(row=4, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self.log.insert(
            "end",
            "选择输入后点「开始脱敏」。输出目录里已有的文件会被清空。\n"
            "词表在「词表」页编辑，对照表在「对照表」页查看，会去掉的内容见「说明」页。\n",
        )
        self.log.configure(state="disabled")

    def _build_lex_tab(self, tab: ctk.CTkFrame) -> None:
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(2, weight=1)

        lex_help = ctk.CTkLabel(
            tab,
            text=_LEXICON_HELP,
            anchor="w",
            justify="left",
            wraplength=640,
        )
        lex_help.grid(row=0, column=0, sticky="ew", padx=8, pady=(4, 6))
        self._track_wrap(lex_help)

        path_row = ctk.CTkFrame(tab, fg_color="transparent")
        path_row.grid(row=1, column=0, sticky="ew", padx=8, pady=4)
        path_row.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(path_row, text="保存到", width=56, anchor="w").grid(row=0, column=0, padx=(0, 8))
        lex_entry = ctk.CTkEntry(path_row, textvariable=self.lex_var)
        lex_entry.grid(row=0, column=1, sticky="ew")
        self._lock_widgets.append(lex_entry)
        for column, caption, command, width in (
            (2, "浏览", self._pick_lexicon, 72),
            (3, "重新加载", self._reload_lexicon, 88),
            (4, "保存词表", self._save_lexicon, 88),
        ):
            button = ctk.CTkButton(path_row, text=caption, width=width, command=command)
            button.grid(row=0, column=column, padx=(8, 0))
            self._lock_widgets.append(button)

        self.lex_editor = ctk.CTkTextbox(tab, wrap="word", font=ctk.CTkFont(family="Consolas", size=13))
        self.lex_editor.grid(row=2, column=0, sticky="nsew", padx=8, pady=(4, 4))
        self.lex_editor.bind("<<Modified>>", self._on_lex_modified)
        self._lock_widgets.append(self.lex_editor)

        bottom = ctk.CTkFrame(tab, fg_color="transparent")
        bottom.grid(row=3, column=0, sticky="ew", padx=8, pady=(4, 8))
        self.lex_status = ctk.CTkLabel(bottom, text="", anchor="w")
        self.lex_status.grid(row=0, column=0, sticky="w")
        template_btn = ctk.CTkButton(bottom, text="插入示例模板", width=120, command=self._insert_lexicon_template)
        template_btn.grid(row=0, column=1, padx=(12, 0))
        self._lock_widgets.append(template_btn)

    def _build_map_tab(self, tab: ctk.CTkFrame) -> None:
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(3, weight=1)

        tip = ctk.CTkLabel(
            tab,
            text="对照表含原文，只在本窗口查看，不要和脱敏结果一起上传。选中一行后可复制原文。",
            anchor="w",
            justify="left",
            wraplength=480,
            text_color=("gray30", "gray75"),
        )
        tip.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 6))
        self._track_wrap(tip)

        summary_row = ctk.CTkFrame(tab, fg_color="transparent")
        summary_row.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 4))
        summary_row.grid_columnconfigure(0, weight=1)
        self.map_summary = ctk.CTkLabel(
            summary_row,
            text="",
            anchor="w",
            text_color=("gray30", "gray75"),
        )
        self.map_summary.grid(row=0, column=0, sticky="ew")
        self.map_note = ctk.CTkLabel(summary_row, text="", anchor="e", text_color=("#1A5276", "#85C1E9"))
        self.map_note.grid(row=0, column=1, padx=(8, 8))
        ctk.CTkButton(summary_row, text="刷新", width=72, command=self._reload_mapping).grid(
            row=0, column=2, padx=(0, 8)
        )
        ctk.CTkButton(summary_row, text="复制原文", width=88, command=self._copy_mapping_selection).grid(
            row=0, column=3, padx=(0, 8)
        )
        ctk.CTkButton(summary_row, text="清空对照表", width=110, command=self._clear_mapping).grid(
            row=0, column=4
        )

        filter_row = ctk.CTkFrame(tab, fg_color="transparent")
        filter_row.grid(row=2, column=0, sticky="ew", padx=8, pady=(0, 6))
        filter_row.grid_columnconfigure(4, weight=1)
        self._map_filters: dict[str, ctk.CTkButton] = {}
        for index, label in enumerate(("全部", "文件名", "代号", "已移除")):
            button = ctk.CTkButton(
                filter_row,
                text=label,
                width=64 if len(label) == 2 else 84,
                command=lambda name=label: self._set_map_kind(name),
            )
            button.grid(row=0, column=index, padx=(0, 6))
            self._map_filters[label] = button
        self._style_map_filters()
        ctk.CTkEntry(
            filter_row,
            textvariable=self.map_query,
            placeholder_text="筛选代号、原文或文件名",
        ).grid(row=0, column=4, sticky="ew", padx=(6, 0))
        self.map_query.trace_add("write", lambda *_args: self._render_mapping())

        holder = ctk.CTkFrame(tab)
        holder.grid(row=3, column=0, sticky="nsew", padx=8, pady=(0, 8))
        holder.grid_columnconfigure(0, weight=1)
        holder.grid_rowconfigure(0, weight=1)

        scale = float(self._get_window_scaling())
        font_px = max(16, int(round(13 * scale)))
        style = ttk.Style(self)
        style.configure(
            "RedactMap.Treeview",
            font=("Microsoft YaHei UI", -font_px),
            rowheight=int(round(font_px * 2.05)),
        )
        style.configure("RedactMap.Treeview.Heading", font=("Microsoft YaHei UI", -font_px))
        style.map(
            "RedactMap.Treeview",
            background=[("selected", "#1F6AA5")],
            foreground=[("selected", "#FFFFFF")],
        )

        self.map_tree = ttk.Treeview(
            holder,
            columns=_MAP_COLUMNS,
            show="headings",
            style="RedactMap.Treeview",
            selectmode="extended",
        )
        for key in _MAP_COLUMNS:
            self.map_tree.heading(key, text=_MAP_HEADINGS[key], command=lambda column=key: self._sort_mapping(column))
        self.map_tree.column("kind", width=int(88 * scale), minwidth=70, stretch=False, anchor="w")
        self.map_tree.column("masked", width=int(260 * scale), minwidth=140, stretch=True, anchor="w")
        self.map_tree.column("original", width=int(360 * scale), minwidth=160, stretch=True, anchor="w")
        self.map_tree.grid(row=0, column=0, sticky="nsew")
        y_scroll = ctk.CTkScrollbar(holder, command=self.map_tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll = ctk.CTkScrollbar(holder, orientation="horizontal", command=self.map_tree.xview)
        x_scroll.grid(row=1, column=0, sticky="ew")
        self.map_tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.map_tree.bind("<Control-c>", self._copy_mapping_selection)
        self.map_tree.bind("<Double-1>", self._copy_mapping_selection)

    def _build_help_tab(self, tab: ctk.CTkFrame) -> None:
        tab.grid_columnconfigure(0, weight=1)
        tab.grid_rowconfigure(0, weight=1)
        box = ctk.CTkTextbox(tab, wrap="word", font=ctk.CTkFont(size=14))
        box.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        box.insert("end", _HELP_TEXT)
        box.configure(state="disabled")

    def _row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        label: str,
        variable: ctk.StringVar,
        *pickers,
    ) -> None:
        ctk.CTkLabel(parent, text=label, width=72, anchor="w").grid(
            row=row, column=0, sticky="w", pady=4, padx=(0, 8)
        )
        entry = ctk.CTkEntry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", pady=4)
        self._lock_widgets.append(entry)
        buttons = ctk.CTkFrame(parent, fg_color="transparent")
        buttons.grid(row=row, column=2, sticky="e", padx=(8, 0), pady=4)
        captions = ("选文件", "选文件夹") if len(pickers) > 1 else ("浏览",)
        for index, (caption, picker) in enumerate(zip(captions, pickers)):
            button = ctk.CTkButton(buttons, text=caption, width=88, command=picker)
            button.grid(row=0, column=index, padx=(8 if index else 0, 0))
            self._lock_widgets.append(button)

    def _track_wrap(self, label: ctk.CTkLabel) -> None:
        label.bind("<Configure>", lambda event, item=label: self._fit_label(item, event.width), add="+")

    def _fit_label(self, label: ctk.CTkLabel, width: int) -> None:
        """wraplength 会被界面缩放再乘一次，所以先按像素宽度折回去。"""
        if width <= 80:
            return
        target = max(160, int(label._reverse_widget_scaling(width - 16)))
        if int(label.cget("wraplength")) == target:
            return
        label.configure(wraplength=target)

    def _on_close(self) -> None:
        if self._busy and not messagebox.askyesno(
            "正在脱敏",
            "脱敏还在进行。关闭窗口会中断本次处理，Word、Excel 或 WPS 也可能停在半截。\n\n确定关闭？",
            parent=self,
            default=messagebox.NO,
        ):
            return
        self.destroy()

    def _pick_src_file(self) -> None:
        if self._busy:
            return
        path = filedialog.askopenfilename(
            title="选择 Word 或 Excel",
            filetypes=[
                ("Word 和 Excel", "*.doc *.docx *.xls *.xlsx *.xlsm"),
                ("Word 文档", "*.doc *.docx"),
                ("Excel 表格", "*.xls *.xlsx *.xlsm"),
                ("全部文件", "*.*"),
            ],
        )
        if path:
            self.src_var.set(path)

    def _pick_src_dir(self) -> None:
        if self._busy:
            return
        path = filedialog.askdirectory(title="选择报告文件夹")
        if path:
            self.src_var.set(path)

    def _pick_out_dir(self) -> None:
        if self._busy:
            return
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.out_var.set(path)

    def _pick_lexicon(self) -> None:
        if self._busy:
            return
        path = filedialog.askopenfilename(
            title="选择词表",
            filetypes=[("文本", "*.txt"), ("全部文件", "*.*")],
            initialdir=str(ROOT),
        )
        if path:
            self.lex_var.set(path)
            self._load_lexicon_into_editor()

    def _on_lex_modified(self, _event=None) -> None:
        try:
            if self.lex_editor.edit_modified():
                current = self.lex_editor.get("1.0", "end-1c")
                self._lex_dirty = current != self._lex_loaded_text
                self._update_lex_status()
                self.lex_editor.edit_modified(False)
        except Exception:
            pass

    def _update_lex_status(self) -> None:
        path = self.lex_var.get().strip() or str(_DEFAULT_LEXICON)
        flag = "（有未保存修改）" if self._lex_dirty else ""
        self.lex_status.configure(text=f"{path} {flag}".strip())

    def _set_editor_text(self, text: str) -> None:
        self.lex_editor.delete("1.0", "end")
        self.lex_editor.insert("1.0", text)
        self._lex_loaded_text = text
        self._lex_dirty = False
        self.lex_editor.edit_modified(False)
        self._update_lex_status()

    def _load_lexicon_into_editor(self, show_error: bool = True) -> None:
        raw = self.lex_var.get().strip() or str(_DEFAULT_LEXICON)
        path = Path(raw)
        if not path.exists():
            self._set_editor_text(_LEXICON_HEADER)
            if show_error:
                self.lex_status.configure(text=f"文件尚不存在，已载入模板，保存后写入：{path}")
            else:
                self._update_lex_status()
            return
        try:
            text = path.read_text(encoding="utf-8")
        except OSError as exc:
            if show_error:
                messagebox.showerror("错误", f"无法读取词表：\n{exc}")
            return
        self.lex_var.set(str(path))
        self._set_editor_text(text)

    def _reload_lexicon(self) -> None:
        if self._busy:
            return
        if self._lex_dirty and not messagebox.askyesno("重新加载", "有未保存修改，确定放弃并重新加载？"):
            return
        self._load_lexicon_into_editor()

    def _insert_lexicon_template(self) -> None:
        if self._busy:
            return
        current = self.lex_editor.get("1.0", "end-1c").strip()
        if current and not messagebox.askyesno("插入模板", "将用示例模板替换当前编辑内容，继续？"):
            return
        self._set_editor_text(_LEXICON_HEADER)
        self._lex_dirty = True
        self._update_lex_status()

    def _save_lexicon(self, quiet: bool = False) -> bool:
        raw = self.lex_var.get().strip() or str(_DEFAULT_LEXICON)
        path = Path(raw)
        text = self.lex_editor.get("1.0", "end-1c")
        if text and not text.endswith("\n"):
            text += "\n"
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("错误", f"保存词表失败：\n{exc}")
            return False
        self.lex_var.set(str(path))
        self._set_editor_text(path.read_text(encoding="utf-8"))
        if not quiet:
            messagebox.showinfo("词表", f"已保存到：\n{path}")
        self._append_log(f"词表已保存：{path}\n")
        return True

    def _ensure_lexicon_saved(self) -> bool:
        if not self._lex_dirty:
            return True
        return self._save_lexicon(quiet=True)

    def _append_log(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _drain_log(self) -> None:
        try:
            while True:
                self._append_log(self._log_queue.get_nowait())
        except queue.Empty:
            pass
        self.after(80, self._drain_log)

    def _set_status(self, text: str, color: tuple[str, str] | None = None) -> None:
        if color is None:
            self.status.configure(text=text)
        else:
            self.status.configure(text=text, text_color=color)

    def _confirm_output_clear(self, output: Path) -> bool:
        if not output.exists():
            return True
        if not output.is_dir():
            messagebox.showerror("错误", f"输出路径不是文件夹：\n{output}", parent=self)
            return False
        try:
            names = [item.name for item in output.iterdir()]
        except OSError as exc:
            messagebox.showerror("错误", f"无法查看输出目录：\n{exc}", parent=self)
            return False
        if not names:
            return True
        preview = "、".join(name if len(name) <= 40 else name[:37] + "…" for name in names[:5])
        extra = f"\n等共 {len(names)} 项" if len(names) > 5 else ""
        return messagebox.askyesno(
            "清空输出目录",
            f"输出目录里已有文件，开始后会全部删除，且不可恢复：\n{output}\n\n{preview}{extra}\n\n确定继续？",
            parent=self,
            default=messagebox.NO,
        )

    def _set_busy(self, busy: bool) -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.run_btn.configure(state=state, text="处理中…" if busy else "开始脱敏")
        for widget in self._lock_widgets:
            widget.configure(state=state)
        if busy:
            self._set_status("正在脱敏…", _STATUS_BUSY)

    def _start(self) -> None:
        if self._busy:
            return
        src = self.src_var.get().strip()
        out = self.out_var.get().strip()
        if not src:
            messagebox.showwarning("提示", "请先选择输入文件或文件夹。")
            return
        if not out:
            messagebox.showwarning("提示", "请指定输出目录。")
            return
        src_path = Path(src)
        if not src_path.exists():
            messagebox.showerror("错误", f"找不到输入：\n{src}")
            return
        if not self._confirm_output_clear(Path(out)):
            return
        if not self._ensure_lexicon_saved():
            return

        argv = [src, "-o", out]
        lexicon = self.lex_var.get().strip()
        if lexicon:
            argv.extend(["--lexicon", lexicon])

        self._set_busy(True)
        self._append_log("\n—— 开始 ——\n")
        self._worker = threading.Thread(target=self._run_job, args=(argv,), daemon=True)
        self._worker.start()

    def _run_job(self, argv: list[str]) -> None:
        writer = _QueueWriter(self._log_queue)
        old_out, old_err = sys.stdout, sys.stderr
        code = 1
        try:
            sys.stdout = writer  # type: ignore[assignment]
            sys.stderr = writer  # type: ignore[assignment]
            code = cli_main(argv)
        except Exception:
            self._log_queue.put(traceback.format_exc())
            code = 1
        finally:
            sys.stdout = old_out
            sys.stderr = old_err
            self.after(0, lambda: self._finished(code))

    def _finished(self, code: int) -> None:
        self._set_busy(False)
        self._reload_mapping()
        if code == 0:
            self._set_status("完成", _STATUS_OK)
            self._append_log("—— 完成 ——\n")
        else:
            self._set_status("有失败或错误，请看下方日志", _STATUS_BAD)
            self._append_log("—— 结束（有失败或错误）——\n")

    def _open_output(self) -> None:
        path = Path(self.out_var.get().strip() or str(_DEFAULT_OUTPUT))
        path.mkdir(parents=True, exist_ok=True)
        self._reveal(path)

    def _open_mapping(self) -> None:
        self._reload_mapping()
        self.tabs.set("对照表")

    def _set_map_kind(self, kind: str) -> None:
        self.map_kind.set(kind)
        self._style_map_filters()
        self._render_mapping()

    def _style_map_filters(self) -> None:
        current = self.map_kind.get()
        for name, button in self._map_filters.items():
            if name == current:
                button.configure(fg_color=("#1F6AA5", "#1F6AA5"), hover_color=("#1A5C91", "#1A5C91"))
            else:
                button.configure(fg_color=("gray75", "gray35"), hover_color=("gray65", "gray28"))

    def _reload_mapping(self) -> None:
        self._map_rows, self._map_error = _read_mapping_rows(_DEFAULT_MAP)
        self._render_mapping()

    def _render_mapping(self, _value: str | None = None) -> None:
        tree = getattr(self, "map_tree", None)
        if tree is None:
            return
        tree.delete(*tree.get_children())
        shown = _filter_mapping_rows(self._map_rows, self.map_kind.get(), self.map_query.get())
        shown = self._ordered_mapping_rows(shown)
        dark = ctk.get_appearance_mode() == "Dark"
        if dark:
            tree.tag_configure("even", background="#2B2B2B", foreground="#E8E8E8")
            tree.tag_configure("odd", background="#343638", foreground="#E8E8E8")
        else:
            tree.tag_configure("even", background="#F7F8FA", foreground="#1C1C1C")
            tree.tag_configure("odd", background="#FFFFFF", foreground="#1C1C1C")
        for index, (kind, masked, original) in enumerate(shown):
            tree.insert("", "end", values=(kind, masked, original), tags=("odd" if index % 2 else "even",))
        self._paint_mapping_summary(len(shown))

    def _ordered_mapping_rows(self, rows: list[tuple[str, str, str]]) -> list[tuple[str, str, str]]:
        column, descending = self._map_sort
        if column not in _MAP_COLUMNS:
            return rows
        index = _MAP_COLUMNS.index(column)
        return sorted(rows, key=lambda row: row[index].casefold(), reverse=descending)

    def _sort_mapping(self, column: str) -> None:
        current, descending = self._map_sort
        self._map_sort = (column, not descending if current == column else False)
        labels = dict(_MAP_HEADINGS)
        arrow = " ↓" if self._map_sort[1] else " ↑"
        labels[column] = labels[column] + arrow
        for key, text in labels.items():
            self.map_tree.heading(key, text=text, command=lambda name=key: self._sort_mapping(name))
        self._render_mapping()

    def _paint_mapping_summary(self, shown: int) -> None:
        if self._map_error:
            self.map_summary.configure(text=self._map_error, text_color=_STATUS_BAD)
            return
        if not _DEFAULT_MAP.exists():
            self.map_summary.configure(
                text=f"对照表尚未生成。跑完脱敏后会出现在：{_DEFAULT_MAP}",
                text_color=("gray30", "gray75"),
            )
            return
        if not self._map_rows:
            self.map_summary.configure(text="对照表是空的。", text_color=("gray30", "gray75"))
            return
        counts = {kind: 0 for kind in ("文件名", "代号", "已移除")}
        for kind, _masked, _original in self._map_rows:
            counts[kind] = counts.get(kind, 0) + 1
        text = (
            f"文件名 {counts['文件名']} · 代号 {counts['代号']} · 已移除 {counts['已移除']}"
        )
        if shown != len(self._map_rows):
            text = f"显示 {shown} 条 · {text}"
        self.map_summary.configure(text=text, text_color=("gray30", "gray75"))

    def _clear_mapping(self) -> None:
        if self._busy:
            messagebox.showinfo(
                "清空对照表",
                "正在脱敏，对照表还在写入。请等这次处理结束再清空。",
                parent=self,
            )
            return
        path = _DEFAULT_MAP
        if not path.exists():
            self._flash_map_note("对照表还没有内容")
            return
        rows, error = _read_mapping_rows(path)
        if error is None and not rows:
            self._flash_map_note("对照表已经是空的")
            return
        if not messagebox.askyesno(
            "清空对照表",
            "对照表里的原文、代号和文件名对应都会删掉，且不可恢复。已经脱敏的文件不会改动。\n\n确定清空？",
            parent=self,
            default=messagebox.NO,
        ):
            return
        try:
            MappingBook(path).clear()
        except OSError as exc:
            messagebox.showerror("清空对照表", f"无法写入对照表：\n{exc}", parent=self)
            return
        self._reload_mapping()
        self._flash_map_note("已清空对照表")

    def _copy_mapping_selection(self, _event=None) -> str:
        values = [self.map_tree.set(item, "original") for item in self.map_tree.selection()]
        values = [item for item in values if item]
        if not values:
            self._flash_map_note("先选中一行")
            return "break"
        self.clipboard_clear()
        self.clipboard_append("\n".join(values))
        self._flash_map_note("已复制原文" if len(values) == 1 else f"已复制 {len(values)} 条原文")
        return "break"

    def _flash_map_note(self, text: str) -> None:
        self._map_note_token += 1
        token = self._map_note_token
        self.map_note.configure(text=text)
        self.after(2200, lambda: self._clear_map_note(token))

    def _clear_map_note(self, token: int) -> None:
        if token != self._map_note_token or not self.winfo_exists():
            return
        self.map_note.configure(text="")

    @staticmethod
    def _reveal(path: Path) -> None:
        import os

        os.startfile(path)  # noqa: S606 — Windows 本机打开


def main() -> int:
    app = App()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
